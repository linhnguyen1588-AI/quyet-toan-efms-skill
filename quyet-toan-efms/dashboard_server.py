import os
import json
import subprocess
import datetime
import uuid
import sys
import shutil
# Reconfigure stdout/stderr to use UTF-8 encoding
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

from flask import Flask, jsonify, request, send_from_directory

app = Flask(__name__)

# Resolve paths
SKILL_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPTS_DIR = os.path.join(SKILL_DIR, "scripts")
LOGS_DIR = os.path.join(SKILL_DIR, "logs")
CONFIG_FILE = os.path.join(SKILL_DIR, "config.json")
MAP_FILE = os.path.join(SKILL_DIR, "company_map.json")

# Ensure logs directory exists
os.makedirs(LOGS_DIR, exist_ok=True)

# Keep track of running processes
# {task_id: {process: Popen, script: str, start_time: str, params: dict}}
active_tasks = {}

def get_disk_free_gb(path):
    """Return free disk space in GB for the path's drive."""
    try:
        total, used, free = shutil.disk_usage(path)
        return round(free / (1024**3), 2)
    except:
        try:
            total, used, free = shutil.disk_usage("C:\\")
            return round(free / (1024**3), 2)
        except:
            return 0.0

def load_json(filepath):
    if os.path.exists(filepath):
        with open(filepath, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except:
                return {}
    return {}

def save_json(filepath, data):
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

@app.route("/")
def index():
    return send_from_directory(SKILL_DIR, "dashboard.html")

@app.route("/api/status", methods=["GET"])
def get_status():
    config = load_json(CONFIG_FILE)
    company_map = load_json(MAP_FILE)
    output_dir = config.get("output_dir", "")
    
    # Check disk space
    disk_free = get_disk_free_gb(output_dir if output_dir else SKILL_DIR)
    
    # Get active tasks summary (clean up completed ones first)
    finished_tasks = []
    for tid, tinfo in active_tasks.items():
        poll = tinfo["process"].poll()
        if poll is not None:
            finished_tasks.append(tid)
            
    # Mark finished tasks
    for tid in finished_tasks:
        proc = active_tasks[tid]["process"]
        active_tasks[tid]["status"] = "Completed" if proc.returncode == 0 else f"Failed (code {proc.returncode})"
        
    running_count = sum(1 for t in active_tasks.values() if "status" not in t)
    
    tasks_summary = []
    for tid, tinfo in active_tasks.items():
        status = tinfo.get("status", "Running")
        tasks_summary.append({
            "task_id": tid,
            "script": tinfo["script"],
            "start_time": tinfo["start_time"],
            "params": tinfo["params"],
            "status": status
        })
        
    return jsonify({
        "config": config,
        "company_map": company_map,
        "disk_free_gb": disk_free,
        "active_tasks": tasks_summary,
        "running_count": running_count,
        "server_time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })

@app.route("/api/save-config", methods=["POST"])
def save_config():
    data = request.json
    new_config = data.get("config")
    new_map = data.get("company_map")
    
    if new_config:
        save_json(CONFIG_FILE, new_config)
    if new_map:
        save_json(MAP_FILE, new_map)
        
    return jsonify({"success": True, "message": "Cấu hình đã được lưu thành công!"})

@app.route("/api/run-script", methods=["POST"])
def run_script():
    data = request.json
    script_name = data.get("script")
    params = data.get("params", {})
    
    # Generate Task ID
    task_id = str(uuid.uuid4())[:8]
    log_file_path = os.path.join(LOGS_DIR, f"{task_id}.log")
    
    # Construct commands based on script name, with python fallback if uv is missing
    python_bin = sys.executable
    if shutil.which("uv"):
        runner_cmd = ["uv", "run"]
    else:
        runner_cmd = [python_bin]
        
    cmd = list(runner_cmd)
    
    if script_name == "build_efms":
        company = params.get("company", "")
        month = params.get("month", "")
        year = params.get("year", "")
        input_file = params.get("input_file", "")
        split = params.get("split", "1")
        group_tk = params.get("group_tk", True)
        
        bill_col = params.get("bill_col")
        cont_col = params.get("cont_col")
        tk_col = params.get("tk_col")
        
        if shutil.which("uv"):
            cmd.extend([
                "--with", "pandas",
                "--with", "openpyxl",
                "--with", "requests",
                "--with", "python-calamine",
            ])
        cmd.extend([
            os.path.join(SCRIPTS_DIR, "build_efms.py"),
            "--company", str(company),
            "--month", str(month),
            "--year", str(year),
            "--split", str(split)
        ])
        if input_file:
            cmd.extend(["--input-file", input_file])
        if not group_tk:
            cmd.append("--no-group-tk")
            
        if bill_col is not None and str(bill_col).strip() != "":
            cmd.extend(["--bill-col", str(bill_col)])
        if cont_col is not None and str(cont_col).strip() != "":
            cmd.extend(["--cont-col", str(cont_col)])
        if tk_col is not None and str(tk_col).strip() != "":
            cmd.extend(["--tk-col", str(tk_col)])
            
    elif script_name == "download_thuphihatang":
        if shutil.which("uv"):
            cmd.extend(["--with", "playwright", "--with", "pandas", "--with", "openpyxl"])
        cmd.append(os.path.join(SCRIPTS_DIR, "download_thuphihatang.py"))

    elif script_name == "download_invoices_api":
        if shutil.which("uv"):
            cmd.extend(["--with", "playwright", "--with", "pandas", "--with", "openpyxl", "--with", "requests"])
        cmd.append(os.path.join(SCRIPTS_DIR, "download_invoices_api.py"))

    elif script_name == "download_eport_api":
        excel_file = params.get("excel_file", "")
        if shutil.which("uv"):
            cmd.extend(["--with", "requests", "--with", "pandas", "--with", "openpyxl", "--with", "playwright"])
        cmd.append(os.path.join(SCRIPTS_DIR, "download_eport_api.py"))
        if excel_file:
            cmd.append(excel_file)

    elif script_name == "check_customs_clearance":
        if shutil.which("uv"):
            cmd.extend(["--with", "requests", "--with", "pandas", "--with", "openpyxl"])
        cmd.append(os.path.join(SCRIPTS_DIR, "check_customs_clearance.py"))

    elif script_name == "check_buying_fees":
        if shutil.which("uv"):
            cmd.extend(["--with", "pandas", "--with", "openpyxl"])
        cmd.append(os.path.join(SCRIPTS_DIR, "check_buying_fees.py"))

    elif script_name == "check_cutoff_alert":
        if shutil.which("uv"):
            cmd.extend(["--with", "requests", "--with", "pandas", "--with", "openpyxl"])
        cmd.append(os.path.join(SCRIPTS_DIR, "check_cutoff_alert.py"))

    elif script_name == "upload_efms":
        if shutil.which("uv"):
            cmd.extend(["--with", "requests", "--with", "pandas", "--with", "openpyxl"])
        cmd.append(os.path.join(SCRIPTS_DIR, "upload_efms.py"))

    elif script_name == "format_excel":
        if shutil.which("uv"):
            cmd.extend(["--with", "openpyxl"])
        cmd.append(os.path.join(SCRIPTS_DIR, "format_excel.py"))

    elif script_name == "process_buying":
        if shutil.which("uv"):
            cmd.extend(["--with", "openpyxl"])
        cmd.append(os.path.join(SCRIPTS_DIR, "process_buying.py"))

    else:
        return jsonify({"success": False, "message": f"Kịch bản '{script_name}' không hợp lệ!"}), 400
        
    print(f"[INFO] Launching task {task_id}: {' '.join(cmd)}")
    
    # Write starting information to log file
    with open(log_file_path, "w", encoding="utf-8") as lf:
        lf.write(f"=== BẮT ĐẦU CHẠY NHIỆM VỤ [{task_id}] ===\n")
        lf.write(f"Thời gian: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        lf.write(f"Lệnh thực thi: {' '.join(cmd)}\n")
        lf.write("="*60 + "\n\n")
        lf.flush()
        
    # Start subprocess
    try:
        lf = open(log_file_path, "a", encoding="utf-8", errors="replace")
        
        # On Windows, we run via shell
        process = subprocess.Popen(
            cmd,
            stdout=lf,
            stderr=lf,
            cwd=SKILL_DIR,
            shell=True,
            text=True,
            bufsize=1
        )
        
        active_tasks[task_id] = {
            "process": process,
            "script": script_name,
            "start_time": datetime.datetime.now().strftime("%H:%M:%S"),
            "params": params,
            "log_file": log_file_path
        }
        
        return jsonify({
            "success": True,
            "task_id": task_id,
            "message": f"Nhiệm vụ {task_id} đã được khởi chạy ngầm."
        })
    except Exception as e:
        print(f"[ERROR] Failed to run task: {e}")
        return jsonify({"success": False, "message": f"Lỗi khởi chạy: {str(e)}"}), 500

@app.route("/api/task-log/<task_id>", methods=["GET"])
def get_task_log(task_id):
    seek_offset = int(request.args.get("seek", 0))
    log_file = os.path.join(LOGS_DIR, f"{task_id}.log")
    
    if not os.path.exists(log_file):
        return jsonify({"content": "", "offset": 0, "status": "Not Found"})
        
    # Get current status
    status = "Running"
    if task_id in active_tasks:
        poll = active_tasks[task_id]["process"].poll()
        if poll is not None:
            status = "Completed" if poll == 0 else f"Failed (code {poll})"
    else:
        status = "Finished"
        
    # Read new lines
    content = ""
    new_offset = seek_offset
    
    try:
        file_size = os.path.getsize(log_file)
        if file_size > seek_offset:
            with open(log_file, "r", encoding="utf-8", errors="replace") as f:
                f.seek(seek_offset)
                content = f.read()
                new_offset = f.tell()
    except Exception as e:
        content = f"\n[Lỗi đọc file log: {str(e)}]\n"
        
    return jsonify({
        "content": content,
        "offset": new_offset,
        "status": status
    })

@app.route("/api/kill-task", methods=["POST"])
def kill_task():
    data = request.json
    task_id = data.get("task_id")
    
    if task_id not in active_tasks:
        return jsonify({"success": False, "message": "Nhiệm vụ không tồn tại hoặc đã kết thúc!"}), 400
        
    tinfo = active_tasks[task_id]
    process = tinfo["process"]
    
    try:
        if os.name == 'nt':
            subprocess.run(f"taskkill /F /T /PID {process.pid}", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            process.terminate()
            process.wait(timeout=2)
            
        tinfo["status"] = "Killed"
        return jsonify({"success": True, "message": f"Đã dừng nhiệm vụ {task_id} thành công."})
    except Exception as e:
        return jsonify({"success": False, "message": f"Không thể dừng nhiệm vụ: {str(e)}"}), 500

@app.route("/api/files", methods=["GET"])
def list_files():
    config = load_json(CONFIG_FILE)
    output_dir = config.get("output_dir", "")
    
    if not output_dir or not os.path.exists(output_dir):
        return jsonify({"files": [], "error": f"Thư mục '{output_dir}' không tồn tại!"})
        
    files_list = []
    try:
        for entry in os.scandir(output_dir):
            if entry.is_file() and entry.name.endswith((".xls", ".xlsx")):
                stats = entry.stat()
                size_kb = round(stats.st_size / 1024, 1)
                mod_time = datetime.datetime.fromtimestamp(stats.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                
                files_list.append({
                    "name": entry.name,
                    "size_kb": size_kb,
                    "mtime": mod_time,
                    "path": entry.path
                })
        
        files_list.sort(key=lambda x: x["mtime"], reverse=True)
        return jsonify({"files": files_list})
    except Exception as e:
        return jsonify({"files": [], "error": f"Lỗi liệt kê file: {str(e)}"})

@app.route("/api/open-file", methods=["POST"])
def open_file():
    data = request.json
    filename = data.get("filename")
    open_folder = data.get("open_folder", False)
    
    config = load_json(CONFIG_FILE)
    output_dir = config.get("output_dir", "")
    
    if not output_dir or not os.path.exists(output_dir):
        return jsonify({"success": False, "message": "Thư mục đầu ra không tồn tại hoặc chưa được cấu hình!"}), 400
        
    try:
        if open_folder:
            print(f"[INFO] Opening directory: {output_dir}")
            os.startfile(output_dir)
            return jsonify({"success": True, "message": "Đã mở thư mục kết quả."})
        elif filename:
            full_path = os.path.join(output_dir, filename)
            if os.path.exists(full_path):
                print(f"[INFO] Opening file: {full_path}")
                os.startfile(full_path)
                return jsonify({"success": True, "message": f"Đã mở file {filename}."})
            else:
                return jsonify({"success": False, "message": f"File {filename} không tồn tại!"}), 404
        else:
            return jsonify({"success": False, "message": "Yêu cầu không hợp lệ!"}), 400
    except Exception as e:
        return jsonify({"success": False, "message": f"Không thể thực hiện mở file: {str(e)}"}), 500

if __name__ == "__main__":
    print("\n" + "="*80)
    print("      EFMS DECISION SETTLEMENT SYSTEM DASHBOARD SERVER IS RUNNING")
    print("      Địa chỉ truy cập: http://localhost:5000")
    print("="*80 + "\n")
    app.run(host="127.0.0.1", port=5000, debug=True)
