def patch():
    with open('scripts/build_efms.py', 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    start_idx = -1
    for i, l in enumerate(lines):
        if l.startswith('def fetch_local_excel_data'):
            start_idx = i
            break
            
    end_idx = -1
    for i in range(start_idx + 1, len(lines)):
        if lines[i].startswith('def extract_tariff_matrix'):
            end_idx = i
            break
            
    new_func = """def fetch_local_excel_data(file_path, month=None, year=None, company=None):
    print(f"[INFO] Fetching data from local Excel: {file_path}")
    shipments = []
    try:
        try:
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names
        except Exception as e:
            print(f"[WARNING] pd.ExcelFile failed ({e}). Retrying with engine='calamine'...")
            xls = pd.ExcelFile(file_path, engine='calamine')
            sheet_names = xls.sheet_names
            
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            except Exception:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None, engine='calamine')
                
            block_parsing_mode = False
            in_month_block = False
            
            if not month:
                block_parsing_mode = True
                in_month_block = True
                
            for index, row in df.iterrows():
                try:
                    col_b = str(row[1]).strip().upper() if pd.notna(row[1]) else ""
                    val = col_b
                    
                    if not block_parsing_mode:
                        if str(month) in col_b and str(year) in col_b:
                            block_parsing_mode = True
                            in_month_block = True
                            print(f"[DEBUG] Found month header in sheet '{sheet_name}': {val}")
                            continue
                    else:
                        if not in_month_block:
                            continue
                            
                        if "THÁNG" in val or "MONTH" in val or "-" in val:
                            import re
                            if re.search(r'\d{4}-\d{2}', val) or re.search(r'\d{1,2}/\d{4}', val) or "THÁNG" in val:
                                if str(month) not in val or str(year) not in val:
                                    in_month_block = False
                                    continue
                                    
                    if not in_month_block:
                        continue
                        
                    so_bill = str(row[2]).strip() if pd.notna(row[2]) else ""
                    if not so_bill or so_bill.lower() == 'nan' or len(so_bill) < 4:
                        continue
                        
                    is_special = False
                    if " " in so_bill:
                        is_special = True
                        
                    raw_bill = so_bill
                    so_bill = so_bill.replace(" ", "")
                    
                    try:
                        cont_20 = float(str(row[5]).replace(',', '')) if pd.notna(row[5]) and str(row[5]).strip() != '' else 0
                    except:
                        cont_20 = 0
                    try:
                        cont_40 = float(str(row[6]).replace(',', '')) if pd.notna(row[6]) and str(row[6]).strip() != '' else 0
                    except:
                        cont_40 = 0
                        
                    total_cont = cont_20 + cont_40
                    if total_cont == 0:
                        total_cont = 1
                        
                    default_luong = "xanh"
                    if company:
                        c_lower = company.lower()
                        if "paritas" in c_lower or "toàn lực" in c_lower or "toan luc" in c_lower:
                            default_luong = "vàng"
                            
                    try:
                        phan_luong = str(row[4]).strip().lower() if pd.notna(row[4]) and str(row[4]).strip().lower() != 'nan' else default_luong
                    except Exception:
                        phan_luong = default_luong
                        
                    declarations_str = str(row[17]).strip()
                    if declarations_str and declarations_str.lower() != 'nan':
                        declarations = [d.strip() for d in declarations_str.replace('\\n', ',').split(',')]
                    else:
                        declarations = [""]
                        
                    to_khai_count = len([d for d in declarations if d])
                    if to_khai_count == 0:
                        to_khai_count = 1
                        
                    phu_thu_hq = 0
                    try:
                        val_o = str(row[14]).strip() if pd.notna(row[14]) else ""
                        if val_o and val_o.lower() != 'nan':
                            phu_thu_hq = float(val_o.replace(',', ''))
                    except ValueError:
                        pass

                    phu_thu_bx = 0
                    try:
                        val_p = str(row[15]).strip() if pd.notna(row[15]) else ""
                        if val_p and val_p.lower() != 'nan':
                            phu_thu_bx = float(val_p.replace(',', ''))
                    except ValueError:
                        pass
                        
                    note = str(row[20]).strip() if pd.notna(row[20]) and str(row[20]).strip().lower() != 'nan' else ""

                    for decl in declarations:
                        shipment = {
                            "so_bill": so_bill,
                            "raw_bill": raw_bill,
                            "is_special": is_special,
                            "phan_luong": phan_luong,
                            "cont_20": cont_20,
                            "cont_40": cont_40,
                            "total_cont": total_cont,
                            "to_khai_count": 1, 
                            "declarations": [decl] if decl else [],
                            "note": note,
                            "phu_thu_hq": phu_thu_hq,
                            "phu_thu_bx": phu_thu_bx
                        }
                        shipments.append(shipment)
                except Exception as row_e:
                    print(f"[DEBUG] Skipping row due to error: {row_e}")
                    continue
        
        return shipments
    except Exception as e:
        print(f"[ERROR] Local Excel parsing failed: {e}")
        return None

"""
    lines[start_idx:end_idx] = [new_func]
    
    with open('scripts/build_efms.py', 'w', encoding='utf-8') as f:
        f.writelines(lines)

patch()
