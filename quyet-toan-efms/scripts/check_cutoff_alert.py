import urllib.request
import csv
import io
import datetime
import ctypes
import re
import sys

CSV_URL = "https://docs.google.com/spreadsheets/d/194s0s029dpS2Y1OnUYYe8zJ2QBIM8M1O_N2jggD2LJg/export?format=csv&gid=616508979"
ALERT_HOURS = 24

def parse_date(date_str):
    if not date_str:
        return None
    m = re.search(r'(?:(\d{1,2}):(\d{1,2})\s+)?(\d{1,2})/(\d{1,2})/(\d{4})', date_str)
    if m:
        hour = int(m.group(1)) if m.group(1) else 0
        minute = int(m.group(2)) if m.group(2) else 0
        day = int(m.group(3))
        month = int(m.group(4))
        year = int(m.group(5))
        try:
            return datetime.datetime(year, month, day, hour, minute)
        except ValueError:
            return None
    return None

def main():
    try:
        req = urllib.request.Request(CSV_URL, headers={'User-Agent': 'Mozilla/5.0'})
        response = urllib.request.urlopen(req)
        content = response.read().decode('utf-8')
    except Exception as e:
        print(f"Error fetching CSV: {e}")
        if "--interactive" in sys.argv:
            ctypes.windll.user32.MessageBoxW(0, f"Lỗi không tải được dữ liệu: {e}", "Báo Động Cut Off - Lỗi", 0x10 | 0x0)
        return

    reader = csv.reader(io.StringIO(content))
    data = list(reader)
    if not data:
        return

    now = datetime.datetime.now()
    alerts = []

    for i in range(1, len(data)):
        row = data[i]
        if len(row) < 9:
            continue
        
        status = str(row[0]).strip().upper()
        if status in ["DONE", "HỦY", "HUY", "DONE,"]:
            continue
        
        booking_no = row[3]
        vessel = row[5]
        si_cutoff_raw = row[7]
        cy_cutoff_raw = row[8]
        
        si_cutoff = parse_date(si_cutoff_raw)
        cy_cutoff = parse_date(cy_cutoff_raw)
        
        is_alerting = False
        msg = []
        
        if si_cutoff:
            delta = (si_cutoff - now).total_seconds() / 3600
            # Cảnh báo trước 24h hoặc quá hạn (tối đa 48h để không nhắc mãi những cái cũ rích)
            if -48 < delta <= ALERT_HOURS: 
                if delta < 0:
                    msg.append(f"SI Cut Off ĐÃ QUÁ HẠN ({si_cutoff.strftime('%H:%M %d/%m/%Y')})")
                else:
                    msg.append(f"SI Cut Off sắp tới ({si_cutoff.strftime('%H:%M %d/%m/%Y')})")
                is_alerting = True
                
        if cy_cutoff:
            delta = (cy_cutoff - now).total_seconds() / 3600
            if -48 < delta <= ALERT_HOURS:
                if delta < 0:
                    msg.append(f"CY Cut Off ĐÃ QUÁ HẠN ({cy_cutoff.strftime('%H:%M %d/%m/%Y')})")
                else:
                    msg.append(f"CY Cut Off sắp tới ({cy_cutoff.strftime('%H:%M %d/%m/%Y')})")
                is_alerting = True
                
        if is_alerting:
            alerts.append(f"📦 Booking: {booking_no} | Tàu: {vessel}\n    - " + "\n    - ".join(msg))
            
    if alerts:
        alert_text = f"🚨 CẢNH BÁO CUT OFF TRONG {ALERT_HOURS} GIỜ TỚI!\n\n" + "\n\n".join(alerts)
        ctypes.windll.user32.MessageBoxW(0, alert_text, "🚨 Báo Động Cut Off", 0x30 | 0x0)
        print(f"Triggered {len(alerts)} alerts.")
    else:
        if "--interactive" in sys.argv:
            ctypes.windll.user32.MessageBoxW(0, f"✅ Hiện tại không có Booking nào sắp tới hạn Cut Off trong {ALERT_HOURS}h tới.", "Báo Động Cut Off", 0x40 | 0x0)
        print("No alerts within 24h.")

if __name__ == "__main__":
    main()
