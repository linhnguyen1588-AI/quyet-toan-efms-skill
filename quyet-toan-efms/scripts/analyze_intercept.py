import json

with open("eport_intercept_logs.json", "r", encoding="utf-8") as f:
    logs = json.load(f)

print(f"Tổng số request: {len(logs)}")
for log in logs:
    url = log.get("url", "")
    method = log.get("method", "")
    
    # Bỏ qua các API không liên quan
    if "signalr" in url or "GetListNotifycation" in url or "Notification" in url:
        continue
        
    post_data = log.get("post_data", "")
    
    if "Invoice" in url or "FullContainer" in url or "Search" in url or "Order" in url:
        print(f"[{method}] {url}")
        if post_data:
            print(f"   Payload: {post_data[:500]}")
            
    # Search for the booking number anywhere in the stringified log
    log_str = json.dumps(log, ensure_ascii=False)
    if "258527559" in log_str:
        print(f"\n[!!!] TÌM THẤY SỐ BOOK TRONG REQUEST NÀY:")
        print(f"[{method}] {url}")
        print(f"Payload: {post_data}")
        print("-" * 50)
