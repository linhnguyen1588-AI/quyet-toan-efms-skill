from bs4 import BeautifulSoup
import re

with open('invoice_page.html', 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f, 'html.parser')

tables = soup.find_all('table')
print(f"Found {len(tables)} tables")

if tables:
    for t_idx, t in enumerate(tables):
        print(f"\n--- TABLE {t_idx} ---")
        thead = t.find('thead')
        if thead:
            headers = [th.text.strip() for th in thead.find_all(['th', 'td'])]
            print("Headers:", headers)
        
        rows = t.find_all('tr')
        # Skip header rows for data
        data_rows = [r for r in rows if r.parent.name == 'tbody'] or rows
        print(f"Total Data Rows: {len(data_rows)}")
        
        if data_rows:
            # Look at first data row
            r = data_rows[0]
            tds = r.find_all('td')
            print(f"Row 0 has {len(tds)} columns")
            for c_idx, td in enumerate(tds):
                text = td.text.strip()
                if len(text) > 50: text = text[:50] + "..."
                
                buttons = td.find_all(['button', 'a'])
                btn_info = []
                for b in buttons:
                    classes = b.get('class', [])
                    title = b.get('title', '')
                    btn_text = b.text.strip()
                    btn_info.append(f"<{b.name}> '{btn_text}' class={classes} title='{title}'")
                
                if btn_info:
                    print(f"  Col {c_idx+1}: {text} -> {btn_info}")
                elif text:
                    print(f"  Col {c_idx+1}: {text}")
else:
    # If no tables, look for something that might be a grid
    grids = soup.find_all(class_=re.compile("grid|row|table", re.I))
    print(f"Found {len(grids)} potential grid elements")
    
# also search for any buttons with 'tải' or 'download' or 'pdf'
btns = soup.find_all(['button', 'a'])
print(f"\nTotal buttons/links on page: {len(btns)}")
for b in btns:
    text = b.text.lower()
    title = (b.get('title') or '').lower()
    if 'tải' in text or 'tải' in title or 'download' in text or 'download' in title or 'pdf' in text or 'pdf' in title:
        print(f"Found related button: <{b.name}> '{b.text.strip()}' class={b.get('class')} title='{b.get('title')}'")
        parent = b.parent
        print(f"  Parent: <{parent.name}> class={parent.get('class')}")
