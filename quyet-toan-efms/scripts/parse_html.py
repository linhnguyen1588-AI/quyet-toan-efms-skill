from bs4 import BeautifulSoup

with open('invoice_page.html', 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f, 'html.parser')

table = soup.find('table')
if table:
    print('Found table!')
    thead = table.find('thead')
    if thead:
        headers = [th.text.strip() for th in thead.find_all('th')]
        print('Headers:', headers)
    tbody = table.find('tbody')
    if tbody:
        rows = tbody.find_all('tr')
        print(f'Found {len(rows)} rows in tbody.')
        if rows:
            for r_idx, row in enumerate(rows):
                if 'ant-table-row' in row.get('class', []):
                    print(f'--- ROW {r_idx} ---')
                    tds = row.find_all('td')
                    for i, td in enumerate(tds):
                        buttons = td.find_all('button')
                        if buttons:
                            print(f'  Column {i+1}:')
                            for b in buttons:
                                print(f'    [BUTTON] text="{b.text.strip()}" class="{b.get("class")}" title="{b.get("title")}"')
                                icons = b.find_all(class_=lambda c: c and 'anticon' in c)
                                for icon in icons:
                                    print(f'      [ICON] class="{icon.get("class")}"')
                    break
        else:
            print("No rows in tbody")
    else:
        print("No tbody")
else:
    print("No table found")
