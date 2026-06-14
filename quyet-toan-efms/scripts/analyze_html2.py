from bs4 import BeautifulSoup
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open(r"d:\workspace-ai\quyet-toan-efms-skill\quyet-toan-efms\scripts\search_page_dump.html", "r", encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

print("--- INPUT FIELDS ---")
for inp in soup.find_all('input'):
    print(f"Input: name='{inp.get('name')}', id='{inp.get('id')}', placeholder='{inp.get('placeholder')}'")

print("\n--- BUTTONS ---")
for btn in soup.find_all('button'):
    print(f"Button: text='{btn.text.strip()}', id='{btn.get('id')}', class='{btn.get('class')}'")

print("\n--- TABLES ---")
for table in soup.find_all('table'):
    print(f"Table id='{table.get('id')}', class='{table.get('class')}'")
    headers = [th.text.strip() for th in table.find_all('th')]
    print(f"  Headers: {headers}")
