from bs4 import BeautifulSoup
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open(r"d:\workspace-ai\quyet-toan-efms-skill\quyet-toan-efms\scripts\search_page_dump.html", "r", encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

btn_search = soup.find('button', class_='btnSearch')
if btn_search:
    # go up to common parent (like form or div)
    parent = btn_search.parent
    for _ in range(3):
        if parent: parent = parent.parent
    
    if parent:
        print("--- INPUTS NEAR SEARCH BUTTON ---")
        for inp in parent.find_all(['input', 'select', 'textarea']):
            print(f"{inp.name}: id='{inp.get('id')}', name='{inp.get('name')}', class='{inp.get('class')}'")
else:
    print("No btnSearch found.")
