import urllib.request
import re

html = open('eirsrv_dump.html', 'r', encoding='utf-8').read()
js_files = re.findall(r'src="([^"]+\.js[^"]*)"', html)
for js in js_files:
    if 'main' not in js: continue
    url = 'https://smartport.gemadept.com.vn' + (js if js.startswith('/') else '/' + js)
    url = url.split('?')[0]
    print('Downloading', url)
    try:
        content = urllib.request.urlopen(url).read().decode('utf-8')
        apis = re.findall(r'[\'"](/apismp/[^\'"]+(?:pdf|download|print|invoice)[^\'"]*)[\'"]', content, re.IGNORECASE)
        for api in set(apis):
            print('   API:', api)
    except Exception as e:
        print('Error:', e)
