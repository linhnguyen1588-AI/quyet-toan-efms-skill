import urllib.request
import re

html = open('tracking_inv_dump.html', 'r', encoding='utf-8').read()
js_files = re.findall(r'src="([^"]+\.js[^"]*)"', html)
for js in js_files:
    url = 'https://smartport.gemadept.com.vn' + (js if js.startswith('/') else '/' + js)
    url = url.split('?')[0]
    try:
        content = urllib.request.urlopen(url).read().decode('utf-8')
        if 'vnpt' in content.lower():
            print('FOUND VNPT in', url)
            matches = re.finditer(r'.{0,60}vnpt.{0,60}', content, re.IGNORECASE)
            for m in matches:
                print('SNIPPET:', m.group(0))
    except Exception as e:
        pass
