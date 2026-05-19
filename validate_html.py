import json, re

with open(r'd:\claude2\看板.html', 'r', encoding='utf-8') as f:
    html = f.read()

size_kb = len(html.encode('utf-8')) / 1024
print(f'File size: {size_kb:.0f} KB')

# Check Doctype
print('Has DOCTYPE:', html.startswith('<!DOCTYPE html>'))

# Check script tag
m = re.search(r'<script id="embeddedData" type="application/json">(.+?)</script>', html, re.DOTALL)
if m:
    data_str = m.group(1)
    print(f'Embedded data tag found: {len(data_str)} chars')
    try:
        data = json.loads(data_str)
        print('JSON parsed OK')
        print(f'  Keys: {list(data.keys())}')
        print(f'  Current contracts: {len(data["current"]["contracts"])}')
        print(f'  Last contracts: {len(data["last"]["contracts"])}')
        # Check for ending script tag in data
        if data_str.lower().find('</script>') >= 0:
            print('WARNING: </script> found in data!')
        else:
            print('No script tag conflicts in data')
    except Exception as e:
        print(f'JSON parse error: {e}')
else:
    print('ERROR: Embedded data tag not found!')

# Check key JS functions
for fn in ['switchTab', 'render', 'renderKPI', 'renderDepartmentTable', 'renderChangeDetail', 'renderBottomSection', 'computeTransferAnalysis']:
    if fn in html:
        print(f'  JS OK: {fn}')
    else:
        print(f'  MISSING: {fn}')

# Verify closing tags
open_body = html.count('<body')
close_body = html.count('</body>')
open_html_tag = html.count('<html')
close_html_tag = html.count('</html>')
print(f'Tags: body({open_body}/{close_body}) html({open_html_tag}/{close_html_tag})')
