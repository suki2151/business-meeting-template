import json
with open(r'd:\claude2\data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

last_p4 = data['last']['ppt_overview']['p4_forecast']
cur_p4 = data['current']['ppt_overview']['p4_forecast']

print('Last P4 rows:', len(last_p4))
for r in last_p4:
    print(f"  {r['dept']}: full_year={r['full_year_forecast']}")

print('Cur P4 rows:', len(cur_p4))
for r in cur_p4:
    print(f"  {r['dept']}: prev={r['prev_forecast']}")

print('\nDebug company_summary...')
cs = data['current']['company_summary']
print(f"产值 rows: {len(cs['产值'])}")
for r in cs['产值']:
    print(f"  dept={r['dept']}, annual_plan={r['annual_plan']}")
