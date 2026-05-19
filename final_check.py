import json
with open(r'd:\claude2\data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

c = data['current']
print('=== KPI ===')
for r in c['ppt_overview']['p3_revenue']:
    if r['dept'] == '合计':
        print(f"Revenue total: {r['revenue_current']} (expect ~678)")
        print(f"Self revenue: {r['self_rev_current']} (expect ~246)")
        print(f"Collection: {r['collection']} (expect ~4865)")

for r in c['ppt_overview']['p4_forecast']:
    if r['dept'] == '合计':
        print(f"Full year forecast: {r['full_year_forecast']} (expect ~13224)")
        print(f"Prev forecast: {r['prev_forecast']}")

print(f"\nContracts: {len(c['contracts'])}")
print(f"With plans: {sum(1 for x in c['contracts'] if sum(x['monthly_plan']) > 0)}")
print(f"With collection plans: {sum(1 for x in c['contracts'] if sum(x['monthly_collection_plan']) > 0)}")

# Check company summary
cs = c['company_summary']
print(f"\nCompany Summary - Value: {len(cs['产值'])} depts")
for d in cs['产值']:
    non_zero = sum(1 for m in d['months'] if m['plan'] > 0 or m['actual'] > 0)
    if non_zero > 0:
        print(f"  {d['dept']}: annual_plan={d['annual_plan']}, active_months={non_zero}")

print(f"\nLast meeting P3 total:")
for r in data['last']['ppt_overview']['p3_revenue']:
    if r['dept'] == '合计':
        print(f"  Revenue: {r['revenue_current']}, Self: {r['self_rev_current']}, Collection: {r['collection']}")
