import json
with open(r'd:\claude2\data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

cur = data['current']
last = data['last']

print('=== 本次会议 P3 Revenue ===')
for r in cur['ppt_overview']['p3_revenue']:
    print(f"  {r['dept'][:6]}: 营收={r['revenue_current']} 指标={r['revenue_target']} 完成率={r['revenue_rate']} | 自营={r['self_rev_current']} 收款={r['collection']}")

print('\n=== 本次会议 P4 Forecast (with prev from last) ===')
for r in cur['ppt_overview']['p4_forecast']:
    print(f"  {r['dept'][:6]}: 完成={r['current']} 存量={r['backlog_forecast']} 待签={r['pending_forecast']} 全年={r['full_year_forecast']} 上期预测={r['prev_forecast']}")

print('\n=== Contract sample with non-zero monthly plans ===')
for c in cur['contracts']:
    if sum(c['monthly_plan']) > 0:
        print(f"  {c['contract_no']} {c['contract_name'][:40]}")
        print(f"    金额={c['contract_amount']} 部门={c['dept']}")
        print(f"    产值计划: {[v for v in c['monthly_plan'] if v > 0]}")
        print(f"    收款计划: {[v for v in c['monthly_collection_plan'] if v > 0]}")
        break

# Count contracts with non-zero plans
plan_count = sum(1 for c in cur['contracts'] if sum(c['monthly_plan']) > 0)
col_plan_count = sum(1 for c in cur['contracts'] if sum(c['monthly_collection_plan']) > 0)
print(f'\nContracts with non-zero monthly plan: {plan_count}')
print(f'Contracts with non-zero collection plan: {col_plan_count}')

print('\n=== LAST meeting contracts with plans ===')
last_plan_count = sum(1 for c in last['contracts'] if sum(c['monthly_plan']) > 0)
print(f'Total: {len(last["contracts"])}, with plans: {last_plan_count}')

print('\n=== TOP10 ===')
for t in cur['top10']['top10_list'][:3]:
    print(f"  {t['dept']} | {t['contract_name'][:30]}... | {t['plan_amount']}万")

print('\n=== Company Summary ===')
for d in cur['company_summary']['产值']:
    non_zero = [(i+1, m['plan'], m['actual']) for i, m in enumerate(d['months']) if m['plan'] > 0 or m['actual'] > 0]
    if non_zero:
        print(f"  {d['dept'][:6]}: annual_plan={d['annual_plan']} non_zero_months={non_zero}")
