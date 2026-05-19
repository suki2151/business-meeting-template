"""Embed data.json into HTML template"""
import json

# Read data
with open(r'd:\claude2\data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Remove raw contract_check data from HTML (too large, not needed for display)
# Keep essential fields only
for key in ['last', 'current']:
    if key in data:
        # Shrink contracts: keep only needed fields for display
        data[key]['contracts'] = [
            {k: v for k, v in c.items() if k in [
                'dept', 'contract_no', 'contract_name', 'contract_amount',
                'monthly_plan', 'monthly_actual', 'monthly_collection_plan',
                'monthly_collection_actual', 'manager', 'team', 'is_leading',
                'confirmed_2026', 'remaining_value'
            ]}
            for c in data[key]['contracts']
        ]
        # Remove huge contract_check
        if 'contract_check' in data[key]:
            data[key]['contract_check'] = data[key]['contract_check'][:50]

# Minify JSON
json_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))

# Read template
with open(r'd:\claude2\看板_template.html', 'r', encoding='utf-8') as f:
    template = f.read()

# Embed data (into script tag — raw JSON, no quoting needed)
html = template.replace("{{DATA_PLACEHOLDER}}", json_str)

# Write final file
with open(r'd:\claude2\看板.html', 'w', encoding='utf-8') as f:
    f.write(html)

import os
size_kb = os.path.getsize(r'd:\claude2\看板.html') / 1024
print(f'Generated 看板.html: {size_kb:.0f} KB')
