import json
raw = open('sandbox/_t/inv.json', encoding='utf-8').read()
start = raw.find('\n{\n')
data = json.loads(raw[start:])
f = data['fields']; roll = data['rollup']; po = data['po']
print('vendor:', f.get('vendor_name'), '| state:', f.get('state'), '| tax_charged:', f.get('tax_charged'))
print('lines:', len(data['lines']), '| capex:', roll['capex_total'], 'opex:', roll['opex_total'])
print('tax_status:', roll['tax_status'], '| expected:', roll['expected_tax_total'],
      'charged:', roll['tax_charged'], 'over:', roll.get('over_collected'), 'use_tax:', roll['use_tax_owed'])
print('routing:', data['routing'], '| PO:', po['matched'],
      (po['record']['po_number'] if po['record'] else None), '| disc:', po['discrepancies'])
print('multi_juris:', data['multi_jurisdiction'])
print('precedent:', data['precedent']['summary'])
print('quality:', data['quality']['passed'], data['quality']['score'], '| safety:', data['safety']['passed'])
# state citation source (bill-to vs ship-to check)
cite = data['evidence']['citations'].get('state')
print('state citation:', (cite or {}).get('quote'))
print('--- per line ---')
for r in data['lines']:
    print(f"  L{r['n']:>2} {r['description'][:32]:32} cap={r['capex_opex']:5} task={str(r.get('task_code')):8} "
          f"item={str(r.get('item_type')):20} tax={r.get('tax_verdict')} conf={r['confidence']} route={r['route']}")
