import json

with open('verify_progress.json', 'r', encoding='utf-8') as f:
    d = json.load(f)

stats = {}
methods = {}
for v in d.values():
    s = v.get('status', '?')
    m = v.get('method', '?')
    stats[s] = stats.get(s, 0) + 1
    methods[m] = methods.get(m, 0) + 1

print(f"Total processed: {len(d)}")
print(f"\nBy status:")
for k, v in sorted(stats.items(), key=lambda x: -x[1]):
    print(f"  {k}: {v}")
print(f"\nBy method:")
for k, v in sorted(methods.items(), key=lambda x: -x[1]):
    print(f"  {k}: {v}")
