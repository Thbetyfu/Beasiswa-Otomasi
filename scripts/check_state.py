import requests, json

# Check Ollama
try:
    r = requests.get('http://localhost:11434/api/tags', timeout=5)
    models = [m['name'] for m in r.json().get('models', [])]
    print(f"Ollama models: {models}")
    print(f"minicpm-v available: {'minicpm-v:latest' in models or any('minicpm' in m for m in models)}")
except Exception as e:
    print(f"Ollama error: {e}")

# Check progress
with open('verify_progress.json', 'r', encoding='utf-8') as f:
    d = json.load(f)
stats = {}
for v in d.values():
    s = v.get('status', '?')
    stats[s] = stats.get(s, 0) + 1
print(f"\nProgress: {len(d)} items")
for k, v in sorted(stats.items(), key=lambda x: -x[1]):
    print(f"  {k}: {v}")
