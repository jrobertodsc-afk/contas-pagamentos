import json

with open(r'C:\Users\Roberto\.gemini\antigravity\brain\90f97107-8e39-4406-acff-ac4ac2916622\.system_generated\logs\overview.txt', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

output = []
capture = False
for l in lines:
    if "extract_phase1.py" in l and "write_to_file" in l:
        capture = True
    if capture:
        output.append(l)
        if "}]" in l or "}\n" == l:
            break

with open('extract_phase1_recovered.txt', 'w', encoding='utf-8') as f:
    f.write(''.join(output))
