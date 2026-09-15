# -*- coding: utf-8 -*-
import json

with open('pca_lengua_7mo.json', 'r', encoding='utf-8') as f:
    pca = json.load(f)
pca_json_str = json.dumps(pca, ensure_ascii=False)

with open('index.html', 'r', encoding='utf-8') as f:
    content = f.read()

# Actualizar PCA_LENGUA_DATA
old_marker = 'const PCA_LENGUA_DATA = ['
start_idx = content.find(old_marker)
if start_idx != -1:
    end_idx = content.find(';\n\nlet currentPcaTrimFilter', start_idx)
    if end_idx != -1:
        content = content[:start_idx] + 'const PCA_LENGUA_DATA = ' + pca_json_str + content[end_idx:]
        print('PCA_LENGUA_DATA reemplazado')

# Actualizar filtro p1 a 6 a 11
content = content.replace(
    "currentPcaTrimFilter === 'p1' && (item.sem < 6 || item.sem > 10)",
    "currentPcaTrimFilter === 'p1' && (item.sem < 6 || item.sem > 11)"
)
content = content.replace(
    "📘 Planificación N° 1 (05 Oct · 6 sem)",
    "📘 Planif. Microcurricular (05 Oct · 6 sem)"
)

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(content)

print('index.html sincronizado correctamente')
