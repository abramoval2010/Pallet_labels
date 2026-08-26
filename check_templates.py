# check_templates.py
import os
import re

templates_dir = 'templates'

for filename in os.listdir(templates_dir):
    if not filename.endswith('.html'):
        continue

    filepath = os.path.join(templates_dir, filename)
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # Ищем все url_for
    matches = re.findall(r"url_for\('([^']+)'\)", content)

    if matches:
        print(f"\n📁 {filename}:")
        for match in matches:
            # Проверяем, есть ли точка в имени (значит уже исправлено)
            if '.' in match:
                print(f"  ✅ {match}")
            else:
                print(f"  ❌ {match} (нужно исправить)")