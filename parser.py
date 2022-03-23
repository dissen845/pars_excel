import json
import pandas as pd
from pathlib import Path


this_dir = Path(__file__).resolve().parent
parts = []
new_list = []

for path in this_dir.glob("*.xlsm"):
    print(f'Чтение {path.name}')
    part = pd.read_excel(path, sheet_name='СП', skiprows=1, index_col=0, nrows=30, usecols="A,C:H")
    parts.append(part)

df = pd.concat(parts, axis=1)
df.to_excel("new/new_file.xlsm")
df = pd.ExcelFile('new/new_file.xlsm').parse(index_col=0)
df = df.to_dict()

for key, value in df.items():
    new_list.append(df[key])

y = pd.DataFrame(new_list)\
    .drop_duplicates(["Состав передаваемых прав на объект ",
                      "собственность",
                      "Область/ край",
                      "Муниципальный район",
                      "Цена предложения"])\
    .to_dict('records')

with open("new/file_new_1.json", 'w') as f:
    json.dump(y, f, ensure_ascii=False, default=str)
