import json
import os
import pandas as pd
from tkinter.filedialog import askdirectory, asksaveasfilename

path = askdirectory(initialdir="~")
finalData = []
cols = ['List Type', 'Variable Name', 'English']

def handle_data_types(filename, key, value):
    match value:
        case str():
            finalData.append([filename, key, value])
        case list():
            for index, item in enumerate(value):
                handle_data_types(filename, f'{key}[{index}]', item)
        case dict():
            for k, v in value.items():
                handle_data_types(filename, f'{key} - {k}', value.get(k))


for file in os.listdir(path):
    if not file.endswith('.json'):
        continue
    with open(path + '/' + file, 'r') as jsonFile:
        print(file)
        thisData = json.loads(jsonFile.read())
        for key, value in thisData.items():
            handle_data_types(file[:-5], key, value)
df = pd.DataFrame(data=finalData, columns=cols)
save_dir = asksaveasfilename(defaultextension=".xlsx",
                             initialfile="translationData.xlsx", )
with pd.ExcelWriter(save_dir, engine='openpyxl', mode='w') as writer:
    df.to_excel(writer, index=False)
