import openpyxl
from pathlib import Path
import json
import os


xlsx_file = Path('temp', '/Users/aakardhakal/Desktop/temp/7212264827585187114.xlsx')
wb_obj = openpyxl.load_workbook(xlsx_file) 

# Read the active sheet:
sheet = wb_obj.active

data = []
comments = []
j = 16
for i in range(int(sheet["B12"].value)):
    w = "E" + str(j)
    comments.append(sheet[w].value)
    j += 1

if os.stat("Phase3/final_data.json").st_size == 0:
    comment_id = 1
else:
    with open("Phase3/final_data.json", "r") as file:
        temp = (json.load(file))
        comment_id = len(temp) +1

data.append({
    "comment_id" : comment_id,
    "author" : sheet["B4"].value,
    "url" : sheet["B2"].value,
    "published_on" : sheet["B6"].value,
    "total_comments" : int(sheet["B12"].value),
    "video_description" : sheet["B9"].value,
    "comments" : comments
})

if os.stat("Phase3/final_data.json").st_size == 0:
    # json_object = json.dumps(data, indent=4)
    with open("Phase3/final_data.json", "w") as file:
        json.dump(data, file, ensure_ascii=False, indent=4)

else:
    with open("Phase3/final_data.json", "r") as file:
        content = json.load(file)
    for i in data:
        content.append(i)
    with open("Phase3/final_data.json", "w") as file:
        json.dump(content, file, ensure_ascii=False, indent=4)

