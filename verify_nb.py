import json
with open(r"d:\movie-app\backend\results_analysis.ipynb", "r", encoding="utf-8") as f:
    nb = json.load(f)
print("Total cells:", len(nb["cells"]))
for i, cell in enumerate(nb["cells"]):
    src = "".join(cell["source"])[:100].replace("\n", " ")
    print(f"Cell {i:02d} [{cell['cell_type']}]: {src}")
