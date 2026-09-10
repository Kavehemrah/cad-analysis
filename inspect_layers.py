import win32com.client
from collections import Counter


acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

print(f"Drawing: {doc.Name}")
print("=" * 100)

layer_counts = Counter()

for obj in doc.ModelSpace:
    try:
        layer_counts[obj.Layer] += 1
    except Exception:
        pass


print(f"Total Layers Used: {len(layer_counts)}")
print("=" * 100)

for layer, count in layer_counts.most_common():
    print(f"{count:5}  {layer}")

print("=" * 100)
print("POSSIBLE INFRASTRUCTURE LAYERS")
print("=" * 100)

keywords = [
    "rail",
    "traverse",
    "sleeper",
    "travers",
    "signal",
    "pipe",
    "drain",
    "drainage",
    "ditch",
    "culvert",
    "canal",
    "water",
    "sewer",
    "زهکش",
    "تراورس",
    "لوله",
    "سیگنال",
    "آب",
]

for layer, count in layer_counts.most_common():
    layer_lower = layer.lower()

    if any(keyword.lower() in layer_lower for keyword in keywords):
        print(f"{count:5}  {layer}")
        