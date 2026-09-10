import win32com.client
from collections import Counter


acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

counter = Counter()

for obj in doc.ModelSpace:
    try:
        counter[obj.Layer] += 1
    except Exception:
        pass

print(f"Drawing: {doc.Name}")
print(f"Total objects: {doc.ModelSpace.Count}")
print("-" * 70)

for layer, count in counter.most_common():
    print(f"{count:6}  {layer}")

print("-" * 70)
print(f"Total layers used: {len(counter)}")