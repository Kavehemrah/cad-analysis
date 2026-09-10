import win32com.client
from collections import Counter


acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

layers = {
    "Rail",
    "Rail_Gap",
    "Km_Rail",
    "_flg-Signaling-Pipe 110",
    "Signaling Pipe",
    "_flg-Point Machine",
    "ازبیلت رآهن",
}

counter = Counter()

print(f"Drawing: {doc.Name}")
print("-" * 70)

for obj in doc.ModelSpace:
    try:
        if obj.Layer in layers:
            counter[(obj.Layer, obj.ObjectName)] += 1
    except Exception:
        pass

for (layer, object_type), count in sorted(counter.items()):
    print(f"{count:5}  {layer:35}  {object_type}")

print("-" * 70)