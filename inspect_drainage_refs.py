import win32com.client
from collections import Counter

acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

rows = []

for obj in doc.ModelSpace:
    try:
        if obj.ObjectName != "AcDbBlockReference":
            continue

        name = str(obj.Name)

        try:
            effective = str(obj.EffectiveName)
        except:
            effective = "N/A"

        layer = str(obj.Layer)

        if "drainage" in name.lower() or "drainage" in effective.lower():
            rows.append((layer, name, effective))

    except:
        pass

print("=" * 100)
print("DRAINAGE BLOCK STRUCTURE")
print("=" * 100)

print("Found:", len(rows))

counts = Counter(rows)

for (layer, name, effective), count in counts.most_common():
    print(
        f"Count={count:4d} | "
        f"Layer={layer:<30} | "
        f"Name={name:<30} | "
        f"EffectiveName={effective}"
    )
