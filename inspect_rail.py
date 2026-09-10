import win32com.client


acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

print(f"Drawing: {doc.Name}")
print("-" * 100)

count = 0

for obj in doc.ModelSpace:

    if obj.Layer != "Rail":
        continue

    count += 1

    print(f"\nRail #{count}")
    print(f"  Handle      : {obj.Handle}")
    print(f"  ObjectName  : {obj.ObjectName}")
    print(f"  EntityName  : {obj.EntityName}")
    print(f"  Layer       : {obj.Layer}")

    # بررسی متدها / خصوصیات مهم
    for attr in [
        "StartPoint",
        "EndPoint",
        "Length",
        "Coordinates",
        "NumberOfVertices",
        "Elevation",
        "Closed",
        "Linetype",
    ]:
        try:
            value = getattr(obj, attr)
            print(f"  {attr:<18}: {value}")
        except Exception as e:
            print(f"  {attr:<18}: ERROR -> {e}")

    if count >= 20:
        break

print("-" * 100)
print(f"Inspected: {count} objects")