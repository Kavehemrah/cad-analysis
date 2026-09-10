import win32com.client


acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

print(f"Drawing: {doc.Name}")
print(f"Total objects: {doc.ModelSpace.Count}")
print("-" * 60)

count = 0

for obj in doc.ModelSpace:
    if obj.ObjectName == "AcDbLine":
        start = obj.StartPoint
        end = obj.EndPoint

        print(f"Line #{count + 1}")
        print(f"  Layer : {obj.Layer}")
        print(f"  Start : X={start[0]:.3f}, Y={start[1]:.3f}, Z={start[2]:.3f}")
        print(f"  End   : X={end[0]:.3f}, Y={end[1]:.3f}, Z={end[2]:.3f}")
        print(f"  Length: {obj.Length:.3f}")
        print()

        count += 1

        if count >= 10:
            break

print("-" * 60)
print(f"Inspected: {count} lines")