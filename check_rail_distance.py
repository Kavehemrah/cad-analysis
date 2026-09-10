import win32com.client
import math

acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

px = 563259.065
py = 3594765.270

results = []

for obj in doc.ModelSpace:
    if obj.Layer != "Rail" or obj.ObjectName != "AcDbPolyline":
        continue

    coords = list(obj.Coordinates)
    best = (float("inf"), 0, 0, 0)

    for i in range(0, len(coords) - 2, 2):
        x1, y1 = coords[i], coords[i + 1]
        x2, y2 = coords[i + 2], coords[i + 3]

        dx = x2 - x1
        dy = y2 - y1
        length2 = dx * dx + dy * dy

        if length2 == 0:
            continue

        t = ((px - x1) * dx + (py - y1) * dy) / length2
        t = max(0, min(1, t))

        qx = x1 + t * dx
        qy = y1 + t * dy

        distance = math.hypot(px - qx, py - qy)
        angle = math.degrees(math.atan2(dy, dx))

        if distance < best[0]:
            best = (distance, qx, qy, angle)

    results.append(
        (
            best[0],
            obj.Handle,
            obj.Length,
            best[1],
            best[2],
            best[3],
        )
    )

results.sort()

print("DRAINAGE 7462A")
print(f"Position = ({px}, {py})")
print("=" * 100)

for distance, handle, length, qx, qy, angle in results[:10]:
    print(
        f"Handle={handle}  "
        f"Distance={distance:.3f} m  "
        f"RailPoint=({qx:.3f},{qy:.3f})  "
        f"RailAxis={angle:.3f} deg  "
        f"Length={length:.3f} m"
    )
