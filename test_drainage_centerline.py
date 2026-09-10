import win32com.client
import math

acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

handle = "7462A"
obj = doc.HandleToObject(handle)

ins = tuple(obj.InsertionPoint)
rot = obj.Rotation
scale = obj.XScaleFactor

def local_to_world(x, y):
    x *= scale
    y *= scale

    return (
        ins[0] + x * math.cos(rot) - y * math.sin(rot),
        ins[1] + x * math.sin(rot) + y * math.cos(rot),
    )

block = doc.Blocks.Item(obj.Name)

lines = []

for e in block:

    if e.ObjectName != "AcDbLine":
        continue

    if e.Layer != "E Pipe":
        continue

    sp = tuple(e.StartPoint)
    ep = tuple(e.EndPoint)

    p1 = local_to_world(sp[0], sp[1])
    p2 = local_to_world(ep[0], ep[1])

    length = math.hypot(
        p2[0] - p1[0],
        p2[1] - p1[1]
    )

    lines.append((p1, p2, length))


lines.sort(key=lambda x: x[2], reverse=True)

print("DRAINAGE:", handle)
print("=" * 90)

for i, (p1, p2, length) in enumerate(lines, 1):

    angle = math.degrees(
        math.atan2(
            p2[1] - p1[1],
            p2[0] - p1[0]
        )
    )

    print(
        f"Line {i}: "
        f"P1=({p1[0]:.4f},{p1[1]:.4f}) "
        f"P2=({p2[0]:.4f},{p2[1]:.4f}) "
        f"Length={length:.4f} "
        f"Axis={angle:.4f}°"
    )


# دو خط اصلی با بیشترین فاصله عرضی را پیدا می‌کنیم
best_pair = None
best_offset = -1

for i in range(len(lines)):
    for j in range(i + 1, len(lines)):

        a1, a2, _ = lines[i]
        b1, b2, _ = lines[j]

        offset = math.hypot(
            ((a1[0] + a2[0]) / 2) - ((b1[0] + b2[0]) / 2),
            ((a1[1] + a2[1]) / 2) - ((b1[1] + b2[1]) / 2),
        )

        if offset > best_offset:
            best_offset = offset
            best_pair = (lines[i], lines[j])


a1, a2, _ = best_pair[0]
b1, b2, _ = best_pair[1]

center1 = (
    (a1[0] + b1[0]) / 2,
    (a1[1] + b1[1]) / 2,
)

center2 = (
    (a2[0] + b2[0]) / 2,
    (a2[1] + b2[1]) / 2,
)

center_length = math.hypot(
    center2[0] - center1[0],
    center2[1] - center1[1],
)

center_axis = math.degrees(
    math.atan2(
        center2[1] - center1[1],
        center2[0] - center1[0],
    )
)

print()
print("CALCULATED CENTER AXIS")
print("=" * 90)

print(
    f"P1=({center1[0]:.4f},{center1[1]:.4f})"
)

print(
    f"P2=({center2[0]:.4f},{center2[1]:.4f})"
)

print(f"Length={center_length:.4f} m")
print(f"Axis={center_axis:.4f}°")
print(f"Side-line separation={best_offset:.4f} m")
