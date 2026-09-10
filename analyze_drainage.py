import win32com.client
import math
from collections import Counter

acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

BLOCK_NAME = "Drainage Pipe E"


def local_to_world(x, y, ins, rot, scale):
    x *= scale
    y *= scale

    wx = ins[0] + x * math.cos(rot) - y * math.sin(rot)
    wy = ins[1] + x * math.sin(rot) + y * math.cos(rot)

    return wx, wy


def normalize_angle(angle):
    while angle >= 180:
        angle -= 360
    while angle < -180:
        angle += 360
    return angle


def angle_diff(a, b):
    d = abs(normalize_angle(a - b))
    return min(d, 360 - d)


results = []

for obj in doc.ModelSpace:

    if obj.ObjectName != "AcDbBlockReference":
        continue

    if obj.Name != BLOCK_NAME:
        continue

    ins = tuple(obj.InsertionPoint)
    rot = obj.Rotation
    scale = obj.XScaleFactor

    block_def = doc.Blocks.Item(BLOCK_NAME)

    lines = []

    for entity in block_def:

        if entity.Layer != "E Pipe":
            continue

        if entity.ObjectName != "AcDbLine":
            continue

        sp = tuple(entity.StartPoint)
        ep = tuple(entity.EndPoint)

        p1 = local_to_world(
            sp[0], sp[1], ins, rot, scale
        )

        p2 = local_to_world(
            ep[0], ep[1], ins, rot, scale
        )

        length = math.hypot(
            p2[0] - p1[0],
            p2[1] - p1[1]
        )

        angle = math.degrees(
            math.atan2(
                p2[1] - p1[1],
                p2[0] - p1[0]
            )
        )

        lines.append(
            {
                "p1": p1,
                "p2": p2,
                "length": length,
                "angle": angle,
            }
        )

    if not lines:
        continue

    # خطوط اصلی را بر اساس طول مرتب می‌کنیم
    lines.sort(key=lambda x: x["length"], reverse=True)

    # فعلاً بلندترین خط را به عنوان محور نماینده انتخاب می‌کنیم.
    main = lines[0]

    results.append(
        {
            "handle": obj.Handle,
            "insertion": ins,
            "rotation": math.degrees(rot),
            "scale": scale,
            "line_count": len(lines),
            "length": main["length"],
            "axis": normalize_angle(main["angle"]),
            "p1": main["p1"],
            "p2": main["p2"],
        }
    )


print()
print("DRAINAGE PIPE E ANALYSIS")
print("=" * 100)
print(f"Block: {BLOCK_NAME}")
print(f"Instances found: {len(results)}")

if results:

    print()
    print("STATISTICS")
    print("-" * 100)

    lengths = [r["length"] for r in results]
    axes = [r["axis"] for r in results]

    print(f"Length min : {min(lengths):.3f} m")
    print(f"Length max : {max(lengths):.3f} m")
    print(f"Length avg : {sum(lengths)/len(lengths):.3f} m")

    print()
    print("FIRST 20 DRAINAGES")
    print("-" * 100)

    for r in results[:20]:

        print(
            f"Handle={r['handle']}  "
            f"Pos=({r['insertion'][0]:.3f},{r['insertion'][1]:.3f})  "
            f"Axis={r['axis']:.3f}°  "
            f"Length={r['length']:.3f} m  "
            f"Lines={r['line_count']}  "
            f"P1=({r['p1'][0]:.3f},{r['p1'][1]:.3f})  "
            f"P2=({r['p2'][0]:.3f},{r['p2'][1]:.3f})"
        )
