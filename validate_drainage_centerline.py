import win32com.client
import math

acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

BLOCK_NAME = "Drainage Pipe E"


def local_to_world(x, y, ins, rot, scale):
    x *= scale
    y *= scale

    return (
        ins[0] + x * math.cos(rot) - y * math.sin(rot),
        ins[1] + x * math.sin(rot) + y * math.cos(rot),
    )


def get_centerline(obj):
    ins = tuple(obj.InsertionPoint)
    rot = obj.Rotation
    scale = obj.XScaleFactor

    block = doc.Blocks.Item(BLOCK_NAME)

    lines = []

    for e in block:
        if e.ObjectName != "AcDbLine":
            continue
        if e.Layer != "E Pipe":
            continue

        sp = tuple(e.StartPoint)
        ep = tuple(e.EndPoint)

        p1 = local_to_world(sp[0], sp[1], ins, rot, scale)
        p2 = local_to_world(ep[0], ep[1], ins, rot, scale)

        length = math.hypot(
            p2[0] - p1[0],
            p2[1] - p1[1]
        )

        lines.append((p1, p2, length))

    if len(lines) < 2:
        return None

    # حذف خطوط کاملاً تکراری
    unique = []

    for line in lines:
        p1, p2, length = line

        duplicate = False

        for old in unique:
            a1, a2, _ = old

            d1 = math.hypot(p1[0] - a1[0], p1[1] - a1[1])
            d2 = math.hypot(p2[0] - a2[0], p2[1] - a2[1])

            d3 = math.hypot(p1[0] - a2[0], p1[1] - a2[1])
            d4 = math.hypot(p2[0] - a1[0], p2[1] - a1[1])

            if (d1 < 0.001 and d2 < 0.001) or (
                d3 < 0.001 and d4 < 0.001
            ):
                duplicate = True
                break

        if not duplicate:
            unique.append(line)

    lines = unique

    # پیدا کردن دو خط تقریباً موازی با بیشترین فاصله عرضی
    best = None

    for i in range(len(lines)):
        for j in range(i + 1, len(lines)):

            a1, a2, _ = lines[i]
            b1, b2, _ = lines[j]

            angle_a = math.atan2(
                a2[1] - a1[1],
                a2[0] - a1[0]
            )

            angle_b = math.atan2(
                b2[1] - b1[1],
                b2[0] - b1[0]
            )

            diff = abs(
                math.degrees(angle_a - angle_b)
            ) % 180

            if diff > 90:
                diff = 180 - diff

            # فقط خطوط تقریباً موازی
            if diff > 2:
                continue

            ma = (
                (a1[0] + a2[0]) / 2,
                (a1[1] + a2[1]) / 2,
            )

            mb = (
                (b1[0] + b2[0]) / 2,
                (b1[1] + b2[1]) / 2,
            )

            separation = math.hypot(
                ma[0] - mb[0],
                ma[1] - mb[1]
            )

            if best is None or separation > best[0]:
                best = (separation, lines[i], lines[j])

    if best is None:
        return None

    separation, line_a, line_b = best

    a1, a2, _ = line_a
    b1, b2, _ = line_b

    p1 = (
        (a1[0] + b1[0]) / 2,
        (a1[1] + b1[1]) / 2,
    )

    p2 = (
        (a2[0] + b2[0]) / 2,
        (a2[1] + b2[1]) / 2,
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

    return p1, p2, length, angle, separation, len(lines)


# ---------------------------------------------------------
# جمع Drainageها
# ---------------------------------------------------------

drainages = []

for obj in doc.ModelSpace:

    if obj.ObjectName != "AcDbBlockReference":
        continue

    if obj.Name != BLOCK_NAME:
        continue

    centerline = get_centerline(obj)

    if centerline:
        drainages.append(
            (
                obj.Handle,
                tuple(obj.InsertionPoint),
                centerline
            )
        )


print()
print("DRAINAGE CENTERLINE VALIDATION")
print("=" * 110)
print(f"Total Drainage Pipe E: {len(drainages)}")


# ---------------------------------------------------------
# انتخاب ۱۰ نمونه از ابتدا تا انتها
# ---------------------------------------------------------

count = len(drainages)

indices = sorted(
    set(
        [
            0,
            1,
            2,
            3,
            count // 4,
            count // 4 + 1,
            count // 2,
            count // 2 + 1,
            3 * count // 4,
            count - 1,
        ]
    )
)


print()
print("SELECTED SAMPLES")
print("-" * 110)

for index in indices:

    handle, insertion, data = drainages[index]

    p1, p2, length, angle, separation, line_count = data

    print(
        f"[{index + 1:3d}] "
        f"Handle={handle}  "
        f"Pos=({insertion[0]:.2f},{insertion[1]:.2f})  "
        f"Length={length:.4f} m  "
        f"Axis={angle:.3f}°  "
        f"Width={separation:.4f} m  "
        f"UniqueLines={line_count}"
    )
