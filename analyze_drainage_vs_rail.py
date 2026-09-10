import win32com.client
import math

acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

DRAINAGE_BLOCK = "Drainage Pipe E"


def local_to_world(x, y, ins, rot, scale):
    x *= scale
    y *= scale
    return (
        ins[0] + x * math.cos(rot) - y * math.sin(rot),
        ins[1] + x * math.sin(rot) + y * math.cos(rot),
    )


def normalize_angle(angle):
    while angle >= 180:
        angle -= 360
    while angle < -180:
        angle += 360
    return angle


def angle_difference(a, b):
    d = abs(normalize_angle(a - b))
    return min(d, 180 - d)


def point_to_segment(px, py, ax, ay, bx, by):
    dx = bx - ax
    dy = by - ay
    length2 = dx * dx + dy * dy

    if length2 == 0:
        return math.hypot(px - ax, py - ay)

    t = ((px - ax) * dx + (py - ay) * dy) / length2
    t = max(0, min(1, t))

    qx = ax + t * dx
    qy = ay + t * dy

    return math.hypot(px - qx, py - qy)


def segment_distance(a, b, c, d):
    ax, ay = a
    bx, by = b
    cx, cy = c
    dx, dy = d

    def cross(x1, y1, x2, y2):
        return x1 * y2 - y1 * x2

    rx = bx - ax
    ry = by - ay
    sx = dx - cx
    sy = dy - cy

    denominator = cross(rx, ry, sx, sy)

    if abs(denominator) > 1e-12:
        qx = cx - ax
        qy = cy - ay

        t = cross(qx, qy, sx, sy) / denominator
        u = cross(qx, qy, rx, ry) / denominator

        if 0 <= t <= 1 and 0 <= u <= 1:
            return 0.0

    return min(
        point_to_segment(ax, ay, cx, cy, dx, dy),
        point_to_segment(bx, by, cx, cy, dx, dy),
        point_to_segment(cx, cy, ax, ay, bx, by),
        point_to_segment(dx, dy, ax, ay, bx, by),
    )


def get_drainage_centerline(obj):
    ins = tuple(obj.InsertionPoint)
    rot = obj.Rotation
    scale = obj.XScaleFactor

    block = doc.Blocks.Item(DRAINAGE_BLOCK)

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

    # حذف خطوط تکراری
    unique = []

    for line in lines:

        p1, p2, length = line
        duplicate = False

        for old in unique:

            a1, a2, _ = old

            d1 = math.hypot(
                p1[0] - a1[0],
                p1[1] - a1[1]
            )

            d2 = math.hypot(
                p2[0] - a2[0],
                p2[1] - a2[1]
            )

            d3 = math.hypot(
                p1[0] - a2[0],
                p1[1] - a2[1]
            )

            d4 = math.hypot(
                p2[0] - a1[0],
                p2[1] - a1[1]
            )

            if (
                (d1 < 0.001 and d2 < 0.001)
                or
                (d3 < 0.001 and d4 < 0.001)
            ):
                duplicate = True
                break

        if not duplicate:
            unique.append(line)

    if len(unique) < 2:
        return None

    # دو خط موازی با بیشترین فاصله را پیدا می‌کنیم
    best = None

    for i in range(len(unique)):
        for j in range(i + 1, len(unique)):

            a1, a2, _ = unique[i]
            b1, b2, _ = unique[j]

            angle_a = math.degrees(
                math.atan2(
                    a2[1] - a1[1],
                    a2[0] - a1[0]
                )
            )

            angle_b = math.degrees(
                math.atan2(
                    b2[1] - b1[1],
                    b2[0] - b1[0]
                )
            )

            if angle_difference(angle_a, angle_b) > 2:
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
                best = (separation, unique[i], unique[j])

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

    axis = math.degrees(
        math.atan2(
            p2[1] - p1[1],
            p2[0] - p1[0]
        )
    )

    return p1, p2, length, axis


# ---------------------------------------------------------
# جمع ریل‌ها
# ---------------------------------------------------------

rails = []

for obj in doc.ModelSpace:

    if obj.Layer != "Rail":
        continue

    if obj.ObjectName != "AcDbPolyline":
        continue

    coords = list(obj.Coordinates)

    if len(coords) < 4:
        continue

    segments = []

    min_x = min(coords[0::2])
    max_x = max(coords[0::2])
    min_y = min(coords[1::2])
    max_y = max(coords[1::2])

    for i in range(0, len(coords) - 2, 2):

        p1 = (coords[i], coords[i + 1])
        p2 = (coords[i + 2], coords[i + 3])

        angle = math.degrees(
            math.atan2(
                p2[1] - p1[1],
                p2[0] - p1[0]
            )
        )

        segments.append((p1, p2, angle))

    rails.append(
        {
            "handle": obj.Handle,
            "segments": segments,
            "bbox": (min_x, max_x, min_y, max_y),
            "length": obj.Length,
        }
    )


print()
print("RAIL DATA")
print("=" * 100)
print(f"Rail polylines: {len(rails)}")


# ---------------------------------------------------------
# جمع Drainageها
# ---------------------------------------------------------

drainages = []

for obj in doc.ModelSpace:

    if obj.ObjectName != "AcDbBlockReference":
        continue

    if obj.Name != DRAINAGE_BLOCK:
        continue

    centerline = get_drainage_centerline(obj)

    if centerline:

        p1, p2, length, axis = centerline

        drainages.append(
            {
                "handle": obj.Handle,
                "p1": p1,
                "p2": p2,
                "length": length,
                "axis": axis,
            }
        )


print(f"Drainage objects: {len(drainages)}")


# ---------------------------------------------------------
# تحلیل
# ---------------------------------------------------------

results = []

for d in drainages:

    p1 = d["p1"]
    p2 = d["p2"]

    min_x = min(p1[0], p2[0]) - 20
    max_x = max(p1[0], p2[0]) + 20
    min_y = min(p1[1], p2[1]) - 20
    max_y = max(p1[1], p2[1]) + 20

    candidates = []

    for rail in rails:

        rminx, rmaxx, rminy, rmaxy = rail["bbox"]

        if rmaxx < min_x or rminx > max_x:
            continue

        if rmaxy < min_y or rminy > max_y:
            continue

        candidates.append(rail)

    rail_results = []

    for rail in candidates:

        best_distance = float("inf")
        best_angle = None

        for r1, r2, rail_angle in rail["segments"]:

            distance = segment_distance(
                p1,
                p2,
                r1,
                r2
            )

            if distance < best_distance:

                best_distance = distance
                best_angle = rail_angle

        if best_angle is not None:

            rail_results.append(
                (
                    best_distance,
                    rail["handle"],
                    best_angle,
                )
            )

    rail_results.sort()

    if rail_results:

        nearest = rail_results[:4]

        results.append(
            {
                "drainage": d,
                "rails": nearest,
            }
        )


# ---------------------------------------------------------
# نمایش موارد نزدیک / متقاطع
# ---------------------------------------------------------

print()
print("NEAREST RAILS PER DRAINAGE")
print("=" * 120)

for item in results:

    d = item["drainage"]
    nearest = item["rails"]

    # فقط مواردی که حداقل یکی از ریل‌ها حداکثر 3 متر فاصله دارد
    if nearest[0][0] > 3:
        continue

    print()
    print(
        f"Drainage={d['handle']}  "
        f"Axis={d['axis']:.3f}°"
    )

    for distance, rail_handle, rail_axis in nearest:

        diff = angle_difference(
            d["axis"],
            rail_axis
        )

        print(
            f"    Rail={rail_handle}  "
            f"Distance={distance:.3f} m  "
            f"RailAxis={rail_axis:.3f}°  "
            f"AngleDiff={diff:.3f}°"
        )
