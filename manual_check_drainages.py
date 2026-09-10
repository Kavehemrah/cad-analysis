import math
import win32com.client


DRAINAGES = {"74630", "7467E", "746F7"}

SLEEPER_LAYER = "AG_Print"
SLEEPER_BLOCKS = {"AG_t", "AG_tttt"}
DRAINAGE_BLOCK = "Drainage Pipe E"


def local_to_world(x, y, insertion, rotation, sx=1.0, sy=1.0):

    x *= sx
    y *= sy

    c = math.cos(rotation)
    s = math.sin(rotation)

    return (
        insertion[0] + x * c - y * s,
        insertion[1] + x * s + y * c,
    )


def distance(a, b):
    return math.hypot(
        b[0] - a[0],
        b[1] - a[1],
    )


def orientation(a, b):
    return math.degrees(
        math.atan2(
            b[1] - a[1],
            b[0] - a[0],
        )
    )


def get_block_polyline(acad, block_name):

    block = acad.ActiveDocument.Blocks.Item(block_name)

    for ent in block:

        try:
            if ent.ObjectName != "AcDbPolyline":
                continue

            coords = tuple(ent.Coordinates)

            vertices = []

            for i in range(0, len(coords), 2):
                vertices.append(
                    (
                        float(coords[i]),
                        float(coords[i + 1]),
                    )
                )

            return vertices

        except Exception:
            continue

    return []


def get_drainage_centerline(acad):

    block = acad.ActiveDocument.Blocks.Item(
        DRAINAGE_BLOCK
    )

    lines = []

    for ent in block:

        try:

            if ent.ObjectName != "AcDbLine":
                continue

            if str(ent.Layer) != "E Pipe":
                continue

            p1 = tuple(ent.StartPoint)
            p2 = tuple(ent.EndPoint)

            p1 = (float(p1[0]), float(p1[1]))
            p2 = (float(p2[0]), float(p2[1]))

            if distance(p1, p2) > 1.0:
                lines.append((p1, p2))

        except Exception:
            continue

    # حذف duplicate
    unique = []

    for a, b in lines:

        duplicate = False

        for ua, ub in unique:

            direct = (
                distance(a, ua) < 1e-6
                and
                distance(b, ub) < 1e-6
            )

            reverse = (
                distance(a, ub) < 1e-6
                and
                distance(b, ua) < 1e-6
            )

            if direct or reverse:
                duplicate = True
                break

        if not duplicate:
            unique.append((a, b))

    best = None

    for i in range(len(unique)):

        for j in range(i + 1, len(unique)):

            a1, a2 = unique[i]
            b1, b2 = unique[j]

            ang1 = orientation(a1, a2)
            ang2 = orientation(b1, b2)

            d = abs((ang1 - ang2) % 180.0)
            d = min(d, 180.0 - d)

            if d > 5.0:
                continue

            # فاصله بین دو خط
            dx = b1[0] - a1[0]
            dy = b1[1] - a1[1]

            sx = a2[0] - a1[0]
            sy = a2[1] - a1[1]

            denom = math.hypot(sx, sy)

            if denom == 0:
                continue

            separation = abs(
                dx * sy - dy * sx
            ) / denom

            if best is None or separation > best[0]:
                best = (
                    separation,
                    (a1, a2),
                    (b1, b2),
                )

    if best is None:
        return None

    _, line1, line2 = best

    p1 = (
        (line1[0][0] + line2[0][0]) / 2,
        (line1[0][1] + line2[0][1]) / 2,
    )

    p2 = (
        (line1[1][0] + line2[1][0]) / 2,
        (line1[1][1] + line2[1][1]) / 2,
    )

    return p1, p2


def point_to_segment_distance(p, a, b):

    dx = b[0] - a[0]
    dy = b[1] - a[1]

    length_sq = dx * dx + dy * dy

    if length_sq == 0:
        return distance(p, a)

    t = (
        (p[0] - a[0]) * dx
        +
        (p[1] - a[1]) * dy
    ) / length_sq

    t = max(0.0, min(1.0, t))

    q = (
        a[0] + t * dx,
        a[1] + t * dy,
    )

    return distance(p, q)


def segment_distance(a, b, c, d):

    # ساده ولی دقیق برای این تست
    def cross(p, q, r):
        return (
            (q[0] - p[0]) * (r[1] - p[1])
            -
            (q[1] - p[1]) * (r[0] - p[0])
        )

    c1 = cross(a, b, c)
    c2 = cross(a, b, d)
    c3 = cross(c, d, a)
    c4 = cross(c, d, b)

    if (
        ((c1 > 0 and c2 < 0) or (c1 < 0 and c2 > 0))
        and
        ((c3 > 0 and c4 < 0) or (c3 < 0 and c4 > 0))
    ):
        return 0.0

    return min(
        point_to_segment_distance(a, c, d),
        point_to_segment_distance(b, c, d),
        point_to_segment_distance(c, a, b),
        point_to_segment_distance(d, a, b),
    )


def sleeper_clearance(pipe_a, pipe_b, polygon):

    min_d = float("inf")

    for i in range(len(polygon)):

        a = polygon[i]
        b = polygon[(i + 1) % len(polygon)]

        d = segment_distance(
            pipe_a,
            pipe_b,
            a,
            b,
        )

        min_d = min(min_d, d)

    # Ø63 mm
    return min_d - 0.0315


def main():

    acad = win32com.client.GetActiveObject(
        "AutoCAD.Application"
    )

    doc = acad.ActiveDocument

    sleeper_defs = {
        name: get_block_polyline(
            acad,
            name,
        )
        for name in SLEEPER_BLOCKS
    }

    sleepers = []

    for obj in doc.ModelSpace:

        try:

            if str(obj.Layer) != SLEEPER_LAYER:
                continue

            if obj.ObjectName != "AcDbBlockReference":
                continue

            name = str(obj.EffectiveName)

            if name not in SLEEPER_BLOCKS:
                continue

            ins = tuple(obj.InsertionPoint)

            verts = sleeper_defs[name]

            world = [
                local_to_world(
                    x,
                    y,
                    ins,
                    float(obj.Rotation),
                    float(obj.XScaleFactor),
                    float(obj.YScaleFactor),
                )
                for x, y in verts
            ]

            sleepers.append({
                "handle": str(obj.Handle),
                "insertion": (
                    float(ins[0]),
                    float(ins[1]),
                ),
                "rotation_deg": math.degrees(
                    float(obj.Rotation)
                ),
                "polygon": world,
            })

        except Exception:
            continue

    centerline = get_drainage_centerline(acad)

    drainages = {}

    for obj in doc.ModelSpace:

        try:

            if obj.ObjectName != "AcDbBlockReference":
                continue

            if str(obj.EffectiveName) != DRAINAGE_BLOCK:
                continue

            handle = str(obj.Handle)

            if handle not in DRAINAGES:
                continue

            ins = tuple(obj.InsertionPoint)

            a, b = centerline

            wa = local_to_world(
                a[0],
                a[1],
                ins,
                float(obj.Rotation),
                float(obj.XScaleFactor),
                float(obj.YScaleFactor),
            )

            wb = local_to_world(
                b[0],
                b[1],
                ins,
                float(obj.Rotation),
                float(obj.XScaleFactor),
                float(obj.YScaleFactor),
            )

            drainages[handle] = {
                "insertion": (
                    float(ins[0]),
                    float(ins[1]),
                ),
                "a": wa,
                "b": wb,
                "axis": orientation(wa, wb),
                "rotation": math.degrees(
                    float(obj.Rotation)
                ),
            }

        except Exception:
            continue

    print("=" * 110)
    print("MANUAL CHECK DATA")
    print("=" * 110)

    for handle in DRAINAGES:

        d = drainages.get(handle)

        if d is None:
            print(f"\n{handle}: NOT FOUND")
            continue

        # تمام تراورس‌های اطراف
        candidates = []

        center = (
            (d["a"][0] + d["b"][0]) / 2,
            (d["a"][1] + d["b"][1]) / 2,
        )

        for s in sleepers:

            if distance(
                center,
                s["insertion"],
            ) > 1.5:
                continue

            clearance = sleeper_clearance(
                d["a"],
                d["b"],
                s["polygon"],
            )

            candidates.append(
                (clearance, s)
            )

        candidates.sort(
            key=lambda x: x[0]
        )

        print("\n" + "-" * 110)

        print(
            f"DRAINAGE = {handle}"
        )

        print(
            f"Insertion     = "
            f"({d['insertion'][0]:.6f}, "
            f"{d['insertion'][1]:.6f})"
        )

        print(
            f"Centerline P1 = "
            f"({d['a'][0]:.6f}, "
            f"{d['a'][1]:.6f})"
        )

        print(
            f"Centerline P2 = "
            f"({d['b'][0]:.6f}, "
            f"{d['b'][1]:.6f})"
        )

        print(
            f"Pipe Axis     = "
            f"{d['axis']:.6f} deg"
        )

        print(
            f"Pipe Rotation = "
            f"{d['rotation']:.6f} deg"
        )

        print("\nNearest Sleepers:")

        for clearance, s in candidates[:5]:

            print(
                f"  Sleeper={s['handle']} "
                f"| Pos=({s['insertion'][0]:.6f}, "
                f"{s['insertion'][1]:.6f}) "
                f"| Rot={s['rotation_deg']:.6f}° "
                f"| Clearance={clearance * 1000:.2f} mm"
            )


if __name__ == "__main__":
    main()
