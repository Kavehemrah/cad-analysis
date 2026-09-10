import math
import win32com.client


DRAINAGE_BLOCK = "Drainage Pipe E"
SLEEPER_LAYER = "AG_Print"
SLEEPER_BLOCKS = {"AG_t", "AG_tttt"}


def local_to_world(x, y, insertion, rotation, sx=1.0, sy=1.0):
    x *= sx
    y *= sy

    c = math.cos(rotation)
    s = math.sin(rotation)

    return (
        insertion[0] + x * c - y * s,
        insertion[1] + x * s + y * c,
    )


def orientation(a, b):
    return math.degrees(
        math.atan2(
            b[1] - a[1],
            b[0] - a[0]
        )
    )


def angle_diff(a, b):
    d = abs((a - b) % 180.0)
    return min(d, 180.0 - d)


def point_to_segment_distance(px, py, ax, ay, bx, by):
    dx = bx - ax
    dy = by - ay

    length_sq = dx * dx + dy * dy

    if length_sq == 0:
        return math.hypot(px - ax, py - ay)

    t = (
        (px - ax) * dx +
        (py - ay) * dy
    ) / length_sq

    t = max(0.0, min(1.0, t))

    qx = ax + t * dx
    qy = ay + t * dy

    return math.hypot(
        px - qx,
        py - qy
    )


def segments_intersect(a, b, c, d, tol=1e-9):

    def cross(p, q, r):
        return (
            (q[0] - p[0]) * (r[1] - p[1])
            -
            (q[1] - p[1]) * (r[0] - p[0])
        )

    def on_segment(p, q, r):
        return (
            min(p[0], r[0]) - tol <= q[0] <= max(p[0], r[0]) + tol
            and
            min(p[1], r[1]) - tol <= q[1] <= max(p[1], r[1]) + tol
        )

    c1 = cross(a, b, c)
    c2 = cross(a, b, d)
    c3 = cross(c, d, a)
    c4 = cross(c, d, b)

    if (
        ((c1 > tol and c2 < -tol) or (c1 < -tol and c2 > tol))
        and
        ((c3 > tol and c4 < -tol) or (c3 < -tol and c4 > tol))
    ):
        return True

    if abs(c1) <= tol and on_segment(a, c, b):
        return True

    if abs(c2) <= tol and on_segment(a, d, b):
        return True

    if abs(c3) <= tol and on_segment(c, a, d):
        return True

    if abs(c4) <= tol and on_segment(c, b, d):
        return True

    return False


def segment_to_segment_distance(a, b, c, d):

    if segments_intersect(a, b, c, d):
        return 0.0

    return min(
        point_to_segment_distance(
            a[0], a[1],
            c[0], c[1],
            d[0], d[1],
        ),
        point_to_segment_distance(
            b[0], b[1],
            c[0], c[1],
            d[0], d[1],
        ),
        point_to_segment_distance(
            c[0], c[1],
            a[0], a[1],
            b[0], b[1],
        ),
        point_to_segment_distance(
            d[0], d[1],
            a[0], a[1],
            b[0], b[1],
        ),
    )


def get_block_vertices(acad, block_name):

    block_def = acad.ActiveDocument.Blocks.Item(block_name)

    for ent in block_def:

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


def get_drainage_centerline(acad, block_name):

    block_def = acad.ActiveDocument.Blocks.Item(block_name)

    lines = []

    for ent in block_def:

        try:
            if ent.ObjectName != "AcDbLine":
                continue

            if str(ent.Layer) != "E Pipe":
                continue

            p1 = tuple(ent.StartPoint)
            p2 = tuple(ent.EndPoint)

            p1 = (
                float(p1[0]),
                float(p1[1]),
            )

            p2 = (
                float(p2[0]),
                float(p2[1]),
            )

            length = math.hypot(
                p2[0] - p1[0],
                p2[1] - p1[1],
            )

            if length > 1.0:
                lines.append((p1, p2))

        except Exception:
            continue

    # حذف خطوط تکراری
    unique = []

    for line in lines:

        a, b = line
        duplicate = False

        for ua, ub in unique:

            direct = (
                math.hypot(
                    a[0] - ua[0],
                    a[1] - ua[1]
                ) < 1e-6
                and
                math.hypot(
                    b[0] - ub[0],
                    b[1] - ub[1]
                ) < 1e-6
            )

            reverse = (
                math.hypot(
                    a[0] - ub[0],
                    a[1] - ub[1]
                ) < 1e-6
                and
                math.hypot(
                    b[0] - ua[0],
                    b[1] - ua[1]
                ) < 1e-6
            )

            if direct or reverse:
                duplicate = True
                break

        if not duplicate:
            unique.append(line)

    if len(unique) < 2:
        return None

    best = None

    for i in range(len(unique)):

        for j in range(i + 1, len(unique)):

            a1, a2 = unique[i]
            b1, b2 = unique[j]

            angle1 = orientation(a1, a2)
            angle2 = orientation(b1, b2)

            if angle_diff(angle1, angle2) > 5.0:
                continue

            separation = point_to_segment_distance(
                b1[0],
                b1[1],
                a1[0],
                a1[1],
                a2[0],
                a2[1],
            )

            if best is None or separation > best[0]:

                best = (
                    separation,
                    unique[i],
                    unique[j],
                )

    if best is None:
        return None

    _, line1, line2 = best

    centerline_p1 = (
        (line1[0][0] + line2[0][0]) / 2.0,
        (line1[0][1] + line2[0][1]) / 2.0,
    )

    centerline_p2 = (
        (line1[1][0] + line2[1][0]) / 2.0,
        (line1[1][1] + line2[1][1]) / 2.0,
    )

    return (
        centerline_p1,
        centerline_p2,
    )


def build_sleepers(acad):

    sleepers = []

    for obj in acad.ActiveDocument.ModelSpace:

        try:

            if str(obj.Layer) != SLEEPER_LAYER:
                continue

            if obj.ObjectName != "AcDbBlockReference":
                continue

            try:
                block_name = str(obj.EffectiveName)
            except Exception:
                block_name = str(obj.Name)

            if block_name not in SLEEPER_BLOCKS:
                continue

            insertion = tuple(obj.InsertionPoint)

            sleepers.append({
                "handle": str(obj.Handle),
                "block": block_name,
                "insertion": (
                    float(insertion[0]),
                    float(insertion[1]),
                    float(insertion[2]),
                ),
                "rotation": float(obj.Rotation),
                "sx": float(obj.XScaleFactor),
                "sy": float(obj.YScaleFactor),
            })

        except Exception:
            continue

    return sleepers


def build_drainages(acad):

    drainages = []

    for obj in acad.ActiveDocument.ModelSpace:

        try:

            if obj.ObjectName != "AcDbBlockReference":
                continue

            try:
                block_name = str(obj.EffectiveName)
            except Exception:
                block_name = str(obj.Name)

            if block_name != DRAINAGE_BLOCK:
                continue

            insertion = tuple(obj.InsertionPoint)

            drainages.append({
                "handle": str(obj.Handle),
                "block": block_name,
                "insertion": (
                    float(insertion[0]),
                    float(insertion[1]),
                    float(insertion[2]),
                ),
                "rotation": float(obj.Rotation),
                "sx": float(obj.XScaleFactor),
                "sy": float(obj.YScaleFactor),
            })

        except Exception:
            continue

    return drainages


def transform_sleeper_vertices(vertices, sleeper):

    return [
        local_to_world(
            x,
            y,
            sleeper["insertion"],
            sleeper["rotation"],
            sleeper["sx"],
            sleeper["sy"],
        )
        for x, y in vertices
    ]


def make_sleeper_segments(vertices):

    segments = []

    for i in range(len(vertices) - 1):

        segments.append(
            (
                vertices[i],
                vertices[i + 1],
            )
        )

    # اگر پلی‌لاین صراحتاً بسته نشده باشد،
    # فقط در صورتی آخرین-اولین را اضافه می‌کنیم
    # که دو سر عملاً یک نقطه نباشند.
    if len(vertices) > 2:

        if math.hypot(
            vertices[-1][0] - vertices[0][0],
            vertices[-1][1] - vertices[0][1],
        ) > 1e-9:

            segments.append(
                (
                    vertices[-1],
                    vertices[0],
                )
            )

    return segments


def analyze():

    acad = win32com.client.GetActiveObject(
        "AutoCAD.Application"
    )

    print("=" * 100)
    print("DRAINAGE vs SLEEPER - GEOMETRIC ANALYSIS")
    print("=" * 100)

    # -------------------------
    # Sleeper definitions
    # -------------------------

    sleeper_geometry = {}

    print("\nReading sleeper definitions...")

    for block_name in sorted(SLEEPER_BLOCKS):

        vertices = get_block_vertices(
            acad,
            block_name,
        )

        sleeper_geometry[block_name] = vertices

        print(
            f"{block_name}: "
            f"{len(vertices)} vertices"
        )

    sleepers = build_sleepers(acad)

    print(
        f"Sleepers found: "
        f"{len(sleepers)}"
    )

    # -------------------------
    # Drainage
    # -------------------------

    drainages = build_drainages(acad)

    print(
        f"Drainage '{DRAINAGE_BLOCK}': "
        f"{len(drainages)}"
    )

    centerline = get_drainage_centerline(
        acad,
        DRAINAGE_BLOCK,
    )

    if centerline is None:

        print(
            "\nERROR: "
            "Drainage centerline not found."
        )

        return

    print(
        "Drainage centerline: OK"
    )

    # -------------------------
    # Analyze
    # -------------------------

    results = []

    for idx, drainage in enumerate(
        drainages,
        start=1
    ):

        local_p1, local_p2 = centerline

        cp1 = local_to_world(
            local_p1[0],
            local_p1[1],
            drainage["insertion"],
            drainage["rotation"],
            drainage["sx"],
            drainage["sy"],
        )

        cp2 = local_to_world(
            local_p2[0],
            local_p2[1],
            drainage["insertion"],
            drainage["rotation"],
            drainage["sx"],
            drainage["sy"],
        )

        drainage_axis = orientation(
            cp1,
            cp2,
        )

        best = None

        for sleeper in sleepers:

            dx = (
                drainage["insertion"][0]
                - sleeper["insertion"][0]
            )

            dy = (
                drainage["insertion"][1]
                - sleeper["insertion"][1]
            )

            center_dist = math.hypot(
                dx,
                dy,
            )

            if center_dist > 5.0:
                continue

            local_vertices = sleeper_geometry[
                sleeper["block"]
            ]

            world_vertices = transform_sleeper_vertices(
                local_vertices,
                sleeper,
            )

            sleeper_segments = make_sleeper_segments(
                world_vertices
            )

            # مهم:
            # محور تراورس = Block Rotation
            sleeper_axis = math.degrees(
                sleeper["rotation"]
            )

            min_distance = float("inf")
            intersects = False

            for a, b in sleeper_segments:

                distance = segment_to_segment_distance(
                    cp1,
                    cp2,
                    a,
                    b,
                )

                if distance < min_distance:
                    min_distance = distance

                if distance <= 1e-9:
                    intersects = True
                    break

            candidate = {
                "drainage": drainage["handle"],
                "sleeper": sleeper["handle"],
                "distance": min_distance,
                "intersection": intersects,
                "drainage_axis": drainage_axis,
                "sleeper_axis": sleeper_axis,
                "angle_diff": angle_diff(
                    drainage_axis,
                    sleeper_axis,
                ),
            }

            if (
                best is None
                or candidate["distance"]
                < best["distance"]
            ):
                best = candidate

        if best is not None:
            results.append(best)

        if idx % 25 == 0:
            print(
                f"Processed: "
                f"{idx}/{len(drainages)}"
            )

    # -------------------------
    # Sort
    # -------------------------

    results.sort(
        key=lambda x: x["distance"]
    )

    print("\n" + "=" * 100)
    print("RESULTS - 100 CLOSEST")
    print("=" * 100)

    for r in results[:100]:

        status = (
            "INTERSECTION"
            if r["intersection"]
            else "NEAR"
        )

        print(
            f"Drainage={r['drainage']}  "
            f"Sleeper={r['sleeper']}  "
            f"Distance={r['distance']:.3f} m  "
            f"DrainageAxis={r['drainage_axis']:.3f}°  "
            f"SleeperAxis={r['sleeper_axis']:.3f}°  "
            f"AngleDiff={r['angle_diff']:.3f}°  "
            f"{status}"
        )

    # -------------------------
    # Summary
    # -------------------------

    intersections = sum(
        1
        for r in results
        if r["intersection"]
    )

    print("\n" + "=" * 100)
    print("SUMMARY")
    print("=" * 100)

    print(
        f"Drainages analyzed : "
        f"{len(results)}"
    )

    print(
        f"Intersections      : "
        f"{intersections}"
    )

    print(
        f"Minimum distance   : "
        f"{min(r['distance'] for r in results):.3f} m"
    )

    print(
        f"Maximum distance   : "
        f"{max(r['distance'] for r in results):.3f} m"
    )


if __name__ == "__main__":
    analyze()
