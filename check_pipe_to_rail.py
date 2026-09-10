import win32com.client
import math

acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument

# Drainage Pipe E مورد آزمایش
drainage_handle = "7462A"

drainage = doc.HandleToObject(drainage_handle)

print("DRAINAGE:", drainage_handle)
print("Block:", drainage.Name)
print("Insertion:", tuple(drainage.InsertionPoint))
print("Rotation:", math.degrees(drainage.Rotation))
print("Scale:", drainage.XScaleFactor)

# ---------------------------------------------------------
# تبدیل نقطه Local Block به World
# ---------------------------------------------------------

ins = tuple(drainage.InsertionPoint)
rot = drainage.Rotation
scale = drainage.XScaleFactor

def local_to_world(x, y):
    x *= scale
    y *= scale

    wx = ins[0] + x * math.cos(rot) - y * math.sin(rot)
    wy = ins[1] + x * math.sin(rot) + y * math.cos(rot)

    return wx, wy


# ---------------------------------------------------------
# پیدا کردن خطوط اصلی E Pipe در Block Definition
# ---------------------------------------------------------

block_def = doc.Blocks.Item(drainage.Name)

pipe_lines = []

for obj in block_def:
    if obj.Layer == "E Pipe" and obj.ObjectName == "AcDbLine":
        sp = tuple(obj.StartPoint)
        ep = tuple(obj.EndPoint)

        p1 = local_to_world(sp[0], sp[1])
        p2 = local_to_world(ep[0], ep[1])

        pipe_lines.append((p1, p2))

print()
print("PIPE AXIS CANDIDATES")
print("=" * 90)

for i, (p1, p2) in enumerate(pipe_lines, 1):
    angle = math.degrees(
        math.atan2(p2[1] - p1[1], p2[0] - p1[0])
    )

    length = math.hypot(
        p2[0] - p1[0],
        p2[1] - p1[1]
    )

    print(
        f"{i}: "
        f"P1=({p1[0]:.3f},{p1[1]:.3f})  "
        f"P2=({p2[0]:.3f},{p2[1]:.3f})  "
        f"Length={length:.3f} m  "
        f"Axis={angle:.3f} deg"
    )


# ---------------------------------------------------------
# فاصله نقطه از Segment
# ---------------------------------------------------------

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


# ---------------------------------------------------------
# نزدیک‌ترین فاصله بین دو Segment
# ---------------------------------------------------------

def segment_distance(a, b, c, d):

    ax, ay = a
    bx, by = b
    cx, cy = c
    dx, dy = d

    # تقاطع دو Segment
    def cross(ax, ay, bx, by):
        return ax * by - ay * bx

    r_x = bx - ax
    r_y = by - ay
    s_x = dx - cx
    s_y = dy - cy

    denom = cross(r_x, r_y, s_x, s_y)

    if abs(denom) > 1e-12:
        qpx = cx - ax
        qpy = cy - ay

        t = cross(qpx, qpy, s_x, s_y) / denom
        u = cross(qpx, qpy, r_x, r_y) / denom

        if 0 <= t <= 1 and 0 <= u <= 1:
            return 0.0

    return min(
        point_to_segment(ax, ay, cx, cy, dx, dy),
        point_to_segment(bx, by, cx, cy, dx, dy),
        point_to_segment(cx, cy, ax, ay, bx, by),
        point_to_segment(dx, dy, ax, ay, bx, by),
    )


# ---------------------------------------------------------
# مقایسه با ریل‌ها
# ---------------------------------------------------------

print()
print("DISTANCE TO RAILS")
print("=" * 90)

for pipe_index, (p1, p2) in enumerate(pipe_lines, 1):

    results = []

    for obj in doc.ModelSpace:

        if obj.Layer != "Rail" or obj.ObjectName != "AcDbPolyline":
            continue

        coords = list(obj.Coordinates)

        best = float("inf")

        for i in range(0, len(coords) - 2, 2):

            r1 = (coords[i], coords[i + 1])
            r2 = (coords[i + 2], coords[i + 3])

            d = segment_distance(p1, p2, r1, r2)

            if d < best:
                best = d

        results.append((best, obj.Handle, obj.Length))

    results.sort()

    print()
    print(f"PIPE LINE {pipe_index}")

    for distance, handle, length in results[:4]:

        print(
            f"Rail={handle}  "
            f"Distance={distance:.3f} m  "
            f"RailLength={length:.3f} m"
        )
