import math

import pythoncom
import win32com.client
from win32com.client import VARIANT


CASES = [
    {"drainage": "74630", "sleeper": "68351"},
    {"drainage": "7467E", "sleeper": "68093"},
    {"drainage": "746F7", "sleeper": "67C53"},
]

DRAINAGE_BLOCK_NAME = "Drainage Pipe E"
SLEEPER_HALF_WIDTH_M = 0.175
PIPE_RADIUS_M = 0.0315
CLEARANCE_M = 0.100
TARGET_HALF_DISTANCE_M = SLEEPER_HALF_WIDTH_M + PIPE_RADIUS_M + CLEARANCE_M
TEXT_HEIGHT = 0.05


def point_variant(x, y, z=0.0):
    return VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        (float(x), float(y), float(z)),
    )


def add_point(ms, p):
    ms.AddPoint(point_variant(p[0], p[1], 0.0))


def add_line(ms, a, b):
    ms.AddLine(
        point_variant(a[0], a[1], 0.0),
        point_variant(b[0], b[1], 0.0),
    )


def add_text(ms, text, p):
    ms.AddText(
        text,
        point_variant(p[0] + 0.03, p[1] + 0.03, 0.0),
        TEXT_HEIGHT,
    )


def distance(a, b):
    return math.hypot(b[0] - a[0], b[1] - a[1])


def normalize(x, y):
    length = math.hypot(x, y)
    if length < 1e-12:
        raise RuntimeError("Zero-length vector.")
    return x / length, y / length


def local_to_world(x, y, insertion, rotation, sx=1.0, sy=1.0):
    x *= sx
    y *= sy
    c = math.cos(rotation)
    s = math.sin(rotation)
    return (
        insertion[0] + x * c - y * s,
        insertion[1] + x * s + y * c,
    )


def world_to_local(x, y, insertion, rotation):
    dx = x - insertion[0]
    dy = y - insertion[1]
    c = math.cos(rotation)
    s = math.sin(rotation)
    return (
        dx * c + dy * s,
        -dx * s + dy * c,
    )


def orientation(a, b):
    return math.degrees(math.atan2(b[1] - a[1], b[0] - a[0]))


def get_handle_object(doc, handle):
    return doc.HandleToObject(handle)


def get_drainage_centerline(doc):
    block = doc.Blocks.Item(DRAINAGE_BLOCK_NAME)
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

    unique = []
    for a, b in lines:
        duplicate = False
        for ua, ub in unique:
            direct = distance(a, ua) < 1e-6 and distance(b, ub) < 1e-6
            reverse = distance(a, ub) < 1e-6 and distance(b, ua) < 1e-6
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

            angle1 = orientation(a1, a2)
            angle2 = orientation(b1, b2)
            diff = abs((angle1 - angle2) % 180.0)
            diff = min(diff, 180.0 - diff)
            if diff > 5.0:
                continue

            sx = a2[0] - a1[0]
            sy = a2[1] - a1[1]
            length = math.hypot(sx, sy)
            if length < 1e-12:
                continue

            dx = b1[0] - a1[0]
            dy = b1[1] - a1[1]
            separation = abs(dx * sy - dy * sx) / length

            if best is None or separation > best[0]:
                best = (separation, (a1, a2), (b1, b2))

    if best is None:
        raise RuntimeError("Drainage centerline not found.")

    _, line1, line2 = best
    return (
        ((line1[0][0] + line2[0][0]) / 2.0,
         (line1[0][1] + line2[0][1]) / 2.0),
        ((line1[1][0] + line2[1][0]) / 2.0,
         (line1[1][1] + line2[1][1]) / 2.0),
    )


def extract_pipe(doc, handle, local_centerline):
    obj = get_handle_object(doc, handle)
    ins = tuple(obj.InsertionPoint)
    insertion = (float(ins[0]), float(ins[1]))
    rotation = float(obj.Rotation)
    sx = float(obj.XScaleFactor)
    sy = float(obj.YScaleFactor)

    lp1, lp2 = local_centerline
    p1 = local_to_world(lp1[0], lp1[1], insertion, rotation, sx, sy)
    p2 = local_to_world(lp2[0], lp2[1], insertion, rotation, sx, sy)

    center = ((p1[0] + p2[0]) / 2.0, (p1[1] + p2[1]) / 2.0)
    ux, uy = normalize(p2[0] - p1[0], p2[1] - p1[1])

    return {
        "handle": handle,
        "center": center,
        "p1": p1,
        "p2": p2,
        "axis_deg": orientation(p1, p2),
        "normal_a": (-uy, ux),
        "normal_b": (uy, -ux),
    }


def extract_sleeper(doc, handle):
    obj = get_handle_object(doc, handle)
    ins = tuple(obj.InsertionPoint)
    insertion = (float(ins[0]), float(ins[1]))
    rotation = float(obj.Rotation)

    return {
        "handle": handle,
        "block": str(obj.EffectiveName),
        "insertion": insertion,
        "rotation": rotation,
        "rotation_deg": math.degrees(rotation),
    }


def calculate_target(pipe, sleeper):
    px, py = pipe["center"]
    local_x, local_y = world_to_local(
        px,
        py,
        sleeper["insertion"],
        sleeper["rotation"],
    )

    if local_y >= 0.0:
        side_sign = 1.0
        side_name = "+Y"
    else:
        side_sign = -1.0
        side_name = "-Y"

    sr = sleeper["rotation"]
    sleeper_y_world = (-math.sin(sr), math.cos(sr))

    na = pipe["normal_a"]
    nb = pipe["normal_b"]
    dot_a = na[0] * sleeper_y_world[0] + na[1] * sleeper_y_world[1]
    dot_b = nb[0] * sleeper_y_world[0] + nb[1] * sleeper_y_world[1]

    if side_sign > 0:
        move_dir = na if dot_a >= dot_b else nb
    else:
        move_dir = na if dot_a <= dot_b else nb

    current_abs_y = abs(local_y)
    move_m = max(0.0, TARGET_HALF_DISTANCE_M - current_abs_y)

    target = (
        px + move_dir[0] * move_m,
        py + move_dir[1] * move_m,
    )

    return {
        "local_x": local_x,
        "local_y": local_y,
        "side": side_name,
        "move_dir": move_dir,
        "move_m": move_m,
        "target": target,
    }


def draw_result(ms, pipe, result):
    center = pipe["center"]
    target = result["target"]

    add_point(ms, center)
    add_point(ms, target)
    add_line(ms, center, target)
    add_text(ms, "TARGET", target)


def main():
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    print()
    print("=" * 95)
    print("THREE DRAINAGES - GEOMETRIC TARGET TEST")
    print("=" * 95)
    print(f"Target center distance = {TARGET_HALF_DISTANCE_M * 1000:.1f} mm")
    print("(175 mm sleeper half-width + 31.5 mm pipe radius + 100 mm clearance)")

    local_centerline = get_drainage_centerline(doc)

    for case_no, case in enumerate(CASES, start=1):
        print()
        print("=" * 95)
        print(f"CASE {case_no}: {case['drainage']} / {case['sleeper']}")

        try:
            pipe = extract_pipe(doc, case["drainage"], local_centerline)
            sleeper = extract_sleeper(doc, case["sleeper"])
            result = calculate_target(pipe, sleeper)
            draw_result(ms, pipe, result)

            print(f"Pipe center     = {pipe['center']}")
            print(f"Sleeper insert  = {sleeper['insertion']}")
            print(f"Sleeper axis    = {sleeper['rotation_deg']:.6f} deg")
            print(f"Pipe axis       = {pipe['axis_deg']:.6f} deg")
            print(f"Local Y         = {result['local_y']:.6f} m")
            print(f"Side            = {result['side']}")
            print(f"Move direction  = {result['move_dir']}")
            print(f"Move            = {result['move_m'] * 1000:.3f} mm")
            print(f"Target          = {result['target']}")

        except Exception as exc:
            print(f"ERROR: {type(exc).__name__}: {exc}")

    try:
        doc.Regen(1)
    except Exception:
        pass

    print()
    print("=" * 95)
    print("DONE")
    print("=" * 95)
    print("No object was moved.")
    print("Only TARGET points and center-to-target lines were created.")
    print("=" * 95)


if __name__ == "__main__":
    main()
