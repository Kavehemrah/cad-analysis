import math
import time

import pythoncom
import win32com.client
from win32com.client import VARIANT


# ============================================================
# TEST CASES
# First 10 MOVE cases requested for live test.
# ============================================================

CASES = [
    {"drainage": "74630", "sleeper": "68351"},
    {"drainage": "7467E", "sleeper": "68093"},
    {"drainage": "746BB", "sleeper": "67E6F"},
    {"drainage": "746BC", "sleeper": "67E66"},
    {"drainage": "746BD", "sleeper": "67E5D"},
    {"drainage": "746BE", "sleeper": "67E54"},
    {"drainage": "746BF", "sleeper": "67E4B"},
    {"drainage": "746C0", "sleeper": "67E42"},
    {"drainage": "746C1", "sleeper": "67E39"},
    {"drainage": "746C2", "sleeper": "67E30"},
]


# ============================================================
# GEOMETRY / PROJECT CONSTANTS
# ============================================================

DRAINAGE_BLOCK_NAME = "Drainage Pipe E"

SLEEPER_HALF_WIDTH_M = 0.175       # 350 mm / 2
PIPE_RADIUS_M = 0.0315             # 63 mm / 2
CLEARANCE_M = 0.100                # required final clearance

TARGET_HALF_DISTANCE_M = (
    SLEEPER_HALF_WIDTH_M
    + PIPE_RADIUS_M
    + CLEARANCE_M
)

# Debug marker settings
MARKER_TEXT_HEIGHT = 0.05
MARKER_POINT_SIZE = 0.08

# COM retry settings
COM_RETRIES = 12
COM_DELAY = 0.20


# ============================================================
# AUTOCAD HELPERS
# ============================================================

def point_variant(x, y, z=0.0):
    return VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        (float(x), float(y), float(z)),
    )


def wait_for_autocad(acad):
    try:
        state = acad.GetAcadState()
    except Exception:
        return

    for _ in range(COM_RETRIES):
        try:
            if state.IsQuiescent:
                return
        except Exception:
            return

        pythoncom.PumpWaitingMessages()
        time.sleep(COM_DELAY)


def is_busy_error(exc):
    text = str(exc)
    return (
        "Call was rejected by callee" in text
        or "RPC_E_CALL_REJECTED" in text
        or "-2147418111" in text
    )


def call_with_retry(func):
    last_error = None

    for _ in range(COM_RETRIES):
        try:
            return func()
        except Exception as exc:
            last_error = exc

            if not is_busy_error(exc):
                raise

            pythoncom.PumpWaitingMessages()
            time.sleep(COM_DELAY)

    raise last_error


def add_point(ms, p):
    call_with_retry(
        lambda: ms.AddPoint(
            point_variant(
                p[0],
                p[1],
                0.0,
            )
        )
    )


def add_line(ms, a, b):
    call_with_retry(
        lambda: ms.AddLine(
            point_variant(
                a[0],
                a[1],
                0.0,
            ),
            point_variant(
                b[0],
                b[1],
                0.0,
            ),
        )
    )


def add_text(ms, text, p):
    call_with_retry(
        lambda: ms.AddText(
            str(text),
            point_variant(
                p[0] + 0.03,
                p[1] + 0.03,
                0.0,
            ),
            MARKER_TEXT_HEIGHT,
        )
    )


# ============================================================
# BASIC GEOMETRY
# ============================================================

def distance(a, b):
    return math.hypot(
        b[0] - a[0],
        b[1] - a[1],
    )


def normalize(x, y):
    length = math.hypot(x, y)

    if length < 1e-12:
        raise RuntimeError("Zero-length vector.")

    return (
        x / length,
        y / length,
    )


def orientation(a, b):
    return math.degrees(
        math.atan2(
            b[1] - a[1],
            b[0] - a[0],
        )
    )


def local_to_world(
    x,
    y,
    insertion,
    rotation,
    sx=1.0,
    sy=1.0,
):
    x *= sx
    y *= sy

    c = math.cos(rotation)
    s = math.sin(rotation)

    return (
        insertion[0]
        + x * c
        - y * s,

        insertion[1]
        + x * s
        + y * c,
    )


def world_to_local(
    x,
    y,
    insertion,
    rotation,
):
    dx = x - insertion[0]
    dy = y - insertion[1]

    c = math.cos(rotation)
    s = math.sin(rotation)

    return (
        dx * c + dy * s,
        -dx * s + dy * c,
    )


# ============================================================
# HANDLE
# ============================================================

def get_handle_object(doc, handle):
    return call_with_retry(
        lambda: doc.HandleToObject(handle)
    )


# ============================================================
# PIPE CENTERLINE
# ============================================================

def get_drainage_centerline(doc):
    block = doc.Blocks.Item(
        DRAINAGE_BLOCK_NAME
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

            p1 = (
                float(p1[0]),
                float(p1[1]),
            )

            p2 = (
                float(p2[0]),
                float(p2[1]),
            )

            if distance(p1, p2) > 1.0:
                lines.append((p1, p2))

        except Exception:
            continue

    unique = []

    for a, b in lines:
        duplicate = False

        for ua, ub in unique:
            direct = (
                distance(a, ua) < 1e-6
                and distance(b, ub) < 1e-6
            )

            reverse = (
                distance(a, ub) < 1e-6
                and distance(b, ua) < 1e-6
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

            angle1 = orientation(a1, a2)
            angle2 = orientation(b1, b2)

            diff = abs(
                (angle1 - angle2) % 180.0
            )

            diff = min(
                diff,
                180.0 - diff,
            )

            if diff > 5.0:
                continue

            sx = a2[0] - a1[0]
            sy = a2[1] - a1[1]

            length = math.hypot(
                sx,
                sy,
            )

            if length < 1e-12:
                continue

            dx = b1[0] - a1[0]
            dy = b1[1] - a1[1]

            separation = abs(
                dx * sy
                - dy * sx
            ) / length

            if (
                best is None
                or separation > best[0]
            ):
                best = (
                    separation,
                    (a1, a2),
                    (b1, b2),
                )

    if best is None:
        raise RuntimeError(
            "Drainage centerline not found."
        )

    _, line1, line2 = best

    return (
        (
            (
                line1[0][0]
                + line2[0][0]
            ) / 2.0,

            (
                line1[0][1]
                + line2[0][1]
            ) / 2.0,
        ),

        (
            (
                line1[1][0]
                + line2[1][0]
            ) / 2.0,

            (
                line1[1][1]
                + line2[1][1]
            ) / 2.0,
        ),
    )


# ============================================================
# EXTRACT PIPE
# ============================================================

def extract_pipe(
    doc,
    handle,
    centerline,
):
    obj = get_handle_object(
        doc,
        handle,
    )

    ins = tuple(obj.InsertionPoint)

    insertion = (
        float(ins[0]),
        float(ins[1]),
    )

    rotation = float(obj.Rotation)
    sx = float(obj.XScaleFactor)
    sy = float(obj.YScaleFactor)

    lp1, lp2 = centerline

    p1 = local_to_world(
        lp1[0],
        lp1[1],
        insertion,
        rotation,
        sx,
        sy,
    )

    p2 = local_to_world(
        lp2[0],
        lp2[1],
        insertion,
        rotation,
        sx,
        sy,
    )

    center = (
        (p1[0] + p2[0]) / 2.0,
        (p1[1] + p2[1]) / 2.0,
    )

    ux, uy = normalize(
        p2[0] - p1[0],
        p2[1] - p1[1],
    )

    return {
        "object": obj,
        "handle": handle,
        "center": center,
        "p1": p1,
        "p2": p2,
        "axis_deg": orientation(p1, p2),
        "normal_a": (-uy, ux),
        "normal_b": (uy, -ux),
    }


# ============================================================
# EXTRACT SLEEPER
# ============================================================

def extract_sleeper(doc, handle):
    obj = get_handle_object(
        doc,
        handle,
    )

    ins = tuple(obj.InsertionPoint)

    insertion = (
        float(ins[0]),
        float(ins[1]),
    )

    rotation = float(obj.Rotation)

    return {
        "object": obj,
        "handle": handle,
        "block": str(obj.EffectiveName),
        "insertion": insertion,
        "rotation": rotation,
        "rotation_deg": math.degrees(rotation),
    }


# ============================================================
# CALCULATE MOVE
# ============================================================

def calculate_move(pipe, sleeper):
    px, py = pipe["center"]

    local_x, local_y = world_to_local(
        px,
        py,
        sleeper["insertion"],
        sleeper["rotation"],
    )

    # Which side of sleeper centerline contains the Pipe?
    if local_y >= 0.0:
        side_sign = 1.0
        side = "+Y"
    else:
        side_sign = -1.0
        side = "-Y"

    # Sleeper local +Y direction expressed in World coordinates.
    r = sleeper["rotation"]

    sleeper_y_world = (
        -math.sin(r),
        math.cos(r),
    )

    na = pipe["normal_a"]
    nb = pipe["normal_b"]

    dot_a = (
        na[0] * sleeper_y_world[0]
        + na[1] * sleeper_y_world[1]
    )

    dot_b = (
        nb[0] * sleeper_y_world[0]
        + nb[1] * sleeper_y_world[1]
    )

    # Select the normal pointing toward the current Pipe side.
    if side_sign > 0.0:
        move_dir = (
            na if dot_a >= dot_b else nb
        )
    else:
        move_dir = (
            na if dot_a <= dot_b else nb
        )

    current_abs_y = abs(local_y)

    move_m = max(
        0.0,
        TARGET_HALF_DISTANCE_M
        - current_abs_y,
    )

    target = (
        px + move_dir[0] * move_m,
        py + move_dir[1] * move_m,
    )

    return {
        "local_x": local_x,
        "local_y": local_y,
        "side": side,
        "move_dir": move_dir,
        "move_m": move_m,
        "target": target,
    }


# ============================================================
# MOVE OBJECT
# ============================================================

def move_object(obj, dx, dy):
    """Move BlockReference by world-coordinate vector."""

    base = point_variant(
        0.0,
        0.0,
        0.0,
    )

    destination = point_variant(
        dx,
        dy,
        0.0,
    )

    call_with_retry(
        lambda: obj.Move(
            base,
            destination,
        )
    )


# ============================================================
# DEBUG MARKER
# ============================================================

def draw_movement_marker(
    ms,
    old_center,
    new_center,
    handle,
):
    # Original position
    add_point(
        ms,
        old_center,
    )

    # New position
    add_point(
        ms,
        new_center,
    )

    # Permanent movement trace
    add_line(
        ms,
        old_center,
        new_center,
    )

    # Short label
    add_text(
        ms,
        handle,
        new_center,
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 100)
    print("MOVE TEST - 10 DRAINAGES")
    print("=" * 100)

    print(
        f"Target center distance: "
        f"{TARGET_HALF_DISTANCE_M * 1000:.1f} mm"
    )

    print(
        "175 + 31.5 + 100 = 306.5 mm"
    )

    print()
    print(
        "IMPORTANT: Save As a BACKUP of the DWG before running."
    )

    try:
        acad = win32com.client.GetActiveObject(
            "AutoCAD.Application"
        )

    except Exception as exc:
        print(
            f"ERROR connecting to AutoCAD: {exc}"
        )
        return

    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    print(
        f"Drawing: {doc.Name}"
    )

    wait_for_autocad(acad)

    # Read shared Drainage block definition once.
    print()
    print(
        "Reading Drainage Pipe E definition..."
    )

    centerline = get_drainage_centerline(
        doc
    )

    # One AutoCAD undo mark for the entire test.
    try:
        doc.StartUndoMark()
    except Exception:
        pass

    moved = 0
    failed = 0

    for index, case in enumerate(
        CASES,
        start=1,
    ):

        print()
        print("-" * 100)
        print(
            f"[{index}/10] "
            f"PIPE={case['drainage']} "
            f"SLEEPER={case['sleeper']}"
        )

        try:
            # ------------------------------------------------
            # Read objects by exact Handle.
            # ------------------------------------------------

            pipe = extract_pipe(
                doc,
                case["drainage"],
                centerline,
            )

            sleeper = extract_sleeper(
                doc,
                case["sleeper"],
            )

            # ------------------------------------------------
            # Calculate geometry BEFORE Move.
            # ------------------------------------------------

            result = calculate_move(
                pipe,
                sleeper,
            )

            old_center = pipe["center"]
            new_center = result["target"]

            dx = (
                new_center[0]
                - old_center[0]
            )

            dy = (
                new_center[1]
                - old_center[1]
            )

            move_mm = result["move_m"] * 1000.0

            print(
                f"Local Y      = "
                f"{result['local_y']:.6f} m"
            )

            print(
                f"Side         = "
                f"{result['side']}"
            )

            print(
                f"Move dir     = "
                f"({result['move_dir'][0]:.9f}, "
                f"{result['move_dir'][1]:.9f})"
            )

            print(
                f"Move         = "
                f"{move_mm:.3f} mm"
            )

            print(
                f"Old center   = "
                f"({old_center[0]:.6f}, "
                f"{old_center[1]:.6f})"
            )

            print(
                f"New target   = "
                f"({new_center[0]:.6f}, "
                f"{new_center[1]:.6f})"
            )

            # ------------------------------------------------
            # Move ONLY the drainage BlockReference.
            # ------------------------------------------------

            if move_mm > 0.0:
                move_object(
                    pipe["object"],
                    dx,
                    dy,
                )

                moved += 1

                # Leave permanent visual trace.
                draw_movement_marker(
                    ms,
                    old_center,
                    new_center,
                    case["drainage"],
                )

                print(
                    "RESULT      = MOVED"
                )

            else:
                print(
                    "RESULT      = NO MOVE NEEDED"
                )

        except Exception as exc:

            failed += 1

            print(
                f"RESULT      = FAILED"
            )

            print(
                f"ERROR       = "
                f"{type(exc).__name__}: {exc}"
            )

    try:
        doc.EndUndoMark()
    except Exception:
        pass

    try:
        doc.Regen(1)
    except Exception:
        pass

    print()
    print("=" * 100)
    print("MOVE TEST FINISHED")
    print("=" * 100)

    print(
        f"Moved       : {moved}"
    )

    print(
        f"Failed      : {failed}"
    )

    print()
    print(
        "A permanent line was created from each OLD center "
        "to its NEW center."
    )

    print(
        "The DWG was NOT saved automatically."
    )

    print(
        "Inspect the result before saving."
    )

    print("=" * 100)


if __name__ == "__main__":
    main()
