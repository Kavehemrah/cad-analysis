import csv
import math
import win32com.client


# ============================================================
# SETTINGS
# ============================================================

DRAINAGE_BLOCK = "Drainage Pipe E"

SLEEPER_LAYER = "AG_Print"
SLEEPER_BLOCKS = {"AG_t", "AG_tttt"}

PIPE_DIAMETER = 0.063       # 63 mm
PIPE_RADIUS = PIPE_DIAMETER / 2.0

SLEEPER_WIDTH = 0.350       # 350 mm
INITIAL_OFFSET = SLEEPER_WIDTH / 2.0   # 175 mm
OFFSET_STEP = 0.050         # 50 mm

MAX_OFFSET = 3.000          # safety limit: 3 meters

OUTPUT_CSV = "drainage_sleeper_move_results.csv"


# ============================================================
# BASIC GEOMETRY
# ============================================================

def local_to_world(x, y, insertion, rotation, sx=1.0, sy=1.0):

    x *= sx
    y *= sy

    c = math.cos(rotation)
    s = math.sin(rotation)

    return (
        insertion[0] + x * c - y * s,
        insertion[1] + x * s + y * c,
    )


def distance(p1, p2):

    return math.hypot(
        p2[0] - p1[0],
        p2[1] - p1[1],
    )


def orientation(a, b):

    return math.degrees(
        math.atan2(
            b[1] - a[1],
            b[0] - a[0],
        )
    )


def angle_diff(a, b):

    d = abs((a - b) % 180.0)

    return min(
        d,
        180.0 - d,
    )


def point_to_segment_distance(point, a, b):

    px, py = point
    ax, ay = a
    bx, by = b

    dx = bx - ax
    dy = by - ay

    length_sq = dx * dx + dy * dy

    if length_sq == 0:

        return math.hypot(
            px - ax,
            py - ay,
        )

    t = (
        (px - ax) * dx
        +
        (py - ay) * dy
    ) / length_sq

    t = max(
        0.0,
        min(1.0, t),
    )

    qx = ax + t * dx
    qy = ay + t * dy

    return math.hypot(
        px - qx,
        py - qy,
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
            min(p[0], r[0]) - tol
            <= q[0]
            <= max(p[0], r[0]) + tol
            and
            min(p[1], r[1]) - tol
            <= q[1]
            <= max(p[1], r[1]) + tol
        )

    c1 = cross(a, b, c)
    c2 = cross(a, b, d)
    c3 = cross(c, d, a)
    c4 = cross(c, d, b)

    if (
        (
            (c1 > tol and c2 < -tol)
            or
            (c1 < -tol and c2 > tol)
        )
        and
        (
            (c3 > tol and c4 < -tol)
            or
            (c3 < -tol and c4 > tol)
        )
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
        point_to_segment_distance(a, c, d),
        point_to_segment_distance(b, c, d),
        point_to_segment_distance(c, a, b),
        point_to_segment_distance(d, a, b),
    )


def point_in_polygon(point, polygon):

    x, y = point

    inside = False

    j = len(polygon) - 1

    for i in range(len(polygon)):

        xi, yi = polygon[i]
        xj, yj = polygon[j]

        intersects = (
            ((yi > y) != (yj > y))
            and
            (
                x
                <
                (xj - xi)
                * (y - yi)
                / (yj - yi)
                + xi
            )
        )

        if intersects:
            inside = not inside

        j = i

    return inside


# ============================================================
# POLYGON / PIPE COLLISION
# ============================================================

def pipe_to_sleeper_clearance(pipe_a, pipe_b, polygon):

    """
    Returns:
        clearance in meters.

    Positive  -> clear
    Zero      -> tangent
    Negative  -> pipe overlaps sleeper
    """

    # --------------------------------------------------------
    # If either endpoint is inside sleeper
    # --------------------------------------------------------

    if point_in_polygon(pipe_a, polygon):
        return -PIPE_RADIUS

    if point_in_polygon(pipe_b, polygon):
        return -PIPE_RADIUS

    # Midpoint check for unusual cases where
    # the whole pipe centerline segment is inside polygon.
    mid = (
        (pipe_a[0] + pipe_b[0]) / 2.0,
        (pipe_a[1] + pipe_b[1]) / 2.0,
    )

    if point_in_polygon(mid, polygon):
        return -PIPE_RADIUS

    min_distance = float("inf")

    # --------------------------------------------------------
    # Centerline against sleeper boundary
    # --------------------------------------------------------

    for i in range(len(polygon)):

        a = polygon[i]
        b = polygon[(i + 1) % len(polygon)]

        d = segment_to_segment_distance(
            pipe_a,
            pipe_b,
            a,
            b,
        )

        if d < min_distance:
            min_distance = d

    # Subtract actual pipe radius.
    return min_distance - PIPE_RADIUS


# ============================================================
# AUTO CAD BLOCK GEOMETRY
# ============================================================

def get_block_polyline(acad, block_name):

    block_def = acad.ActiveDocument.Blocks.Item(
        block_name
    )

    for ent in block_def:

        try:

            if ent.ObjectName != "AcDbPolyline":
                continue

            coords = tuple(ent.Coordinates)

            vertices = []

            for i in range(
                0,
                len(coords),
                2,
            ):

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

    block_def = acad.ActiveDocument.Blocks.Item(
        DRAINAGE_BLOCK
    )

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

            length = distance(
                p1,
                p2,
            )

            if length > 1.0:
                lines.append(
                    (p1, p2)
                )

        except Exception:

            continue

    # --------------------------------------------------------
    # Remove duplicate lines
    # --------------------------------------------------------

    unique = []

    for line in lines:

        a, b = line

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
            unique.append(line)

    if len(unique) < 2:
        return None

    # --------------------------------------------------------
    # Find two side lines
    # --------------------------------------------------------

    best = None

    for i in range(len(unique)):

        for j in range(
            i + 1,
            len(unique),
        ):

            a1, a2 = unique[i]
            b1, b2 = unique[j]

            angle1 = orientation(
                a1,
                a2,
            )

            angle2 = orientation(
                b1,
                b2,
            )

            if angle_diff(
                angle1,
                angle2,
            ) > 5.0:
                continue

            separation = point_to_segment_distance(
                b1,
                a1,
                a2,
            )

            if (
                best is None
                or separation > best[0]
            ):

                best = (
                    separation,
                    unique[i],
                    unique[j],
                )

    if best is None:
        return None

    _, line1, line2 = best

    p1 = (
        (line1[0][0] + line2[0][0]) / 2.0,
        (line1[0][1] + line2[0][1]) / 2.0,
    )

    p2 = (
        (line1[1][0] + line2[1][0]) / 2.0,
        (line1[1][1] + line2[1][1]) / 2.0,
    )

    return p1, p2


# ============================================================
# MODELSPACE OBJECTS
# ============================================================

def build_sleepers(acad, block_geometry):

    sleepers = []

    for obj in acad.ActiveDocument.ModelSpace:

        try:

            if str(obj.Layer) != SLEEPER_LAYER:
                continue

            if obj.ObjectName != "AcDbBlockReference":
                continue

            try:
                block_name = str(
                    obj.EffectiveName
                )
            except Exception:
                block_name = str(
                    obj.Name
                )

            if block_name not in SLEEPER_BLOCKS:
                continue

            insertion = tuple(
                obj.InsertionPoint
            )

            local_vertices = block_geometry[
                block_name
            ]

            world_vertices = [
                local_to_world(
                    x,
                    y,
                    insertion,
                    float(obj.Rotation),
                    float(obj.XScaleFactor),
                    float(obj.YScaleFactor),
                )
                for x, y in local_vertices
            ]

            # ------------------------------------------------
            # Bounding radius around insertion point
            # ------------------------------------------------

            max_radius = 0.0

            for p in world_vertices:

                r = distance(
                    (
                        float(insertion[0]),
                        float(insertion[1]),
                    ),
                    p,
                )

                max_radius = max(
                    max_radius,
                    r,
                )

            sleepers.append({
                "handle": str(obj.Handle),
                "block": block_name,
                "insertion": (
                    float(insertion[0]),
                    float(insertion[1]),
                ),
                "rotation_deg": math.degrees(
                    float(obj.Rotation)
                ),
                "polygon": world_vertices,
                "radius": max_radius,
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
                block_name = str(
                    obj.EffectiveName
                )
            except Exception:
                block_name = str(
                    obj.Name
                )

            if block_name != DRAINAGE_BLOCK:
                continue

            insertion = tuple(
                obj.InsertionPoint
            )

            drainages.append({
                "handle": str(obj.Handle),
                "insertion": (
                    float(insertion[0]),
                    float(insertion[1]),
                ),
                "rotation": float(
                    obj.Rotation
                ),
                "sx": float(
                    obj.XScaleFactor
                ),
                "sy": float(
                    obj.YScaleFactor
                ),
            })

        except Exception:

            continue

    return drainages


# ============================================================
# PIPE POSITION
# ============================================================

def transform_centerline(
    local_centerline,
    drainage,
):

    p1 = local_to_world(
        local_centerline[0][0],
        local_centerline[0][1],
        (
            drainage["insertion"][0],
            drainage["insertion"][1],
            0.0,
        ),
        drainage["rotation"],
        drainage["sx"],
        drainage["sy"],
    )

    p2 = local_to_world(
        local_centerline[1][0],
        local_centerline[1][1],
        (
            drainage["insertion"][0],
            drainage["insertion"][1],
            0.0,
        ),
        drainage["rotation"],
        drainage["sx"],
        drainage["sy"],
    )

    return p1, p2


def move_pipe(
    pipe_a,
    pipe_b,
    offset,
):

    dx = pipe_b[0] - pipe_a[0]
    dy = pipe_b[1] - pipe_a[1]

    length = math.hypot(
        dx,
        dy,
    )

    ux = dx / length
    uy = dy / length

    # Left-hand normal
    nx = -uy
    ny = ux

    shift_x = nx * offset
    shift_y = ny * offset

    return (
        (
            pipe_a[0] + shift_x,
            pipe_a[1] + shift_y,
        ),
        (
            pipe_b[0] + shift_x,
            pipe_b[1] + shift_y,
        ),
    )


# ============================================================
# TEST A PIPE POSITION
# ============================================================

def evaluate_pipe(
    pipe_a,
    pipe_b,
    sleepers,
):

    min_clearance = float("inf")
    blocking_sleeper = None

    # --------------------------------------------------------
    # Pipe center
    # --------------------------------------------------------

    pipe_center = (
        (pipe_a[0] + pipe_b[0]) / 2.0,
        (pipe_a[1] + pipe_b[1]) / 2.0,
    )

    pipe_half_length = (
        distance(
            pipe_a,
            pipe_b,
        ) / 2.0
    )

    # --------------------------------------------------------
    # Only inspect nearby sleepers
    # --------------------------------------------------------

    for sleeper in sleepers:

        center_distance = distance(
            pipe_center,
            sleeper["insertion"],
        )

        search_radius = (
            pipe_half_length
            +
            sleeper["radius"]
            +
            PIPE_RADIUS
            +
            0.10
        )

        if center_distance > search_radius:
            continue

        clearance = pipe_to_sleeper_clearance(
            pipe_a,
            pipe_b,
            sleeper["polygon"],
        )

        if clearance < min_clearance:

            min_clearance = clearance
            blocking_sleeper = sleeper

            if min_clearance < 0:
                # We still continue because another
                # sleeper might have a worse overlap.
                pass

    if blocking_sleeper is None:

        return {
            "clear": True,
            "min_clearance": float("inf"),
            "blocking_sleeper": None,
        }

    return {
        "clear": min_clearance >= 0.0,
        "min_clearance": min_clearance,
        "blocking_sleeper": blocking_sleeper,
    }


# ============================================================
# FIND REQUIRED MOVE
# ============================================================

def find_required_move(
    pipe_a,
    pipe_b,
    sleepers,
):

    # --------------------------------------------------------
    # First: current position
    # --------------------------------------------------------

    current = evaluate_pipe(
        pipe_a,
        pipe_b,
        sleepers,
    )

    if current["clear"]:

        return {
            "current_clear": True,
            "required_offset": 0.0,
            "direction": "NONE",
            "final_a": pipe_a,
            "final_b": pipe_b,
            "final_clearance": current["min_clearance"],
            "blocking_sleeper": (
                current["blocking_sleeper"]
            ),
        }

    # --------------------------------------------------------
    # Try 175, 225, 275...
    # in both normal directions.
    # --------------------------------------------------------

    offset = INITIAL_OFFSET

    while offset <= MAX_OFFSET:

        candidates = []

        for sign, label in (
            (+1, "+Normal"),
            (-1, "-Normal"),
        ):

            candidate_a, candidate_b = move_pipe(
                pipe_a,
                pipe_b,
                sign * offset,
            )

            evaluation = evaluate_pipe(
                candidate_a,
                candidate_b,
                sleepers,
            )

            candidates.append({
                "label": label,
                "offset": offset,
                "a": candidate_a,
                "b": candidate_b,
                "evaluation": evaluation,
            })

        clear_candidates = [
            c
            for c in candidates
            if c["evaluation"]["clear"]
        ]

        if clear_candidates:

            # If both directions work at the same
            # offset, choose the one with greater
            # final clearance.
            clear_candidates.sort(
                key=lambda c:
                c["evaluation"]["min_clearance"],
                reverse=True,
            )

            selected = clear_candidates[0]

            return {
                "current_clear": False,
                "required_offset": selected["offset"],
                "direction": selected["label"],
                "final_a": selected["a"],
                "final_b": selected["b"],
                "final_clearance": (
                    selected["evaluation"]
                    ["min_clearance"]
                ),
                "blocking_sleeper": (
                    current["blocking_sleeper"]
                ),
            }

        offset += OFFSET_STEP

    # --------------------------------------------------------
    # Could not clear inside safety limit
    # --------------------------------------------------------

    return {
        "current_clear": False,
        "required_offset": None,
        "direction": "NOT_FOUND",
        "final_a": None,
        "final_b": None,
        "final_clearance": None,
        "blocking_sleeper": (
            current["blocking_sleeper"]
        ),
    }


# ============================================================
# MAIN
# ============================================================

def main():

    acad = win32com.client.GetActiveObject(
        "AutoCAD.Application"
    )

    print("=" * 110)
    print("DRAINAGE -> SLEEPER MOVE CALCULATION")
    print("=" * 110)

    # --------------------------------------------------------
    # Sleeper geometry
    # --------------------------------------------------------

    block_geometry = {}

    for block_name in sorted(
        SLEEPER_BLOCKS
    ):

        vertices = get_block_polyline(
            acad,
            block_name,
        )

        block_geometry[
            block_name
        ] = vertices

        print(
            f"{block_name}: "
            f"{len(vertices)} vertices"
        )

    sleepers = build_sleepers(
        acad,
        block_geometry,
    )

    print(
        f"Sleepers found: "
        f"{len(sleepers)}"
    )

    # --------------------------------------------------------
    # Drainage
    # --------------------------------------------------------

    drainages = build_drainages(
        acad
    )

    print(
        f"Drainage Pipe E found: "
        f"{len(drainages)}"
    )

    if not drainages:
        print("ERROR: No drainages found.")
        return

    local_centerline = get_drainage_centerline(
        acad
    )

    if local_centerline is None:

        print(
            "ERROR: Drainage centerline not found."
        )

        return

    print(
        "Drainage centerline: OK"
    )

    # --------------------------------------------------------
    # Results
    # --------------------------------------------------------

    results = []

    current_clear_count = 0
    moved_count = 0
    not_found_count = 0

    for index, drainage in enumerate(
        drainages,
        start=1,
    ):

        pipe_a, pipe_b = transform_centerline(
            local_centerline,
            drainage,
        )

        pipe_axis = orientation(
            pipe_a,
            pipe_b,
        )

        result = find_required_move(
            pipe_a,
            pipe_b,
            sleepers,
        )

        blocking = result[
            "blocking_sleeper"
        ]

        if result["current_clear"]:

            status = "CLEAR"
            current_clear_count += 1

        elif result["required_offset"] is None:

            status = "NOT_FOUND"
            not_found_count += 1

        else:

            status = "MOVE"
            moved_count += 1

        sleeper_axis = None
        angle_difference = None
        blocking_handle = None

        if blocking is not None:

            sleeper_axis = (
                blocking["rotation_deg"]
            )

            angle_difference = angle_diff(
                pipe_axis,
                sleeper_axis,
            )

            blocking_handle = (
                blocking["handle"]
            )

        required_mm = None

        if result["required_offset"] is not None:

            required_mm = (
                result["required_offset"]
                * 1000.0
            )

        current_clearance_mm = None

        # Re-evaluate current position for
        # exact clearance value.
        current_eval = evaluate_pipe(
            pipe_a,
            pipe_b,
            sleepers,
        )

        if math.isfinite(
            current_eval["min_clearance"]
        ):

            current_clearance_mm = (
                current_eval["min_clearance"]
                * 1000.0
            )

        final_clearance_mm = None

        if result["final_clearance"] is not None:

            if math.isfinite(
                result["final_clearance"]
            ):

                final_clearance_mm = (
                    result["final_clearance"]
                    * 1000.0
                )

        row = {
            "Drainage": drainage["handle"],
            "Status": status,
            "BlockingSleeper": blocking_handle,
            "PipeAxis_deg": round(
                pipe_axis,
                3,
            ),
            "SleeperAxis_deg": (
                round(
                    sleeper_axis,
                    3,
                )
                if sleeper_axis is not None
                else None
            ),
            "AngleDiff_deg": (
                round(
                    angle_difference,
                    3,
                )
                if angle_difference is not None
                else None
            ),
            "CurrentClearance_mm": (
                round(
                    current_clearance_mm,
                    1,
                )
                if current_clearance_mm is not None
                else None
            ),
            "MoveDirection": (
                result["direction"]
            ),
            "RequiredOffset_mm": (
                round(
                    required_mm,
                    1,
                )
                if required_mm is not None
                else None
            ),
            "FinalClearance_mm": (
                round(
                    final_clearance_mm,
                    1,
                )
                if final_clearance_mm is not None
                else None
            ),
        }

        results.append(row)

        print(
            f"{drainage['handle']:>6} | "
            f"{status:<9} | "
            f"Sleeper={str(blocking_handle):<6} | "
            f"Current={str(current_clearance_mm):>8} mm | "
            f"Move={str(required_mm):>6} mm | "
            f"{result['direction']:<9}"
        )

    # --------------------------------------------------------
    # CSV
    # --------------------------------------------------------

    fields = [
        "Drainage",
        "Status",
        "BlockingSleeper",
        "PipeAxis_deg",
        "SleeperAxis_deg",
        "AngleDiff_deg",
        "CurrentClearance_mm",
        "MoveDirection",
        "RequiredOffset_mm",
        "FinalClearance_mm",
    ]

    with open(
        OUTPUT_CSV,
        "w",
        newline="",
        encoding="utf-8-sig",
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=fields,
        )

        writer.writeheader()

        writer.writerows(results)

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    print("\n" + "=" * 110)
    print("SUMMARY")
    print("=" * 110)

    print(
        f"Drainages analyzed : "
        f"{len(results)}"
    )

    print(
        f"Already clear      : "
        f"{current_clear_count}"
    )

    print(
        f"Need movement      : "
        f"{moved_count}"
    )

    print(
        f"Not found <= 3 m   : "
        f"{not_found_count}"
    )

    print(
        f"Initial offset     : "
        f"{INITIAL_OFFSET * 1000:.0f} mm"
    )

    print(
        f"Offset step        : "
        f"{OFFSET_STEP * 1000:.0f} mm"
    )

    print(
        f"Pipe diameter      : "
        f"{PIPE_DIAMETER * 1000:.0f} mm"
    )

    print(
        f"CSV saved          : "
        f"{OUTPUT_CSV}"
    )


if __name__ == "__main__":
    main()
