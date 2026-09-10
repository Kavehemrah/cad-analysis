import math

import pythoncom
import win32com.client
from win32com.client import VARIANT


# ============================================================
# REAL TEST CASE
# ============================================================
PIPE_X = 563273.517648
PIPE_Y = 3594736.029122
PIPE_Z = 0.0

P1_X = 563274.664455
P1_Y = 3594736.625886
P2_X = 563272.381242
P2_Y = 3594735.437862

DRAINAGE_BLOCK = "Drainage Pipe E"
SLEEPER_LAYER = "AG_Print"
SLEEPER_BLOCKS = {"AG_t", "AG_tttt"}

OVERLAP_MM = 50.0
CLEARANCE_MM = 100.0
OFFSET_MM = OVERLAP_MM + CLEARANCE_MM
OFFSET_M = OFFSET_MM / 1000.0


def point_variant(x, y, z=0.0):
    return VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        (float(x), float(y), float(z)),
    )


def local_to_world(x, y, insertion, rotation, sx=1.0, sy=1.0):
    x *= sx
    y *= sy
    c = math.cos(rotation)
    s = math.sin(rotation)
    return (
        insertion[0] + x * c - y * s,
        insertion[1] + x * s + y * c,
    )


def point_to_segment_distance(px, py, ax, ay, bx, by):
    dx = bx - ax
    dy = by - ay
    length_sq = dx * dx + dy * dy
    if length_sq == 0.0:
        return math.hypot(px - ax, py - ay)

    t = ((px - ax) * dx + (py - ay) * dy) / length_sq
    t = max(0.0, min(1.0, t))
    qx = ax + t * dx
    qy = ay + t * dy
    return math.hypot(px - qx, py - qy)


def point_in_polygon(point, polygon):
    x, y = point
    inside = False
    j = len(polygon) - 1

    for i in range(len(polygon)):
        xi, yi = polygon[i]
        xj, yj = polygon[j]
        if ((yi > y) != (yj > y)):
            x_cross = (xj - xi) * (y - yi) / (yj - yi) + xi
            if x < x_cross:
                inside = not inside
        j = i

    return inside


def polygon_centroid(polygon):
    if not polygon:
        raise RuntimeError("Sleeper polygon is empty.")

    # Average is sufficient for choosing which side of the sleeper
    # contains the pipe; this is a direction test, not a survey result.
    x = sum(p[0] for p in polygon) / len(polygon)
    y = sum(p[1] for p in polygon) / len(polygon)
    return x, y


def get_block_vertices(doc, block_name):
    block_def = doc.Blocks.Item(block_name)

    for ent in block_def:
        try:
            if ent.ObjectName != "AcDbPolyline":
                continue

            coords = tuple(ent.Coordinates)
            vertices = []
            for i in range(0, len(coords), 2):
                vertices.append((float(coords[i]), float(coords[i + 1])))

            if len(vertices) >= 3:
                return vertices
        except Exception:
            continue

    return []


def read_sleepers(doc):
    geometry = {}
    for block_name in SLEEPER_BLOCKS:
        geometry[block_name] = get_block_vertices(doc, block_name)

    sleepers = []
    for obj in doc.ModelSpace:
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
            local_vertices = geometry[block_name]
            if len(local_vertices) < 3:
                continue

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

            sleepers.append({
                "handle": str(obj.Handle),
                "block": block_name,
                "insertion": (float(insertion[0]), float(insertion[1])),
                "rotation": float(obj.Rotation),
                "polygon": world_vertices,
            })
        except Exception:
            continue

    return sleepers


def distance_to_polygon(point, polygon):
    if point_in_polygon(point, polygon):
        return 0.0

    best = float("inf")
    for i in range(len(polygon)):
        a = polygon[i]
        b = polygon[(i + 1) % len(polygon)]
        d = point_to_segment_distance(
            point[0], point[1],
            a[0], a[1],
            b[0], b[1],
        )
        best = min(best, d)
    return best


acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument
ms = doc.ModelSpace

sleepers = read_sleepers(doc)
if not sleepers:
    raise RuntimeError(
        f"No sleeper blocks found on layer '{SLEEPER_LAYER}'."
    )

# Select the physically nearest sleeper to the pipe insertion.
pipe_point = (PIPE_X, PIPE_Y)
nearest = min(
    sleepers,
    key=lambda s: distance_to_polygon(pipe_point, s["polygon"]),
)

sleeper_polygon = nearest["polygon"]
sleeper_center = polygon_centroid(sleeper_polygon)

# Pipe axis unit vector from the actual P1/P2 data.
axis_dx = P2_X - P1_X
axis_dy = P2_Y - P1_Y
axis_length = math.hypot(axis_dx, axis_dy)
if axis_length < 1e-12:
    raise RuntimeError("P1 and P2 define a zero-length axis.")

ux = axis_dx / axis_length
uy = axis_dy / axis_length

# Two possible perpendicular directions to the pipe.
n1x, n1y = -uy, ux
n2x, n2y = uy, -ux

# Critical step:
# choose the normal that points from the sleeper interior/center
# toward the current pipe position. Therefore the offset moves OUT
# of the sleeper, rather than assuming Y+ or Y-.
to_pipe_x = PIPE_X - sleeper_center[0]
to_pipe_y = PIPE_Y - sleeper_center[1]

if to_pipe_x * n1x + to_pipe_y * n1y >= to_pipe_x * n2x + to_pipe_y * n2y:
    away_x, away_y = n1x, n1y
    direction_name = "NORMAL 1 (-uy, +ux)"
else:
    away_x, away_y = n2x, n2y
    direction_name = "NORMAL 2 (+uy, -ux)"

# Candidate coordinates.
new_x = PIPE_X + away_x * OFFSET_M
new_y = PIPE_Y + away_y * OFFSET_M

# Also create the opposite candidate so both directions are visible.
opposite_x = PIPE_X - away_x * OFFSET_M
opposite_y = PIPE_Y - away_y * OFFSET_M

# Draw test markers.
ms.AddPoint(point_variant(PIPE_X, PIPE_Y, PIPE_Z))
ms.AddPoint(point_variant(new_x, new_y, PIPE_Z))
ms.AddPoint(point_variant(opposite_x, opposite_y, PIPE_Z))

ms.AddLine(
    point_variant(PIPE_X, PIPE_Y, PIPE_Z),
    point_variant(new_x, new_y, PIPE_Z),
)
ms.AddLine(
    point_variant(PIPE_X, PIPE_Y, PIPE_Z),
    point_variant(opposite_x, opposite_y, PIPE_Z),
)

# Connect pipe position to sleeper center for visual verification.
ms.AddLine(
    point_variant(sleeper_center[0], sleeper_center[1], PIPE_Z),
    point_variant(PIPE_X, PIPE_Y, PIPE_Z),
)

text_height = 50.0
ms.AddText(
    "ORIGINAL PIPE",
    point_variant(PIPE_X, PIPE_Y, PIPE_Z),
    text_height,
)
ms.AddText(
    "AWAY FROM SLEEPER +150mm",
    point_variant(new_x, new_y, PIPE_Z),
    text_height,
)
ms.AddText(
    "TOWARD SLEEPER -150mm",
    point_variant(opposite_x, opposite_y, PIPE_Z),
    text_height,
)
ms.AddText(
    "NEAREST SLEEPER CENTER",
    point_variant(sleeper_center[0], sleeper_center[1], PIPE_Z),
    text_height,
)

doc.Regen(1)

print("=" * 76)
print("PIPE OFFSET DIRECTION TEST - SLEEPER AWARE")
print("=" * 76)
print(f"Original      : X={PIPE_X:.6f}, Y={PIPE_Y:.6f}")
print(f"Nearest sleeper: {nearest['handle']} ({nearest['block']})")
print(f"Sleeper center: X={sleeper_center[0]:.6f}, Y={sleeper_center[1]:.6f}")
print(f"Sleeper distance from pipe: {distance_to_polygon(pipe_point, sleeper_polygon):.6f} m")
print(f"Pipe axis     : ux={ux:.9f}, uy={uy:.9f}")
print(f"Chosen normal : {direction_name}")
print(f"Away vector   : dx={away_x:.9f}, dy={away_y:.9f}")
print(f"Offset        : {OFFSET_MM:.1f} mm")
print("-" * 76)
print(f"AWAY position : X={new_x:.6f}, Y={new_y:.6f}")
print(f"TOWARD pos    : X={opposite_x:.6f}, Y={opposite_y:.6f}")
print("-" * 76)
print("The +150 mm candidate is chosen from sleeper geometry, not a fixed Y direction.")
print("Drawing was NOT saved automatically.")
print("=" * 76)
