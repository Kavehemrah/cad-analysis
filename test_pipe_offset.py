import math

import pythoncom
import win32com.client
from win32com.client import VARIANT


# ============================================================
# TEST CASE
# ============================================================
PIPE_X = 563273.517648
PIPE_Y = 3594736.029122
PIPE_Z = 0.0

P1_X = 563274.664455
P1_Y = 3594736.625886
P2_X = 563272.381242
P2_Y = 3594735.437862

OVERLAP_MM = 50.0
CLEARANCE_MM = 100.0
OFFSET_MM = OVERLAP_MM + CLEARANCE_MM  # 150 mm
OFFSET_M = OFFSET_MM / 1000.0

SLEEPER_LAYER = "AG_Print"
SLEEPER_BLOCKS = {"AG_t", "AG_tttt"}
SEARCH_RADIUS_M = 5.0


def point_variant(x: float, y: float, z: float = 0.0):
    return VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        (float(x), float(y), float(z)),
    )


def distance2(ax, ay, bx, by):
    dx = bx - ax
    dy = by - ay
    return dx * dx + dy * dy


def local_to_world(x, y, insertion, rotation, sx=1.0, sy=1.0):
    x *= sx
    y *= sy
    c = math.cos(rotation)
    s = math.sin(rotation)
    return (
        insertion[0] + x * c - y * s,
        insertion[1] + x * s + y * c,
    )


def polygon_centroid(points):
    if not points:
        raise RuntimeError("Sleeper polygon has no vertices.")

    area2 = 0.0
    cx = 0.0
    cy = 0.0

    for i, (x1, y1) in enumerate(points):
        x2, y2 = points[(i + 1) % len(points)]
        cross = x1 * y2 - x2 * y1
        area2 += cross
        cx += (x1 + x2) * cross
        cy += (y1 + y2) * cross

    if abs(area2) < 1e-12:
        return (
            sum(x for x, _ in points) / len(points),
            sum(y for _, y in points) / len(points),
        )

    return cx / (3.0 * area2), cy / (3.0 * area2)


def get_sleeper_geometry(doc, block_name):
    block_def = doc.Blocks.Item(block_name)

    for ent in block_def:
        try:
            if ent.ObjectName != "AcDbPolyline":
                continue

            coords = tuple(ent.Coordinates)
            vertices = [
                (float(coords[i]), float(coords[i + 1]))
                for i in range(0, len(coords), 2)
            ]

            if len(vertices) >= 3:
                return vertices
        except Exception:
            continue

    raise RuntimeError(f"No usable polyline found in sleeper block: {block_name}")


def find_nearest_sleeper(doc):
    """Find nearest sleeper using insertion points only."""
    best = None
    limit2 = SEARCH_RADIUS_M * SEARCH_RADIUS_M

    for obj in doc.ModelSpace:
        try:
            if obj.ObjectName != "AcDbBlockReference":
                continue
            if str(obj.Layer) != SLEEPER_LAYER:
                continue

            try:
                block_name = str(obj.EffectiveName)
            except Exception:
                block_name = str(obj.Name)

            if block_name not in SLEEPER_BLOCKS:
                continue

            ins = tuple(obj.InsertionPoint)
            ix = float(ins[0])
            iy = float(ins[1])
            d2 = distance2(PIPE_X, PIPE_Y, ix, iy)

            if d2 > limit2:
                continue

            if best is None or d2 < best["d2"]:
                best = {
                    "obj": obj,
                    "block": block_name,
                    "insertion": (ix, iy),
                    "rotation": float(obj.Rotation),
                    "sx": float(obj.XScaleFactor),
                    "sy": float(obj.YScaleFactor),
                    "d2": d2,
                }
        except Exception:
            continue

    return best


# ============================================================
# MAIN
# ============================================================
acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument
ms = doc.ModelSpace

# Pipe axis.
dx = P2_X - P1_X
dy = P2_Y - P1_Y
pipe_length = math.hypot(dx, dy)
if pipe_length < 1e-12:
    raise RuntimeError("Pipe P1/P2 are identical.")

ux = dx / pipe_length
uy = dy / pipe_length

# Two possible perpendicular directions to the pipe.
n1 = (-uy, ux)
n2 = (uy, -ux)

# Fast search: find the nearest sleeper from insertion points.
sleeper = find_nearest_sleeper(doc)
if sleeper is None:
    raise RuntimeError(
        f"No {SLEEPER_LAYER} sleeper block found within {SEARCH_RADIUS_M:.1f} m."
    )

# Geometry is transformed only for that one sleeper.
local_vertices = get_sleeper_geometry(doc, sleeper["block"])
world_vertices = [
    local_to_world(
        x,
        y,
        sleeper["insertion"],
        sleeper["rotation"],
        sleeper["sx"],
        sleeper["sy"],
    )
    for x, y in local_vertices
]

sleeper_cx, sleeper_cy = polygon_centroid(world_vertices)

# Pick the normal that points away from the sleeper center.
# This avoids the previous bad assumption that Y- is always "down".
side_x = PIPE_X - sleeper_cx
side_y = PIPE_Y - sleeper_cy

score1 = side_x * n1[0] + side_y * n1[1]
score2 = side_x * n2[0] + side_y * n2[1]

away_x, away_y = n1 if score1 >= score2 else n2

new_x = PIPE_X + away_x * OFFSET_M
new_y = PIPE_Y + away_y * OFFSET_M

# Draw minimal debug geometry.
ms.AddPoint(point_variant(PIPE_X, PIPE_Y, PIPE_Z))
ms.AddPoint(point_variant(sleeper_cx, sleeper_cy, PIPE_Z))
ms.AddPoint(point_variant(new_x, new_y, PIPE_Z))
ms.AddLine(
    point_variant(PIPE_X, PIPE_Y, PIPE_Z),
    point_variant(new_x, new_y, PIPE_Z),
)

text_height = 50.0
ms.AddText("ORIGINAL PIPE", point_variant(PIPE_X, PIPE_Y, PIPE_Z), text_height)
ms.AddText("SLEEPER CENTER", point_variant(sleeper_cx, sleeper_cy, PIPE_Z), text_height)
ms.AddText("OFFSET 150mm AWAY", point_variant(new_x, new_y, PIPE_Z), text_height)

doc.Regen(1)

print("=" * 72)
print("PIPE OFFSET DIRECTION TEST - FAST")
print("=" * 72)
print(f"Original        X={PIPE_X:.6f}, Y={PIPE_Y:.6f}")
print(f"Pipe axis       ({ux:.9f}, {uy:.9f})")
print(f"Nearest sleeper {sleeper['block']} / handle={sleeper['obj'].Handle}")
print(f"Sleeper insert  X={sleeper['insertion'][0]:.6f}, Y={sleeper['insertion'][1]:.6f}")
print(f"Sleeper center  X={sleeper_cx:.6f}, Y={sleeper_cy:.6f}")
print(f"Insert distance = {math.sqrt(sleeper['d2']):.6f} m")
print(f"Overlap         = {OVERLAP_MM:.1f} mm")
print(f"Clearance       = {CLEARANCE_MM:.1f} mm")
print(f"Total offset    = {OFFSET_MM:.1f} mm")
print("-" * 72)
print(f"Chosen AWAY normal = ({away_x:.9f}, {away_y:.9f})")
print(f"Offset vector      = ({away_x * OFFSET_M:.6f}, {away_y * OFFSET_M:.6f})")
print(f"New position       X={new_x:.6f}, Y={new_y:.6f}")
print("-" * 72)
print("Direction is selected from the nearest sleeper geometry.")
print("Drawing was NOT saved automatically.")
print("=" * 72)
