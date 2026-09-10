import math

import pythoncom
import win32com.client
from win32com.client import VARIANT


PIPE_X = 563273.517648
PIPE_Y = 3594736.029122
PIPE_Z = 0.0

P1_X = 563274.664455
P1_Y = 3594736.625886
P2_X = 563272.381242
P2_Y = 3594735.437862

# Test assumptions discussed for this case.
OVERLAP_MM = 50.0
CLEARANCE_MM = 100.0
OFFSET_MM = OVERLAP_MM + CLEARANCE_MM  # 150 mm
OFFSET_M = OFFSET_MM / 1000.0


def point_variant(x: float, y: float, z: float = 0.0):
    return VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        (float(x), float(y), float(z)),
    )


acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument
ms = doc.ModelSpace

# Centerline direction.
dx = P2_X - P1_X
dy = P2_Y - P1_Y
length = math.hypot(dx, dy)
if length == 0:
    raise RuntimeError("Centerline P1/P2 are identical.")

tx = dx / length
ty = dy / length

# Find the closest point on the centerline to the pipe insertion.
rx = PIPE_X - P1_X
ry = PIPE_Y - P1_Y
projection = rx * tx + ry * ty
foot_x = P1_X + projection * tx
foot_y = P1_Y + projection * ty

# The vector from the centerline toward the pipe is the desired
# 'away from sleeper/centerline' direction for this test.
side_x = PIPE_X - foot_x
side_y = PIPE_Y - foot_y
side_length = math.hypot(side_x, side_y)
if side_length < 1e-9:
    raise RuntimeError("Pipe is exactly on the centerline; side is undefined.")

away_x = side_x / side_length
away_y = side_y / side_length

offset_x = away_x * OFFSET_M
offset_y = away_y * OFFSET_M
new_x = PIPE_X + offset_x
new_y = PIPE_Y + offset_y

# Draw original, projection foot, and proposed offset point.
ms.AddPoint(point_variant(PIPE_X, PIPE_Y, PIPE_Z))
ms.AddPoint(point_variant(foot_x, foot_y, PIPE_Z))
ms.AddPoint(point_variant(new_x, new_y, PIPE_Z))

# Draw the direction from centerline to pipe and the proposed offset.
ms.AddLine(
    point_variant(foot_x, foot_y, PIPE_Z),
    point_variant(PIPE_X, PIPE_Y, PIPE_Z),
)
ms.AddLine(
    point_variant(PIPE_X, PIPE_Y, PIPE_Z),
    point_variant(new_x, new_y, PIPE_Z),
)

text_height = 50.0
ms.AddText(
    "ORIGINAL PIPE",
    point_variant(PIPE_X, PIPE_Y, PIPE_Z),
    text_height,
)
ms.AddText(
    "CENTERLINE FOOT",
    point_variant(foot_x, foot_y, PIPE_Z),
    text_height,
)
ms.AddText(
    "OFFSET 150mm AWAY",
    point_variant(new_x, new_y, PIPE_Z),
    text_height,
)

doc.Regen(1)

print("=" * 72)
print("PIPE OFFSET DIRECTION TEST")
print("=" * 72)
print(f"Original      X={PIPE_X:.6f}, Y={PIPE_Y:.6f}")
print(f"Centerline P1 X={P1_X:.6f}, Y={P1_Y:.6f}")
print(f"Centerline P2 X={P2_X:.6f}, Y={P2_Y:.6f}")
print(f"Overlap       = {OVERLAP_MM:.1f} mm")
print(f"Clearance     = {CLEARANCE_MM:.1f} mm")
print(f"Total offset  = {OFFSET_MM:.1f} mm")
print("-" * 72)
print(f"Centerline foot X={foot_x:.6f}, Y={foot_y:.6f}")
print(f"Away direction dx={away_x:.9f}, dy={away_y:.9f}")
print(f"Offset vector   dx={offset_x:.6f}, dy={offset_y:.6f}")
print(f"New position    X={new_x:.6f}, Y={new_y:.6f}")
print("-" * 72)
print("Direction is derived from the pipe side of the centerline.")
print("Drawing was NOT saved automatically.")
print("=" * 72)
