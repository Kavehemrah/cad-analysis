import math
import time
from dataclasses import dataclass, asdict
from typing import Optional

import pythoncom
import win32com.client
from win32com.client import VARIANT


DRAINAGE_BLOCK = "Drainage Pipe E"
SLEEPER_LAYER = "AG_Print"
SLEEPER_BLOCKS = {"AG_t", "AG_tttt"}

EXPECTED_PIPE_DIAMETER_M = 0.063
DEFAULT_CLEARANCE_M = 0.100
MIN_PIPE_LINE_LENGTH_M = 1.0
PAIR_ANGLE_TOLERANCE_DEG = 2.0
SLEEPER_ANGLE_TOLERANCE_DEG = 8.0
MATCH_MARGIN_M = 0.080
GRID_SIZE_M = 5.0
COM_RETRIES = 12
COM_DELAY = 0.20
MARKER_LAYER = "CAD_ANALYSIS_MOVE"


def distance(a, b):
    return math.hypot(b[0] - a[0], b[1] - a[1])


def angle_diff(a, b):
    d = abs((a - b) % 180.0)
    return min(d, 180.0 - d)


def orientation(a, b):
    return math.degrees(math.atan2(b[1] - a[1], b[0] - a[0]))


def normalize(x, y):
    length = math.hypot(x, y)
    if length < 1e-12:
        raise ValueError("Zero-length vector")
    return x / length, y / length


def local_to_world(x, y, insertion, rotation, sx=1.0, sy=1.0):
    x *= sx
    y *= sy
    c = math.cos(rotation)
    s = math.sin(rotation)
    return (insertion[0] + x * c - y * s,
            insertion[1] + x * s + y * c)


def world_to_local(x, y, insertion, rotation):
    dx = x - insertion[0]
    dy = y - insertion[1]
    c = math.cos(rotation)
    s = math.sin(rotation)
    return (dx * c + dy * s, -dx * s + dy * c)


def is_busy(exc):
    text = str(exc)
    return ("Call was rejected by callee" in text or
            "RPC_E_CALL_REJECTED" in text or
            "-2147418111" in text)


def retry(call):
    last = None
    for _ in range(COM_RETRIES):
        try:
            return call()
        except Exception as exc:
            last = exc
            if not is_busy(exc):
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(COM_DELAY)
    raise last


@dataclass
class Pipe:
    handle: str
    obj: object
    insertion: tuple
    center: tuple
    p1: tuple
    p2: tuple
    axis_deg: float
    radius_m: float
    normal_a: tuple
    normal_b: tuple


@dataclass
class Sleeper:
    handle: str
    obj: object
    block: str
    insertion: tuple
    rotation: float
    width_m: float
    length_m: float
    center_local: tuple
    center_world: tuple


@dataclass
class Result:
    drainage: str
    sleeper: str = ""
    status: str = ""
    confidence: str = ""
    pipe_axis_deg: float = 0.0
    sleeper_axis_deg: float = 0.0
    angle_diff_deg: float = 0.0
    pipe_x: float = 0.0
    pipe_y: float = 0.0
    sleeper_x: float = 0.0
    sleeper_y: float = 0.0
    local_x: float = 0.0
    local_y: float = 0.0
    pipe_diameter_mm: float = 0.0
    sleeper_width_mm: float = 0.0
    current_clearance_mm: float = 0.0
    required_move_mm: float = 0.0
    final_clearance_mm: float = 0.0
    move_dir_x: float = 0.0
    move_dir_y: float = 0.0
    old_x: float = 0.0
    old_y: float = 0.0
    new_x: float = 0.0
    new_y: float = 0.0
    note: str = ""


class Detector:
    def __init__(self, doc, clearance_m=DEFAULT_CLEARANCE_M):
        self.doc = doc
        self.clearance_m = clearance_m
        self.centerline_local, diameter_local = self._discover_pipe_profile()
        self.pipe_radius_local = diameter_local / 2.0
        self.sleeper_profiles = {
            name: self._discover_sleeper_profile(name)
            for name in SLEEPER_BLOCKS
        }

    def _discover_pipe_profile(self):
        block = self.doc.Blocks.Item(DRAINAGE_BLOCK)
        lines = []
        for ent in block:
            try:
                if ent.ObjectName != "AcDbLine" or str(ent.Layer) != "E Pipe":
                    continue
                a = tuple(ent.StartPoint)
                b = tuple(ent.EndPoint)
                p1 = (float(a[0]), float(a[1]))
                p2 = (float(b[0]), float(b[1]))
                length = distance(p1, p2)
                if length >= MIN_PIPE_LINE_LENGTH_M:
                    lines.append((p1, p2, length, orientation(p1, p2)))
            except Exception:
                continue
        if len(lines) < 2:
            raise RuntimeError("Could not find two E Pipe side lines")
        best = None
        for i in range(len(lines)):
            for j in range(i + 1, len(lines)):
                a1, a2, la, aa = lines[i]
                b1, b2, lb, ab = lines[j]
                if angle_diff(aa, ab) > PAIR_ANGLE_TOLERANCE_DEG:
                    continue
                min_len = min(la, lb)
                ratio = min_len / max(la, lb)
                if ratio < 0.90:
                    continue
                sx = a2[0] - a1[0]
                sy = a2[1] - a1[1]
                denom = math.hypot(sx, sy)
                if denom < 1e-12:
                    continue
                dx = b1[0] - a1[0]
                dy = b1[1] - a1[1]
                sep = abs(dx * sy - dy * sx) / denom
                if not (0.005 <= sep <= 0.20):
                    continue
                # Prefer long, similar lines and a diameter close to the known project size.
                score = min_len + 2.0 * ratio - 20.0 * abs(sep - EXPECTED_PIPE_DIAMETER_M)
                if best is None or score > best[0]:
                    best = (score, (a1, a2), (b1, b2), sep)
        if best is None:
            raise RuntimeError("Could not identify pipe side pair")
        _, line1, line2, sep = best
        p1 = ((line1[0][0] + line2[0][0]) / 2.0,
              (line1[0][1] + line2[0][1]) / 2.0)
        p2 = ((line1[1][0] + line2[1][0]) / 2.0,
              (line1[1][1] + line2[1][1]) / 2.0)
        return (p1, p2), sep

    def _entity_bbox(self, ent):
        try:
            lo = tuple(ent.GeometricExtents.MinPoint)
            hi = tuple(ent.GeometricExtents.MaxPoint)
            return float(lo[0]), float(lo[1]), float(hi[0]), float(hi[1])
        except Exception:
            return None

    def _discover_sleeper_profile(self, block_name):
        block = self.doc.Blocks.Item(block_name)
        boxes = []
        for ent in block:
            try:
                name = str(ent.ObjectName)
                if name not in {"AcDbLine", "AcDbPolyline", "AcDb2dPolyline", "AcDbCircle", "AcDbArc"}:
                    continue
                box = self._entity_bbox(ent)
                if box:
                    boxes.append(box)
            except Exception:
                continue
        if not boxes:
            raise RuntimeError(f"No usable geometry in sleeper block {block_name}")
        min_x = min(b[0] for b in boxes)
        min_y = min(b[1] for b in boxes)
        max_x = max(b[2] for b in boxes)
        max_y = max(b[3] for b in boxes)
        width = max_y - min_y
        length = max_x - min_x
        return {
            "min_x": min_x, "min_y": min_y,
            "max_x": max_x, "max_y": max_y,
            "width_m": min(width, length),
            "length_m": max(width, length),
            "center_local": ((min_x + max_x) / 2.0, (min_y + max_y) / 2.0),
        }

    def collect_instances(self):
        drainages = []
        sleepers = []
        for obj in retry(lambda: self.doc.ModelSpace):
            try:
                if obj.ObjectName != "AcDbBlockReference":
                    continue
                name = str(obj.EffectiveName)
                ins = tuple(obj.InsertionPoint)
                insertion = (float(ins[0]), float(ins[1]))
                rot = float(obj.Rotation)
                sx = float(obj.XScaleFactor)
                sy = float(obj.YScaleFactor)
                handle = str(obj.Handle)
                if name == DRAINAGE_BLOCK:
                    drainages.append({"handle": handle, "obj": obj, "insertion": insertion,
                                      "rotation": rot, "sx": sx, "sy": sy})
                elif str(obj.Layer) == SLEEPER_LAYER and name in SLEEPER_BLOCKS:
                    profile = self.sleeper_profiles[name]
                    center_local = profile["center_local"]
                    center_world = local_to_world(center_local[0], center_local[1], insertion, rot, sx, sy)
                    sleepers.append(Sleeper(handle, obj, name, insertion, rot,
                                            profile["width_m"] * max(abs(sx), abs(sy)),
                                            profile["length_m"] * max(abs(sx), abs(sy)),
                                            center_local, center_world))
            except Exception:
                continue
        return drainages, sleepers

    @staticmethod
    def _grid_key(p, size):
        return (math.floor(p[0] / size), math.floor(p[1] / size))

    def _build_grid(self, sleepers):
        grid = {}
        for s in sleepers:
            key = self._grid_key(s.center_world, GRID_SIZE_M)
            grid.setdefault(key, []).append(s)
        return grid

    def extract_pipe(self, raw, scale_radius=True):
        obj = raw["obj"]
        ins = raw["insertion"]
        rot = raw["rotation"]
        sx = float(raw["sx"])
        sy = float(raw["sy"])
        lp1, lp2 = self.centerline_local
        p1 = local_to_world(lp1[0], lp1[1], ins, rot, sx, sy)
        p2 = local_to_world(lp2[0], lp2[1], ins, rot, sx, sy)
        center = ((p1[0] + p2[0]) / 2.0, (p1[1] + p2[1]) / 2.0)
        ux, uy = normalize(p2[0] - p1[0], p2[1] - p1[1])
        scale = (abs(sx) + abs(sy)) / 2.0 if scale_radius else 1.0
        return Pipe(raw["handle"], obj, ins, center, p1, p2,
                    orientation(p1, p2), self.pipe_radius_local * scale,
                    (-uy, ux), (uy, -ux))

    def match_sleeper(self, pipe, grid, sleepers):
        key = self._grid_key(pipe.center, GRID_SIZE_M)
        candidates = []
        for ix in range(key[0] - 1, key[0] + 2):
            for iy in range(key[1] - 1, key[1] + 2):
                for s in grid.get((ix, iy), []):
                    lx, ly = world_to_local(pipe.center[0], pipe.center[1], s.center_world, s.rotation)
                    # Convert relative to sleeper centroid, not insertion point.
                    ad = angle_diff(pipe.axis_deg, math.degrees(s.rotation))
                    half_w = s.width_m / 2.0
                    half_l = s.length_m / 2.0
                    if ad > SLEEPER_ANGLE_TOLERANCE_DEG:
                        continue
                    if abs(lx) > half_l + MATCH_MARGIN_M:
                        continue
                    if abs(ly) > half_w + pipe.radius_m + MATCH_MARGIN_M:
                        continue
                    score = abs(ly) + 0.25 * abs(lx) + 0.02 * ad
                    candidates.append((score, s, lx, ly, ad))
        candidates.sort(key=lambda x: x[0])
        if not candidates:
            return None, "", []
        best = candidates[0]
        if len(candidates) > 1 and candidates[1][0] - best[0] < 0.020:
            return None, "AMBIGUOUS", candidates[:3]
        confidence = "HIGH" if best[4] <= 2.0 else "MEDIUM"
        return best, confidence, candidates[:3]

    def analyze_all(self):
        raw_drainages, sleepers = self.collect_instances()
        grid = self._build_grid(sleepers)
        results = []
        for raw in raw_drainages:
            try:
                pipe = self.extract_pipe(raw)
                match, confidence, alternatives = self.match_sleeper(pipe, grid, sleepers)
                r = Result(drainage=pipe.handle,
                           pipe_axis_deg=pipe.axis_deg,
                           pipe_x=pipe.center[0], pipe_y=pipe.center[1],
                           pipe_diameter_mm=pipe.radius_m * 2000.0)
                if match is None:
                    r.status = "UNMATCHED" if confidence == "" else "AMBIGUOUS"
                    r.confidence = confidence
                    r.note = "No unique blocking sleeper detected"
                    if alternatives:
                        r.note += "; alternatives=" + ",".join(x[1].handle for x in alternatives)
                    results.append(r)
                    continue
                _, sleeper, lx, ly, ad = match
                r.sleeper = sleeper.handle
                r.status = "BLOCKED" if abs(ly) <= sleeper.width_m / 2.0 + pipe.radius_m else "CLEAR"
                r.confidence = confidence
                r.sleeper_axis_deg = math.degrees(sleeper.rotation)
                r.angle_diff_deg = ad
                r.sleeper_x, r.sleeper_y = sleeper.center_world
                r.local_x, r.local_y = lx, ly
                r.sleeper_width_mm = sleeper.width_m * 1000.0
                r.current_clearance_mm = (abs(ly) - sleeper.width_m / 2.0 - pipe.radius_m) * 1000.0
                target = sleeper.width_m / 2.0 + pipe.radius_m + self.clearance_m
                move = max(0.0, target - abs(ly))
                r.required_move_mm = move * 1000.0
                r.final_clearance_mm = max(r.current_clearance_mm, self.clearance_m * 1000.0) if move == 0 else self.clearance_m * 1000.0
                sign = 1.0 if ly >= 0 else -1.0
                # Pick pipe normal that points away from the sleeper side.
                sleeper_y_world = (-math.sin(sleeper.rotation), math.cos(sleeper.rotation))
                na, nb = pipe.normal_a, pipe.normal_b
                da = na[0] * sleeper_y_world[0] + na[1] * sleeper_y_world[1]
                db = nb[0] * sleeper_y_world[0] + nb[1] * sleeper_y_world[1]
                toward_side = na if da >= db else nb
                if sign < 0:
                    toward_side = na if da <= db else nb
                r.move_dir_x, r.move_dir_y = toward_side
                r.old_x, r.old_y = pipe.center
                r.new_x = pipe.center[0] + toward_side[0] * move
                r.new_y = pipe.center[1] + toward_side[1] * move
                if r.status == "BLOCKED" and move > 0:
                    r.note = "Move required"
                else:
                    r.note = "No movement required"
                results.append(r)
            except Exception as exc:
                results.append(Result(drainage=raw["handle"], status="ERROR", note=f"{type(exc).__name__}: {exc}"))
        return results

    def _ensure_marker_layer(self):
        try:
            return self.doc.Layers.Item(MARKER_LAYER)
        except Exception:
            return self.doc.Layers.Add(MARKER_LAYER)

    def apply_moves(self, results, save=False):
        marker_layer = self._ensure_marker_layer()
        try:
            self.doc.StartUndoMark()
        except Exception:
            pass
        moved = 0
        for r in results:
            if r.status != "BLOCKED" or r.required_move_mm <= 0 or r.confidence not in {"HIGH", "MEDIUM"}:
                continue
            try:
                obj = retry(lambda h=r.drainage: self.doc.HandleToObject(h))
                dx = r.new_x - r.old_x
                dy = r.new_y - r.old_y
                retry(lambda obj=obj, dx=dx, dy=dy: obj.Move(
                    VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (0.0, 0.0, 0.0)),
                    VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (dx, dy, 0.0)),
                ))
                line = retry(lambda dx=dx, dy=dy, r=r: self.doc.ModelSpace.AddLine(
                    VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (r.old_x, r.old_y, 0.0)),
                    VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (r.new_x, r.new_y, 0.0)),
                ))
                line.Layer = MARKER_LAYER
                moved += 1
            except Exception as exc:
                r.status = "ERROR"
                r.note = f"Move failed: {type(exc).__name__}: {exc}"
        try:
            self.doc.Regen(1)
        except Exception:
            pass
        try:
            self.doc.EndUndoMark()
        except Exception:
            pass
        if save:
            retry(lambda: self.doc.Save())
        return moved
