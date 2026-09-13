import math
import time
from dataclasses import dataclass

import pythoncom
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
MOVED_LAYER = "CAD_ANALYSIS_MOVED"
MOVED_COLOR_INDEX = 1


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
    return (
        insertion[0] + x * c - y * s,
        insertion[1] + x * s + y * c,
    )


def world_to_local(x, y, insertion, rotation):
    dx = x - insertion[0]
    dy = y - insertion[1]
    c = math.cos(rotation)
    s = math.sin(rotation)
    return (dx * c + dy * s, -dx * s + dy * c)


def is_busy(exc):
    text = str(exc)
    return (
        "Call was rejected by callee" in text
        or "RPC_E_CALL_REJECTED" in text
        or "-2147418111" in text
    )


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
                if str(ent.ObjectName) != "AcDbLine" or str(ent.Layer) != "E Pipe":
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

                score = (
                    min_len
                    + 2.0 * ratio
                    - 20.0 * abs(sep - EXPECTED_PIPE_DIAMETER_M)
                )
                if best is None or score > best[0]:
                    best = (score, (a1, a2), (b1, b2), sep)

        if best is None:
            raise RuntimeError("Could not identify pipe side pair")

        _, line1, line2, sep = best
        p1 = (
            (line1[0][0] + line2[0][0]) / 2.0,
            (line1[0][1] + line2[0][1]) / 2.0,
        )
        p2 = (
            (line1[1][0] + line2[1][0]) / 2.0,
            (line1[1][1] + line2[1][1]) / 2.0,
        )
        return (p1, p2), sep

    @staticmethod
    def _entity_points(ent):
        name = str(ent.ObjectName)

        if name == "AcDbLine":
            a = tuple(ent.StartPoint)
            b = tuple(ent.EndPoint)
            return [
                (float(a[0]), float(a[1])),
                (float(b[0]), float(b[1])),
            ]

        if name in {"AcDbPolyline", "AcDb2dPolyline", "AcDb3dPolyline"}:
            try:
                coords = tuple(ent.Coordinates)
                if name == "AcDb3dPolyline":
                    return [
                        (float(coords[i]), float(coords[i + 1]))
                        for i in range(0, len(coords), 3)
                    ]
                return [
                    (float(coords[i]), float(coords[i + 1]))
                    for i in range(0, len(coords), 2)
                ]
            except Exception:
                pass

            points = []
            try:
                for vertex in ent:
                    coords = tuple(vertex.Coordinates)
                    if len(coords) >= 2:
                        points.append((float(coords[0]), float(coords[1])))
            except Exception:
                pass
            return points

        return []

    def _discover_sleeper_profile(self, block_name):
        block = self.doc.Blocks.Item(block_name)
        points = []
        entity_count = 0

        for ent in block:
            try:
                pts = self._entity_points(ent)
                if pts:
                    entity_count += 1
                    points.extend(pts)
            except Exception:
                continue

        if not points:
            raise RuntimeError(
                f"No usable point geometry in sleeper block {block_name}"
            )

        min_x = min(p[0] for p in points)
        min_y = min(p[1] for p in points)
        max_x = max(p[0] for p in points)
        max_y = max(p[1] for p in points)

        span_x = max_x - min_x
        span_y = max_y - min_y
        if span_x <= 1e-9 or span_y <= 1e-9:
            raise RuntimeError(
                f"Degenerate sleeper geometry in block {block_name}: "
                f"span=({span_x:.6f}, {span_y:.6f})"
            )

        return {
            "min_x": min_x,
            "min_y": min_y,
            "max_x": max_x,
            "max_y": max_y,
            "width_m": min(span_x, span_y),
            "length_m": max(span_x, span_y),
            "center_local": (
                (min_x + max_x) / 2.0,
                (min_y + max_y) / 2.0,
            ),
            "entity_count": entity_count,
        }

    def collect_instances(self):
        drainages = []
        sleepers = []

        model_space = retry(lambda: self.doc.ModelSpace)
        count = int(retry(lambda: model_space.Count))

        for index in range(count):
            try:
                obj = retry(lambda index=index: model_space.Item(index))
                if str(obj.ObjectName) != "AcDbBlockReference":
                    continue

                name = str(obj.EffectiveName)
                layer = str(obj.Layer)
                ins = tuple(obj.InsertionPoint)
                insertion = (float(ins[0]), float(ins[1]))
                rot = float(obj.Rotation)
                sx = float(obj.XScaleFactor)
                sy = float(obj.YScaleFactor)
                handle = str(obj.Handle)

                if name == DRAINAGE_BLOCK:
                    # Never re-analyze copies produced by this tool.
                    if layer == MOVED_LAYER:
                        continue
                    drainages.append({
                        "handle": handle,
                        "obj": obj,
                        "insertion": insertion,
                        "rotation": rot,
                        "sx": sx,
                        "sy": sy,
                    })

                elif layer == SLEEPER_LAYER and name in SLEEPER_BLOCKS:
                    profile = self.sleeper_profiles[name]
                    center_local = profile["center_local"]
                    center_world = local_to_world(
                        center_local[0],
                        center_local[1],
                        insertion,
                        rot,
                        sx,
                        sy,
                    )
                    scale = max(abs(sx), abs(sy))
                    sleepers.append(
                        Sleeper(
                            handle,
                            obj,
                            name,
                            insertion,
                            rot,
                            profile["width_m"] * scale,
                            profile["length_m"] * scale,
                            center_local,
                            center_world,
                        )
                    )

            except Exception as exc:
                if is_busy(exc):
                    pythoncom.PumpWaitingMessages()
                    time.sleep(COM_DELAY)
                continue

        return drainages, sleepers

    @staticmethod
    def _grid_key(p, size):
        return (math.floor(p[0] / size), math.floor(p[1] / size))

    def _build_grid(self, sleepers):
        grid = {}
        for sleeper in sleepers:
            key = self._grid_key(sleeper.center_world, GRID_SIZE_M)
            grid.setdefault(key, []).append(sleeper)
        return grid

    def extract_pipe(self, raw, scale_radius=True):
        ins = raw["insertion"]
        rot = raw["rotation"]
        sx = float(raw["sx"])
        sy = float(raw["sy"])
        lp1, lp2 = self.centerline_local

        p1 = local_to_world(lp1[0], lp1[1], ins, rot, sx, sy)
        p2 = local_to_world(lp2[0], lp2[1], ins, rot, sx, sy)
        center = (
            (p1[0] + p2[0]) / 2.0,
            (p1[1] + p2[1]) / 2.0,
        )
        ux, uy = normalize(p2[0] - p1[0], p2[1] - p1[1])
        scale = (abs(sx) + abs(sy)) / 2.0 if scale_radius else 1.0

        return Pipe(
            raw["handle"],
            raw["obj"],
            ins,
            center,
            p1,
            p2,
            orientation(p1, p2),
            self.pipe_radius_local * scale,
            (-uy, ux),
            (uy, -ux),
        )

    def match_sleeper(self, pipe, grid, sleepers):
        key = self._grid_key(pipe.center, GRID_SIZE_M)
        candidates = []

        for ix in range(key[0] - 1, key[0] + 2):
            for iy in range(key[1] - 1, key[1] + 2):
                for sleeper in grid.get((ix, iy), []):
                    lx, ly = world_to_local(
                        pipe.center[0],
                        pipe.center[1],
                        sleeper.center_world,
                        sleeper.rotation,
                    )
                    ad = angle_diff(
                        pipe.axis_deg,
                        math.degrees(sleeper.rotation),
                    )
                    half_w = sleeper.width_m / 2.0
                    half_l = sleeper.length_m / 2.0

                    if ad > SLEEPER_ANGLE_TOLERANCE_DEG:
                        continue
                    if abs(lx) > half_l + MATCH_MARGIN_M:
                        continue
                    if abs(ly) > half_w + pipe.radius_m + MATCH_MARGIN_M:
                        continue

                    score = abs(ly) + 0.25 * abs(lx) + 0.02 * ad
                    candidates.append((score, sleeper, lx, ly, ad))

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
                match, confidence, alternatives = self.match_sleeper(
                    pipe,
                    grid,
                    sleepers,
                )

                result = Result(
                    drainage=pipe.handle,
                    pipe_axis_deg=pipe.axis_deg,
                    pipe_x=pipe.center[0],
                    pipe_y=pipe.center[1],
                    pipe_diameter_mm=pipe.radius_m * 2000.0,
                )

                if match is None:
                    result.status = (
                        "UNMATCHED" if confidence == "" else "AMBIGUOUS"
                    )
                    result.confidence = confidence
                    result.note = "No unique blocking sleeper detected"
                    if alternatives:
                        result.note += "; alternatives=" + ",".join(
                            item[1].handle for item in alternatives
                        )
                    results.append(result)
                    continue

                _, sleeper, lx, ly, ad = match
                result.sleeper = sleeper.handle
                result.status = (
                    "BLOCKED"
                    if abs(ly) <= sleeper.width_m / 2.0 + pipe.radius_m
                    else "CLEAR"
                )
                result.confidence = confidence
                result.sleeper_axis_deg = math.degrees(sleeper.rotation)
                result.angle_diff_deg = ad
                result.sleeper_x, result.sleeper_y = sleeper.center_world
                result.local_x, result.local_y = lx, ly
                result.sleeper_width_mm = sleeper.width_m * 1000.0
                result.current_clearance_mm = (
                    abs(ly) - sleeper.width_m / 2.0 - pipe.radius_m
                ) * 1000.0

                target = (
                    sleeper.width_m / 2.0
                    + pipe.radius_m
                    + self.clearance_m
                )
                move = max(0.0, target - abs(ly))
                result.required_move_mm = move * 1000.0
                result.final_clearance_mm = (
                    max(
                        result.current_clearance_mm,
                        self.clearance_m * 1000.0,
                    )
                    if move == 0
                    else self.clearance_m * 1000.0
                )

                sign = 1.0 if ly >= 0 else -1.0
                sleeper_y_world = (
                    -math.sin(sleeper.rotation),
                    math.cos(sleeper.rotation),
                )
                na, nb = pipe.normal_a, pipe.normal_b
                da = na[0] * sleeper_y_world[0] + na[1] * sleeper_y_world[1]
                db = nb[0] * sleeper_y_world[0] + nb[1] * sleeper_y_world[1]
                move_dir = na if da >= db else nb
                if sign < 0:
                    move_dir = na if da <= db else nb

                result.move_dir_x, result.move_dir_y = move_dir
                result.old_x, result.old_y = pipe.center
                result.new_x = pipe.center[0] + move_dir[0] * move
                result.new_y = pipe.center[1] + move_dir[1] * move

                if result.status == "BLOCKED" and move > 0:
                    result.note = "Move required; original will be preserved"
                else:
                    result.note = "No movement required"

                results.append(result)

            except Exception as exc:
                results.append(
                    Result(
                        drainage=raw["handle"],
                        status="ERROR",
                        note=f"{type(exc).__name__}: {exc}",
                    )
                )

        return results

    def _ensure_moved_layer(self):
        try:
            layer = retry(lambda: self.doc.Layers.Item(MOVED_LAYER))
        except Exception:
            layer = retry(lambda: self.doc.Layers.Add(MOVED_LAYER))

        try:
            layer.Color = MOVED_COLOR_INDEX
        except Exception:
            pass
        return layer

    def apply_moves(self, results, save=False):
        moved = 0
        moved_layer = self._ensure_moved_layer()

        try:
            self.doc.StartUndoMark()
        except Exception:
            pass

        for result in results:
            if (
                result.status != "BLOCKED"
                or result.required_move_mm <= 0
                or result.confidence not in {"HIGH", "MEDIUM"}
            ):
                continue

            try:
                original = retry(
                    lambda handle=result.drainage: self.doc.HandleToObject(handle)
                )
                dx = result.new_x - result.old_x
                dy = result.new_y - result.old_y

                # Preserve the original object. Move only its copied reference.
                copied = retry(lambda original=original: original.Copy())
                retry(
                    lambda copied=copied, dx=dx, dy=dy: copied.Move(
                        VARIANT(
                            pythoncom.VT_ARRAY | pythoncom.VT_R8,
                            (0.0, 0.0, 0.0),
                        ),
                        VARIANT(
                            pythoncom.VT_ARRAY | pythoncom.VT_R8,
                            (dx, dy, 0.0),
                        ),
                    )
                )

                try:
                    copied.Layer = MOVED_LAYER
                except Exception:
                    pass
                try:
                    copied.Color = MOVED_COLOR_INDEX
                except Exception:
                    pass

                result.note = (
                    f"Moved copy created; original preserved; "
                    f"move={result.required_move_mm:.1f} mm; "
                    f"direction=({result.move_dir_x:.4f}, {result.move_dir_y:.4f})"
                )
                moved += 1

            except Exception as exc:
                result.status = "ERROR"
                result.note = (
                    f"Move copy failed: {type(exc).__name__}: {exc}"
                )

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
