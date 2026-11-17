"""The Format commands of VB6's form designer without Qt (phase 8.4), as the VB6 IDE did them (VB6 oracle
DEV-8/fd1). Rectangles are (Left, Top, Width, Height) in twips; 15 twips are a pixel.

  Horizontal / Vertical Spacing: the controls in the order of their Left (Top); the first stays.
      Make Equal: the first and the last stay, the gaps between them become equal (twips; each Left is cut
                  to a whole twip: 1240 + 586.67 -> 1826)
      Increase / Decrease: in pixels: each control moves by its place in the order times the grid (8 px for 120
                  twips): Left 1820 (121.33 px -> 121) + 8 -> 129 px = 1935
      Remove: in pixels, each control against the one before it: 16 px + 67 px (1000 twips) -> 83 px = 1245
  Size to Grid: Left and Top to the nearest grid line, the right edge to the nearest grid line + 1 pixel; the
      Height stays (VB6 does not change it)
  Ctrl+arrow: moves by the grid (in pixels), 1 pixel without Align Controls to Grid; Shift+arrow sizes the same
      way; both also for locked controls
  arrow alone: the next control in that direction: one whose gap from this one along that axis is larger than
      the gap along the other axis (0 when they overlap), the nearest one
"""
import math

TPP = 15


def px(tw):
    """twips to whole pixels, a half up (1000 -> 67)"""
    return int(math.floor(tw / TPP + 0.5))


def snap(v, grid):
    """to the nearest grid line, a half up (VB6: 1260 on a 120 grid -> 1320)"""
    if grid <= 0:
        return v
    return int(math.floor(v / grid + 0.5)) * grid


def spacing(rects, vertical, how, grid):
    """Format > Horizontal / Vertical Spacing: how is "equal", "increase", "decrease" or "remove"; grid is the
    grid in twips (Width or Height of Tools > Options). The new rectangles, in the order given"""
    a, s = (1, 3) if vertical else (0, 2)
    out = [list(r) for r in rects]
    n = len(rects)
    if n < 2:
        return [tuple(r) for r in out]
    order = sorted(range(n), key=lambda i: (rects[i][a], i))
    if how == "equal":
        if n >= 3:
            first, last, mid = order[0], order[-1], order[1:-1]
            cur = rects[first][a] + rects[first][s]
            gap = (rects[last][a] - cur - sum(rects[i][s] for i in mid)) / (n - 1)
            for i in mid:
                p = int(math.floor(cur + gap))
                out[i][a] = p
                cur = p + rects[i][s]
    elif how in ("increase", "decrease"):
        g = px(grid) * (1 if how == "increase" else -1)
        for k, i in enumerate(order[1:], 1):
            out[i][a] = (px(rects[i][a]) + k * g) * TPP
    elif how == "remove":
        first = order[0]
        cur = px(rects[first][a]) + px(rects[first][s])
        for i in order[1:]:
            out[i][a] = cur * TPP
            cur += px(rects[i][s])
    return [tuple(r) for r in out]


def size_to_grid(rect, gx, gy):
    """Format > Size to Grid of one control"""
    x, y, w, h = rect
    nx, ny = snap(x, gx), snap(y, gy)
    return (nx, ny, max(TPP, snap(x + w, gx) - nx + TPP), h)


def nudge(v, sign, grid, align, size=False):
    """Ctrl+arrow (a position) / Shift+arrow (a size): by the grid, or by a pixel without Align Controls to
    Grid (360 -> 375); a size stays at least a pixel"""
    step = px(grid) if align else 1
    n = px(v) + sign * step
    return max(1, n) * TPP if size else n * TPP


def _gaps(a, b):
    """the horizontal and the vertical gap between two rectangles (0 where they overlap)"""
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    hg = bx - (ax + aw) if bx >= ax + aw else (ax - (bx + bw) if bx + bw <= ax else 0)
    vg = by - (ay + ah) if by >= ay + ah else (ay - (by + bh) if by + bh <= ay else 0)
    return hg, vg


def neighbor(rects, i, direction):
    """the arrow keys alone: the index of the control they select from rects[i], or None (it stays)"""
    me = rects[i]
    best = None
    for j, r in enumerate(rects):
        if j == i:
            continue
        hg, vg = _gaps(me, r)
        if direction in ("left", "right"):
            side = r[0] >= me[0] + me[2] if direction == "right" else r[0] + r[2] <= me[0]
            if not side or hg <= vg:
                continue
            key = (hg, vg)
        else:
            side = r[1] >= me[1] + me[3] if direction == "down" else r[1] + r[3] <= me[1]
            if not side or vg <= hg:
                continue
            key = (vg, hg)
        if best is None or key < best[0]:
            best = (key, j)
    return best[1] if best else None
