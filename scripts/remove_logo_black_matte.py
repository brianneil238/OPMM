"""
One-off / repeatable: remove solid black (or near-black) matte from bsu_neu_logo.png
by flood-filling transparency from image edges. Preserves the seal interior.
"""
from __future__ import annotations

from collections import deque
from pathlib import Path

from PIL import Image


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    path = root / "core" / "static" / "core" / "img" / "bsu_neu_logo.png"
    if not path.is_file():
        raise SystemExit(f"Missing logo: {path}")

    img = Image.open(path).convert("RGBA")
    w, h = img.size
    px = img.load()

    def is_matte(r: int, g: int, b: int, a: int) -> bool:
        if a < 10:
            return True
        # Near-black / very dark matte only (seal colors stay above this)
        return max(r, g, b) < 78 and (r + g + b) < 200

    seen = set()
    q: deque[tuple[int, int]] = deque()

    def push(x: int, y: int) -> None:
        if 0 <= x < w and 0 <= y < h and (x, y) not in seen:
            seen.add((x, y))
            q.append((x, y))

    for x in range(w):
        push(x, 0)
        push(x, h - 1)
    for y in range(h):
        push(0, y)
        push(w - 1, y)

    while q:
        x, y = q.popleft()
        r, g, b, a = px[x, y]
        if not is_matte(r, g, b, a):
            continue
        px[x, y] = (r, g, b, 0)
        for nx, ny in ((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)):
            if 0 <= nx < w and 0 <= ny < h and (nx, ny) not in seen:
                seen.add((nx, ny))
                q.append((nx, ny))

    bak = path.with_name("bsu_neu_logo_bak.png")
    if not bak.exists():
        import shutil

        shutil.copy2(path, bak)

    img.save(path, format="PNG", optimize=True)
    print(f"Wrote transparent matte: {path} (backup at {bak.name} if first run)")


if __name__ == "__main__":
    main()
