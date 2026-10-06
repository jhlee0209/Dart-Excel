"""'DART 데이터북.app'을 만든다.

사용: .venv.nosync/bin/python tools/build_mac_app.py

앱은 이 프로젝트 폴더와 가상환경을 그대로 쓰는 실행기다(파이썬을 통째로 담지 않는다).
그래서 앱 파일은 어디로 옮겨도 되지만, 프로젝트 폴더를 옮기면 이 스크립트를 다시 돌려야 한다.
"""
from __future__ import annotations

import plistlib
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
APP = ROOT / "dist" / "DART 데이터북.app"
PYTHON = ROOT / ".venv.nosync" / "bin" / "python"
NAVY, BRONZE, IVORY = (31, 58, 95), (201, 164, 108), (250, 248, 244)


def draw_icon(size: int = 1024) -> Image.Image:
    """짙은 네이비 바탕에 오르는 막대 세 개."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    pad = size * 0.09
    d.rounded_rectangle([pad, pad, size - pad, size - pad], radius=size * 0.2, fill=NAVY)
    base, left, width, gap = size * 0.72, size * 0.27, size * 0.115, size * 0.055
    for i, (height, color) in enumerate([(0.17, IVORY), (0.29, IVORY), (0.43, BRONZE)]):
        x = left + i * (width + gap)
        d.rounded_rectangle([x, base - size * height, x + width, base], radius=size * 0.022, fill=color)
    d.rounded_rectangle([left - size * 0.02, base + size * 0.035, left + 3 * width + 2 * gap + size * 0.02,
                         base + size * 0.055], radius=size * 0.01, fill=(*IVORY, 150))
    return img


def make_icns(target: Path) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        iconset = Path(tmp) / "icon.iconset"
        iconset.mkdir()
        master = draw_icon()
        for size in (16, 32, 128, 256, 512):
            master.resize((size, size), Image.LANCZOS).save(iconset / f"icon_{size}x{size}.png")
            master.resize((size * 2, size * 2), Image.LANCZOS).save(iconset / f"icon_{size}x{size}@2x.png")
        subprocess.run(["iconutil", "-c", "icns", str(iconset), "-o", str(target)], check=True)


def main() -> None:
    if not PYTHON.exists():
        sys.exit(f"가상환경이 없습니다: {PYTHON}")
    if APP.exists():
        shutil.rmtree(APP)
    (APP / "Contents" / "MacOS").mkdir(parents=True)
    (APP / "Contents" / "Resources").mkdir()

    launcher = APP / "Contents" / "MacOS" / "launcher"
    launcher.write_text(
        "#!/bin/bash\n"
        f'ROOT="{ROOT}"\n'
        'if [ ! -x "$ROOT/.venv.nosync/bin/python" ]; then\n'
        "  osascript -e 'display alert \"DART 데이터북을 열 수 없습니다\" message \"프로젝트 폴더나 가상환경을 찾지 못했습니다. "
        "tools/build_mac_app.py를 다시 실행해 주세요.\"'\n"
        "  exit 1\n"
        "fi\n"
        'cd "$ROOT" || exit 1\n'
        'exec "$ROOT/.venv.nosync/bin/python" "$ROOT/desktop.py"\n',
        encoding="utf-8",
    )
    launcher.chmod(0o755)

    with (APP / "Contents" / "Info.plist").open("wb") as f:
        plistlib.dump({
            "CFBundleName": "DART 데이터북",
            "CFBundleDisplayName": "DART 데이터북",
            "CFBundleIdentifier": "com.jhlee.dart-databook",
            "CFBundleExecutable": "launcher",
            "CFBundleIconFile": "icon",
            "CFBundlePackageType": "APPL",
            "CFBundleShortVersionString": "1.0",
            "CFBundleVersion": "1",
            "LSMinimumSystemVersion": "11.0",
            "NSHighResolutionCapable": True,
        }, f)
    make_icns(APP / "Contents" / "Resources" / "icon.icns")
    print(f"만들었습니다: {APP}")


if __name__ == "__main__":
    main()
