# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

import tkinterdnd2


PROJECT_DIR = Path(SPECPATH)
TKDND_ROOT = Path(tkinterdnd2.__file__).resolve().parent / "tkdnd"
ICON_PATH = PROJECT_DIR / "icon.ico"

if not ICON_PATH.exists():
    raise SystemExit(f"Build icon not found: {ICON_PATH}")

datas = [(str(ICON_PATH), ".")]
binaries = []
hiddenimports = ["tkinterdnd2.TkinterDnD"]

for arch_name in ("win-arm64", "win-x64", "win-x86"):
    arch_dir = TKDND_ROOT / arch_name
    if not arch_dir.exists():
        continue

    dest_dir = f"tkinterdnd2/tkdnd/{arch_name}"
    for item in arch_dir.iterdir():
        if not item.is_file():
            continue
        if item.suffix.lower() == ".dll":
            binaries.append((str(item), dest_dir))
        else:
            datas.append((str(item), dest_dir))

EXCLUDES = [
    "IPython",
    "PIL.ImageQt",
    "PyQt5",
    "PyQt6",
    "PySide2",
    "PySide6",
    "astropy",
    "bokeh",
    "cv2",
    "dask",
    "datasets",
    "django",
    "docutils",
    "fastapi",
    "gradio",
    "holoviews",
    "imageio",
    "ipykernel",
    "ipywidgets",
    "jax",
    "jaxlib",
    "jedi",
    "jinja2",
    "joblib",
    "jupyter",
    "jupyter_client",
    "jupyter_core",
    "keras",
    "llvmlite",
    "matplotlib",
    "mistune",
    "moviepy",
    "nbconvert",
    "nbformat",
    "networkx",
    "nltk",
    "notebook",
    "numba",
    "onnxruntime",
    "paddle",
    "pandas",
    "panel",
    "plotly",
    "prompt_toolkit",
    "pyarrow",
    "pycocotools",
    "pygments",
    "pytest",
    "qtpy",
    "scipy",
    "seaborn",
    "sentencepiece",
    "skimage",
    "sklearn",
    "spacy",
    "sqlalchemy",
    "streamlit",
    "sympy",
    "tensorboard",
    "tensorflow",
    "tokenizers",
    "torch",
    "torchaudio",
    "torchvision",
    "tqdm",
    "transformers",
    "twisted",
    "xarray",
    "zmq",
]


a = Analysis(
    ["main.py"],
    pathex=[str(PROJECT_DIR)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="AcademicFigureCropper",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(ICON_PATH)],
)
