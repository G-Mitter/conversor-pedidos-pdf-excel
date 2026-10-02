# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files


a = Analysis(
    ['app_gui.py'],
    pathex=[],
    binaries=[],
    # modelos .onnx, config.yaml e os .py do RapidOCR: ele carrega os submódulos pelo nome
    # em tempo de execução, então sem os arquivos soltos o OCR não inicia dentro do exe
    datas=collect_data_files('rapidocr_onnxruntime', include_py_files=True) + [('icone.ico', '.')],
    # dependências que os .py soltos do RapidOCR importam (o PyInstaller não as enxerga sozinho)
    hiddenimports=['pyclipper', 'shapely', 'shapely.geometry', 'six', 'yaml', 'cv2', 'onnxruntime'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='Conversor_de_Pedidos',
    icon='icone.ico',
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
)
