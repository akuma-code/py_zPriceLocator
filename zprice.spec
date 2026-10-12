# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec для приложения расчёта цен.
Сборка:
    pyinstaller zprice.spec
Результат:
    dist/zprice.exe  (Windows)
    dist/zprice      (Linux/macOS)
"""

from PyInstaller.utils.hooks import collect_submodules, collect_data_files


block_cipher = None


# ---------- Скрытые импорты ----------
# pandas и openpyxl активно используют динамические импорты,
# поэтому их нужно явно перечислить.
hiddenimports = [
    # pandas
    "pandas",
    "pandas._libs",
    "pandas._libs.tslibs",
    "pandas._libs.tslibs.base",
    "pandas._libs.tslibs.conversion",
    "pandas._libs.tslibs.parsing",
    "pandas._libs.tslibs.timedeltas",
    "pandas._libs.tslibs.timestamps",
    "pandas.io.formats.style",
    "pandas.io.excel._base",
    "pandas.io.excel._openpyxl",
    "pandas.io.excel._xlrd",
    # openpyxl (xlsx)
    "openpyxl",
    "openpyxl.cell._writer",
    # xlrd (xls)
    "xlrd",
    # tkinter (обычно подхватывается сам, но на всякий случай)
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "tkinter.font",
]

# Дополнительно подтянем все подмодули openpyxl и xlrd
hiddenimports += collect_submodules("openpyxl")
hiddenimports += collect_submodules("xlrd")

# ---------- Данные (json и др.) ----------
# Если хотите, чтобы groups.json и adresses.json были рядом с .exe —
# раскомментируйте datas ниже и передайте их внутрь бандла.
# В текущей конфигурации они читаются из рабочей директории,
# поэтому просто положите их рядом с собранным exe.
datas = [
     ("groups.json", "."),
     ("adresses.json", "."),
]

# Собираем пакеты данных pandas/openpyxl (нужны для Excel)
datas += collect_data_files("openpyxl")
datas += collect_data_files("pandas")

# ---------- Анализ ----------
a = Analysis(
    ["main.py"],                # основной файл
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # уменьшаем размер — выкидываем ненужные тяжелые модули
        "matplotlib",
        "scipy",
        "PyQt5",
        "PyQt6",
        "PySide2",
        "PySide6",
        "IPython",
        "jupyter",
        "notebook",
        "pytest",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ---------- Одним файлом (--onefile) ----------
# exe = EXE(
#     pyz,
#     a.scripts,
#     a.binaries,
#     a.zipfiles,
#     a.datas,
#     [],
#     name="zprice",
#     debug=False,
#     bootloader_ignore_signals=False,
#     strip=False,
#     upx=True,
#     upx_exclude=[
#         # UPX не трогаем эти библиотеки, иначе ломается импорт
#         "vcruntime140.dll",
#         "python3.dll",
#         "python3*.dll",
#         "tcl86t.dll",
#         "tk86t.dll",
#     ],
#     runtime_tmpdir=None,
#     console=False,          # окно без консоли
#     disable_windowed_traceback=False,
#     argv_emulation=False,
#     target_arch=None,
#     codesign_identity=None,
#     entitlements_file=None,
#     # icon="icon.ico",      # ← если есть иконка — раскомментируйте
# )

# ---------- Вариант "папка" (--onedir) ----------
# Если onefile-вариант запускается слишком долго (распаковка pandas
# каждый раз занимает время), закомментируйте блок EXE выше
# и раскомментируйте COLLECT ниже:
#
exe = EXE(
     pyz,
     a.scripts,
     [],
     exclude_binaries=True,
     name="zprice",
     debug=False,
     bootloader_ignore_signals=False,
     strip=False,
     upx=True,
     console=False,
     disable_windowed_traceback=False,
     argv_emulation=False,
     target_arch=None,
     codesign_identity=None,
     entitlements_file=None,
     icon="",
 )

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="zprice",
)