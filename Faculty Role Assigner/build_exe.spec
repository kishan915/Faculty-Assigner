# ── Faculty Role Assigner — PyInstaller spec ─────────────────────────────────
# Run this on a Windows machine:
#   pip install pyinstaller flask pandas openpyxl
#   pyinstaller build_exe.spec
#
# Output: dist\FacultyRoleAssigner.exe  (double-click to run)

block_cipher = None

a = Analysis(
    ['launcher.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('app_prot.py', '.'),
    ],
    hiddenimports=[
        'flask',
        'werkzeug',
        'werkzeug.serving',
        'werkzeug.routing',
        'werkzeug.exceptions',
        'werkzeug.middleware',
        'jinja2',
        'jinja2.ext',
        'click',
        'itsdangerous',
        'markupsafe',
        'pandas',
        'pandas.io.formats.style',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.writer.excel',
        'openpyxl.reader.excel',
        'et_xmlfile',
        'webbrowser',
        'threading',
        'socket',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'PIL', 'tkinter', 'PyQt5'],
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='FacultyRoleAssigner',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,        # keeps a console window so errors are visible
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
