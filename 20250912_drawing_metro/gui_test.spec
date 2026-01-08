
# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['gui_test.py'],
             pathex=[],
             binaries=[],
             datas=[('favicon.icns','img')],
             hiddenimports=[
             'matplotlib',
             'matplotlib.backends.backend_agg',
             'matplotlib.backends.backend_tkagg',
             'matplotlib.pyplot',
             'numpy',
             'PIL',
             'tkinter',
             'docx',
             'pandas',
             'six',
             'packaging',
             'pyparsing',
             'dateutil',
             'cycler',
             'kiwisolver',
             ],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,  
          [],
          name='地铁线路绘图软件',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None , icon='favicon.ico')
app = BUNDLE(
    exe,
    name='地铁线路绘图软件.app',
    icon='favicon.icns',  # macOS 图标文件
    bundle_identifier='com.wty.plot',
    info_plist={
        'NSHighResolutionCapable': 'True',
        'LSMinimumSystemVersion': '10.13.0',
        'CFBundleShortVersionString': '1.0.0',
        'CFBundleVersion': '1.0.0',
    },
)
