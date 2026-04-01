# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['login.py'],
             pathex=[],
             binaries=[],
             datas=[('template', 'template'), ('favicon.ico', 'img'), ('E:/Anaconda3/Lib/site-packages/docxcompose/templates/*.*', 'docxcompose/templates')],
             hiddenimports=['docx', 'pandas', 'docxcompose', 'lxml', 'lxml._elementpath'],
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
          name='gui_output',
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
          entitlements_file=None, icon='favicon.ico')
