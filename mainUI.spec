# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['mainUI.py'],
             pathex=['/Users/coyote/PycharmProjects/untitled/venv/xmindtoxls/lib/python3.9/site-packages', '/Users/coyote/Desktop/git_repository/xmindtoxls'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
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
          name='mainUI',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
app = BUNDLE(exe,
             name='mainUI.app',
             icon=None,
             bundle_identifier=None)
