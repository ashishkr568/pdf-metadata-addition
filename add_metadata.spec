# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['add_metadata.py'],
             pathex=[],
             binaries=[],
             datas=[],
             hiddenimports=[],
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

import shutil
import os
spec_root = os.path.realpath(SPECPATH)
shutil.rmtree(spec_root+'\dist', ignore_errors=False, onerror=None)
shutil.copytree(spec_root+'/img', spec_root+'/dist/img/',dirs_exist_ok=False)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,  
          [],
          name='PDF Metadata Addition Tool',
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
          entitlements_file=None , icon='img\\appIcon.ico')
