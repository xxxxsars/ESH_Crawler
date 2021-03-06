# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['main.py'],
             pathex=['E:\\PycharmProjects\\ESH_Crawler'],
             binaries=[],
             datas=[("package_lib/*",".")],
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
          name='main',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )


import shutil
shutil.copyfile('setting_sample.ini', 'dist/setting.ini')
shutil.copyfile('chromedriver.exe', 'dist/chromedriver.exe')
shutil.copyfile('keyword.json', 'dist/keyword.json')

if not os.path.exists('dist/templates'):
    shutil.copytree('templates', 'dist/templates')