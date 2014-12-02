# -*- mode: python -*-
a = Analysis(['paster.py'],
             pathex=['C:\\Users\\Chad\\mystuff\\projects\\crn paster'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None)
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='paster.exe',
          debug=False,
          strip=None,
          upx=True,
          console=False , icon='clip.ico')
