pip install nuitka

python -m nuitka --onefile --follow-imports ^
  --lto=yes ^
  --show-memory --show-progress ^
  --plugin-enable=numpy ^
  --plugin-enable=pyqt5 ^
  --include-package=pywinauto ^
  --include-package=pyautogui ^
  --include-package=comtypes ^
  --include-module=comtypes.stream ^
  --noinclude-unittest ^
  --nofollow-import-to=tkinter,numpy
  --remove-output ^
  --assume-yes-for-downloads ^
  --windows-disable-console ^
  --windows-company-name="Corder Labs" ^
  --windows-product-name="Keyence Auto Namer" ^
  --windows-file-version=1.0.2 ^
  --windows-product-version=1.0.2 ^
  main.py

pyinstaller --noconfirm --onefile --console --clean ^
  --hidden-import "pywinauto" ^
  --hidden-import "pyautogui" ^
  --hidden-import "comtypes" ^
  --hidden-import "comtypes.stream" ^
  --exclude-module "tkinter" ^
  --exclude-module "numpy" ^
  --exclude-module "matplotlib" ^
  --exclude-module "pandas" ^
  --exclude-module "scipy" ^
  --exclude-module "PyQt5" ^
  --exclude-module "PySide2" ^
  --exclude-module "wx" ^
  --exclude-module "pyglet" ^
  --exclude-module "OpenGL" ^
  --exclude-module "PyCrypto" ^
  --exclude-module "cryptography" ^
  --exclude-module "asyncio" ^
  --exclude-module "sqlite3" ^
  --exclude-module "docutils" ^
  --exclude-module "pydoc" ^
  --exclude-module "unittest" ^
  --exclude-module "distutils" ^
  --exclude-module "setuptools" ^
  --exclude-module "email" ^
  --exclude-module "html" ^
  --exclude-module "http" ^
  --exclude-module "xml" ^
  --exclude-module "pydoc_data" ^
  main.py