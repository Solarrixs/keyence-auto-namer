import os
import logging
from pywinauto.application import Application
from pywinauto import Desktop

# File paths
LOG_FILE = 'keyence_auto_namer.log'
IMAGE_PATH = os.path.join(os.path.dirname(__file__), 'image.png')

# Window titles
WIDE_IMAGE_VIEWER_TITLE = 'BZ-X800 Wide Image Viewer'
ANALYZER_TITLE = 'BZ-X800 Analyzer'

# Other constants
MAX_DELAY_TIME = 20
REQUIRED_FIELDS = ['Run Name', 'Stitch Type', 'Overlay', 'Naming Template', 'Filepath', 'XY Name']

# Logging setup
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler(LOG_FILE),
                    ])

def get_main_window():
    try:
        app = Application(backend="uia").connect(title=ANALYZER_TITLE)
        main_window = app.window(title=ANALYZER_TITLE)
        return main_window
    except Exception:
        logging.error("BZ-X800 Analyzer not found. Open the program and press enter to try again.")
        input()
        return get_main_window()

def get_desktop():
    return Desktop(backend="uia")