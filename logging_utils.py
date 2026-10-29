import os
import os
import sys
import logging
import traceback
from sys import platform as _platform

def setup_logging_to_file(filename, level=logging.ERROR):
    global logger
    logger = logging.getLogger(__name__)
    logger.setLevel(level)
    handler = logging.FileHandler(filename, mode='w', delay=True)
    handler.setLevel(level)
    logger.addHandler(handler)
        
def handle_errors(eclass, e, etrace):
    global logger
    logger.error('Traceback (most recent call last):')
    efile, eline, efunc, esource = traceback.extract_tb(etrace)[-1]
    logger.error('  File "%s", line %d in %s\n%s: %s' % (os.path.basename(efile), eline, efunc, eclass.__name__, str(e)))
    if _platform == "win32" or _platform == "win64":
        import atexit
        import win32api
        atexit.register(win32api.MessageBox, 0,
            'See the logfile "%s.log" for details' % sys.executable,
            'Errors occurred')