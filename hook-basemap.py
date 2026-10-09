# hook-basemap.py
# Runtime hook: runs at the start of the frozen app to fix Basemap's data path.

import os
import sys

if getattr(sys, 'frozen', False):
    # sys._MEIPASS is the directory where PyInstaller extracts bundled data
    base = sys._MEIPASS
    basemap_data = os.path.join(base, 'mpl_toolkits', 'basemap_data')

    # Basemap looks for its data in basemap_datadir (module attribute) and
    # respects the BASEMAPDATA environment variable.
    os.environ['BASEMAPDATA'] = basemap_data

    # Also patch the module attribute directly, in case the module has
    # already been imported by the time this hook runs.
    try:
        import mpl_toolkits.basemap as _bm
        _bm.basemap_datadir = basemap_data
    except Exception:
        # If Basemap hasn't been imported yet, the env var above will
        # be picked up when it is.
        pass