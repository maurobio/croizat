#! /usr/bin/python
# -*- coding: utf-8 -*-
#===================================================================================#
#  Croizat - A Software Package for Quantitative Analysis in Panbiogeography        #
#                 Copyright 2003-2026 Mauro J. Cavalcanti                           #
#                           maurobio@gmail.com                                      #
#                                                                                   #
#  This program is free software; you can redistribute it and/or modify             #
#  it under the terms of the GNU General Public License as published by             #
#  the Free Software Foundation; either version 3 of the License, or                #
#  (at your option) any later version.                                              #
#                                                                                   # 
#  This program is distributed in the hope that it will be useful,                  #
#  but WITHOUT ANY WARRANTY; without even the implied warranty of                   #
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                    #
#  GNU General Public License for more details.                                     #
#                                                                                   #
#  You should have received a copy of the GNU General Public License                #
#  along with this program. If not, see <http://www.gnu.org/licenses/>.             #
#                                                                                   #
#  Requirements:                                                                    #
#    Python 3.14+ (www.python.org)                                                  #
#    PyQt 6.11+ (www.riverbankcomputing.com/software/pyqt)                          #
#    PyShp 3.1.6+ (github.com/GeospatialPython/pyshp)                               #
#    NumPy 2.5.3+ (www.numpy.org)                                                   #
#    Matplotlib 3.1.1+ (www.matplotlib.org)                                         #
#    Basemap 2.0.0+ (www.matplotlib.org/basemap)                                    # 
#    simplekml 1.3+ (www.simplekml.com)                                             #
#    xlrd 1.1+ (www.python-excel.org)                                               #
#    ezodf 0.2+ (pythonhosted.org/ezodf)                                            #
#    xlwt 0.7+ (www.python-excel.org)                                               # 
#    xlsxwriter 1.0+ (github.com/jmcnamara/XlsxWriter)                              #
#    dbfpy3 4.2.3+ (github.com/frankyxhl/dbfpy3)                                    # 
#    openpyxl 3.1.5+ (foss.heptapod.net/openpyxl/openpyxl)                          #
#                                                                                   #
#  REVISION HISTORY:                                                                #
#    Version 3.00, 26 Oct 2026 - First public release ("Vancouver")                 #
#      * Added mapping of main msssings with choice of grid size and color palette  #
#      * Added cross-platform (Linux/Windows) call to MartiTracks/Panbiotracks      #
#      * Added GeoJSON format support for point and line data                       #
#      * Added Panbiotracks integration for generating generalized tracks and nodes #
#      * Added confirmation dialog for removing files from project                  #
#      * Added option for creating incidence matrices for PAE analysis              #
#      * Changed shortcut for adding files to Ctrl+Ins                              #
#      * Changed shortcut for deleting files to Ctrl+Del                            #
#      * Changed file reading for better handling of table formatted data           #
#      * Removed Import option from the main menu                                   $
#      * Fixed many minor bugs                                                      # 
#===================================================================================#

import sys
import os
import os.path
import time
import random
import math
import shutil 
import subprocess
import platform
import warnings
import csv
import io
import shapefile
import simplekml
import xml.etree.ElementTree as et
import configparser
import numpy as np
import xlrd
import xlwt
import xlsxwriter
import openpyxl
import re
import json
from ezodf import opendoc
from dbfpy3 import dbf
from os.path import basename
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtGui import QStandardItemModel, QStandardItem
from formlayout import fedit
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
from mpl_toolkits.basemap import Basemap
import matplotlib.colors as colors
from matplotlib.colors import cnames, LinearSegmentedColormap, LogNorm
        
warnings.filterwarnings("ignore")

VERSION = "3.0 <i>Vancouver</i>"

symbols = {"o":0,"^":1,"v":2,"<":3,">":4,"s":5,"+":6,"x":7,"D":8,"d":9,"h":10,"p":11,"|":12,"_":13}

MASSING_PALETTES = {
    "Yellow-Orange-Red": ["#ffffcc", "#fd8d3c", "#800026"],
    "White-Green":       ["#ffffff", "#74c476", "#00441b"],
    "White-Blue-Purple": ["#f7fbff", "#6baed6", "#08306b"],
    "Yellow-Green-Blue": ["#ffffcc", "#78c679", "#004529"],
    "White-Red-Black":   ["#ffffff", "#fc9272", "#67000d"],
    "White-Gray-Black":  ["#ffffff", "#969696", "#000000"],
}

def random_color():
    return random.choice(list(cnames.values()))

def get_basename(fname):
    return os.path.splitext(os.path.basename(fname))[0]
  
def getExecPath():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    try:
        sFile = os.path.abspath(sys.modules['__main__'].__file__)
    except Exception:
        sFile = sys.executable
    return os.path.dirname(sFile)    
    
def martitracks_executable():
    system = platform.system()
    if system == "Windows":
        name = "mt05-win64.exe"
    elif system == "Linux":
        name = "mt05-64"
    else:
        QtWidgets.QMessageBox.warning(None, "Unsupported platform",
            "MartiTracks is only supported on Windows and Linux.\n"
            "Detected platform: %s" % system)
        return None
    w = shutil.which(name)
    if w:
        return w
    base = getExecPath()
    path = os.path.join(base, name)
    if os.path.exists(path):
        return path
    alt = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), name)
    if os.path.exists(alt):
        return alt
    return path   

def panbiotracks_executable():
    system = platform.system()
    if system == "Windows":
        name = "panbiotracks_win_x64_0.3.0.exe"
    elif system == "Linux":
        name = "panbiotracks_linux_x64_0.3.0"
    else:
        QtWidgets.QMessageBox.warning(None, "Unsupported platform",
            "Panbiotracks is only supported on Windows and Linux.\n"
            "Detected platform: %s" % system)
        return None
    w = shutil.which(name)
    if w:
        return w
    base = getExecPath()
    path = os.path.join(base, name)
    if os.path.exists(path):
        return path
    return path    

def getIcon(name):
    # Images are stored in the images directory next to the program.
    path = os.path.join(getExecPath(), "images", name)
    if os.path.exists(path):
        return QtGui.QIcon(path)
    return QtGui.QIcon()

def getImagePath(name):
    return os.path.join(getExecPath(), "images", name)
    
def iif(boolVar, ifTrue, ifFalse):
    if boolVar:
        return ifTrue
    else:
        return ifFalse

def to_float(in_val):
    try:
        ret_val = float(in_val)
        return ret_val
    except:
        return 0.0
        
class _GeoJSONLineLayer:
    # Adapter so a list of Line2D objects can be used where Croizat
    # expects the collection returned by Basemap.readshapefile().

    def __init__(self, lines):
        self.lines = lines
        self._visible = True

    def set_visible(self, v):
        self._visible = v
        for line in self.lines:
            line.set_visible(v)

    def get_visible(self):
        return self._visible

    def get_linewidth(self):
        if not self.lines:
            return [2]
        return [self.lines[0].get_linewidth()]

    def set_linewidth(self, w):
        for line in self.lines:
            line.set_linewidth(w)

    def get_color(self):
        if not self.lines:
            return ['#000000']
        return [self.lines[0].get_color()]

    def set_color(self, c):
        for line in self.lines:
            line.set_color(c)

def loadGeoJSONLines(filename):
    # Read a GeoJSON file and return a list of line segments.
    #Each segment is a list of (lon, lat) tuples.  Point and multipoint
    #features are ignored; LineString and MultiLineString features are
    #accepted.  Returns an empty list if the file has no line features.
    if not os.path.exists(filename):
        return []
    with open(filename, "r", encoding="utf-8") as f:
        gj = json.load(f)

    if gj.get("type") == "FeatureCollection":
        features = gj.get("features", [])
    elif gj.get("type") == "Feature":
        features = [gj]
    elif gj.get("type") in ("LineString", "MultiLineString"):
        # Bare geometry, not wrapped in a Feature
        features = [{"type": "Feature", "geometry": gj}]
    else:
        features = []

    segments = []
    for feat in features:
        geom = feat.get("geometry") or {}
        gtype = geom.get("type")
        coords = geom.get("coordinates")
        if not coords:
            continue
        if gtype == "LineString":
            segments.append([(float(pt[0]), float(pt[1])) for pt in coords])
        elif gtype == "MultiLineString":
            for line in coords:
                segments.append([(float(pt[0]), float(pt[1])) for pt in line])
        # Point / MultiPoint / Polygon / etc. are ignored
    return segments
    
def geojsonHasLines(filename):
    # Return True if the GeoJSON file contains any line features.
    if not os.path.exists(filename):
        return False
    try:
        with open(filename, "r", encoding="utf-8") as f:
            gj = json.load(f)
    except (OSError, ValueError):
        return False

    if gj.get("type") == "FeatureCollection":
        features = gj.get("features", [])
    elif gj.get("type") == "Feature":
        features = [gj]
    else:
        features = []

    for feat in features:
        geom = feat.get("geometry") or {}
        if geom.get("type") in ("LineString", "MultiLineString"):
            return True
    return False

def distance_matrix(pts):
    # Returns matrix of pairwise Euclidean distances, vectorized numpy version
    return np.sum((pts[None,:] - pts[:, None])**2, -1)**0.5

def minimum_spanning_tree(X, copy_X=True):
    # Compute the minimum spanning tree of a fully connected graph.

    # Implements Prim's algorithm (Prim, 1957): start with vertex 0, and at
    # each step add the minimum-weight edge connecting the current tree to
    # a vertex not yet in it. X is the edge-weight matrix (X[i,j] is the
    # weight of edge i-j). Returns an array of edges, one per row, as
    # pairs of vertex indices.

    # Note: this is the matrix-based O(n^3) form of Prim's, adequate for
    #cthe sizes Croizat handles (tens to low hundreds of points per track).
    # X are edge weights of fully connected graph
    if copy_X:
        X = X.copy()

    if X.shape[0] != X.shape[1]:
        raise ValueError("X needs to be square matrix of edge weights")
    n_vertices = X.shape[0]
    spanning_edges = []

    # initialize with node 0:
    visited_vertices = [0]
    num_visited = 1
    # exclude self connections:
    diag_indices = np.arange(n_vertices)
    X[diag_indices, diag_indices] = np.inf

    while num_visited != n_vertices:
        new_edge = np.argmin(X[visited_vertices], axis=None)
        # 2d encoding of new_edge from flat, get correct indices
        new_edge = divmod(new_edge, n_vertices)
        new_edge = [visited_vertices[new_edge[0]], new_edge[1]]
        # add edge to tree
        spanning_edges.append(new_edge)
        visited_vertices.append(new_edge[1])
        # remove all edges inside current tree
        X[visited_vertices, new_edge[1]] = np.inf
        X[new_edge[1], visited_vertices] = np.inf
        num_visited += 1
    return np.vstack(spanning_edges)

class ChecklistDialog(QtWidgets.QDialog):
    def __init__(self, name, stringlist=None, checked=False, icon=None, parent=None):
        super(ChecklistDialog, self).__init__(parent)

        self.name = name
        self.icon = icon
        self.model = QStandardItemModel()
        self.listView = QtWidgets.QListView()

        if stringlist is not None:
            for string in stringlist:
                item = QStandardItem(string)
                item.setCheckable(True)
                check = QtCore.Qt.CheckState.Checked if checked else QtCore.Qt.CheckState.Unchecked
                item.setCheckState(check)
                self.model.appendRow(item)

        self.listView.setModel(self.model)

        self.okButton = QtWidgets.QPushButton("OK")
        self.cancelButton = QtWidgets.QPushButton("Cancel")
        self.selectButton = QtWidgets.QPushButton("Select All")
        self.unselectButton = QtWidgets.QPushButton("Unselect All")

        hbox = QtWidgets.QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(self.okButton)
        hbox.addWidget(self.cancelButton)
        hbox.addWidget(self.selectButton)
        hbox.addWidget(self.unselectButton)

        vbox = QtWidgets.QVBoxLayout()
        vbox.addWidget(self.listView)
        vbox.addStretch(1)
        vbox.addLayout(hbox)

        self.setLayout(vbox)
        self.setWindowTitle(self.name)
        if self.icon is not None: self.setWindowIcon(self.icon)

        self.okButton.clicked.connect(self.accept)
        self.cancelButton.clicked.connect(self.reject)
        self.selectButton.clicked.connect(self.select)
        self.unselectButton.clicked.connect(self.unselect)

        self.listView.setFocus()

    def reject(self):
        QtWidgets.QDialog.reject(self)

    def accept(self):
        self.choices = []
        i = 0
        while self.model.item(i):
            if self.model.item(i).checkState() == QtCore.Qt.CheckState.Checked:
                self.choices.append(self.model.item(i).text())
            i += 1
        QtWidgets.QDialog.accept(self)

    def select(self):
        for i in range(self.model.rowCount()):
            item = self.model.item(i)
            item.setCheckState(QtCore.Qt.CheckState.Checked)

    def unselect(self):
        for i in range(self.model.rowCount()):
            item = self.model.item(i)
            item.setCheckState(QtCore.Qt.CheckState.Unchecked)

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_DeleteOnClose)
        self.initUI()

    def initUI(self):
        self.filename = ""
        self.filelist = {}
        self.init_parameters()

        self.fileMenu = QtWidgets.QMenu("&Project", self)
        self.newAction = QtGui.QAction(getIcon("new.png"), "&New...", self,
                        shortcut="Ctrl+N", statusTip="Create a new project",
                        triggered=self.newFile)
        self.fileMenu.addAction(self.newAction)

        self.openAction = QtGui.QAction(getIcon("open.png"), "&Open...", self,
                        shortcut="Ctrl+O", statusTip="Open an existing project",
                        triggered=self.openFile)
        self.fileMenu.addAction(self.openAction)

        self.saveAction = QtGui.QAction(getIcon("save.png"), "Save &As...", self,
                        shortcut="Ctrl+S", statusTip="Save the current project",
                        triggered=self.saveFile)
        self.fileMenu.addAction(self.saveAction)

        self.fileMenu.addSeparator()
        self.addAction = QtGui.QAction(getIcon("add.png"), "&Add File...", self,
                        shortcut="Ctrl+Ins", statusTip="Add file to project",
                        triggered=self.addFile)
        self.fileMenu.addAction(self.addAction)

        self.delAction = QtGui.QAction(getIcon("remove.png"), "&Remove File...", self,
                        shortcut="Ctrl+Del", statusTip="Remove file from project",
                        triggered=self.delFile)
        self.fileMenu.addAction(self.delAction)

        self.fileMenu.addSeparator()
        self.exportAction = QtGui.QAction(getIcon("export.png"), "Export Data...", self,
                        statusTip="Export data",
                        triggered=self.exportFile)
        self.fileMenu.addAction(self.exportAction)
        
        self.fileMenu.addSeparator()
        self.paeAction = QtGui.QAction(getIcon("tree.png"), "PAE Mmatrix...", self,
                        statusTip="Build a presence/absence matrix for Parsimony Analysis of Endemicity",
                        triggered=self.paeMatrix)
        self.fileMenu.addAction(self.paeAction)
        
        self.outlierAction = QtGui.QAction(getIcon("outlier.png"), "&Detect Outliers...", self,
                statusTip="Identify potential outlier points in the loaded data",
                triggered=self.detectOutliers)
        self.fileMenu.addAction(self.outlierAction)

        self.fileMenu.addSeparator()
        self.closeAction = QtGui.QAction(getIcon("exit.png"), "&Quit", self,
                        shortcut="Ctrl+Q", statusTip="Quit the program",
                        triggered=self.close)
        self.fileMenu.addAction(self.closeAction)
        self.menuBar().addMenu(self.fileMenu)

        self.mapMenu = QtWidgets.QMenu("&Map", self)
        self.borderAction = QtGui.QAction("Borders", self, checkable=True,
                        statusTip="Draw country borders",
                        triggered=self.borders)
        self.mapMenu.addAction(self.borderAction)

        self.riverAction = QtGui.QAction("Rivers", self, checkable=True,
                        statusTip="Draw rivers",
                        triggered=self.rivers)
        self.mapMenu.addAction(self.riverAction)

        self.mapMenu.addSeparator()
        self.limitsAction = QtGui.QAction("Limits...", self,
                        statusTip="Set map limits",
                        triggered=self.maplimits)
        self.mapMenu.addAction(self.limitsAction)

        self.centerAction = QtGui.QAction("Center...", self,
                        statusTip="Set map center",
                        triggered=self.mapcenter)
        self.mapMenu.addAction(self.centerAction)

        self.mapMenu.addSeparator()
        self.resolutionAction = QtGui.QAction("Resolution...", self,
                        statusTip="Set map resolution",
                        triggered=self.mapresolution)
        self.mapMenu.addAction(self.resolutionAction)

        self.imageMenu = self.mapMenu.addMenu("Background")
        group = QtGui.QActionGroup(self.mapMenu)
        group.setExclusive(True)

        self.defaultmapAction = QtGui.QAction("Default", self, checkable=True,
                        statusTip="Display default map",
                        triggered=self.defaultmap)
        action = group.addAction(self.defaultmapAction)
        self.imageMenu.addAction(action)
        self.defaultmapAction.setChecked(self.basemap)

        self.bluemarbleAction = QtGui.QAction("Blue Marble...", self, checkable=True,
                        statusTip="Display NASA Blue Marble image",
                        triggered=self.bluemarble)
        action = group.addAction(self.bluemarbleAction)
        self.imageMenu.addAction(action)

        self.etopomodelAction = QtGui.QAction("ETOPO Global Relief...", self, checkable=True,
                        statusTip="Display NOAA ETOPO Global Relief Model image",
                        triggered=self.etopo)
        action = group.addAction(self.etopomodelAction)
        self.imageMenu.addAction(action)

        self.shadedreliefAction = QtGui.QAction("Shaded Relief...", self, checkable=True,
                        statusTip="Display shaded relief image",
                        triggered=self.shaded)
        action = group.addAction(self.shadedreliefAction)
        self.imageMenu.addAction(action)

        self.mapMenu.addSeparator()
        self.colorAction = QtGui.QAction("Colors...", self,
                        statusTip="Set map colors",
                        triggered=self.mapcolors)
        self.mapMenu.addAction(self.colorAction)

        self.titleAction = QtGui.QAction("Title...", self,
                        statusTip="Set map title",
                        triggered=self.maptitle)
        self.mapMenu.addAction(self.titleAction)

        self.mapMenu.addSeparator()
        self.copyAction = QtGui.QAction(getIcon("copy.png"), "Copy", self,
                        shortcut="Ctrl+C", statusTip="Copy map to the clipboard",
                        triggered=self.copytoclipboard)
        self.mapMenu.addAction(self.copyAction)

        self.reloadAction = QtGui.QAction(getIcon("reload.png"), "Reload", self,
                        shortcut="Ctrl+R", statusTip="Reload data and redraw map",
                        triggered=self.reload)
        self.mapMenu.addAction(self.reloadAction)
        self.menuBar().addMenu(self.mapMenu)

        self.runMenu = QtWidgets.QMenu("&Analysis", self)
        self.trackAction = QtGui.QAction(getIcon("track.png"), "&Individual Tracks...", self,
                        statusTip="Compute individual tracks",
                        triggered=self.track)
        self.runMenu.addAction(self.trackAction)

        self.gentrackMenu = self.runMenu.addMenu(getIcon("clique.png"),
                        "&Generalized Tracks")

        self.gentracksAction = QtGui.QAction(getIcon("clique.png"), "Marti&Tracks...", self,
                        statusTip="Compute generalized tracks with MartiTracks",
                        triggered=self.gentracks)
        self.gentrackMenu.addAction(self.gentracksAction)

        self.panbiotracksAction = QtGui.QAction(getIcon("clique.png"), "&Panbiotracks...", self,
                        statusTip="Compute generalized tracks with Panbiotracks",
                        triggered=self.panbiotracks)
        self.gentrackMenu.addAction(self.panbiotracksAction)

        self.nodesAction = QtGui.QAction(getIcon("node.png"), "&Nodes...", self,
                        statusTip="Compute panbiogeographic nodes with Panbiotracks",
                        triggered=self.nodes)
        self.runMenu.addAction(self.nodesAction)
        
        self.massingAction = QtGui.QAction(getIcon("grid.png"), "&Main Massings...", self,
                statusTip="Compute main massings",
                triggered=self.massings)
        self.runMenu.addAction(self.massingAction)

        self.menuBar().addMenu(self.runMenu)

        self.helpAction = QtGui.QAction(getIcon("about.png"), "&About", self,
                        statusTip="Display information about the program",
                        triggered=self.about)

        self.helpMenu = QtWidgets.QMenu("&Help", self)
        self.helpMenu.addAction(self.helpAction)
        self.menuBar().addMenu(self.helpMenu)

        self.statusBar().setSizeGripEnabled(True)
        self.statusBar().showMessage("Ready")

        toolbar = self.addToolBar("Toolbar")
        toolbar.setObjectName("Toolbar")
        toolbar.addAction(self.newAction)
        toolbar.addAction(self.openAction)
        toolbar.addSeparator()
        toolbar.addAction(self.addAction)
        toolbar.addAction(self.delAction)
        toolbar.addSeparator()
        toolbar.addAction(self.reloadAction)
        toolbar.addSeparator()
        toolbar.addAction(self.trackAction)
        toolbar.addAction(self.gentracksAction)
        toolbar.addAction(self.panbiotracksAction)
        toolbar.addAction(self.nodesAction)
        toolbar.addAction(self.massingAction)
        toolbar.addSeparator()
        toolbar.addAction(self.closeAction)

        self.items = QtWidgets.QDockWidget("Layers", self)
        self.items.setFloating(False)
        self.items.setFeatures(QtWidgets.QDockWidget.DockWidgetFeature.NoDockWidgetFeatures)

        self.listWidget = QtWidgets.QListWidget()
        self.listWidget.setSortingEnabled(True)
        self.listWidget.itemChanged.connect(self.listWidgetItemChanged)
        self.listWidget.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.listWidget.customContextMenuRequested.connect(self.itemContextMenu)
        addFileKey = QtGui.QShortcut(QtGui.QKeySequence("Alt+Ins"), self.listWidget)
        delFileKey = QtGui.QShortcut(QtGui.QKeySequence("Alt+Del"), self.listWidget)
        editPropertiesKey = QtGui.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key.Key_Return), self.listWidget)
        addFileKey.activated.connect(self.addFile)
        delFileKey.activated.connect(self.delFile)
        editPropertiesKey.activated.connect(self.editProperties)
        self.items.setWidget(self.listWidget)

        self.plotWidget = QtWidgets.QWidget(self)
        layout = QtWidgets.QVBoxLayout(self.plotWidget)
        self.figure = Figure(dpi=self.dpi)
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)
        self.canvas.mpl_connect("pick_event", self.pick)
        self.compute_initial_figure()
        layout.addWidget(self.canvas)
        layout.addWidget(self.toolbar)
        self.setCentralWidget(self.plotWidget)

        self.setGeometry(100,100,650,350)
        self.addDockWidget(QtCore.Qt.DockWidgetArea.LeftDockWidgetArea, self.items)
        self.setWindowTitle("Croizat")
        self.setWindowIcon(getIcon("croizat.ico"))
        self.updateUI()
        self.showMaximized()

    def itemContextMenu(self, pos):
        menu = QtWidgets.QMenu(self)
        insAction = QtGui.QAction("Add file...", self,
                        statusTip="Add file to project",
                        triggered=self.addFile)
        menu.addAction(insAction)
        enable = len(self.filename) > 0
        insAction.setEnabled(enable)

        delAction = QtGui.QAction("Remove file...", self,
                        statusTip="Remove file from project",
                        triggered=self.delFile)
        menu.addAction(delAction)
        delAction.setEnabled(self.listWidget.count() > 0)
        menu.addSeparator()

        editAction = QtGui.QAction("Properties...", self,
                        statusTip="Modify file properties",
                        triggered=self.editProperties)
        menu.addAction(editAction)
        editAction.setEnabled(self.listWidget.count() > 0)
        menu.addSeparator()

        recordsAction = QtGui.QAction("View/Hide records", self,
                        statusTip="Display or hide all records",
                        triggered=self.displayRecords)
        menu.addAction(recordsAction)
        recordsAction.setEnabled(self.listWidget.count() > 0)

        treesAction = QtGui.QAction("View/Hide individual tracks", self,
                        statusTip="Display or hide individual tracks",
                        triggered=self.displayTrees)
        menu.addAction(treesAction)
        treesAction.setEnabled(self.listWidget.count() > 0)

        tracksAction = QtGui.QAction("View/Hide generalized tracks", self,
                        statusTip="Display or hide generalized tracks",
                        triggered=self.displayTracks)
        menu.addAction(tracksAction)
        tracksAction.setEnabled(self.listWidget.count() > 0)

        action = menu.exec(self.listWidget.mapToGlobal(pos))

    def displayRecords(self):
        if self.listWidget.count() > 0:
            for i in range(self.listWidget.count()):
                item = self.listWidget.item(i)
                filename = str(item.text())
                if filename.find("_individualtrack") != -1 or filename.find("_ MST") != -1:
                    pass
                else:
                    checked = item.checkState() == QtCore.Qt.CheckState.Checked
                    item.setCheckState(iif(checked, QtCore.Qt.CheckState.Unchecked, QtCore.Qt.CheckState.Checked))
            self.canvas.draw_idle()

    def displayTrees(self):
        if self.listWidget.count() > 0:
            for i in range(self.listWidget.count()):
                item = self.listWidget.item(i)
                filename = str(item.text())
                if filename.find("_individualtrack") != -1:
                    checked = item.checkState() == QtCore.Qt.CheckState.Checked
                    item.setCheckState(iif(checked, QtCore.Qt.CheckState.Unchecked, QtCore.Qt.CheckState.Checked))
            self.canvas.draw_idle()

    def displayTracks(self):
        if self.listWidget.count() > 0:
            for i in range(self.listWidget.count()):
                item = self.listWidget.item(i)
                filename = str(item.text())
                if filename.find("_ MST") != -1:
                    checked = item.checkState() == QtCore.Qt.CheckState.Checked
                    item.setCheckState(iif(checked, QtCore.Qt.CheckState.Unchecked, QtCore.Qt.CheckState.Checked))
            self.canvas.draw_idle()

    def compute_initial_figure(self):
        self._massing_colorbar = None
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
        self.figure.clear()
        axes = self.figure.add_subplot(111)
        self.figure.set_tight_layout(True)

        # plot map
        self.map = Basemap(ax=axes, llcrnrlon=self.west, llcrnrlat=self.south, urcrnrlon=self.east, urcrnrlat=self.north, \
            projection=self.projection, resolution=self.resolution, lat_0=self.clat, lon_0=self.clon)
        if self.visibleearth:
            self.map.drawcoastlines(color='#777799')
            self.map.drawmapboundary(fill_color='none')
            self.map.bluemarble(scale=self.scale)
        elif self.etopomodel:
            self.map.drawcoastlines(color='#777799')
            self.map.drawmapboundary(fill_color='none')
            self.map.etopo(scale=self.scale)
        elif self.shadedrelief:
            self.map.drawcoastlines(color='#777799')
            self.map.drawmapboundary(fill_color='none')
            self.map.shadedrelief(scale=self.scale)
        else:
            self.map.drawmapboundary(fill_color=self.watercolor)
            self.map.fillcontinents(color=self.landcolor,lake_color=self.watercolor)
            self.map.drawcoastlines()

        if self.borderAction.isChecked():
            color = iif((self.visibleearth or self.etopomodel or self.shadedrelief), '#ccccee', self.bordercolor)
            self.borders = self.map.drawcountries(color=color)

        if self.riverAction.isChecked():
            color = iif((self.visibleearth or self.etopomodel or self.shadedrelief), '#ccccee', self.bordercolor)
            self.rivers = self.map.drawrivers(color=color)

        if self.listWidget.count() > 0:
            for i in range(self.listWidget.count()):
                item = self.listWidget.item(i)
                filename = str(item.text())
                checked = item.checkState() == QtCore.Qt.CheckState.Checked
                if self.is_line_layer(filename):
                    if filename not in self.shape_list:
                        continue
                    shape = self.shape_list[filename]
                    if not isinstance(shape, tuple) or len(shape) < 5:
                        continue
                    file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
                    if file_ext == "shp":
                        width = shape[4].get_linewidth()[0]
                        color = colors.to_hex(shape[4].get_color()[0])
                        self.shape_list[filename] = self.map.readshapefile(
                            os.path.splitext(filename)[0], filename,
                            linewidth=float(width), color=color)
                    else:
                        # line-GeoJSON: re-parse and re-plot
                        old = shape[4]
                        width = old.get_linewidth()[0]
                        color = old.get_color()[0]
                        segments = loadGeoJSONLines(filename)
                        plotted_lines = []
                        for seg in segments:
                            lons = [p[0] for p in seg]
                            lats = [p[1] for p in seg]
                            line, = self.map.plot(lons, lats, '-',
                                                  color=color, linewidth=float(width))
                            plotted_lines.append(line)
                        self.shape_list[filename] = (None, None, None, None,
                                                     _GeoJSONLineLayer(plotted_lines))
                    if checked:
                        shape = self.shape_list[filename]
                        if isinstance(shape, tuple) and len(shape) >= 5:
                            shape[4].set_visible(checked)
                else:
                    if filename not in self.plot_list:
                        continue
                    plot = self.plot_list[filename]
                    marker = plot[0].get_marker()
                    size = plot[0].get_markersize()
                    color = plot[0].get_markerfacecolor()
                    headers, data = self.loadData(filename)
                    coords = []
                    for row in data:
                        longitude = row[2]
                        latitude = row[3]
                        coords.append((longitude, latitude))
                    coords = np.array(coords, dtype=np.float32)
                    self.plot_list[filename] = self.map.ax.plot(coords[:,0], coords[:,1], marker, markersize=int(size), color=color, picker=5)
                    if checked:
                        plot = self.plot_list[filename]
                        plot[0].set_visible(checked)

        self.map.ax.callbacks.connect("xlim_changed", self.change_xlimits)
        self.map.ax.callbacks.connect("ylim_changed", self.change_ylimits)

        self.map.ax.set_title(self.title)
        self.canvas.draw()
        QtWidgets.QApplication.restoreOverrideCursor()

    def init_parameters(self):
        self.plot_list = {}
        self.edge_list = []
        self.graph_list = {}
        self.shape_list = {}
        self.title = "Untitled"
        self.landcolor = "#00ff00" #lime
        self.watercolor = "#00ffff" #cyan
        self.bordercolor = "#000000" #black
        self.massingpalette = "Yellow-Orange-Red"
        self.north = 90  ##urcrnrlat=north
        self.south = -90 ##llcrnrlat=south
        self.east = 180  ##urcrnrlon=east
        self.west = -180 ##llcrnrlon=west
        self.clat = 0
        self.clon = 0
        self.projection = "cyl"
        self.resolution = 'c'
        self.basemap = True
        self.visibleearth = False
        self.etopomodel = False
        self.shadedrelief = False
        self.scale = 0.25
        self.dpi = 100

    def updateUI(self):
        enable = len(self.filename) > 0
        self.addAction.setEnabled(enable)
        self.delAction.setEnabled(enable)
        self.exportAction.setEnabled(enable)
        self.paeAction.setEnabled(enable)
        self.outlierAction.setEnabled(enable)
        self.reloadAction.setEnabled(enable)
        self.trackAction.setEnabled(enable)
        self.gentrackMenu.setEnabled(enable)
        self.gentracksAction.setEnabled(enable)
        self.panbiotracksAction.setEnabled(enable)
        self.massingAction.setEnabled(enable)
        self.nodesAction.setEnabled(enable)
        #has_gentracks = any(
        #    "_ mst" in f.lower() or "generalized" in f.lower()
        #    for f in self.filelist)
        #self.nodesAction.setEnabled(enable and has_gentracks)

    def loadData(self, filename):
        file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
        data = []
        headers = []
        if not os.path.exists(filename):
            QtWidgets.QMessageBox.critical(self, 
                "Error", "File " + filename + " not found!")
            return headers, data

        if file_ext == "xls":
            wb = xlrd.open_workbook(filename)
            sh1 = wb.sheet_by_index(0)
            for rownum in range(sh1.nrows):
                data += [sh1.row_values(rownum)]
            headers = data[0]
            del data[0]
        elif file_ext == "xlsx":
            wb = openpyxl.load_workbook(filename, data_only=True)
            sh1 = wb.active
            for row in sh1.iter_rows(values_only=True):
                data.append(list(row))
            headers = data[0]
            del data[0]
        elif file_ext == "ods":
            wb = opendoc(filename)
            s = wb.sheets[0]
            for row in range(s.nrows()):
                values = []
                for col in range(s.ncols()):
                    values += [s[(row,col)].value]
                data += [values]
            headers = data[0]
            del data[0]
        elif file_ext == "csv":
            with open(filename, "r", newline="", encoding="utf-8-sig") as csvfile:
                sample = csvfile.readline()
                csvfile.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, [',',';','\t','|'])
                except csv.Error:
                    dialect = csv.excel
                reader = csv.reader(csvfile, dialect)
                try:
                    headers = next(reader)
                except StopIteration:
                    headers = []
                for row in reader:
                    if row:
                        data.append(row)
        elif file_ext == "dbf":
            db = dbf.Dbf(filename, readOnly=True)
            data = [row.asList() for row in db]
            headers = db.fieldNames
        elif file_ext == "kml":
            doc = et.parse(filename)
            nmsp = '{http://www.opengis.net/kml/2.2}'
            headers = ["scientific name", "locality", "longitude", "latitude"]
            for pm in doc.iterfind('.//{0}Placemark'.format(nmsp)):
                name = pm.find('{0}name'.format(nmsp)).text
                description = pm.find('{0}description'.format(nmsp)).text
                if description is None:
                    description = ""
                description = description.replace('\n', ' ').strip()
                longitude = latitude = None
                for ls in pm.iterfind('{0}Point/{0}coordinates'.format(nmsp)):
                    parts = ls.text.strip().split(',')
                    longitude = float(parts[0])
                    latitude = float(parts[1])
                if longitude is not None and latitude is not None:
                    data.append((name, description, longitude, latitude))
        elif file_ext in ("json", "geojson", "gjs"):
            headers = ["scientific name", "locality", "longitude", "latitude"]
            with open(filename, "r", encoding="utf-8") as gjf:
                gj = json.load(gjf)
            if gj.get("type") == "FeatureCollection":
                features = gj.get("features", [])
            elif gj.get("type") == "Feature":
                features = [gj]
            else:
                features = []
            for feat in features:
                geom = feat.get("geometry") or {}
                if geom.get("type") != "Point":
                    continue
                coords = geom.get("coordinates")
                if not coords or len(coords) < 2:
                    continue
                longitude = float(coords[0])
                latitude = float(coords[1])
                props = feat.get("properties") or {}
                # Case-insensitive property lookup
                plower = {k.lower(): v for k, v in props.items()}
                name = (plower.get("scientificname")
                        or plower.get("species")
                        or plower.get("name")
                        or plower.get("scientific_name")
                        or "")
                locality = (plower.get("locality")
                            or plower.get("description")
                            or "")
                data.append((str(name), str(locality), longitude, latitude))

        # --- Column remapping ---------------------------------------------
        # If the header row contains DarwinCore-style column names (or any
        # of the common aliases), reorder each row so that the first four
        # columns are always [species, locality, longitude, latitude].
        # If no recognizable header is found, leave data as-is (the old
        # positional behaviour).
        if headers and len(headers) >= 2:
            lower = [str(h).lower().strip() for h in headers]

            def find_index(*names):
                for n in names:
                    n_norm = n.replace(" ", "").replace("_", "")
                    for i, h in enumerate(lower):
                        if h.replace(" ", "").replace("_", "") == n_norm:
                            return i
                return None

            sci_idx = find_index("scientificname", "scientific_name",
                                 "species", "speciesname", "taxon", "name")
            loc_idx = find_index("locality", "verbatimlocality",
                                 "verbatim_locality", "location", "description")
            lon_idx = find_index("decimallongitude", "decimal_longitude",
                                 "longitude", "lon", "x")
            lat_idx = find_index("decimallatitude", "decimal_latitude",
                                 "latitude", "lat", "y")

            # Only remap if we found at least longitude and latitude.
            # Species and locality are optional (empty string if absent).
            if lon_idx is not None and lat_idx is not None:
                new_data = []
                for row in data:
                    try:
                        species = str(row[sci_idx]) if sci_idx is not None else ""
                    except (IndexError, TypeError):
                        species = ""
                    try:
                        locality = str(row[loc_idx]) if loc_idx is not None else ""
                    except (IndexError, TypeError):
                        locality = ""
                    try:
                        longitude = float(row[lon_idx])
                    except (IndexError, ValueError, TypeError):
                        continue
                    try:
                        latitude = float(row[lat_idx])
                    except (IndexError, ValueError, TypeError):
                        continue
                    new_data.append((species, locality, longitude, latitude))
                data = new_data
                # After remapping, the effective headers are the four we
                # now produce, so downstream code sees the expected layout.
                headers = ["Scientific name", "Locality", "Longitude", "Latitude"]
            else:
                # No recognized coordinate columns; fall back to positional
                # interpretation (assumes columns 0..3 are species, locality,
                # longitude, latitude, which is Croizat's native format).
                new_data = []
                for row in data:
                    if len(row) < 4:
                        continue
                    try:
                        longitude = float(row[2])
                        latitude = float(row[3])
                    except (ValueError, TypeError):
                        continue
                    new_data.append((str(row[0]), str(row[1]), longitude, latitude))
                data = new_data
                headers = ["Scientific name", "Locality", "Longitude", "Latitude"]

        # Normalise header titles for display
        for i in range(len(headers)):
            headers[i] = str(headers[i]).title().strip('\r')
        return headers, data

    def importData(self, filename, infmt, outfmt):
        infile = os.path.basename(str(filename))
        outfile = os.path.splitext(infile)[0] + "_imported_from_" + infmt

        if infmt == "gbif":
            delim = '\t'
            scientificname = 12
            locality = 14
            decimallatitude = 16
            decimallongitude = 17
        elif infmt == "idigbio":
            delim = ','
            scientificname = 145
            locality = 101
            decimallatitude = 43
            decimallongitude = 44
        elif infmt == "obis":
            delim = ','
            scientificname = 4
            locality = 7
            decimallatitude = 9
            decimallongitude = 10
        elif infmt == "vertnet":
            delim = '\t'
            scientificname = 129
            locality = 71
            verbatimlocality = 72
            decimallatitude = 83
            decimallongitude = 84

        headers = ["Scientific name", "Locality", "Longitude", "Latitude"]
        if outfmt == "csv":
            outfile += ".csv"
            writer = csv.writer(open(outfile, "w", newline="", encoding="utf-8"), delimiter=',', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
            writer.writerow(headers)
        elif outfmt == "xls":
            outfile += ".xls"
            book = xlwt.Workbook()
            sheet1 = book.add_sheet("Sheet1")
            row1 = sheet1.row(0)
            for i in range(len(headers)):
                row1.write(i, headers[i])
        elif outfmt == "xlsx":
            outfile += ".xlsx"
            book = xlsxwriter.Workbook(outfile)
            sheet1 = book.add_worksheet("Sheet1")
            for i in range(1, len(headers) + 1):
                sheet1.write(0, i - 1, headers[i - 1])
        elif outfmt == "ods":
            outfile += ".ods"
            book = newdoc(doctype="ods", filename=outfile)
            book.sheets.append(Table("Sheet1"))
            sheet1 = book.sheets[0]
            sheet1.append_columns(len(headers))
            for i in range(len(headers)):
                sheet1[0, i].set_value(headers[i])

        with open(infile, "r", newline="", encoding="utf-8-sig") as infile:
            reader = csv.reader(infile, delimiter=delim)
            next(reader)
            indata = list(reader)

        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
        recs = 0
        recnum = 0
        for row in indata:
            if infmt == "vertnet":
                Locality = iif(len(row[locality]) == 0,row[verbatimlocality], row[locality])
            else:
                Locality = row[locality]
            if len(Locality) == 0:
                Locality = "N/A"
            else:
                Locality = str(Locality)
            Latitude = to_float(row[decimallatitude])
            Longitude = to_float(row[decimallongitude])
            if Latitude == 0.0 or Longitude == 0.0: continue
            ScientificName = row[scientificname]

            if outfmt == "csv":
                data = (ScientificName, Locality, Longitude, Latitude)
                writer.writerow(data)
            elif outfmt == "xls":
                sheet1.row(recnum + 1).write(0, ScientificName)
                sheet1.row(recnum + 1).write(1, Locality)
                sheet1.row(recnum + 1).write(2, Longitude)
                sheet1.row(recnum + 1).write(3, Latitude)
            elif outfmt == "xlsx":
                sheet1.write(recnum + 1, 0, ScientificName)
                sheet1.write(recnum + 1, 1, Locality)
                sheet1.write(recnum + 1, 2, Longitude)
                sheet1.write(recnum + 1, 3, Latitude)
            elif outfmt == "ods":
                sheet1.append_rows(1)
                sheet1[recnum + 1, 0].set_value(ScientificName)
                sheet1[recnum + 1, 1].set_value(Locality)
                sheet1[recnum + 1, 2].set_value(Longitude)
                sheet1[recnum + 1, 3].set_value(Latitude)
            recs += 1
            recnum = recs

        if outfmt == "xls":
            book.save(outfile)
        elif outfmt == "xlsx":
            book.close()
        elif outfmt == "ods":
            book.save()
        QtWidgets.QApplication.restoreOverrideCursor()

        return outfile, recs

    def readProject(self):
        self.listWidget.clear()
        self.filelist.clear()
        config = configparser.RawConfigParser()
        config.read(str(self.filename))
        self.title = config.get("Title", "title")
        self.landcolor = config.get("Colors", "land")
        self.watercolor = config.get("Colors", "water")
        self.bordercolor = config.get("Colors", "border")
        self.massing_palette = config.get("Colors", "massing_palette",
                                           fallback=self.massingpalette)
        self.borderAction.setChecked(config.getboolean("Options", "borders"))
        self.riverAction.setChecked(config.getboolean("Options", "rivers"))
        self.projection = config.get("Options", "projection")
        self.resolution = config.get("Options", "resolution")
        self.north = config.getfloat("View", "north")
        self.south = config.getfloat("View", "south")
        self.east = config.getfloat("View", "east")
        self.west = config.getfloat("View", "west")
        self.clat = config.getfloat("View", "clat")
        self.clon = config.getfloat("View", "clon")
        self.basemap = config.getboolean("Image", "basemap")
        self.defaultmapAction.setChecked(self.basemap)
        self.visibleearth = config.getboolean("Image", "bluemarble")
        self.bluemarbleAction.setChecked(self.visibleearth)
        self.etopomodel = config.getboolean("Image", "etopomodel")
        self.etopomodelAction.setChecked(self.etopomodel)
        self.shadedrelief = config.getboolean("Image", "shadedrelief")
        self.shadedreliefAction.setChecked(self.shadedrelief)
        self.scale = config.getfloat("Image", "scale")
        self.dpi = int(config.get("Image", "dpi"))
        for name, value in config.items("Files"):
            for line in value.splitlines():
                if line:
                    parts = line.split(',')
                    key, attributes = parts[0], parts[1:]
                    self.filelist[key] = attributes

    def saveProject(self):
        if len(str(self.filename)) > 0 and len(self.filelist) > 0:
            config = configparser.RawConfigParser(allow_no_value=True)
            config.add_section("Title")
            config.set("Title", "title", self.title)
            config.add_section("Colors")
            config.set("Colors", "land", self.landcolor)
            config.set("Colors", "water", self.watercolor)
            config.set("Colors", "border", self.bordercolor)
            config.set("Colors", "massingpalette", self.massing_palette)
            config.add_section("Options")
            config.set("Options", "borders", str(self.borderAction.isChecked()))
            config.set("Options", "rivers", str(self.riverAction.isChecked()))
            config.set("Options", "projection", self.projection)
            config.set("Options", "resolution", self.resolution)
            config.add_section("View")
            config.set("View", "north", str(self.north))
            config.set("View", "south", str(self.south))
            config.set("View", "east", str(self.east))
            config.set("View", "west", str(self.west))
            config.set("View", "clat", str(self.clat))
            config.set("View", "clon", str(self.clon))
            config.add_section("Image")
            config.set("Image", "basemap", str(self.basemap))
            config.set("Image", "bluemarble", str(self.visibleearth))
            config.set("Image", "etopomodel", str(self.etopomodel))
            config.set("Image", "shadedrelief", str(self.shadedrelief))
            config.set("Image", "scale", str(self.scale))
            config.set("Image", "dpi", str(self.dpi))
            config.add_section("Files")
            i = 0
            for k in range(self.listWidget.count()):
                item = self.listWidget.item(k)
                filename = str(item.text())
                if os.path.exists(filename):
                    if self.is_line_layer(filename) and filename in self.shape_list:
                        shape = self.shape_list[filename]
                        if not isinstance(shape, tuple) or len(shape) < 5:
                            continue                        
                        style = '-'
                        width = shape[4].get_linewidth()[0]
                        color = colors.to_hex(shape[4].get_color()[0])
                        state = iif(item.checkState() == QtCore.Qt.CheckState.Checked, 1, 0)
                        outstr = ',' + style + ',' + str(int(width)) + ',' + color.lower() + ',' + str(state)
                    else:
                        plot = self.plot_list[filename]
                        marker = plot[0].get_marker()
                        size = plot[0].get_markersize()
                        color = plot[0].get_markerfacecolor()
                        state = iif(item.checkState() == QtCore.Qt.CheckState.Checked, 1, 0)
                        outstr = ',' + marker + ',' + str(int(size)) + ',' + color.lower() + ',' + str(state)
                    config.set("Files", "file" + str(i), filename + outstr)
                    i += 1
            configfile = open(str(self.filename), "w", encoding="utf-8")
            config.write(configfile)

    def newFile(self):
        self.saveProject()
        self.filename, _ = QtWidgets.QFileDialog.getSaveFileName(self,
                    "Save As",
                    self.filename,
                    "Project (*.prj)")
        if self.filename:
            directory = os.path.dirname(str(self.filename))
            os.chdir(directory)
            open(str(self.filename), 'a').close()
            self.filelist.clear()
            self.listWidget.clear()
            self.setWindowTitle("Croizat - " + os.path.basename(str(self.filename)))
            self.updateUI()
            self.init_parameters()
            self.compute_initial_figure()
           
    def openFile(self):
        self.saveProject()
        self.filename, _ = QtWidgets.QFileDialog.getOpenFileName(self,
                "Open File",
                "",
                "Project (*.prj)")
        if self.filename:
            directory = os.path.dirname(str(self.filename))
            os.chdir(directory)
            self.readProject()
            self.compute_initial_figure()
            if len(self.filelist) == 0:
                QtWidgets.QMessageBox.warning(self, "Warning", "Project has no data files!")
                return
            QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
            filecount = 0
            for filename in self.filelist:
                attributes = self.filelist[filename]
                symbol = attributes[0]
                size = attributes[1]
                color = attributes[2]
                state = attributes[3].replace('\n','')
                checked = int(state) == 1
                if os.path.exists(filename):
                    item = QtWidgets.QListWidgetItem(filename)
                    item.setCheckState(iif(checked, QtCore.Qt.CheckState.Checked, QtCore.Qt.CheckState.Unchecked))
                    self.listWidget.addItem(item)
                    if self.is_line_layer(filename):
                        file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
                        if file_ext == "shp":
                            # Inspect the shapefile's geometry type before
                            # deciding whether to plot as a scatter (points)
                            # or as a line layer (polylines).
                            import shapefile as _shapefile
                            try:
                                reader = _shapefile.Reader(
                                    os.path.splitext(filename)[0] + ".shp")
                                is_point = reader.shapeType in (
                                    _shapefile.POINT,
                                    _shapefile.POINTZ,
                                    _shapefile.POINTM)
                                if is_point:
                                    coords = []
                                    for shape_obj in reader.shapes():
                                        pts = shape_obj.points
                                        if len(pts) > 0:
                                            coords.append((pts[0][0], pts[0][1]))
                                reader.close()
                            except Exception as e:
                                QtWidgets.QMessageBox.warning(self, "Warning",
                                    "Could not read shapefile %s: %s"
                                    % (filename, e))
                                continue
                            if is_point:
                                if len(coords) == 0:
                                    continue
                                coords = np.array(coords, dtype=np.float32)
                                self.plot_list[filename] = self.map.ax.plot(
                                    coords[:,0], coords[:,1],
                                    'o', markersize=int(size),
                                    color=color, picker=5)
                                self.plot_list[filename][0].set_visible(checked)
                            else:
                                self.shape_list[filename] = self.map.readshapefile(
                                    os.path.splitext(filename)[0], filename,
                                    linewidth=float(size), color=color)
                                shape = self.shape_list[filename]
                                if isinstance(shape, tuple) and len(shape) >= 5:
                                    shape[4].set_visible(checked)
                        else:
                            # line-GeoJSON
                            segments = loadGeoJSONLines(filename)
                            plotted_lines = []
                            for seg in segments:
                                lons = [p[0] for p in seg]
                                lats = [p[1] for p in seg]
                                line, = self.map.plot(lons, lats, '-',
                                                      color=color, linewidth=float(size))
                                plotted_lines.append(line)
                            self.shape_list[filename] = (None, None, None, None,
                                                         _GeoJSONLineLayer(plotted_lines))
                            self.shape_list[filename][4].set_visible(checked)
                    else:
                        headers, data = self.loadData(filename)
                        if len(headers) == 0:
                            QtWidgets.QMessageBox.critical(self, "Error", "Missing header line in data file!")
                            QtWidgets.QApplication.restoreOverrideCursor()
                            return
                        elif len(data) == 0:
                            QtWidgets.QMessageBox.critical(self, "Error", "Data file empty!")
                            QtWidgets.QApplication.restoreOverrideCursor()
                            return
                        coords = []
                        for row in data:
                            longitude = row[2]
                            latitude = row[3]
                            coords.append((longitude, latitude))
                        coords = np.array(coords, dtype=np.float32)
                        self.plot_list[filename] = self.map.ax.plot(
                            coords[:,0], coords[:,1], symbol,
                            markersize=int(size), color=color, picker=5)
                        plot = self.plot_list[filename]
                        plot[0].set_visible(checked)
                    filecount += 1
                else:
                    QtWidgets.QMessageBox.warning(self, "Warning",
                        "File " + filename + " not found!")
            QtWidgets.QApplication.restoreOverrideCursor()
            if filecount == 0:
                QtWidgets.QMessageBox.warning(self, "Warning", "No data files found for this project!")
                return
            self.listWidget.sortItems()
            self.setWindowTitle("Croizat - " + os.path.basename(str(self.filename)))
            self.updateUI()
            self.canvas.draw_idle()       

    def saveFile(self):
        self.filename, _ = QtWidgets.QFileDialog.getSaveFileName(self,
                    "Save File As", self.filename,
                    "Project (*.prj)")
        if self.filename:
            directory = os.path.dirname(str(self.filename))
            os.chdir(directory)
            self.saveProject()
            self.setWindowTitle("Croizat - " + os.path.basename(self.filename))
            self.updateUI()

    def addFile(self):
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(self,
                "Add File",
                "",
                "Comma-Separated Values (*.csv);; Spreadsheets (*.xls *.xlsx *.ods);; Dbase (*.dbf);; Keyhole Markup Language (*.kml);; Shapefile (*.shp);; GeoJSON (*.json *.geojson *.gjs)")
        if not filename:
            return

        src_path = str(filename)
        base_name = os.path.basename(src_path)
        if base_name in self.filelist:
            return

        file_ext = os.path.splitext(base_name)[1].lower().lstrip(".")

        # Copy the file (or, for shapefiles, the whole companion set) into
        # the project directory if it is not already there.  This keeps
        # the project self-contained and lets loadData() open the file
        # by basename, which is what the rest of the code expects.
        dst_path = os.path.join(self.project_dir(), base_name)
        if os.path.abspath(src_path) != os.path.abspath(dst_path):
            if file_ext == "shp":
                src_base = os.path.splitext(src_path)[0]
                dst_base = os.path.splitext(dst_path)[0]
                # The five core companions plus the two optional spatial
                # index files that ArcGIS sometimes produces.
                for ext in ('.shp', '.shx', '.dbf', '.prj', '.cpg',
                            '.sbn', '.sbx'):
                    src_ext = src_base + ext
                    if os.path.exists(src_ext):
                        try:
                            shutil.copy2(src_ext, dst_base + ext)
                        except (OSError, shutil.Error) as e:
                            QtWidgets.QMessageBox.critical(self, "Error",
                                "Cannot copy shapefile component %s:\n%s"
                                % (ext, e))
                            return
            else:
                try:
                    shutil.copy2(src_path, dst_path)
                except (OSError, shutil.Error) as e:
                    QtWidgets.QMessageBox.critical(self, "Error",
                        "Cannot copy file into project directory:\n%s" % e)
                    return

        filename = base_name
        attributes = []
        item = QtWidgets.QListWidgetItem(filename)
        item.setCheckState(QtCore.Qt.CheckState.Checked)
        self.listWidget.addItem(item)

        if file_ext in ["csv", "xls", "lsx", "xlsx", "ods", "dbf", "kml",
                        "json", "geojson", "gjs"]:     
            # GeoJSON can be points or lines.  Detect which, and route
            # line files to the line-loading path instead.
            if file_ext in ("json", "geojson", "gjs") and geojsonHasLines(base_name):
                segments = loadGeoJSONLines(base_name)
                if len(segments) == 0:
                    QtWidgets.QMessageBox.critical(self, "Error",
                        "GeoJSON file contains no usable line features.")
                    self.listWidget.takeItem(self.listWidget.row(item))
                    return
                color = random_color()
                width = 4
                plotted_lines = []
                for seg in segments:
                    lons = [p[0] for p in seg]
                    lats = [p[1] for p in seg]
                    line, = self.map.plot(lons, lats, '-',
                                          color=color, linewidth=width)
                    plotted_lines.append(line)
                # Store in shape_list under the .shp-style name so
                # everything downstream treats it as a line layer.
                display_name = base_name
                self.shape_list[display_name] = (nome, Nome, Nome, Nome,
                                                    _GeoJSONLineLayer(plotted_lines))
                self.filelist[display_name] = ['-', width, color, 1]
                item.setText(display_name)
                self.saveProject()
                self.updateUI()
                self.canvas.draw_idle()
                return                
            headers, data = self.loadData(filename)
            if len(headers) == 0:
                QtWidgets.QMessageBox.critical(self, "Error",
                    "Missing header line in data file!")
                self.listWidget.takeItem(self.listWidget.row(item))
                return
            elif len(data) == 0:
                QtWidgets.QMessageBox.critical(self, "Error",
                    "Data file empty!")
                self.listWidget.takeItem(self.listWidget.row(item))
                return
            coords = []
            for row in data:
                longitude = row[2]
                latitude = row[3]
                coords.append((longitude, latitude))
            coords = np.array(coords, dtype=np.float32)
            symbol = "o"
            size = 6
            color = random_color()
            attributes.append([symbol, size, color])
            self.plot_list[filename] = self.map.ax.plot(
                coords[:,0], coords[:,1], symbol,
                markersize=int(size), color=color, picker=5)
        elif file_ext == "shp":
            style = '-'
            width = 2
            color = random_color()
            attributes.append([style, width, color, 1])
            self.shape_list[filename] = self.map.readshapefile(
                os.path.splitext(filename)[0], filename,
                linewidth=float(width), color=color)

        self.filelist[filename] = attributes
        self.saveProject()
        self.updateUI()
        self.canvas.draw_idle()    

    def delFile(self):
        reply = QtWidgets.QMessageBox.question(self, "Remove file",
            "Remove file from project?",
            QtWidgets.QMessageBox.StandardButton.Yes 
            | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        if self.listWidget.count() > 0 and self.listWidget.selectedIndexes():
            index = self.listWidget.selectedIndexes()[0].row()
            item = self.listWidget.takeItem(index)
            filename = str(item.text())
            del item
            if self.is_line_layer(filename) and filename in self.shape_list:
                shape = self.shape_list[filename]
                if isinstance(shape, tuple) and len(shape) >= 5:
                    shape[4].set_visible(False)
            elif filename in self.plot_list:
                plot = self.plot_list[filename]
                plot[0].set_visible(False)    
                
            if filename in self.filelist:   
                self.filelist.pop(filename)
            self.saveProject()
            self.updateUI()
            self.canvas.draw()
            
    def detectOutliers(self):
        # Collect point-data files from the project
        datafiles = []
        for filename in self.filelist:
            file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
            if file_ext in ["csv", "xls", "lsx", "xlsx", "ods", "dbf",
                            "kml", "json", "geojson", "gjs"]:
                datafiles.append(filename)
        if len(datafiles) == 0:
            QtWidgets.QMessageBox.warning(self, "Detect Outliers",
                "No point-data files are loaded.")
            return

        # Ask for the threshold (how many standard deviations)
        k, ok = QtWidgets.QInputDialog.getDouble(
            self, "Detect Outliers",
            "Detect points beyond this many standard deviations\n"
            "from the dataset centre, or from their nearest neighbour:",
            value=3.0, min=1.0, max=10.0, decimals=1)
        if not ok:
            return

        datafiles.sort()
        self.statusBar().showMessage("Detecting outliers...")
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
        self.setCursor(QtCore.Qt.CursorShape.WaitCursor)
        QtWidgets.QApplication.processEvents()

        # Compute outliers for each file
        report_lines = []
        total_flagged = 0
        for filename in datafiles:
            headers, data = self.loadData(filename)
            if len(data) < 3:
                # Too few points for a meaningful outlier test
                continue

            lats, lons = [], []
            for row in data:
                try:
                    lons.append(float(row[2]))
                    lats.append(float(row[3]))
                except (ValueError, TypeError):
                    lons.append(None)
                    lats.append(None)

            # Compute centroid using only valid points
            valid = [(lo, la) for lo, la in zip(lons, lats)
                     if lo is not None and la is not None]
            if len(valid) < 3:
                continue
            cx = sum(p[0] for p in valid) / len(valid)
            cy = sum(p[1] for p in valid) / len(valid)

            # Distance of each point from centroid
            dists = []
            for lo, la in zip(lons, lats):
                if lo is None or la is None:
                    dists.append(None)
                else:
                    # Euclidean in degrees; adequate for outlier detection
                    d = ((lo - cx)**2 + (la - cy)**2) ** 0.5
                    dists.append(d)

            valid_dists = [d for d in dists if d is not None]
            mean_c = sum(valid_dists) / len(valid_dists)
            var_c = sum((d - mean_c)**2 for d in valid_dists) / len(valid_dists)
            sd_c = var_c ** 0.5
            cutoff_c = mean_c + k * sd_c

            # Nearest-neighbour distance for each point
            nn_dists = []
            for i in range(len(valid)):
                lo_i, la_i = valid[i]
                nearest = None
                for j in range(len(valid)):
                    if i == j:
                        continue
                    lo_j, la_j = valid[j]
                    d = ((lo_i - lo_j)**2 + (la_i - la_j)**2) ** 0.5
                    if nearest is None or d < nearest:
                        nearest = d
                nn_dists.append(nearest)

            mean_n = sum(nn_dists) / len(nn_dists)
            var_n = sum((d - mean_n)**2 for d in nn_dists) / len(nn_dists)
            sd_n = var_n ** 0.5
            cutoff_n = mean_n + k * sd_n

            # Detect points that exceed either cutoff
            flagged = []
            for i, (lo, la) in enumerate(zip(lons, lats)):
                if lo is None or la is None:
                    continue
                # Find this point's index in `valid` (for nn_dists access)
                # Simpler: recompute nearest for this specific point
                nearest = None
                for j, (lo_j, la_j) in enumerate(valid):
                    d = ((lo - lo_j)**2 + (la - la_j)**2) ** 0.5
                    if d > 0 and (nearest is None or d < nearest):
                        nearest = d
                if dists[i] is not None and dists[i] > cutoff_c:
                    flagged.append((i, lo, la, dists[i], "far from centre"))
                elif nearest is not None and nearest > cutoff_n:
                    flagged.append((i, lo, la, nearest, "far from nearest neighbour"))

            if len(flagged) > 0:
                report_lines.append("")
                report_lines.append("=== %s ===" % filename)
                report_lines.append("%d potential outlier(s) of %d points:"
                                    % (len(flagged), len(data)))
                for idx, lo, la, dist, reason in flagged:
                    report_lines.append(
                        "  Row %d: (%.4f, %.4f)  %s (%.3f°)"
                        % (idx + 2, lo, la, reason, dist))   # +2 for header + 1-based
                total_flagged += len(flagged)

        self.statusBar().showMessage("Ready")
        self.unsetCursor()
        QtWidgets.QApplication.restoreOverrideCursor()
        
        if total_flagged == 0:
            QtWidgets.QMessageBox.information(self, "Flag Outliers",
                "No outliers found at threshold %.1f standard deviations."
                % k)
            return

        # Show the report in a scrollable dialog
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Potential Outliers")
        dialog.resize(700, 500)
        text = QtWidgets.QPlainTextEdit()
        text.setReadOnly(True)
        text.setPlainText("\n".join(report_lines))
        text.setStyleSheet("font-family: Consolas, monospace;")
        closeBtn = QtWidgets.QPushButton("Close")
        closeBtn.clicked.connect(dialog.accept)
        button_row = QtWidgets.QHBoxLayout()
        button_row.addStretch(1)
        button_row.addWidget(closeBtn)
        button_row.addStretch(1)
        layout = QtWidgets.QVBoxLayout(dialog)
        layout.addWidget(text)
        layout.addLayout(button_row)
        dialog.exec()

    def exportFile(self):
        output_formats = ["shp", "kml", "geojson"]
        items = ("Shapefile (.shp)", "Keyhole Markup Language (.kml)", "GeoJSON (.geojson)")
        item, ok = QtWidgets.QInputDialog.getItem(self, "Export", "Output Formats:", items, 0, False)
        if ok and item:
            outfmt = output_formats[items.index(str(item))]
            datafiles = []
            for filename in self.filelist:
                file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
                if file_ext in ["csv", "xls", "xlsx", "ods", "dbf"] + iif(outfmt == "shp", ["kml"], []):
                    datafiles.append(filename)
            datafiles.sort()
            total = 0
            form = ChecklistDialog("Data Files", datafiles, checked=True)
            if form.exec():
                QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
                for s in form.choices:
                    headers, data = self.loadData(str(s))
                    if len(headers) == 0:
                        QtWidgets.QMessageBox.critical(self, "Error", 
                            "Missing header line in data file!")
                        return
                    elif len(data) == 0:
                        QtWidgets.QMessageBox.critical(self, "Error", 
                            "Data file empty!")
                        return
                    total += 1

                    if outfmt == "shp":
                        # Create shapefile
                        w = shapefile.Writer(get_basename(str(s)), shapeType=shapefile.POINT)
                        w.autoBalance = 1
                        # Define database scheme
                        w.field("NAME", 'C', 22)
                        w.field("DESCRIPTION", 'C', 72)
                        w.field("LONGITUDE", 'F', 11, 6)
                        w.field("LATITUDE", 'F', 11, 6)

                        count = 0
                        for row in data:
                            name = row[0]
                            description = row[1]
                            longitude = row[2]
                            latitude = row[3]
                            # Write point coordinates and database record
                            w.point(float(longitude), float(latitude))
                            w.record(name, description, float(longitude), float(latitude))
                            count += 1
                        w.close()

                    elif outfmt == "kml":
                        # Create KML file
                        kml = simplekml.Kml()
                        style = simplekml.Style()
                        style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"

                        count = 0
                        for row in data:
                            name = str(row[0])
                            description = str(row[1])
                            longitude = float(row[2])
                            latitude = float(row[3])
                            point = kml.newpoint(name=name, description=description, coords=[(longitude, latitude)])
                            point.style = style
                            count += 1
                        kml.save(get_basename(str(s)) + ".kml")
                        
                    elif outfmt == "geojson":
                        # Create GeoJSON FeatureCollection
                        features = []
                        for row in data:
                            name = str(row[0])
                            description = str(row[1])
                            longitude = float(row[2])
                            latitude = float(row[3])
                            feature = {
                                "type": "Feature",
                                "geometry": {
                                    "type": "Point",
                                    "coordinates": [longitude, latitude]
                                },
                                "properties": {
                                    "scientificName": name,
                                    "locality": description
                                }
                            }
                            features.append(feature)
                        geojson = {
                            "type": "FeatureCollection",
                            "features": features
                        }
                        with open(get_basename(str(s)) + ".geojson", "w", encoding="utf-8") as gjf:
                            json.dump(geojson, gjf, ensure_ascii=False, indent=2)    

                QtWidgets.QApplication.restoreOverrideCursor()
                QtWidgets.QMessageBox.information(self, "Information",
                    "%d file(s) exported." % (total))
                    
    def paeMatrix(self):
        # Build a PAE presence/absence matrix from the loaded point data.
        
        # 1. Collect point-data files
        datafiles = []
        for filename in self.filelist:
            file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
            if file_ext in ["csv", "xls", "lsx", "xlsx", "ods", "dbf",
                            "kml", "json", "geojson", "gjs"]:
                datafiles.append(filename)
        if len(datafiles) == 0:
            QtWidgets.QMessageBox.warning(self, "PAE Matrix",
                "No point-data files are loaded.")
            return False

        datafiles.sort()
        form = ChecklistDialog("Data Files for PAE", datafiles, checked=True)
        if not form.exec():
            return False
        if len(form.choices) < 1:
            QtWidgets.QMessageBox.warning(self, "PAE Matrix",
                "At least one data file must be selected.")
            return False

        # 2. Parameters
        options = [
            ('Grid cell size (degrees):', 5.0),
            ('Output format:', [1, 'NEXUS (.nex)', 'CSV (.csv)',
                                'TNT (.tnt)', 'PHYLIP (.phy)']),
            ('Output name:', 'pae_matrix'),
            ('Exclude cells with only one species', False),
            ('Include outgroup row', True),
        ]
        result = fedit(options, title="PAE Matrix",
                       comment="Build a presence/absence matrix of grid cells "
                               "(taxa) x species (characters) for Parsimony "
                               "Analysis of Endemicity.")
        if result is None:
            return False

        grid_size = float(result[0])       
        dx = result[1] - 1
        if idx < 0 or idx > 3:
            idx = 0
        out_format = ['nex', 'csv', 'tnt', 'phy'][idx]
        
        out_name = str(result[2]).strip()
        exclude_singletons = bool(result[3])
        include_outgroup = bool(result[4])

        if grid_size <= 0:
            QtWidgets.QMessageBox.warning(self, "PAE Matrix",
                "Grid cell size must be positive.")
            return False
        if not out_name:
            QtWidgets.QMessageBox.warning(self, "PAE Matrix",
                "Please provide an output name.")
            return False

        # 3. Build the presence/absence data
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
        QtWidgets.QApplication.processEvents()

        species_set = set()
        cell_set = set()
        presence = set()

        try:
            for filename in form.choices:
                self.statusBar().showMessage("Reading " + str(filename))
                QtWidgets.QApplication.processEvents()
                headers, data = self.loadData(str(filename))
                
                for row in data:
                    try:
                        species = str(row[0]).strip()
                        lon = float(row[2])
                        lat = float(row[3])
                    except (ValueError, TypeError, IndexError):
                        continue
                    if not species:
                        continue
                    cx = int(math.floor(lon / grid_size))
                    cy = int(math.floor(lat / grid_size))
                    cell = (cx, cy)
                    species_set.add(species)
                    cell_set.add(cell)
                    presence.add((species, cell))
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()
            self.statusBar().showMessage("Ready")

        if len(species_set) == 0 or len(cell_set) == 0:
            QtWidgets.QMessageBox.warning(self, "PAE Matrix",
                "No valid records found in the selected files.")
            return False

        # 4. Optionally drop singleton cells (only one species)
        if exclude_singletons:
            species_count = {}
            for (sp, cell) in presence:
                species_count[cell] = species_count.get(cell, 0) + 1
            cell_set = {c for c in cell_set if species_count.get(c, 0) > 1}
            if len(cell_set) == 0:
                QtWidgets.QMessageBox.warning(self, "PAE Matrix",
                    "All cells were singletons and were excluded. "
                    "Try a smaller grid cell size.")
                return False

        # 5. Sort cells north-to-south, then west-to-east
        cells_sorted = sorted(cell_set, key=lambda c: (-c[1], c[0]))
        species_sorted = sorted(species_set)

        def cell_label(cell):
            cx, cy = cell
            center_lon = cx * grid_size + grid_size / 2.0
            center_lat = cy * grid_size + grid_size / 2.0
            ns = 'N' if center_lat >= 0 else 'S'
            ew = 'E' if center_lon >= 0 else 'W'
            return "%s%02d_%s%03d" % (ns, abs(int(center_lat)),
                                      ew, abs(int(center_lon)))

        cell_labels = [cell_label(c) for c in cells_sorted]

        if len(cells_sorted) < 3:
            QtWidgets.QMessageBox.warning(self, "PAE Matrix",
                "The matrix has fewer than 3 cells; parsimony analysis "
                "requires at least 3 taxa plus outgroup. "
                "Try a smaller grid cell size.")
            return False

        # 6. Write the matrix
        proj_dir = self.project_dir()
        out_path = os.path.join(proj_dir, out_name)
        filepath = out_path + "." + out_format

        try:
            if out_format == "nex":
                self._write_pae_nexus(filepath, cell_labels, species_sorted,
                                      presence, cells_sorted, include_outgroup)
            elif out_format == "csv":
                self._write_pae_csv(filepath, cell_labels, species_sorted,
                                    presence, cells_sorted, include_outgroup)
            elif out_format == "tnt":
                self._write_pae_tnt(filepath, cell_labels, species_sorted,
                                    presence, cells_sorted, include_outgroup)
            elif out_format == "phy":
                self._write_pae_phylip(filepath, cell_labels, species_sorted,
                                       presence, cells_sorted, include_outgroup)                        
        except OSError as e:
            QtWidgets.QMessageBox.critical(self, "PAE Matrix",
                "Could not write matrix file:\n%s" % e)
            return False

        # 7. Success message
        n_taxa = len(cells_sorted) + (1 if include_outgroup else 0)
        n_chars = len(species_sorted)
        QtWidgets.QMessageBox.information(self, "PAE Matrix",
            "PAE matrix written to:\n%s\n\n"
            "%d cells (taxa) x %d species (characters).\n\n"
            "Open the file in TNT, PAUP*, PHYLIP, or another parsimony program "
            "to run the analysis."
            % (filepath, n_taxa, n_chars))
        return True
        
    def _write_pae_nexus(self, filepath, cell_labels, species_sorted,
                         presence, cells_sorted, include_outgroup):
        n_chars = len(species_sorted)
        n_taxa = len(cells_sorted) + (1 if include_outgroup else 0)
        label_width = max(len(l) for l in cell_labels + ["outgroup"]) + 2
        with open(filepath, "w", encoding="utf-8") as f:
            f.write("#NEXUS\n")
            f.write("BEGIN DATA;\n")
            f.write("    DIMENSIONS NTAX=%d NCHAR=%d;\n" % (n_taxa, n_chars))
            f.write('    FORMAT DATATYPE=STANDARD SYMBOLS="01" MISSING=? GAP=-;\n')
            f.write("    MATRIX\n")
            for i, cell in enumerate(cells_sorted):
                row = ''.join('1' if (sp, cell) in presence else '0'
                              for sp in species_sorted)
                f.write("    " + cell_labels[i].ljust(label_width) + row + "\n")
            if include_outgroup:
                f.write("    " + "outgroup".ljust(label_width) + "0" * n_chars + "\n")
            f.write("    ;\n")
            f.write("END;\n")

    def _write_pae_csv(self, filepath, cell_labels, species_sorted,
                       presence, cells_sorted, include_outgroup):
        with open(filepath, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Cell"] + species_sorted)
            for i, cell in enumerate(cells_sorted):
                row = [1 if (sp, cell) in presence else 0
                       for sp in species_sorted]
                writer.writerow([cell_labels[i]] + row)
            if include_outgroup:
                writer.writerow(["outgroup"] + [0] * len(species_sorted))

    def _write_pae_tnt(self, filepath, cell_labels, species_sorted,
                       presence, cells_sorted, include_outgroup):
        n_chars = len(species_sorted)
        n_taxa = len(cells_sorted) + (1 if include_outgroup else 0)
        label_width = max(len(l) for l in cell_labels + ["outgroup"]) + 2
        with open(filepath, "w", encoding="utf-8") as f:
            f.write("nstates 2;\n")
            f.write("xread\n")
            f.write("%d %d\n" % (n_taxa, n_chars))
            for i, cell in enumerate(cells_sorted):
                row = ''.join('1' if (sp, cell) in presence else '0'
                              for sp in species_sorted)
                f.write(cell_labels[i].ljust(label_width) + row + "\n")
            if include_outgroup:
                f.write("outgroup".ljust(label_width) + "0" * n_chars + "\n")
            f.write(";\n")
            f.write("proc/;\n")
            
    def _write_pae_phylip(self, filepath, cell_labels, species_sorted,
                          presence, cells_sorted, include_outgroup):
        n_chars = len(species_sorted)
        n_taxa = len(cells_sorted) + (1 if include_outgroup else 0)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write("%d %d\n" % (n_taxa, n_chars))
            for i, cell in enumerate(cells_sorted):
                # PHYLIP's traditional format: name padded to 10 columns,
                # then the character string.
                name = cell_labels[i][:10].ljust(10)
                row = ''.join('1' if (sp, cell) in presence else '0'
                              for sp in species_sorted)
                f.write(name + row + "\n")
            if include_outgroup:
                f.write("outgroup".ljust(10) + "0" * n_chars + "\n")            

    def about(self):
        from matplotlib import __version__
        matplotlib_version = __version__
        from mpl_toolkits.basemap import __version__
        basemap_version = __version__
        QtWidgets.QMessageBox.about(self, "About Croizat",
                """<b>Croizat</b> v {0}
                <p>Copyright &copy; 2003-2026 Mauro J. Cavalcanti.
                <p>A Software Package for Quantitative Analysis in Panbiogeography.
                <p><i>Earth and life evolve together.</i></p>
                <p>Running on Python: {1} <br>
                Qt version: {2} <br>
                PyQt version: {3} <br>
                NumPy version: {4} <br>
                Matplotlib version: {5} <br>
                Basemap version: {6} <br>
                PyShp version: {7} <br>
                Platform: {8} {9}""".format(
                VERSION, platform.python_version(),
                QtCore.qVersion(), QtCore.PYQT_VERSION_STR,
                np.__version__, matplotlib_version,
                basemap_version, shapefile.__version__,
                platform.system(), platform.release()))

    def closeEvent(self, event):
        reply = QtWidgets.QMessageBox.question(self, "Confirmation",
                        "Quit program?",
                        QtWidgets.QMessageBox.StandardButton.Yes, QtWidgets.QMessageBox.StandardButton.No)
        if reply == QtWidgets.QMessageBox.StandardButton.Yes:
            self.saveProject()
            event.accept()
        else:
            event.ignore()

    def track(self):
        datafiles = []
        for filename in self.filelist:
            file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
            if file_ext in ["csv", "xls", "xlsx", "ods", "dbf", "kml", "json", "geojson", "gjs"]:
                track_base = get_basename(str(filename)) + "_individualtrack"
                track_file = track_base + ".shp"
                if (track_file not in self.filelist or not os.path.exists(track_file)):
                    datafiles.append(filename)
                    
        if len(datafiles) == 0:
            QtWidgets.QMessageBox.warning(self, "Warning", 
                "All tracks have already been computed!")
            return False
        datafiles.sort()
        form = ChecklistDialog("Data Files", datafiles, checked=True)
        if form.exec():
            if len(form.choices) < 1:
                QtWidgets.QMessageBox.warning(self, "Warning", 
                    "At least one data file must be selected!")
                return False
            for s in form.choices:
                attributes = []
                headers, data = self.loadData(str(s))
                if len(headers) == 0:
                    QtWidgets.QMessageBox.critical(self, "Error", 
                        "Missing header line in data file!")
                    return
                elif len(data) == 0:
                    QtWidgets.QMessageBox.critical(self, "Error", 
                        "Data file empty!")
                    return

                self.statusBar().showMessage("Computing individual track for " + s)
                QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
                try:
                    plot = self.plot_list[str(s)]
                    color = plot[0].get_markerfacecolor()
                    width = 2
                    coords = []
                    for row in data:
                        longitude = row[2]
                        latitude = row[3]
                        coords.append((longitude, latitude))

                    if len(coords) < 2:
                        QtWidgets.QMessageBox.warning(
                            self, "Individual Track",
                            "At least two points are required to compute an individual track.")
                        continue

                    coords = np.array(coords, dtype=np.float32)
                    d = distance_matrix(coords)
                    edges = minimum_spanning_tree(d)
                    edge_list = []
                    for edge in edges:
                        i, j = edge
                        edge_list.append([(coords[i, 0], coords[i, 1]),
                                          (coords[j, 0], coords[j, 1])])

                    if len(edge_list) == 0:
                        QtWidgets.QMessageBox.warning(
                            self, "Individual Track",
                            "No track could be computed from the selected points.")
                        continue

                    # Save the minimum spanning tree
                    track_base = get_basename(str(s)) + "_individualtrack"
                    track_file = track_base + ".shp"
                    
                    for ext in ('.shp', '.shx', '.dbf', '.prj', '.cpg'):
                        f = track_base + ext
                        try:
                            if os.path.exists(f):
                                os.remove(f)
                        except OSError:
                            pass    
                    
                    w = shapefile.Writer(track_base, shapeType=shapefile.POLYLINE)
                    w.line(edge_list)
                    w.field("COMMON_ID", 'C')
                    w.record("Point")
                    w.close()

                    # Save KML file
                    kml = simplekml.Kml()
                    for i in range(len(edge_list)):
                        line = kml.newlinestring(name="Track" + str(i + 1), coords=edge_list[i])
                        line.style.linestyle.width = 3 #width
                        line.style.linestyle.color = simplekml.Color.red #color
                    kml.save(track_base + ".kml")

                    # Plot the minimum spanning tree.
                    # The layer name passed to readshapefile must not contain a dot.
                    layer_name = get_basename(track_file)
                    self.shape_list[track_file] = self.map.readshapefile(
                        track_base, layer_name, linewidth=width, color=color)
                    self.shape_list[track_file][4].set_visible(True)

                    # Store 4-element attributes: [style, width, color, state]
                    attributes = ['-', width, color, 1]
                    self.filelist[track_file] = attributes

                    item = QtWidgets.QListWidgetItem(track_file)
                    item.setCheckState(QtCore.Qt.CheckState.Checked)
                    self.listWidget.addItem(item)
                    self.saveProject()
                    self.canvas.draw_idle()
                except Exception as e:
                    QtWidgets.QMessageBox.critical(
                        self, "Individual Track",
                        "Error computing individual track:\n%s: %s" %
                        (type(e).__name__, str(e)))
                finally:
                    QtWidgets.QApplication.restoreOverrideCursor()
            self.statusBar().showMessage("Ready")

    def gentracks(self):
        datafiles = []
        for filename in self.filelist:
            file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
            if file_ext in ["csv", "xls", "xlsx", "ods", "dbf", "kml", "json", "geojson", "gjs"]:
                datafiles.append(filename)
        datafiles.sort()
        form = ChecklistDialog("Data Files", datafiles, checked=True)
        if form.exec():
            if len(form.choices) < 2:
                QtWidgets.QMessageBox.warning(self, "Warning", 
                    "At least two data files must be selected!")
                return False

            if os.path.exists("output0.dat"):
                os.remove("output0.dat")

            for s in form.choices:
                headers, indata = self.loadData(str(s))
                if len(headers) == 0:
                    QtWidgets.QMessageBox.critical(self, "Error", 
                        "Missing header line in data file!")
                    return
                elif len(indata) == 0:
                    QtWidgets.QMessageBox.critical(self, "Error", 
                        "Data file empty!")
                    return
                self.statusBar().showMessage("Preparing data for " + s)

                outdata = []
                for row in indata:
                    id = row[0].replace(' ', '_').replace('.', '')
                    lon = float(row[2])
                    lat = float(row[3])
                    outdata.append((id, lat, lon))

                # Output data for MartiTracks
                with open("output0.dat", "a", newline="", encoding="utf-8") as csvfile:
                    writer = csv.writer(csvfile, delimiter=' ', quoting=csv.QUOTE_NONE)
                    for row in outdata:
                        writer.writerow(row)

            self.statusBar().showMessage("Ready")

            options = [
                ('cv', 1.5),
                ('lmax', 2.0),
                ('lmin', 0.6),
                ('maxline', 3.0),
                ('ci', 0.8),
                ('strategy:', [0, 'croizat0', ' croizat1']),
                ('Create shapefile', True),
            ]
            result = fedit(options, title="MartiTracks Parameters", comment="Set parameters for generalized tracks.")
            if result is not None:
                # Create parameters file
                outfile = open("params.dat", "w", encoding="utf-8")
                outfile.write("set cv " + str(result[0]) + "\n")
                outfile.write("set lmax " + str(result[1]) + "\n")
                outfile.write("set lmin " + str(result[2]) + "\n")
                outfile.write("set maxline " + str(result[3]) + "\n")
                outfile.write("set ci " + str(result[4]) + "\n")
                outfile.write("kmlgen\n")
                outfile.write(iif(result[5] == 1, "croizat0", "croizat1"))
                outfile.close()

                # Execute MartiTracks
                if os.path.exists("output0.kml"):
                    os.remove("output0.kml")
                    
                self.statusBar().showMessage("Computing generalized tracks (this may take a while!)")
                QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
                exe = martitracks_executable()
                mt05 = subprocess.call([exe, "output0.dat", "output0.kml", "params.dat"], shell=True)
                
                QtWidgets.QApplication.restoreOverrideCursor()
                self.statusBar().showMessage("Ready")
                if mt05 != 0:
                    QtWidgets.QMessageBox.critical(self, "Error", 
                        "Cannot run MartiTracks program!")
                    return
                
                # Create shapefile
                if result[6]:
                    with open("output0.kml", "rb") as kmlf:
                        raw = kmlf.read()
                    try:
                        text = raw.decode("utf-8")
                    except UnicodeDecodeError:
                        text = raw.decode("latin-1")
                    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', text)
                    text = re.sub(r'&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9A-Fa-f]+);)', '&amp;', text)
                    try:
                        doc = et.fromstring(text)
                    except et.ParseError as e:
                        QtWidgets.QMessageBox.critical(self, "Generalized Tracks",
                            "MartiTracks produced invalid XML:\n%s\n\n"
                            "This is a limitation of the MartiTracks program; "
                            "the KML file 'output0.kml' has been left in the "
                            "project directory for inspection." % e)
                        return
                    nmsp = '{http://earth.google.com/kml/2.1}'

                    for pm in doc.iterfind('.//{0}Placemark'.format(nmsp)):
                        name = pm.find('{0}name'.format(nmsp)).text.replace(',', '_')[:100]
                        self.statusBar().showMessage("Computing generalized track " + name)

                        lines = []
                        for ls in pm.iterfind('{0}MultiGeometry/{0}LineString/{0}coordinates'.format(nmsp)):
                            x = ls.text.strip().replace('\n','').split()[0]
                            y = ls.text.strip().replace('\n','').split()[1]
                            x0 = float(x.split(',')[0])
                            y0 = float(y.split(',')[0])
                            x1 = float(x.split(',')[1])
                            y1 = float(y.split(',')[1])
                            lines.append([(x0, x1),(y0, y1)])
                        lines = list(sorted(lines))
                        
                        track_base = name

                        for ext in ('.shp', '.shx', '.dbf', '.prj', '.cpg'):
                            f = track_base + ext
                            try:
                                if os.path.exists(f):
                                    os.remove(f)
                            except OSError:
                                pass

                        w = shapefile.Writer(name, shapeType=shapefile.POLYLINE)
                        w.line(lines)
                        w.field("COMMON_ID", 'C')
                        w.record("Point")
                        w.close()

                        color = random_color()
                        width = 4
                        style = '-'

                        # Plot the generalized track
                        self.shape_list[name + ".shp"] = self.map.readshapefile(track_base, track_base, linewidth=width, color=color)
                        self.filelist[name + ".shp"] = [style, width, color, 1]
                        item = QtWidgets.QListWidgetItem(name + ".shp")
                        item.setCheckState(QtCore.Qt.CheckState.Checked)
                        self.listWidget.addItem(item)
                        self.saveProject()
                        self.canvas.draw_idle()
                    self.statusBar().showMessage("Ready")

    def panbiotracks(self):
        datafiles = []
        for filename in self.filelist:
            if filename.endswith("_individualtrack.shp"):
                datafiles.append(filename)
        if len(datafiles) < 2:
            QtWidgets.QMessageBox.warning(self, "Warning",
                "At least two individual tracks must be computed first!")
            return False

        datafiles.sort()
        form = ChecklistDialog("Individual Track Files", datafiles, checked=True)
        if form.exec():
            if len(form.choices) < 2:
                QtWidgets.QMessageBox.warning(self, "Warning",
                    "At least two individual tracks must be selected!")
                return False

            options = [
                ('Output format:', [0,
                    'ESRI Shapefile (.shp)',
                    'GeoJSON (.geojson)']),
                ('Output name:', 'generalized_tracks'),
            ]
            result = fedit(options, title="Panbiotracks Generalized Tracks",
                           comment="Set parameters for generalized-track computation.")
            if result is not None:
                out_format = ['shp', 'gjs'][result[0] - 1]
                out_name = str(result[1]).strip()
                if not out_name:
                    QtWidgets.QMessageBox.warning(self, "Warning",
                        "Please provide an output name.")
                    return False

                exe = panbiotracks_executable()
                if exe is None:
                    return False
                if not os.path.exists(exe):
                    QtWidgets.QMessageBox.critical(self, "Error",
                        "Panbiotracks executable not found: " + exe)
                    return False

                proj_dir = self.project_dir()
                out_path = os.path.join(proj_dir, out_name)

                cmd = [exe, "-m", "P", "-i"] + [str(c) for c in form.choices] + \
                      ["-o", out_path, "-of", out_format]

                self.statusBar().showMessage("Computing generalized tracks with Panbiotracks")
                QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
                try:
                    proc = subprocess.run(cmd, capture_output=True, text=True)
                finally:
                    QtWidgets.QApplication.restoreOverrideCursor()
                    self.statusBar().showMessage("Ready")

                if proc.returncode != 0:
                    QtWidgets.QMessageBox.critical(self, "Panbiotracks error",
                        "Exit code %d\n\nSTDOUT:\n%s\n\nSTDERR:\n%s" %
                        (proc.returncode, proc.stdout, proc.stderr))
                    return False

                if out_format == "shp":
                    shp_path = out_path + ".shp"
                    if not os.path.exists(shp_path):
                        QtWidgets.QMessageBox.warning(self, "Warning",
                            "Expected output shapefile not found: " + shp_path)
                        return False
                    base = os.path.splitext(os.path.basename(shp_path))[0]
                    color = random_color()
                    self.shape_list[os.path.basename(shp_path)] = self.map.readshapefile(
                        os.path.splitext(shp_path)[0], base,
                        linewidth=4, color=color)
                    self.filelist[os.path.basename(shp_path)] = ['-', 4, color, 1]
                    item = QtWidgets.QListWidgetItem(os.path.basename(shp_path))
                    item.setCheckState(QtCore.Qt.CheckState.Checked)
                    self.listWidget.addItem(item)
                    self.saveProject()
                    self.updateUI()
                    self.canvas.draw_idle()

                elif out_format == "gjs":
                    gjs_path = out_path + ".geojson"
                    if not os.path.exists(gjs_path):
                        for ext in (".json", ".gjs"):
                            candidate = out_path + ext
                            if os.path.exists(candidate):
                                gjs_path = candidate
                                break
                        else:
                            QtWidgets.QMessageBox.warning(self, "Warning",
                                "Expected output GeoJSON file not found: "
                                + gjs_path)
                            return False

                    segments = loadGeoJSONLines(gjs_path)
                    if len(segments) == 0:
                        QtWidgets.QMessageBox.warning(self, "Warning",
                            "Panbiotracks output contains no line features.")
                        return False

                    color = random_color()
                    width = 4
                    plotted_lines = []
                    for seg in segments:
                        lons = [p[0] for p in seg]
                        lats = [p[1] for p in seg]
                        line, = self.map.plot(lons, lats, '-',
                                              color=color, linewidth=width)
                        plotted_lines.append(line)

                    real_name = os.path.basename(gjs_path)
                    self.shape_list[real_name] = (None, None, None, None,
                                                  _GeoJSONLineLayer(plotted_lines))
                    self.filelist[real_name] = ['-', width, color, 1]
                    item = QtWidgets.QListWidgetItem(real_name)
                    item.setCheckState(QtCore.Qt.CheckState.Checked)
                    self.listWidget.addItem(item)
                    self.saveProject()
                    self.updateUI()
                    self.canvas.draw_idle()

                self.statusBar().showMessage("Ready")
                return True
        return False

    def nodes(self):
        datafiles = []
        for filename in self.filelist:
            if self.is_line_layer(filename):
                lowered = filename.lower()
                if ("_ mst" in lowered
                        or "generalized" in lowered
                        or "panbiotracks" in lowered):
                    datafiles.append(filename)
        if len(datafiles) < 2:
            QtWidgets.QMessageBox.warning(self, "Warning",
                "At least two generalized tracks must be computed first.")
            return False

        datafiles.sort()
        form = ChecklistDialog("Generalized Track Files", datafiles, checked=True)
        if form.exec():
            if len(form.choices) < 2:
                QtWidgets.QMessageBox.warning(self, "Warning",
                    "At least two generalized tracks must be selected!")
                return False

            options = [
                ('Output format:', [0,
                    'ESRI Shapefile (.shp)',
                    'GeoJSON (.geojson)']),
                ('Output name:', 'nodes_output'),
            ]
            result = fedit(options, title="Panbiotracks Nodes",
                           comment="Set parameters for node detection.")
            if result is not None:
                out_format = ['shp', 'gjs'][result[0] - 1]
                out_name = str(result[1]).strip()
                if not out_name:
                    QtWidgets.QMessageBox.warning(self, "Warning",
                        "Please provide an output name.")
                    return False

                exe = panbiotracks_executable()
                if exe is None:
                    return False
                if not os.path.exists(exe):
                    QtWidgets.QMessageBox.critical(self, "Error",
                        "Panbiotracks executable not found: " + exe)
                    return False

                proj_dir = self.project_dir()
                out_path = os.path.join(proj_dir, out_name)

                cmd = [exe, "-m", "N", "-i"] + [str(c) for c in form.choices] + \
                      ["-o", out_path, "-of", out_format]

                self.statusBar().showMessage("Computing panbiogeographic nodes")
                QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
                try:
                    proc = subprocess.run(cmd, capture_output=True, text=True)
                finally:
                    QtWidgets.QApplication.restoreOverrideCursor()
                    self.statusBar().showMessage("Ready")

                if proc.returncode != 0:
                    QtWidgets.QMessageBox.critical(self, "Panbiotracks error",
                        "Exit code %d\n\nSTDOUT:\n%s\n\nSTDERR:\n%s" %
                        (proc.returncode, proc.stdout, proc.stderr))
                    return False

                if out_format == "shp":
                    shp_path = out_path + ".shp"
                    if not os.path.exists(shp_path):
                        QtWidgets.QMessageBox.warning(self, "Warning",
                            "Expected output shapefile not found: " + shp_path)
                        return False

                    # Nodes are POINT features. readshapefile() returns a
                    # 4-tuple for point shapefiles but a 5-tuple for
                    # polyline shapefiles, so it cannot be used here.
                    # Read the points directly and plot them as a scatter,
                    # the same way GeoJSON nodes are plotted.
                    import shapefile as _shapefile
                    reader = _shapefile.Reader(shp_path)
                    coords = []
                    for shape in reader.shapes():
                        if shape.shapeType in (_shapefile.POINT,
                                               _shapefile.POINTZ,
                                               _shapefile.POINTM):
                            pts = shape.points
                            if len(pts) > 0:
                                coords.append((pts[0][0], pts[0][1]))
                    reader.close()
                    if len(coords) == 0:
                        QtWidgets.QMessageBox.warning(self, "Warning",
                            "Nodes shapefile contains no point features.")
                        return False

                    coords = np.array(coords, dtype=np.float32)
                    color = random_color()
                    real_name = os.path.basename(shp_path)
                    self.plot_list[real_name] = self.map.ax.plot(
                        coords[:,0], coords[:,1],
                        'o', markersize=6, color=color, picker=5)
                    self.filelist[real_name] = ['o', 6, color, 1]
                    item = QtWidgets.QListWidgetItem(real_name)
                    item.setCheckState(QtCore.Qt.CheckState.Checked)
                    self.listWidget.addItem(item)
                    self.saveProject()
                    self.updateUI()
                    self.canvas.draw_idle()

                elif out_format == "gjs":
                    gjs_path = out_path + ".geojson"
                    if not os.path.exists(gjs_path):
                        for ext in (".json", ".gjs"):
                            candidate = out_path + ext
                            if os.path.exists(candidate):
                                gjs_path = candidate
                                break
                        else:
                            QtWidgets.QMessageBox.warning(self, "Warning",
                                "Expected output GeoJSON file not found: "
                                + gjs_path)
                            return False

                    if geojsonHasLines(gjs_path):
                        segments = loadGeoJSONLines(gjs_path)
                        color = random_color()
                        width = 4
                        plotted_lines = []
                        for seg in segments:
                            lons = [p[0] for p in seg]
                            lats = [p[1] for p in seg]
                            line, = self.map.plot(lons, lats, '-',
                                                  color=color, linewidth=width)
                            plotted_lines.append(line)
                        real_name = os.path.basename(gjs_path)
                        self.shape_list[real_name] = (None, None, None, None,
                                                      _GeoJSONLineLayer(plotted_lines))
                        self.filelist[real_name] = ['-', width, color, 1]
                        item = QtWidgets.QListWidgetItem(real_name)
                        item.setCheckState(QtCore.Qt.CheckState.Checked)
                        self.listWidget.addItem(item)
                    else:
                        headers, data = self.loadData(gjs_path)
                        if len(data) == 0:
                            QtWidgets.QMessageBox.warning(self, "Warning",
                                "Panbiotracks output contains no features.")
                            return False
                        coords = []
                        for row in data:
                            coords.append((row[2], row[3]))
                        coords = np.array(coords, dtype=np.float32)
                        color = random_color()
                        real_name = os.path.basename(gjs_path)
                        self.plot_list[real_name] = self.map.ax.plot(
                            coords[:,0], coords[:,1],
                            'o', markersize=6, color=color, picker=5)
                        self.filelist[real_name] = ['o', 6, color, 1]
                        item = QtWidgets.QListWidgetItem(real_name)
                        item.setCheckState(QtCore.Qt.CheckState.Checked)
                        self.listWidget.addItem(item)

                    self.saveProject()
                    self.updateUI()
                    self.canvas.draw_idle()

                self.statusBar().showMessage("Ready")
                return True
        return False
    
    def massings(self):
        # Palette selection (persisted in the project file)
        palette_names = list(MASSING_PALETTES.keys())
        try:
            current_index = palette_names.index(self.massing_palette)
        except ValueError:
            current_index = 0
            
        options = [
            ('Color palette:', [current_index + 1, *palette_names]),
            ('Cell size (degrees):', 50),
        ]
        result = fedit(options, title="Main Massings",
                               comment="Select a color palette and a grid cell size "
                                       "for the species-richness map.")
        if result is None:
            return
        self.massing_palette = palette_names[result[0] - 1]
        self.saveProject()
        
        ##size, ok = QtWidgets.QInputDialog.getInt(
        ##self, "Main Massings", "Enter cell size (degrees): ",
        ##value=50, min=1, max=90, step=1)
        ##if not ok or not size:
        ##    return
        
        size = int(result[1])
        if size < 1 or size > 90:
            QtWidgets.QMessageBox.warning(self, "Main Massings",
                "Cell size must be between 1 and 90 degrees.")
            return

        # Collect all point data from the currently loaded data files
        all_data = []
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
        try:
            for i in range(self.listWidget.count()):
                item = self.listWidget.item(i)
                filename = str(item.text())
                file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
                if file_ext in ["csv", "xls", "xlsx", "ods", "dbf", "kml", "json", "gjs"]:
                    headers, data = self.loadData(filename)
                    all_data += data
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()

        if len(all_data) == 0:
            QtWidgets.QMessageBox.warning(self, "Main Massings", 
                "No point data available.")
            return

        lats, lons = [], []
        for row in all_data:
            try:
                lat = float(row[3])
                lon = float(row[2])
            except (ValueError, TypeError):
                continue
            lats.append(lat)
            lons.append(lon)

        if len(lats) < 2:
            QtWidgets.QMessageBox.warning(self, "Main Massings", 
                "Not enough valid coordinates to compute massings.")
            return

        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
        
        # Reuse the current figure/axes; do not create a new window.
        axes = self.map.ax
        axes.clear()

        # Redraw the basemap on the freshly cleared axes
        self.map.drawcoastlines()
        self.map.drawmapboundary(fill_color=self.watercolor)
        self.map.fillcontinents(color=self.landcolor, lake_color=self.watercolor)
        if self.borderAction.isChecked():
            self.map.drawcountries()
        if self.riverAction.isChecked():
            self.map.drawrivers()

        # Bin the data into a size x size grid covering the data extent
        nx = int(size)
        ny = int(size)
        db = 1.0   # bin padding in degrees
        
        # Determine the visible extent from the Basemap object.
        # For a cylindrical projection these are just degrees; for others,
        # Basemap still exposes llcrnrlon etc.
        lon_min = self.map.llcrnrlon
        lon_max = self.map.urcrnrlon
        lat_min = self.map.llcrnrlat
        lat_max = self.map.urcrnrlat
        
        lon_bins = np.linspace(lon_min, lon_max, nx + 1)
        lat_bins = np.linspace(lat_min, lat_max, ny + 1)

        # density[i, j] = number of points in lat_bins[i] x lon_bins[j]
        density, _, _ = np.histogram2d(lats, lons, [lat_bins, lon_bins])

        # Project bin edges to map coordinates
        lon_bins_2d, lat_bins_2d = np.meshgrid(lon_bins, lat_bins)
        xs, ys = self.map(lon_bins_2d, lat_bins_2d)

        # Custom colormap
        ##cmap = LinearSegmentedColormap.from_list("massings", ["#ffffcc", "#fd8d3c", "#800026"])
        palette_colors = MASSING_PALETTES.get(
            self.massing_palette, MASSING_PALETTES["Yellow-Orange-Red"])
        cmap = LinearSegmentedColormap.from_list("massings", palette_colors)

        # Plot the density grid on the map
        mesh = self.map.pcolormesh(xs, ys, density, cmap=cmap, norm=LogNorm(vmin=1, vmax=density.max() if density.max() > 0 else 1))
        
        # Add a colorbar to the right of the map axes
        # Remove any previous colorbar to avoid stacking on repeated runs
        if hasattr(self, "_massing_colorbar") and self._massing_colorbar is not None:
            try:
                self._massing_colorbar.remove()
            except Exception:
                pass
        self._massing_colorbar = self.figure.colorbar(
            mesh, ax=axes, location="right", shrink=0.7,
            label="Species richness")

        axes.set_title(self.title)
        self.canvas.draw_idle()                
        QtWidgets.QApplication.restoreOverrideCursor()

    def borders(self):
        if self.borderAction.isChecked():
            color = iif((self.visibleearth or self.etopomodel or self.shadedrelief), '#ccccee', self.bordercolor)
            self.borders = self.map.drawcountries(color=color)
        else:
            self.borders.remove()
        self.canvas.draw_idle()

    def rivers(self):
        if self.riverAction.isChecked():
            color = iif((self.visibleearth or self.etopomodel or self.shadedrelief), '#ccccee', self.bordercolor)
            QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
            self.rivers = self.map.drawrivers(color=color)
            QtWidgets.QApplication.restoreOverrideCursor()
        else:
            self.rivers.remove()
        self.canvas.draw_idle()

    def defaultmap(self):
        if self.defaultmapAction.isChecked():
            self.visibleearth = False
            self.etopomodel = False
            self.shadedrelief = False
            self.compute_initial_figure()

    def bluemarble(self):
        if self.bluemarbleAction.isChecked():
            scale, ok = QtWidgets.QInputDialog.getDouble(self,"Display Blue Marble","Enter scale factor: ",
                        value=0.50, min=0.25, max=0.75)
            if ok:
                self.etopomodel = False
                self.shadedrelief = False
                self.visibleearth = True
                self.scale = scale
        else:
            self.visibleearth = False
        self.compute_initial_figure()

    def etopo(self):
        if self.etopomodelAction.isChecked():
            scale, ok = QtWidgets.QInputDialog.getDouble(self,"Display ETOPO Global Relief Model","Enter scale factor: ",
                        value=0.50, min=0.25, max=0.75)
            if ok:
                self.visibleearth = False
                self.shadedrelief = False
                self.etopomodel = True
                self.scale = scale
        else:
            self.etopomodel = False
        self.compute_initial_figure()

    def shaded(self):
        if self.shadedreliefAction.isChecked():
            scale, ok = QtWidgets.QInputDialog.getDouble(self,"Display Shaded Relief","Enter scale factor: ",
                        value=0.35, min=0.15, max=0.45)
            if ok:
                self.visibleearth = False
                self.etopomodel = False
                self.shadedrelief = True
                self.scale = scale
        else:
            self.shadedrelief = False
        self.compute_initial_figure()

    def mapresolution(self):
        items = ["crude", "low", "intermediate", "high"]
        matching = [s for s in items if any(xs in s for xs in self.resolution)]
        if not matching:
            matching = ['crude']
        pos = items.index(matching[0])
        options = [
                ("Resolution:", [pos] + items),
                ("DPI:", self.dpi),
        ]
        result = fedit(options, title="Map Resolution", comment="Set map resolution.")
        if result is not None:
            self.resolution = items[result[0] - 1][0]
            self.dpi = int(result[1])
            self.compute_initial_figure()

    def maplimits(self):
        options = [
                ("North (Latitude):", self.north),
                ("South (Latitude):", self.south),
                ("East (Longitude):", self.east),
                ("West (Longitude):", self.west),
            ]
        result = fedit(options, title="Map Limits", comment="Set map limits.")
        if result is not None:
            self.north = int(result[0])
            self.south = int(result[1])
            self.east = int(result[2])
            self.west = int(result[3])
            self.compute_initial_figure()

    def mapcenter(self):
        options = [
                ("Latitude:", self.clat),
                ("Longitude:", self.clon),
        ]
        result = fedit(options, title="Map Center", comment="Set map center.")
        if result is not None:
            self.clat = int(result[0])
            self.clon = int(result[1])
            self.compute_initial_figure()

    def maptitle(self):
        text, ok = QtWidgets.QInputDialog.getText(self, "Map Title", "Enter map title: ", QtWidgets.QLineEdit.EchoMode.Normal, self.title)
        if ok:
            self.title = str(text)
            self.compute_initial_figure()

    def mapcolors(self):
        options = [
            ("Land:", self.landcolor),
            ("Water:", self.watercolor),
            ("Border:", self.bordercolor),
        ]
        result = fedit(options, title="Map Colors", comment="Set map colors.")
        if result is not None:
            self.landcolor = result[0]
            self.watercolor = result[1]
            self.bordercolor = result[2]
            self.compute_initial_figure()

    def copytoclipboard(self):
        buffer = io.BytesIO()
        self.figure.savefig(buffer)
        QtWidgets.QApplication.clipboard().setImage(QtGui.QImage.fromData(buffer.getvalue()))
        buffer.close()

    def reload(self):
        if self.listWidget.count() > 0:
            self.compute_initial_figure()

    def editProperties(self):
        index = self.listWidget.currentIndex()
        if index.isValid():
            item = self.listWidget.itemFromIndex(index)
            filename = str(item.text())
            if self.is_line_layer(filename) and filename in self.shape_list:
                shape = self.shape_list[filename]
                if not isinstance(shape, tuple) or len(shape) < 5:
                    return
                width = int(shape[4].get_linewidth()[0])
                color = colors.to_hex(shape[4].get_color()[0])
                line_options = [
                    (None, "File: " + filename),
                    (None, None),
                    ("Width:", width),
                    ("Color:", color),
                ]
                result = fedit(line_options, title="Line Properties",
                               comment="Set line properties.")
                if result is not None:
                    color = result[1]
                    shape[4].set_linewidth(result[0])
                    shape[4].set_color(color)
            else:
                #headers, data = self.loadData(filename)
                #recs = len(data)
                file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
                if file_ext == "shp":
                    # Shapefiles store their attribute records in the DBF
                    # companion file.  loadData() cannot read shapefiles,
                    # so query the record count directly.
                    import shapefile as _shapefile
                    try:
                        reader = _shapefile.Reader(
                            os.path.splitext(filename)[0] + ".shp")
                        recs = reader.numRecords
                        reader.close()
                    except Exception:
                        recs = 0
                else:
                    headers, data = self.loadData(filename)
                    recs = len(data)
                plot = self.plot_list[filename]
                pos = symbols[plot[0].get_marker()]
                size = int(plot[0].get_markersize())
                color = plot[0].get_markerfacecolor()
                #point_options = [
                #    (None, "File: " + filename),
                #    (None, "Records: " + str(recs)),
                #    (None, None),
                if recs > 0:
                    point_options = [
                        (None, "File: " + filename),
                        (None, "Records: " + str(recs)),
                        (None, None),
                        ("Symbol:", [pos, "o (circle)", "^ (triangle up)",
                            "v (triangle down)", "< (triangle left)",
                            "> (triangle right)", "s (square)", "+ (plus)",
                            "x (cross)", "D (diamond)", "d (thin diamond)",
                            "h (hexagon)", "p (pentagon)", "I (vertical line)",
                            "_ (horizontal line)"]),
                        ("Size:", size),
                        ("Color:", color),
                    ]
                else:
                    point_options = [
                        (None, "File: " + filename),
                        (None, None),
                        ("Symbol:", [pos, "o (circle)", "^ (triangle up)",
                            "v (triangle down)", "< (triangle left)",
                            "> (triangle right)", "s (square)", "+ (plus)",
                            "x (cross)", "D (diamond)", "d (thin diamond)",
                            "h (hexagon)", "p (pentagon)", "I (vertical line)",
                            "_ (horizontal line)"]),
                        ("Size:", size),
                        ("Color:", color),
                    ]
                result = fedit(point_options, title="Point Properties",
                               comment="Set point properties.")
                if result is not None:
                    idx = result[0] - 1
                    if idx < 0 or idx >= len(symbols):
                        idx = 0
                    symbol = list(symbols.keys())[idx]
                    size = result[1]
                    color = result[2]
                    plot[0].set_marker(symbol)
                    plot[0].set_markersize(size)
                    plot[0].set_markerfacecolor(color)
                    plot[0].set_markeredgecolor(color)
            self.canvas.draw_idle()

    def listWidgetItemChanged(self, item):
        filename = str(item.text())
        checked = item.checkState() == QtCore.Qt.CheckState.Checked
        if self.is_line_layer(filename) and filename in self.shape_list:
            shape = self.shape_list[filename]
            if not isinstance(shape, tuple) or len(shape) < 5:
                return
            shape[4].set_visible(checked)
        else:
            plot = self.plot_list[filename]
            plot[0].set_visible(checked)
        self.canvas.draw_idle()

    def pick(self, event):
        if self.listWidget.count() == 0: return
        thisline = event.artist
        xdata = thisline.get_xdata()
        ydata = thisline.get_ydata()
        ind = event.ind
        if event.ind is None:
            return
        points = tuple(zip(xdata[ind], ydata[ind]))
        x = round(points[0][0],4)
        y = round(points[0][1],4)
        all_data = []
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
        for i in range(self.listWidget.count()):
            item = self.listWidget.item(i)
            filename = str(item.text())
            file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
            if file_ext in ["csv", "xls", "xlsx", "ods", "dbf", "kml", "json", "geojson", "gjs"]:
                headers, data = self.loadData(filename)
                all_data += data
        QtWidgets.QApplication.restoreOverrideCursor()
        match = [coord for coord in all_data if round(float(coord[2]),4) == x and round(float(coord[3]),4) == y]
        if not match:
            return
        recinfo = "Scientific Name: %s\n Location: %s\n Latitude: %f\n Longitude: %f" % (match[0][0], match[0][1], match[0][2], match[0][3])
        QtWidgets.QMessageBox.information(self, 
            "Record Information", recinfo)

    def project_dir(self):
        if self.filename:
            return os.path.dirname(os.path.abspath(str(self.filename)))
        return os.getcwd()
        
    def is_line_layer(self, filename):
        # Return True if `filename` is stored as a line layer.
        if filename in self.shape_list:
            return True
        if filename in self.plot_list:
            return False
        file_ext = os.path.splitext(filename)[1].lower().lstrip(".")
        if file_ext == "shp":
            return True
        if file_ext in ("json", "geojson", "gjs"):
            return geojsonHasLines(filename)
        return False
    
    def change_xlimits(self, axes):
        self.west = axes.get_xlim()[0]
        self.east = axes.get_xlim()[1]

    def change_ylimits(self, axes):
        self.south = axes.get_ylim()[0]
        self.north = axes.get_ylim()[1]

def main():
    app = QtWidgets.QApplication(sys.argv)
    QtCore.QSettings.setDefaultFormat(QtCore.QSettings.Format.IniFormat)
    app.setOrganizationName("Ecoinformatics Studio")
    app.setOrganizationDomain("croizat.sourceforge.net")
    app.processEvents()

    splash = None
    splash_path = getImagePath("croizat.jpg")
    if os.path.exists(splash_path):
        pixmap = QtGui.QPixmap(splash_path)
        splash = QtWidgets.QSplashScreen(
            pixmap, QtCore.Qt.WindowType.WindowStaysOnTopHint)
        splash.show()
        app.processEvents()

    main = MainWindow()
    main.showMaximized()
    if splash is not None:
        splash.finish(main)
    sys.exit(app.exec())

if __name__ == "__main__":
    if getattr(sys, "frozen", False):
        from logging_utils import setup_logging_to_file, handle_errors
        setup_logging_to_file(os.path.basename(sys.executable + '.log'))
        sys.excepthook = handle_errors
    main()