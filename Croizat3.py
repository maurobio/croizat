#! /usr/bin/python
# -*- coding: utf-8 -*-
#=================================================================================#
#  Croizat - A Software Package for Quantitative Analysis in Panbiogeography      #
#                 Copyright 2003-2026 Mauro J. Cavalcanti                         #
#                           maurobio@gmail.com                                    #
#                                                                                 #
#  This program is free software; you can redistribute it and/or modify           #
#  it under the terms of the GNU General Public License as published by           #
#  the Free Software Foundation; either version 3 of the License, or              #
#  (at your option) any later version.                                            #
#                                                                                 #
#  This program is distributed in the hope that it will be useful,                #
#  but WITHOUT ANY WARRANTY; without even the implied warranty of                 #
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                  #
#  GNU General Public License for more details.                                   #
#                                                                                 #
#  You should have received a copy of the GNU General Public License              #
#  along with this program. If not, see <http://www.gnu.org/licenses/>.           #
#                                                                                 #
#  Requirements:                                                                  #
#    Python 3.14+ (www.python.org)                                                #
#    PyQt 6.11+ (www.riverbankcomputing.com/software/pyqt)                        #
#    PyShp 3.1.6+ (github.com/GeospatialPython/pyshp)                             #
#    NumPy 2.5.3+ (www.numpy.org)                                                 #
#    Matplotlib 3.1.1+ (www.matplotlib.org)                                       #
#    Basemap 2.0.0+ (www.matplotlib.org/basemap)                                  #                                                                               #
#    simplekml 1.3+ (www.simplekml.com)                                           #
#    xlrd 1.1+ (www.python-excel.org)                                             #
#    ezodf 0.2+ (pythonhosted.org/ezodf)                                          #
#    xlwt 0.7+ (www.python-excel.org)                                             #                                                                                #
#    xlsxwriter 1.0+ (github.com/jmcnamara/XlsxWriter)                            #
#    dbfpy3 4.2.3+ (github.com/frankyxhl/dbfpy3)                                  # 
#                                                                                 #
#  REVISION HISTORY:                                                              #
#    Version 2.00, 20 Oct 2017 - First public release ("Africa")                  #
#    Version 2.10, 25 Oct 2017 - Seconde public release ("Tenasserim")            #
#      * Added new Background submenu to Map menu, with                           #
#        Blue Marble, ETOPO and shaded relief images as options                   #
#      * Added several minor improvements to the interface                        #
#   Version 2.20, 3 Jun 2018 = Third public release ("Neo-Caledonia")             #
#      * Added buttons to select all/unselect all items in checkbox dialogues     #
#      * Added options for toggle on/off the display of records and tracks in     #
#        the layers context menu                                                  #
#      * Changed the state of checkbox items to be selected by default            #
#      * Fixed a bug which caused the map coordinates to be saved only as         #
#        integers, causing a program crash with user-selected map coordinates     #
#      * Fixed a bug which caused a program crash when opening a project file     #
#        when another was already opened                                          #
#      * Fixed a bug in the generation of the output file for MartiTracks that    #
#        caused the program to treat the names of the same species as different   #
#        when they included a dot                                                 #
#      * Fixed a bug that caused the generalized tracks to be displayed in a      #
#        single color                                                             #
#   Version 2.30, 26 Jun 2018 - Fourth public release ("Altai")                   #
#      * Changed the Background submenu to a mutually exclusive list of options   #
#      * Added an option to the Background submenu to return the display to the   #
#        default basemap                                                          #
#   Version 2.40, 3 Nov 2020 - Fifth public release ("Caribbean")                 #
#      * Removed SciPy dependence                                                 #
#      * Updated code to the new PyShp API                                        #
#   Version 2.50, 5 Oct 2026 - Sixth public release ("Flores-Soembawa")           #
#      * Updated to Python3, PyQt6, NumPy2, Matplotlib3, Basemap2                 # 
#      * The program is now a fully 64-bit application; support for 32-bit        #
#        operating systems has been dropped                                       #
#=================================================================================#

import sys
import os
import os.path
import time
import random
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
from matplotlib.colors import cnames, LinearSegmentedColormap

warnings.filterwarnings("ignore")

VERSION = "2.5.0 <i>Flores-Soembawa</i>"

symbols = {"o":0,"^":1,"v":2,"<":3,">":4,"s":5,"+":6,"x":7,"D":8,"d":9,"h":10,"p":11,"|":12,"_":13}

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

def distance_matrix(pts):
    # Returns matrix of pairwise Euclidean distances, vectorized numpy version
    return np.sum((pts[None,:] - pts[:, None])**2, -1)**0.5

def minimum_spanning_tree(X, copy_X=True):
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
            if self.model.item(i).checkState():
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
        self.dirname = getExecPath() + '\\'
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
                        shortcut="Ctrl+A", statusTip="Add file to project",
                        triggered=self.addFile)
        self.fileMenu.addAction(self.addAction)

        self.delAction = QtGui.QAction(getIcon("remove.png"), "&Remove File...", self,
                        shortcut="Ctrl+R", statusTip="Remove file from project",
                        triggered=self.delFile)
        self.fileMenu.addAction(self.delAction)

        self.fileMenu.addSeparator()
        self.importAction = QtGui.QAction(getIcon("import.png"), "Import...", self,
                        statusTip="Import data",
                        triggered=self.importFile)
        self.fileMenu.addAction(self.importAction)

        self.exportAction = QtGui.QAction(getIcon("export.png"), "Export...", self,
                        statusTip="Export data",
                        triggered=self.exportFile)
        self.fileMenu.addAction(self.exportAction)

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
        self.trackAction = QtGui.QAction(getIcon("track.png"), "&Individual track...", self,
                        statusTip="Compute individual track",
                        triggered=self.track)
        self.runMenu.addAction(self.trackAction)

        self.gentrackAction = QtGui.QAction(getIcon("track.png"), "&Generalized tracks...", self,
                        statusTip="Compute generalized tracks",
                        triggered=self.gentrack)
        self.runMenu.addAction(self.gentrackAction)

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
        toolbar.addAction(self.gentrackAction)
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

        recordsAction = QtGui.QAction("View records", self,
                        statusTip="Display all records",
                        triggered=self.displayRecords)
        menu.addAction(recordsAction)
        recordsAction.setEnabled(self.listWidget.count() > 0)

        treesAction = QtGui.QAction("View individual tracks", self,
                        statusTip="Display individual tracks",
                        triggered=self.displayTrees)
        menu.addAction(treesAction)
        treesAction.setEnabled(self.listWidget.count() > 0)

        tracksAction = QtGui.QAction("View generalized tracks", self,
                        statusTip="Display generalized tracks",
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
                file_ext = filename[-3:]
                if file_ext in ["csv", "xls", "lsx", "ods", "dbf", "kml"]:
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
                elif file_ext == "shp":
                    shape = self.shape_list[filename]
                    width = shape[4].get_linewidth()[0]
                    color = colors.to_hex(shape[4].get_color()[0])
                    self.shape_list[filename] = self.map.readshapefile(os.path.splitext(filename)[0], filename, linewidth=float(width), color=color)
                    if checked:
                        shape = self.shape_list[filename]
                        shape[4].set_visible(checked)

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
        self.reloadAction.setEnabled(enable)
        self.trackAction.setEnabled(enable)
        self.gentrackAction.setEnabled(enable)

    def loadData(self, filename):
        file_ext = filename[-3:]
        data = []
        if not os.path.exists(filename):
            QtWidgets.QMessageBox.critical(self, "Error", "File " + filename + " not found!")
        else:
            if file_ext == "xls" or file_ext == "lsx":
                wb = xlrd.open_workbook(filename)
                sh1 = wb.sheet_by_index(0)
                for rownum in range(sh1.nrows):
                    data += [sh1.row_values(rownum)]
                headers = data[0]
                del data[0]
            if file_ext == "ods":
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
                    dialect = csv.Sniffer().sniff(csvfile.readline(), [',',';','\t'])
                    csvfile.seek(0)
                    reader = csv.reader(csvfile, dialect)
                    headers = next(reader)
                    parser = ((row[0], row[1], float(row[2]), float(row[3])) for row in reader)
                    for row in parser:
                        data += [row]
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
                    description = description.encode("ascii", "ignore").replace('\n', ' ')
                    for ls in pm.iterfind('{0}Point/{0}coordinates'.format(nmsp)):
                        longitude = float(ls.text.split(',')[0])
                        latitude = float(ls.text.split(',')[1])
                    data.append((name, description, longitude, latitude))
            for i in range(len(headers)):
                headers[i] = headers[i].title().strip('\r')
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
            for i in range(self.listWidget.count()):
                item = self.listWidget.item(i)
                filename = str(item.text())
                if os.path.exists(filename):
                    file_ext = filename[-3:]
                    if file_ext in ["csv", "xls", "lsx", "ods", "dbf", "kml"]:
                        plot = self.plot_list[filename]
                        marker = plot[0].get_marker()
                        size = plot[0].get_markersize()
                        color = plot[0].get_markerfacecolor()
                        state = iif(item.checkState() == QtCore.Qt.CheckState.Checked, 1, 0)
                        outstr = ',' + marker + ',' + str(int(size)) + ',' + color.lower() + ',' + str(state)
                    elif file_ext == "shp":
                        shape = self.shape_list[filename]
                        style = '-'
                        width = shape[4].get_linewidth()[0]
                        color = colors.to_hex(shape[4].get_color()[0])
                        state = iif(item.checkState() == QtCore.Qt.CheckState.Checked, 1, 0)
                        outstr = ',' + style + ',' + str(int(width)) + ',' + color.lower() + ',' + str(state)
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
                    file_ext = filename[-3:]
                    if file_ext in ["csv", "xls", "lsx", "ods", "dbf", "kml"]:
                        headers, data = self.loadData(filename)
                        if len(headers) == 0:
                            QtWidgets.QMessageBox.critical(self, "Error", "Missing header line in data file!")
                            return
                        elif len(data) == 0:
                            QtWidgets.QMessageBox.critical(self, "Error", "Data file empty!")
                            return
                        coords = []
                        for row in data:
                            longitude = row[2]
                            latitude = row[3]
                            coords.append((longitude, latitude))
                        coords = np.array(coords, dtype=np.float32)
                        self.plot_list[filename] = self.map.ax.plot(coords[:,0], coords[:,1], symbol, markersize=int(size), color=color, picker=5)
                        plot = self.plot_list[filename]
                        plot[0].set_visible(checked)
                    elif file_ext == "shp":
                        self.shape_list[filename] = self.map.readshapefile(os.path.splitext(filename)[0], filename, linewidth=float(size), color=color)
                        shape = self.shape_list[filename]
                        shape[4].set_visible(checked)
                    filecount += 1
                else:
                    QtWidgets.QMessageBox.warning(self, "Warning", "File " + filename + " not found!")
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
                "Comma-Separated Values (*.csv);; Spreadsheets (*.xls *.xlsx *.ods);; Dbase (*.dbf);; Keyhole Markup Language (*.kml);; Shapefile (*.shp)")
        if filename:
            filename = os.path.basename(str(filename))
            if not filename in self.filelist:
                attributes = []
                item = QtWidgets.QListWidgetItem(filename)
                item.setCheckState(QtCore.Qt.CheckState.Checked)
                self.listWidget.addItem(item)
                file_ext = filename[-3:]
                if file_ext in ["csv", "xls", "lsx", "ods", "dbf", "kml"]:
                    headers, data = self.loadData(filename)
                    if len(headers) == 0:
                        QtWidgets.QMessageBox.critical(self, "Error", "Missing header line in data file!")
                        return
                    elif len(data) == 0:
                        QtWidgets.QMessageBox.critical(self, "Error", "Data file empty!")
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
                    self.plot_list[filename] = self.map.ax.plot(coords[:,0], coords[:,1], symbol, markersize=int(size), color=color, picker=5)
                elif file_ext == "shp":
                    style = '-'
                    width = 2
                    color = random_color()
                    attributes.append([style, width, color])
                    self.shape_list[filename] = self.map.readshapefile(os.path.splitext(filename)[0], filename, linewidth=float(width), color=color)
                self.filelist[filename] = attributes
                self.saveProject()
                self.updateUI()
                self.canvas.draw_idle()

    def delFile(self):
        if self.listWidget.count() > 0 and self.listWidget.selectedIndexes():    
            index = self.listWidget.selectedIndexes()[0].row()
            item = self.listWidget.takeItem(index)
            filename = str(item.text())
            file_ext = filename[-3:]
            del item
            if file_ext in ["csv", "xls", "lsx", "ods", "dbf", "kml"]:
                plot = self.plot_list[filename]
                plot[0].set_visible(False)
            elif file_ext == "shp":
                shape = self.shape_list[filename]
                shape[4].set_visible(False)
            self.filelist.pop(filename)
            self.saveProject()
            self.updateUI()
            self.canvas.draw_idle()

    def importFile(self):
        if self.listWidget.count() == 0:
            self.newFile()
            if not self.filename:
                return

        output_formats = ["csv", "xls", "xlsx", "ods"]
        input_formats = ["gbif", "idigbio", "obis", "vertnet"]
        options = [("Input format:",
                    [0, "GBIF", "iDigBio", "OBIS", "VertNet"]),
                    ("Output format:",
                    [0, "CSV Text (.csv)",
                        "Microsoft Excel 97/2000/XP/2003 (.xls)",
                        "Microsoft Excel 2007/2010 (.xlsx)",
                        "ODF Spreadsheet (.ods)"]),
                  ]
        result = fedit(options, title="Import", comment="Select output format:")
        if result is not None:
            infmt = input_formats[result[0]]
            outfmt = output_formats[result[1]]
            datasource, _ = QtWidgets.QFileDialog.getOpenFileName(self,
                "Open File",
                "",
                "Comma-Separated Values (*.csv);; Text (*.txt)")
            if datasource:
                filename, records = self.importData(datasource, infmt, outfmt)
                if records == 0:
                    QtWidgets.QMessageBox.warning(self, "Warning", "No records imported!")
                    return
                else:
                    QtWidgets.QMessageBox.information(self, "Information", "%d records imported." % records)

                if filename in self.filelist:
                    self.filelist.pop(filename)
                    items = self.listWidget.findItems(filename, QtCore.Qt.MatchFlag.MatchExactly)
                    if len(items) > 0:
                        for item in items:
                            del item

                attributes = []
                item = QtWidgets.QListWidgetItem(filename)
                item.setCheckState(QtCore.Qt.CheckState.Checked)
                self.listWidget.addItem(item)
                headers, data = self.loadData(filename)
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
                self.plot_list[filename] = self.map.ax.plot(coords[:,0], coords[:,1], symbol, markersize=int(size), color=color, picker=5)
                self.filelist[filename] = attributes
                self.saveProject()
                self.updateUI()
                self.canvas.draw_idle()

    def exportFile(self):
        output_formats = ["shp", "kml"]
        items = ("Shapefile (.shp)", "Keyhole Markup Language (.kml)")
        item, ok = QtWidgets.QInputDialog.getItem(self, "Export", "Output Formats:", items, 0, False)
        if ok and item:
            outfmt = output_formats[items.index(str(item))]
            datafiles = []
            for filename in self.filelist:
                file_ext = filename[-3:]
                if file_ext in ["csv", "xls", "lsx", "ods", "dbf"] + iif(outfmt == "shp", ["kml"], []):
                    datafiles.append(filename)
            datafiles.sort()
            total = 0
            form = ChecklistDialog("Data Files", datafiles, checked=True)
            if form.exec():
                QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
                for s in form.choices:
                    headers, data = self.loadData(str(s))
                    if len(headers) == 0:
                        QtWidgets.QMessageBox.critical(self, "Error", "Missing header line in data file!")
                        return
                    elif len(data) == 0:
                        QtWidgets.QMessageBox.critical(self, "Error", "Data file empty!")
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

                QtWidgets.QApplication.restoreOverrideCursor()
                QtWidgets.QMessageBox.information(self, "Information",
                    "%d file(s) exported." % (total))

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
            event.accept()
            self.saveProject()
        else:
            event.ignore()

    def track(self):
        datafiles = []
        for filename in self.filelist:
            file_ext = filename[-3:]
            if file_ext in ["csv", "xls", "lsx", "ods", "dbf", "kml"]:
                if get_basename(str(filename)) + "_individualtrack.shp" not in self.filelist:
                    datafiles.append(filename)
        if len(datafiles) == 0:
            QtWidgets.QMessageBox.warning(self, "Warning", "All tracks have already been computed!")
            return False
        datafiles.sort()
        form = ChecklistDialog("Data Files", datafiles, checked=True)
        if form.exec():
            if len(form.choices) < 1:
                QtWidgets.QMessageBox.warning(self, "Warning", "At least one data file must be selected!")
                return False
            for s in form.choices:
                attributes = []
                headers, data = self.loadData(str(s))
                if len(headers) == 0:
                    QtWidgets.QMessageBox.critical(self, "Error", "Missing header line in data file!")
                    return
                elif len(data) == 0:
                    QtWidgets.QMessageBox.critical(self, "Error", "Data file empty!")
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
                    attributes = ['-', width, color]
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

    def gentrack(self):
        datafiles = []
        for filename in self.filelist:
            file_ext = filename[-3:]
            if file_ext in ["csv", "xls", "lsx", "ods", "dbf", "kml"]:
                datafiles.append(filename)
        datafiles.sort()
        form = ChecklistDialog("Data Files", datafiles, checked=True)
        if form.exec():
            if len(form.choices) < 2:
                QtWidgets.QMessageBox.warning(self, "Warning", "At least two data files must be selected!")
                return False

            if os.path.exists("output0.dat"):
                os.remove("output0.dat")

            for s in form.choices:
                headers, indata = self.loadData(str(s))
                if len(headers) == 0:
                    QtWidgets.QMessageBox.critical(self, "Error", "Missing header line in data file!")
                    return
                elif len(indata) == 0:
                    QtWidgets.QMessageBox.critical(self, "Error", "Data file empty!")
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
                outfile.write(iif(result[5] == 0, "croizat0", "croizat1"))
                outfile.close()

                # Execute MartiTracks
                if os.path.exists("output0.kml"):
                    os.remove("output0.kml")

                self.statusBar().showMessage("Computing generalized tracks")
                QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
                mt05 = subprocess.call([self.dirname + "mt05-win64.exe", "output0.dat", "output0.kml", "params.dat"], shell=True)
                QtWidgets.QApplication.restoreOverrideCursor()
                self.statusBar().showMessage("Ready")
                if mt05 != 0:
                    QtWidgets.QMessageBox.critical(self, "Error", "Cannot run MartiTracks program!")
                    return

                # Create shapefile
                if result[6]:
                    doc = et.parse("output0.kml")
                    nmsp = '{http://earth.google.com/kml/2.1}'

                    for pm in doc.iterfind('.//{0}Placemark'.format(nmsp)):
                        name = pm.find('{0}name'.format(nmsp)).text
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
                        self.filelist[name + ".shp"] = [style, width, color]
                        item = QtWidgets.QListWidgetItem(name + ".shp")
                        item.setCheckState(QtCore.Qt.CheckState.Checked)
                        self.listWidget.addItem(item)
                        self.saveProject()
                        self.canvas.draw_idle()
                    self.statusBar().showMessage("Ready")

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
            self.resolution = items[result[0]][0]
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
            file_ext = filename[-3:]
            if file_ext in ["csv", "xls", "lsx", "ods", "dbf", "kml"]:
                headers, data = self.loadData(filename)
                recs = len(data)
                plot = self.plot_list[filename]
                pos = symbols[plot[0].get_marker()]
                size = int(plot[0].get_markersize())
                color = plot[0].get_markerfacecolor()
                point_options = [
                    (None, "File: " + filename),
                    (None, "Records: " + str(recs)),
                    (None, None),
                    ("Symbol:", [pos, "o (circle)",
                            "^ (triangle up)",
                            "v (triangle down)",
                            "< (triangle left)",
                            "> (triangle right)",
                            "s (square)",
                            "+ (plus)",
                            "x (cross)",
                            "D (diamond)",
                            "d (thin diamond)",
                            "h (hexagon)",
                            "p (pentagon)",
                            "I (vertical line)",
                            "_ (horizontal line)"]),
                    ("Size:", size),
                    ("Color:", color),
                ]
                result = fedit(point_options, title="Point Properties", comment="Set point properties.")
                if result is not None:
                    symbol = list(symbols.keys())[list(symbols.values()).index(result[0])]
                    size = result[1]
                    color = result[2]
                    plot[0].set_marker(symbol)
                    plot[0].set_markersize(size)
                    plot[0].set_markerfacecolor(color)
                    plot[0].set_markeredgecolor(color)

            elif file_ext == "shp":
                shape = self.shape_list[filename]
                width = int(shape[4].get_linewidth()[0])
                color = colors.to_hex(shape[4].get_color()[0])
                line_options = [
                    (None, "File: " + filename),
                    (None, None),
                    ("Width:", width),
                    ("Color:", color),
                ]
                result = fedit(line_options, title="Line Properties", comment="Set line properties.")
                if result is not None:
                    color = result[1]
                    shape[4].set_linewidth(result[0])
                    shape[4].set_color(color)
            self.canvas.draw_idle()

    def listWidgetItemChanged(self, item):
        filename = str(item.text())
        checked = item.checkState() == QtCore.Qt.CheckState.Checked
        file_ext = filename[-3:]
        if file_ext in ["csv", "xls", "lsx", "ods", "dbf", "kml"]:
            plot = self.plot_list[filename]
            plot[0].set_visible(checked)
        elif file_ext == "shp":
            shape = self.shape_list[filename]
            shape[4].set_visible(checked)
        self.canvas.draw_idle()

    def pick(self, event):
        if self.listWidget.count() == 0: return
        thisline = event.artist
        xdata = thisline.get_xdata()
        ydata = thisline.get_ydata()
        ind = event.ind
        points = tuple(zip(xdata[ind], ydata[ind]))
        x = round(points[0][0],4)
        y = round(points[0][1],4)
        all_data = []
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CursorShape.WaitCursor)
        for i in range(self.listWidget.count()):
            item = self.listWidget.item(i)
            filename = str(item.text())
            file_ext = filename[-3:]
            if file_ext in ["csv", "xls", "lsx", "ods", "dbf", "kml"]:
                headers, data = self.loadData(filename)
                all_data += data
        QtWidgets.QApplication.restoreOverrideCursor()
        match = [coord for coord in all_data if round(float(coord[2]),4) == x and round(float(coord[3]),4) == y]
        if not match:
            return
        recinfo = "Scientific Name: %s\n Location: %s\n Latitude: %f\n Longitude: %f" % (match[0][0], match[0][1], match[0][2], match[0][3])
        QtWidgets.QMessageBox.information(self, "Record Information", recinfo)

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