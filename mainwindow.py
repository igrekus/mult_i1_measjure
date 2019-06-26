from PyQt5 import uic
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QDialog, QAction
from PyQt5.QtCore import Qt, QStateMachine, QState, pyqtSignal, pyqtSlot

from instrumentcontroller import InstrumentController
from connectionwidget import ConnectionWidget
from measuremodel import MeasureModel
from measurewidget import MeasureWidget, MeasureWidgetWithSecondaryParameters


class MainWindow(QMainWindow):

    instrumentsFound = pyqtSignal()
    sampleFound = pyqtSignal()
    measurementFinished = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)

        self.setAttribute(Qt.WA_QuitOnClose)
        self.setAttribute(Qt.WA_DeleteOnClose)

        # create instance variables
        self._ui = uic.loadUi('mainwindow.ui', self)
        self._instrumentController = InstrumentController(parent=self)
        self._connectionWidget = ConnectionWidget(parent=self, controller=self._instrumentController)
        self._measureWidget = MeasureWidgetWithSecondaryParameters(parent=self, controller=self._instrumentController)
        self._measureModel = MeasureModel(parent=self, controller=self._instrumentController)

        # init UI
        self._ui.layInstrs.insertWidget(0, self._connectionWidget)
        self._ui.layInstrs.insertWidget(1, self._measureWidget)

        self._init()

    def _init(self):
        self._connectionWidget.connected.connect(self.on_instrumens_connected)
        self._connectionWidget.connected.connect(self._measureWidget.on_instrumentsConnected)

        self._measureWidget.measureComplete.connect(self._measureModel.update)

        self._ui.tableMeasure.setModel(self._measureModel)

        self.refreshView()

    # UI utility methods
    def refreshView(self):
        self.resizeTable()
        # twidth = self.ui.tableSuggestions.frameGeometry().width() - 30
        # self.ui.tableSuggestions.setColumnWidth(0, twidth * 0.05)
        # self.ui.tableSuggestions.setColumnWidth(1, twidth * 0.10)
        # self.ui.tableSuggestions.setColumnWidth(2, twidth * 0.55)
        # self.ui.tableSuggestions.setColumnWidth(3, twidth * 0.10)
        # self.ui.tableSuggestions.setColumnWidth(4, twidth * 0.15)
        # self.ui.tableSuggestions.setColumnWidth(5, twidth * 0.05)

    def resizeTable(self):
        self._ui.tableMeasure.resizeRowsToContents()
        self._ui.tableMeasure.resizeColumnsToContents()

    # event handlers
    def resizeEvent(self, event):
        self.refreshView()

    @pyqtSlot()
    def on_instrumens_connected(self):
        print(f'connected {self._instrumentController}')
