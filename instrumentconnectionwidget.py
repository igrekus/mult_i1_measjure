from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot, pyqtSignal
from PyQt5.QtWidgets import QWidget

from instrumentwidget import InstrumentWidget


class InstrumentConnectionWidget(QWidget):

    connected = pyqtSignal()

    def __init__(self, parent=None, controller=None):
        super().__init__(parent=parent)

        self._ui = uic.loadUi('instrumentconnectionwidget.ui', self)
        self._controller = controller

        self._widgets = {
            k: InstrumentWidget(parent=self, title=f'{k}', addr=f'{v.addr}')
            for k, v in self._controller.requiredInstruments.items()
        }

        self._setupUi()

    def _setupUi(self):
        for i, iw in enumerate(self._widgets.items()):
            self._ui.layInstruments.insertWidget(i, iw[1])

    @pyqtSlot()
    def on_btnConnect_clicked(self):
        print('connect')

        if not self._controller.connect({k: w.address for k, w in self._widgets.items()}):
            print('connect error check connection')
            return

        for w, s in zip(self._widgets.values(), self._controller.status):
            w.status = s
        self.connected.emit()
