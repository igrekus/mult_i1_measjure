from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot, pyqtSignal, QRunnable, QThreadPool
from PyQt5.QtWidgets import QWidget


class MeasureTask(QRunnable):

    def __init__(self, fn, end, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.end = end
        self.args = args
        self.kwargs = kwargs

    def run(self):
        self.fn(*self.args, **self.kwargs)
        self.end()


class MeasureWidget(QWidget):

    def __init__(self, parent=None, controller=None):
        super().__init__(parent=parent)

        self._ui = uic.loadUi('measurewidget.ui', self)
        self._controller = controller
        self._threads = QThreadPool()

    def check(self):
        print('connect')
        self._modeDuringCheck()
        self._threads.start(MeasureTask(self._controller.check,
                                        self.checkTaskComplete,
                                        {'params': 'parampampams'}))

    def checkTaskComplete(self):
        print('check complete')
        if not self._controller.present:
            print('sample not found')
            return

        print('found sample')
        self._modePreMeasure()

    def measure(self):
        print('measuring...')

    @pyqtSlot()
    def on_instrumentsConnected(self):
        self._modePreCheck()

    @pyqtSlot()
    def on_btnCheck_clicked(self):
        print('checking sample presence')
        self.check()

    @pyqtSlot()
    def on_btnMeasure_clicked(self):
        print('start measure')

        self.measure()

        print('measure complete')
        self._modePreCheck()

    def _modePreConnect(self):
        self._ui.btnCheck.setEnabled(False)
        self._ui.btnMeasure.setEnabled(False)
        # TODO also update param widgets

    def _modePreCheck(self):
        self._ui.btnCheck.setEnabled(True)
        self._ui.btnMeasure.setEnabled(False)

    def _modeDuringCheck(self):
        self._ui.btnCheck.setEnabled(False)
        self._ui.btnMeasure.setEnabled(False)

    def _modePreMeasure(self):
        self._ui.btnCheck.setEnabled(False)
        self._ui.btnMeasure.setEnabled(True)
