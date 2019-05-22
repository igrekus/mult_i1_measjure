from PyQt5.QtCore import pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QRadioButton, QButtonGroup


class DeviceSelectWidget(QWidget):

    selectedChanged = pyqtSignal(str)

    def __init__(self, parent=None, params=None):
        super().__init__(parent=parent)

        self._layout = QVBoxLayout()
        self._group = QButtonGroup()

        for i, label in enumerate(params.keys()):
            w = QRadioButton(label)
            self._group.addButton(w, i)
            self._layout.addWidget(w)

        self.setLayout(self._layout)
        self._group.button(0).setChecked(True)
        self._group.buttonToggled[int, bool].connect(self.on_buttonToggled)

        self._enabled = True

    @property
    def selected(self):
        return self._group.checkedButton().text()

    @pyqtSlot(int, bool)
    def on_buttonToggled(self, id, toggled):
        if toggled:
            self.selectedChanged.emit(self._group.button(id).text())

    @property
    def enabled(self):
        return self._enabled

    @enabled.setter
    def enabled(self, value: bool):
        self._enabled = value
        for b in self._group.buttons():
            b.setEnabled(self._enabled)
