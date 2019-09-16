from PyQt5.QtCore import Qt, QAbstractTableModel, QVariant, pyqtSlot


class ControlModel(QAbstractTableModel):
    def __init__(self, parent=None, controller=None):
        super().__init__(parent)

        self._controller = controller
        self._params = dict()

        self._data = list()
        self._headers = ['№', 'Частота, ГГц']

        # self._init()

    def _init(self, value):
        self._params = self._controller.deviceParams[value]

        self._data = [[i + 1, f] for i, f in enumerate(self._params['F'])]

    def headerData(self, section, orientation, role=None):
        if orientation == Qt.Horizontal:
            if role == Qt.DisplayRole:
                if section < len(self._headers):
                    return QVariant(self._headers[section])
        return QVariant()

    def rowCount(self, parent=None, *args, **kwargs):
        if parent.isValid():
            return 0
        return len(self._data)

    def columnCount(self, parent=None, *args, **kwargs):
        return len(self._headers)

    def data(self, index, role=None):
        if not index.isValid():
            return QVariant()
        if role == Qt.DisplayRole:
            return QVariant(self._data[index.row()][index.column()])
        return QVariant()

    def getParamsForRow(self, row):
        print(f'getting params for row {row}')
        return {
            'F': self._params['F'][row],
            'Freal': self._params['Freal'][row],
            'Fmul': self._params['Fmul'][row],
            'Poffs1': self._params['Poffs1'][row],
            'Poffs2': self._params['Poffs2'][row],
            'Poffs3': self._params['Poffs3'][row],
            'p': self._params['P1']
        }

    @pyqtSlot(str)
    def on_deviceChanged(self, value):
        self.beginResetModel()
        self._init(value)
        self.endResetModel()
