from PyQt5.QtCore import Qt, QAbstractTableModel, QVariant, pyqtSlot, QModelIndex


class ControlModel(QAbstractTableModel):
    def __init__(self, parent=None, controller=None):
        super().__init__(parent)

        self._controller = controller
        self._params = dict()

        self._data = list()
        self._headers = ['№', 'Частота, ГГц', 'x1', 'x2', 'x3']

        # self._init()

    def _init(self, value, second=0):
        self.beginResetModel()

        self._params = self._controller.deviceParams[value]

        if len(self._params['F']) != 3:
            self._data = []
            return

        self._data = [[i + 1, f, 1, 2, 3] for i, f in enumerate(self._params['F'][second])]
        self.endResetModel()

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

    def getParamsForRow(self, row, second):
        print(f'getting params for row {row}')
        return {
            'F': self._params['F'][second][row],
            'Freal': self._params['Freal'][second][row],
            'Fmul': self._params['Fmul'][second][row],
            'Poffs1': self._params['Poffs1'][second][row],
            'Poffs2': self._params['Poffs2'][second][row],
            'Poffs3': self._params['Poffs3'][second][row],
            'P1': self._params['P1'],
            'P2': self._params['P2'],
            'Istat': self._params['Istat'],
            'Idyn': self._params['Idyn']
        }

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        col = index.column()
        flags = super().flags(index)
        if col == 0 or col == 1:
            return flags ^ Qt.ItemIsSelectable
        return flags


    @pyqtSlot(str)
    def on_deviceChanged(self, value):
        self._init(value)

    @pyqtSlot(str, int)
    def on_secondaryChanged(self, first, second):
        self._init(first, second)
