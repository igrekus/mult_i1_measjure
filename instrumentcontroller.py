import random
import time
import pandas
import visa

from os import listdir
from os.path import isfile, join
from collections import defaultdict

from PyQt5.QtCore import QObject

# from analyzer import Analyzer
# from analyzermock import AnalyzerMock
# from generator import Generator
# from generatormock import GeneratorMock
# from source import Source
# from sourcemock import SourceMock

# MOCK
from agilent34410amock import Agilent34410AMock
from agilente3644amock import AgilentE3644AMock
from agilentn5183amock import AgilentN5183AMock
from agilentn9030amock import AgilentN9030AMock
from instr.agilent34410a import Agilent34410A
from instr.agilente3644a import AgilentE3644A
from instr.agilentn5183a import AgilentN5183A
from instr.agilentn9030a import AgilentN9030A

mock_enabled = True


class Src:
    addr = 'src addr'
    model = 'Src'
    status = f'{model} at {addr}'
    def __str__(self):
        return 'Src'


class InstrumentFactory:
    def __init__(self, addr, label):
        self.applicable = None
        self.addr = addr
        self.label = label
    def find(self):
        # TODO remove applicable instrument when found one if needed more than one instrument of the same type
        # TODO: idea: pass list of applicable instruments to differ from the model of the same type?
        instr = self.from_address()
        if not instr:
            return self.try_find()
        return instr
    def from_address(self):
        raise NotImplementedError()
    def try_find(self):
        raise NotImplementedError()


class GeneratorFactory(InstrumentFactory):
    def __init__(self, addr):
        super().__init__(addr=addr, label='Генератор')
        self.applicable = ['N5183A']
    def from_address(self):
        if mock_enabled:
            return AgilentN5183A(self.addr, '1,N5183A mock,1', AgilentN5183AMock())
        try:
            rm = visa.ResourceManager()
            inst = rm.open_resource(self.addr)
            idn = inst.query('*IDN?')
            name = idn.split(',')[1].strip()
            if name in self.applicable:
                return AgilentN5183A(self.addr, idn, inst)
        except Exception as ex:
            print('Generator find error:', ex)
            exit(1)


class AnalyzerFactory(InstrumentFactory):
    def __init__(self, addr):
        super().__init__(addr=addr, label='Анализатор')
        self.applicable = ['N9030A']
    def from_address(self):
        if mock_enabled:
            return AgilentN9030A(self.addr, '1,N9030A mock,1', AgilentN9030AMock())
        try:
            rm = visa.ResourceManager()
            inst = rm.open_resource(self.addr)
            idn = inst.query('*IDN?')
            name = idn.split(',')[1].strip()
            if name in self.applicable:
                return AgilentN9030A(self.addr, idn, inst)
        except Exception as ex:
            print('Analyzer find error:', ex)
            exit(2)


class MultimeterFactory(InstrumentFactory):
    def __init__(self, addr):
        super().__init__(addr=addr, label='Мультиметр')
        self.applicable = ['34410A']
    def from_address(self):
        if mock_enabled:
            return Agilent34410A(self.addr, '1,34410A mock,1', Agilent34410AMock())
        try:
            rm = visa.ResourceManager()
            inst = rm.open_resource(self.addr)
            idn = inst.query('*IDN?')
            name = idn.split(',')[1].strip()
            if name in self.applicable:
                return Agilent34410A(self.addr, idn, inst)
        except Exception as ex:
            print('Multimeter find error:', ex)
            exit(3)


class SourceFactory(InstrumentFactory):
    def __init__(self, addr):
        super().__init__(addr=addr, label='Исчточник питания')
        self.applicable = ['SRC MODEL']
    def from_address(self):
        if mock_enabled:
            return AgilentE3644A(self.addr, '1,SRC MODEL mock,1', AgilentE3644AMock())
        try:
            rm = visa.ResourceManager()
            inst = rm.open_resource(self.addr)
            idn = inst.query('*IDN?')
            name = idn.split(',')[1].strip()
            if name in self.applicable:
                return AgilentE3644A(self.addr, idn, inst)
        except Exception as ex:
            print('Source find error:', ex)
            exit(4)


class MeasureResult:
    def __init__(self):
        self.headers = list()
    def init(self):
        raise NotImplementedError()
    def process_raw_data(self, *args, **kwargs):
        raise NotImplementedError()


class MeasureResultMock(MeasureResult):
    def __init__(self, device, secondary):
        super().__init__()
        self.devices: list = list(device.keys())
        self.secondary: dict = secondary

        self.headersCache = dict()
        self._generators = defaultdict(list)
        self.data = list()

    def init(self):
        self.headersCache.clear()
        self._generators.clear()
        self.data.clear()

        # check task table presence
        def getFileList(data_path):
            return [l for l in listdir(data_path) if isfile(join(data_path, l)) and '.xlsx' in l]

        files = getFileList('.')
        if len(files) != 1:
            print('working dir should have only one task table')
            return False

        self._parseTaskTable(files[0])
        return True

    def _parseTaskTable(self, filename):
        print(f'using task table: {filename}')
        for dev in self.devices:
            raw_data: pandas.DataFrame = pandas.read_excel(filename, sheet_name=dev)
            name, _, *headers = raw_data.columns.tolist()
            self.headersCache[name] = headers
            for g in raw_data.groupby(name):
                _, df = g
                for h in headers:
                    self._generators[f'{name} {df[name].tolist()[0]}'].append(df[h].tolist())

    def process_raw_data(self, device, secondary, raw_data):
        print('processing', device, secondary, raw_data)
        self.headers = self.headersCache[device]
        self.data = [self.generateValue(data) for data in self._generators[f'{device} {secondary}']]

    def generateValue(self, data):

        if not data or '-' in data or chr(0x2212) in data or not all(data):
            return '-'

        span, step, mean = data
        start = mean - span
        stop = mean + span
        return random.randint(0, int((stop - start) / step)) * step + start


class InstrumentController(QObject):

    def __init__(self, parent=None):
        super().__init__(parent=parent)

        self.requiredInstruments = {
            'Источник питания': SourceFactory('GPIB0::10::INSTR'),
            'Мультиметр': MultimeterFactory('GPIB0::11::INSTR'),
            'Генератор': GeneratorFactory('GPIB0::12::INSTR'),
            'Анализатор': AnalyzerFactory('GPIB0::5::INSTR'),
        }

        # TODO generate parameter list from .xlsx
        self.deviceParams = {
            'Тип 1 (1324ПП11У)': {
                'F': [0.6, 0.75, 0.89, 1.04, 1.18, 1.33, 1.47, 1.62, 1.76, 1.9, 2.20],
                'mul': 2,
                'P1': 13,
                'P2': 21,
                'Istat': [None, None, None],
                'Idyn': [None, None, None]
            },
            'Тип 5 (1324ПП15У)': {
                'F': [0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5, 0.55, 0.6],
                'mul': 3,
                'P1': 13,
                'P2': 21,
                'Istat': [None, None, None],
                'Idyn': [None, None, None]
            },
            'Тип 12 (1324ПП19У)': {
                'F': [1.2, 1.35, 1.5, 1.65, 1.8, 1.95, 2.1, 2.25, 2.4, 2.55, 2.8],
                'mul': 4,
                'P1': 0,
                'P2': 7,
                'Istat': [(160, 170), (123, 125), (220, 230)],
                'Idyn': [(130, 140), (110, 120), (200, 215)]
            },
            'Тип 1а (1324ПП23У)': {
                'F': [0.6, 0.73, 0.86, 0.99, 1.12, 1.25, 1.38, 1.51, 1.64, 1.77, 2.0],
                'mul': 2,
                'P1': 13,
                'P2': 21,
                'Istat': [(220, 230), (170, 190), (240, 250)],
                'Idyn': [(200, 210), (155, 165), (220, 230)]
            },
        }

        # TODO generate combo for secondary params
        self.secondaryParams = {
            0: 0,
            1: 1,
            2: 2
        }

        self._instruments = {}
        self.found = False
        self.present = False
        self.hasResult = False

        self.result = MeasureResult() if not mock_enabled \
            else MeasureResultMock(self.deviceParams, self.secondaryParams)

    def __str__(self):
        return f'{self._instruments}'

    def connect(self, addrs):
        print(f'searching for {addrs}')
        for k, v in addrs.items():
            self.requiredInstruments[k].addr = v
        self.found = self._find()
        # time.sleep(5)

    def _find(self):
        self._instruments = {
            k: v.find() for k, v in self.requiredInstruments.items()
        }
        return all(self._instruments.values())

    def check(self, params):
        print(f'call check with {params}')
        device, secondary = params
        self.present = self._check(device, secondary)
        print('sample pass')

    def _check(self, device, secondary):
        print(f'launch check with {self.deviceParams[device]} {self.secondaryParams[secondary]}')
        return self.result.init() and self._runCheck(self.deviceParams[device], self.secondaryParams[secondary])

    def _runCheck(self, param, secondary):
        threshold = -5

        if param['Istat'][0] is not None:
            self._instruments['Источник питания'].set_current(chan=1, value=300, unit='mA')
            self._instruments['Источник питания'].set_voltage(chan=1, value=5, unit='V')
            self._instruments['Источник питания'].set_output(chan=1, state='ON')

        self._instruments['Генератор'].set_modulation(state='OFF')
        self._instruments['Генератор'].set_freq(value=param['F'][6], unit='GHz')
        self._instruments['Генератор'].set_pow(value=param['P1'], unit='dBm')
        self._instruments['Генератор'].set_output(state='ON')

        self._instruments['Анализатор'].set_autocalibrate(state='OFF')
        self._instruments['Анализатор'].set_span(value=1, unit='MHz')

        center_freq = param['F'][6] * param['mul']
        self._instruments['Анализатор'].set_measure_center_freq(value=center_freq, unit='GHz')
        self._instruments['Анализатор'].set_marker1_x_center(value=center_freq, unit='GHz')
        pow = self._instruments['Анализатор'].read_pow(marker=1)

        self._instruments['Анализатор'].remove_marker(marker=1)
        self._instruments['Анализатор'].set_autocalibrate(state='ON')
        self._instruments['Генератор'].set_output(state='OFF')
        self._instruments['Источник питания'].set_output(chan=1, state='OFF')

        print('smaple response:', pow)
        return pow > threshold

    def measure(self, params):
        print(f'call measure with {params}')
        device, secondary = params
        raw_data = self._ref_measure(device, secondary)
        self.hasResult = bool(raw_data)

        if self.hasResult:
            self.result.process_raw_data(device, secondary, raw_data)

    def _ref_measure(self, device, secondary):
        print(f'launch measure with {self.deviceParams[device]} {self.secondaryParams[secondary]}')
        # TODO implement actual measure algorithm
        # time.sleep(3)
        return ['raw data']

    @property
    def status(self):
        return [i.status for i in self._instruments.values()]

    def checkSample(self):
        print('instrument manager: check sample')

        if isfile('settings.ini'):
            with open('settings.ini') as f:
                line = f.readline()
                self.pow_limit = float(line.split('=')[1].strip())

        # TODO: report proper condition
        return True

    def checkTaskTable(self):

        def getFileList(data_path):
            return [l for l in listdir(data_path) if isfile(join(data_path, l)) and self.excel_extension in l]

        # TODO: extract measurement manager class
        print('instrument manager: check task table')

        files = getFileList('.')
        length = len(files)
        if length > 1:
            print('too many task tables, abort')
            return False
        elif length <= 0:
            print('no task table found, abort')
            return False

        self._fileName = files[0]
        print('using task table:', self._fileName)

        wb = openpyxl.load_workbook(self._fileName)
        ws = wb.active

        cols = ws.max_column
        rows = ws.max_row
        letters = [chr(ch_code).upper() for ch_code in range(ord('a'), ord('z') + 1)]
        start_col = 2

        count = int((rows - 1) / 3)

        # TODO: validate table
        for row in range(count):
            self._params[row + 1] = [[ws[char + str(row * 3 + 1 + num)].value for num in range(1, 4)] for char in letters[start_col:cols]]

        headers = ['№'] + [ws[char + '1'].value for char in letters[start_col:cols]]

        for model in self.measureModels.values():
            model.initHeader(headers)

        wb.close()
        print('done read')
        return True

    # TODO implement proper measurement algorithm
    def _measure(self, letter):
        print(f'start measure letter={letter}')

        self.measureTask(letter)

        self.measureModels[letter].initModel(self._params[letter])

        print('end measure')

    def generator_freq_sweep(self, generator, letter):

        def do_branch(letter, branch, step):

            self.control_voltage = self.step_voltage[step]

            if branch == 1:
                chan1voltage = self.control_voltage
                chan2voltage = 0.0
            else:
                chan1voltage = 0.0
                chan2voltage = self.control_voltage

            self._source.set_voltage(1, chan1voltage, 'V')
            self._source.set_voltage(2, chan2voltage, 'V')
            self._source.set_output(1, 'ON')
            self._source.set_output(2, 'ON')

            generator.set_output('ON')

            pows = list()

            for freq in self.measure_freq[letter].values():
                generator.set_freq(freq, self.measure_freq_unit)
                self._analyzer.set_measure_center_freq(freq, self.measure_freq_unit)
                self._analyzer.set_marker1_x_center(freq, self.measure_freq_unit)

                pows.append(self._analyzer.read_pow(1))
                if not mock_enabled:
                    time.sleep(0.3)

            generator.set_output('OFF')
            self._source.set_output(1, 'OFF')
            self._source.set_output(2, 'OFF')

            return pows

        generator.set_pow(self.measure_pow[1]['p1'], self.measure_pow_unit)

        pows = list()
        for cycle in range(3):
            pows = [do_branch(letter, branch, cycle) for branch in [1, 2]]
        return pows

    def measureTask(self, letter: int):

        print(f'measurement task run, letter={letter}')

        self._generator1.set_modulation('OFF')
        self._generator2.set_modulation('OFF')

        self._analyzer.set_autocalibrate('OFF')
        self._analyzer.set_span(10, 'MHz')

        self._source.set_current(1, self.measure_current, 'mA')
        self._source.set_current(2, self.measure_current, 'mA')

        self._analyzer.set_marker_mode(1, 'POS')

        # gen 1 freq sweep
        pows = self.generator_freq_sweep(self._generator1, letter)
        print(pows)

        # gen 2 freq sweep
        pows = self.generator_freq_sweep(self._generator2, letter)
        print(pows)

        # gen 1 pow sweep
        freq = self.measure_freq[letter]['f3']
        freq_unit = self.measure_freq_unit
        pow_unit = self.measure_pow_unit

        self._generator1.set_freq(freq, freq_unit)

        self._analyzer.set_autocalibrate('OFF')
        self._analyzer.set_span(1, 'MHz')
        self._analyzer.set_marker_mode(1, 'POS')
        self._analyzer.set_measure_center_freq(freq, 'MHz')
        self._analyzer.set_marker1_x_center(freq, 'MHz')

        self._source.set_voltage(1, 3, 'V')
        self._source.set_voltage(2, 0, 'V')
        self._source.set_output(1, 'ON')
        self._source.set_output(2, 'ON')

        self._generator1.set_output('ON')

        for pow in self.measure_pow[letter].values():
            self._generator1.set_pow(pow, pow_unit)

            if not mock_enabled:
                time.sleep(0.3)

            read_pow = self._analyzer.read_pow(1)
            print(read_pow)

        self._generator1.set_output('OFF')

        self._source.set_output(1, 'OFF')
        self._source.set_output(2, 'OFF')
        self._source.set_system_local()

        self._analyzer.set_autocalibrate('ON')
