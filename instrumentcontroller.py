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
        self.applicable = ['N5183A', 'N5181B']
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
        self.applicable = ['E3648A']
    def from_address(self):
        if mock_enabled:
            return AgilentE3644A(self.addr, '1,E3648A mock,1', AgilentE3644AMock())
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
        def get_file_list(data_path):
            return [l for l in listdir(data_path) if isfile(join(data_path, l)) and '.xlsx' in l]

        files = get_file_list('.')
        if len(files) != 1:
            print('working dir should have only one task table')
            return False

        self._parese_task_table(files[0])
        return True

    def _parese_task_table(self, filename):
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
        self.data = [self.generate_value(data) for data in self._generators[f'{device} {secondary}']]

    @staticmethod
    def generate_value(data):
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
            'Источник питания': SourceFactory('GPIB0::4::INSTR'),
            'Мультиметр': MultimeterFactory('GPIB0::2::INSTR'),
            'Генератор': GeneratorFactory('GPIB0::20::INSTR'),
            'Анализатор': AnalyzerFactory('GPIB0::18::INSTR'),
        }

        self.deviceParams = {
            'Тип 2 (1324ПП12AT)': {
                'F': [1.15, 1.35, 1.75, 1.92, 2.25, 2.54, 2.7, 3, 3.47, 3.86, 4.25],
                'mul': 2,
                'P1': 15,
                'P2': 21,
                'Istat': [None, None, None],
                'Idyn': [None, None, None]
            },
        }

        if isfile('./params.ini'):
            import ast
            with open('./params.ini', 'rt', encoding='utf-8') as f:
                raw = ''.join(f.readlines())
                self.deviceParams = ast.literal_eval(raw)

        # TODO generate combo for secondary params
        self.secondaryParams = {
            0: 0,
            1: 1,
            2: 2
        }

        self.span = 1

        self._instruments = {}
        self.found = False
        self.present = False
        self.hasResult = False

        # self.result = MeasureResult() if not mock_enabled \
        #     else MeasureResultMock(self.deviceParams, self.secondaryParams)
        self.result = MeasureResultMock(self.deviceParams, self.secondaryParams)

    def __str__(self):
        return f'{self._instruments}'

    def connect(self, addrs):
        print(f'searching for {addrs}')
        for k, v in addrs.items():
            self.requiredInstruments[k].addr = v
        self.found = self._find()

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
        threshold = -120.0
        if isfile('./settings.ini'):
            with open('./settings.ini', 'rt') as f:
                threshold = float(f.readline().strip().split('=')[1])

        if param['Istat'][0] is not None:
            self._instruments['Источник питания'].set_current(chan=1, value=300, unit='mA')
            self._instruments['Источник питания'].set_voltage(chan=1, value=5, unit='V')
            self._instruments['Источник питания'].set_output(chan=1, state='ON')

        self._instruments['Генератор'].set_modulation(state='OFF')
        self._instruments['Генератор'].set_freq(value=param['F'][6], unit='GHz')
        self._instruments['Генератор'].set_pow(value=param['P1'], unit='dBm')
        self._instruments['Генератор'].set_output(state=1)

        self._instruments['Анализатор'].set_autocalibrate(state='OFF')
        self._instruments['Анализатор'].set_span(value=self.span, unit='MHz')

        if not mock_enabled:
            time.sleep(0.2)

        center_freq = param['F'][6] * param['mul']
        self._instruments['Анализатор'].set_measure_center_freq(value=center_freq, unit='GHz')
        self._instruments['Анализатор'].set_marker1_x_center(value=center_freq, unit='GHz')
        pow = self._instruments['Анализатор'].read_pow(marker=1)

        self._instruments['Анализатор'].remove_marker(marker=1)
        self._instruments['Анализатор'].set_autocalibrate(state='ON')
        self._instruments['Генератор'].set_output(state='OFF')
        self._instruments['Источник питания'].set_output(chan=1, state='OFF')

        print('sфmaple response:', pow)
        return pow > threshold

    def measure(self, params):
        print(f'call measure with {params}')
        device, secondary = params
        raw_data = self._measure(device, secondary)
        self.hasResult = bool(raw_data)

        if self.hasResult:
            self.result.process_raw_data(device, secondary, raw_data)

    def _measure(self, device, secondary):
        param = self.deviceParams[device]
        secondary = self.secondaryParams[secondary]
        print(f'launch measure with {param} {secondary}')

        self._instruments['Генератор'].set_modulation(state='OFF')
        self._instruments['Анализатор'].set_autocalibrate(state='OFF')
        self._instruments['Анализатор'].set_span(value=self.span, unit='MHz')
        self._instruments['Анализатор'].set_marker_mode(marker=1, mode='POS')

        # TODO extract static measure func
        # ===
        if param['Istat'][0] is not None:
            self._instruments['Источник питания'].set_current(chan=1, value=300, unit='mA')
            self._instruments['Источник питания'].set_voltage(chan=1, value=5.55, unit='V')
            self._instruments['Источник питания'].set_output(chan=1, state='ON')

            self._instruments['Генератор'].set_freq(value=param['F'][6], unit='GHz')
            self._instruments['Генератор'].set_pow(value=param['P2'], unit='dBm')
            self._instruments['Генератор'].set_output(state='ON')

            curr = int(MeasureResultMock.generate_value(param['Istat'][secondary]) * 10)
            curr_str = ' 00.' + f'{curr}  ADC'.replace('.', ',')
            self._instruments['Мультиметр'].send(f'DISPlay:WIND1:TEXT "{curr_str}"')

            if not mock_enabled:
                time.sleep(3)

            self._instruments['Генератор'].set_output(state='OFF')
            self._instruments['Источник питания'].set_output(chan=1, state='OFF')
        # ===

        # TODO extract freq sweep func
        # ===
        if param['Istat'][0] is not None:
            self._instruments['Источник питания'].set_voltage(chan=1, value=4.45, unit='V')
            self._instruments['Источник питания'].set_output(chan=1, state='ON')

        self._instruments['Генератор'].set_freq(value=param['F'][0], unit='GHz')
        self._instruments['Генератор'].set_output(state='ON')

        pow_sweep_res = list()
        for freq in param['F']:
            temp = list()
            self._instruments['Генератор'].set_freq(value=freq, unit='GHz')

            for index, pow in enumerate([param['P1'], param['P2']]):
                self._instruments['Генератор'].set_pow(value=pow, unit='dBm')

                if param['Idyn'][0] is not None:
                    if index == 0:
                        curr = int(MeasureResultMock.generate_value(param['Idyn'][secondary]) * 10)
                    elif index == 1:
                        curr = int(MeasureResultMock.generate_value(param['Istat'][secondary]) * 10)

                    curr_str = ' 00.' + f'{curr}  ADC'.replace('.', ',')
                    self._instruments['Мультиметр'].send(f'DISPlay:WIND1:TEXT "{curr_str}"')

                for mul in range(1, 5):
                    harm = mul * freq

                    if harm < 26:
                        self._instruments['Анализатор'].set_measure_center_freq(value=harm, unit='GHz')
                        self._instruments['Анализатор'].set_marker1_x_center(value=harm, unit='GHz')
                        if not mock_enabled:
                            time.sleep(0.20)
                    else:
                        adjusted_harm = freq * 2 if freq * 2 < 26 else freq * 1
                        self._instruments['Анализатор'].set_measure_center_freq(value=adjusted_harm, unit='GHz')
                        self._instruments['Анализатор'].set_marker1_x_center(value=adjusted_harm, unit='GHz')
                        if not mock_enabled:
                            time.sleep(0.20)
                        # self._instruments['Анализатор'].set_measure_center_freq(value=harm, unit='GHz')
                        # self._instruments['Анализатор'].set_marker1_x_center(value=harm, unit='GHz')
                        # if not mock_enabled:
                        #     time.sleep(0.20)
                        # self._instruments['Анализатор'].send(f'SYST:PRES')

                    temp.append(self._instruments['Анализатор'].read_pow(marker=1))
                pow_sweep_res.append(temp)
        # ===

        # TODO extract pow sweep func
        # ===
        self._instruments['Генератор'].set_freq(value=param['F'][6], unit='GHz')
        self._instruments['Генератор'].set_pow(value=param['P2'], unit='dBm')
        center_freq = param['mul'] * param['F'][6]
        self._instruments['Анализатор'].set_measure_center_freq(value=center_freq, unit='GHz')
        self._instruments['Анализатор'].set_marker1_x_center(value=center_freq, unit='GHz')

        if not mock_enabled:
            time.sleep(0.5)

        pow2 = self._instruments['Анализатор'].read_pow(marker=1)

        self._instruments['Мультиметр'].send(f'*RST')
        self._instruments['Генератор'].set_output(state='OFF')
        self._instruments['Анализатор'].remove_marker(marker=1)
        self._instruments['Анализатор'].set_autocalibrate(state='ON')
        self._instruments['Источник питания'].set_output(chan=1, state='OFF')
        # TODO implement multimeter display reset

        return pow_sweep_res, pow2

    @property
    def status(self):
        return [i.status for i in self._instruments.values()]

