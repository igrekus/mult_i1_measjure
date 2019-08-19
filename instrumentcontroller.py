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
from agilentn5230amock import AgilentN5230AMock
from agilentn9030amock import AgilentN9030AMock
from instr.agilent34410a import Agilent34410A
from instr.agilentN5230A import AgilentN5230A
from instr.agilente3644a import AgilentE3644A
from instr.agilentn5183a import AgilentN5183A
from instr.agilentn9030a import AgilentN9030A

mock_enabled = False


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


class SpectrumAnalyzerFactory(InstrumentFactory):
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


class NetworkAnalyzerFactory(InstrumentFactory):
    def __init__(self, addr):
        super().__init__(addr=addr, label='Анализатор')
        self.applicable = ['N5230A']
    def from_address(self):
        if mock_enabled:
            return AgilentN5230A(self.addr, '1,N5230A mock,1', AgilentN5230AMock())
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
            try:
                raw_data: pandas.DataFrame = pandas.read_excel(filename, sheet_name=dev)
            except Exception as ex:
                print('Error:', ex)
                continue
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
            'Источник питания': SourceFactory('GPIB1::4::INSTR'),
            'Мультиметр': MultimeterFactory('GPIB1::2::INSTR'),
            'Генератор': GeneratorFactory('GPIB1::20::INSTR'),
            'Анализатор': NetworkAnalyzerFactory('GPIB1::10::INSTR'),
        }

        self.deviceParams = {
            'Тип 3': {
                'F': 6.0,
                'Pmin': 15,
                'Pmax': 21,
                'Istat': [
                    [None, None, None],
                    [None, None, None],
                    [None, None, None]
                ],
                'Idyn': [
                    [None, None, None],
                    [None, None, None],
                    [None, None, None]
                ]
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
        # TODO implement pna-check
        return self.result.init() and self._runCheck(self.deviceParams[device], self.secondaryParams[secondary])

    def _runCheck(self, param, secondary):
        return True
        # 50 mA
        threshold = -120.0
        if isfile('./settings.ini'):
            with open('./settings.ini', 'rt') as f:
                threshold = float(f.readline().strip().split('=')[1])

        if param['Istat'][0][0] is not None:
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

    def _pna_init(self):
        user_preset_name = r'C:\Program Files\Agilent\Network Analyzer\Documents\UserPreset.sta'

        self._instruments['Анализатор'].send('SYST:PRES')
        self._instruments['Анализатор'].query('*OPC?')

        self._instruments['Анализатор'].send(f'SYST:UPR:LOAD "{user_preset_name}"')
        self._instruments['Анализатор'].send(f'SYST:UPR')

        self._instruments['Анализатор'].send('CALC:PAR:DEL:ALL')
        # self._instruments['Анализатор'].send('DISP:WIND2 ON')

        self._instruments['Анализатор'].send('CALC1:PAR:DEF "MEAS_1",B,1')   # TODO use required meas param
        self._instruments['Анализатор'].send('CALC1:FORM MLOG')
        self._instruments['Анализатор'].send('DISP:WIND1:TRAC1:FEED "MEAS_1"')
        self._instruments['Анализатор'].send('DISP:WIND1:TRAC1:Y:SCAL:AUTO')

        # c:\program files\agilent\newtowrk analyzer\UserCalSets
        # self._instruments['Анализатор'].send('SENS1:CORR:CSET:ACT "-20dBm_1.1-1.4G",1')

    def _syncRig(self):
        sweep_points = 301
        smooth_points = 30

        self._instruments['Анализатор'].send(f'TRIG:SOUR EXT')
        self._instruments['Анализатор'].send(f'TRIG:SCOP CURR')
        self._instruments['Анализатор'].send(f'SENS1:SWE:MODE CONT')
        self._instruments['Анализатор'].send(f'SENS1:SWE:TRIG:MODE POIN')

        # self._instruments['Анализатор'].send(f'TRIG:ROUTE:INP MAIN')   # error TODO replace with preset load
        self._instruments['Анализатор'].send(f'TRIG:TYPE EDGE')
        self._instruments['Анализатор'].send(f'TRIG:SLOP POS')
        self._instruments['Анализатор'].send(f'CONT:SIGN:TRIG:ATBA ON')
        self._instruments['Анализатор'].send(f'TRIG:READ:POL LOW')

        self._instruments['Анализатор'].send(f'TRIG:CHAN1:AUX1 ON')
        self._instruments['Анализатор'].send(f'TRIG:CHAN1:AUX1:OPOL POS')
        self._instruments['Анализатор'].send(f'TRIG:CHAN1:AUX1:POS AFT')
        # self._instruments['Анализатор'].send(f'TRIG:CHAN1:AUX1:DUR?')
        self._instruments['Анализатор'].send(f'SENS1:SWE:POIN {sweep_points}')
        self._instruments['Анализатор'].send(f'SENS1:FOM ON')

        # ass plot smothing
        self._instruments['Анализатор'].send(f'CALC1:SMO ON')
        self._instruments['Анализатор'].send(f'CALC1:SMO:POIN {smooth_points}')

        # self._instruments['Генератор'].set_pow(value=15, unit='dBm')

        self._instruments['Генератор'].send(f':FREQ:MODE LIST')
        self._instruments['Генератор'].send(f':LIST:TYPE STEP')
        self._instruments['Генератор'].send(f':INIT:CONT OFF')
        self._instruments['Генератор'].send(f':SWE:POIN {sweep_points}')
        self._instruments['Генератор'].send(f':LIST:TRIG:SOUR EXT')
        self._instruments['Генератор'].send(f':LIST:MODE AUTO')
        self._instruments['Генератор'].send(f':TRIG:SOUR IMM')
        self._instruments['Генератор'].send(f':POW:ATT:AUTO ON')

        self._set_harmonic(harmonic=1)

        # self._instruments['Генератор'].send('SWE:DWEL .5')
        # self._instruments['Генератор'].send('INIT')

    def _set_harmonic(self, harmonic=1):
        harm_offset = {
            1: (0.1, 40),
            2: (0.1, 25),
            3: (0.1, 16.6),
            4: (0.1, 12.5)
        }

        self._instruments['Анализатор'].send(f'SENS1:FOM:RANG1:FREQ:STAR {harm_offset[harmonic][0]}GHz')
        self._instruments['Анализатор'].send(f'SENS1:FOM:RANG1:FREQ:STOP {harm_offset[harmonic][1]}GHz')
        self._instruments['Анализатор'].send(f'SENS1:FOM:RANG3:FREQ:MULT {harmonic}')

        self._instruments['Генератор'].send(f':FREQ:STAR {harm_offset[harmonic][0]}GHz')
        self._instruments['Генератор'].send(f':FREQ:STOP {harm_offset[harmonic][1]}GHz')

    def _measure(self, device, secondary):
        param = self.deviceParams[device]
        secondary = self.secondaryParams[secondary]
        print(f'launch measure with {param} {secondary}')

        self._instruments['Генератор'].send('*CLS')
        self._instruments['Генератор'].set_modulation(state='OFF')
        self._instruments['Генератор'].send(f':POW:ATT:AUTO ON')

        self._pna_init()

        # TODO extract static measure func
        # ===
        if param['Istat'][0][0] is not None:
            self._instruments['Источник питания'].set_current(chan=1, value=420, unit='mA')
            self._instruments['Источник питания'].set_voltage(chan=1, value=5.55, unit='V')
            self._instruments['Источник питания'].set_output(chan=1, state='ON')

            self._instruments['Генератор'].set_freq(value=param['F'], unit='GHz')
            self._instruments['Генератор'].set_pow(value=param['Pmax'], unit='dBm')
            self._instruments['Генератор'].set_output(state='ON')

            # TODO adjust current value gen towards new algorithm
            curr = int(MeasureResultMock.generate_value(param['Istat'][secondary]) * 10)
            curr_str = ' 00.' + f'{curr}  ADC'.replace('.', ',')
            self._instruments['Мультиметр'].send(f'DISPlay:WIND1:TEXT "{curr_str}"')

            if not mock_enabled:
                time.sleep(3)

            self._instruments['Генератор'].set_output(state='OFF')
            self._instruments['Источник питания'].set_output(chan=1, state='OFF')
            # self._instruments['Мультиметр'].send(f'SYST:PRES')

        self._syncRig()

        # TODO extract dynamic measure func
        # ===
        if param['Istat'][0][0] is not None:
            self._instruments['Источник питания'].set_current(chan=1, value=300, unit='mA')
            self._instruments['Источник питания'].set_voltage(chan=1, value=4.45, unit='V')
            self._instruments['Источник питания'].set_output(chan=1, state='ON')

        # TODO extract pow sweep
        # ===
        self._instruments['Генератор'].set_pow(value=param['Pmin'], unit='dBm')
        # self._instruments['Генератор'].set_freq(value=param['F'], unit='GHz')
        self._instruments['Генератор'].set_output(state='ON')

        if param['Idyn'][0][0] is not None:
            curr = int(MeasureResultMock.generate_value(param['Idyn'][secondary]) * 10)
            curr_str = ' 00.' + f'{curr}  ADC'.replace('.', ',')
            self._instruments['Мультиметр'].send(f'DISPlay:WIND1:TEXT "{curr_str}"')

        for mul in [1, 2, 3, 4]:
            self._set_harmonic(harmonic=mul)
            self._instruments['Генератор'].send(f':INIT')
            self._instruments['Генератор'].query('*OPC?')
            #if not mock_enabled:
            #    time.sleep(0.3)
            self._instruments['Анализатор'].send('DISP:WIND1:TRAC1:Y:SCAL:AUTO')
            #if not mock_enabled:
            #    time.sleep(0.3)

        # TODO extract freq sweep func
        # ===
        self._instruments['Генератор'].set_pow(value=param['Pmax'], unit='dBm')
        # self._instruments['Генератор'].set_freq(value=param['F'], unit='GHz')
        self._instruments['Генератор'].set_output(state='ON')

        if param['Idyn'][0][0] is not None:
            curr = int(MeasureResultMock.generate_value(param['Idyn'][secondary]) * 10)
            curr_str = ' 00.' + f'{curr}  ADC'.replace('.', ',')
            self._instruments['Мультиметр'].send(f'DISPlay:WIND1:TEXT "{curr_str}"')

        for mul in [1, 2, 3, 4]:
            self._set_harmonic(harmonic=mul)
            self._instruments['Генератор'].send(f':INIT')
            self._instruments['Генератор'].query('*OPC?')
            #if not mock_enabled:
            #    time.sleep(0.3)
            self._instruments['Анализатор'].send('DISP:WIND1:TRAC1:Y:SCAL:AUTO')
            #if not mock_enabled:
            #    time.sleep(0.3)

        self._instruments['Анализатор'].send('CALC:PAR:DEL:ALL')

        self._instruments['Мультиметр'].send(f'SYST:PRES')
        self._instruments['Генератор'].set_output(state='OFF')
        self._instruments['Источник питания'].set_output(chan=1, state='OFF')

        return ['ok'], 'ok'

    @property
    def status(self):
        return [i.status for i in self._instruments.values()]

