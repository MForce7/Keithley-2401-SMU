import numpy as np 
import time
import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
import serial.tools.list_ports
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt5.QtGui import QIcon, QDoubleValidator
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QComboBox, QLabel, QVBoxLayout, QHBoxLayout, QPushButton, QGroupBox, QLineEdit, QStyle, QFileDialog, QGridLayout

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Keithley 2401 SMU")
        self.setWindowIcon(QIcon(r"Icon_keithley-min.png"))
        self.setGeometry(100, 100, 1280, 720)
        
        #COMM
        self.title_comm = QLabel('COMM')
        self.list_comm = QComboBox()
        self.list_comm.addItems([])
        self.list_comm.currentTextChanged.connect(self.comm_changed)
        self.comm = ''
        self.refresh_comm = QPushButton()
        self.refresh_icon = self.style().standardIcon(QStyle.SP_BrowserReload) 
        self.refresh_comm.setIcon(self.refresh_icon)
        self.refresh_comm.clicked.connect(self.detect_com_ports)
        
        # baudrate  
        self.title_baudrate = QLabel('Baud')
        self.list_baud = QComboBox()
        self.list_baud.addItems(['9600', '19200', '38400', '57600', '115200'])
        self.list_baud.currentTextChanged.connect(self.baud_changed)
        self.baud = '9600'
        
        # mode
        self.title_mode = QLabel('Mode')
        self.list_mode = QComboBox()
        self.list_mode.addItems(['2-Probe', '4-Probe'])
        self.list_mode.currentTextChanged.connect(self.probe_changed)
        self.probe = '2-Probe'
        
        # source
        self.title_source = QLabel('Source')
        self.list_source = QComboBox()
        self.list_source.addItems(['Current', 'Voltage'])
        self.list_source.currentTextChanged.connect(self.source_changed)
        self.source = 'Current'
        
        # sense
        self.title_sense = QLabel('Sense')
        self.list_sense = QComboBox()
        self.list_sense.addItems(['Voltage', 'Current'])
        self.list_sense.currentTextChanged.connect(self.sense_changed)
        self.sense = 'Voltage'
        
        # Start
        self.title_Start = QLabel('Start')
        self.box_start = QLineEdit()
        self.box_start.setFixedWidth(100)
        self.float_validator = QDoubleValidator(self.box_start)
        self.float_validator.setDecimals(10000)
        self.float_validator.setRange(-1000.0, 1000.0)  
        self.box_start.setText("-5")
        self.box_start.textChanged.connect(self.start_value)
        self.box_start.setValidator(self.float_validator)
        self.start_point = "-5"
        
        # Step
        self.title_Step = QLabel('Step')
        self.box_step = QLineEdit()
        self.box_step.setFixedWidth(100)
        self.float_validator = QDoubleValidator(self.box_step)
        self.float_validator.setDecimals(10000)
        self.float_validator.setRange(-1000.0, 1000.0)  
        self.box_step.setText("10")
        self.box_step.textChanged.connect(self.step_value)
        self.box_step.setValidator(self.float_validator)
        self.step_point = "10"
   
        # Stop
        self.title_Stop = QLabel('Stop')
        self.box_stop = QLineEdit()
        self.box_stop.setFixedWidth(100)
        self.float_validator = QDoubleValidator(self.box_stop)
        self.float_validator.setDecimals(10000)
        self.float_validator.setRange(-1000.0, 1000.0)  
        self.box_stop.setText("5")
        self.box_stop.textChanged.connect(self.stop_value)
        self.box_stop.setValidator(self.float_validator)
        self.stop_point = "5"
        self.stop_point_string = str(self.stop_point)
        
        # Protection
        self.title_Prot = QLabel('Protect')
        self.box_Prot = QLineEdit()
        self.box_Prot.setFixedWidth(100)
        self.float_validator = QDoubleValidator(self.box_Prot)
        self.float_validator.setDecimals(10000)
        self.float_validator.setRange(-1000.0, 1000.0)  
        self.box_Prot.setText("0.030")
        self.box_Prot.textChanged.connect(self.prot_value)
        self.box_Prot.setValidator(self.float_validator)
        self.prot_point = "0.030"
        self.protection = float(self.prot_point)
        self.prot_point_string = str(self.prot_point)
        
        # Path
        self.title_path = QLabel('Path')
        self.box_path = QLabel()
        # self.box_path.setFixedWidth(100)
        self.box_path.setText(os.path.join(os.path.expanduser("~"), "Documents"))
        # self.box_path.textChanged.connect(self.stop_value)
        self.output_file_path = os.path.join(os.path.expanduser("~"), "Documents")
        self.browse_path = QPushButton()
        self.browse_icon = self.style().standardIcon(QStyle.SP_DialogOpenButton)  # Ikon tempat sampah
        self.browse_path.setIcon(self.browse_icon)
        self.browse_path.clicked.connect(self.open_directory)
        
        # File Name
        self.title_file_name = QLabel('File Name')
        self.box_file_name = QLineEdit()
        self.box_file_name.setFixedWidth(100)
        self.box_file_name.setText('Output_keithley')
        self.box_file_name.textChanged.connect(self.file_name_value)
        self.output_file_name = 'Output_keithley'
        self.real_output_file_name = self.output_file_path + "\\" + self.output_file_name + ".xlsx"
        
        self.addition = 1
        while os.path.isfile(self.real_output_file_name):
            self.real_file_name = self.output_file_name + "_" + str(self.addition)
            self.real_output_file_name = self.output_file_path + "\\" + self.output_file_name +"_" + str(self.addition) + ".xlsx"
            self.box_file_name.setText(self.real_file_name)
            self.addition += 1
        
        self.canvas = FigureCanvas(plt.figure())
        self.line_graph = QVBoxLayout()
        self.line_graph.addWidget(self.canvas)
        
        #Play Button
        tombol_play = QPushButton("Mulai")
        tombol_play.clicked.connect(self.kode_utama)
        
        #Plot Button
        tombol_plot = QPushButton("Plot")
        tombol_plot.clicked.connect(self.plot_graph)
        line_button_control = QHBoxLayout()
        line_button_control.addWidget(tombol_play)
        line_button_control.addWidget(tombol_plot)
        
        setup_layout = QGridLayout()
        setup_layout.setColumnStretch(0, 2)
        setup_layout.setColumnStretch(1, 2)
        setup_layout.setColumnStretch(2, 1)
        setup_layout.addWidget(self.title_comm, 0, 0)
        setup_layout.addWidget(self.list_comm, 0, 1)
        setup_layout.addWidget(self.refresh_comm, 0, 2)
        setup_layout.addWidget(self.title_baudrate, 1, 0)
        setup_layout.addWidget(self.list_baud, 1, 1)
        setup_layout.addWidget(self.title_mode, 2, 0)
        setup_layout.addWidget(self.list_mode, 2, 1)
        setup_layout.addWidget(self.title_source, 3, 0)
        setup_layout.addWidget(self.list_source, 3, 1)
        setup_layout.addWidget(self.title_sense, 4, 0)
        setup_layout.addWidget(self.list_sense, 4, 1)
        
        measure_layout = QGridLayout()
        measure_layout.setColumnStretch(0, 2)
        measure_layout.setColumnStretch(1, 2)
        measure_layout.setColumnStretch(2, 1)
        measure_layout.addWidget(self.title_Start, 0, 0)
        measure_layout.addWidget(self.box_start, 0, 1)
        measure_layout.addWidget(self.title_Step, 1, 0)
        measure_layout.addWidget(self.box_step, 1, 1)
        measure_layout.addWidget(self.title_Stop, 2, 0)
        measure_layout.addWidget(self.box_stop, 2, 1)
        measure_layout.addWidget(self.title_Prot, 3, 0)
        measure_layout.addWidget(self.box_Prot, 3, 1)
        
        path_layout = QGridLayout()
        path_layout.setColumnStretch(0, 2)
        path_layout.setColumnStretch(1, 2)
        path_layout.setColumnStretch(2, 1)
        path_layout.addWidget(self.title_path, 0, 0)
        path_layout.addWidget(self.box_path, 0, 1)
        path_layout.addWidget(self.browse_path, 0, 2)
        path_layout.addWidget(self.title_file_name, 1, 0)
        path_layout.addWidget(self.box_file_name, 1, 1)
        
        #Setup Group
        setup_grup_box = QGroupBox("Setup Connection")
        setup_line = QVBoxLayout()
        setup_line.addLayout(setup_layout)
        setup_grup_box.setLayout(setup_line)
        
        #Measure Group
        measure_grup_box = QGroupBox("Setup Measurement")
        measure_line = QVBoxLayout()
        measure_line.addLayout(measure_layout)
        measure_grup_box.setLayout(measure_line)

        #Path Group
        path_grup_box = QGroupBox("Setup Output File")
        path_line = QVBoxLayout()
        path_line.addLayout(path_layout)
        path_grup_box.setLayout(path_line)
        
        nav_container = QWidget()
        nav_side_line = QVBoxLayout()
        nav_side_line.setContentsMargins(10, 20, 10, 20)
        nav_side_line.addWidget(setup_grup_box)
        nav_side_line.addWidget(measure_grup_box)
        nav_side_line.addWidget(path_grup_box)
        nav_container.setLayout(nav_side_line)
        nav_container.setFixedSize(300, 720)
        
        show_side_line = QVBoxLayout()
        show_side_line.addLayout(self.line_graph)
        show_side_line.addLayout(line_button_control)

        window_line = QHBoxLayout()
        window_line.addWidget(nav_container)
        window_line.addLayout(show_side_line)
        
        container = QWidget()
        container.setLayout(window_line)
        self.setCentralWidget(container)
    
    def printting(self):
        print(type(self.step_point))
        
    def baud_changed(self, text):
        self.baud = text    
    def probe_changed(self, text):
        self.probe = text    
    def source_changed(self, text):
        self.source = text    
    def sense_changed(self, text):
        self.sense = text    
    def comm_changed(self, text):
        self.comm = text    
    
    def start_value(self, text):
        self.start_point = text    
    def step_value(self, text):
        self.step_point = text    
    def stop_value(self, text):
        self.stop_point = text    
    def prot_value(self, text):
        self.prot_point = text    
        
    def file_name_value(self, text):
        self.output_file_name = text  
        self.real_output_file_name = self.output_file_path + "\\" + self.output_file_name + ".xlsx"
        
    def open_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Pilih Directory")
        self.output_file_path = directory
        self.box_path.setText(self.output_file_path)
        
        self.real_output_file_name = self.output_file_path + "\\" + self.output_file_name + ".xlsx"
        self.addition = 1
        while os.path.isfile(self.real_output_file_name):
            self.real_file_name = self.output_file_name + "_" + str(self.addition)
            self.real_output_file_name = self.output_file_path + "\\" + self.output_file_name +"_" + str(self.addition) + ".xlsx"
            self.box_file_name.setText(self.real_file_name)
            self.addition += 1
        
    def detect_com_ports(self):
        ports = serial.tools.list_ports.comports()
        available_ports = []
        for port in ports:
            available_ports.append(port.device)
        self.list_comm.clear()
        self.list_comm.addItems(available_ports)
    
    def plot_data(self, data):
        # Bersihkan grafik lama
        self.canvas.figure.clear()
        # Membuat grafik baru
        ax = self.canvas.figure.add_subplot(111)
        ax.plot(data[self.source], data[self.sense], marker='o', linestyle='-', color='b', label=f'{self.source} vs {self.sense}')
        # Tambahkan label, judul, dll.
        ax.set_xlabel(f'Source ({self.source})')
        ax.set_ylabel(f'Sense({self.sense})')
        ax.set_title(f'Grafik {self.sense} vs {self.source}')
        ax.legend()
        ax.grid(True)
        # Perbarui canvas
        self.canvas.draw()

    def load_excel_file(self): 
        self.real_output_file_name = self.output_file_path + "\\" + self.output_file_name + ".xlsx"
        if self.real_output_file_name:
            try:
                data = pd.read_excel(self.real_output_file_name)
                if self.source in data.columns and self.sense in data.columns:
                    self.plot_data(data)
                else:
                    print("excel tidak memiliki column source dan sense")
            except Exception as e:
                print(f"Error saat membaca file: {e}")

    def kode_utama(self):
            # Konfigurasi awal
        max_current = float(self.prot_point)  # Maksimum arus dalam Amps
        min_voltage = float(self.start_point)      # Minimum tegangan
        max_voltage = float(self.stop_point)      # Maksimum tegangan
        number_of_steps = float(self.step_point)
        step_voltage = (max_voltage - min_voltage) / number_of_steps

        num_steps = str(number_of_steps + 1)
        min_v = str(min_voltage)
        max_v = str(max_voltage)
        step_v = str(step_voltage)
        max_c = str(max_current)

        try:
            ser = serial.Serial(self.comm, baudrate=9600, timeout=1, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE)

            # Fungsi untuk mengirim perintah ke Keithley
            def send_command(command):
                ser.write((command + "\n").encode())
                time.sleep(0.1)

            def read_response():
                response = ser.readline().decode().strip()
                return response

            # Model 2400 Spesifik
            send_command('*RST')
            send_command('*ESE 0')
            send_command('*CLS')
            send_command('STAT:MEAS:ENAB 1024')
            send_command('*SRE 1')

            # Konfigurasi buffer
            send_command('TRAC:CLE')
            send_command(f'TRAC:POIN {num_steps}')

            # Konfigurasi Sweep
            send_command('SOUR:FUNC:MODE VOLT')
            send_command(f'SOUR:VOLT:STAR {min_v}')
            send_command(f'SOUR:VOLT:STOP {max_v}')
            send_command(f'SOUR:VOLT:STEP {step_v}')
            send_command('SOUR:CLE:AUTO ON')
            send_command('SOUR:VOLT:MODE SWE')
            send_command('SOUR:SWE:SPAC LIN')
            send_command('SOUR:DEL:AUTO OFF')
            send_command('SOUR:DEL 0.10')

            # Konfigurasi pengukuran
            send_command('SENS:FUNC "CURR"')
            send_command('SENS:FUNC:CONC ON')
            send_command('SENS:CURR:RANG:AUTO ON')
            send_command(f'SENS:CURR:PROT:LEV {max_c}')
            send_command('SENS:CURR:NPLC 1')
            send_command('FORM:ELEM:SENS VOLT,CURR')
            send_command(f'TRIG:COUN {num_steps}')
            send_command('TRIG:DEL 0.001')
            send_command('SYST:AZER:STAT OFF')
            send_command('SYST:TIME:RES:AUTO ON')
            send_command('TRAC:TST:FORM ABS')
            send_command('TRAC:FEED:CONT NEXT')

            # Memulai Sweep
            send_command('OUTP ON')
            send_command('INIT')

            time.sleep(2)

            # Membaca data dari buffer
            send_command('TRAC:DATA?')
            data = read_response()
            data = np.array(data.split(','), dtype=float)

            # Parsing data
            currents = data[1::2]
            voltages = data[0::2]
            
            data_hasil = {
                self.source: voltages,
                self.sense : currents
            }
            print("data frame done")
            df = pd.DataFrame(data_hasil)
            self.real_output_file_name = self.output_file_path + "\\" + self.output_file_name + ".xlsx"
            self.addition = 1
            while os.path.isfile(self.real_output_file_name):
                self.real_file_name = self.output_file_name + "_" + str(self.addition)
                self.real_output_file_name = self.output_file_path + "\\" + self.output_file_name +"_" + str(self.addition) + ".xlsx"
                self.box_file_name.setText(self.real_file_name)
                self.addition += 1
                
            if self.real_output_file_name:
                df.to_excel(self.real_output_file_name, index=False)
            else:
                df.to_excel(r"Output_01\output.xlsx", index=False)
            print("excel aman")
            
            self.load_excel_file()

            # Reset dan bersihkan
            send_command('*RST')
            send_command('*CLS')
            send_command('*SRE 0')
            ser.close()
        except Exception as e:
            print("comm tidak ditemukan")
            print(f"Error = {e}")

    
    def plot_graph(self):
        try:
            self.load_excel_file()
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())