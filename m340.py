# BY Thanapon Techasen (Palm)#

# This is an automate test based on Modbus TCP
# It reads data from M172 and M340 at the same time,
# Then log in the data into the Excel file
import threading
from datetime import datetime
from float_rw import FloatModbusClient
from threading import *
from tkinter import *
# from tkinter.ttk import Scale
from ttkwidgets import TickScale
from matplotlib.figure import Figure
from matplotlib import pyplot as plt
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from pyModbusTCP import utils
import os.path

sem = threading.Semaphore(1)

##############################################################################################################
# M340_IP = "192.168.1.103"
# M340_PORT = 502

##############################################################################################################
# M172_IP = "192.168.1.182"
# M172_PORT = 502

devices = list()
devices.append("M340")
devices.append("M172")

Diffusers = 3
DiffuserSpacing = 200
DiffuserEnds = []

for i in range(Diffusers):
    DiffuserEnds.append('_D'+str(i+1))

DiffuserEndsTuple = tuple(DiffuserEnds)

addresses = {"M340": ["192.168.1.103", 502, True], "M172": ["192.168.1.182", 502, False]}
# Boolean: are all the variables pre-processed and ready to use?

connections = {}

for i in range(len(devices)):
	connections[devices[i]] = FloatModbusClient(host=addresses[devices[i]][0], port=addresses[devices[i]][1], auto_open=True)

# establish the connection via modbus

# start creating GUI
win = Tk()
win.title("M340&M172 Control")

# text describing the file name and the label in the UI
# The first group is for read-write registers and the second group is for read-only registers
##############################################################################################################
variables_read_write = list()
variables_read_write.append("FCU Mode")  # index 0
variables_read_write.append("FCU Fan Speed")
variables_read_write.append("FCU Temperature SP")
variables_read_write.append("Main Lab Right Damper")
variables_read_write.append("FCU Fan State")
variables_read_write.append("T_SP_D1")
variables_read_write.append("T_db_D1")

variables_read_only = list()
variables_read_only.append("Temp1")
variables_read_only.append("Temp3")
variables_read_only.append("FCU_supply_temp_PV")
# variables_read_only.append("DesignAirflow_D3")
# variables_read_only.append("RawReads[3,11]")
# variables_read_only.append("R3_Perf_CO2")
# variables_read_only.append("R3_Bel_Temp")
variables_read_only.append("Room_T_D1")
variables_read_only.append("Supply_T_D1")
variables_read_only.append("Pressure_D1")
variables_read_only.append("CO2_D1")
variables_read_only.append("Position_D1")
variables_read_only.append("Airflow_D1")
variables_read_only.append("T_CO2_D1")
variables_read_only.append("PresRelief_D1")
variables_read_only.append("PresOffset_D1")

rows = len(variables_read_write)+len(variables_read_only)
Drows = 0
for i in range(len(variables_read_write)):
    if(variables_read_write[i].endswith('_D1')):
        Drows = Drows + 1
        for j in range(Diffusers - 1):
            variables_read_write.append(variables_read_write[i].replace('_D1', '_D' + str(j+2)))

for i in range(len(variables_read_only)):
    if(variables_read_only[i].endswith('_D1')):
        for j in range(Diffusers - 1):
            variables_read_only.append(variables_read_only[i].replace('_D1', '_D' + str(j+2)))

variables = variables_read_write+variables_read_only

values_write = [0 for i in range(len(variables_read_write))]  # initialise the corresponding values for writing out
values_read = [66666 for i in range(len(variables))]  # initialise the corresponding values for reading in

# define a dictionary of register addresses and the datatype for M340 and M172: [address, datatype]
#############################################################################################################
# register_addr_type = {"FCU Mode": [219, 'INT', "M340"], "FCU Fan Speed": [213, 'INT', "M340"], "FCU Temperature SP": [370, 'FLOAT', "M340"],
#                       "Main Lab Right Damper": [221, 'INT', "M340"], "FCU Fan State": [215, 'INT', "M340"],"Temp1": [338, 'FLOAT', "M340"], "Temp3": [332, 'FLOAT', "M340"],
#                       "FCU_supply_temp_PV": [320, 'FLOAT', "M340"], "R3_Perf_CO2": [8972,'INT', "M172", 0.1], "CO2_D3": [9302,'INT', "M172", 0.1],
#                       "R3_Bel_Temp": [8989,'INT', "M172", 4], "Position_D3": [9305,'INT', "M172", 1.0], "Airflow_D3": [9309, 'INT', "M172", 0.1],
#                       "Pressure_D3": [9301, 'INT', "M172", 0.1], "Room_T_D3": [9299, 'INT', "M172", 0.1], "Supply_T_D3": [9300, 'INT', "M172", 0.1],
#                       "T_SP_D3": [9319, 'INT', "M172", 0.1], "T_deadband_D3": [9320, 'INT', "M172", 0.1], "Position_D1": [9105,'INT', "M172", 1.0], "Position_D2": [9205,'INT', "M172", 1.0],
#                       "TemperatureCO2_D1": [9124, 'INT', "M172", 0.1], "TemperatureCO2_D2": [9224, 'INT', "M172", 0.1], "TemperatureCO2_D3": [9324, 'INT', "M172", 0.1],
#                       "Room_T_D1": [9099, 'INT', "M172", 0.1],"Room_T_D2": [9199, 'INT', "M172", 0.1], "PresRelief_D1": [9121, 'INT', "M172", 0.001], "PresOffset_D1": [9163, 'INT', "M172", 0.1],
#                       "PresRelief_D2": [9221, 'INT', "M172", 0.001], "PresOffset_D2": [9263, 'INT', "M172", 0.1],"PresRelief_D3": [9321, 'INT', "M172", 0.001], "PresOffset_D3": [9363, 'INT', "M172", 0.1]}

# register_addr_type = {"FCU Mode": [219, 'INT', "M340"], "FCU Fan Speed": [213, 'INT', "M340"], "FCU Temperature SP": [370, 'FLOAT', "M340"],
#                       "Main Lab Right Damper": [221, 'INT', "M340"], "FCU Fan State": [215, 'INT', "M340"],"Temp1": [338, 'FLOAT', "M340"], "Temp3": [332, 'FLOAT', "M340"],
#                       "FCU_supply_temp_PV": [320, 'FLOAT', "M340"], 
#                       "Room_T_D3": [9499, 'INT', "M172", 0.1], "Supply_T_D3": [9500, 'INT', "M172", 0.1],"Pressure_D3": [9501, 'INT', "M172", 0.1], "CO2_D3": [9502,'INT', "M172", 0.1],
#                       "Position_D3": [9504,'INT', "M172", 1.0], "Airflow_D3": [9507, 'INT', "M172", 0.1],
#                       "T_SP_D3": [9523, 'INT', "M172", 0.1], "T_deadband_D3": [9524, 'INT', "M172", 0.1], "Position_D1": [9104,'INT', "M172", 1.0], "Position_D2": [9304,'INT', "M172", 1.0],
#                       "TemperatureCO2_D1": [9193, 'INT', "M172", 0.1], "TemperatureCO2_D2": [9393, 'INT', "M172", 0.1], "TemperatureCO2_D3": [9593, 'INT', "M172", 0.1],
#                       "Room_T_D1": [9099, 'INT', "M172", 0.1],"Room_T_D2": [9299, 'INT', "M172", 0.1], "PresRelief_D1": [9128, 'INT', "M172", 0.001], "PresOffset_D1": [9113, 'INT', "M172", 0.1],
#                       "PresRelief_D2": [9328, 'INT', "M172", 0.001], "PresOffset_D2": [9313, 'INT', "M172", 0.1],"PresRelief_D3": [9528, 'INT', "M172", 0.001], "PresOffset_D3": [9513, 'INT', "M172", 0.1]}

register_addr_type = {"FCU Mode": [219, 'INT', "M340"], "FCU Fan Speed": [213, 'INT', "M340"], "FCU Temperature SP": [370, 'FLOAT', "M340"],
                      "Main Lab Right Damper": [221, 'INT', "M340"], "FCU Fan State": [215, 'INT', "M340"],"Temp1": [338, 'FLOAT', "M340"], "Temp3": [332, 'FLOAT', "M340"],
                      "FCU_supply_temp_PV": [320, 'FLOAT', "M340"], 
                      "Room_T_D1": [9099, 'INT', "M172", 0.1], "Supply_T_D1": [9100, 'INT', "M172", 0.1],"Pressure_D1": [9101, 'INT', "M172", 0.1], "CO2_D1": [9102,'INT', "M172", 0.1],
                      "Position_D1": [9104,'INT', "M172", 1.0], "Airflow_D1": [9107, 'INT', "M172", 0.1],
                      "T_SP_D1": [9123, 'INT', "M172", 0.1], "T_db_D1": [9124, 'INT', "M172", 0.1], 
                      "T_CO2_D1": [9193, 'INT', "M172", 0.1], "PresRelief_D1": [9128, 'INT', "M172", 0.001], "PresOffset_D1": [9113, 'INT', "M172", 0.1]}

for i in list(register_addr_type.keys()):
    if(i.endswith('_D1')):
        for j in range(Diffusers - 1):
            if addresses[register_addr_type[i][2]][2] == False:
                register_addr_type.update({i.replace('_D1', '_D' + str(j+2)): [register_addr_type[i][0]+DiffuserSpacing*(j+1), register_addr_type[i][1], register_addr_type[i][2], register_addr_type[i][3]]})
            else:
                register_addr_type.update({i.replace('_D1', '_D' + str(j+2)): [register_addr_type[i][0]+DiffuserSpacing*(j+1), register_addr_type[i][1], register_addr_type[i][2]]})

# write value to the plc only once per click
write_clicked = False

# can be logged multiple at a time
b_log_clicked = [False for i in range(len(variables))]
b_filename = [variables[i]+".csv" for i in range(len(variables))]

# plot from the file once per click
plot_clicked = False

# widgets for reading  in and writing out input register
labels_text = list()
labels_read = list()
entries = list()
buttons_write = list()
buttons_log = list()
buttons_stop_log = list()
scales_plot = list()
buttons_plot = list()
# create string variables for entry widget
entry = list()
for i in range(len(variables_read_write)):
    entry.append(StringVar())

## ExternalFrame = Frame(win, width=1700, height=990, background="orange")
## DiffusersScrollCanvas = Canvas(ExternalFrame, bg="green")
## DiffusersScrollFrame = Frame(DiffusersScrollCanvas, background="white")
DiffusersScrollFrame = Frame(win, background="white")
# initialise the GUI widgets
# The first four are read-write registers and the last three are read-only registers
for i in range(len(variables)):
    if(variables[i].endswith(DiffuserEndsTuple)):
        root = DiffusersScrollFrame
    else:
        root = win
    labels_text.append(Label(root, text=variables[i]))
    labels_read.append(Label(root, text=""))

    if i < len(variables_read_write):  # widget only for writing the holding registers from index 0-3
        entries.append(Entry(root, textvariable=entry[i]))
        buttons_write.append(Button(root, text="write", command=lambda k=i: write_select(k)))

    buttons_log.append(Button(root, text="log", command=lambda k=i: log_select(k)))
    buttons_stop_log.append(Button(root, text="stop", command=lambda k=i: log_deselect(k), state=DISABLED))
    scales_plot.append(TickScale(root, from_=0, to=len(variables)-1, orient=HORIZONTAL, length=50, resolution=1, showvalue=True, labelpos="e")) 
    buttons_plot.append(Button(root, text="plot", command=lambda k=i: plot(k,scales_plot[k].get())))

j = 0
columnextra = 0
for i in range(8*Diffusers):
    win.grid_columnconfigure(i, weight=1)
for i in range(rows):
    win.grid_rowconfigure(i, weight=1)
for i in range(8*Diffusers):
    DiffusersScrollFrame.grid_columnconfigure(i, weight=1)
for i in range(Drows):
    DiffusersScrollFrame.grid_rowconfigure(i, weight=1)
# arrange the GUI
# The first four are read-write registers and the last three are read-only registers
k = 0
rowD = 0
rowsDict = {}
for i in range(len(variables)):
    if(variables[i].endswith(DiffuserEndsTuple)):
        columnextra = (int(variables[i].rpartition("_D")[2])-1)*8
        if(variables[i].endswith('_D1')):
            rowD = j
        else:
            rowD = rowsDict[variables[i].rpartition("_D")[0]+'_D1']
        labels_text[i].grid(row=rowD, column=(0 + columnextra), sticky="NWES")
        labels_read[i].grid(row=rowD, column=(1 + columnextra), sticky="NWES")
        if i < len(variables_read_write):  # widget only for writing the holding registers from index 0-3
            entries[i].grid(row=rowD, column=(2 + columnextra), sticky="NWES")
            buttons_write[i].grid(row=rowD, column=(3 + columnextra), sticky="NWES")
        buttons_log[i].grid(row=rowD, column=(4 + columnextra), sticky="NWES")
        buttons_stop_log[i].grid(row=rowD, column=(5 + columnextra), sticky="NWES")
        scales_plot[i].grid(row=rowD, column=(6 + columnextra), sticky="NWES")
        buttons_plot[i].grid(row=rowD, column=(7 + columnextra), sticky="NWES")
        if(variables[i].endswith('_D1')):
            rowsDict.update({variables[i]: j})
            j = j+1
    else:
        columnextra = 0
        labels_text[i].grid(row=k, column=(0 + columnextra), sticky="NWES")
        labels_read[i].grid(row=k, column=(1 + columnextra), sticky="NWES")
        if i < len(variables_read_write):  # widget only for writing the holding registers from index 0-3
            entries[i].grid(row=k, column=(2 + columnextra), sticky="NWES")
            buttons_write[i].grid(row=k, column=(3 + columnextra), sticky="NWES")
        buttons_log[i].grid(row=k, column=(4 + columnextra), sticky="NWES")
        buttons_stop_log[i].grid(row=k, column=(5 + columnextra), sticky="NWES")
        scales_plot[i].grid(row=k, column=(6 + columnextra), sticky="NWES")
        buttons_plot[i].grid(row=k, column=(7 + columnextra), sticky="NWES")
        k = k+1
DiffusersScrollFrame.grid(row=k, column=0, columnspan=8*Diffusers, rowspan=Drows, sticky="NWES")
l=Label(win, text=str(Diffusers)).grid(column=8*Diffusers-1, row=0)
## ExternalFrame.grid(row=j, column=0, columnspan=8, sticky="nsew")
## DiffusersScrollCanvas.grid(row=0, column=0, sticky="nsew")
## DiffusersScrollCanvas.create_window(0, 0, window=DiffusersScrollFrame, anchor='nw')
## DiffusersScrollbar = Scrollbar(ExternalFrame, orient=HORIZONTAL)
## DiffusersScrollbar.config(command=DiffusersScrollCanvas.yview)
## DiffusersScrollCanvas.config(yscrollcommand=DiffusersScrollbar.set)
## DiffusersScrollbar.grid(row=0, column=1, sticky="ew")
## DiffusersScrollCanvas.configure(scrollregion=DiffusersScrollCanvas.bbox('all'))
# plot the file
def plot(option,figure):
    global plot_clicked

    # select the filename to be plotted
    filename = b_filename[int(option)]

    # open and create a list of values from the file
    value_series = []
    time_series = []

    # separate x and y axes
    with open(filename, 'r') as f:

        line = f.readline()  # skip the header line of the file

        for line in f:
            plot_time = line.split(',')[0]
            plot_value = line.split(',')[1]

            # turn it into int or float
            # if register_addr_type[variables[option]][1].upper() == 'INT':
            #     plot_value = int(plot_value)
            # elif register_addr_type[variables[option]][1].upper() == 'FLOAT':
            #     plot_value = float(plot_value)
            plot_value = float(plot_value)
            #if len(value_series) == 0 or value_series[-1] != plot_value:
            time_series.append(plot_time)
            value_series.append(plot_value)

    # plotting the graph
    plt.figure(figure)
    plt.plot(time_series, value_series, lw=0, marker='o', label=filename)
    plt.legend()
    plt.show()

# log a line into a file in csv format
def log(filename, *data):
    filepath = './/'+filename
    file_exist = os.path.isfile(filepath) # check if the file already exist before opening

    # open the file and write the header if not previously exist
    f = open(filename, 'a')
    if not file_exist:
        f.write("Time"+","+filename.split('.')[0])  # print the header row using Time and header without .csv or .txt
        f.write("\n")

    # replace with a list as the index is needed to determine if the comma will be printed
    l_data = list(data)
    length = len(l_data)

    for i in range(length):
        f.write(str(l_data[i]))

        if i < length-1:
            f.write(',')

    f.write("\n")
    f.close()


# select what data to be logged
def log_select(option):
    b_log_clicked[option] = True
    buttons_log[option].config(state=DISABLED)
    buttons_stop_log[option].config(state=NORMAL)


# de-select the logged data
def log_deselect(option):
    b_log_clicked[option] = False
    buttons_log[option].config(state=NORMAL)
    buttons_stop_log[option].config(state=DISABLED)


# select which data to be written
def write_select(option):
    global write_clicked
    write_clicked = True  # a button clicked
    values_write[int(option)] = float(entry[option].get())


# write and read data in a while-true loop
def write_read():

    global write_clicked
    while True:

        if write_clicked:
            write_registers() # write only when clicked
            write_clicked = False
        read_registers()


# write all the registers
def write_registers():

    sem.acquire()  # To ensure the connection is being used by only read or write

    for i in range(len(variables_read_write)):

        # select to write out as float or int
        if register_addr_type[variables_read_write[i]][1].upper() == 'INT':
            connections[register_addr_type[variables[i]][2]].write_multiple_registers(register_addr_type[variables_read_write[i]][0], [int(values_write[i])])
        elif register_addr_type[variables_read_write[i]][1].upper() == 'FLOAT':
            connections[register_addr_type[variables[i]][2]].write_float(register_addr_type[variables_read_write[i]][0], [float(values_write[i])])

    sem.release()


# read all the registers
def read_registers():

    sem.acquire() # To ensure the connection is being used by only read or write

    time = datetime.utcnow().isoformat(sep=' ', timespec='milliseconds')

    # read the read-write registers to ensure the value is already set after a click, not overwritten by others
    for i in range(len(variables)):
        # select to read as int or float
        if register_addr_type[variables[i]][1].upper() == 'INT':
            temp = utils.get_2comp(connections[register_addr_type[variables[i]][2]].read_holding_registers(register_addr_type[variables[i]][0], 1)[0], 16)
        elif register_addr_type[variables[i]][1].upper() == 'FLOAT':
            temp = connections[register_addr_type[variables[i]][2]].read_float(register_addr_type[variables[i]][0], 1)[0]

        if addresses[register_addr_type[variables[i]][2]][2] == False:
            temp = temp * register_addr_type[variables[i]][3]
        else:
            temp = temp

        labels_read[i].config(text=round(temp, 1))

        if b_log_clicked[i] and (values_read[i] == 66666 or values_read[i] != temp) and (temp != 0 or values_read[i] != 0):
            log(b_filename[i], time, str(temp))

        values_read[i] = temp

    sem.release()


if __name__ == '__main__':
    t1 = Thread(target=write_read)
    t1.start()

    win.mainloop()

for i in range(len(devices)):
	connections[devices[i]].close()
