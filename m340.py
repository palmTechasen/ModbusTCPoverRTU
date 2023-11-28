# BY Thanapon Techasen (Palm)#

# This is an automate test based on Modbus TCP
# It reads data from M172 and M340 at the same time,
# Then log in the data into the Excel file
import threading
from datetime import datetime
from float_rw import FloatModbusClient
from threading import *
from tkinter import *
from matplotlib.figure import Figure
from matplotlib import pyplot as plt
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
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

variables_read_only = list()
variables_read_only.append("Temp1")
variables_read_only.append("Temp3")
variables_read_only.append("FCU_supply_temp_PV")
# variables_read_only.append("DesignAirflow_D3")
# variables_read_only.append("RawReads[3,11]")
variables_read_only.append("R3_Perf_CO2")
variables_read_only.append("CO2_D3")
variables_read_only.append("R3_Bel_Temp")

variables = variables_read_write+variables_read_only

values_write = [0 for i in range(len(variables_read_write))]  # initialise the corresponding values for writing out
values_read = [0 for i in range(len(variables))]  # initialise the corresponding values for reading in

# define a dictionary of register addresses and the datatype for M340 and M172: [address, datatype]
#############################################################################################################
register_addr_type = {"FCU Mode": [219, 'INT', "M340"], "FCU Fan Speed": [213, 'INT', "M340"], "FCU Temperature SP": [370, 'FLOAT', "M340"],
                      "Main Lab Right Damper": [221, 'INT', "M340"], "FCU Fan State": [215, 'INT', "M340"],"Temp1": [338, 'FLOAT', "M340"], "Temp3": [332, 'FLOAT', "M340"],
                      "FCU_supply_temp_PV": [320, 'FLOAT', "M340"], "R3_Perf_CO2": [8972,'INT', "M172", 0.1], "CO2_D3": [9302,'INT', "M172", 0.1], "R3_Bel_Temp": [8989,'INT', "M172", 4]}

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
figures = list()
canva = list()
# create string variables for entry widget
entry = list()
for i in range(len(variables_read_write)):
    entry.append(StringVar())

# initialise the GUI widgets
# The first four are read-write registers and the last three are read-only registers
for i in range(len(variables)):
    labels_text.append(Label(win, text=variables[i]))
    labels_read.append(Label(win, text=""))

    if i < len(variables_read_write):  # widget only for writing the holding registers from index 0-3
        entries.append(Entry(win, textvariable=entry[i]))
        buttons_write.append(Button(win, text="write", command=lambda k=i: write_select(k)))

    buttons_log.append(Button(win, text="log", command=lambda k=i: log_select(k)))
    buttons_stop_log.append(Button(win, text="stop log", command=lambda k=i: log_deselect(k), state=DISABLED))
    scales_plot.append(Scale(win, label="Figure", from_=0, to=len(variables)-1, orient=HORIZONTAL, length=200)) 
    buttons_plot.append(Button(win, text="plot from log", command=lambda k=i: plot(k,scales_plot[k].get())))
    figures.append(Figure(figsize=(1, 1), dpi=50))
    canva.append(FigureCanvasTkAgg(figures[i], master=win))

# arrange the GUI
# The first four are read-write registers and the last three are read-only registers
for i in range(len(variables)):
    labels_text[i].grid(row=i, column=0)
    labels_read[i].grid(row=i, column=1)

    if i < len(variables_read_write):  # widget only for writing the holding registers from index 0-3
        entries[i].grid(row=i, column=2)
        buttons_write[i].grid(row=i, column=3)

    buttons_log[i].grid(row=i, column=4)
    buttons_stop_log[i].grid(row=i, column=5)
    scales_plot[i].grid(row=i, column=6)
    buttons_plot[i].grid(row=i, column=7)
    canva[i].get_tk_widget().grid(row=i, column=8)

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
            time_series.append(plot_time)
            value_series.append(plot_value)

    # create the figure that will contain the plot
    fig = figures[figure]
    # fig.set_size_inches(5, 5, forward=True)
    # fig.set_dpi(100)
    print(f"{figure} contains {option} from {filename}")

    # adding the subplot
    if option < 10:
        plot1 = fig.add_subplot(option*111)
    elif option < 91:
        plot1 = fig.add_subplot(option*11)
    elif option < 1000:
        plot1 = fig.add_subplot(option*1)

    # plotting the graph
    plot1.plot(time_series, value_series)
    plot1.set_title(filename)
    plt.show()
    canva[figure].draw()

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
            temp = connections[register_addr_type[variables[i]][2]].read_holding_registers(register_addr_type[variables[i]][0], 1)[0]
        elif register_addr_type[variables[i]][1].upper() == 'FLOAT':
            temp = connections[register_addr_type[variables[i]][2]].read_float(register_addr_type[variables[i]][0], 1)[0]

        if addresses[register_addr_type[variables[i]][2]][2] == False:
            values_read[i] = temp * register_addr_type[variables[i]][3]
        else:
            values_read[i] = temp

        labels_read[i].config(text=round(values_read[i], 1))
        if b_log_clicked[i]:
            log(b_filename[i], time, str(values_read[i]))
    sem.release()


if __name__ == '__main__':
    t1 = Thread(target=write_read)
    t1.start()

    win.mainloop()

for i in range(len(devices)):
	connections[devices[i]].close()
