# BY Thanapon Techasen (Palm)#

# This is an automate test based on Modbus TCP
# It reads data from M172 and M340 at the same time,
# Then log in the data into the Excel file
import threading
from datetime import datetime
from float_rw import FloatModbusClient
from threading import *
from tkinter import *
from sys import exit
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
import matplotlib.pyplot as plt

sem = threading.Semaphore(1)

##############################################################################################################
M340_IP = "192.168.1.103"
M340_PORT = 502


# establish the connection via modbus
m340_modbus = FloatModbusClient(host=M340_IP, port=M340_PORT, auto_open=True)

# start creating GUI
win = Tk()
win.title("M340 Control")

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

variables = variables_read_write+variables_read_only

values_write = [0 for i in range(len(variables_read_write))]  # initialise the corresponding values for writing out
values_read = [0 for i in range(len(variables))]  # initialise the corresponding values for reading in

# define a dictionary of register addresses and the datatype for M340 and M172: [address, datatype]
#############################################################################################################
register_addr_type = {"FCU Mode": [219, 'INT'], "FCU Fan Speed": [213, 'INT'], "FCU Temperature SP": [370, 'FLOAT'],
                      "Main Lab Right Damper": [221, 'INT'], "FCU Fan State": [215, 'INT'],"Temp1": [338, 'FLOAT'], "Temp3": [332, 'FLOAT'],
                      "FCU_supply_temp_PV": [320, 'FLOAT']}

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
buttons_plot = list()

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
    buttons_plot.append(Button(win, text="plot from log", command=lambda k=i: plot(k)))


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
    buttons_plot[i].grid(row=i, column=6)


# plot the file
def plot(option):
    global plot_clicked

    # select the filename to be plotted
    filename = b_filename[int(option)]

    # open and create a list of values from the file
    value_series = []
    time_series = []

    # separate x and y axes
    with open(filename, 'r') as f:
        for line in f:
            plot_time = line.split(',')[0]
            plot_value = line.split(',')[1]

            # turn it into int or float
            if register_addr_type[variables[option]][1].upper() == 'INT':
                plot_value = int(plot_value)
            elif register_addr_type[variables[option]][1].upper() == 'FLOAT':
                plot_value = float(plot_value)

            time_series.append(plot_time)
            value_series.append(plot_value)

    # create the figure that will contain the plot
    fig = Figure(figsize=(5, 5), dpi=100)

    # adding the subplot
    plot1 = fig.add_subplot(111)

    # plotting the graph
    plot1.plot(time_series, value_series)
    plot1.set_title(filename)

    # creating the Tkinter canvas
    # containing the Matplotlib figure
    canvas = FigureCanvasTkAgg(fig,
                               master=win)
    canvas.draw()

    # placing the canvas on the Tkinter window
    canvas.get_tk_widget().grid(row=option, column=7)


# log a line into a file in csv format
def log(filename, *data):
    f = open(filename, 'a')

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
            m340_modbus.write_multiple_registers(register_addr_type[variables_read_write[i]][0], [int(values_write[i])])
        elif register_addr_type[variables_read_write[i]][1].upper() == 'FLOAT':
            m340_modbus.write_float(register_addr_type[variables_read_write[i]][0], [float(values_write[i])])

    sem.release()


# read all the registers
def read_registers():

    sem.acquire() # To ensure the connection is being used by only read or write

    time = datetime.utcnow().isoformat(sep=' ', timespec='milliseconds')

    # read the read-write registers to ensure the value is already set after a click, not overwritten by others
    for i in range(len(variables)):

        # select to read as int or float
        if register_addr_type[variables[i]][1].upper() == 'INT':
            values_read[i] = m340_modbus.read_holding_registers(register_addr_type[variables[i]][0], 1)[0]
            labels_read[i].config(text=values_read[i])
        elif register_addr_type[variables[i]][1].upper() == 'FLOAT':
            values_read[i] = m340_modbus.read_float(register_addr_type[variables[i]][0], 1)[0]
            labels_read[i].config(text=round(values_read[i], 1))

        if b_log_clicked[i]:
            log(b_filename[i], time, str(values_read[i]))

    sem.release()


if __name__ == '__main__':
    t1 = Thread(target=write_read)
    t1.start()

    win.mainloop()

m340_modbus.close()
