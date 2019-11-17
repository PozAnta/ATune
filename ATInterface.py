from tkinter import *
import serial
import ATScript_ver_01


def exit_gui():
    master.destroy()


def ok():

    parameters = ['LMJR', 'KNLD', 'KNLI', 'KNLIV', 'KNLP', 'KNLUSERGAIN', 'NLANTIVIBGAIN2', 'NLANTIVIBGAIN3',
                  'NLANTIVIBHZ2', 'NLANTIVIBHZ3', 'NLANTIVIBSHARP2', 'NLANTIVIBSHARP3', 'NLANTIVIBQ3',
                  'MOVESMOOTHLPFHZ', 'NLTFBW', 'NLFILTDAMPING', 'NLFILTT1', 'KNLAFRC', 'HDTUNEDEBUG', 'IGRAV',
                  'gearfiltmode', 'gearfiltdepth']

    additional_parameters = ['AccDec', 'Velocity', 'Distance', 'Result']

    print("Number Iterations: ", numiterations.get())
    print("Path for save results: ", path.get())

    adv_traj = [accVar.get(), disVar.get(), vcruiseVar.get()]
    adv_options = [hdAvTuneVar.get(), hdFfTuneVar.get(), igrav_app.get()]
    mail_param = [mail_check.get(), mail_app.get(), mail_app_pass.get()]

    tune = ATScript_ver_01.Tune(name_tune.get(), parameters, int(numiterations.get()), port_var.get(), path.get(),
                         additional_parameters, mechSetup.get(), adv_traj, adv_options, mail_param,
                         smart_factor_check.get())

    print("Velocity: ", vcruiseVar.get())
    print("Distance: ", disVar.get())
    print("Acc: ", accVar.get())

    if name_tune.get() == "Express":

        print("Selected Tuning: Express")

        master.destroy()
        tune.easy_tune()

    if name_tune.get() == "Advance":

        if vcruiseVar.get() == "0" or disVar.get() == "0" or accVar.get() == "0":
            print("One of trajectory parameters is zero!")
            print("Set correct trajectory")
            print("==================================\n")
            return

        print("Selected Tuning: Advance")

        master.destroy()
        tune.advance_tune()

    if name_tune.get() == "Express External":

        print("Selected Tuning: Express External")

        master.destroy()
        tune.easy_external_tune()

    if name_tune.get() == "Advance External":

        print("Selected Tuning: Advance External")

        master.destroy()
        tune.advance_external_tune()


def serial_ports():
    ports = ['COM%s' % (i + 1) for i in range(256)]

    result = []
    for port in ports:
        try:
            s = serial.Serial(port)
            s.close()
            result.append(port)
        except (OSError, serial.SerialException):
            # result.append('No find ports')
            # return result
            pass

    return result


ver_file = "Ver 2.0"
'''

Added smart factor for PTPVCMD plot
Option for select use smart factor
Add path for mail password
add record calculation time interval 4

'''
master = Tk()
col = 'powder blue'
OPTIONS = ["Express", "Advance", "Express External", "Advance External"]

port_var = StringVar(master)
port_var.set(serial_ports()[0])

name_tune = StringVar(master)
name_tune.set(OPTIONS[0])  # default value

hdAvTuneVar = BooleanVar()
hdFfTuneVar = BooleanVar()
igrav_app = BooleanVar()
mechSetup = BooleanVar()
seDrive = BooleanVar()
mail_check = BooleanVar()
smart_factor_check = BooleanVar()

disVar = StringVar(master)
vcruiseVar = StringVar(master)
accVar = StringVar(master)

master.title("Auto Tune " + ver_file)
label_1 = Label(master, text="Number Iterations", font=('arial', 11, 'bold'))
label_1.grid(row=0, column=0, sticky=W)

numiterations = Entry(master, bd=5, insertwidth=1, bg=col, width=7, justify='left')
numiterations.grid(row=1, column=0, padx=10, sticky=W)
numiterations.delete(0, END)
numiterations.insert(0, "1")

# ===============Path for results========================
label_2 = Label(master, text="Path to save results", font=('arial', 11, 'bold'))
label_2.grid(row=2, column=0, sticky=W)

path = Entry(master, bd=5, insertwidth=1, bg=col, width=29, justify='left', font=('arial', 10, 'bold'))
path.grid(row=3, column=0, padx=10, sticky=W)
path.delete(0, END)
path.insert(0, "C:\AT\Python")

# ===============Menu for Ports========================
label_3 = Label(master, text="Select Port", font=('arial', 11, 'bold'))
label_3.grid(row=0, column=2, sticky=W)

# ===============Menu for Options========================
win_ports = OptionMenu(master, port_var, *serial_ports())
win_ports.config(bd=4, bg=col, font=('arial', 10, 'bold'))
win_ports.grid(row=1, column=2, sticky=W)

label_4 = Label(master, text="Mode of Tuning", font=('arial', 12, 'bold'))
label_4.grid(row=2, column=2, sticky=W)

w = OptionMenu(master, name_tune, *OPTIONS)
w.config(bd=5, bg=col, font=('arial', 10, 'bold'))
w.grid(row=3, column=2, sticky=W)

# ===============Menu for Advance tuning========================
label_4 = Label(master, text="Advance Options", font=('arial', 12, 'bold'))
label_4.grid(sticky=W)

cb = Checkbutton(master, text="HDTUNEAVMODE", variable=hdAvTuneVar, font=('arial', 10, 'bold'))
cb.grid(row=5, column=0, sticky=W)

cb = Checkbutton(master, text="HDTUNENAFRC", variable=hdFfTuneVar, font=('arial', 10, 'bold'))
cb.grid(row=6, column=0, sticky=W)

cb = Checkbutton(master, text="IGRAV", variable=igrav_app, font=('arial', 10, 'bold'))
cb.grid(row=7, column=0, sticky=W)

cb = Checkbutton(master, text="Belt or Ball Screw?", variable=mechSetup, font=('arial', 10, 'bold'))
cb.grid(row=8, column=0, sticky=W)

cb = Checkbutton(master, text="SE Drive", variable=seDrive, font=('arial', 10, 'bold'))
cb.grid(row=9, column=0, sticky=W)

# ================Menu for smart factor================
cb = Checkbutton(master, text="Smart factor for PTPVCMD", variable=smart_factor_check, font=('arial', 10, 'bold'))
cb.grid(row=10, column=0, sticky=W)

# ===============Menu for Ports========================
label_9 = Label(master, text="Trajectory Advance", font=('arial', 11, 'bold'))
label_9.grid(row=4, column=2, sticky=W)

label_5 = Label(master, text="Vcruise [rpm]", font=('arial', 9, 'bold')).grid(row=5, column=2)
vcruiseVar = Entry(master, bd=5, insertwidth=1, bg=col, justify='left', font=('arial', 10, 'bold'))
vcruiseVar.grid(row=6, column=2, sticky=W)
vcruiseVar.delete(0, END)
vcruiseVar.insert(0, "0")

label_6 = Label(master, text="Distance [rev]", font=('arial', 9, 'bold')).grid(row=7, column=2)
disVar = Entry(master, bd=5, insertwidth=1, bg=col, justify='left', font=('arial', 10, 'bold'))
disVar.grid(row=8, column=2, sticky=W)
disVar.delete(0, END)
disVar.insert(0, "0")

label_7 = Label(master, text="AccDec [rpm/s]", font=('arial', 9, 'bold')).grid(row=9, column=2)
accVar = Entry(master, bd=5, insertwidth=1, bg=col, justify='left', font=('arial', 10, 'bold'))
accVar.grid(row=10, column=2, sticky=W)
accVar.delete(0, END)
accVar.insert(0, "0")

# ================Menu for mail=====================
cb = Checkbutton(master, text="Send mail?", variable=mail_check, font=('arial', 10, 'bold'))
cb.grid(row=14, column=0, sticky=W)

label_7_1 = Label(master, text="Address mail", font=('arial', 9, 'bold')).grid(row=15, column=0, sticky=W)

mail_app = Entry(master, bd=5, insertwidth=1, bg=col, justify='left', width=35)
mail_app.grid(row=16, padx=10, column=0, sticky=W)

label_7_2 = Label(master, text="Password mail", font=('arial', 9, 'bold')).grid(row=18, column=0, sticky=W)

mail_app_pass = Entry(master, show="*", bd=5, insertwidth=1, bg=col, justify='left', width=35)
mail_app_pass.grid(row=19, padx=10, column=0, sticky=W)
# ================Menu for buttons=====================
label_8 = Label(master, text="                                     Start Tuning", font=('arial', 11, 'bold')).grid(
    row=24)

button = Button(master, text="OK", font='bolt', bd=5, bg=col, width=10, command=ok).grid(row=30, column=0)
button1 = Button(master, text="Exit", font='bolt', bd=5, bg=col, width=10, command=exit_gui).grid(row=30, column=1,
                                                                                                  sticky=W)

mainloop()
