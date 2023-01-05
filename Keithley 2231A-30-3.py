import pyvisa
import time
import datetime
import openpyxl
import keyboard

d = datetime.datetime.now()
wb = openpyxl.Workbook()
ws = wb.create_sheet("TEST")  # Name of the sheet

rm = 0
my_PS = 0
Total_time = 0
Loop_count = 0

def datalogging(x,y,data):
    ws.cell(row=x, column=y , value=data)
def dataoutput(finish_time):
    wb.save(r'D:\Dropbox\노승범\실험\TOF\Laser\20220728\220%\Fabrication_170g\20221214\90g_retry\Practice_'
            +finish_time.strftime('%Y%m%d%H%M')+'.xlsx')
    print('Data is saved!')

def KEI2231_Connect(rsrcString, getIdStr, timeout, doRst):
    my_PS = rm.open_resource(rsrcString, baud_rate = 9600, data_bits = 8)	#opens desired resource and sets it variable my_instrument
    my_PS.write_termination = '\n'
    my_PS.read_termination = '\n'
    my_PS.send_end = True
    my_PS.StopBits = 1
    if getIdStr == 1:
        print(my_PS.query("*IDN?"))
    my_PS.write('SYST:REM')
    my_PS.timeout = timeout
    if doRst == 1:
        return my_PS

def KEI2231A_Disconnect():
    my_PS.write('SYST:LOC')
    my_PS.close
    return

def KEI2231A_SelectChannel(myChan):
    my_PS.write("INST:NSEL %d" % myChan)
    return

def KEI2231A_SetVoltage(myV,myI):
    my_PS.write('APPLy CH1,%0.4f %0.4f' %(myV,myI)) #마지막은 current 범위
    return

def KEI2231A_OutputState(myState):
    if myState == 0:
        my_PS.write("OUTP 0")
    else:
        my_PS.write("OUTP 1")
    return
#================================================================================
#    MAIN CODE GOES HERE
#================================================================================

rm = pyvisa.ResourceManager()   # Opens the resource manager and sets it to variable rm

my_PS = KEI2231_Connect("ASRL5::INSTR", 1, 20000, 1) #VISA communication
KEI2231A_SelectChannel(1)
KEI2231A_OutputState(1) # 1일 경우에는 ON

########Getting a initial resistance code
KEI2231A_SetVoltage(1,1)
my_PS.write('FETC:CURR?')
Current = my_PS.read()
Initial_resistance = 1/Current
print(Initial_resistance)
#########################################
while True:
    if keyboard.is_pressed("a"):
        break
    t1 = time.time()    # Capture start time....a
    Input_voltage = 5
    KEI2231A_SetVoltage(Input_voltage,1)
    my_PS.write('FETC:CURR?')
    Current = my_PS.read()
    Resistance = Input_voltage/float(Current)
    print('Current value: ',"%0.4f" %float(Current))
    print('Resistance value:', "%0.4f" %Resistance)
    t2 = time.time()    # Capture stop time...
    print("Loop time: ","{0:.4f} s".format((t2-t1)))
    Total_time = (t2 - t1) + Total_time
    Loop_count = Loop_count + 1
    datalogging(Loop_count+2,1,Total_time)
    datalogging(Loop_count+2,2,Input_voltage)
    datalogging(Loop_count+2,3,Resistance)

dataoutput(d)
KEI2231A_OutputState(0)
KEI2231A_Disconnect()

rm.close



