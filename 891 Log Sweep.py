import pyvisa
import csv
import time
from datetime import datetime
import xlsxwriter
#####_____ initialization ______#####
print("#######################################")
print("#######################################")
print("#### 891 LCR Meter Sweep Function  ####")
print("#######################################")
print("#######################################\n\n")
print("     ##### INITIALIZATION #####\n")
rm = pyvisa.ResourceManager()
li = rm.list_resources()
choice = ''
while(choice == ''):
    for index in range(len(li)):
        print(str(index)+" - "+li[index])
    choice = input("Select DUT: ")
    try: 
        if(int(choice) > len(li) - 1 or int(choice) < 0):
            choice = ''
            print("Invalid Input\n")
    except:
        print("Invalid Input\n")
        choice = ''
inst = rm.open_resource(li[int(choice)])
inst.timeout= 1000
DATE_TIME = "%s"%datetime.now() 

#####_____ Command ______#####
ID = "*IDN?\n"
print(inst.query(ID))
IDE= "%s"%(inst.query("*IDN?\n"))

# Set or return the sweep mode <LINear/0 | LOGarithm/1 >
inst.write("SWEEP:MODE 0\n")
#Set or return function in sweep linear graph mode <CSQ/0 | CSD/1 | 
#CSR/2 | CPQ/3 | 
#CPD/4 | CPR/5 | 
#CPG/6 | LSQ/7 | 
#LSD/8 | LSR/9 | 
#LPQ/10 | LPD/11 | 
#LPR/12 | LPG/13 | 
#ZTH/14 | YTH/15 | 
#RX/16 | GB/17>
inst.write("SWEEP:DISP:LIN:GRAP ZTH\n")
#Set or return sweep auto frequency status <1/ON | 0/OFF> 
inst.write("SWEEP:AUTO:FREQ 0\n")
#Set or return sweep steps <NR1 | ?>
inst.write("SWEEP:STEP 10\n")
#Set stop and start frequency
inst.write("SWEEP:FREQUENCY:START 35000\n")
inst.write("SWEEP:FREQUENCY:STOP 41000\n")
#Set or return the sweep primary and secondary swap parameter <1/ON | 0/OFF> 
inst.write("SWEEP:SWAP OFF")
#Start sweep
inst.write("SWEEP:START ON")
print("Please wait while the sweep completes. This can take a couple of seconds")
sweepstate= "%s"%(inst.query("SWEEP:BUSY?\n"))

while True:
    sweepstate= "%s"%(inst.query("SWEEP:BUSY?\n"))
#    print(sweepstate)
    if(sweepstate=="OFF\r\n"):
        print("\n")
        print("sweep complete\n\n")
        break

#create file workbook and worksheet 

name = input("Enter a name for the Excel File: ")
outWorkbook  = xlsxwriter.Workbook(name + ".xlsx")
Sheet1 = outWorkbook.add_worksheet("Sheet1") 
cell_format = outWorkbook.add_format()
cell_format.set_num_format("0.00E+00")
# Create a format to use in the merged range.  
merge_format = outWorkbook.add_format({ 
 
    'bold': 1, 
    'border': 1, 
    'align': 'center', 
    'valign': 'vcenter'})

centerbold_format = outWorkbook.add_format({ 
 
    'bold': 1, 
    'border': 0, 
    'align': 'center', 
    'valign': 'vcenter'})
cell_format.set_num_format(11) 
Sheet1.set_column(0, 1, 16)

# Merge 6 columns for ID 
Sheet1.merge_range('A1:F1', IDE, merge_format) 

#Basic information of acquired data 
Sheet1.write('A3', 'Date and Time', centerbold_format) 
Sheet1.merge_range('B3:D3', DATE_TIME,centerbold_format) 
Sheet1.write('A5', "Frequency", centerbold_format) 
Sheet1.write('B5', "Impedance", centerbold_format) 

 
    
i = 0
j = 6
k = 6
while i <301:
    point= "SWEEP:POINT %d"%i
    print(point)
    inst.write(point)
    inst.timeout= 1000
    data = "%s"%(inst.query("SWEEP:DATA?\n"))
    print(data)
    datae = data.replace("\r\n","")
    Freq_Data = float(datae.split(",",1)[0])
    print(Freq_Data)
    Imp_Data = float(datae.split(",",1)[1])
    print(Imp_Data)
    inst.timeout= 1000
    freqline = "A%d"%j
    impline = "B%d"%k
    Sheet1.write(freqline, Freq_Data, cell_format) 
    Sheet1.write(impline, Imp_Data, cell_format) 

    i += 1    
    j += 1
    k += 1
print("Data Acquired")
#Create a chart
chart = outWorkbook.add_chart({"type":"scatter", "subtype":"smooth_with_markers"})
chart.set_size({'width': 720, 'height': 450})
Sheet1.insert_chart("E6",chart)
#formatting axis
chart.set_x_axis({
    'name': 'Frequency',
    'name_font': {'size': 14, 'bold': True},
    'num_font':  {'italic': True },
})
chart.set_y_axis({
    'name': 'Impedance',
    'name_font': {'size': 14, 'bold': True},
    'num_font':  {'italic': True },
})

#Add a sereis to the chart
chart.add_series({
    'categories': '=Sheet1!$A$6:$A$306',
    'values':     '=Sheet1!$B$6:$B$306',
    'line':       {'color': 'blue'},
})

outWorkbook.close()

inst.close()