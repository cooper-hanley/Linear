from labjack import ljm
import time
import csv
import numpy as np
from openpyxl import load_workbook

# Input Parameters
xValue = '500' # Input your known x-value (i.e. pressure, weight, temperature, etc.)
avename = 'Calibration_Values' # Existing excel file name where value will be logged
Record_Length = 1000 # Number of data points taken (sample rate = .004s), then averaged, for sensor output
#changechange

ave_filename = avename + '.xlsx'
if __name__ == "__main__":

    handle = ljm.openS("T7", "ANY", "ANY")  # T7 device, Any connection, Any identifier

    info = ljm.getHandleInfo(handle)
    print("Opened a LabJack with Device type: %i, Connection type: %i,\n"
      "Serial number: %i, IP address: %s, Port: %i,\nMax bytes per MB: %i" %
      (info[0], info[1], info[2], ljm.numberToIP(info[3]), info[4], info[5]))

    # Setup and call eReadName to read from AIN0 on the LabJack.
    sig_psi = "AIN2"

    # List variables
    start = float(time.time())
    t_stamp_act = []
    t_stamp = ['T_Stamp',]
    v_out = ['V_Out',]

    # Get voltage output from sensor and build v_out list
    #Average .004sec per data point
    #2500 = 10sec
    while len(t_stamp_act) < Record_Length:
        sig = ljm.eReadName(handle, sig_psi)
        if sig >= 0 and sig <=3:
            t_stamp_act.append(time.time())
            v_out.append(sig)
            #print(sig)
            #time.sleep(1)

        else:
            print("INVALID " + str(sig))
            #time.sleep(1)

    # Adjust time stamp to start at zero. Build time stamp list
    for x in range(len(t_stamp_act)):
        tick = float(t_stamp_act[x]) - start
        t_stamp.append(tick)

    # Build dataframe with time stamps and v_out data
    data_r = list(zip(t_stamp,v_out))

    # Write full dataframe to csv file
    filename = xValue + '.csv'
    file = open(filename, 'w', newline='')
    wrapper = csv.writer(file, dialect='excel')
    for line in data_r:
        wrapper.writerow(line)
    file.close()

    sensorOutputAverage = float(np.average(v_out[1:]))
    sensorOutputSTD = float(np.std(v_out[1:]))

    # Write Data Averages and Standard Deviations to new line in Averages Excel Doc
    data_ave = [[xValue, sensorOutputAverage, sensorOutputSTD]]
    wb = load_workbook(ave_filename)
    # Select First Worksheet
    ws = wb.worksheets[0]

    # Append 1 new Row
    for row_data in data_ave:
        # Append Row Values
        ws.append(row_data)

    wb.save(ave_filename)

    # Close handle
    ljm.close(handle)


