import Header
from init import *
from openpyxl import Workbook
from datetime import datetime


for machines in apn_mach:
    step_3 = []
    if machines["status"] and machines["type"] == "GROB":
        # Open text file and put each line in a list
        with open(machines["path"], "r") as tf:
            line_list = tf.read().split("\n")
            tf.close()
        # Step 1: Transform START and END BLOC into lines
        step_1 = []
        head_1 = ["Status", "Date", "Time", "Pallet", "Name", "Post Version", "Post Date", "Force Version", "Force Date", "Otr",
                  "Line in text file"]
        step_1.append(head_1)
        for i in range(0, len(line_list)):
            if len(Header.line_decode_step_1(line_list, i)) > 1:
                step_1.append(Header.line_decode_step_1(line_list, i))
        # Step 2: Transform START and END lines into execution
        step_2 = []
        head_2 = ["Name", "Cycle time (min)", "Post", "Force", "Otr", "Start Date", "Start Time", "End Date", "End time"]
        step_2.append(head_2)
        for i in range(0, len(step_1)-1):
            if len(Header.line_decode_step_2(step_1, i)) > 1:
                step_2.append(Header.line_decode_step_2(step_1, i))
        # Step 2.5: Add identifier at the end of each execution
        step_2_5 = step_2
        head_2_5 = ["Name", "Cycle time (min)", "Post", "Force", "Otr", "Start Date", "Start Time", "End Date",
                    "End time", "Identifier"]
        step_2_5[0] = head_2_5
        for i in range(1, len(step_2_5)):
            step_2[i].append(step_2[i][0] + "-" + step_2[i][2] + "-" + step_2[i][3] + "-" + step_2[i][4])
        # Step 3: Compress into single line and number of execution for each unique program
        head_3 = ["Program Name", "Cycle time (min)", "Execution", "Start Date", "Start time", "End date", "End time",
                  "Post", "Force", "Otr", "Machine"]
        step_3 = Header.compress(step_2_5)
        # Add machine Number
        for i in step_3:
            i.append(machines["name"])
    # Add update time + write to each machine dictionary their log
    step_3.append("LAST UPDATE: " + datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
    machines["log"] = step_3


# Print out data to command box
for mach_dico in apn_mach:
    print("\n")
    log = mach_dico["log"]
    for line in log:
        print(line)




