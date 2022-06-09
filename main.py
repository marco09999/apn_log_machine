import Header

# Input variables
path_1 = "P:\Programmes\GROB_7\LOG MACHINE\log_G7.spf"

# Open text file and put each line in a list
with open(path_1, "r") as tf:
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

# Step 3: Compress into single line and number of execution for each unique program
step_3 = []
head_3 = ["Program Name", "Cycle time (min)", "Execution", "Start Date", "Start time", "End date", "End time", "Post",
          "Force", "Otr", "Machine" ]
step_3.append(head_3)
step_3.append(head_3)
step_3.append(head_3)
































