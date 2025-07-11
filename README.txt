--- General Overview ---

This program accepts Excel files and an MF4 file as input. If only an Excel file is selected for a job, it will
convert the Excel file to MF4. If an MF4 and Excel file(s) are selected, the data will be merged into a new MF4 file. It will always output an MF4 file.

Multiple jobs can be queued, if one fails to run it will simply abort the conversion for that particular job and continue running the rest of the queue. The program is designed to run in the background with no user intervention until it is complete. Please note that the program is not particularly fast. Unless the window is frozen, the merging process is still happening in the background and it can be left to run.

Additionally, it will pull the summary sheets from each Excel file and save them as a new Excel file. These summary sheets will hold only the values from the original Excel files without any of the calculation process to reach each value. 

The outputted MF4 file and summary files will be in the directory that the original Excel file is located.

--- Capabilities ---

1. Excel to MF4 direct conversion

When adding a job, the MF4 file selection can be canceled and an Excel file can be given in order to convert an Excel file directly to MF4. This will be displayed in the Job queue as just an Excel file.

2. Excel and MF4 merging

Adding a job will, by default, prompt the user to give an MF4 file and an Excel file. These files will be merged together and channels from both will be in the output. The alignment process conducts a cross-correlation on the engine speed channels from both files and shifts the timing of the Excel file based on the lag. This means that both files need engine speed to be present in order to get a valid output.

3. Multiple Excel files and MF4 merge

Multiple Excel files can be merged into one MF4. As long as they are all from the same set of data, they will merge accurately. In order to do so, select an MF4 file as prompted and use CTRL or SHIFT during the Excel select prompt in order to input multiple Excel files. These files will be time-aligned and append, no differently than the single file process. However, since they will introduce duplicate variables, the signals will be numbered (I.e. file 1's values could be Engine Speed while file 2's are Engine Speed(1)). Take note of this and be sure to search for these additional values when looking through your data.

--- Troubleshooting ---

1. I cannot run it

Upon opening the .exe, Windows may try to block you. The pop-up may only provide you with the button "Don't run". If this message is encountered, hit "More info" and a run option should appear.

2. It is taking a long time to run

Ultimately, this program is fairly slow due to the scale of data that we handle. Allow it to run in the background and it will eventually complete processing. Keeping your computer plugged in while it runs will shorten your runtime significantly.

3. There is no output

If the tasks ran without error, there should be an output file in the directory that housed the original Excel file. It will have the same name with _merged appended and will be an MF4 file. The summary Excel sheets will also be here. If the inputted files are stored on a network drive, there is also a possibility that network desync during the merging process interfered with data transmission. Move the files locally and try again.

If there are any unexpected issues encountered, please message me
JEFFREY.LIU@GMAIL.COM


