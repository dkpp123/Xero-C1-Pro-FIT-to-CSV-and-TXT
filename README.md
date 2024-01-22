# Xero-C1-Pro-FIT-to-CSV-and-CSV-and-TXT
Powershell Script to convert Garmin Xero C1 Pro .FIT files to CSV(s) and TXT

Requires JAVA

Download and extract entire contents somewhere (its a few empty directories, the script, a .jar and a .bat)



Connect your Garmin Xero C1 Pro to the PC via USB cable

In File Explorer Right click the script + run in powershell or use the launcher.bat. A new File Explorer window will appear

Navigate to the .FIT files located on the Garmin Chronograph 

Highlight the session data you want to export and press the Open button (You can select more than 1 file to batch process)

Some Session information will be displayed

When Prompted Enter the Session Name, this will be the TXT file name.  Its OK to include spaces, DO NOT include a .TXT at the end as its unnecessary

Rinse and Repeat the previous step until you have named all your exported sessions and you are done

Check the "Results" folder for the CSV converted from the FIT (no other processing completed on the file, it is raw) and TXT file which is much more human 
readable and is in Feet per Second.  Feel free to modify the script if you want m/s mph kmph whatever, im not doing it for you.


Text File will include information like:

Session Start Time/Date

Session Name

Min

Max

Average

Standard Deviation

Population Standard Deviation (there are 2 standard deviation methods of calcualation, I included both)

Number of Shots 

Individual Shot Number + Velocity + Time Taken etc...
.

may need to need to run  "Set-ExecutionPolicy -ExecutionPolicy Unrestricted" in powershell if running in powershell.  The launcher.bat file uses a "bypass" policy to allow the script to run without the user needing to change the execution policy

304.79999025 is the MM to FT conversion.  Replace this value if you want M/s MPH/ KMPH, etc.  The base unit of measurement from the Garmin is mm/s.  You can do the math!


.

.

.

.

.

.


used part of the Garmin FIT SDK for FIT to CSV conversion. 

-ZenEffect

the creator claims no responsibility for anything what so ever, use at own risk after reviewing code.  free to use and distribute, don't bother me about bugs I don't care.  lawyers go get bent I'm not responsible this is the disclaimer, so it is said, so it shall be.  If you dont like the .bat or .jar files you can get them from the Garmin FIT SDK found here https://developer.garmin.com/fit/download/.  Just run the script in the directory that contains the Java folder.
