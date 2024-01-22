# .
# Notes: Requires Java... Using parts of the Garmin SDK for this part and thats how they do it
#
# Copy entire folder from google drive
# https://drive.google.com/file/d/1SMq9lcJLsMsSC06hnqZXBz2wTmKVcxlh/view?usp=drive_link
# Connect Garmin Xero C1 Pro to Computer
# Run the script (right click + run with powershell) or use the launcher.bat
# Navigate to the .FIT files on the device in the pop up and select them, press "Open" button
# Name the sessions when prompted in terminal window.  Spaces are OK, dont include .txt
#
#
#
# Created by ZenEffect

clear-host
Add-Type -AssemblyName System.Windows.Forms

# temp directory cleaning to be lazy.  suppressed errors will be present if re-running the same fit files but wont break function
remove-item -path .\fit_temp\*.csv -ErrorAction SilentlyContinue
remove-item -path .\fit_temp\*.fit -ErrorAction SilentlyContinue

# file selection window
$Fit = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Garmin fit (*.fit)|*.fit'
    Multiselect = $true}
$null = $Fit.ShowDialog()

# copies to the "fit_temp" folder first for processing.  Does not interact with device except to copy the files to computer first.
if (!$fit.filenames) 
    {
        Write-Output "No Files Selectd `nExiting"
        pause
        break
    }

Copy-Item $fit.filenames .\fit_Temp -ErrorAction SilentlyContinue

# sends the selected files to the bat file that came with the Garmin SDK.  Only modification to the .BAT was to remove the "Pause" at the end.  FIT files
# are processed into CSV by some Garmin SDK Java magic. 
$fit_temp = get-childitem -Path .\fit_temp\*.fit 
foreach ($fitfile in $fit_temp)
    {
        $fit2 = "`"$fitfile`""
        start-process -filepath '.\java\fittocsv.bat' -argumentlist $fit2 -wait
    }

# gets list of newly created .csv from fit_temp directory for processing
$items = get-childitem -Path .\fit_temp\*.csv 
 
# Gets shot data and calculates standard deviation
foreach ($item in $items)
    {
        $csv1 = $item.fullname | import-csv
        $csv_data = $csv1 | where-object 'type' -eq 'data' | where-object 'local number' -eq 4
        $measured = $csv_data.'value 2' | measure-object -average
        $mean = $null
        ForEach ($number in $csv_data) 
            {$mean += [Math]::Pow(($number.'value 2' - $measured.Average), 2)}
        $stdDev = [math]::Sqrt($($mean / ($measured.Count - 1)))
        $stdDevp = [math]::Sqrt($($mean / ($measured.Count)))
        $stddev_fps=$stddev / 304.79999025
        $stddevp_fps=$stddevp / 304.79999025

# Begin CSV to TXT file processing
# csv location detection for general statistics and bullet weight
        foreach ($csv in $csv1)
            {
                $type = $csv.type
		        $localnumber = $csv."local number"

# value correction to feet per second, local time, bullet weight 
                if ($type -eq "Data" -and $localnumber -eq "3")
                    {
                        $min = ($csv."Value 2" / 304.79999025)
                        $max = ($csv."Value 3" / 304.79999025)
                        $average = ($csv."Value 4" / 304.79999025)
                        $shots = $csv."Value 6"
                        $time_hack=([int]$csv."Value 1" + [int]631065600)
                        $time = (([System.DateTimeOffset]::FromUnixTimeSeconds($time_hack)).DateTime.ToLocalTime()).ToString("s")
                        $extreme_spread= ($max - $min)
                        $weight = ($csv."Value 5" / 10)
                        $tdelta = $null                
                        $kenergya = (([Math]::Pow($average ,2) * $weight) / 450437)
                
# sessionn info, user prompt for session name. if empty will still process using file name as base           
                        $Session_name = Read-Host -Prompt "Session Time/Date: $time `nNum of Shots $shots `nBullet Weight gr: $weight `n `nMin: $min `nMax: $max `nAverage: $average `n `nStandard Deviation: $stddev_fps `nPopulation Standard Deviation: $stddevp_fps `nExtreme Spread: $extreme_spread `nAverage Kinetic Energy: $kenergya `n `n `nEnter Session Name"
                        
                        if (!$Session_name) 
                            {$Session_name = $item.Name.Replace('.', '_')                            }
                        
                        $filecheck = get-childitem -Path .\Results\$session_name.txt 
                        if ($Session_name = $filecheck.name)
                            {$Session_name = -join($session_name,"_new")}

# Write TXT file and dummy CSV for later processing into a real CSV              
                        write-output "Session Time/Date: $time `nSession Name: $session_name `nNum of Shots: $shots `nBullet Weight: $weight `n `nMin: $min `nMax: $Max 'nStandard Deviation: $stddev_fps `nPopulation Standard Deviation: $stddevp_fps 'nExtreme Spread: $extreme_spread  `nAverage Kinetic Energy FPE: $kenergya `n" | out-file -filepath .\Results\$Session_name.txt -Append
                        Write-output '"Shot #","Speed (FPS)","Delta from AVG (FPS)","Delta from Last (FPS)","Kinetic Energy (Ft-LBS)","Power Factor (kgr-ft/s)", "Time"' | out-file -filepath .\Results\$Session_name.csx -Append
                        clear-host
                    }
        
# csv location detection for individual shot data and value processing               
                if ($type -eq "Data" -and $localnumber -eq "4")
                    {
                        $shot_num = $csv.'Value 3'
                        $speed = ($csv.'Value 2' / 304.79999025)
                        $time_hack1=([int]$csv."value 1" + [int]631065600)
                        $time1 = (([System.DateTimeOffset]::FromUnixTimeSeconds($time_hack1)).DateTime.ToLocalTime()).ToString("s")
                        $kenergy = (([Math]::Pow($speed ,2) * $weight) / 450437)                     
                        $adelta = $speed - $average  
                        $ldelta =  $speed - $tdelta
                        $pfactor = ($weight * $speed) / 1000
                        $tdelta = $speed
# Write TXT file and dummy CSV for later processing into a real CSV
                        Write-Output "`n `nShot #: $shot_num `nSpeed: $speed `nDelta from Average: $adelta `nDelta from Last: $ldelta `nKinetic Energy FPE: $kenergy `nTime of Shot: $time1 `n" | out-file -filepath .\Results\$Session_name.txt -Append
                        Write-Output "$shot_num,$speed,$adelta,$ldelta,$kenergy,$pfactor,$time1" | out-file -filepath .\Results\$Session_name.csx -Append  
                    } 
            }

# Additional information written to CSV at end of file after shot data processing completes
        write-output "-,,,,,, `nWeight(gr): , $weight ,,,,,,,,,, `nAverage: , $average ,,,,,,,,,, `nSample Standard Deviation (N-1): , $stddev_fps ,,,,,,,,, `nPopulation Standard Deviation: , $stddevp_fps ,,,,,,,,,, `nExtreme Spread: , $extreme_spread ,,,,,,,,,, `nAverage Kinetic Energy: , $kenergya ,,,,,,,,,, `n-,,,,,, `nSession Name: , $Session_name `nSession Time: , $time" | out-file -filepath .\Results\$Session_name.csx -Append

# turning into a real CSV.  Imports the dummy csv then exports it as properly formatted with file headers so it becomes a real csv
        $csvprocessing = import-csv .\results\$session_name.csx
        $csvprocessing | export-csv .\results\$session_name.csv -NoTypeInformation
        remove-item .\results\$session_name.csx
    }  

# cleanup actions
move-item -path .\fit_temp\*.csv  -destination .\results -ErrorAction SilentlyContinue
move-item -path .\fit_temp\*.fit -destination .\fit_backup -ErrorAction SilentlyContinue

# open file explorer window to results
explorer '.\Results'



    