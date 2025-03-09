# Global variables
$global:CCID = $null
$global:Case = $null

# Load error keywords from config file at startup
$keywordFile = "$env:USERPROFILE\Desktop\EcoDesk\Config\ErrorKeywords.json"
if (Test-Path $keywordFile) {
    try {
        $keywords = Get-Content -Path $keywordFile -Raw | ConvertFrom-Json -ErrorAction Stop
        $global:ErrorKeywords = $keywords.ErrorKeywords
        $global:WarningKeywords = $keywords.WarningKeywords
    } catch {
        Write-Host "Failed to load error keywords from $keywordFile $_" -ForegroundColor Red
        Write-Host "Using default keywords." -ForegroundColor Yellow
        $global:ErrorKeywords = @("error", "fail", "exception", "critical", "fatal")
        $global:WarningKeywords = @("warning", "caution")
    }
} else {
    $global:ErrorKeywords = @("error", "fail", "exception", "critical", "fatal")
    $global:WarningKeywords = @("warning", "caution")
}

function Main {
    param (
        [string]$CCID = $global:CCID,
        [string]$Case = $global:Case
    )

    Write-Host "Extract Logs from CELogs.." -ForegroundColor Cyan
    Clear-Host
    
    $global:CCID = if ([string]::IsNullOrWhiteSpace($CCID)) { Read-Host -Prompt "Enter CCID" } else { $CCID }
    $global:Case = if ([string]::IsNullOrWhiteSpace($Case)) { Read-Host -Prompt "Enter Ticket Number" } else { $Case }
    if ($global:CCID -match "\s") {
        do {
            $global:CCID = Read-Host -Prompt "CCID cannot contain spaces. Enter CCID again"
        } while ($global:CCID -match "\s")
    }
    Write-Host "Using CCID: $global:CCID, Case: $global:Case" -ForegroundColor Green

    $Date = (Get-Date).ToString('yyyy-MM-dd')
    $Year = (Get-Date).Year
    $attempts = 0
    $maxAttempts = 3
    $validChoice = $false
    
    while ($attempts -lt $maxAttempts -and -not $validChoice) { 
        Clear-Host
        Write-Host "******************************************************"
        Write-Host "* Select an option:                                  *"
        Write-Host "* ==================                                 *"
        Write-Host "* 1. Select 1 for CE LOGS                            *"
        Write-Host "* 2. Select 2 for TITAN LOGS                         *"
        Write-Host "* 3. Select 3 to extract all files from CELogs       *"
        Write-Host "******************************************************"
        $choice = Read-Host "Enter your Choice"
        switch ($choice) { 
            "1" { 
                $Source = "\\eng\celogs\$global:CCID"
                Write-Host "Checking source: $Source" -ForegroundColor Yellow
                if (-not (Test-Path $Source)) {
                    Write-Host "Error: CCID '$global:CCID' not found in CELogs location." -ForegroundColor Red
                    Read-Host "Press Enter to return"
                    return
                }
                Invoke-Item $Source
                $Destination = "\\englog\escalationlogs\$global:CCID\$global:Case"
                $JobID = Read-Host -Prompt "Enter the JobID"
                $Pathchk = Test-Path $Destination
                Write-Host "Destination: $Destination, Path exists: $Pathchk" -ForegroundColor Yellow
                $validChoice = $true
            } 
            "2" { 
                $Source = "\\titan\cloudriver\celogs\$global:CCID"
                Write-Host "Checking source: $Source" -ForegroundColor Yellow
                if (-not (Test-Path $Source)) {
                    Write-Host "Error: CCID '$global:CCID' not found in Titan location." -ForegroundColor Red
                    Read-Host "Press Enter to return"
                    return 
                }
                Invoke-Item $Source
                $Destination = "\\englog\escalationlogs\$global:CCID\$global:Case"
                Write-Host "Destination: $Destination" -ForegroundColor Yellow
                Write-Host "CCID Found. Triggering ProjectTitan." -ForegroundColor Green
                ProjectTitan
                $validChoice = $true
                return 
            }
            "3" {
                $Source = "\\eng\celogs\$global:CCID"
                Write-Host "Checking source: $Source" -ForegroundColor Yellow
                if (-not (Test-Path $Source)) {
                    Write-Host "Error: CCID '$global:CCID' not found in CELogs location." -ForegroundColor Red
                    Read-Host "Press Enter to return"
                    return 
                }
                Invoke-Item $Source
                $Destination = "\\englog\escalationlogs\$global:CCID\$global:Case"
                Write-Host "Destination: $Destination" -ForegroundColor Yellow
                Write-Host "CCID Found. Extracting all files from CELogs." -ForegroundColor Green
                ExtractAllFiles $Source $Destination $Date
                $validChoice = $true
                return 
            }
            default {
                $attempts++
                Write-Host "Invalid choice. Attempts remaining: $(($maxAttempts - $attempts))" -ForegroundColor Red
                Start-Sleep -Seconds 3 
                if ($attempts -ge $maxAttempts) {
                    Write-Host "Maximum attempts reached. Exiting." -ForegroundColor Red
                    Read-Host "Press Enter to return"
                    return
                }
            }
        }
    }
    
    if ($JobID) {
        $Directory = Get-ChildItem -Recurse -Path $Source -Filter "*$JobID*" | Where-Object { $_.Extension -ne ".json" }
        Write-Host "Files found for JobID '$JobID': $($Directory.Count)" -ForegroundColor Yellow
        If ($Directory.Count -gt 0) { 
            Write-Host "Files matching JobID '$JobID' in $Source :" -ForegroundColor Green
            for ($i = 0; $i -lt $Directory.Count; $i++) {
                Write-Host "$($i + 1). $($Directory[$i].Name)" 
            } 
            Start-Sleep -Seconds 5
            $selectedFiles = $Directory 
        } else {
            Write-Host "No files found matching JobID '$JobID' in $Source." -ForegroundColor Red
            Read-Host "Press Enter to return"
            return
        }
    } else {
        $FileExists = Test-Path "\\eng\celogs\$global:CCID"
        $Pathchk = Test-Path "\\englog\escalationlogs\$global:CCID"
        $Directory = Get-ChildItem -Recurse -Path $Source
        Write-Host "Files found in $Source $($Directory.Count)" -ForegroundColor Yellow
        If (($FileExists -eq $true) -and ($Directory.count -gt 0)) { 
            Write-Host "The Commcell ID Exists" -ForegroundColor Green
            Write-Host "Files in $Source :" -ForegroundColor Green
            for ($i = 0; $i -lt $Directory.Count; $i++) {
                Write-Host "$($i + 1). $($Directory[$i].Name)" 
            }
            $selectedIndices = Read-Host "Enter the numbers of the files you want to select (comma-separated)"
            $selectedIndices = $selectedIndices -split ',' | ForEach-Object { $_.Trim() }
            $validIndices = $selectedIndices -as [int[]]
            if ($null -eq $validIndices) {
                Write-Host "Invalid input. Please enter valid numbers separated by commas." -ForegroundColor Red
                Read-Host "Press Enter to return"
                return 
            }
            $selectedFiles = @()
            foreach ($index in $validIndices) {
                if ($index -ge 1 -and $index -le $Directory.Count) {
                    $selectedFiles += $Directory[$index - 1] 
                } else {
                    Write-Host "Invalid index: $index. Skipping." -ForegroundColor Yellow
                }
            }
        } else {
            Write-Host "CCID doesn't exist or no log files found" -ForegroundColor Red
            Read-Host "Press Enter to return"
            return
        }
    }
    
    if ($selectedFiles.Count -gt 0) {
        Write-Host "Selected Files ($($selectedFiles.Count)):" -ForegroundColor Green
        foreach ($file in $selectedFiles) {
            Write-Host $file.FullName 
        }
    } else {
        Write-Host "No files selected." -ForegroundColor Yellow
        Read-Host "Press Enter to return"
        return
    }
    
    foreach ($file in $selectedFiles) {
        $Folder = $file.Basename
        $PurgePattern = ".*$global:CCID(.*?)_2024_(.*?)\.(tar\.7z|7z|tar)"
        $Trim = $Folder -replace $PurgePattern, '$1_$2'
        $extractedFolder = Join-Path "$Destination" $Trim
        Write-Host "Processing file: $($file.Name)" -ForegroundColor Yellow
        Write-Host "Target folder: $extractedFolder" -ForegroundColor Yellow
        If ($Pathchk -or $Pathchk -ne $true) { 
            Write-Host "Creating or using directory: $extractedFolder" -ForegroundColor Yellow
            New-Item -ItemType Directory -Path "$extractedFolder" -Force | Out-Null
            Write-Host "Folder Created Successfully. Extracting File now...." -ForegroundColor Green
            & 'C:\Program Files\7-Zip\7z.exe' x "$($file.FullName)" -o"$($extractedFolder)" -y
            Write-Host "Extracted $($file.Name) to $($extractedFolder)" -ForegroundColor Green
            $allFiles = Get-ChildItem -Path $extractedFolder -Recurse
            $7zFiles = $allFiles | Where-Object { $_.Name -like "*.zip" -or $_.Name -like "*.7z" -or $_.Name -like "*.7z.001"}
            $tarFiles = $allFiles | Where-Object { $_.Name -like "*.tar" -or $_.Name -like "*.tar.gz" }
            Write-Host "Found 7z files: $($7zFiles.Count), Tar files: $($tarFiles.Count)" -ForegroundColor Yellow
            if ($7zFiles.Count -ne 0 -or $tarFiles.Count -ne 0) {
                Write-Host "7z, Zip or tar files found in destination path. Triggering Extraction..." -ForegroundColor Green
                Extraction -AutomaticallyTriggered $true -Destination $extractedFolder  # Pass $extractedFolder directly
            } else {
                Write-Host "No 7z or tar files found in this destination..." -ForegroundColor Yellow
            }
        }
    }
    CheckForDMPFiles -Destination $Destination
    ii "$Destination"
}

function Extraction {
    param (
        [bool]$AutomaticallyTriggered = $false,
        [string]$Destination,  # Required parameter, no default clipboard
        [string]$extractedFolder
    )
    
    Write-Host "Extracting TAR and 7z Files..."
    if (-not $Destination) {
        if ($AutomaticallyTriggered) {
            Write-Host "No destination provided with AutomaticallyTriggered. Please specify a destination." -ForegroundColor Red
            Read-Host "Press Enter to return"
            return
        }
        $Destination = Read-Host -Prompt "Enter the destination path"
        if ([string]::IsNullOrWhiteSpace($Destination)) {
            Write-Host "No path entered, back to main screen."
            Start-Sleep -Seconds 2
            return
        }
    }
    if (-not (Test-Path -Path $Destination -PathType Container)) {
        Write-Host "The specified path $Destination does not exist."
        return
    }
    try {
        $allFiles = Get-ChildItem -Path $Destination -Recurse
        $tarFiles = $allFiles | Where-Object { $_.Name -like "*.tar" -or $_.Name -like "*.tar.gz"}
        $7zFiles = $allFiles | Where-Object { $_.Name -like "*.zip" -or $_.Name -like "*.7z" -or $_.Name -like "*.7z.001"}
        if ($tarFiles.Count -eq 0 -and $7zFiles.Count -eq 0) {
            Write-Host "No tar or 7z files found in the specified directory."
            return
        }
        foreach ($tarFile in $tarFiles) {
            $Dir = $tarFile.Basename
            $PurgePattern = ".*$global:CCID(.*?)_2024_(.*?)\.(tar\.7z|7z|tar)"
            $Purge = $Dir -replace $PurgePattern, '$1_$2'
            $extractedFolder = Join-Path -Path $Destination $Purge
            New-Item -ItemType Directory -Path "$extractedFolder" -Force
            Write-Host "Folder Created Successfully. Extracting File now...."
            & tar -xvf "$($tarFile.FullName)" -C "$($extractedFolder)"
            Write-Host "Extracted $($tarFile.Name) to $($extractedFolder)"
        }
        Get-ChildItem -Filter *.zip -Recurse "$Destination" | % { $_.FullName } | Split-Path | Get-Unique | % { cd $_ ; &'C:\Program Files\7-Zip\7z.exe' x "*.zip" -y}
        Get-ChildItem -Filter *.7z -Recurse "$Destination" | % { $_.FullName } | Split-Path | Get-Unique | % { cd $_ ; &'C:\Program Files\7-Zip\7z.exe' x "*.7z" -o* -y}
        Get-ChildItem -Filter *.7z.001 -Recurse "$Destination" | % { $_.FullName } | Split-Path | Get-Unique | % { cd $_ ; &'C:\Program Files\7-Zip\7z.exe' x "*.7z.001" -o* -y}
        Get-ChildItem -Filter *.bz2 -Recurse "$Destination" | % { $_.FullName } | Split-Path | Get-Unique | % { cd $_ ; &'C:\Program Files\7-Zip\7z.exe' x "*.bz2" -y}
        Get-ChildItem -Filter *.zip -Recurse "$Destination" | % { $_.FullName } | Split-Path | Get-Unique | % { cd $_ ; &'C:\Program Files\7-Zip\7z.exe' x "*.zip" -y}
        Get-ChildItem -Include *.bz2, *.7z, *.zip, *.001, *.002, *.003, *.004, *.005 -Recurse "$Destination" | Remove-Item
    } catch {
        Write-Host "Error occurred: $_"
        Write-Host "Re-running Extraction with the same destination path..."
        Extraction -AutomaticallyTriggered $AutomaticallyTriggered -Destination $Destination  # Use $Destination directly
    }
    Get-ChildItem -Path "$Destination" -Directory | Where-Object { $_.Name -match 'DBFiles' } | ForEach-Object { Rename-Item -Path $_.FullName -NewName "CSDB" }
}

function ExtractAllFiles {
    param ([string]$Source, [string]$Destination, [string]$Date)
    $allFiles = Get-ChildItem -Recurse -Path $Source | Where-Object { $_.Name -like "*.zip" -or $_.Name -like "*.7z" -or $_.Name -like "*.7z.001" -or $_.Name -like "*.tar" -or $_.Name -like "*.tar.gz"}
    $allDestFiles = Get-ChildItem -Recurse -Path $Destination | Where-Object { $_.Name -like "*.zip" -or $_.Name -like "*.7z" -or $_.Name -like "*.7z.001" -or $_.Name -like "*.tar" -or $_.Name -like "*.tar.gz"}
    foreach ($file in $allFiles) {
        $Folder = $file.Basename
        $PurgePattern = ".*$global:CCID(.*?)_2024_(.*?)\.(tar\.7z|7z|tar)"
        $Trim = $Folder -replace $PurgePattern, '$1_$2'
        $extractedFolder = Join-Path (Join-Path $Destination $Date) $Trim
        if (-not (Test-Path $extractedFolder)) {
            New-Item -ItemType Directory -Path $extractedFolder -Force
        }
        & 'C:\Program Files\7-Zip\7z.exe' x "$($file.FullName)" -o"$($extractedFolder)"
        if ($allDestFiles.Count -ne 0 -or $allFiles.Count -ne 0) {
            Write-Host "7z, Zip or tar files found in destination path. Triggering Extraction..." -ForegroundColor Green
            # Updated to pass destination directly
            Extraction -AutomaticallyTriggered $true -Destination $extractedFolder
        } else {
            Write-Host "No 7z or tar files found in this destination..." -ForegroundColor Yellow
        }
        ii $Destination
    }
    return
}

function ProcessDMPFiles {
    param ([System.IO.FileInfo[]]$DMPFiles, [string]$Case, [string]$Destination)

    Clear-Host
    Write-Host "Found CSDB. Processing .dmp files found in $Destination folder:" -ForegroundColor Yellow
    if ($DMPFiles.Count -gt 0) {
        Write-Host "Processing file: $($DMPFiles.FullName)" -ForegroundColor Green
        Write-Host "******************************************************"
        Write-Host "* Select an option:                                  *"
        Write-Host "* ==================                                 *"
        Write-Host "* 1. Open Staging URL for Manual Reservation         *"
        Write-Host "* 2. Ignore                                          *"
        Write-Host "******************************************************"
        $createStagingChoice = Read-Host "Enter Your Choice"
        
        if ($createStagingChoice -eq "1") {
            Write-Host "Opening staging URL for manual reservation..." -ForegroundColor Green
            Write-Host "Please follow these steps on the staging site (https://ce-staging.commvault.com/):" -ForegroundColor Yellow
            Write-Host "- Navigate to https://ce-staging.commvault.com/Reservation/Create/$Case" -ForegroundColor Yellow
            Write-Host "- Select the following options:" -ForegroundColor Yellow
            Write-Host "  1. Would you like to restore the CommServe database now? -> Yes" -ForegroundColor Yellow
            Write-Host "  2. Would you like existing jobs to be killed when the restore is completed? -> No" -ForegroundColor Yellow
            Write-Host "  3. Do you want hotfixes installed? -> Yes" -ForegroundColor Yellow
            Write-Host "  4. Is a Simpana infofile included with your CommServe data? -> No" -ForegroundColor Yellow
            Write-Host "  5. Simpana Platform: Select the template based on infofile.html in the LogExtraction folder (e.g., Simpana 11 SP36 - SQL 2022)" -ForegroundColor Yellow
            Write-Host "     - Enter the LogExtraction folder location when prompted to locate infofile.html." -ForegroundColor Yellow
            Write-Host "  6. Where is the CommServe data at this time? -> On the \\englog\escalationlogs fileshare" -ForegroundColor Yellow
            Write-Host "  7. In the file navigation window, expand CCID > Ticket Number > CSDB, and select the .dmp file (e.g., CommServ_cshcpemea01_2025_02_27_07_48_FULL.dmp)" -ForegroundColor Yellow
            Write-Host "  8. Submit the reservation." -ForegroundColor Yellow
            Write-Host "After completing these steps, press Enter to continue..." -ForegroundColor Yellow
            Open-URLInChrome -Case $Case  # Opens the create URL manually
            Read-Host "Press Enter after completing the reservation"
        } elseif ($createStagingChoice -eq "2") {
            Write-Host "Staging not requested. Continuing..." -ForegroundColor Yellow
        } else {
            Write-Host "Invalid choice. Staging not requested." -ForegroundColor Red
        }
    } else {
        Write-Host "No .dmp files found to process." -ForegroundColor Yellow
    }
}

function Open-URLInChrome {
    param ([string]$Case)

    $url = "https://ce-staging.commvault.com/Reservation/Details/US1/$Case"
    try { 
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
        if ($response.StatusCode -eq 200) {
            Write-Host "URL exists: Opening $url in Chrome." -ForegroundColor Green
            Start-Process "chrome.exe" $url -PassThru 
        } else {
            throw "URL returned status code $($response.StatusCode)" 
        }
    } catch {
        Write-Host "URL $url does not exist or could not be reached. Loading alternative URL for manual reservation..." -ForegroundColor Yellow
        $alternativeURL = "https://ce-staging.commvault.com/Reservation/Create/$Case"
        Write-Host "Please manually create a new reservation by following these steps:" -ForegroundColor Yellow
        Write-Host "- On the page ($alternativeURL), select the following options:" -ForegroundColor Yellow
        Write-Host "  1. Would you like to restore the CommServe database now? -> Yes" -ForegroundColor Yellow
        Write-Host "  2. Would you like existing jobs to be killed when the restore is completed? -> No" -ForegroundColor Yellow
        Write-Host "  3. Do you want hotfixes installed? -> Yes" -ForegroundColor Yellow
        Write-Host "  4. Is a Simpana infofile included with your CommServe data? -> No" -ForegroundColor Yellow
        Write-Host "  5. Simpana Platform: Select the template based on infofile.html in the LogExtraction folder (e.g., Simpana 11 SP36 - SQL 2022)" -ForegroundColor Yellow
        Write-Host "     - Enter the LogExtraction folder location when prompted to locate infofile.html." -ForegroundColor Yellow
        Write-Host "  6. Where is the CommServe data at this time? -> On the \\englog\escalationlogs fileshare" -ForegroundColor Yellow
        Write-Host "  7. In the file navigation window, expand CCID > Ticket Number > CSDB, and select the .dmp file (e.g., CommServ_cshcpemea01_2025_02_27_07_48_FULL.dmp)" -ForegroundColor Yellow
        Write-Host "  8. Submit the reservation." -ForegroundColor Yellow
        Write-Host "After completing these steps, press Enter to continue..." -ForegroundColor Yellow
        Start-Process "chrome.exe" $alternativeURL -PassThru 
        Read-Host "Press Enter after completing the reservation"
    } 
    return 
}

function ProjectTitan {
    param ([int]$MaxParallelJobs = 5)
    Write-Host "+++ ProjectTitan has been triggered +++" -ForegroundColor Green
    Write-Host "Parent folders in Titan location:"
    $TitanParentFolders = Get-ChildItem -Directory $Source
    for ($i = 0; $i -lt $TitanParentFolders.Count; $i++) {
        Write-Host "$($i + 1). $($TitanParentFolders[$i])"    
    }
    $selectedFolderIndex = Read-Host "Enter the number of the parent folder you want to select"
    $selectedFolderIndex = [int]$selectedFolderIndex
    if ($selectedFolderIndex -ge 1 -and $selectedFolderIndex -le $TitanParentFolders.Count) {
        $selectedFolderName = $TitanParentFolders[$selectedFolderIndex - 1]
        $Source = Join-Path $Source $selectedFolderName
        Write-Host "Selected folder: $selectedFolderName" -ForegroundColor Yellow
        Write-Host "Folders within the selected parent folder:"
        $SubFolders = Get-ChildItem -Directory $Source
        for ($i = 0; $i -lt $SubFolders.Count; $i++) {
            Write-Host "$($i + 1). $($SubFolders[$i])"        
        }
        $selectedSubFolderIndex = Read-Host "Enter the number of the folder you want to select"
        $selectedSubFolderIndex = [int]$selectedSubFolderIndex
        if ($selectedSubFolderIndex -ge 1 -and $selectedSubFolderIndex -le $SubFolders.Count) {
            $selectedSubFolderName = $SubFolders[$selectedSubFolderIndex - 1]
            $Source = Join-Path $Source $selectedSubFolderName
            Write-Host "Selected folder: $selectedSubFolderName" -ForegroundColor Yellow        
        } else {
            Write-Host "Invalid selection. Exiting script." -ForegroundColor Red
            return        
        }
        if ($selectedSubFolderName -eq "Compressed") {
            Write-Host "7z files within 'Compressed' folder:" -ForegroundColor Yellow
            $Files = Get-ChildItem -Path $Source -Filter *.7z
            for ($i = 0; $i -lt $Files.Count; $i++) {
                Write-Host "$($i + 1). $($Files[$i].Name)"            
            }
            $selected7zIndex = Read-Host "Enter the number of the .7z file you want to select"
            if ($selected7zIndex -ge 1 -and $selected7zIndex -le $Files.Count) {
                $selected7zFile = $Files[$selected7zIndex - 1]
                Write-Host "Selected .7z file: $($selected7zFile.Name)" -ForegroundColor Yellow
            } else {
                Write-Host "Invalid selection. Exiting script." -ForegroundColor Red
                return            
            }        
        }
        if ($selectedSubFolderName -ne "Uncompressed") {
            $parentFolderName = (Get-Item $Source).Parent.Name
            $destinationParentFolder = Join-Path -Path $Destination -ChildPath $parentFolderName
            if (!(Test-Path -Path $destinationParentFolder)) {
                New-Item -ItemType Directory -Path $destinationParentFolder -Force | Out-Null            
            }
            $sourceItems = Get-ChildItem -Path $Source
            $jobs = @()
            foreach ($item in $sourceItems) {
                if ($jobs.Count -ge $MaxParallelJobs) {
                    $completedJob = Wait-Job -Job $jobs -Any
                    $jobs = $jobs | Where-Object { $_.State -eq 'Running' }
                    Receive-Job -Job $completedJob | Out-Null
                    Remove-Job -Job $completedJob                
                }
                $job = Start-Job -ScriptBlock {
                    param ($source, $destination)
                    Copy-Item -Path $source -Destination $destination -Recurse -Force -Verbose
                } -ArgumentList "$($item.FullName)", "$destinationParentFolder\$($item.Name)"
                $jobs += $job            
            }
            $jobs | ForEach-Object {
                Wait-Job -Job $_
                Receive-Job -Job $_ | Out-Null
                Remove-Job -Job $_            
            }
            Write-Host "Uncompressed folder and its contents copied to $destinationParentFolder" -ForegroundColor Green
            # Updated to pass destination directly
            Extraction -AutomaticallyTriggered $true -Destination $destinationParentFolder
            ii $destinationParentFolder
            SearchLogs -Destination $Destination
            return        
        }
    } else {
        Write-Host "Invalid selection. Exiting script." -ForegroundColor Red
        return    
    }
}

function CheckForDMPFiles {
    param (
        [string]$Destination,
        [bool]$AutomaticallyTriggered = $false
    )
    if ($AutomaticallyTriggered) {
        $Destination = Get-Clipboard 
    } else {
        if (-not $Destination) {
            $global:CCID = Read-Host -Prompt "Enter CCID"
            $global:Case = Read-Host -Prompt "Enter Ticket Number"
            $Destination = "\\englog\escalationlogs\$global:CCID\$global:Case"
        }
    }
    $dmpFiles = Get-ChildItem -Path $Destination -Recurse -Filter *.dmp
    if ($dmpFiles.Count -gt 0) {
        Write-Host ".dmp files found in $Destination or its child folders. Proceeding..."
        ProcessDMPFiles -DMPFiles $dmpFiles -Case $global:Case -Destination $Destination 
    }
}

function SearchLogs {
    param ([bool]$AutomaticallyTriggered = $false, [string]$Destination)

    try {
        Clear-Host
        Write-Host "🔍 Search Logs for Job ID" -ForegroundColor Cyan
        Write-Host "========================" -ForegroundColor Cyan

        # Set default destination if not provided
        if (-not $Destination) {
            $defaultDestination = "\\englog\escalationlogs\$global:CCID\$global:Case"
            if (-not (Test-Path $defaultDestination)) {
                Write-Host "Default log location $defaultDestination not found." -ForegroundColor Yellow
                $manualDestination = Read-Host "Please enter the manual log location (e.g., path to log files)"
                if (-not [string]::IsNullOrWhiteSpace($manualDestination)) {
                    $Destination = $manualDestination
                    if (-not (Test-Path $Destination)) {
                        Write-Host "Manual log location $Destination not found. Please verify the path." -ForegroundColor Red
                        Read-Host "Press Enter to return to menu"
                        return
                    }
                } else {
                    Write-Host "No location entered. Exiting." -ForegroundColor Red
                    Read-Host "Press Enter to return to menu"
                    return
                }
            } else {
                $Destination = $defaultDestination
            }
        }

        if ($AutomaticallyTriggered) {
            $Destination = Get-Clipboard 
        }

        $JobID = Read-Host -Prompt "Enter the Job ID to search logs"
        if (-not $JobID) {
            Write-Host "No Job ID entered. Exiting." -ForegroundColor Red
            return
        }
        Write-Host "🔍 Searching for Job ID '$JobID' in extracted logs..."
        $OutputFile = Join-Path -Path $Destination -ChildPath "$JobID.txt"
        $JobFolder = Join-Path -Path $Destination -ChildPath $JobID
        New-Item -ItemType Directory -Path $JobFolder -Force | Out-Null
        
        $foundLogs = @{}
        $logFiles = Get-ChildItem -Path $Destination -Recurse -Filter "*.log"
        $txtFiles = Get-ChildItem -Path $Destination -Recurse -Filter "*.txt"
        $allFiles = $logFiles + $txtFiles
        
        foreach ($file in $allFiles) {
            $filePath = $file.FullName
            $matchingLines = Select-String -Path $filePath -Pattern $JobID | ForEach-Object { $_.Line }
            if ($matchingLines.Count -gt 0) {
                $foundLogs[$filePath] = $matchingLines
                $logCutFile = Join-Path -Path $JobFolder -ChildPath "$($file.BaseName)_logcut.txt"
                "Log cut from: $filePath`n" + ($matchingLines -join "`n") | Set-Content -Path $logCutFile
            }
        }
        
        if ($foundLogs.Count -gt 0) {
            $outputContent = @()
            foreach ($file in $foundLogs.Keys) {
                $outputContent += "`n$file"
                $outputContent += $foundLogs[$file]
            }
            $outputContent | Set-Content -Path $OutputFile
            Write-Host "Job ID '$JobID' found in logs. Results saved to $OutputFile"
            LogAnalysis -FilePath $OutputFile -Destination $Destination
        } else {
            Write-Host "No matching logs found for Job ID '$JobID'."
        }
        Write-Host "`nPress Enter to continue..."
        Read-Host
    } catch {
        Write-Host "Error in SearchLogs: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}

function LogAnalysis {
    param (
        [string]$FilePath,
        [string]$Destination
    )

    try {
        if (-not (Test-Path $FilePath)) {
            Write-Host "Log file $FilePath not found. Skipping analysis." -ForegroundColor Red
            return
        }

        $errorReportPath = Join-Path -Path $Destination -ChildPath "ErrorReport_$((Get-Item $FilePath).BaseName).txt"
        
        # Use global keywords with fallbacks
        $errorPatterns = if ($global:ErrorKeywords) { $global:ErrorKeywords } else { @("error", "fail", "exception") }
        $errorPatternRegex = ($errorPatterns | ForEach-Object { [regex]::Escape($_) }) -join "|"
        
        $errorsFound = Select-String -Path $FilePath -Pattern $errorPatternRegex -CaseSensitive:$false
        
        if ($errorsFound.Count -gt 0) {
            # Build a formatted report with each error on separate lines
            $reportContent = "Errors found in ${FilePath}:`n"
            foreach ($error in $errorsFound) {
                $reportContent += "Log Location: $($error.Filename)`n"
                $reportContent += "Log Line: $($error.Line)`n"
                $reportContent += "------------------------`n"
            }
            $reportContent | Set-Content -Path $errorReportPath
            Write-Host "Errors detected. Report saved to $errorReportPath" -ForegroundColor Green
            CheckReferenceLog -ErrorReportPath $errorReportPath -Destination $Destination
        } else {
            Write-Host "No errors found in $FilePath" -ForegroundColor Green
        }
    } catch {
        Write-Host "Error in LogAnalysis: $_" -ForegroundColor Red
    }
}

function CreateKBArticle {
    param (
        [string]$Destination = "$env:USERPROFILE\Desktop\EcoDesk\KBArticles"
    )

    try {
        Clear-Host
        Write-Host "Creating KB Article..." -ForegroundColor Cyan

        # Auto-populate data from latest summary file
        $summaryPath = "$env:USERPROFILE\Desktop\EcoDesk\Solution Tree\$global:Case"
        $latestSummary = Get-ChildItem -Path $summaryPath -File | 
                        Where-Object { $_.Name -match "Date_\d+\.txt" } | 
                        Sort-Object Name -Descending | 
                        Select-Object -First 1
        
        # Default values if no summary exists
        $issueDescription = "Not found in summary file"
        $causes = "Not specified"
        $findings = "Not specified"
        $logCuts = "No log cuts available"
        $resolutions = "Not resolved"
        $agentType = "Unknown"
        $issueType = "Unknown"

        if ($latestSummary) {
            $content = Get-Content -Path $latestSummary.FullName -Raw
            # Issue Description from summary
            if ($content -match "Issue: (.+?)(?=\r?\n\r?\n)") {
                $issueDescription = $matches[1]
            }
            # Causes (from Notes or Steps Taken, assuming Notes might hint at causes)
            if ($content -match "Notes:.*?------+(.+?)(?=\r?\n\r?\nResolution:)" -replace "`r`n", "`n") {
                $causes = $matches[1].Trim()
            }
            # Findings (from Steps Taken)
            if ($content -match "Steps Taken:.*?------+(.+?)(?=\r?\n\r?\nNotes:)" -replace "`r`n", "`n") {
                $findings = $matches[1].Trim()
            }
            # Resolutions (from Resolution section)
            if ($content -match "Resolution:.*?-----------+(.+)" -replace "`r`n", "`n") {
                $resolutions = $matches[1].Trim()
            }
            # Agent Type (from Agent line)
            if ($content -match "Agent: (.+?)(?=\r?\n)") {
                $agentType = $matches[1]
            }
        }

        # Auto-populate Log Cuts from SearchLogs output if available
        $logPath = "\\englog\escalationlogs\$global:CCID\$global:Case"
        $logCutFile = Get-ChildItem -Path $logPath -Filter "*_logcut.txt" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($logCutFile) {
            $logCuts = Get-Content -Path $logCutFile.FullName -Raw
        }

        # Determine Issue Type (simple heuristic: check if "backup" or "restore" is in issue description)
        if ($issueDescription -match "backup") { $issueType = "Backup" }
        elseif ($issueDescription -match "restore") { $issueType = "Restore" }

        # Check if KB already exists
        $kbFileName = "KB_$global:Case_$agentType_$issueType"
        $kbTxtPath = Join-Path -Path $Destination -ChildPath "$kbFileName.txt"
        $kbHtmlPath = Join-Path -Path $Destination -ChildPath "$kbFileName.html"
        if ((Test-Path $kbTxtPath) -or (Test-Path $kbHtmlPath)) {
            Write-Host "KB Article for Case $global:Case already exists at $Destination" -ForegroundColor Yellow
            Read-Host "Press Enter to return to menu"
            return
        }

        # Display and allow edits
        Write-Host "`nKB Article Preview (Edit if needed):" -ForegroundColor Cyan
        Write-Host "1. Issue Description: " -NoNewline -ForegroundColor Yellow
        Write-Host $issueDescription
        $edit = Read-Host "Edit? (Y/N)"
        if ($edit.ToUpper() -eq "Y") { $issueDescription = Read-Host "Enter new Issue Description" }

        Write-Host "2. Causes: " -NoNewline -ForegroundColor Yellow
        Write-Host $causes
        $edit = Read-Host "Edit? (Y/N)"
        if ($edit.ToUpper() -eq "Y") { $causes = Read-Host "Enter new Causes" }

        Write-Host "3. Findings: " -NoNewline -ForegroundColor Yellow
        Write-Host $findings
        $edit = Read-Host "Edit? (Y/N)"
        if ($edit.ToUpper() -eq "Y") { $findings = Read-Host "Enter new Findings" }

        Write-Host "4. Log Cuts: " -NoNewline -ForegroundColor Yellow
        Write-Host $logCuts
        $edit = Read-Host "Edit? (Y/N)"
        if ($edit.ToUpper() -eq "Y") { $logCuts = Read-Host "Enter new Log Cuts" }

        Write-Host "5. Resolutions: " -NoNewline -ForegroundColor Yellow
        Write-Host $resolutions
        $edit = Read-Host "Edit? (Y/N)"
        if ($edit.ToUpper() -eq "Y") { $resolutions = Read-Host "Enter new Resolutions" }

        # Choose format
        Write-Host "`nSave as:" -ForegroundColor Cyan
        Write-Host "1. Plain Text (.txt)"
        Write-Host "2. HTML (.html)"
        $formatChoice = Read-Host "Enter choice (1 or 2)"
        
        # Create directory if it doesn’t exist
        if (-not (Test-Path $Destination)) {
            New-Item -ItemType Directory -Path $Destination -Force | Out-Null
        }

        # Prepare content
        $kbContent = @"
Issue Description:
------------------
$issueDescription

Causes:
-------
$causes

Findings:
---------
$findings

Log Cuts:
---------
$logCuts

Resolutions:
------------
$resolutions
"@

        if ($formatChoice -eq "2") {
            # HTML format with color coding
            $htmlContent = @"
<html>
<head>
    <title>KB Article - Case $global:Case</title>
    <style>
        body { font-family: Arial, sans-serif; }
        h2 { color: #0066cc; }
        .section { margin-bottom: 20px; }
        .label { color: #ff9900; font-weight: bold; }
    </style>
</head>
<body>
    <h2>KB Article - Case $global:Case</h2>
    <div class='section'>
        <span class='label'>Issue Description:</span><br>
        $issueDescription
    </div>
    <div class='section'>
        <span class='label'>Causes:</span><br>
        $causes
    </div>
    <div class='section'>
        <span class='label'>Findings:</span><br>
        $findings
    </div>
    <div class='section'>
        <span class='label'>Log Cuts:</span><br>
        <pre>$logCuts</pre>
    </div>
    <div class='section'>
        <span class='label'>Resolutions:</span><br>
        $resolutions
    </div>
</body>
</html>
"@
            $htmlContent | Set-Content -Path $kbHtmlPath
            Write-Host "KB Article saved as HTML to $kbHtmlPath" -ForegroundColor Green
            Start-Process $kbHtmlPath  # Open in default browser
        } else {
            # Plain text format
            $kbContent | Set-Content -Path $kbTxtPath
            Write-Host "KB Article saved as text to $kbTxtPath" -ForegroundColor Green
            Start-Process notepad.exe $kbTxtPath  # Open in Notepad
        }

        Read-Host "Press Enter to return to menu"
    } catch {
        Write-Host "Error creating KB Article: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}

function CheckReferenceLog {
    param (
        [string]$ErrorReportPath,
        [string]$Destination = "\\englog\escalationlogs\$global:CCID\$global:Case"
    )

    $refFolder = "\\englog\escalationlogs\references"
    if (-not (Test-Path $refFolder)) {
        Write-Host "No reference logs available at $refFolder." -ForegroundColor Yellow
        return
    }
    
    try {
        $errorLines = Get-Content -Path $ErrorReportPath -ErrorAction Stop
        if (-not $errorLines -or $errorLines.Count -eq 0) {
            Write-Host "No error lines found in $ErrorReportPath to match." -ForegroundColor Yellow
            return
        }

        $refFiles = Get-ChildItem -Path $refFolder -Filter "*.txt" -ErrorAction Stop
        if ($refFiles.Count -eq 0) {
            Write-Host "No reference log files found in $refFolder." -ForegroundColor Yellow
            return
        }

        foreach ($refFile in $refFiles) {
            try {
                $refContent = Get-Content -Path $refFile.FullName -ErrorAction Stop
                $refLogCutMatch = $refContent | Select-String -Pattern "Log Cut:([\s\S]*?)Solution:" -AllMatches
                $solutionMatch = $refContent | Select-String -Pattern "Solution: (.*)$" -AllMatches

                if ($refLogCutMatch.Matches.Count -gt 0 -and $solutionMatch.Matches.Count -gt 0) {
                    $refLogCut = $refLogCutMatch.Matches[0].Groups[1].Value.Trim()
                    $solution = $solutionMatch.Matches[0].Groups[1].Value.Trim()
                    
                    $refLogCutWords = $refLogCut -split '\s+' | Where-Object { $_ }
                    $matchThreshold = 0.75

                    foreach ($errorLine in $errorLines) {
                        $errorLineWords = $errorLine -split '\s+' | Where-Object { $_ }
                        $commonWords = $refLogCutWords | Where-Object { $errorLineWords -contains $_ }
                        $matchPercentage = $commonWords.Count / [Math]::Max($refLogCutWords.Count, $errorLineWords.Count)

                        if ($matchPercentage -ge $matchThreshold -and $commonWords.Count -ge 3) {
                            Write-Host "Match found in reference log: $refFile (Match: $($matchPercentage*100)%)" -ForegroundColor Green
                            Write-Host "Reference Log Cut: $refLogCut" -ForegroundColor Yellow
                            Write-Host "Matching Error Line: $errorLine" -ForegroundColor Yellow
                            Write-Host "Suggested Solution: $solution" -ForegroundColor Yellow

                            # Simulate saving/updating reference log
                            $saveRef = Read-Host "Save this as a new reference log? (Y/N)"
                            if ($saveRef.ToUpper() -eq "Y") {
                                $refFileName = "RefLog_$((Get-Date).ToString('yyyyMMddHHmmss')).txt"
                                $refFilePath = Join-Path -Path $refFolder -ChildPath $refFileName
                                $maskedRefLogCut = Mask-PII -LogCut $refLogCut  # Mask PII here
                                "Log Cut:`n$maskedRefLogCut`nSolution:`n$solution" | Set-Content -Path $refFilePath
                                Write-Host "Reference log saved with PII masked to $refFilePath" -ForegroundColor Green
                            }
                            return
                        }
                    }
                }
            } catch {
                Write-Host "Error reading reference file $refFile $_" -ForegroundColor Red
                continue
            }
        }
        Write-Host "No matching reference log found." -ForegroundColor Yellow

        # Option to create a new reference log if no match
        $createNew = Read-Host "No match found. Create a new reference log? (Y/N)"
        if ($createNew.ToUpper() -eq "Y") {
            $newLogCut = Read-Host "Enter the log cut to save"
            $newSolution = Read-Host "Enter the solution"
            $refFileName = "RefLog_$((Get-Date).ToString('yyyyMMddHHmmss')).txt"
            $refFilePath = Join-Path -Path $refFolder -ChildPath $refFileName
            $maskedNewLogCut = Mask-PII -LogCut $newLogCut  # Mask PII here
            "Log Cut:`n$maskedNewLogCut`nSolution:`n$newSolution" | Set-Content -Path $refFilePath
            Write-Host "New reference log saved with PII masked to $refFilePath" -ForegroundColor Green
        }
    } catch {
        Write-Host "Error in CheckReferenceLog: $_" -ForegroundColor Red
    }
}

function SolutionTree {
    param (
        [string]$Destination = "$env:USERPROFILE\Desktop\EcoDesk\Solution Tree\$global:Case"
    )

    # Load custom steps if they exist
    $stepsFile = "$env:USERPROFILE\Desktop\EcoDesk\Config\CustomSteps.json"
    $customSteps = if (Test-Path $stepsFile) {
        Get-Content -Path $stepsFile -Raw | ConvertFrom-Json
    } else {
        @{}
    }

    # Define Focus Groups and their Agents (unchanged)
    $focusGroups = @{
        "Messaging" = @{
            "SQL Server" = @("SQL Server", "Azure SQL", "AWS SQL", "Availability Group", "SQL Cluster")
            "Exchange" = @("Mailbox On-Prem", "Mailbox Online", "Exchange Database", "Exchange DAG", "Journaling", "Content Store SMTP")
            "Sharepoint" = @("Sharepoint On-Prem", "Sharepoint Online")
            "All Other" = @("Cosmos DB", "Onedrive", "Teams", "GitHub", "GitLab", "Lotus Notes", "Dynamo DB", "Cockroach DB", "Splunk", "Data Insight")
        }
        "Media Management" = @{
            "General" = @("Worm Copy", "Auxiliary Copy", "Content Indexing and Search", "DASH full (Synthetic Full)", "Data Verification", "Deduplication", "Gridstor", "Index Cache", "Library and Drive Configuration", "Media Agent Server", "Media Refresh", "Media Explorer", "Multiplexing", "NAS NDMP Agents", "NFS ObjectStore / Hybrid File Store", "Vault Tracker", "Edge", "Data Analytics", "Catalog Media")
            "Hyperscale" = @("Hyperscale 1.5 (Gluster)", "Hyperscale X (CDS)", "Remote Office Appliance")
        }
        "Server" = @("MongoDB Security Issue", "Worm Copy", "cyberark credential access integration", "Alerts", "Cloud Services Website", "CommCell Console", "Command Line Interface/QCommand (CLI)", "CommCell Merge", "CommServe Database", "CommServe DR Process", "CommServe Install/Upgrade", "CommServe Services and Performance", "Configuration", "Data Aging", "Downloads and Updates", "Encryption", "End of Life", "Job Controller Management and Job History", "Licensing", "Log File Upload", "Policies and Groups", "Reports", "User Security (CommVault level)", "Schedules", "Security Concerns", "T2 Training", "Workflows (graphical method to chain multiple commands in sequence)", "Error Deploying V11 Service Pack 11", "I need to access my CommServe Database")
        "Client" = @("Windows App Aware snap mounts take a long time or seem to be hung", "Application Aware", "Error Code [72:106] Caught Access Violation Exception during LVM processing", "Microsoft Azure Cosmos DB", "Kubernetes", "Metadata Collection Fails for Virtual Machine Backups", "CVVD.SYS driver may cause reboot of machine due to issue when using Live Browse or Block Level restore", "CVE-2021-4034 PKEXEC exploit - Actions to take", "Missing DR VMs within the Cloud after Upgrade to CPR2023E", "Amazon Client (Lifecycle and EC2 Instance Backup)", "Hyper-V", "Microsoft Azure", "Google Cloud Platform", "Microsoft Azure Stack", "Docker", "Huawei FusionCompute", "Nutanix Acropolis", "Oracle Cloud Infrastructure", "Oracle VM", "Red Hat Enterprise Virtualization", "OpenStack", "XenServer", "VMware", "v2 Indexing", "How do I configure Metadata Collection/Granular Recovery for VSA within SP11 and on?", "Failed to create the backup statistics information object from a database record")
        "Unix" = @("Amazon Redshift", "DynamoDB", "Cassandra", "DB2", "DB2 MultiNode", "Greenplum", "HANA", "HBase", "Hive", "Informix", "MaxDB", "Microsoft SQL Server", "MongoDB", "MySQL/MariaDB", "Oracle", "Oracle RAC", "PostgreSQL", "SAP Archive Link", "SAP Oracle", "SAP Landscape Management Integration", "Sybase")
        "Metallic" = @("All subscriptions", "Microsoft 365", "Risk Analysis for Microsoft 365", "Integration with Microsoft 365 Backup Storage", "Endpoints", "Files & Objects", "Archive for Files & Objects", "Threat Scan for File & Object", "VMs", "Auto Recovery for VM", "Kubernetes", "Databases", "Google Workspace", "Salesforce", "Microsoft Dynamics 365", "Active Directory", "Migration as a Service", "Threatwise", "Security IQ", "Cleanroom", "Cloud Rewind", "Clumio", "CV Cloud SaaS & Hyperscale X", "Air Gap Protect", "Lighthouse", "Onboarding Tracking (SoftwareONE ONLY)")
    }

    # Define Issue Types
    $issueTypes = @("Backup", "Restore", "Installation", "Discovery", "Browse")

    # Versions in descending order
    $commvaultVersions = @("FR40", "FR36", "FR32", "FR28", "FR24")

    try {
        Clear-Host
        # Select Focus Group
        Write-Host "Select Focus Group:" -ForegroundColor Cyan
        $focusGroupKeys = $focusGroups.Keys | Sort-Object
        for ($i = 0; $i -lt $focusGroupKeys.Count; $i++) {
            Write-Host "$($i+1). $($focusGroupKeys[$i])"
        }
        $focusChoice = [int](Read-Host "Enter number (1-$($focusGroupKeys.Count))") - 1
        if ($focusChoice -lt 0 -or $focusChoice -ge $focusGroupKeys.Count) { throw "Invalid focus group selection" }
        $selectedFocus = $focusGroupKeys[$focusChoice]

        # Select Agent/Subcategory
        Clear-Host
        Write-Host "Selected Focus Group: $selectedFocus" -ForegroundColor Cyan
        $agents = $focusGroups[$selectedFocus]
        if ($agents -is [hashtable]) {
            Write-Host "Select Category:" -ForegroundColor Cyan
            $categoryKeys = $agents.Keys | Sort-Object
            for ($i = 0; $i -lt $categoryKeys.Count; $i++) {
                Write-Host "$($i+1). $($categoryKeys[$i])"
            }
            $categoryChoice = [int](Read-Host "Enter number (1-$($categoryKeys.Count))") - 1
            if ($categoryChoice -lt 0 -or $categoryChoice -ge $categoryKeys.Count) { throw "Invalid category selection" }
            $selectedCategory = $categoryKeys[$categoryChoice]
            $agentList = $agents[$selectedCategory]
        } else {
            $agentList = $agents
        }

        Write-Host "`nSelect Agent:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $agentList.Count; $i++) {
            Write-Host "$($i+1). $($agentList[$i])"
        }
        $agentChoice = [int](Read-Host "Enter number (1-$($agentList.Count))") - 1
        if ($agentChoice -lt 0 -or $agentChoice -ge $agentList.Count) { throw "Invalid agent selection" }
        $selectedAgent = $agentList[$agentChoice]

        # New: Select Issue Type
        Clear-Host
        Write-Host "Selected: $selectedFocus > $selectedAgent" -ForegroundColor Cyan
        Write-Host "Select Issue Type:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $issueTypes.Count; $i++) {
            Write-Host "$($i+1). $($issueTypes[$i])"
        }
        $issueTypeChoice = [int](Read-Host "Enter number (1-$($issueTypes.Count))") - 1
        if ($issueTypeChoice -lt 0 -or $issueTypeChoice -ge $issueTypes.Count) { throw "Invalid issue type selection" }
        $issueType = $issueTypes[$issueTypeChoice]

        # Select Commvault Version
        Clear-Host
        Write-Host "Selected: $selectedFocus > $selectedAgent ($issueType)" -ForegroundColor Cyan
        Write-Host "Select Commvault Version:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $commvaultVersions.Count; $i++) {
            Write-Host "$($i+1). $($commvaultVersions[$i])"
        }
        $versionChoice = [int](Read-Host "Enter number (1-$($commvaultVersions.Count))") - 1
        if ($versionChoice -lt 0 -or $versionChoice -ge $commvaultVersions.Count) { throw "Invalid version selection" }
        $selectedVersion = $commvaultVersions[$versionChoice]

        # Enter Issue Summary
        Clear-Host
        Write-Host "Selected: $selectedFocus > $selectedAgent ($issueType, $selectedVersion)" -ForegroundColor Cyan
        $issueSummary = Read-Host "Enter Issue Summary"

        # Update stepKey to include Issue Type
        $stepKey = "$selectedFocus|$selectedAgent|$issueType|$selectedVersion"

        # Load troubleshooting steps (custom or default)
        $troubleshootingSteps = if ($customSteps.PSObject.Properties.Name -contains $stepKey) {
            $customSteps.$stepKey
        } else {
            @(
                "Verify service status and restart if necessary",
                "Check log files for specific error messages",
                "Validate network connectivity to affected component",
                "Confirm user permissions and access rights",
                "Review configuration settings",
                "Check disk space and resource availability",
                "Update to latest patches if applicable",
                "Run diagnostic tools specific to the agent",
                "Verify backup/restore job logs",
                "Consult Commvault documentation for known issues",
                "Escalate to Tier 2 support"
            )
        }

        $currentStep = 0
        $notes = @()
        $kbContent = ""

        while ($currentStep -lt $troubleshootingSteps.Count) {
            Clear-Host
            Write-Host "Troubleshooting: $selectedAgent ($issueType, $selectedVersion)" -ForegroundColor Green
            Write-Host "Issue: $issueSummary" -ForegroundColor Green
            Write-Host "Step $($currentStep + 1): $($troubleshootingSteps[$currentStep])" -ForegroundColor Yellow
            Write-Host "`nOptions:" -ForegroundColor Cyan
            Write-Host "1: Issue persists (Next Step)"
            Write-Host "2: Issue resolved"
            Write-Host "3: Check KB Article"
            Write-Host "4: Back to Menu"
            if ($currentStep -eq 10) { 
                Write-Host "5: Verify KB Article for Tier 2 Escalation"
                Write-Host "6: Verify KB Article for Dev Team Escalation"
            }

            $choice = Read-Host "Enter choice"
            switch ($choice) {
                "1" { 
                    if ($currentStep -lt 10) { $currentStep++ }
                    else { Write-Host "All steps exhausted, consider escalation" -ForegroundColor Yellow; Start-Sleep -Seconds 2 }
                }
                "2" { 
                    $resolution = Read-Host "Enter resolution details (optional)"
                    AutoSummary -Focus $selectedFocus -Agent $selectedAgent -IssueType $issueType -Version $selectedVersion `
                               -Issue $issueSummary -Steps $troubleshootingSteps[0..$currentStep] `
                               -Notes $notes -Resolution $resolution -OpenFile $true
                    return 
                }
                "3" { 
                    $kbResult = CheckKBArticle -Agent $selectedAgent -IssueType $issueType -Issue $issueSummary
                    if ($kbResult -eq "Not Found") {
                        Write-Host "No KB article exists. Would you like to create one? (Y/N)" -ForegroundColor Yellow
                        if ((Read-Host).ToUpper() -eq "Y") {
                            $kbContent = Read-Host "Enter KB article content"
                            SaveKBArticle -Agent $selectedAgent -IssueType $issueType -Issue $issueSummary -Content $kbContent
                            Write-Host "KB Article saved" -ForegroundColor Green
                            Start-Sleep -Seconds 1
                        } else {
                            EscalateToTier2
                            return
                        }
                    } else {
                        Write-Host "KB Article found: $kbResult" -ForegroundColor Green
                        Read-Host "Press Enter to continue"
                    }
                }
                "4" { 
                    AutoSummary -Focus $selectedFocus -Agent $selectedAgent -IssueType $issueType -Version $selectedVersion `
                               -Issue $issueSummary -Steps $troubleshootingSteps[0..$currentStep] `
                               -Notes $notes -OpenFile $false
                    return 
                }
                "5" { 
                    if ($currentStep -eq 10) {
                        $kbResult = CheckKBArticle -Agent $selectedAgent -IssueType $issueType -Issue $issueSummary
                        if ($kbResult -eq "Not Found") {
                            Write-Host "No KB article exists. Would you like to create one? (Y/N)" -ForegroundColor Yellow
                            if ((Read-Host).ToUpper() -eq "Y") {
                                $kbContent = Read-Host "Enter KB article content"
                                SaveKBArticle -Agent $selectedAgent -IssueType $issueType -Issue $issueSummary -Content $kbContent
                                Write-Host "KB Article saved" -ForegroundColor Green
                                Start-Sleep -Seconds 1
                            }
                        } else {
                            Write-Host "KB Article found: $kbResult" -ForegroundColor Green
                            Read-Host "Press Enter to continue"
                        }
                        EscalateToTier2
                        return
                    }
                }
                "6" { 
                    if ($currentStep -eq 10) {
                        $kbResult = CheckKBArticle -Agent $selectedAgent -IssueType $issueType -Issue $issueSummary
                        if ($kbResult -eq "Not Found") {
                            Write-Host "No KB article exists. Would you like to create one? (Y/N)" -ForegroundColor Yellow
                            if ((Read-Host).ToUpper() -eq "Y") {
                                $kbContent = Read-Host "Enter KB article content"
                                SaveKBArticle -Agent $selectedAgent -IssueType $issueType -Issue $issueSummary -Content $kbContent
                                Write-Host "KB Article saved" -ForegroundColor Green
                                Start-Sleep -Seconds 1
                            }
                        } else {
                            Write-Host "KB Article found: $kbResult" -ForegroundColor Green
                            Read-Host "Press Enter to continue"
                        }
                        EscalateToDevTeam
                        return
                    }
                }
                default { Write-Host "Invalid choice" -ForegroundColor Red; Start-Sleep -Seconds 1 }
            }
        }
    } catch {
        Write-Host "Error in SolutionTree: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}

function AutoSummary {
    param (
        [string]$Focus,
        [string]$Agent,
        [string]$Version,
        [string]$Issue,
        [string[]]$Steps,
        [string[]]$Notes,
        [string]$Resolution = "Not resolved",
        [bool]$OpenFile = $false
    )

    $basePath = "$env:USERPROFILE\Desktop\EcoDesk\Solution Tree\$global:Case"
    if (-not (Test-Path $basePath)) {
        New-Item -ItemType Directory -Path $basePath -Force | Out-Null
    }

    # Use today's date for the filename (e.g., 2025-02-28.txt)
    $date = Get-Date -Format "yyyy-MM-dd"
    $baseFileName = "$date"
    $summaryFile = Join-Path -Path $basePath -ChildPath "$baseFileName.txt"

    # Check if a file for today already exists; if so, append a timestamp (e.g., 2025-02-28_143022.txt)
    $counter = 1
    while (Test-Path $summaryFile) {
        $summaryFile = Join-Path -Path $basePath -ChildPath "$baseFileName_$([string]::Format("{0:HHmmss}", (Get-Date))).txt"
        $counter++
        if ($counter -gt 100) {  # Prevent infinite loop
            Write-Host "Warning: Too many files for today. Using timestamp only." -ForegroundColor Yellow
            $summaryFile = Join-Path -Path $basePath -ChildPath "$([string]::Format("{0:yyyy-MM-dd_HHmmss}", (Get-Date))).txt"
            break
        }
    }

    $notesContent = if ($Notes.Count -eq 0) { "No notes added" } else { $Notes -join "`n" }

    $content = @"
Troubleshooting Summary
======================
Focus Group: $Focus
Agent: $Agent
Commvault Version: $Version
Issue: $Issue
Date: $(Get-Date)

Steps Taken:
------------
$($Steps -join "`n")

Notes:
------
$notesContent

Resolution:
-----------
$Resolution
"@

    $content | Set-Content -Path $summaryFile
    Write-Host "Summary saved to $summaryFile" -ForegroundColor Green
    if ($OpenFile) {
        Start-Process notepad.exe $summaryFile
    }
}

function RollupSummary {
    param (
        [string]$Destination = "$env:USERPROFILE\Desktop\EcoDesk\Solution Tree\$global:Case"
    )

    try {
        Clear-Host
        Write-Host "Generating Rollup Summary..." -ForegroundColor Cyan
        
        if (-not (Test-Path $Destination)) {
            Write-Host "No solution tree summaries found at $Destination" -ForegroundColor Yellow
            Read-Host "Press Enter to return to menu"
            return
        }

        # Updated to match date-based filenames (e.g., 2025-02-28.txt or 2025-02-28_*.txt)
        $summaryFiles = Get-ChildItem -Path $Destination -File | Where-Object { $_.Name -match "^\d{4}-\d{2}-\d{2}.*\.txt$" } | Sort-Object Name
        if ($summaryFiles.Count -eq 0) {
            Write-Host "No summary files found in $Destination" -ForegroundColor Yellow
            Read-Host "Press Enter to return to menu"
            return
        }

        $rollupFile = Join-Path -Path $Destination -ChildPath "Rollup_Summary_$(Get-Date -Format 'yyyyMMddHHmmss').txt"
        $rollupContent = "Rollup Summary for Case $global:Case`n" +
                        "=================================`n" +
                        "Generated: $(Get-Date)`n" +
                        "Total Summaries: $($summaryFiles.Count)`n`n"

        foreach ($file in $summaryFiles) {
            $rollupContent += "Summary from $($file.Name):`n" +
                            "-----------------------------`n" +
                            (Get-Content -Path $file.FullName -Raw) +
                            "`n`n"
        }

        $rollupContent | Set-Content -Path $rollupFile
        Write-Host "Rollup summary saved to $rollupFile" -ForegroundColor Green
        Start-Process notepad.exe $rollupFile
    } catch {
        Write-Host "Error generating rollup summary: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}

function EscalateToTier2 {
    param (
        [string]$Destination = "$env:USERPROFILE\Desktop\EcoDesk\Tier2Esc\$global:Case"
    )

    try {
        Clear-Host
        Write-Host "Preparing Tier 2 Escalation Template..." -ForegroundColor Cyan

        # 1. ISSUE SUMMARY: Pick from the latest Summary file
        $summaryPath = "$env:USERPROFILE\Desktop\EcoDesk\Solution Tree\$global:Case"
        $latestSummary = Get-ChildItem -Path $summaryPath -File | 
                        Where-Object { $_.Name -match "^\d{4}-\d{2}-\d{2}.*\.txt$" } | 
                        Sort-Object Name -Descending | 
                        Select-Object -First 1
        $issueSummary = if ($latestSummary) {
            $content = Get-Content -Path $latestSummary.FullName -Raw
            if ($content -match "Issue: (.+?)(?=\r?\n\r?\n)") {
                $matches[1]
            } else {
                "Not found in summary file"
            }
        } else {
            "No summary file available"
        }

        # 2. EXPECTED OUTCOME
        $defaultOutcome = "Unable to identify root cause. Escalating for further investigation."
        Write-Host "Default Expected Outcome: $defaultOutcome" -ForegroundColor Yellow
        $modifyOutcome = Read-Host "Would you like to modify the expected outcome? (Y/N)"
        $expectedOutcome = if ($modifyOutcome.ToUpper() -eq "Y") {
            Read-Host "Enter modified expected outcome"
        } else {
            $defaultOutcome
        }

        # 3. CUSTOMERS TIME ZONE & AVAILABILITY: User input
        $timeZone = Read-Host "Enter customer's time zone (e.g., EST, PST, GMT)"
        $availability = Read-Host "Enter customer's availability (e.g., 9 AM - 5 PM)"

        # 4. TROUBLESHOOTING STEPS / LOG ANALYSIS
        $clientName = Read-Host "Enter client name"
        $troubleshootingSteps = if ($latestSummary) {
            $content = Get-Content -Path $latestSummary.FullName -Raw
            if ($content -match "Steps Taken:.*?-+(.+?)(?=\r?\n\r?\nNotes:)" -replace "`r`n", "`n") {
                $matches[1].Trim()
            } else {
                "No troubleshooting steps found in summary file"
            }
        } else {
            "No troubleshooting steps available"
        }

        # 5. ADDITIONAL INFORMATION
        # Staging Server: Auto-pick from Open-URLInChrome logic
        $stagingServer = "https://ce-staging.commvault.com/Reservation/Details/US1/$global:Case"
        try {
            $response = Invoke-WebRequest -Uri $stagingServer -UseBasicParsing -TimeoutSec 5 -ErrorAction SilentlyContinue
            if ($response.StatusCode -ne 200) {
                $stagingServer = "Not available (checked URL: $stagingServer)"
            }
        } catch {
            $stagingServer = "Not available (checked URL: $stagingServer)"
        }

        $jobID = Read-Host "Enter Job ID (if applicable)"
        $commserveDB = Read-Host "Enter CommServe Database details (if applicable)"
        $screenshotsLocation = Read-Host "Enter screenshots location (if applicable)"

        # Prepare the template
        $template = @"
Tier 2 Escalation Form
=====================
Case Number: $global:Case
CCID: $global:CCID
Date: $(Get-Date)

ISSUE SUMMARY:
--------------
$issueSummary

EXPECTED OUTCOME:
----------------
$expectedOutcome

CUSTOMERS TIME ZONE & AVAILABILITY:
----------------------------------
Time Zone: $timeZone
Availability: $availability

TROUBLESHOOTING STEPS / LOG ANALYSIS:
------------------------------------
Client Name: $clientName
Steps Taken:
$troubleshootingSteps

ADDITIONAL INFORMATION:
----------------------
Staging Server: $stagingServer
Job ID: $jobID
CommServe Database: $commserveDB
Screenshots Location: $screenshotsLocation
"@

        # Save the template with case number in filename
        $escalationPath = "$env:USERPROFILE\Desktop\EcoDesk\Tier2Esc\$global:Case"
        if (-not (Test-Path $escalationPath)) {
            New-Item -ItemType Directory -Path $escalationPath -Force | Out-Null
        }
        $escalationFile = Join-Path -Path $escalationPath -ChildPath "Escalation_$global:Case.txt"  # Updated filename
        $template | Set-Content -Path $escalationFile

        Write-Host "Escalation template saved to $escalationFile" -ForegroundColor Green
        Start-Process notepad.exe $escalationFile  # Open for copy/paste

        Read-Host "Press Enter to return to menu"
    } catch {
        Write-Host "Error preparing escalation template: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}

function EscalateToDevTeam {
    param (
        [string]$Destination = "$env:USERPROFILE\Desktop\EcoDesk\DevTeamEsc\$global:Case"
    )

    try {
        Clear-Host
        Write-Host "Preparing Development Team Escalation Template..." -ForegroundColor Cyan

        # 1. ISSUE SUMMARY: Pick from the latest Summary file
        $summaryPath = "$env:USERPROFILE\Desktop\EcoDesk\Solution Tree\$global:Case"
        $latestSummary = Get-ChildItem -Path $summaryPath -File | 
                        Where-Object { $_.Name -match "^\d{4}-\d{2}-\d{2}.*\.txt$" } | 
                        Sort-Object Name -Descending | 
                        Select-Object -First 1
        $issueSummary = if ($latestSummary) {
            $content = Get-Content -Path $latestSummary.FullName -Raw
            if ($content -match "Issue: (.+?)(?=\r?\n\r?\n)") {
                $matches[1]
            } else {
                "Not found in summary file"
            }
        } else {
            "No summary file available"
        }

        # 2. EXPECTED OUTCOME
        $defaultOutcome = "Unable to identify root cause. Escalating to Development Team for further investigation or resolution."
        Write-Host "Default Expected Outcome: $defaultOutcome" -ForegroundColor Yellow
        $modifyOutcome = Read-Host "Would you like to modify the expected outcome? (Y/N)"
        $expectedOutcome = if ($modifyOutcome.ToUpper() -eq "Y") {
            Read-Host "Enter modified expected outcome (1-2 lines)"
        } else {
            $defaultOutcome
        }

        # 3. CUSTOMERS TIME ZONE & AVAILABILITY: User input
        $timeZone = Read-Host "Enter customer's time zone (e.g., EST, PST, GMT)"
        $availability = Read-Host "Enter customer's availability (e.g., 9 AM - 5 PM)"

        # 4. TROUBLESHOOTING STEPS / LOG ANALYSIS
        $clientName = Read-Host "Enter client name"
        $troubleshootingSteps = if ($latestSummary) {
            $content = Get-Content -Path $latestSummary.FullName -Raw
            if ($content -match "Steps Taken:.*?-+(.+?)(?=\r?\n\r?\nNotes:)" -replace "`r`n", "`n") {
                $matches[1].Trim()
            } else {
                "No troubleshooting steps found in summary file"
            }
        } else {
            "No troubleshooting steps available"
        }

        # 5. ADDITIONAL INFORMATION
        # Staging Server: Auto-pick from Open-URLInChrome logic
        $stagingServer = "https://ce-staging.commvault.com/Reservation/Details/US1/$global:Case"
        try {
            $response = Invoke-WebRequest -Uri $stagingServer -UseBasicParsing -TimeoutSec 5 -ErrorAction SilentlyContinue
            if ($response.StatusCode -ne 200) {
                $stagingServer = "Not available (checked URL: $stagingServer)"
            }
        } catch {
            $stagingServer = "Not available (checked URL: $stagingServer)"
        }

        $jobID = Read-Host "Enter Job ID (if applicable)"
        $commserveDB = Read-Host "Enter CommServe Database details (if applicable)"
        $screenshotsLocation = Read-Host "Enter screenshots location (if applicable)"

        # 6. ENVIRONMENT DETAILS
        $commserver = Read-Host "Enter CommServer details (e.g., hostname)"
        $commserverVersion = if ($latestSummary) {
            $content = Get-Content -Path $latestSummary.FullName -Raw
            if ($content -match "Commvault Version: (.+?)(?=\r?\n)") {
                $matches[1]
            } else {
                "Not found in summary file"
            }
        } else {
            "Not available"
        }
        $accessNode = Read-Host "Enter Access Node details (if applicable)"
        $indexServer = Read-Host "Enter Index Server details (if applicable)"
        $webServer = Read-Host "Enter Web Server details (if applicable)"

        # Prepare the template
        $template = @"
ISSUE SUMMARY:
===============
$issueSummary

EXPECTED OUTCOME:
==================
$expectedOutcome

CUSTOMERS TIME ZONE & AVAILABILITY:
==================================
Time Zone: $timeZone
Availability: $availability

TROUBLESHOOTING STEPS / LOG ANALYSIS:
=====================================
Client Name: $clientName
Steps Taken:
$troubleshootingSteps

ADDITIONAL INFORMATION:
========================
Staging Server: $stagingServer
Job ID: $jobID
CommServe Database: $commserveDB
Screenshots Location: $screenshotsLocation

ENVIRONMENT DETAILS:
========================
CommServer: $commserver
CommServer Version: $commserverVersion
Access Node: $accessNode
Index Server: $indexServer
Web Server: $webServer
"@

        # Save the template with case number in filename
        $escalationPath = "$env:USERPROFILE\Desktop\EcoDesk\DevTeamEsc\$global:Case"
        if (-not (Test-Path $escalationPath)) {
            New-Item -ItemType Directory -Path $escalationPath -Force | Out-Null
        }
        $escalationFile = Join-Path -Path $escalationPath -ChildPath "Escalation_$global:Case.txt"  # Updated filename
        $template | Set-Content -Path $escalationFile

        Write-Host "Escalation template saved to $escalationFile" -ForegroundColor Green
        Start-Process notepad.exe $escalationFile  # Open for copy/paste

        Read-Host "Press Enter to return to menu"
    } catch {
        Write-Host "Error preparing escalation template: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}

function AnalyzeLogs {
    param (
        [string]$Destination = "\\englog\escalationlogs\$global:CCID\$global:Case",
        [int]$ContextLines = 2,
        [switch]$HtmlOutput = $false
    )

    try {
        Clear-Host
        Write-Host "Enhanced Log Analysis Tool" -ForegroundColor Cyan
        Write-Host "=========================" -ForegroundColor Cyan
        
        # Check if default destination exists, prompt for manual input if not
        if (-not (Test-Path $Destination)) {
            Write-Host "Default log location $Destination not found." -ForegroundColor Yellow
            $manualDestination = Read-Host "Please enter the manual log location (e.g., path to log files)"
            if (-not [string]::IsNullOrWhiteSpace($manualDestination)) {
                $Destination = $manualDestination
                if (-not (Test-Path $Destination)) {
                    Write-Host "Manual log location $Destination not found. Please verify the path." -ForegroundColor Red
                    Read-Host "Press Enter to return to menu"
                    return
                }
                Write-Host "Using manual log location: $Destination" -ForegroundColor Green
            } else {
                Write-Host "No location entered. Exiting." -ForegroundColor Red
                Read-Host "Press Enter to return to menu"
                return
            }
        }

        Write-Host "Checking for log files in $Destination..." -ForegroundColor Yellow
        $logFiles = Get-ChildItem -Path $Destination -Recurse -File -Filter "*.log" -ErrorAction Stop
        $txtFiles = Get-ChildItem -Path $Destination -Recurse -File -Filter "*.txt" -ErrorAction Stop
        $allLogFiles = @($logFiles | Where-Object { $_.PSIsContainer -eq $false -and $_.LastWriteTime }) + @($txtFiles | Where-Object { $_.PSIsContainer -eq $false -and $_.LastWriteTime })  # Explicitly filter for files with LastWriteTime

        if ($allLogFiles.Count -eq 0) {
            Write-Host "No log files found in $Destination" -ForegroundColor Yellow
            Read-Host "Press Enter to return to menu"
            return
        }
        Write-Host "Found $($allLogFiles.Count) log files to analyze" -ForegroundColor Green

        # Debug: Verify file types and LastWriteTime, check for reference folder files
        Write-Host "Debug: Verifying file objects in `$allLogFiles..." -ForegroundColor Yellow
        foreach ($file in $allLogFiles) {
            $hasLastWriteTime = if ($file.LastWriteTime) { 'Yes' } else { 'No' }
            Write-Host "Debug: File - Name: $($file.Name), IsContainer: $($file.PSIsContainer), HasLastWriteTime: $hasLastWriteTime, FullPath: $($file.FullName)" -ForegroundColor Green
            if ($file.FullName -like "*references*") {
                Write-Host "Debug: Found potential reference file in `$allLogFiles: $($file.FullName)" -ForegroundColor Red
            }
        }

        Write-Host "Filtering Options:" -ForegroundColor Cyan
        $jobID = Read-Host "Enter JobID to filter logs (or press Enter for all)"
        $dateFilter = Read-Host "Enter date range (e.g., '2025-02-01 to 2025-02-28' or press Enter for all)"
        $excludePattern = Read-Host "Enter pattern to exclude (e.g., 'deprecated', or press Enter for none)"

        $logsToAnalyze = $allLogFiles
        if ($jobID) { 
            Write-Host "Filtering logs for JobID '$jobID'..." -ForegroundColor Yellow
            $logsToAnalyze = $logsToAnalyze | Where-Object { 
                if ($_.PSIsContainer -eq $false -and $_.LastWriteTime) {
                    Write-Host "Debug: Checking file $($_.FullName) for JobID '$jobID'" -ForegroundColor Yellow
                    $_.Name -like "*$jobID*" 
                } else { $false }
            }
        }

        # Handle date range filtering with robust parsing and error handling
        if ($dateFilter -and $dateFilter.Trim()) {
            Write-Host "Filtering logs for date range '$dateFilter'..." -ForegroundColor Yellow
            $dateRange = $dateFilter.Trim() -split " to "
            if ($dateRange.Count -eq 2) {
                try {
                    $startDate = [datetime]::Parse($dateRange[0].Trim())
                    $endDate = [datetime]::Parse($dateRange[1].Trim())
                    if ($startDate -gt $endDate) {
                        Write-Host "Error: Start date ($startDate) must be before end date ($endDate)." -ForegroundColor Red
                        Read-Host "Press Enter to return to menu"
                        return
                    }
                    Write-Host "Debug: Start Date = $startDate, End Date = $endDate" -ForegroundColor Yellow
                    $logsToAnalyze = $logsToAnalyze | Where-Object { 
                        if ($_.PSIsContainer -eq $false -and $_.LastWriteTime) {
                            Write-Host "Debug: Checking file $($_.FullName) with LastWriteTime $($_.LastWriteTime)" -ForegroundColor Yellow
                            $_.LastWriteTime -ge $startDate -and $_.LastWriteTime -le $endDate
                        } else {
                            Write-Host "Debug: Skipping file $($_.FullName) - No LastWriteTime or is a directory" -ForegroundColor Yellow
                            $false  # Skip files without LastWriteTime or directories
                        }
                    }
                } catch {
                    Write-Host "Error parsing date range '$dateFilter': $_" -ForegroundColor Red
                    Write-Host "Please enter dates in format 'YYYY-MM-DD to YYYY-MM-DD'." -ForegroundColor Yellow
                    Read-Host "Press Enter to return to menu"
                    return
                }
            } else {
                Write-Host "Invalid date range format. Use 'YYYY-MM-DD to YYYY-MM-DD'." -ForegroundColor Red
                Read-Host "Press Enter to return to menu"
                return
            }
        } else {
            Write-Host "No date range specified, skipping date filtering." -ForegroundColor Yellow
        }

        if ($excludePattern) {
            Write-Host "Excluding logs with pattern '$excludePattern'..." -ForegroundColor Yellow
            $logsToAnalyze = $logsToAnalyze | Where-Object { 
                if ($_.PSIsContainer -eq $false) {
                    Write-Host "Debug: Checking file $($_.FullName) for exclude pattern '$excludePattern'" -ForegroundColor Yellow
                    $_.FullName -notmatch $excludePattern 
                } else { $false }
            }
        }

        if ($logsToAnalyze.Count -eq 0) {
            Write-Host "No logs match the specified filters." -ForegroundColor Yellow
            Read-Host "Press Enter to return to menu"
            return
        }
        Write-Host "Analyzing $($logsToAnalyze.Count) filtered log files..." -ForegroundColor Green

        # Debug: Verify final $logsToAnalyze
        Write-Host "Debug: Final `$logsToAnalyze count = $($logsToAnalyze.Count)" -ForegroundColor Yellow
        $logsToAnalyze | Select-Object -First 5 | Format-List Name, PSIsContainer, LastWriteTime | Out-Host

        # Use global keywords with fallbacks
        $errorPatterns = if ($global:ErrorKeywords) { $global:ErrorKeywords } else { @("error", "fail", "exception", "critical", "fatal") }
        $warningPatterns = if ($global:WarningKeywords) { $global:WarningKeywords } else { @("warning", "caution") }

        $analysisResults = @{}
        $errorCount = 0
        $warningCount = 0
        $criticalCount = 0
        $fileCount = 0

        foreach ($log in $logsToAnalyze) {
            $fileCount++
            Write-Host "Analyzing log file $fileCount of $($logsToAnalyze.Count): $($log.FullName)" -ForegroundColor Green
            $lines = [System.IO.File]::ReadLines($log.FullName)
            $lineNumber = 0
            $results = @()
            foreach ($line in $lines) {
                if ($excludePattern -and $line -match $excludePattern) { $lineNumber++; continue }
                $errorMatch = $errorPatterns | Where-Object { $line -match $_ }
                $warningMatch = $warningPatterns | Where-Object { $line -match $_ }
                if ($errorMatch -or $warningMatch) {
                    $start = [Math]::Max(0, $lineNumber - $ContextLines)
                    $end = [Math]::Min($lines.Count - 1, $lineNumber + $ContextLines)
                    $context = ($start..$end) | ForEach-Object { "$($_ + 1): $($lines[$_])" } | Out-String
                    
                    if ($errorMatch) {
                        $severity = if ($line -match "critical|fatal") { "Critical" } else { "Error" }
                        $results += [PSCustomObject]@{ Type = $severity; Line = $context }
                        if ($severity -eq "Critical") { $criticalCount++ } else { $errorCount++ }
                    } elseif ($warningMatch) {
                        $results += [PSCustomObject]@{ Type = "Warning"; Line = $context }
                        $warningCount++
                    }
                }
                $lineNumber++
            }
            if ($results.Count -gt 0) { $analysisResults[$log.FullName] = $results }
        }

        Clear-Host
        Write-Host "Log Analysis Summary" -ForegroundColor Cyan
        Write-Host "====================" -ForegroundColor Cyan
        Write-Host "Total Log Files Analyzed: $($logsToAnalyze.Count)"
        Write-Host "Total Critical Errors Found: $criticalCount" -ForegroundColor Magenta
        Write-Host "Total Errors Found: $errorCount" -ForegroundColor Red
        Write-Host "Total Warnings Found: $warningCount" -ForegroundColor Yellow
        Write-Host ""
        
        if ($analysisResults.Count -gt 0) {
            Write-Host "Detailed Findings:" -ForegroundColor Green
            foreach ($log in $analysisResults.Keys) {
                Write-Host "`nLog File: $log" -ForegroundColor White
                foreach ($result in $analysisResults[$log]) {
                    switch ($result.Type) {
                        "Critical" { Write-Host "$($result.Type):" -ForegroundColor Magenta; Write-Host "  $($result.Line)" }
                        "Error" { Write-Host "$($result.Type):" -ForegroundColor Red; Write-Host "  $($result.Line)" }
                        "Warning" { Write-Host "$($result.Type):" -ForegroundColor Yellow; Write-Host "  $($result.Line)" }
                    }
                }
            }
            $outputFile = Join-Path -Path $Destination -ChildPath "LogAnalysis_$(Get-Date -Format 'yyyyMMddHHmmss')"
            $summaryContent = "Log Analysis Summary`n===================`nDate: $(Get-Date)`nTotal Log Files Analyzed: $($logsToAnalyze.Count)`nTotal Critical Errors Found: $criticalCount`nTotal Errors Found: $errorCount`nTotal Warnings Found: $warningCount`n`nDetailed Findings:`n"
            foreach ($log in $analysisResults.Keys) {
                $summaryContent += "`nLog File: $log`n"
                foreach ($result in $analysisResults[$log]) {
                    $summaryContent += "$($result.Type):`n$($result.Line)`n"
                }
            }
            if ($HtmlOutput) {
                $htmlContent = @"
<html>
<head>
    <title>Log Analysis - Case $global:Case</title>
    <style>
        body { font-family: Arial, sans-serif; }
        h2 { color: #0066cc; }
        .critical { color: #ff00ff; font-weight: bold; }
        .error { color: #ff0000; }
        .warning { color: #ffff00; background-color: #333; }
        pre { white-space: pre-wrap; }
    </style>
</head>
<body>
    <h2>Log Analysis Summary - Case $global:Case</h2>
    <p>Date: $(Get-Date)</p>
    <p>Total Log Files Analyzed: $($logsToAnalyze.Count)</p>
    <p><span class='critical'>Total Critical Errors Found: $criticalCount</span></p>
    <p><span class='error'>Total Errors Found: $errorCount</span></p>
    <p><span class='warning'>Total Warnings Found: $warningCount</span></p>
    <h3>Detailed Findings:</h3>
"@
                foreach ($log in $analysisResults.Keys) {
                    $htmlContent += "<h4>Log File: $log</h4>"
                    foreach ($result in $analysisResults[$log]) {
                        $class = switch ($result.Type) { "Critical" { "critical" } "Error" { "error" } "Warning" { "warning" } }
                        $htmlContent += "<p><span class='$class'>$($result.Type):</span><br><pre>$($result.Line)</pre></p>"
                    }
                }
                $htmlContent += "</body></html>"
                $outputFile += ".html"
                $htmlContent | Set-Content -Path $outputFile
                Write-Host "`nAnalysis results saved as HTML to $outputFile" -ForegroundColor Green
                Start-Process $outputFile
            } else {
                $outputFile += ".txt"
                $summaryContent | Set-Content -Path $outputFile
                Write-Host "`nAnalysis results saved as text to $outputFile" -ForegroundColor Green
                Start-Process notepad.exe $outputFile
            }
            $global:LastLogAnalysis = @{
                Findings = $summaryContent
                LogCuts = ($analysisResults.Values | ForEach-Object { $_.Line } | Out-String)
            }
            Write-Host "Opening destination folder $Destination..." -ForegroundColor Yellow
            ii $Destination
        } else {
            Write-Host "No issues found in the analyzed logs" -ForegroundColor Green
        }
        
        Write-Host "Analysis completed." -ForegroundColor Green
        Read-Host "Press Enter to return to menu"
    } catch {
        Write-Host "Error in AnalyzeLogs: $_" -ForegroundColor Red
        if ($_.Exception -is [System.IO.IOException] -or $_.Exception -is [System.UnauthorizedAccessException]) {
            Write-Host "Possible causes: network issue, permission denied, or file/path inaccessible." -ForegroundColor Yellow
        } elseif ($_.Exception -is [System.Management.Automation.MethodInvocationException]) {
            Write-Host "Possible causes: invalid date range format, missing file properties, or type mismatch in filtering." -ForegroundColor Yellow
            # Debug: Output the state of $logsToAnalyze and current file
            Write-Host "Debug: `$logsToAnalyze count = $($logsToAnalyze.Count)" -ForegroundColor Yellow
            Write-Host "Debug: Current file being processed = $($log.FullName)" -ForegroundColor Yellow
            if ($logsToAnalyze) {
                Write-Host "Debug: First few objects in `$logsToAnalyze:" -ForegroundColor Yellow
                $logsToAnalyze | Select-Object -First 5 | Format-List Name, PSIsContainer, LastWriteTime | Out-Host
            }
        }
        Read-Host "Press Enter to return to menu"
    }
}

function CheckKBArticle {
    param (
        [string]$Agent,
        [string]$Issue
    )

    $kbPath = "$env:USERPROFILE\Desktop\EcoDesk\KB\KBArticles.txt"
    if (-not (Test-Path $kbPath)) { return "Not Found" }

    $kbContent = Get-Content -Path $kbPath -Raw
    if ($kbContent -match [regex]::Escape("$Agent|$Issue")) {
        $match = ($kbContent -split "`n" | Where-Object { $_ -match [regex]::Escape("$Agent|$Issue") })[0]
        return ($match -split "\|")[3]
    }
    return "Not Found"
}

function SaveKBArticle {
    param (
        [string]$Agent,
        [string]$Issue,
        [string]$Content
    )

    $kbPath = "$env:USERPROFILE\Desktop\EcoDesk\KB\KBArticles.txt"
    $kbDir = Split-Path $kbPath -Parent
    if (-not (Test-Path $kbDir)) {
        New-Item -ItemType Directory -Path $kbDir -Force | Out-Null
    }

    "$Agent|$Issue|$Content" | Add-Content -Path $kbPath
}

function ManageTroubleshootingSteps {
    param (
        [string]$Destination ="$env:USERPROFILE\Desktop\EcoDesk\Config"
    )

    # Load existing custom steps with error handling
    $stepsFile = Join-Path -Path $Destination -ChildPath "CustomSteps.json"
    $customSteps = try {
        if (Test-Path $stepsFile) {
            Get-Content -Path $stepsFile -Raw | ConvertFrom-Json -ErrorAction Stop
        } else {
            @{}
        }
    } catch {
        Write-Host "Error loading custom steps from $stepsFile $_" -ForegroundColor Red
        Write-Host "Starting with an empty steps collection." -ForegroundColor Yellow
        @{}
    }

    # Define Focus Groups (same as SolutionTree)
    $focusGroups = @{
        "Messaging" = @{
            "SQL Server" = @("SQL Server", "Azure SQL", "AWS SQL", "Availability Group", "SQL Cluster")
            "Exchange" = @("Mailbox On-Prem", "Mailbox Online", "Exchange Database", "Exchange DAG", "Journaling", "Content Store SMTP")
            "Sharepoint" = @("Sharepoint On-Prem", "Sharepoint Online")
            "All Other" = @("Cosmos DB", "Onedrive", "Teams", "GitHub", "GitLab", "Lotus Notes", "Dynamo DB", "Cockroach DB", "Splunk", "Data Insight")
        }
        "Media Management" = @{
            "General" = @("Worm Copy", "Auxiliary Copy", "Content Indexing and Search", "DASH full (Synthetic Full)", "Data Verification", "Deduplication", "Gridstor", "Index Cache", "Library and Drive Configuration", "Media Agent Server", "Media Refresh", "Media Explorer", "Multiplexing", "NAS NDMP Agents", "NFS ObjectStore / Hybrid File Store", "Vault Tracker", "Edge", "Data Analytics", "Catalog Media")
            "Hyperscale" = @("Hyperscale 1.5 (Gluster)", "Hyperscale X (CDS)", "Remote Office Appliance")
        }
        "Server" = @("MongoDB Security Issue", "Worm Copy", "cyberark credential access integration", "Alerts", "Cloud Services Website", "CommCell Console", "Command Line Interface/QCommand (CLI)", "CommCell Merge", "CommServe Database", "CommServe DR Process", "CommServe Install/Upgrade", "CommServe Services and Performance", "Configuration", "Data Aging", "Downloads and Updates", "Encryption", "End of Life", "Job Controller Management and Job History", "Licensing", "Log File Upload", "Policies and Groups", "Reports", "User Security (CommVault level)", "Schedules", "Security Concerns", "T2 Training", "Workflows (graphical method to chain multiple commands in sequence)", "Error Deploying V11 Service Pack 11", "I need to access my CommServe Database")
        "Client" = @("Windows App Aware snap mounts take a long time or seem to be hung", "Application Aware", "Error Code [72:106] Caught Access Violation Exception during LVM processing", "Microsoft Azure Cosmos DB", "Kubernetes", "Metadata Collection Fails for Virtual Machine Backups", "CVVD.SYS driver may cause reboot of machine due to issue when using Live Browse or Block Level restore", "CVE-2021-4034 PKEXEC exploit - Actions to take", "Missing DR VMs within the Cloud after Upgrade to CPR2023E", "Amazon Client (Lifecycle and EC2 Instance Backup)", "Hyper-V", "Microsoft Azure", "Google Cloud Platform", "Microsoft Azure Stack", "Docker", "Huawei FusionCompute", "Nutanix Acropolis", "Oracle Cloud Infrastructure", "Oracle VM", "Red Hat Enterprise Virtualization", "OpenStack", "XenServer", "VMware", "v2 Indexing", "How do I configure Metadata Collection/Granular Recovery for VSA within SP11 and on?", "Failed to create the backup statistics information object from a database record")
        "Unix" = @("Amazon Redshift", "DynamoDB", "Cassandra", "DB2", "DB2 MultiNode", "Greenplum", "HANA", "HBase", "Hive", "Informix", "MaxDB", "Microsoft SQL Server", "MongoDB", "MySQL/MariaDB", "Oracle", "Oracle RAC", "PostgreSQL", "SAP Archive Link", "SAP Oracle", "SAP Landscape Management Integration", "Sybase")
        "Metallic" = @("All subscriptions", "Microsoft 365", "Risk Analysis for Microsoft 365", "Integration with Microsoft 365 Backup Storage", "Endpoints", "Files & Objects", "Archive for Files & Objects", "Threat Scan for File & Object", "VMs", "Auto Recovery for VM", "Kubernetes", "Databases", "Google Workspace", "Salesforce", "Microsoft Dynamics 365", "Active Directory", "Migration as a Service", "Threatwise", "Security IQ", "Cleanroom", "Cloud Rewind", "Clumio", "CV Cloud SaaS & Hyperscale X", "Air Gap Protect", "Lighthouse", "Onboarding Tracking (SoftwareONE ONLY)")
    }
    $commvaultVersions = @("FR40", "FR36", "FR32", "FR28", "FR24")

    try {
        Clear-Host
        Write-Host "Manage Troubleshooting Steps" -ForegroundColor Cyan
        Write-Host "===========================" -ForegroundColor Cyan

        # Select Focus Group
        Write-Host "Select Focus Group:" -ForegroundColor Cyan
        $focusGroupKeys = $focusGroups.Keys | Sort-Object
        for ($i = 0; $i -lt $focusGroupKeys.Count; $i++) {
            Write-Host "$($i+1). $($focusGroupKeys[$i])"
        }
        $focusChoice = [int](Read-Host "Enter number (1-$($focusGroupKeys.Count))") - 1
        if ($focusChoice -lt 0 -or $focusChoice -ge $focusGroupKeys.Count) { throw "Invalid focus group selection" }
        $selectedFocus = $focusGroupKeys[$focusChoice]

        # Select Agent/Subcategory
        Clear-Host
        Write-Host "Selected Focus Group: $selectedFocus" -ForegroundColor Cyan
        $agents = $focusGroups[$selectedFocus]
        if ($agents -is [hashtable]) {
            Write-Host "Select Category:" -ForegroundColor Cyan
            $categoryKeys = $agents.Keys | Sort-Object
            for ($i = 0; $i -lt $categoryKeys.Count; $i++) {
                Write-Host "$($i+1). $($categoryKeys[$i])"
            }
            $categoryChoice = [int](Read-Host "Enter number (1-$($categoryKeys.Count))") - 1
            if ($categoryChoice -lt 0 -or $categoryChoice -ge $categoryKeys.Count) { throw "Invalid category selection" }
            $selectedCategory = $categoryKeys[$categoryChoice]
            $agentList = $agents[$selectedCategory]
        } else {
            $agentList = $agents
        }

        Write-Host "`nSelect Agent:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $agentList.Count; $i++) {
            Write-Host "$($i+1). $($agentList[$i])"
        }
        $agentChoice = [int](Read-Host "Enter number (1-$($agentList.Count))") - 1
        if ($agentChoice -lt 0 -or $agentChoice -ge $agentList.Count) { throw "Invalid agent selection" }
        $selectedAgent = $agentList[$agentChoice]

        # Select Commvault Version
        Clear-Host
        Write-Host "Selected Agent: $selectedAgent" -ForegroundColor Cyan
        Write-Host "Select Commvault Version:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $commvaultVersions.Count; $i++) {
            Write-Host "$($i+1). $($commvaultVersions[$i])"
        }
        $versionChoice = [int](Read-Host "Enter number (1-$($commvaultVersions.Count))") - 1
        if ($versionChoice -lt 0 -or $versionChoice -ge $commvaultVersions.Count) { throw "Invalid version selection" }
        $selectedVersion = $commvaultVersions[$versionChoice]

        # Key for storing steps
        $stepKey = "$selectedFocus|$selectedAgent|$selectedVersion"
        $currentSteps = if ($customSteps.PSObject.Properties.Name -contains $stepKey) {
            $customSteps.$stepKey
        } else {
            # Default steps
            @(
                "Verify service status and restart if necessary",
                "Check log files for specific error messages",
                "Validate network connectivity to affected component",
                "Confirm user permissions and access rights",
                "Review configuration settings",
                "Check disk space and resource availability",
                "Update to latest patches if applicable",
                "Run diagnostic tools specific to the agent",
                "Verify backup/restore job logs",
                "Consult Commvault documentation for known issues",
                "Escalate to Tier 2 support"
            )
        }

        # Menu for managing steps
        do {
            Clear-Host
            Write-Host "Managing Steps for: $selectedFocus > $selectedAgent ($selectedVersion)" -ForegroundColor Cyan
            Write-Host "Current Steps:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $currentSteps.Count; $i++) {
                Write-Host "$($i+1). $($currentSteps[$i])"
            }
            Write-Host "`nOptions:" -ForegroundColor Cyan
            Write-Host "1. Add New Step"
            Write-Host "2. Modify Existing Step"
            Write-Host "3. Delete Step"
            Write-Host "4. Reset to Default Steps"
            Write-Host "5. Save and Exit"
            Write-Host "6. Exit without Saving"

            $choice = Read-Host "Enter your choice (1-6)"
            switch ($choice) {
                "1" {
                    $newStep = Read-Host "Enter new troubleshooting step"
                    $currentSteps += $newStep
                    Write-Host "Step added: $newStep" -ForegroundColor Green
                    Start-Sleep -Seconds 1
                }
                "2" {
                    $stepNum = [int](Read-Host "Enter step number to modify (1-$($currentSteps.Count))") - 1
                    if ($stepNum -ge 0 -and $stepNum -lt $currentSteps.Count) {
                        $currentSteps[$stepNum] = Read-Host "Enter new text for step $($stepNum + 1)"
                        Write-Host "Step $($stepNum + 1) modified" -ForegroundColor Green
                        Start-Sleep -Seconds 1
                    } else {
                        Write-Host "Invalid step number" -ForegroundColor Red
                        Start-Sleep -Seconds 1
                    }
                }
                "3" {
                    $stepNum = [int](Read-Host "Enter step number to delete (1-$($currentSteps.Count))") - 1
                    if ($stepNum -ge 0 -and $stepNum -lt $currentSteps.Count) {
                        $currentSteps = $currentSteps | Where-Object { $_ -ne $currentSteps[$stepNum] }
                        Write-Host "Step $($stepNum + 1) deleted" -ForegroundColor Green
                        Start-Sleep -Seconds 1
                    } else {
                        Write-Host "Invalid step number" -ForegroundColor Red
                        Start-Sleep -Seconds 1
                    }
                }
                "4" {
                    $currentSteps = @(
                        "Verify service status and restart if necessary",
                        "Check log files for specific error messages",
                        "Validate network connectivity to affected component",
                        "Confirm user permissions and access rights",
                        "Review configuration settings",
                        "Check disk space and resource availability",
                        "Update to latest patches if applicable",
                        "Run diagnostic tools specific to the agent",
                        "Verify backup/restore job logs",
                        "Consult Commvault documentation for known issues",
                        "Escalate to Tier 2 support"
                    )
                    Write-Host "Steps reset to default" -ForegroundColor Green
                    Start-Sleep -Seconds 1
                }
                "5" {
                    if (-not (Test-Path $Destination)) {
                        New-Item -ItemType Directory -Path $Destination -Force | Out-Null
                    }
                    $customSteps.$stepKey = $currentSteps
                    $customSteps | ConvertTo-Json | Set-Content -Path $stepsFile
                    Write-Host "Custom steps saved to $stepsFile" -ForegroundColor Green
                    return
                }
                "6" { return }
                default { Write-Host "Invalid choice" -ForegroundColor Red; Start-Sleep -Seconds 1 }
            }
        } while ($true)
    } catch {
        Write-Host "Error in ManageTroubleshootingSteps: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}


function UpdateErrorKeywords {
    param (
        [string]$ConfigFile = "$env:USERPROFILE\Desktop\EcoDesk\Config\ErrorKeywords.json"
    )

    Clear-Host
    Write-Host "Update Error and Warning Keywords" -ForegroundColor Cyan
    Write-Host "=================================" -ForegroundColor Cyan

    # Ensure EcoDesk\Config directory exists
    $ecoDeskDir = Split-Path $ConfigFile -Parent
    if (-not (Test-Path $ecoDeskDir)) {
        New-Item -ItemType Directory -Path $ecoDeskDir -Force | Out-Null
    }

    # Load existing keywords from file if not already loaded
    if (-not $global:ErrorKeywords -or -not $global:WarningKeywords) {
        if (Test-Path $ConfigFile) {
            try {
                $keywords = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json -ErrorAction Stop
                $global:ErrorKeywords = $keywords.ErrorKeywords
                $global:WarningKeywords = $keywords.WarningKeywords
            } catch {
                Write-Host "Error loading keywords from $ConfigFile $_" -ForegroundColor Red
                $global:ErrorKeywords = @("error", "fail", "exception", "critical", "fatal")
                $global:WarningKeywords = @("warning", "caution")
            }
        } else {
            $global:ErrorKeywords = @("error", "fail", "exception", "critical", "fatal")
            $global:WarningKeywords = @("warning", "caution")
        }
    }

    do {
        Write-Host "`nCurrent Error Keywords: " -ForegroundColor Yellow -NoNewline
        Write-Host ($global:ErrorKeywords -join ", ")
        Write-Host "Current Warning Keywords: " -ForegroundColor Yellow -NoNewline
        Write-Host ($global:WarningKeywords -join ", ")
        Write-Host "`nOptions:" -ForegroundColor Cyan
        Write-Host "1. Add Error Keyword"
        Write-Host "2. Add Warning Keyword"
        Write-Host "3. Remove Error Keyword"
        Write-Host "4. Remove Warning Keyword"
        Write-Host "5. Reset to Defaults"
        Write-Host "6. Save and Exit"

        $choice = Read-Host "Enter your choice (1-6)"
        switch ($choice) {
            "1" {
                $newError = Read-Host "Enter new error keyword"
                if ($newError -and $global:ErrorKeywords -notcontains $newError) {
                    $global:ErrorKeywords += $newError
                    Write-Host "Added '$newError' to error keywords." -ForegroundColor Green
                } else {
                    Write-Host "Keyword already exists or invalid." -ForegroundColor Red
                }
                Start-Sleep -Seconds 1
            }
            "2" {
                $newWarning = Read-Host "Enter new warning keyword"
                if ($newWarning -and $global:WarningKeywords -notcontains $newWarning) {
                    $global:WarningKeywords += $newWarning
                    Write-Host "Added '$newWarning' to warning keywords." -ForegroundColor Green
                } else {
                    Write-Host "Keyword already exists or invalid." -ForegroundColor Red
                }
                Start-Sleep -Seconds 1
            }
            "3" {
                $removeError = Read-Host "Enter error keyword to remove"
                if ($global:ErrorKeywords -contains $removeError) {
                    $global:ErrorKeywords = $global:ErrorKeywords | Where-Object { $_ -ne $removeError }
                    Write-Host "Removed '$removeError' from error keywords." -ForegroundColor Green
                } else {
                    Write-Host "Keyword not found." -ForegroundColor Red
                }
                Start-Sleep -Seconds 1
            }
            "4" {
                $removeWarning = Read-Host "Enter warning keyword to remove"
                if ($global:WarningKeywords -contains $removeWarning) {
                    $global:WarningKeywords = $global:WarningKeywords | Where-Object { $_ -ne $removeWarning }
                    Write-Host "Removed '$removeWarning' from warning keywords." -ForegroundColor Green
                } else {
                    Write-Host "Keyword not found." -ForegroundColor Red
                }
                Start-Sleep -Seconds 1
            }
            "5" {
                $global:ErrorKeywords = @("error", "fail", "exception", "critical", "fatal")
                $global:WarningKeywords = @("warning", "caution")
                Write-Host "Reset to default keywords." -ForegroundColor Green
                Start-Sleep -Seconds 1
            }
            "6" {
                # Save to config file
                $keywordConfig = @{
                    ErrorKeywords = $global:ErrorKeywords
                    WarningKeywords = $global:WarningKeywords
                }
                $keywordConfig | ConvertTo-Json | Set-Content -Path $ConfigFile
                Write-Host "Keywords saved to $ConfigFile and updated successfully." -ForegroundColor Green
                return
            }
            default { Write-Host "Invalid choice. Try again." -ForegroundColor Red; Start-Sleep -Seconds 1 }
        }
    } while ($true)
}

function ManageWebConfig {
    param (
        [string]$MainConfigPath = "$env:USERPROFILE\Desktop\EcoDesk\Config",
        [string]$AllConfigPath = "$env:USERPROFILE\Desktop\EcoDesk\Config\All Config"
    )

    try {
        Clear-Host
        Write-Host "Manage WebConfig Files" -ForegroundColor Cyan
        Write-Host "======================" -ForegroundColor Cyan
        
        # Ensure directories exist
        if (-not (Test-Path $MainConfigPath)) {
            New-Item -ItemType Directory -Path $MainConfigPath -Force | Out-Null
        }
        if (-not (Test-Path $AllConfigPath)) {
            New-Item -ItemType Directory -Path $AllConfigPath -Force | Out-Null
        }

        # Define main config files
        $mainStepsFile = Join-Path -Path $MainConfigPath -ChildPath "CustomSteps.json"
        $mainKeywordsFile = Join-Path -Path $MainConfigPath -ChildPath "ErrorKeywords.json"

        # Initialize main files if they don’t exist
        if (-not (Test-Path $mainStepsFile)) {
            @{} | ConvertTo-Json | Set-Content -Path $mainStepsFile
        }
        if (-not (Test-Path $mainKeywordsFile)) {
            @{
                ErrorKeywords = @("error", "fail", "exception", "critical", "fatal")
                WarningKeywords = @("warning", "caution")
            } | ConvertTo-Json | Set-Content -Path $mainKeywordsFile
        }

        # Load main configs
        $mainSteps = Get-Content -Path $mainStepsFile -Raw | ConvertFrom-Json
        $mainKeywords = Get-Content -Path $mainKeywordsFile -Raw | ConvertFrom-Json

        do {
            Write-Host "`nMain CustomSteps.json Content:" -ForegroundColor Yellow
            if ($mainSteps.PSObject.Properties.Count -eq 0) {
                Write-Host "No custom steps defined."
            } else {
                $mainSteps.PSObject.Properties | ForEach-Object { Write-Host "$($_.Name): $($_.Value -join ', ')" }
            }
            Write-Host "`nMain ErrorKeywords.json Content:" -ForegroundColor Yellow
            Write-Host "Error Keywords: " ($mainKeywords.ErrorKeywords -join ", ")
            Write-Host "Warning Keywords: " ($mainKeywords.WarningKeywords -join ", ")

            # Get all shared config files in All Config folder
            $sharedStepsFiles = Get-ChildItem -Path $AllConfigPath -Filter "CustomSteps*.json"
            $sharedKeywordsFiles = Get-ChildItem -Path $AllConfigPath -Filter "ErrorKeywords*.json"

            Write-Host "`nOptions:" -ForegroundColor Cyan
            Write-Host "1. Add New Shared Config Files (CustomSteps.json/ErrorKeywords.json)"
            Write-Host "2. Import Steps from Shared CustomSteps Files"
            Write-Host "3. Import Keywords from Shared ErrorKeywords Files"
            Write-Host "4. Import All Steps and Keywords at Once"
            Write-Host "5. Save and Exit"
            Write-Host "6. Exit without Saving"

            $choice = Read-Host "Enter your choice (1-6)"
            switch ($choice) {
                "1" {
                    # Prompt user to specify paths to shared files
                    $stepsPath = Read-Host "Enter path to shared CustomSteps.json file (or press Enter to skip)"
                    $keywordsPath = Read-Host "Enter path to shared ErrorKeywords.json file (or press Enter to skip)"
                    
                    # Handle CustomSteps.json
                    if ($stepsPath -and (Test-Path $stepsPath)) {
                        $existingStepsNumbers = $sharedStepsFiles | ForEach-Object { 
                            if ($_.BaseName -match "CustomSteps(\d+)") { [int]$matches[1] } 
                        }
                        $nextStepsNumber = if ($existingStepsNumbers) { ($existingStepsNumbers | Measure-Object -Maximum).Maximum + 1 } else { 1 }
                        $newStepsFile = Join-Path -Path $AllConfigPath -ChildPath "CustomSteps$nextStepsNumber.json"
                        Copy-Item -Path $stepsPath -Destination $newStepsFile -Force
                        Write-Host "Copied to $newStepsFile" -ForegroundColor Green
                        $sharedStepsFiles = Get-ChildItem -Path $AllConfigPath -Filter "CustomSteps*.json"
                    }
                    
                    # Handle ErrorKeywords.json
                    if ($keywordsPath -and (Test-Path $keywordsPath)) {
                        $existingKeywordsNumbers = $sharedKeywordsFiles | ForEach-Object { 
                            if ($_.BaseName -match "ErrorKeywords(\d+)") { [int]$matches[1] } 
                        }
                        $nextKeywordsNumber = if ($existingKeywordsNumbers) { ($existingKeywordsNumbers | Measure-Object -Maximum).Maximum + 1 } else { 1 }
                        $newKeywordsFile = Join-Path -Path $AllConfigPath -ChildPath "ErrorKeywords$nextKeywordsNumber.json"
                        Copy-Item -Path $keywordsPath -Destination $newKeywordsFile -Force
                        Write-Host "Copied to $newKeywordsFile" -ForegroundColor Green
                        $sharedKeywordsFiles = Get-ChildItem -Path $AllConfigPath -Filter "ErrorKeywords*.json"
                    }
                    Start-Sleep -Seconds 1
                }
                "2" {
                    if ($sharedStepsFiles.Count -eq 0) {
                        Write-Host "No shared CustomSteps files found in $AllConfigPath" -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                        continue
                    }
                    Write-Host "Available Shared CustomSteps Files:" -ForegroundColor Yellow
                    for ($i = 0; $i -lt $sharedStepsFiles.Count; $i++) {
                        Write-Host "$($i+1). $($sharedStepsFiles[$i].Name)"
                    }
                    $selection = Read-Host "Enter numbers to import from (comma-separated, or 'all' for all)"
                    $toImport = if ($selection -eq "all") { $sharedStepsFiles } else { 
                        $selection -split "," | ForEach-Object { $sharedStepsFiles[[int]$_.Trim() - 1] } 
                    }

                    foreach ($file in $toImport) {
                        $sharedSteps = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json
                        foreach ($prop in $sharedSteps.PSObject.Properties) {
                            if (-not $mainSteps.PSObject.Properties.Name -contains $prop.Name) {
                                $mainSteps | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.Value
                            } else {
                                # Merge steps, keeping unique values
                                $mainSteps.$($prop.Name) = @($mainSteps.$($prop.Name) + $prop.Value | Sort-Object -Unique)
                            }
                        }
                    }
                    Write-Host "Imported steps into CustomSteps.json" -ForegroundColor Green
                    Start-Sleep -Seconds 1
                }
                "3" {
                    if ($sharedKeywordsFiles.Count -eq 0) {
                        Write-Host "No shared ErrorKeywords files found in $AllConfigPath" -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                        continue
                    }
                    Write-Host "Available Shared ErrorKeywords Files:" -ForegroundColor Yellow
                    for ($i = 0; $i -lt $sharedKeywordsFiles.Count; $i++) {
                        Write-Host "$($i+1). $($sharedKeywordsFiles[$i].Name)"
                    }
                    $selection = Read-Host "Enter numbers to import from (comma-separated, or 'all' for all)"
                    $toImport = if ($selection -eq "all") { $sharedKeywordsFiles } else { 
                        $selection -split "," | ForEach-Object { $sharedKeywordsFiles[[int]$_.Trim() - 1] } 
                    }

                    foreach ($file in $toImport) {
                        $sharedKeywords = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json
                        $mainKeywords.ErrorKeywords = @($mainKeywords.ErrorKeywords + $sharedKeywords.ErrorKeywords | Sort-Object -Unique)
                        $mainKeywords.WarningKeywords = @($mainKeywords.WarningKeywords + $sharedKeywords.WarningKeywords | Sort-Object -Unique)
                    }
                    Write-Host "Imported keywords into ErrorKeywords.json" -ForegroundColor Green
                    Start-Sleep -Seconds 1
                }
                "4" {
                    if ($sharedStepsFiles.Count -eq 0 -and $sharedKeywordsFiles.Count -eq 0) {
                        Write-Host "No shared config files found in $AllConfigPath" -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                        continue
                    }
                    # Import all steps
                    foreach ($file in $sharedStepsFiles) {
                        $sharedSteps = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json
                        foreach ($prop in $sharedSteps.PSObject.Properties) {
                            if (-not $mainSteps.PSObject.Properties.Name -contains $prop.Name) {
                                $mainSteps | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.Value
                            } else {
                                $mainSteps.$($prop.Name) = @($mainSteps.$($prop.Name) + $prop.Value | Sort-Object -Unique)
                            }
                        }
                    }
                    # Import all keywords
                    foreach ($file in $sharedKeywordsFiles) {
                        $sharedKeywords = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json
                        $mainKeywords.ErrorKeywords = @($mainKeywords.ErrorKeywords + $sharedKeywords.ErrorKeywords | Sort-Object -Unique)
                        $mainKeywords.WarningKeywords = @($mainKeywords.WarningKeywords + $sharedKeywords.WarningKeywords | Sort-Object -Unique)
                    }
                    Write-Host "Imported all steps and keywords" -ForegroundColor Green
                    Start-Sleep -Seconds 1
                }
                "5" {
                    $mainSteps | ConvertTo-Json | Set-Content -Path $mainStepsFile
                    $mainKeywords | ConvertTo-Json | Set-Content -Path $mainKeywordsFile
                    Write-Host "Changes saved to $mainStepsFile and $mainKeywordsFile" -ForegroundColor Green
                    # Update global variables to reflect changes
                    $global:ErrorKeywords = $mainKeywords.ErrorKeywords
                    $global:WarningKeywords = $mainKeywords.WarningKeywords
                    return
                }
                "6" { return }
                default { Write-Host "Invalid choice" -ForegroundColor Red; Start-Sleep -Seconds 1 }
            }
        } while ($true)
    } catch {
        Write-Host "Error in ManageWebConfig: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}

function ManageTemplates {
    param (
        [string]$TemplatePath = "$env:USERPROFILE\Desktop\EcoDesk\Template"
    )

    try {
        Clear-Host
        Write-Host "Manage Templates" -ForegroundColor Cyan
        Write-Host "================" -ForegroundColor Cyan
        
        # Create directory if it doesn’t exist
        if (-not (Test-Path $TemplatePath)) {
            New-Item -ItemType Directory -Path $TemplatePath -Force | Out-Null
        }

        # Get existing templates and assign numbers
        $templates = Get-ChildItem -Path $TemplatePath -File -Filter "*.txt" | Sort-Object Name
        $templateCount = $templates.Count

        do {
            Write-Host "`nExisting Templates:" -ForegroundColor Yellow
            if ($templates.Count -eq 0) {
                Write-Host "No templates found."
            } else {
                for ($i = 0; $i -lt $templates.Count; $i++) {
                    Write-Host "$($i+1). $($templates[$i].BaseName)"
                }
            }
            Write-Host "`nOptions:" -ForegroundColor Cyan
            Write-Host "1. Create New Template"
            Write-Host "2. Copy Template Content to Clipboard"
            Write-Host "3. Exit"

            $choice = Read-Host "Enter your choice (1-3)"
            switch ($choice) {
                "1" {
                    $title = Read-Host "Enter template title (used as filename)"
                    if ([string]::IsNullOrWhiteSpace($title)) {
                        Write-Host "Title cannot be empty" -ForegroundColor Red
                        Start-Sleep -Seconds 1
                        continue
                    }
                    $content = Read-Host "Enter template content"
                    $fileName = "$title.txt"  # Using title directly as filename
                    $newFile = Join-Path -Path $TemplatePath -ChildPath $fileName
                    $content | Set-Content -Path $newFile
                    Write-Host "Template '$title' saved as $fileName" -ForegroundColor Green
                    $templates = Get-ChildItem -Path $TemplatePath -File -Filter "*.txt" | Sort-Object Name
                    $templateCount = $templates.Count
                    Start-Sleep -Seconds 1
                }
                "2" {
                    if ($templates.Count -eq 0) {
                        Write-Host "No templates available to copy" -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                        continue
                    }
                    $selection = Read-Host "Enter template number to copy (1-$templateCount)"
                    $index = [int]$selection - 1
                    if ($index -ge 0 -and $index -lt $templates.Count) {
                        $content = Get-Content -Path $templates[$index].FullName -Raw
                        Set-Clipboard -Value $content
                        Write-Host "Content of '$($templates[$index].BaseName)' copied to clipboard" -ForegroundColor Green
                    } else {
                        Write-Host "Invalid selection" -ForegroundColor Red
                    }
                    Start-Sleep -Seconds 1
                }
                "3" { return }
                default { Write-Host "Invalid choice" -ForegroundColor Red; Start-Sleep -Seconds 1 }
            }
        } while ($true)
    } catch {
        Write-Host "Error in ManageTemplates: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}

function ScanArticles {
    param (
        [string]$ArticlePath = "$env:USERPROFILE\Desktop\EcoDesk\Articles",
        [string]$LogPath = "\\englog\escalationlogs\$global:CCID\$global:Case"
    )

    try {
        Clear-Host
        Write-Host "Scan Articles" -ForegroundColor Cyan
        Write-Host "=============" -ForegroundColor Cyan
        
        if (-not (Test-Path $ArticlePath)) {
            New-Item -ItemType Directory -Path $ArticlePath -Force | Out-Null
            Write-Host "Created directory $ArticlePath" -ForegroundColor Green
        }

        if (-not (Test-Path $LogPath)) {
            Write-Host "Default log location $LogPath not found." -ForegroundColor Yellow
            $manualLogPath = Read-Host "Please enter the manual log location (e.g., path to log files)"
            if (-not [string]::IsNullOrWhiteSpace($manualLogPath)) {
                $LogPath = $manualLogPath
                if (-not (Test-Path $LogPath)) {
                    Write-Host "Manual log location $LogPath not found. Please verify the path." -ForegroundColor Red
                    Read-Host "Press Enter to return to menu"
                    return
                }
            } else {
                Write-Host "No location entered. Exiting." -ForegroundColor Red
                Read-Host "Press Enter to return to menu"
                return
            }
        }

        Write-Host "Options:" -ForegroundColor Cyan
        Write-Host "1. Scan Local Articles"
        Write-Host "2. Scan Article from URL"
        $choice = Read-Host "Enter your choice (1-2)"

        $articleContent = ""
        if ($choice -eq "1") {
            $articles = Get-ChildItem -Path $ArticlePath -File -Filter "*.txt"
            if ($articles.Count -eq 0) {
                Write-Host "No articles found in $ArticlePath" -ForegroundColor Yellow
                Read-Host "Press Enter to return to menu"
                return
            }
            Write-Host "Available Articles:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $articles.Count; $i++) {
                Write-Host "$($i+1). $($articles[$i].Name)"
            }
            $selection = Read-Host "Enter article number to scan (1-$($articles.Count))"
            $index = [int]$selection - 1
            if ($index -ge 0 -and $index -lt $articles.Count) {
                Write-Host "Loading local article: $($articles[$index].Name)" -ForegroundColor Green
                $articleContent = Get-Content -Path $articles[$index].FullName -Raw
            } else {
                Write-Host "Invalid selection" -ForegroundColor Red
                Read-Host "Press Enter to return to menu"
                return
            }
        } elseif ($choice -eq "2") {
            $url = Read-Host "Enter article URL"
            try {
                Write-Host "Fetching content from $url..." -ForegroundColor Yellow
                $articleContent = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop | Select-Object -ExpandProperty Content
                Write-Host "Content fetched successfully from $url" -ForegroundColor Green

                $save = Read-Host "Save article locally? (Y/N)"
                if ($save.ToUpper() -eq "Y") {
                    $title = Read-Host "Enter article title for filename"
                    $fileName = "$title.txt"
                    $filePath = Join-Path -Path $ArticlePath -ChildPath $fileName
                    Write-Host "Saving content to $filePath..." -ForegroundColor Yellow
                    $maskedContent = Mask-PII -LogCut $articleContent  # Mask PII before saving
                    $maskedContent | Set-Content -Path $filePath
                    Write-Host "Article saved as $fileName with PII masked" -ForegroundColor Green
                }
            } catch {
                Write-Host "Failed to fetch article from $url $_" -ForegroundColor Red
                Read-Host "Press Enter to return to menu"
                return
            }
        } else {
            Write-Host "Invalid choice" -ForegroundColor Red
            Read-Host "Press Enter to return to menu"
            return
        }

        # ... (rest of the scanning logic unchanged, no masking here)
        Write-Host "Searching for log cuts in $LogPath..." -ForegroundColor Yellow
        $logCuts = Get-ChildItem -Path $LogPath -Recurse -Filter "*_logcut.txt" -ErrorAction Stop | ForEach-Object { 
            Write-Host "Processing log cut file: $($_.Name)" -ForegroundColor Green
            Get-Content -Path $_.FullName -Raw 
        }
        if (-not $logCuts) {
            Write-Host "No log cuts found in $LogPath. Run 'Filter logs based on JobID' first." -ForegroundColor Yellow
            Read-Host "Press Enter to return to menu"
            return
        }

        Write-Host "Scanning content for matches with log cuts..." -ForegroundColor Yellow
        $matchesFound = $false
        $matchThreshold = 0.75
        $articleWords = $articleContent -split '\s+' | Where-Object { $_ }

        foreach ($logCut in $logCuts) {
            $logCutWords = $logCut -split '\s+' | Where-Object { $_ }
            $commonWords = $logCutWords | Where-Object { $articleWords -contains $_ }
            $matchPercentage = $commonWords.Count / [Math]::Max($logCutWords.Count, $articleWords.Count)

            if ($matchPercentage -ge $matchThreshold -and $commonWords.Count -ge 3) {
                $matchesFound = $true
                Write-Host "Match found in article! (Match: $($matchPercentage*100)%)" -ForegroundColor Green
                Write-Host "Matching Log Cut: $logCut" -ForegroundColor Yellow
                # ... (rest of the options unchanged)
            }
        }
        if (-not $matchesFound) {
            Write-Host "No matches found between article and log cuts." -ForegroundColor Yellow
        }
        Write-Host "Scan completed." -ForegroundColor Green
        Read-Host "Press Enter to return to menu"
    } catch {
        Write-Host "Error in ScanArticles: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}

function ScanArticlesWithAnalysis {
    param (
        [string]$ArticlePath = "$env:USERPROFILE\Desktop\EcoDesk\Articles",
        [string]$LogPath = "\\englog\escalationlogs\$global:CCID\$global:Case",
        [string]$KBPath = "$env:USERPROFILE\Desktop\EcoDesk\KBArticles"
    )

    try {
        Clear-Host
        Write-Host "Scan Articles with Log Analysis" -ForegroundColor Cyan
        Write-Host "===============================" -ForegroundColor Cyan
        
        if (-not (Test-Path $ArticlePath)) {
            New-Item -ItemType Directory -Path $ArticlePath -Force | Out-Null
        }
        if (-not (Test-Path $KBPath)) {
            New-Item -ItemType Directory -Path $KBPath -Force | Out-Null
        }

        if (-not (Test-Path $LogPath)) {
            Write-Host "Default log location $LogPath not found." -ForegroundColor Yellow
            $manualLogPath = Read-Host "Please enter the manual log location (e.g., path to log files)"
            if (-not [string]::IsNullOrWhiteSpace($manualLogPath)) {
                $LogPath = $manualLogPath
                if (-not (Test-Path $LogPath)) {
                    Write-Host "Manual log location $LogPath not found. Please verify the path." -ForegroundColor Red
                    Read-Host "Press Enter to return to menu"
                    return
                }
            } else {
                Write-Host "No location entered. Exiting." -ForegroundColor Red
                Read-Host "Press Enter to return to menu"
                return
            }
        }

        $logAnalysisReports = Get-ChildItem -Path $LogPath -Filter "ErrorReport_*.txt" | Sort-Object LastWriteTime -Descending
        if ($logAnalysisReports.Count -eq 0) {
            Write-Host "No ErrorReport files found in $LogPath. Run 'Analyze Logs' first." -ForegroundColor Yellow
            Read-Host "Press Enter to return to menu"
            return
        }
        $latestLogAnalysis = $logAnalysisReports[0]
        $errorLines = Get-Content -Path $latestLogAnalysis.FullName | Where-Object { $_ -match "Log Line:" } | ForEach-Object { $_.Replace("Log Line: ", "") }

        if (-not $errorLines) {
            Write-Host "No errors or warnings found in $latestLogAnalysis" -ForegroundColor Yellow
            Read-Host "Press Enter to return to menu"
            return
        }

        Write-Host "Using findings from: $($latestLogAnalysis.Name)" -ForegroundColor Yellow
        Write-Host "Found $($errorLines.Count) error/warning lines to search with." -ForegroundColor Green

        $kbUrls = @()
        do {
            $url = Read-Host "Enter KB article URL (or press Enter to finish)"
            if ($url) { $kbUrls += $url }
        } while ($url)

        $matchesFound = $false
        $matchThreshold = 0.75  # 75% word match tolerance

        # Search local articles
        $localArticles = Get-ChildItem -Path $ArticlePath -File -Filter "*.txt"
        $kbArticles = Get-ChildItem -Path $KBPath -File -Filter "*.txt"
        $allLocalArticles = $localArticles + $kbArticles

        foreach ($article in $allLocalArticles) {
            $articleContent = Get-Content -Path $article.FullName -Raw
            $articleWords = $articleContent -split '\s+' | Where-Object { $_ }
            foreach ($errorLine in $errorLines) {
                $errorLineWords = $errorLine -split '\s+' | Where-Object { $_ }
                $commonWords = $errorLineWords | Where-Object { $articleWords -contains $_ }
                $matchPercentage = $commonWords.Count / [Math]::Max($errorLineWords.Count, $articleWords.Count)

                if ($matchPercentage -ge $matchThreshold -and $commonWords.Count -ge 3) {
                    $matchesFound = $true
                    Write-Host "Match found in local article: $($article.Name) (Match: $($matchPercentage*100)%)" -ForegroundColor Green
                    Write-Host "Matching Error/Warning: $errorLine" -ForegroundColor Yellow
                    HandleMatch -Content $articleContent -Error $errorLine
                }
            }
        }

        # Search KB URLs
        foreach ($url in $kbUrls) {
            try {
                $webContent = Invoke-WebRequest -Uri $url -UseBasicParsing | Select-Object -ExpandProperty Content
                $webWords = $webContent -split '\s+' | Where-Object { $_ }
                foreach ($errorLine in $errorLines) {
                    $errorLineWords = $errorLine -split '\s+' | Where-Object { $_ }
                    $commonWords = $errorLineWords | Where-Object { $webWords -contains $_ }
                    $matchPercentage = $commonWords.Count / [Math]::Max($errorLineWords.Count, $webWords.Count)

                    if ($matchPercentage -ge $matchThreshold -and $commonWords.Count -ge 3) {
                        $matchesFound = $true
                        Write-Host "Match found in KB URL: $url (Match: $($matchPercentage*100)%)" -ForegroundColor Green
                        Write-Host "Matching Error/Warning: $errorLine" -ForegroundColor Yellow
                        HandleMatch -Content $webContent -Error $errorLine -Url $url
                    }
                }
            } catch {
                Write-Host "Failed to fetch $url $_" -ForegroundColor Red
            }
        }

        if (-not $matchesFound) {
            Write-Host "No matches found in local articles or KB URLs." -ForegroundColor Yellow
        }
        Read-Host "Press Enter to return to menu"
    } catch {
        Write-Host "Error in ScanArticlesWithAnalysis: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}

function HandleMatch {
    param (
        [string]$Content,
        [string]$Error,
        [string]$Url = $null
    )

    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "1. Open in Browser (if URL provided)"
    Write-Host "2. Copy Matching Content to Clipboard"
    Write-Host "3. Display Content in PowerShell"
    Write-Host "4. Continue Scanning"
    $action = Read-Host "Enter your choice (1-4)"
    
    switch ($action) {
        "1" {
            if ($Url) {
                Start-Process "chrome.exe" $Url
            } else {
                Write-Host "No URL available for local article" -ForegroundColor Yellow
            }
        }
        "2" {
            Set-Clipboard -Value $Error
            Write-Host "Matching error copied to clipboard" -ForegroundColor Green
        }
        "3" {
            Write-Host "Matching Content:" -ForegroundColor Yellow
            Write-Host $Content
            Read-Host "Press Enter to continue"
        }
        "4" { return }
        default { Write-Host "Invalid choice" -ForegroundColor Red }
    }
}

function CreateReferenceLog {
    param (
        [string]$RefFolder = "\\englog\escalationlogs\references"
    )

    try {
        Clear-Host
        Write-Host "Create Reference Log" -ForegroundColor Cyan
        Write-Host "===================" -ForegroundColor Cyan
        
        # Create directory if it doesn’t exist
        if (-not (Test-Path $RefFolder)) {
            New-Item -ItemType Directory -Path $RefFolder -Force | Out-Null
        }

        $logCut = Read-Host "Enter the Log Cut (e.g., 'Error: Connection failed')"
        $solution = Read-Host "Enter the Solution (e.g., 'Check network connectivity and restart the service')"
        
        if ([string]::IsNullOrWhiteSpace($logCut) -or [string]::IsNullOrWhiteSpace($solution)) {
            Write-Host "Log Cut and Solution cannot be empty. Exiting." -ForegroundColor Red
            Read-Host "Press Enter to return to menu"
            return
        }

        $fileName = "ref_$(Get-Date -Format 'yyyyMMddHHmmss').txt"
        $refFile = Join-Path -Path $RefFolder -ChildPath $fileName
        "Log Cut: $logCut`nSolution: $solution" | Set-Content -Path $refFile
        Write-Host "Reference log saved as $refFile" -ForegroundColor Green
        Read-Host "Press Enter to return to menu"
    } catch {
        Write-Host "Error in CreateReferenceLog: $_" -ForegroundColor Red
        Read-Host "Press Enter to return to menu"
    }
}

function Mask-PII {
    param (
        [string]$LogCut
    )

    if (-not $LogCut) { return $LogCut }

    $patterns = @{
        "ServerName" = "\b(?:[a-zA-Z0-9-]+\.)*[a-zA-Z0-9-]+(?:\.com|\.org|\.net|\.local)?\b(?<!\b(?:error|fail|exception|warning|log|file))\b"
        "Username"   = "\b[a-zA-Z][a-zA-Z0-9_-]*(?:\.[a-zA-Z0-9_-]+)?\b(?<!\b(?:error|fail|exception|warning|log|file))"
    }
    $replacements = @{
        "ServerName" = "[SERVER]"
        "Username"   = "[USERNAME]"
    }

    $maskedLogCut = $LogCut
    foreach ($key in $patterns.Keys) {
        $maskedLogCut = $maskedLogCut -replace $patterns[$key], $replacements[$key]
    }

    return $maskedLogCut
}


function Show-Menu {
    Clear-Host
    Write-Host "**************************************************************"
    Write-Host "* Select an option:                                          *"
    Write-Host "* ==================                                         *"
    Write-Host "* 1. Solution Tree Troubleshooting                           *"
    Write-Host "* 2. Generate Rollup Summary                                 *"
    Write-Host "* 3. Logs and Staging Management                             *"
    Write-Host "* 4. To Analyse Logs                                         *"
    Write-Host "* 5. Escalation Options                                      *"
    Write-Host "* 6. Manage Options                                          *"
    Write-Host "* 7. Change CCID/Case                                        *"
    Write-Host "* 8. Exit this Script                                        *"
    Write-Host "**************************************************************"
}

function Show-ManageOptions {
    Clear-Host
    Write-Host "**************************************************************"
    Write-Host "* Manage Options:                                            *"
    Write-Host "* ===============                                            *"
    Write-Host "* 1. Manage Troubleshooting Steps                            *"
    Write-Host "* 2. Update Error Keywords                                   *"
    Write-Host "* 3. Create KB Article                                       *"
    Write-Host "* 4. Manage WebConfig Files                                  *"
    Write-Host "* 5. Manage Templates                                        *"
    Write-Host "* 6. Create Reference Log                                    *"
    Write-Host "* 7. Back to Main Menu                                       *"
    Write-Host "**************************************************************"
}

# Load error keywords from config file at startup (unchanged)
$keywordFile = "$env:USERPROFILE\Desktop\EcoDesk\Config\ErrorKeywords.json"
if (Test-Path $keywordFile) {
    try {
        $keywords = Get-Content -Path $keywordFile -Raw | ConvertFrom-Json -ErrorAction Stop
        $global:ErrorKeywords = $keywords.ErrorKeywords
        $global:WarningKeywords = $keywords.WarningKeywords
    } catch {
        Write-Host "Failed to load error keywords from $keywordFile $_" -ForegroundColor Red
        Write-Host "Using default keywords." -ForegroundColor Yellow
        $global:ErrorKeywords = @("error", "fail", "exception", "critical", "fatal")
        $global:WarningKeywords = @("warning", "caution")
    }
} else {
    $global:ErrorKeywords = @("error", "fail", "exception", "critical", "fatal")
    $global:WarningKeywords = @("warning", "caution")
}

function Show-AnalyseLogs {
    Clear-Host
    Write-Host "**************************************************************"
    Write-Host "* To Analyse Logs:                                           *"
    Write-Host "* ================                                           *"
    Write-Host "* 1. Filter logs based on JobID                              *"
    Write-Host "* 2. Analyze Logs                                            *"
    Write-Host "* 3. Scan Articles                                           *"
    Write-Host "* 4. Scan Articles with Log Analysis                         *"
    Write-Host "* 5. Back to Main Menu                                       *"
    Write-Host "**************************************************************"
}

function Show-LogsAndStaging {
    Clear-Host
    Write-Host "**************************************************************"
    Write-Host "* Logs and Staging Management:                               *"
    Write-Host "* ===========================                                *"
    Write-Host "* 1. Extract 7z Logs from CELogs/Titan                       *"
    Write-Host "* 2. Extract Logs from Manual Path                           *"
    Write-Host "* 3. Check for DMP & Open Staging URL in Chrome              *"
    Write-Host "* 4. Back to Main Menu                                       *"
    Write-Host "**************************************************************"
}

function Show-EscalationOptions {
    Clear-Host
    Write-Host "**************************************************************"
    Write-Host "* Escalation Options:                                        *"
    Write-Host "* ==================                                         *"
    Write-Host "* 1. Escalate to Tier 2                                      *"
    Write-Host "* 2. Escalate to Development Team                            *"
    Write-Host "* 3. Back to Main Menu                                       *"
    Write-Host "**************************************************************"
}

# Initial CCID/Case input (unchanged)
do {
    Clear-Host
    do {
        $global:CCID = Read-Host -Prompt "Enter CCID"
        if ($global:CCID -match "\s") {
            Write-Host "CCID cannot contain spaces. Please enter again."
        }   
    } while ($global:CCID -match "\s")
    $global:Case = Read-Host -Prompt "Enter Ticket Number"
} while ([string]::IsNullOrWhiteSpace($global:CCID) -or [string]::IsNullOrWhiteSpace($global:Case))

# Main loop with submenu navigation
$mainChoice = $null
while ($mainChoice -ne "8") {
    Show-Menu
    $mainChoice = Read-Host "Enter your choice (1-8)"
    
    switch ($mainChoice) {
        "1" { SolutionTree }
        "2" { RollupSummary }
        "3" {  
            $subChoice = $null
            while ($subChoice -ne "4") {
                Show-LogsAndStaging
                $subChoice = Read-Host "Enter your choice (1-4)"
                switch ($subChoice) {
                    "1" { 
                        if (-not $global:CCID -or -not $global:Case) { Main } 
                        else { Main -CCID $global:CCID -Case $global:Case }
                    }
                    "2" { Extraction }
                    "3" { CheckForDMPFiles -Destination "\\englog\escalationlogs\$global:CCID\$global:Case" }
                    "4" { break }
                    default { Write-Host "Invalid option. Please try again." }
                }
                if ($subChoice -ne "4") { Read-Host "Press Enter to continue..." }
            }
        }
        "4" {  
            $subChoice = $null
            while ($subChoice -ne "5") {
                Show-AnalyseLogs
                $subChoice = Read-Host "Enter your choice (1-5)"
                switch ($subChoice) {
                    "1" { SearchLogs }
                    "2" { AnalyzeLogs }
                    "3" { ScanArticles }
                    "4" { ScanArticlesWithAnalysis }
                    "5" { break }
                    default { Write-Host "Invalid option. Please try again." }
                }
                if ($subChoice -ne "5") { Read-Host "Press Enter to continue..." }
            }
        }
        "5" {  
            $subChoice = $null
            while ($subChoice -ne "3") {
                Show-EscalationOptions
                $subChoice = Read-Host "Enter your choice (1-3)"
                switch ($subChoice) {
                    "1" { EscalateToTier2 }
                    "2" { EscalateToDevTeam }
                    "3" { break }
                    default { Write-Host "Invalid option. Please try again." }
                }
                if ($subChoice -ne "3") { Read-Host "Press Enter to continue..." }
            }
        }
        "6" {  
            $subChoice = $null
            while ($subChoice -ne "7") {
                Show-ManageOptions
                $subChoice = Read-Host "Enter your choice (1-7)"
                switch ($subChoice) {
                    "1" { ManageTroubleshootingSteps }
                    "2" { UpdateErrorKeywords }
                    "3" { CreateKBArticle }
                    "4" { ManageWebConfig }
                    "5" { ManageTemplates }
                    "6" { CreateReferenceLog }
                    "7" { break }
                    default { Write-Host "Invalid option. Please try again." }
                }
                if ($subChoice -ne "7") { Read-Host "Press Enter to continue..." }
            }
        }
        "7" {  
            Clear-Host
            do {
                $global:CCID = Read-Host -Prompt "Enter new CCID"
                if ($global:CCID -match "\s") {
                    Write-Host "CCID cannot contain spaces. Please enter again."
                }   
            } while ($global:CCID -match "\s")
            $global:Case = Read-Host -Prompt "Enter new Ticket Number"
            Write-Host "CCID and Case updated successfully."
        }
        "8" { Write-Host "Exiting..."; break }
        default { Write-Host "Invalid option. Please try again." }
    }
    if ($mainChoice -ne "8") { Read-Host "Press Enter to continue..." }
}
