$db_directory = "X:/Audit/Databases"
# Check if the script is running as administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Script is not running as administrator. Restarting with elevated privileges..."
    $newProcess = Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -PassThru -WindowStyle Hidden
    $newProcess.WaitForExit()
    exit
}

function Get-ComputerName {
    return $env:COMPUTERNAME
}

function Get-SchoolDrive {
    $computerName = Get-ComputerName
    $nameSplit = $computerName -split "-"
    if ($nameSplit) {
        return "\\ad.ccrsb.ca\xadmin-$($nameSplit[0])"
    }
    return ""
}

if ($db_directory.StartsWith("X:")) {
    $db_directory = $db_directory -replace "^X:", (Get-SchoolDrive)
}
$script_directory = Split-Path -Parent $MyInvocation.MyCommand.Path
$db_directory = if (Test-Path $db_directory) { $db_directory } else { Join-Path $script_directory "Databases" }
Write-Host $db_directory 
$localDbPath = Join-Path $db_directory "$env:COMPUTERNAME.db"

# Function to get the current username
function GetCurrentUsername {
    return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[-1]
}

# Map network drive to H:\
$username = & GetCurrentUsername
if ($username.StartsWith("x-")) {
    $username = $username.Substring(2)
}
$networkPath = "\\ad.ccrsb.ca\it-home\IT-SCHOOL-HOME\" + $username
Write-Host $networkPath

# Disconnect any existing mapping for H:
$netUseDeleteCommand = "net use H: /delete /y"
Invoke-Expression $netUseDeleteCommand

if (Test-Path $networkPath) {
    # Use net use command to map the network drive
    $netUseCommand = "net use H: $networkPath /persistent:yes"
    Invoke-Expression $netUseCommand
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Failed to map network drive H: to $networkPath."
        exit
    }
} else {
    Write-Host "Network path $networkPath is not accessible."
    exit
}
$sqlitePath = Join-Path -Path $networkPath -ChildPath "Login\Audit\sqlite3.exe"
if (-not (Test-Path $db_directory )) {
    $script_directory = Split-Path -Parent $MyInvocation.MyCommand.Path
    $db_directory = Join-Path $script_directory "Databases"
}
# The following code handles critical operations such as database setup and system information retrieval.
# Modifying this section without understanding its functionality may cause errors or data loss.
$tempFolder = [System.IO.Path]::GetTempPath()
$tempSqlitePath = Join-Path $tempFolder "sqlite3.exe"

$folderPath = [System.IO.Path]::GetDirectoryName($localDbPath)
# Check if the folder exists, and create it if it doesn't
if (-not (Test-Path -Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath -Force
}

# Ensure SQLite exists
if (!(Test-Path $sqlitePath)) {
    Write-Host "sqlite3.exe not found at $sqlitePath. Please install SQLite by following these steps:`n1. Download sqlite3.exe from https://www.sqlite.org/download.html`n2. Place sqlite3.exe in the same directory as this script or update the script with the correct path."
    Write-Host "sqlite3.exe not found. Please install SQLite and update the script. $sqlitePath"
    Read-Host -Prompt "Press any key to continue"
    exit
}

Copy-Item -Path $sqlitePath -Destination $tempSqlitePath -Force
if (Test-Path $localDbPath) {
    & $tempSqlitePath $localDbPath "DROP TABLE IF EXISTS AdminSystemInfo;"
}

try {
    $null = & $tempSqlitePath $localDbPath "PRAGMA journal_mode=WAL;"

    # Ensure table exists
    & $tempSqlitePath $localDbPath """
    CREATE TABLE IF NOT EXISTS AdminSystemInfo (
        System_Name TEXT, 
        Win11_Compatible TEXT, 
        Incompatible_Items TEXT, 
        Display1 TEXT, 
        Display1_Serial TEXT, 
        Display2 TEXT, 
        Display2_Serial TEXT, 
        Display3 TEXT, 
        Display3_Serial TEXT, 
        Drivers TEXT,
        Updated_Time TEXT
    );"""

    # Get System info
    $SystemName = $env:COMPUTERNAME

    # Run HardwareReadiness.ps1
    Write-Host "Run Hardware Readiness for Windows 11"
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
    & (Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "HardwareReadiness.ps1") | Out-Null

    # Read registry values
    $win11Compatible = Get-ItemProperty -Path "HKLM:\SOFTWARE\CCRCE\Win11" -Name "WIN11COMPATIBLE" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty WIN11COMPATIBLE -ErrorAction SilentlyContinue
    if ($null -eq $win11Compatible) {
        $win11Compatible = "N/A"
    }
    $incompatibleItems = Get-ItemProperty -Path "HKLM:\SOFTWARE\CCRCE\Win11" -Name "IncompatibleItems" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IncompatibleItems -ErrorAction SilentlyContinue
    if ($null -eq $incompatibleItems) {
        $incompatibleItems = "N/A"
    }
        
    # Define the path to HPIA executable and report
    $tempHpiaFolder = Join-Path -Path $tempFolder -ChildPath "hp-hpia-5.3.1"
    Copy-Item -Path (Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "hp-hpia-5.3.1") -Destination $tempHpiaFolder -Recurse -Force
    $hpiaPath = Join-Path -Path $tempHpiaFolder -ChildPath "HPImageAssistant.exe"
    #Write-Host $hpiaPath
    # Find the latest XML report file in the specified directory
    $reportDirectory = "C:\CCRCE\Logs"

    # Run HPIA silently and analyze for critical updates
    Write-Host "Run HPIA silently and analyze for critical updates"
    Start-Process -FilePath $hpiaPath -ArgumentList "/Operation:Analyze /Action:List /Category:All /Silent /ResultFolderPath:$reportDirectory" -Wait -NoNewWindow

    # Wait for the report to be generated
    Start-Sleep -Seconds 10

    $reportPath = Get-ChildItem -Path $reportDirectory -Filter *.xml -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
    if (-not $reportPath) {
        Write-Host "No report file found in $reportDirectory."
    } else {
        # Load the report XML
        Write-Host "Load the report XML"
        [xml]$report = Get-Content $reportPath -ErrorAction Stop

        # Extract and display the Softpaq name, target version, and reference version
        $recommendations = $report.HPIA.Recommendations
        $Drivers = ""
        foreach ($category in $recommendations.ChildNodes) {
            foreach ($recommendation in $category.Recommendation) {
                $softpaqName = $recommendation.Solution.Softpaq.Name
                $targetVersion = $recommendation.TargetVersion
                $referenceVersion = $recommendation.ReferenceVersion
                $Drivers += "Name: $($softpaqName) (Current Version: $($targetVersion), Available Version: $($referenceVersion))`n"
            }
        }
        $Drivers = $Drivers.TrimEnd("`n")
    }
    
    # Get Display Info
    Write-Host "Get Display Info"
    $MonitorsCMD = Get-WmiObject WmiMonitorID -Namespace root\wmi
    $DisplayData = @()

    For ($i = 0; $i -lt $MonitorsCMD.Count; $i++) {
        $Name = ""; $Serial = ""
        $nm = $MonitorsCMD[$i].UserFriendlyName -notmatch '^0$'
        If ($nm -is [System.Array]) { $Name = ($nm | ForEach-Object {[char]$_}) -join "" }
        $sr = $MonitorsCMD[$i].SerialNumberID -notmatch '^0$'
        If ($sr -is [System.Array]) { $Serial = ($sr | ForEach-Object {[char]$_}) -join "" }
        
        $DisplayData += [PSCustomObject]@{ Name = $Name; Serial = $Serial }
    }

    # Assign variables for each display
    if ($DisplayData.Count -ge 1) {
        $Display1 = $DisplayData[0].Name
        $Display1Serial = $DisplayData[0].Serial
    }
    if ($DisplayData.Count -ge 2) {
        $Display2 = $DisplayData[1].Name
        $Display2Serial = $DisplayData[1].Serial
    }
    if ($DisplayData.Count -ge 3) {
        $Display3 = $DisplayData[2].Name
        $Display3Serial = $DisplayData[2].Serial
    }
        
    # Report when this was run
    $UpdatedTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Insert data into SQLite
    Write-Host "Insert data into SQLite"
    & $tempSqlitePath $localDbPath """
        INSERT OR REPLACE INTO AdminSystemInfo (System_Name, Win11_Compatible, Incompatible_Items, Display1, Display1_Serial, Display2, Display2_Serial, Display3, Display3_Serial, Drivers, Updated_Time)
        VALUES ('$SystemName', '$win11Compatible', '$incompatibleItems', '$Display1', '$Display1Serial', '$Display2', '$Display2Serial', '$Display3', '$Display3Serial', '$Drivers', '$UpdatedTime');
        """
    Write-Host "System information saved to $localDbPath."
} finally {
    try {
        Remove-Item -Path $tempSqlitePath -Force
    } catch {
        Write-Host "Failed to remove temporary SQLite file: $($_.Exception.Message)"
    }
}