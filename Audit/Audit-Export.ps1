param (
    [string]$outputExcel = $(Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "SystemInfo.xlsx")
)

# Paths and Configurations
$dbPath = "X:/Audit/Databases"
$mainDbPath = Join-Path -Path $env:TEMP -ChildPath "MainDatabase.db"
$sqlitePath = Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "sqlite3.exe"
if (-not (Test-Path $dbPath)) {
    $script_directory = Split-Path -Parent $MyInvocation.MyCommand.Path
    $dbPath = Join-Path $script_directory "Databases"
}

# Don't modify below here unless you know what your doing.
$maxRetries = 5
$retryInterval = 5  # In seconds

# Ensure SQLite is installed
if (!(Test-Path $sqlitePath)) {
    Write-Host "ERROR: sqlite3.exe not found. Please install SQLite and update the script."
    Read-Host -Prompt "Press any key to continue"
    exit 1
}

# Ensure ImportExcel module is available
if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module is missing. Attempting to install..."

    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "Script needs to be run as Administrator to install the ImportExcel module."
        Start-Process powershell.exe "-Command Start-Process powershell.exe -ArgumentList 'Install-Module -Name ImportExcel -Force' -Verb RunAs" -Wait
        Write-Host "Re-running the script..."
        Start-Process powershell.exe "-File `"$($MyInvocation.MyCommand.Path)`" -ArgumentList @('$outputExcel')" -Wait
        exit
    } else {
        Install-Module -Name ImportExcel -Force
    }
}

Import-Module ImportExcel

# Function to create the main database if it doesn't exist
function Initialize-Database {
    if (!(Test-Path $mainDbPath)) {
        Write-Host "Main database not found. Creating a new one..."
        & $sqlitePath $mainDbPath "CREATE TABLE IF NOT EXISTS SystemInfo (
            System_Name TEXT PRIMARY KEY, Last_User TEXT, Boot_Time TEXT, AD_OU TEXT, System_Model TEXT, Serial_Number TEXT, 
            Processor TEXT, OS_Drive TEXT, Memory TEXT, BIOS_Version TEXT, MAC_Address TEXT, IP_Address TEXT, 
            WiFi_Status TEXT, OS_Version TEXT, OS_Architecture TEXT, OS_Build TEXT, Windows_Update_Date TEXT, Printers TEXT, Display1 TEXT, Display1_Serial TEXT, 
            Display2 TEXT, Display2_Serial TEXT, Display3 TEXT, Display3_Serial TEXT, BitLocker_Status TEXT, Device_Errors TEXT, Updated_Time TEXT
        );"
        Write-Host "Main database created successfully."
    }
}

# Function to close an open Excel workbook
function Close-ExcelWorkbook {
    param ($filePath)

    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $workbook = $excel.Workbooks | Where-Object { $_.FullName -eq $filePath }

        if ($workbook) {
            $workbook.Close($false)
            Write-Host "Closed open Excel workbook: $filePath"
        }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }catch {
        #Write-Host "No open instance of Excel found."
    }
    if (Test-Path $outputExcel) {
        Remove-Item -Path $outputExcel -Force
    }
}

# Function to merge a local database into the main database
function Merge-Database($localDbPath) {
    #Write-Host "Merging $localDbPath into main database..."
    try {
        & $sqlitePath $mainDbPath "
        ATTACH DATABASE '$localDbPath' AS LocalDB;
        BEGIN TRANSACTION;
        INSERT OR REPLACE INTO SystemInfo SELECT * FROM LocalDB.SystemInfo;
        COMMIT;
        DETACH DATABASE LocalDB;"
        Write-Host "Merged $localDbPath successfully."
    } catch {
        Write-Host "ERROR: Failed to merge $localDbPath. Error: $_"
    }
}

# Function to query SQLite database and return structured data
function Get-DatabaseData {
    Write-Host "Retrieving database records..."

    # Fetch column names dynamically
    $columnInfo = & $sqlitePath $mainDbPath "PRAGMA table_info(SystemInfo);"
    $columns = $columnInfo -split "`n" | ForEach-Object { ($_ -split "\|")[1] } | Where-Object { $_ -ne $null -and $_ -ne "" }

    if ($columns.Count -eq 0) {
        Write-Host "ERROR: Failed to retrieve columns from the database. Ensure SystemInfo table exists."
        Read-Host -Prompt "Press any key to continue"
        exit 1
    }

    # Get the latest record for each unique System_Name and Serial_Number
    $latestData = & $sqlitePath $mainDbPath "
    SELECT * FROM SystemInfo 
    WHERE Updated_Time = (
        SELECT MAX(Updated_Time) FROM SystemInfo AS S 
        WHERE S.System_Name = SystemInfo.System_Name OR S.Serial_Number = SystemInfo.Serial_Number
    );"

    # Get superseded (older) records
    $replacedData = & $sqlitePath $mainDbPath "
    SELECT * FROM SystemInfo 
    WHERE (System_Name, Updated_Time) NOT IN (
        SELECT System_Name, MAX(Updated_Time) FROM SystemInfo GROUP BY System_Name
    )
    OR (Serial_Number, Updated_Time) NOT IN (
        SELECT Serial_Number, MAX(Updated_Time) FROM SystemInfo GROUP BY Serial_Number
    );"

    return @{
        Latest    = Convert-SQLiteOutput -data $latestData -columns $columns
        Replaced  = Convert-SQLiteOutput -data $replacedData -columns $columns
    }
}

# Function to convert SQLite output into PowerShell objects
function Convert-SQLiteOutput {
    param ($data, $columns)
    
    $parsedData = @()
    foreach ($line in $data -split "`n") {
        $values = $line -split "\|"
        while ($values.Length -lt $columns.Length) { $values += "N/A" }  # Pad missing values
        if ($values.Length -eq $columns.Length) {
            $obj = [PSCustomObject]@{}
            for ($i = 0; $i -lt $columns.Length; $i++) {
                $obj | Add-Member -MemberType NoteProperty -Name $columns[$i] -Value $values[$i]
            }
            $parsedData += $obj
        }
    }
    return $parsedData
}

# Modify the Export function to handle multiple sheets
function Export-ToExcel($latestData, $replacedData, $filePath) {
    try {
        Close-ExcelWorkbook -filePath $filePath
        $latestData | Export-Excel -Path $filePath -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName "SystemInfo"
        $replacedData | Export-Excel -Path $filePath -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName "Replaced" -Append
        Write-Host "System information exported successfully to $filePath"
    } catch {
        Write-Host "ERROR: Failed to export to Excel. Error: $_"
    }
}

### --- SCRIPT EXECUTION STARTS HERE --- ###
Write-Host "Starting database audit and export..."

# Initialize the main database
if (Test-Path $mainDbPath) {
    Remove-Item -Path $mainDbPath -Force
}
Initialize-Database

# Merge all databases
$localDbs = Get-ChildItem -Path $dbPath -Filter "*.db"
foreach ($dbFile in $localDbs) {
    $localDbPath = $dbFile.FullName
    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            Merge-Database $localDbPath
            break  # Exit retry loop on success
        } catch {
            Write-Host "Attempt $i of $maxRetries failed for $localDbPath. Retrying in $retryInterval seconds..."
            if ($i -eq $maxRetries) {
                Write-Host "ERROR: Failed to merge $localDbPath after $maxRetries attempts."
            } else {
                Start-Sleep -Seconds $retryInterval
            }
        }
    }
}

Write-Host "All databases merged successfully!"
# Fetch data and export to Excel
$data = Get-DatabaseData
Export-ToExcel -latestData $data.Latest -replacedData $data.Replaced -filePath $outputExcel