$sqlitePath = Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "sqlite3.exe"
$db_directory = "X:/Audit/Databases"
if (-not (Test-Path $db_directory)) {
	$script_directory = Split-Path -Parent $MyInvocation.MyCommand.Path
	$db_directory = Join-Path $script_directory "Databases"
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

$localDbPath = Join-Path $db_directory "$env:COMPUTERNAME.db"

$ignoredPrinters = @(
    "CutePDF Writer", "Microsoft XPS Document Writer", "mimio Print Capture", "Fax",
    "Send To OneNote 2007", "Foxit Reader PDF Printer", "Microsoft Office Document Image Writer",
    "SMART Notebook Print Capture", "Microsoft Print to PDF", "Send To OneNote 2013",
    "Send To OneNote 2016", "Send To OneNote 2019", "OneNote (Desktop)", "OneNote for Windows 10", "Win2PDF", "Win2Image", "Inspiration 9 PDF Driver"
)

# Don't modify below here unless you know what your doing.
$tempFolder = [System.IO.Path]::GetTempPath()
$tempSqlitePath = Join-Path $tempFolder "sqlite3.exe"

$folderPath = [System.IO.Path]::GetDirectoryName($localDbPath)
# Check if the folder exists, and create it if it doesn't
if (-not (Test-Path -Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath -Force
}

# Ensure SQLite exists
if (!(Test-Path $sqlitePath)) {
    Write-Host "sqlite3.exe not found. Please install SQLite and update the script. $sqlitePath"
	Read-Host -Prompt "Press any key to continue"
    exit
}

Copy-Item -Path $sqlitePath -Destination $tempSqlitePath -Force
if (Test-Path $localDbPath) {
	& $tempSqlitePath $localDbPath "DROP TABLE IF EXISTS SystemInfo;"
}

try {
	$null = & $tempSqlitePath $localDbPath "PRAGMA journal_mode=WAL;"

	# Ensure table exists
	& $tempSqlitePath $localDbPath """
	CREATE TABLE IF NOT EXISTS SystemInfo (
			System_Name TEXT, Last_User TEXT, Boot_Time TEXT, AD_OU TEXT, System_Model TEXT, Serial_Number TEXT PRIMARY KEY, 
			Processor TEXT, OS_Drive TEXT, Memory TEXT, BIOS_Version TEXT, MAC_Address TEXT, IP_Address TEXT, 
			WiFi_Status TEXT, OS_Version TEXT, OS_Architecture TEXT, OS_Build TEXT, Windows_Update_Date TEXT, Printers TEXT, Display1 TEXT, Display1_Serial TEXT, 
			Display2 TEXT, Display2_Serial TEXT, Display3 TEXT, Display3_Serial TEXT, BitLocker_Status TEXT, Device_Errors TEXT, Updated_Time TEXT
		);
		"""
	# Get System info
    $SystemName = $env:COMPUTERNAME
    $SystemModel = (Get-WmiObject Win32_ComputerSystem).Model
    $SerialNumber = (Get-WmiObject Win32_BIOS).SerialNumber
	$Processor = (Get-WmiObject Win32_Processor).Name
	$bios = Get-WmiObject Win32_BIOS
	$releaseDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($bios.ReleaseDate)
	$BIOSVersion = $bios.SMBIOSBIOSVersion + " (Release date: " + $releaseDate.ToString("yyyy-MM-dd") + ")"
    $OSArchitecture = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
	$memoryInfo = Get-WmiObject -Class Win32_PhysicalMemory
	$totalMemory = 0
	$memoryDetails = @()
	foreach ($dimm in $memoryInfo) {
		$capacityGB = [math]::round($dimm.Capacity / 1GB, 2)
		$memoryType = switch ($dimm.SMBIOSMemoryType) {
			20 { "DDR" }
			21 { "DDR2" }
			24 { "DDR3" }
			26 { "DDR4" }
			27 { "DDR5" }
			28 { "DDR6" }
			default { "Unknown" }
		}
		$memoryDetails += "$($dimm.DeviceLocator) :$capacityGB GB ($memoryType)"
		$totalMemory += $capacityGB
	}
	$Memory = "$totalMemory GB ("+ ($memoryDetails -join ", ") + ")"
	$CTotal = "{0}GB" -f ([math]::Round((Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object -ExpandProperty Size) / 1GB))
	$CFree = "{0}GB" -f ([math]::Round((Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object -ExpandProperty FreeSpace) / 1GB))
	$CUsed = "{0}GB" -f ([math]::Round(([math]::Round((Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object -ExpandProperty Size) / 1GB)) - ([math]::Round((Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object -ExpandProperty FreeSpace) / 1GB))))
	$DriveType = "Unknown"
	Get-PhysicalDisk | ForEach-Object { 
		$physicalDisk = $_ 
		$partitions = $physicalDisk | Get-Disk | Get-Partition | Where-Object DriveLetter -eq 'C'
		if ($partitions) {
			$DriveType = $physicalDisk.MediaType
			Write-Host $DriveType
		}
	}
	$OS_Drive = "Total size: (" + $CTotal + ") Used: (" + $CUsed + ") Free space: (" +  $CFree + ") Drive type: (" + $DriveType + ")"

    # Get Last Logged in User
    $LastUserName = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" | Select-Object -ExpandProperty LastLoggedOnUser
    $LastUserNameDisplay = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" | Select-Object -ExpandProperty LastLoggedOnDisplayName
    $LastUser = "$LastUserName ($LastUserNameDisplay)"
    
	# Get the last boot time
	$lastBootTime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
	$uptime = (Get-Date) - $lastBootTime
	$days = $uptime.Days
	$hours = $uptime.Hours
	$minutes = $uptime.Minutes
	$seconds = $uptime.Seconds
	$bootTime = "Last Boot: $days days, $hours hours, $minutes minutes, $seconds seconds"
    # Active Directory OU
    $DomainRole = (Get-WmiObject Win32_ComputerSystem).DomainRole
    $AD_OU = if ($DomainRole -ge 1) {
        try { ([adsisearcher]"(&(name=$env:COMPUTERNAME)(objectClass=computer))").FindOne().Path -replace '^LDAP://[^,]+,', '' -replace 'OU=|DC=', '' } catch { "Error Retrieving OU" }
    } else { "Not Joined to AD" }
    
    # OS Version
    $RawOSVersion = (Get-WmiObject Win32_OperatingSystem).Version
    $OSVersion = if ($RawOSVersion -like "10.*") { "Win10" } elseif ($RawOSVersion -like "11.*") { "Win11" } else { "Unknown" }
    $OSBuild = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion
    $WinUpdateDate = (Get-WmiObject Win32_QuickFixEngineering | Sort-Object -Property InstalledOn | Select-Object -Last 1).InstalledOn.ToString("yyyy-MM-dd")
    
	# Get network info
	$MACAddress = (Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.MACAddress -and $_.Description -notlike "*Virtual*" -and $_.Description -notlike "*Bluetooth*" -and $_.Description -notlike "*WAN*" } |
    ForEach-Object {
        $connectionType = if ($_.Description -like "*Wireless*" -or $_.Description -like "*Wi-Fi*" -or $_.Description -like "*802.11*" ) { "WLAN" } else { "LAN" }
        "{0} ({1})" -f $_.MACAddress, $connectionType
    }) -join ", "
	$IPAddresses = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike "169.*" -and $_.IPAddress -notlike "127.*"}
	$IPAddress = ($IPAddresses | ForEach-Object { "$(($_.IPAddress)) ($(($_.InterfaceAlias)))" }) -join ", "
	$wifiAdapter = Get-NetAdapter | Where-Object { $_.Name -like "*Wi-Fi*" }
	if ($wifiAdapter.Status -eq "Up") {
		$netshOutput = netsh wlan show interfaces
		$ssid = ($netshOutput | Select-String -Pattern "SSID" | Select-Object -First 1).Line.Split(":")[1].Trim()
		$WiFiStatus = "Connected: ($ssid)"
	} else {
		$WiFiStatus = "Disconnected"
	}
	# Get Display Info
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
	
	# Get Printers (excluding ignored ones) and format as Name(PortName, DriverName)
	$Printers = Get-Printer | Where-Object { $ignoredPrinters -notcontains $_.Name } |
		Select-Object @{Name='PrinterInfo'; Expression={ "$($_.Name)($($_.PortName))" }} 

	# Join the printer information into a single string
	$PrintersList = $Printers.PrinterInfo -join ", "
    
	# Get Bitlocker Status
	$BitlockerStatus = "Unknown"
	try {
		$shell = New-Object -ComObject Shell.Application
		$bitlockerProtection = $shell.NameSpace('C:').Self.ExtendedProperty('System.Volume.BitLockerProtection')
		switch ($bitlockerProtection) {
			0 { $BitlockerStatus = "None" }
			1 { $BitlockerStatus = "Encrypted" }
			2 { $BitlockerStatus = "Locked" }
			default { $BitlockerStatus = "Unknown" }
		}
	} catch {
		Write-Error "Failed to retrieve BitLocker status: $_"
	}
	# Get device driver errors
	$BasicDisplayDriver = Get-WmiObject Win32_VideoController | Where-Object { $_.Name -like "*Microsoft Basic Display Adapter*" }
	$DeviceErrors = Get-PnpDevice -PresentOnly | Where-Object { 
		$_.Status -ne "OK" -and 
		$_.FriendlyName -notlike "*Keyboard*" -and 
		$_.FriendlyName -notlike "*Mouse*" 
	} |
		Select-Object @{Name='DeviceInfo'; Expression={ "$( $_.FriendlyName )($( $_.InstanceId ))" }}

	$DeviceErrorList = $DeviceErrors.DeviceInfo -join ", "
	if ($BasicDisplayDriver) {
		if ($DeviceErrorList) {
			$DeviceErrorList = "Microsoft Basic Display Adapter, " + $DeviceErrorList
		} else {
			$DeviceErrorList = "Microsoft Basic Display Adapter"
		}
	}
		
	# Report when this was run
    $UpdatedTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

	# Insert data into SQLite
	& $tempSqlitePath $localDbPath """
	INSERT OR REPLACE INTO SystemInfo (System_Name, Last_User, Boot_Time, AD_OU, System_Model, Serial_Number, Processor, OS_Drive, Memory, 
		BIOS_Version, MAC_Address, IP_Address, WiFi_Status, OS_Version, OS_Architecture, OS_Build, Windows_Update_Date, Printers, Display1, Display1_Serial, 
		Display2, Display2_Serial, Display3, Display3_Serial, BitLocker_Status, Device_Errors, Updated_Time)
	VALUES ('$SystemName', '$LastUser', '$bootTime', '$AD_OU', '$SystemModel', '$SerialNumber', '$Processor', '$OS_Drive', '$Memory', 
		'$BIOSVersion', '$MACAddress', '$IPAddress', '$WiFiStatus', '$OSVersion', '$OSArchitecture', '$OSBuild', '$WinUpdateDate', '$PrintersList', '$Display1', '$Display1Serial', 
		'$Display2', '$Display2Serial', '$Display3', '$Display3Serial', '$BitlockerStatus', '$DeviceErrorList', '$UpdatedTime');
	"""
	Write-Host "System information saved to $localDbPath."
} finally {
    Remove-Item -Path $tempSqlitePath -Force
}