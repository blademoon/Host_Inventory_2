#MANUAL
$SCRIPT_VERSION = "v5.7.4" # Добавлен DISPLAY_NAME
$NETWORK_PATH = "\\FS_SERVER\reports$\Windows_Host_Hardware_Inventory\DATA"
$SCRIPT_PREFIX = "INV_5_" 

#AUTO
$SCRIPT_REPORT = [xml] '<?xml version="1.0" encoding="UTF-8"?><HOST><DATE></DATE><TIME></TIME><HOSTNAME></HOSTNAME><HOST_MANUFACTURER></HOST_MANUFACTURER><HOST_MODEL></HOST_MODEL><HOST_PRODUCT_NUMBER></HOST_PRODUCT_NUMBER><HOST_SYSTEM_TYPE></HOST_SYSTEM_TYPE><SERIAL></SERIAL><DOMAIN></DOMAIN><IP></IP><USERS></USERS><DISPLAY_NAME></DISPLAY_NAME><LOCAL_ADMINS></LOCAL_ADMINS><CITY></CITY><AD_OU></AD_OU><DomainAdminUsers></DomainAdminUsers><UPTIME></UPTIME><OS_NAME></OS_NAME><OS_VERSION></OS_VERSION><OS_ARCHITECTURE></OS_ARCHITECTURE><OS_BUILD></OS_BUILD><OS_INSTALLATION_DATE></OS_INSTALLATION_DATE><CPU_NAME></CPU_NAME><CPU_PHYSICAL_NUMBER></CPU_PHYSICAL_NUMBER><CPU_CORES_TOTAL></CPU_CORES_TOTAL><CPU_TEMPERATURE></CPU_TEMPERATURE><RAM_TOTAL></RAM_TOTAL><RAM_FREE></RAM_FREE><VIDEOCARD_NAME></VIDEOCARD_NAME><DISK_INFO></DISK_INFO><HDD_SMART_PREDICT></HDD_SMART_PREDICT><POWERSHELL_VERSION></POWERSHELL_VERSION><SCRIPT_VERSION></SCRIPT_VERSION></HOST>'
$SCRIPT_REPORT_FILE_NAME = $SCRIPT_PREFIX + $env:COMPUTERNAME + ".xml"
$LOCAL_PATH = $NETWORK_PATH + "\" + $SCRIPT_REPORT_FILE_NAME

function Is-File-Locked {
    Param(
        [Parameter(Position=0,Mandatory=$true)][ValidateNotNullOrEmpty()][String[]]$File_Full_Path,
        [Parameter(Position=1,Mandatory=$false)][ValidateNotNullOrEmpty()][ValidateSet($false,$true)]$DEBUG_MODE
    )
    
    if ((Test-Path -Path $File_Full_Path) -eq $false) {
        return $false
    }

    $oFile = New-Object System.IO.FileInfo $File_Full_Path
    
    try {
        $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        if ($oStream) {
            $oStream.Close()
            return $false    
        } 
    } 
    catch {
        return $true
    }
}

function Get-PS-Version {
    $TEMP = $PSVersionTable.PSVersion
    if ($TEMP -eq $null) {return "EXCEPTION: Get-PS-Version return empty value"}
    [string]$result = ($PSVersionTable.PSVersion.Major).ToString() + "." + ($PSVersionTable.PSVersion.Minor).ToString() + "." + ($PSVersionTable.PSVersion.Build).ToString() + "." + ($PSVersionTable.PSVersion.Revision).ToString()
    return $result
}

function Get_LOCAL_ADMINS {
    $LOCAL_ADMINS = $null
    try {
        $LOCAL_ADMINS = (Get-LocalGroupMember -Group "Администраторы" -ErrorAction Stop).Name
    }
    catch {
        try {
            $LOCAL_ADMINS = (Get-LocalGroupMember -Group "Administrators" -ErrorAction Stop).Name
        }
        catch {
            $LOCAL_ADMINS = $null
        }
    }
    if ($LOCAL_ADMINS -eq $null) {
        return "EXCEPTION: Get_LOCAL_ADMINS"
    }
    
    $result = ""
    foreach ($item in $LOCAL_ADMINS) {
        $result = $result + $item + "`r`n"
    }
    $result = $result.Substring(0,$result.Length-2)
    return $result
}

function CreateDiskStr {
    Param(
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [Object[]] $DiskInfo
    )
    ForEach ($disk in $DiskInfo) {
        $str1 = $disk.FreeSpace
        $str1 = $str1/1024/1024/1024
        $str1 = [math]::Round($str1)
        $str1 = $str1.ToString()
        $str2 = $disk.Size
        $str2 = $str2/1024/1024/1024
        $str2 = [math]::Round($str2)
        $str2 = $str2.ToString()
        $Procent = [math]::Round(($str1/$str2)*100)
        
        $return = $return + $Disk.DeviceID + " Общий объём: " + $str2  + " Гбайт. Из них свободно “ + $str1 + " Гбайт. Использовано “ + (100-$Procent) + "%." + "`n"
    }
    
    return $return    
}

function Get-AD-CanonicalName {
    Param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String[]]$Hostname
    )

    $filter = "(&(objectClass=computer)(cn=$Hostname))"
    $DistinguishedName = (([adsisearcher]$filter).FindOne().Properties.distinguishedname)

    foreach ($dn in $DistinguishedName) {      
        $d = $dn.Split(',') 
        $arr = (@(($d | Where-Object { $_ -notmatch 'DC=' }) | ForEach-Object { $_.Substring(3) })) 
        [array]::Reverse($arr) 
 
        "{0}/{1}" -f  (($d | Where-Object { $_ -match 'dc=' } | ForEach-Object { $_.Replace('DC=','') }) -join '.'), ($arr -join '/') 
    } 
}

Function Get-Local-IPs {
    param (
        [parameter(Position=0,Mandatory=$true)][ValidateNotNullOrEmpty()][System.Object[]][ref]$Win32_NetworkAdapterConfiguration
    )

    $Result = ""

    foreach ($NIC in $Win32_NetworkAdapterConfiguration) {
        
        $Hardware_NIC_Configuration = $NIC.GetRelated('Win32_NetworkAdapter') | Select *
        #$Result1 = $Hardware_NIC_Configuration.NetConnectionID + " IP: "

        foreach ($IP in $NIC.IPAddress) {
             
             if ($IP -like "*::*") {continue}

             if (($IP -like "10.242*") -or ($IP -like "10.243*") -or ($IP -like "10.244*") -or ($IP -like "10.227.124*") -or ($IP -like "192.168.*")) {continue}  
             

             #$Result = $Hardware_NIC_Configuration.NetConnectionID + ", IP: "
             $Result += $IP 

             #switch ($Hardware_NIC_Configuration.Speed) {
             #   20000000000 { $Result = $Result + " SPEED: 20 Gbps"; break}
             #   10000000000 { $Result = $Result + " SPEED: 10 Gbps"; break}
             #   1000000000 { $Result = $Result + " SPEED: 1 Gbps"; break}
             #   100000000 { $Result = $Result + " SPEED: 100 Mbps"; break}
             #   10000000 { $Result = $Result + " SPEED: 10 Mbps"; break}
             #
             #   default {$Result = " " + $Result + " Unknown speed " + $Hardware_NIC_Configuration.Speed}
             #}

             $Result += "`n"
        }
    }

    return $Result  
}


function GetTemperature {

    try {
        $tempWMI = Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" -ErrorAction Stop
        if( $tempWMI ) {
            $nprepTemp = 0
            $resTemp = 0
            
            if ( $tempWMI.count ) { 
                $tempWMI | % { [int]$nprepTemp += $_.CurrentTemperature }
                $nprepTemp = $nprepTemp/($tempWMI.count)
            }
            else { [int]$nprepTemp += $tempWMI.CurrentTemperature }
        
            [int]$resTemp += ( $( $nprepTemp )/10 -273.15)*1.8 +32
            if ([string]::IsNullOrEmpty(($resTemp).ToString())) {throw}
            return (($resTemp).ToString())
        }
    
    }
    catch {
        return "NOT SUPPORTED"
    }  
}

function Can-Write-To-Path {
    Param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String[]] $TestPath
    )

    Try { 
        [io.file]::OpenWrite($outfile).close() 
        [io.file]::Delete($outfile)
        return $true
    }
    Catch { 
        #Unable to write to output file $outputfile
        return $false
    }
}

function Get-City {
    Param(
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string] $AD_OU_Path
    )

    $return = ""

    switch -Wildcard ($AD_OU_Path) {

        "*Disabled Computers*" {$return ="Отключенные ПК"; break}

        "*/AG Firm/External Computers/*" {$return ="Внешние ПК"; break}

        "*Computers for *Nefteyugansk*" {$return ="г. Нефтеюганск, SD"; break}

        "*OK/Nefteyugansk/*" {$return ="г. Нефтеюганск, Отдел кадров"; break}

        default {$return ="Другое"}
    }

    return $return    
}

function Get-Logged-User {
    Param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][System.Object[]]$Win32_ComputerSystem,
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][System.Object[]]$Win32_OperatingSystem
    )

    $return = "NOT SUPPORTED"

    try {
        $OSType = $Win32_OperatingSystem.ProductType

        if ($OSType -isnot [UInt32]) {
            $return = "EXCEPTION: Get-Logged-User() type"
            throw
        }

        if ($OSType -eq $null) {
            $return = "EXCEPTION: Get-Logged-User() null"
            throw
        }
        
        if ($OSType -ne 1) {
            $return = "Unsupported Product Type"
            throw
        }

        $User = ($Win32_ComputerSystem.username).ToString()
        if ([string]::IsNullOrEmpty(($User).ToString())) {throw}
        $return = $User
    }
    catch {
         return $return
    }

    return $return
}

function Get-PC-Type {
Param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][System.Object[]]$Win32_ComputerSystem
)

    $return = "NOT_DETECTED"

    $SystemType = ($Win32_ComputerSystem).PCSystemType

    switch($SystemType) {
    1 {$return = "Desktop"}
    2 {$return = "Mobile"}
    3 {$return = "Workstation"}
    4 {$return = "Enterprise Server"}
    5 {$return = "SOHO Server"}
    6 {$return = "Appliance PC"}
    7 {$return =  "Performance Server"}
    default {$return = "Other"}
    
    }

    return $return
}

$SCRIPT_REPORT.HOST.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
$SCRIPT_REPORT.HOST.TIME = (Get-Date -Format "HH:mm:ss").ToString()
$SCRIPT_REPORT.HOST.HOSTNAME = ($env:COMPUTERNAME).ToString()
$SCRIPT_REPORT.HOST.POWERSHELL_VERSION = ($PSVersionTable.PSVersion.Major).ToString() + "." + ($PSVersionTable.PSVersion.Minor).ToString() + "." + ($PSVersionTable.PSVersion.Build).ToString() + "." + ($PSVersionTable.PSVersion.Revision).ToString()
$SCRIPT_REPORT.HOST.LOCAL_ADMINS = Get_LOCAL_ADMINS
$SCRIPT_REPORT.HOST.SCRIPT_VERSION = $SCRIPT_VERSION

# GetTemperature()
try {
    $SCRIPT_REPORT.HOST.CPU_TEMPERATURE = GetTemperature
}
catch {
    $SCRIPT_REPORT.HOST.CPU_TEMPERATURE = "EXCEPTION: GetTemperature()"
}

# Get-AD-CanonicalName() Get-City()
try {
    $TEMP_CanonicallName = Get-AD-CanonicalName($env:COMPUTERNAME)
    if ([string]::IsNullOrEmpty($TEMP_CanonicallName)) {throw}
    $SCRIPT_REPORT.HOST.AD_OU = $TEMP_CanonicallName
    $SCRIPT_REPORT.HOST.CITY = Get-City -AD_OU_Path $TEMP_CanonicallName
}
catch {
    $SCRIPT_REPORT.HOST.AD_OU ="EXCEPTION: Get-AD-CanonicalName()"
    $SCRIPT_REPORT.HOST.CITY = "EXCEPTION: Get-AD-CanonicalName()"
}

# Win32_Processor
try {
    $CPU = Get-WmiObject Win32_Processor -namespace "root\CIMV2" -ErrorAction Stop
    if ($CPU -eq $null) {throw}
    #---------------------------------------------------------------------------------
    $CPU_Names = $CPU.Name
    $CPU_Names_result = ""
    if ($CPU_Names -is [string]) {
        $CPU_Names_result = ($CPU_Names -replace '(^\s+|\s+$)','' -replace '\s+',' ')
    }
    
    if ($CPU_Names -is [Object]) {
        $CPU_Names_result = ($CPU_Names -replace '(^\s+|\s+$)','' -replace '\s+',' ') -join "`r`n"
    }
    $SCRIPT_REPORT.HOST.CPU_NAME = $CPU_Names_result
    #---------------------------------------------------------------------------------
    $CPUs_Core_Count = $CPU.NumberOfCores
    $result_core = 0
    if ($CPUs_Core_Count -is [UInt32]) {
        $result_core = $CPUs_Core_Count
    }
    if ($CPUs_Core_Count -is [object]) {
        foreach ($item in $CPUs_Core_Count) {
            $result_core = $result_core + $item
        }  
    }
    $SCRIPT_REPORT.HOST.CPU_CORES_TOTAL = (($result_core).ToString())
    #---------------------------------------------------------------------------------
    
}
catch {
    $CPU = "EXCEPTION: Win32_Processor"
    $SCRIPT_REPORT.HOST.CPU_NAME = $CPU
    $SCRIPT_REPORT.HOST.CPU_CORES_TOTAL = $CPU
}

# Win32_ComputerSystem
try {
    $CS = Get-WmiObject -class Win32_ComputerSystem -namespace "root\CIMV2" -ErrorAction Stop
    if ($CS -eq $null) {throw}
    $SCRIPT_REPORT.HOST.CPU_PHYSICAL_NUMBER = (($CS.NumberOfProcessors).ToString())
    $SCRIPT_REPORT.HOST.HOST_MANUFACTURER = ($CS.Manufacturer).ToString()
    $SCRIPT_REPORT.HOST.HOST_MODEL = ($CS.Model).ToString()
    $SCRIPT_REPORT.HOST.HOST_SYSTEM_TYPE = Get-PC-Type -Win32_ComputerSystem $CS
    $SCRIPT_REPORT.HOST.DOMAIN = ($CS.Domain).ToString()
}
catch {
    $CS = "EXCEPTION: Win32_ComputerSystem"
    $SCRIPT_REPORT.HOST.CPU_PHYSICAL_NUMBER = $CS
    $SCRIPT_REPORT.HOST.HOST_MANUFACTURER = $CS
    $SCRIPT_REPORT.HOST.HOST_MODEL = $CS
    $SCRIPT_REPORT.HOST.HOST_SYSTEM_TYPE =  $CS
    $SCRIPT_REPORT.HOST.DOMAIN = $CS
}

# SystemSKUNumber
try {
    $SKUNumber = ($CS.SystemSKUNumber).ToString()
    if ([string]::IsNullOrEmpty($SKUNumber)) {throw}
    $SCRIPT_REPORT.HOST.HOST_PRODUCT_NUMBER = $SKUNumber
}
catch {
    $SKUNumber = "NOT SUPPORTED"
    $SCRIPT_REPORT.HOST.HOST_PRODUCT_NUMBER = $SKUNumber
}

# Win32_physicalmemory
try {
    $PM = get-wmiobject -class Win32_physicalmemory -namespace "root\CIMV2" -ErrorAction Stop
    if ($PM -eq $null) {throw}
    $SCRIPT_REPORT.HOST.RAM_TOTAL = "$(([math]::Round((($PM).Capacity | Measure-Object -Sum).Sum/1GB))) GB"
}
catch {
    $PM = "EXCEPTION: Win32_physicalmemory"
    $SCRIPT_REPORT.HOST.RAM_TOTAL = $PM
}

# Win32_OperatingSystem
try {
    $OS = Get-WmiObject -class Win32_OperatingSystem -namespace "root\CIMV2" -ErrorAction Stop
    if ($OS -eq $null) {throw}
    
    $Uptime = (Get-Date) - ([Management.ManagementDateTimeConverter]::ToDateTime($OS.LastBootUpTime))
    $Uptime_str = ($Uptime.Days).ToString() + " " + ($Uptime.Hours).ToString() + ":" + ($Uptime.Minutes).ToString() + ":" + ($Uptime.Seconds).ToString()
    $SCRIPT_REPORT.HOST.UPTIME = $Uptime_str

    $SCRIPT_REPORT.HOST.RAM_FREE = ([math]::Round((($OS).FreePhysicalMemory)/1024/1024)).ToString() + " GB"
    $SCRIPT_REPORT.HOST.OS_NAME = ($OS.Caption).ToString()
    $SCRIPT_REPORT.HOST.OS_VERSION = ($OS.Version).ToString()
    $SCRIPT_REPORT.HOST.OS_ARCHITECTURE = (($OS).OSArchitecture).ToString()
    $SCRIPT_REPORT.HOST.OS_BUILD = ($OS.BuildNumber).ToString()
    $SCRIPT_REPORT.HOST.OS_INSTALLATION_DATE = (([WMI]'').ConvertToDateTime(($OS).InstallDate)).ToString()
}
catch {
    $OS = "EXCEPTION: Win32_OperatingSystem"
    $SCRIPT_REPORT.HOST.UPTIME = $OS
    $SCRIPT_REPORT.HOST.RAM_FREE = $OS
    $SCRIPT_REPORT.HOST.OS_NAME = $OS
    $SCRIPT_REPORT.HOST.OS_VERSION = $OS
    $SCRIPT_REPORT.HOST.OS_ARCHITECTURE = $OS
    $SCRIPT_REPORT.HOST.OS_BUILD = $OS
    $SCRIPT_REPORT.HOST.OS_INSTALLATION_DATE = $OS
}

# Get-Logged-User()
try {
    $SCRIPT_REPORT.HOST.USERS = Get-Logged-User -Win32_ComputerSystem $CS -Win32_OperatingSystem $OS
}
catch {
    $SCRIPT_REPORT.HOST.USERS = "NOT SUPPORTED"
}

# Get-Display-Name (On Host)
try {
    $temp = ($SCRIPT_REPORT.HOST.USERS).Split("\")
    $dom = $temp[0]
    $usr = $temp[1] 

    #$SCRIPT_REPORT.HOST.DISPLAY_NAME = ([adsi]"WinNT://$dom/$usr,user").fullname
    $SCRIPT_REPORT.HOST.DISPLAY_NAME = ([adsi]"WinNT://$dom/$usr,user").fullname.Value
    }
catch {
    $SCRIPT_REPORT.HOST.DISPLAY_NAME = "Unsupported Product Type"
}

# Win32_BIOS
try {
    $B = Get-WmiObject -class Win32_BIOS -namespace "root\CIMV2" -ErrorAction Stop
    if ($B -eq $null) {throw}
    $SCRIPT_REPORT.HOST.SERIAL = (($B).SerialNumber -replace '(^\s+|\s+$)','').ToString()
}
catch {
    $B = "NOT SUPPORTED"
    $SCRIPT_REPORT.HOST.SERIAL = $B
}

# Win32_VideoController
try {
    $VC = Get-WmiObject Win32_VideoController -namespace "root\CIMV2" -ErrorAction Stop
    if ($VC -eq $null) {throw}
    $VideoCard = ($VC).Name

    if ($VideoCard -is [string]) {
    $SCRIPT_REPORT.HOST.VIDEOCARD_NAME = $VideoCard
    }
    elseif ($VideoCard -is [Object]) {
    $SCRIPT_REPORT.HOST.VIDEOCARD_NAME = $VideoCard -join "`n"
    }
    else {trow}
}
catch {
    $VC = "EXCEPTION: Win32_VideoController"
    $SCRIPT_REPORT.HOST.VIDEOCARD_NAME = $VC
}

# Win32_LogicalDisk
try {
    $LD = Get-WmiObject Win32_LogicalDisk -namespace "root\CIMV2" -ErrorAction Stop | where {$_.DriveType -eq 3}
    if ($LD -eq $null) {throw}
    $SCRIPT_REPORT.HOST.DISK_INFO = CreateDiskStr($LD)
}
catch {
    $LD = "EXCEPTION: Win32_LogicalDisk"
    $SCRIPT_REPORT.HOST.DISK_INFO = $LD
}

# SMART_PREDICT
try {
    $SMART_Predict = (Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction Stop).PredictFailure
    if ($SMART_Predict -eq $null) {throw}
    $SCRIPT_REPORT.HOST.HDD_SMART_PREDICT = $SMART_Predict.ToString()
}
catch {
    $SMART_Predict = $_.Exception.Message 
    $SCRIPT_REPORT.HOST.HDD_SMART_PREDICT = $SMART_Predict
}

# Win32_NetworkAdapterConfiguration
try {
    $NICs = Get-WmiObject Win32_NetworkAdapterConfiguration -namespace "root\CIMV2" -ErrorAction Stop | Where-Object {($_.IPEnabled -eq 'TRUE')}
    if ($NICs -eq $null) {throw}
    $SCRIPT_REPORT.HOST.IP = Get-Local-IPs -Win32_NetworkAdapterConfiguration ([ref]$NICs)
}
catch {
    $LD = "EXCEPTION: Win32_NetworkAdapterConfiguration"
    $SCRIPT_REPORT.HOST.IP = $LD
}

# Check lock
if (Is-File-Locked -File_Full_Path $LOCAL_PATH) {
    Start-Sleep -Seconds 60
    if (Is-File-Locked -File_Full_Path $LOCAL_PATH) {
        exit
    }
}

try {
    if (Test-Path -Path $LOCAL_PATH -ErrorAction Stop) {
        Remove-Item -Path $LOCAL_PATH -Force -ErrorAction Stop
    }
}
catch {
    exit
}

try {
    $SCRIPT_REPORT.Save($LOCAL_PATH)
}
catch {
    exit
}