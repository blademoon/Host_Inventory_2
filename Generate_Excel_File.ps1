
# MANUAL
$RESULT_PATH = "\\FS_SERVER\reports$\Windows_Host_Hardware_Inventory\RESULT"
$LOG_PATH = "\\FS_SERVER\reports$\Windows_Host_Hardware_Inventory\LOG" # Server log file path.
$Excel_File_Template = "\\FS_SERVER\reports$\Windows_Host_Hardware_Inventory\SCRIPT\Server\template.xlsx" # Excel file template

#AUTO
$XML_File = $RESULT_PATH + "\" + "REPORT.xml" # Full path for server XML REPORT file for Excel.
$LOG_FILE_FULL_PATH = $LOG_PATH + "\" + "SERVER_LOG.xml"
$RESULT_EXCEL_FILE = $RESULT_PATH + "\" + "RESULT.xlsx"

$XML_Log_Record = [xml] '<?xml version="1.0" encoding="UTF-8"?><LOG><DATE></DATE><TIME></TIME><FILENAME></FILENAME><STATE></STATE><ERROR_MESSAGE></ERROR_MESSAGE><SCRIPT_VERSION></SCRIPT_VERSION></LOG>'
$XML_Document_Dest = [xml] '<?xml version="1.0" encoding="UTF-8"?><HOSTS></HOSTS>'
$XML_Log_File = [xml] '<?xml version="1.0" encoding="UTF-8"?><RESULTS></RESULTS>'


$Excel = New-Object -ComObject Excel.Application
$Excel.DisplayAlerts = $False
$Excel.ScreenUpdating = $False
$Excel.Visible = $False
#$Excel.UpdateLinks = $False

try {
    $WorkBook = $Excel.workbooks.Open($Excel_File_Template)
    $WorkSheetName = "Report"
    $WorkSheet = $WorkBook.Worksheets.Item($WorkSheetName)
    $Cells=$WorkSheet.Cells

    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = ""
    $XML_Log_Record.LOG.STATE = "SUCCESS"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Excel started successfully."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);

}
catch {
    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $Excel_File_Template
    $XML_Log_Record.LOG.STATE = "ERROR"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Can't open excel template file."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);
    $XML_Log_File.Save($LOG_FILE_FULL_PATH)
    
    $Excel.Quit()
    Stop-Process -Name EXCEL
    exit
}

try {
    [xml]$XML_Report = Get-Content -Encoding UTF8 $XML_File -ReadCount -1 -ErrorAction Stop

    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $XML_File
    $XML_Log_Record.LOG.STATE = "SUCCESS"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "REPORT.XML file load successfully."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);
}
catch {

    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $XML_File
    $XML_Log_Record.LOG.STATE = "ERROR"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Can't load REPORT.XML file."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);
    $XML_Log_File.Save($LOG_FILE_FULL_PATH)

    $Excel.Quit()
    Stop-Process -Name EXCEL
    exit
}

$HOSTS = $XML_Report.HOSTS.HOST
$i = 2

Foreach ($HOST1 in $HOSTS) {
    $Cells.item($i,1) = $HOST1.DATE
    $Cells.item($i,2) = $HOST1.TIME
    $Cells.item($i,3) = $HOST1.HOSTNAME
    $Cells.item($i,4) = $HOST1.HOST_MANUFACTURER
    $Cells.item($i,5) = $HOST1.HOST_MODEL
    $Cells.item($i,6) = $HOST1.HOST_PRODUCT_NUMBER
    $Cells.item($i,7) = $HOST1.HOST_SYSTEM_TYPE
    $Cells.item($i,8) = $HOST1.SERIAL
    $Cells.item($i,9) = $HOST1.DOMAIN
    $Cells.item($i,10) = $HOST1.IP
    $Cells.item($i,11) = $HOST1.USERS
    $Cells.item($i,12) = $HOST1.LOCAL_ADMINS
    $Cells.item($i,13) = $HOST1.CITY
    $Cells.item($i,14) = $HOST1.AD_OU
    $Cells.item($i,15) = $HOST1.DomainAdminUsers
    $Cells.item($i,16) = $HOST1.UPTIME
    $Cells.item($i,17) = $HOST1.OS_NAME
    $Cells.item($i,18) = $HOST1.OS_VERSION
    $Cells.item($i,19) = $HOST1.OS_ARCHITECTURE
    $Cells.item($i,20) = $HOST1.OS_BUILD
    $Cells.item($i,21) = $HOST1.OS_INSTALLATION_DATE
    $Cells.item($i,22) = $HOST1.CPU_NAME
    $Cells.item($i,23) = $HOST1.CPU_PHYSICAL_NUMBER
    $Cells.item($i,24) = $HOST1.CPU_CORES_TOTAL
    $Cells.item($i,25) = $HOST1.CPU_TEMPERATURE
    $Cells.item($i,26) = $HOST1.RAM_TOTAL
    $Cells.item($i,27) = $HOST1.RAM_FREE
    $Cells.item($i,28) = $HOST1.VIDEOCARD_NAME
    $Cells.item($i,29) = $HOST1.DISK_INFO
    $Cells.item($i,30) = $HOST1.POWERSHELL_VERSION
    $Cells.item($i,31) = $HOST1.SCRIPT_VERSION

    $i++
}

$usedRange = $WorkSheet.UsedRange                                                                                              
$usedRange.WrapText = $False

$usedRange.EntireColumn.AutoFit() | Out-Null
$usedRange.EntireRow.AutoFit() | Out-Null

$usedRange.VerticalAlignment = -4160
$usedRange.HorizontalAlignment = -4130

$usedRange.EntireColumn.AutoFit() | Out-Null
$usedRange.EntireRow.AutoFit() | Out-Null
$WorkBook.RefreshAll()

try {
    if (Test-Path -Path $RESULT_EXCEL_FILE -ErrorAction Stop) {
        Write-Host "Removing Excel file " + $RESULT_EXCEL_FILE
        Remove-Item -Path $RESULT_EXCEL_FILE -Force -ErrorAction Stop
    }

    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $RESULT_EXCEL_FILE
    $XML_Log_Record.LOG.STATE = "SUCCESS"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Old RESULT.XLSX file deleted successfully."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);

}
catch {
    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $RESULT_EXCEL_FILE
    $XML_Log_Record.LOG.STATE = "ERROR"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Can't delete old RESULT.XLSX file."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);
    $XML_Log_File.Save($LOG_FILE_FULL_PATH)
    
    $Excel.Quit()
    Stop-Process -Name EXCEL
    exit
}

try {
    $XML_Log_File.Save($LOG_FILE_FULL_PATH)

    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $LOG_FILE_FULL_PATH
    $XML_Log_Record.LOG.STATE = "SUCCESS"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "New LOG.XML file saved successfully."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);

}
catch {
    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $LOG_FILE_FULL_PATH
    $XML_Log_Record.LOG.STATE = "ERROR"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Can't save new LOG.XML file"
    exit
}

$WorkBook.SaveAs($RESULT_EXCEL_FILE, 51, [Type]::Missing, [Type]::Missing, $False, $False, 1);
$Excel.Quit()
Stop-Process -Name EXCEL