$StopWatch = new-object system.diagnostics.stopwatch
$StopWatch.Start()

#MANUAL
$LOG_PATH = "\\FS_SERVER\reports$\Windows_Host_Hardware_Inventory\LOG" # Server log file path.
$Directory_path = "\\FS_SERVER\reports$\Windows_Host_Hardware_Inventory\RESULT\DATA" # Folder with clients results
$RESULT_PATH = "\\FS_SERVER\reports$\Windows_Host_Hardware_Inventory\RESULT" # Path where stored RESULT.XML and REPORT.xlsx

#AUTO
$XML_File = $RESULT_PATH + "\" + "REPORT.xml" # Full path for server XML REPORT file for Excel.

$LOG_FILE_FULL_PATH = $LOG_PATH + "\" + "SERVER_LOG.xml"
$RESULT_FILE_FULL_PATH = $RESULT_PATH + "\" + "REPORT.xml"

$XML_Log_Record = [xml] '<?xml version="1.0" encoding="UTF-8"?><LOG><DATE></DATE><TIME></TIME><FILENAME></FILENAME><STATE></STATE><ERROR_MESSAGE></ERROR_MESSAGE><SCRIPT_VERSION></SCRIPT_VERSION></LOG>'
$XML_Document_Dest = [xml] '<?xml version="1.0" encoding="UTF-8"?><HOSTS></HOSTS>'
$XML_Log_File = [xml] '<?xml version="1.0" encoding="UTF-8"?><RESULTS></RESULTS>'

function Check_CRC {
    Param(
        [Parameter(Position=0,Mandatory=$true)][ValidateNotNullOrEmpty()][String[]]$File_Full_Path,
        [Parameter(Position=1,Mandatory=$false)][ValidateNotNullOrEmpty()][ValidateSet($false,$true)]$DEBUG_MODE
    )
    
    if ((Test-Path -Path $File_Full_Path) -eq $false) {
        if ($DEBUG_MODE -eq $true) { 
           
        }

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
        if ($DEBUG_MODE -eq $true){
           
        }
        return $true
    }
}

try {   
    $XMLFiles = Get-ChildItem $Directory_path -Filter *.xml -ErrorAction Stop
    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $Directory_path
    $XML_Log_Record.LOG.STATE = "SUCCESS"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "List of xml files successfully retrieved."
    
    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);

}
catch {
    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $Directory_path
    $XML_Log_Record.LOG.STATE = "EXCEPTION"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Сan't get *.xml files."
    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);
    $XML_Log_File.Save($LOG_FILE_FULL_PATH)
    exit
}

foreach($XMLfile in $XMLFiles) {

    $File_full_path = $XMLfile.FullName


    if (Check_CRC -File_Full_Path $File_full_path) {
        
        $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
        $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
        $XML_Log_Record.LOG.FILENAME = $File_full_path
        $XML_Log_Record.LOG.STATE = "ERROR"
        $XML_Log_Record.LOG.ERROR_MESSAGE = "File is blocked"

        $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
        $XML_Log_File.DocumentElement.AppendChild($insert);
        $XML_Log_File.Save($LOG_FILE_FULL_PATH)
        continue
    }

    try {
        [xml]$XmlDocument_Source = Get-Content -Encoding UTF8 $File_full_path -ReadCount -1 -ErrorAction Stop


        $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
        $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
        $XML_Log_Record.LOG.FILENAME = $XMLfile.FullName
        $XML_Log_Record.LOG.STATE = "SUCCESS"
        $XML_Log_Record.LOG.ERROR_MESSAGE = "File processed successfully."

        $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
        $XML_Log_File.DocumentElement.AppendChild($insert);

        $XML_Log_File.Save($LOG_FILE_FULL_PATH)
    }
    catch {
        $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
        $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
        $XML_Log_Record.LOG.FILENAME = $XMLfile.FullName 
        $XML_Log_Record.LOG.STATE = "ERROR"
        $XML_Log_Record.LOG.ERROR_MESSAGE = $_.CategoryInfo.Activity + ":" + " " + $_.CategoryInfo.Reason 

        $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
        $XML_Log_File.DocumentElement.AppendChild($insert);
        
        $XML_Log_File.Save($LOG_FILE_FULL_PATH)
    }

    $insert = $XML_Document_Dest.ImportNode(($XmlDocument_Source.HOST), $true);
    $XML_Document_Dest.DocumentElement.AppendChild($insert);
}

try {
    if (Test-Path -Path $RESULT_FILE_FULL_PATH -ErrorAction Stop) {
        Remove-Item -Path $RESULT_FILE_FULL_PATH -Force -ErrorAction Stop
    }

    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $RESULT_FILE_FULL_PATH
    $XML_Log_Record.LOG.STATE = "SUCCESS"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Old RESULT.XML file deleted successfully."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);

}
catch {
    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $RESULT_FILE_FULL_PATH
    $XML_Log_Record.LOG.STATE = "ERROR"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Can't delete old RESULT.XML file."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);

    $XML_Log_File.Save($LOG_FILE_FULL_PATH)
    exit
}

try {
    if (Test-Path -Path $LOG_FILE_FULL_PATH -ErrorAction Stop) {
        Remove-Item -Path $LOG_FILE_FULL_PATH -Force -ErrorAction Stop
    }

    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $LOG_FILE_FULL_PATH
    $XML_Log_Record.LOG.STATE = "SUCCESS"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Old LOG.XML file deleted successfully."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);


}
catch {
    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $LOG_FILE_FULL_PATH
    $XML_Log_Record.LOG.STATE = "ERROR"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Can't delete old LOG.XML file"

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);
    # Переработать, здесь явная проблема.
    $XML_Log_File.Save($LOG_FILE_FULL_PATH)

    exit
}

try {

    $XML_Document_Dest.Save($RESULT_FILE_FULL_PATH)

    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $RESULT_FILE_FULL_PATH
    $XML_Log_Record.LOG.STATE = "SUCCESS"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "RESULT.xml file saved successfully."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);

}
catch {

    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $RESULT_FILE_FULL_PATH
    $XML_Log_Record.LOG.STATE = "ERROR"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "Can't save new RESULT.xml file"

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true);
    $XML_Log_File.DocumentElement.AppendChild($insert);
    $XML_Log_File.Save($LOG_FILE_FULL_PATH)

    exit
}

$XML_Log_File.Save($LOG_FILE_FULL_PATH)

$StopWatch.Stop()
Write-Host "Time elapsed (Milliseconds): "
$StopWatch.Elapsed.TotalMilliseconds