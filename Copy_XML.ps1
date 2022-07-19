$StopWatch = new-object system.diagnostics.stopwatch

$StopWatch.Start()

$XMLS_Path = "\\FS_SERVER\reports$\Windows_Host_Hardware_Inventory\DATA"

$Dest_Path = "\\FS_SERVER\reports$\Windows_Host_Hardware_Inventory\RESULT\DATA"
$XML_Log_file_path = "\\FS_SERVER\reports$\Windows_Host_Hardware_Inventory\LOG\FILE_COPY_LOG.xml"

$XML_Log_File = [xml] '<?xml version="1.0" encoding="UTF-8"?><RESULTS></RESULTS>'
$XML_Log_Record = [xml] '<?xml version="1.0" encoding="UTF-8"?><LOG><DATE></DATE><TIME></TIME><FILENAME></FILENAME><STATE></STATE><ERROR_MESSAGE></ERROR_MESSAGE></LOG>'


function exit_handler {
    $XML_Log_File.Save($XML_Log_file_path)

    $StopWatch.Stop()
    Write-Host "Time elapsed (Milliseconds): "
    $StopWatch.Elapsed.TotalMilliseconds

    exit
}

function Is-Lock {
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
            if ($DEBUG_MODE -eq $true){
               
            }
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
    $XMLS = Get-ChildItem $XMLS_Path -Filter *.xml -ErrorAction Stop
    
    $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
    $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
    $XML_Log_Record.LOG.FILENAME = $XMLS_Path
    $XML_Log_Record.LOG.STATE = "SUCCESS"
    $XML_Log_Record.LOG.ERROR_MESSAGE = "List of xml files successfully retrieved."

    $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true)
    $XML_Log_File.DocumentElement.AppendChild($insert)

    if ($XMLS.Count -le 0) {
        
        $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
        $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
        $XML_Log_Record.LOG.FILENAME = $Directory_path
        $XML_Log_Record.LOG.STATE = "EXCEPTION"
        $XML_Log_Record.LOG.ERROR_MESSAGE = "There are no XML files in the specified folder."
        
        $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true)
        $XML_Log_File.DocumentElement.AppendChild($insert)


        throw
    }
    
}
catch {
    exit_handler
}

foreach ($file in $XMLS) {
    $path = $file.FullName

    $path

    if (Is-Lock -File_Full_Path $path) {
        $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
        $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
        $XML_Log_Record.LOG.FILENAME = $path
        $XML_Log_Record.LOG.STATE = "EXCEPTION"
        $XML_Log_Record.LOG.ERROR_MESSAGE = "File skipped due to file lock."
        $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true)
        $XML_Log_File.DocumentElement.AppendChild($insert)

        continue
    }

    try {
        Copy-Item -Path $path -Destination $Dest_Path -Force -ErrorAction Stop
        
        $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
        $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
        $XML_Log_Record.LOG.FILENAME = $path
        $XML_Log_Record.LOG.STATE = "SUCCESS"
        $XML_Log_Record.LOG.ERROR_MESSAGE = "File copied successfully"
        $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true)
        $XML_Log_File.DocumentElement.AppendChild($insert)
    }
    catch {
        $XML_Log_Record.LOG.DATE = (Get-Date -Format "dd.MM.yyyy").ToString()
        $XML_Log_Record.LOG.TIME = (Get-Date -Format "HH:mm:ss").ToString()
        $XML_Log_Record.LOG.FILENAME = $path
        $XML_Log_Record.LOG.STATE = "EXCEPTION"
        $XML_Log_Record.LOG.ERROR_MESSAGE = "File cannot be copied. Reason: " + $_.CategoryInfo.Activity + ":" + " " + $_.CategoryInfo.Reason
        $insert = $XML_Log_File.ImportNode(($XML_Log_Record.LOG), $true)
        $XML_Log_File.DocumentElement.AppendChild($insert)

        continue
    }
}

exit_handler
