<#
    Script to extract outlook logs,
    and save logs to flat file.
#>

try{
    Get-WinEvent -FilterHashtable @{LogName=”application”;ProviderName=”Outlook”;ID=45} |
    Select -ExpandProperty message |
    Out-File outlookLogs.txt
}
catch{
    if($error.count -gt 0){
        Write-Output "Encountered error: " $error[0].Exception.Message;
    }
}finally{
    if($error.count -eq 0){
        Write-Output "All completed.";
    }
    else{
        Write-Output "Copy above error message send it over to support team.";
    }
}

$key = Read-Host "Hit enter to exit"
if ($key -eq [string]::empty){
    [System.Diagnostics.Process]::GetCurrentProcess().Kill()
}else{
    Write-Host "You can close the window."
}