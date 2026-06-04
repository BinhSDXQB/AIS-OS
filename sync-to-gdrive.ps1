# Dong bo AIS-OS -> Google Drive (chay tai may NHA truoc khi tat may)
# Loai tru ca 2 script sync de tranh vong lap
robocopy "D:\Troly\Duan\AIS-OS" "G:\My Drive\Claude_Duan\AIS-OS" /E /MIR /NP /XF "sync-to-gdrive.ps1" "sync-from-gdrive.ps1" /LOG+:"D:\Troly\Duan\AIS-OS\archives\robocopy-sync.log"
Write-Output "Da day len Google Drive: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
