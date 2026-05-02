reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\GameDVR" /v AppCaptureEnabled /t REG_DWORD /d 0 /f
reg add "HKCU\System\GameConfigStore" /v GameDVR_Enabled /t REG_DWORD /d 0 /f

Start-Process "C:\Users\WDAGUtilityAccount\Documents\winrar-x64-61b1.exe" -ArgumentList "/S" -Wait

Start-Sleep 3

& "C:\Program Files\WinRAR\WinRAR.exe" x `
"C:\Users\WDAGUtilityAccount\Documents\AofSro.rar" `
"C:\Users\WDAGUtilityAccount\Desktop\" -y

Start-Sleep 3

Start-Process "C:\Users\WDAGUtilityAccount\Documents\npp.8.9.3.Installer.x64.exe" -ArgumentList "/S" -Wait

Start-Sleep 3

$FontPath = "C:\Users\WDAGUtilityAccount\Documents\VnTahoma.ttf" # Update with your actual path
$Shell = New-Object -ComObject Shell.Application
$FontsFolder = $Shell.Namespace(0x14) # 0x14 is the ID for C:\Windows\Fonts
$FontsFolder.CopyHere($FontPath)

Start-Sleep 3

$currentPath = [Environment]::GetEnvironmentVariable("PATH", "User")
$newPath = "C:\Users\WDAGUtilityAccount\Documents\mbot_manager\python;C:\Users\WDAGUtilityAccount\Documents\mbot_manager\python\Scripts;" + $currentPath
[Environment]::SetEnvironmentVariable("PATH", $newPath, "User")

Start-Sleep 2

$env:PATH = $newPath + ";" + $env:PATH

Start-Sleep 20


Start-Process "python.exe" -ArgumentList "C:\Users\WDAGUtilityAccount\Documents\mbot_manager\mbot_manager.py --autologin" -WorkingDirectory "C:\Users\WDAGUtilityAccount\Documents\mbot_manager" -Verb RunAs

