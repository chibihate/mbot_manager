# Disable Xbox game
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\GameDVR" /v AppCaptureEnabled /t REG_DWORD /d 0 /f
reg add "HKCU\System\GameConfigStore" /v GameDVR_Enabled /t REG_DWORD /d 0 /f

# Install font
$FontPath = "C:\Users\WDAGUtilityAccount\Documents\VnTahoma.ttf"
$Shell = New-Object -ComObject Shell.Application
$FontsFolder = $Shell.Namespace(0x14) # 0x14 is the ID for C:\Windows\Fonts
$FontsFolder.CopyHere($FontPath)

# Setup python enviroment
Copy-Item `
"C:\Users\WDAGUtilityAccount\Documents\mbot_manager" `
"C:\Users\WDAGUtilityAccount\Desktop\" `
-Recurse -Force

$currentPath = [Environment]::GetEnvironmentVariable("PATH", "User")
$newPath = "C:\Users\WDAGUtilityAccount\Desktop\mbot_manager\python;C:\Users\WDAGUtilityAccount\Desktop\mbot_manager\python\Scripts;" + $currentPath
[Environment]::SetEnvironmentVariable("PATH", $newPath, "User")
$env:PATH = $newPath + ";" + $env:PATH

# Install required software
Start-Process "C:\Users\WDAGUtilityAccount\Documents\npp.8.9.3.Installer.x64.exe" -ArgumentList "/S"
Start-Process "C:\Users\WDAGUtilityAccount\Documents\Opera_131.0.5877.5_Setup_x64.exe" -ArgumentList "/silent=1", "/launchbrowser=0", "/install=1"
Start-Process "C:\Users\WDAGUtilityAccount\Documents\vlc-3.0.23-win32.exe" -ArgumentList "/S"
Start-Process "C:\Users\WDAGUtilityAccount\Documents\winrar-x64-61b1.exe" -ArgumentList "/S" -Wait

Start-Sleep 2

Start-Process `
  -FilePath "C:\Program Files\WinRAR\WinRAR.exe" `
  -ArgumentList 'x "C:\Users\WDAGUtilityAccount\Documents\AofSro.rar" "C:\Users\WDAGUtilityAccount\Desktop\" -y' `
  -Wait

Start-Sleep 2

Start-Process cmd.exe -ArgumentList '/c "C:\Users\WDAGUtilityAccount\Desktop\mbot_manager\launch_autologin.bat"'