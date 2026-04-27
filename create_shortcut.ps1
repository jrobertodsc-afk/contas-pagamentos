$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("C:\Users\Roberto\Desktop\Com System.lnk")
$Shortcut.TargetPath = "pythonw.exe"
$Shortcut.Arguments = "`"c:\Users\Roberto\Documents\ROBO\main.py`""
$Shortcut.WorkingDirectory = "c:\Users\Roberto\Documents\ROBO"
$Shortcut.IconLocation = "c:\Users\Roberto\Documents\ROBO\aliados_hub.ico"
$Shortcut.Description = "Com System - Agência de Atendimento Digital"
$Shortcut.Save()
