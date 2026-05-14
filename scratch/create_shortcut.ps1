$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("c:\Users\kus\OneDrive\Escritorio\GeneraEtiquetas.lnk")
$Shortcut.TargetPath = "c:\Users\kus\OneDrive\Escritorio\Proyectos\Etiquetas Claro\Lanzar_Etiquetas.bat"
$Shortcut.WorkingDirectory = "c:\Users\kus\OneDrive\Escritorio\Proyectos\Etiquetas Claro"
$Shortcut.Save()
