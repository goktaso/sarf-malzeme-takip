Option Explicit

Dim files, shell, launcher
Set files = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
launcher = files.BuildPath(files.GetParentFolderName(WScript.ScriptFullName), "Sarf-Malzeme-Takip.ps1")
WScript.Quit shell.Run("powershell.exe -NoProfile -WindowStyle Hidden -File " & Chr(34) & launcher & Chr(34), 0, False)
