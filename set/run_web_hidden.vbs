Option Explicit
Dim sh, fso, setDir, rootDir, projDir, cmd, rc
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
setDir = fso.GetParentFolderName(WScript.ScriptFullName)
rootDir = fso.GetParentFolderName(setDir)
projDir = rootDir & "\project"
sh.CurrentDirectory = projDir

cmd = "pyw.exe -3 """ & projDir & "\run_web_launcher.py"""
rc = sh.Run(cmd, 0, True)
If rc <> 0 Then
  cmd = "pythonw.exe """ & projDir & "\run_web_launcher.py"""
  rc = sh.Run(cmd, 0, True)
End If
WScript.Quit rc
