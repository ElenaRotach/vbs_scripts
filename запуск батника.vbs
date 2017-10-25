dim fso, path, name
Set fso = CreateObject("Scripting.FileSystemObject")
path = "\\192.168.0.0\schedule\"
name = "F1EasyUpdateJob.bat"
If fso.FileExists(name) = False Then
MsgBox("Не найден файл F1EasyUpdateJob.bat")
Else
Dim WshShell
Set WshShell = CreateObject("Wscript.Shell")
WshShell.Run path & name
End If
