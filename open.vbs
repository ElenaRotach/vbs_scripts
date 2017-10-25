'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso, file, name
'-----------------------------------------------------------------------------------------------------------------------------------------------
Set fso = CreateObject("Scripting.FileSystemObject")
name = "c:\1\test.html"
If fso.FileExists(name) Then
fso.DeleteFile name, True
End If
    fso.CreateTextFile(name)
Set Log_file = fso.OpenTextFile(name, 2, True)

Log_file.writeLine "<!DOCTYPE>"
Log_file.writeLine "<html>"
Log_file.writeLine "<head>"
Log_file.writeLine "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & "content=" & Chr(34) & "text/html; charset=windows-1251" & Chr(34) & " />"
Log_file.writeLine "<title>...</title>"
Log_file.writeLine "</head>"
Log_file.writeLine "<body>"
Log_file.writeLine "<form id=" & Chr(34) & "fileload" & Chr(34) & " action=" & Chr(34) & "#" & Chr(34) & " method=" & Chr(34) & "post" & Chr(34) & " enctype=" & Chr(34) & "multipart/form-data" & Chr(34) & "><input type=" & Chr(34) & "file" & Chr(34) & " onchange=" & Chr(34) & "TestValue(this)" & Chr(34) & " name=" & Chr(34) & "anyfile" & Chr(34) & " id=" & Chr(34) & "inpField" & Chr(34) & "/></form>"
Log_file.writeLine "<script type=" & Chr(34) & "text/javascript" & Chr(34) & ">"
Log_file.writeLine "function TestValue(a){"
Log_file.writeLine "alert(a.value);"
Log_file.writeLine "}"
Log_file.writeLine "</script>"
Log_file.writeLine "</body>"
Log_file.writeLine "</html>"
Log_file.Close

'Set objShell = CreateObject("shell.application")
'        objShell.ShellExecute "c:\1\test.html", "", "", "runas", 1

  Set objIE = CreateObject("InternetExplorer.Application")  
    objIE.Navigate(name) ' ????? ?????????????????
    Do While objIE.Busy : Wscript.Sleep 700 : Loop

  ' ???????????? ???????????? ??????
  objIE.Top = 350
  objIE.Left = 100
  objIE.Height = 400
  objIE.Width  = 750
  
'' ??????? ??? ?????? ? IE
'  objIE.AddressBar = False
'  objIE.MenuBar = False
'  objIE.ToolBar = False
'  objIE.StatusBar  = False
'  objIE.RegisterAsBrowser = True
  objIE.Visible = True
 
'WScript.Quit 0

Dim Processes, myProcEnum, myProc, Proc
'------------------------------------------------------------------------------------------------------------------------------------------------
Set Processes = GetObject("winmgmts://localhost") 
Set myProcEnum = Processes.ExecQuery("select * from Win32_Process") 
myProc = False 
For Each Proc In myProcEnum 
If Proc.Name = "iexplore.exe" Then 
Proc.Terminate()
End If 
Next



