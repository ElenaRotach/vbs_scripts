Dim Processes, myProcEnum, myProc, Proc
'------------------------------------------------------------------------------------------------------------------------------------------------
Set Processes = GetObject("winmgmts://localhost") 
Set myProcEnum = Processes.ExecQuery("select * from Win32_Process") 
myProc = False 
For Each Proc In myProcEnum 
If Proc.Name = "EXCEL.EXE" Then 
Proc.Terminate()
End If 
Next