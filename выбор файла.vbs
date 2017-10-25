' Создание диалогового окна для выбор файла на VBScript
Option Explicit
' Флаги.
Const BIF_returnonlyfsdirs       = &H0001
Const BIF_dontgobelowdomain           = &H0002
Const BIF_statustext                 = &H0004
Const BIF_returnfsancestors      = &H0008
Const BIF_editbox                     = &H0010
Const BIF_validate                    = &H0020
Const BIF_browseforcomputer  = &H1000
Const BIF_browseforprinter      = &H2000
Const BIF_browseincludefiles   = &H4000
Dim file
file = BrowseForFoldr("Выберите файл или папку", BIF_returnonlyfsdirs + BIF_browseincludefiles, "")
If file = "-5" Then
WScript.Echo "Выбран файл в корневой папке"
Else
If file = "-1" Then
WScript.Echo "Объект не выбран"
Else
WScript.Echo "Объект: ", file
End If
End If
 
' Получение полного пути к объекту
Function BrowseForFoldr(title, flag, dir)
On Error Resume Next
Dim oShell, oItem, tmp
Set oShell = WScript.CreateObject("Shell.Application")
' Взывать диалоговое окно Browse For Folder.
Set oItem = oShell.BrowseForFolder(&H0, title, flag, dir)
If Err.Number <> 0 Then
If Err.Number = 5 Then
BrowseForFoldr="-5"
Err.Clear
Set oShell = Nothing
Set oItem = Nothing
Exit Function
End If
End If
' Теперь попытаемся получить полный путь.
BrowseForFoldr = oItem.ParentFolder.ParseName(oItem.title).Path
' Обработка нажатия кнопки Cancel и выбора диска.

If Err<> 0 Then
If Err.Number = 424 Then                 ' Обработать кнопку Cancel.
BrowseForFoldr ="-1"
Else
Err.Clear
' Обработать ситуацию, в которой пользователь выбирает диск.
tmp = InStr(1, oItem.title, ":")
If tmp > 0 Then          ' Символ ":" найден; взять два символа и добавить \.
BrowseForFoldr = Mid(oItem.Title, (tmp - 1), 2) & "\"
End If
End If
End If
Set oShell = Nothing
Set oItem = Nothing
On Error GoTo 0
End Function