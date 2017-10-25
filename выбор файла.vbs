' �������� ����������� ���� ��� ����� ����� �� VBScript
Option Explicit
' �����.
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
file = BrowseForFoldr("�������� ���� ��� �����", BIF_returnonlyfsdirs + BIF_browseincludefiles, "")
If file = "-5" Then
WScript.Echo "������ ���� � �������� �����"
Else
If file = "-1" Then
WScript.Echo "������ �� ������"
Else
WScript.Echo "������: ", file
End If
End If
 
' ��������� ������� ���� � �������
Function BrowseForFoldr(title, flag, dir)
On Error Resume Next
Dim oShell, oItem, tmp
Set oShell = WScript.CreateObject("Shell.Application")
' ������� ���������� ���� Browse For Folder.
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
' ������ ���������� �������� ������ ����.
BrowseForFoldr = oItem.ParentFolder.ParseName(oItem.title).Path
' ��������� ������� ������ Cancel � ������ �����.

If Err<> 0 Then
If Err.Number = 424 Then                 ' ���������� ������ Cancel.
BrowseForFoldr ="-1"
Else
Err.Clear
' ���������� ��������, � ������� ������������ �������� ����.
tmp = InStr(1, oItem.title, ":")
If tmp > 0 Then          ' ������ ":" ������; ����� ��� ������� � �������� \.
BrowseForFoldr = Mid(oItem.Title, (tmp - 1), 2) & "\"
End If
End If
End If
Set oShell = Nothing
Set oItem = Nothing
On Error GoTo 0
End Function