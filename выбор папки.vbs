    on error resume next
    Dim oFolder: Set oFolder = CreateObject("Shell.Application").BrowseForFolder(0, "����� ����� � ������� ��� ������ XML ���� AVZ", 16 + 16384, StartFolder)
    If not (oFolder is Nothing) Then OpenFileDialogue = oFolder.Self.Path
    if Err.Number <> 0 or len(OpenFileDialogue) = 0 then msgbox "�������� ����� ������ ����� !",,"ALF": WScript.Quit 1
        msgbox OpenFileDialogue
    set oFolder = Nothing