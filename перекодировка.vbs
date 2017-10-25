    Set fso = CreateObject("Scripting.FileSystemObject")
    Set str = CreateObject("ADODB.Stream")
        path = "C:\1\"
        naim_file = "1150000002499==16022015.txt"
        str.Type = 2
        str.Charset = "Windows-1251"
        str.Open()
        str.LoadFromFile(Path & naim_file)
        Text = str.ReadText()
        str.Close()
        fso.DeleteFile(Path & naim_file)
        str.Charset = "CP 866"
        str.Open()
        str.WriteText(Text)
        str.SaveToFile Path & naim_file, 2
        str.Close()