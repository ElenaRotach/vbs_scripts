'>> Скрипт: "Скрипт1Отчет5000011179484366"
Dim CommonDialog 
Set CommonDialog = CreateObject("UserAccounts.CommonDialog") 
Set WinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
Set WebForm = New WebFormClass

Const cdlOFNExplorer = &H80000 
Const cdlOFNFileMustExist = &H1000 
Const cdlOFNHideReadOnly = &H4 
Const cdlOFNPathMustExist = &H800 

With CommonDialog 
    .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist 
    .Filter = "Файлы изображений (*.bmp;*.gif;*.jpg;*.jpeg;*.png)|*.bmp;*.gif;*.jpg;*.jpeg;*.png" 
    .ShowOpen 'Display dialog 
End With 

If Trim(CommonDialog.FileName) = "" Then WScript.Quit
result = upload(CommonDialog.FileName)
If Trim(result) = "" Then WScript.Quit
MsgBox "Файл успешно загружен на сервер."+vbCrLf+"Ссылка на файл:"+vbCrLf+vbCrLf+result,vbInformation,"Upload Complete"

Function upload(ByVal filepath)
    Dim action_url, p0, p1, s0
       action_url="http://pics.kz/"' или http://pics.kz/mini
        WebForm.Action = action_url
        WebForm.Method = "POST"
        WebForm.Enctype = "multipart/form-data"
        WebForm.AddFile "file_0", filepath

        WinHttp.Open WebForm.Method, WebForm.Action, False
        WinHttp.setRequestHeader "Content-type", WebForm.Enctype
        WinHttp.send WebForm.VarBody
        Select Case WinHttp.Status
        Case 200 'OK
            s0 = WinHttp.responseText '  возвращает ту же страницу (обработчик)
            wscript.echo WinHttp.responseText ' debug
            p0 = InStr(1, s0, "[URL=")
            If p0 Then
                p0 = p0 + 5 '5 = "[URL="
                p1 = InStr(p0, s0, "]")
                If p1 > p0 Then _
                    upload = Mid(s0, p0, p1 - p0)
            End If
        Case Else 'Error
            MsgBox WinHttp.statusText, vbCritical, "Error" + WinHttp.Status
        End Select
End Function


'/// Класс формы 
Class WebFormClass
    Private Fields, Files
    Private PropertyEnctype, PropertyMethod, PropertyBoundary, PropertyAction

    Private Sub Class_Initialize()
        Fields = Array()
        Files = Array()
        PropertyEnctype = "application/x-www-form-urlencoded"
        PropertyMethod = "GET"
        PropertyBoundary = String(27, "-") & GenerateBoundary
        PropertyAction = "about:blank"
    End Sub

    Public Property Let Action(Value)
        PropertyAction = Value
    End Property

    Public Property Get Action()
        Action = PropertyAction
        If PropertyMethod = "GET" Then
            Dim Params
            Params = VarBody
            If VarBody <> "" Then Action = Action & "?" & Params
        End If
    End Property

    Public Property Get Boundary()
        Boundary = PropertyBoundary
    End Property

    Public Property Get Method()
        Method = PropertyMethod
    End Property

    Public Property Let Method(Value)
        Value = UCase(Value)
        If Value = "GET" Or Value = "POST" Then PropertyMethod = Value
    End Property

    Public Property Get Enctype()
        Enctype = PropertyEnctype
        If PropertyEnctype = "multipart/form-data" Then Enctype = Enctype & "; boundary=" & PropertyBoundary
    End Property

    Public Property Let Enctype(Value)
        Value = LCase(Value)
        If Value = "multipart/form-data" Or Value = "application/x-www-form-urlencoded" Then PropertyEnctype = Value
    End Property

    Public Sub AddField(Name, Value)
        ReDim Preserve Fields(UBound(Fields) + 1)
        Fields(UBound(Fields)) = Array(Name, Value)
    End Sub

    Public Sub AddFile(Name, Value)
        ReDim Preserve Files(UBound(Files) + 1)
        Files(UBound(Files)) = Array(Name, Value)
    End Sub
      

    Public Property Get VarBody()
        If PropertyMethod = "POST" And PropertyEnctype = "multipart/form-data" Then
            Const DefaultBoundary = "--"
            Dim Stream
            Set Stream = CreateObject("ADODB.Stream")
            Stream.Type = 2
            Stream.Mode = 3
            Stream.Charset = "Windows-1251"
            Stream.Open
            
            Dim FieldHeader, FieldsBody
            
            For Each Field In Fields
                FieldHeader = "Content-Disposition: form-data; name=""" & Field(0) & """"
                FieldsBody = FieldsBody & DefaultBoundary & PropertyBoundary & vbCrLf & FieldHeader & vbCrLf & Field(1) & vbCrLf
            Next
            
            Stream.WriteText FieldsBody
            
            Dim FileHeader
            
            For Each File In Files
                If LoadFile(File(1), Data) Then
                    FileHeader = DefaultBoundary & Boundary & vbCrLf & "Content-Disposition: form-data; name=""" & File(0) & """; filename=""" & File(1) & """" & vbCrLf & "Content-Type: octet/stream" & vbCrLf & vbCrLf
                    Stream.WriteText FileHeader
                    Stream.Position = 0
                    Stream.Type = 1
                    Stream.Position = Stream.Size
                    Stream.write Data
                    Stream.Position = 0
                    Stream.Type = 2
                    Stream.Position = Stream.Size
                End If
            Next
            
            Stream.Position = 0
            Stream.Type = 2
            Stream.Position = Stream.Size
            Stream.WriteText vbCrLf & DefaultBoundary & PropertyBoundary & DefaultBoundary
            
            Stream.Position = 0
            Stream.Type = 1
            
            VarBody = Stream.Read
        Else
            For Each Field In Fields
                VarBody = VarBody & URLEncode(Field(0)) & "=" & URLEncode(Field(1)) & "&"
            Next
            For Each File In Files
                VarBody = VarBody & URLEncode(File(0)) & "=" & URLEncode(File(1)) & "&"
            Next
            If Len(VarBody) > 0 Then VarBody = Left(VarBody, Len(VarBody) - 1)
        End If
    End Property

    Private Function URLEncode(Data)
        Dim CharPosition, CharCode
        For CharPosition = 1 To Len(Data)
            CharCode = Asc(Mid(Data, CharPosition, 1))
            If CharCode = 32 Then
                URLEncode = URLEncode + "+"
            ElseIf (CharCode < 48 Or CharCode > 126) Or (CharCode > 56 And CharCode <= 64) Then
                URLEncode = URLEncode + "%" + Right("0" & Hex(CharCode), 2)
            Else
                URLEncode = URLEncode + Chr(CharCode)
            End If
        Next
    End Function

    Private Function LoadFile(Path, Data)
        On Error Resume Next
        Dim Stream
        Set Stream = CreateObject("ADODB.Stream")
        Stream.Type = 1
        Stream.Mode = 3
        Stream.Open
        Stream.LoadFromFile Path
        If Err.Number <> 0 Then Exit Function
        Data = Stream.Read
        LoadFile = True
    End Function

    Private Function GenerateBoundary()
        Dim Char
        Dim N, Start
        Const Chars = "abcdefghijklmnopqrstuvxyz0123456789"
        Randomize
        For N = 1 To 12
            Start = CLng(Rnd * (Len(Chars) - 1)) + 1
            Char = Mid(Chars, Start, 1)
            If Start Mod 2 Then Char = UCase(Char)
            GenerateBoundary = GenerateBoundary & Char
        Next
    End Function
End Class