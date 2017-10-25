Class VbsJson
    'Author: Demon
    'Date: 2012/5/3
    'Website: http://demon.tw
    Private Whitespace, NumberRegex, StringChunk
    Private b, f, r, n, t

    Private Sub Class_Initialize
        Whitespace = " " & vbTab & vbCr & vbLf
        b = ChrW(8)
        f = vbFormFeed
        r = vbCr
        n = vbLf
        t = vbTab

        Set NumberRegex = New RegExp
        NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
        NumberRegex.Global = False
        NumberRegex.MultiLine = True
        NumberRegex.IgnoreCase = True

        Set StringChunk = New RegExp
        StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
        StringChunk.Global = False
        StringChunk.MultiLine = True
        StringChunk.IgnoreCase = True
    End Sub
    
    'Return a JSON string representation of a VBScript data structure
    'Supports the following objects and types
    '+-------------------+---------------+
    '| VBScript          | JSON          |
    '+===================+===============+
    '| Dictionary        | object        |
    '+-------------------+---------------+
    '| Array             | array         |
    '+-------------------+---------------+
    '| String            | string        |
    '+-------------------+---------------+
    '| Number            | number        |
    '+-------------------+---------------+
    '| True              | true          |
    '+-------------------+---------------+
    '| False             | false         |
    '+-------------------+---------------+
    '| Null              | null          |
    '+-------------------+---------------+
    Public Function Encode(ByRef obj)
        Dim buf, i, c, g
        Set buf = CreateObject("Scripting.Dictionary")
        Select Case VarType(obj)
            Case vbNull
                buf.Add buf.Count, "null"
            Case vbBoolean
                If obj Then
                    buf.Add buf.Count, "true"
                Else
                    buf.Add buf.Count, "false"
                End If
            Case vbInteger, vbLong, vbSingle, vbDouble
                buf.Add buf.Count, obj
            Case vbString
                buf.Add buf.Count, """"
                For i = 1 To Len(obj)
                    c = Mid(obj, i, 1)
                    Select Case c
                        Case """" buf.Add buf.Count, "\"""
                        Case "\"  buf.Add buf.Count, "\\"
                        Case "/"  buf.Add buf.Count, "/"
                        Case b    buf.Add buf.Count, "\b"
                        Case f    buf.Add buf.Count, "\f"
                        Case r    buf.Add buf.Count, "\r"
                        Case n    buf.Add buf.Count, "\n"
                        Case t    buf.Add buf.Count, "\t"
                        Case Else
                            If AscW(c) >= 0 And AscW(c) <= 31 Then
                                c = Right("0" & Hex(AscW(c)), 2)
                                buf.Add buf.Count, "\u00" & c
                            Else
                                buf.Add buf.Count, c
                            End If
                    End Select
                Next
                buf.Add buf.Count, """"
            Case vbArray + vbVariant
                g = True
                buf.Add buf.Count, "["
                For Each i In obj
                    If g Then g = False Else buf.Add buf.Count, ","
                    buf.Add buf.Count, Encode(i)
                Next
                buf.Add buf.Count, "]"
            Case vbObject
                If TypeName(obj) = "Dictionary" Then
                    g = True
                    buf.Add buf.Count, "{"
                    For Each i In obj
                        If g Then g = False Else buf.Add buf.Count, ","
                        buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
                    Next
                    buf.Add buf.Count, "}"
                Else
                    Err.Raise 8732,,"None dictionary object"
                End If
            Case Else
                buf.Add buf.Count, """" & CStr(obj) & """"
        End Select
        Encode = Join(buf.Items, "")
    End Function

    'Return the VBScript representation of ``str(``
    'Performs the following translations in decoding
    '+---------------+-------------------+
    '| JSON          | VBScript          |
    '+===============+===================+
    '| object        | Dictionary        |
    '+---------------+-------------------+
    '| array         | Array             |
    '+---------------+-------------------+
    '| string        | String            |
    '+---------------+-------------------+
    '| number        | Double            |
    '+---------------+-------------------+
    '| true          | True              |
    '+---------------+-------------------+
    '| false         | False             |
    '+---------------+-------------------+
    '| null          | Null              |
    '+---------------+-------------------+
    Public Function Decode(ByRef str)
        Dim idx
        idx = SkipWhitespace(str, 1)

        If Mid(str, idx, 1) = "{" Then
            Set Decode = ScanOnce(str, 1)
        Else
            Decode = ScanOnce(str, 1)
        End If
    End Function
    
    Private Function ScanOnce(ByRef str, ByRef idx)
        Dim c, ms

        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "{" Then
            idx = idx + 1
            Set ScanOnce = ParseObject(str, idx)
            Exit Function
        ElseIf c = "[" Then
            idx = idx + 1
            ScanOnce = ParseArray(str, idx)
            Exit Function
        ElseIf c = """" Then
            idx = idx + 1
            ScanOnce = ParseString(str, idx)
            Exit Function
        ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = Null
            Exit Function
        ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = True
            Exit Function
        ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
            idx = idx + 5
            ScanOnce = False
            Exit Function
        End If
        
        Set ms = NumberRegex.Execute(Mid(str, idx))
        If ms.Count = 1 Then
            idx = idx + ms(0).Length
            ScanOnce = CDbl(ms(0))
            Exit Function
        End If
        
        Err.Raise 8732,,"No JSON object could be ScanOnced"
    End Function

    Private Function ParseObject(ByRef str, ByRef idx)
        Dim c, key, value
        Set ParseObject = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)
        
        If c = "}" Then
            Exit Function
        ElseIf c <> """" Then
            Err.Raise 8732,,"Expecting property name"
        End If

        idx = idx + 1
        
        Do
            key = ParseString(str, idx)

            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) <> ":" Then
                Err.Raise 8732,,"Expecting : delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            ParseObject.Add key, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "}" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            c = Mid(str, idx, 1)
            If c <> """" Then
                Err.Raise 8732,,"Expecting property name"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
    End Function
    
    Private Function ParseArray(ByRef str, ByRef idx)
        Dim c, values, value
        Set values = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "]" Then
            ParseArray = values.Items
            Exit Function
        End If

        Do
            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            values.Add values.Count, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "]" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
        ParseArray = values.Items
    End Function
    
    Private Function ParseString(ByRef str, ByRef idx)
        Dim chunks, content, terminator, ms, esc, char
        Set chunks = CreateObject("Scripting.Dictionary")

        Do
            Set ms = StringChunk.Execute(Mid(str, idx))
            If ms.Count = 0 Then
                Err.Raise 8732,,"Unterminated string starting"
            End If
            
            content = ms(0).Submatches(0)
            terminator = ms(0).Submatches(1)
            If Len(content) > 0 Then
                chunks.Add chunks.Count, content
            End If
            
            idx = idx + ms(0).Length
            
            If terminator = """" Then
                Exit Do
            ElseIf terminator <> "\" Then
                Err.Raise 8732,,"Invalid control character"
            End If
            
            esc = Mid(str, idx, 1)

            If esc <> "u" Then
                Select Case esc
                    Case """" char = """"
                    Case "\"  char = "\"
                    Case "/"  char = "/"
                    Case "b"  char = b
                    Case "f"  char = f
                    Case "n"  char = n
                    Case "r"  char = r
                    Case "t"  char = t
                    Case Else Err.Raise 8732,,"Invalid escape"
                End Select
                idx = idx + 1
            Else
                char = ChrW("&H" & Mid(str, idx + 1, 4))
                idx = idx + 5
            End If

            chunks.Add chunks.Count, char
        Loop

        ParseString = Join(chunks.Items, "")
    End Function

    Private Function SkipWhitespace(ByRef str, ByVal idx)
        Do While idx <= Len(str) And _
            InStr(Whitespace, Mid(str, idx, 1)) > 0
            idx = idx + 1
        Loop
        SkipWhitespace = idx
    End Function

End Class

Function Init
	Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")
		adr = "http://185.27.195.77:9000/export_json.php"
		auth=""

		xmlHTTP.Open "GET", adr, False
		xmlHTTP.Send'(auth)
		rez = xmlHTTP.getAllResponseHeaders
		rez = Split(rez, Chr(10))
		Set heders = CreateObject("Scripting.Dictionary")
		For i=0 To UBound(rez)-2
			str = Split(rez(i), ": ")
			If heders(str(0)) = False Then
				heders(str(0)) = str(1)
			End If
		Next
		Do While xmlHTTP.readystate <> 4
		Loop
        coc = Split(heders("Set-Cookie"), ";")
        cookie = Mid(coc(0), InStr(coc(0), "=")+1)
        Set md5 = new md5Hash
        md5.Initialize(cookie & "tatsoc")
    Set Init = md5
End Function

Function AllForTheDays(operDate1, operDate2)
    Set md5 = Init
    Set uslColl = new serviceCollection
    uslColl.Init md5

    allTrTest = Array()
    For IndUsl=0 To uslColl.count-1
        Set trtest = new transactionList
        Set md5 = Init
        trtest.Init(md5)
        Set id = uslColl.collection(IndUsl)
        trtest.allStartEnd id.id, operDate1, operDate2
        For Ind=0 To UBound(trtest.collection)
            ReDim Preserve allTrTest(UBound(allTrTest)+1)
            Set allTrTest(UBound(allTrTest)) = trtest.collection(Ind)            
        Next
    Next
	AllForTheDays = allTrTest
End Function
Class md5Hash
    Public hash
    Public Sub Initialize(str) 

        Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")
        adr="http://www.md5online.org/md5-encrypt.html"
        num = str
        auth="md5=" & num & "&action=encrypt&a=11111111"

        xmlHTTP.Open "POST", adr, False
        xmlHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        xmlHTTP.setRequestHeader "X-Requested-With", "XMLHttpRequest"
        xmlHTTP.Send(auth)
        rez = xmlHTTP.responseText 
        serchStr = "MD5 hash for " & num & " is : "
        serchInd = InStr(rez, serchStr)
        If serchInd > 0 Then
            rez = Mid(rez, Len(serchStr)+serchInd+3, 32)
            hash = rez
        End If

    End Sub
End Class
Class serviceCollection
    Public count, collection, md5
    Private xmlHTTP
    Public Sub Init(md5) 
        count = 0
        WScript.Sleep 1000

		
		Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")

        console = "console=0"
        PHPSESSID = "PHPSESSID=" & cookie
        collection = Array()
        adr = "http://185.27.195.77:9000/export_json.php"
        param = "{""USER"":""tatsoc"",""PWD"":""" & md5.hash & """,""CMD"":""list""}"
			 xmlHTTP.Open "POST", adr, "False"

        xmlHTTP.SetRequestHeader "Cookie", console' PHPSESSID=" & cookie
        xmlHTTP.SetRequestHeader "Cookie", PHPSESSID
        xmlHTTP.SetRequestHeader "Content-Length", CStr(Len(param))

        xmlHTTP.Send CStr(param)
			 Do While xmlHTTP.readystate <> 4
			 Loop
			 rez = xmlHTTP.responseText 
        Set json = New VbsJson
        Set o = json.Decode(rez)
        If o("RESULT") = "0" Then
        
        For Each usl In o("DATA")
            Set uslInColl = new service
            uslInColl.Initialize usl("ID"), usl("Name"), usl("Time")

            ReDim Preserve collection(UBound(collection) + 1)                                         
            Set collection(UBound(collection)) = uslInColl
            count = count + 1
        Next
    End If

    End Sub
    Public Function GetElementByNum(value)
        Set usluga = new service
        usluga.Initialize collection(value).id, collection(value).name, collection(value).TimeS
        GetElementByNum = usluga
    End Function
    Public Property Get GetElementById 
        For Each usl In collection
            If usl.key = value Then : GetElementById = usl : End If
        Next
    End Property
    Public Property Get GetElementByName 
        For Each usl In collection
            If usl.value = value Then : GetElementByName = usl : End If
        Next
    End Property
End Class
Class service
    Public id, name, TimeS
    Public Sub Initialize(idServ, nameServ, TimeServ)
        id = idServ : name = nameServ : TimeS = TimeServ
    End Sub
End Class	

Class transactionList
    Public collection
    Private adr, md5, xmlHTTP
	  Private Sub Class_Initialize() 
        Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")

        count = 0
        collection = Array()
        adr = "http://185.27.195.77:9000/export_json.php"
    End Sub
    Public Sub Init(md)
    Set md5 = md
    End Sub
    Public Sub allnew(id)
        param = "{""USER"":""tatsoc"",""PWD"":""" & md5.hash & """,""CMD"":""export"",""ID"":""" & id & """}   "
			 xmlHTTP.Open "POST", adr, "false"
        xmlHTTP.SetRequestHeader "Content-Length", CStr(Len(param))
        xmlHTTP.SetRequestHeader "Cookie", Chr(34) & "console=0" & Chr(34)' PHPSESSID=" & cookie
        xmlHTTP.SetRequestHeader "Cookie", Chr(34) & "PHPSESSID=" & cookie & Chr(34)
        xmlHTTP.Send CStr(param)
			 Do While xmlHTTP.readystate <> 4
			 Loop
        rez = xmlHTTP.responseText 
        Set json = New VbsJson
        Set o = json.Decode(rez)
        If o("RESULT") = "0" Then
            For Each tr In o("DATA")
                Set newTransaction = new transaction
                newTransaction.Initialize tr("TransactionID"), tr("OperationID"), tr("TerminalID"), tr("NoticeType"), tr("NoticeCard"), tr("NoticeAmount"), tr("NoticePaymentNumber"), _
                                          tr("NoticeData1"), tr("NoticeData2"), tr("NoticeData3"), tr("NoticeData4"), tr("NoticeData5"), tr("ResponseCode"), tr("ResponseNumber"), _
                                          tr("ResponseText"), tr("ResponseData1"), tr("ResponseData2"), tr("ResponseData3"), tr("ResponseData4"), tr("ResponseData5"), tr("ResponseData6"), _
                                          tr("ResponseData7"), tr("ResponseData8"), tr("ResponseData9"), tr("ResponseData10"), tr("TimeT"), tr("NoticeDatetime"), tr("PackData")
                ReDim Preserve collection(UBound(collection) + 1)                                         
                    Set collection(UBound(collection)) = newTransaction
                    count = count + 1
            Next
        End If
    End Sub
    Public Sub allSTART(id, DateStart)
        param = "{""USER"":""tatsoc"",""PWD"":""" & md5.hash & """,""CMD"":""export"",""ID"":""" & id & """,""START"":""" & DateStart & """}   "
			 xmlHTTP.Open "POST", adr, "false"
        xmlHTTP.SetRequestHeader "Content-Length", CStr(Len(param))
        xmlHTTP.SetRequestHeader "Cookie", Chr(34) & "console=0" & Chr(34)' PHPSESSID=" & cookie
        xmlHTTP.SetRequestHeader "Cookie", Chr(34) & "PHPSESSID=" & cookie & Chr(34)
        xmlHTTP.Send CStr(param)
			 Do While xmlHTTP.readystate <> 4
			 Loop
        rez = xmlHTTP.responseText 
        Set json = New VbsJson
        Set o = json.Decode(rez)
        If o("RESULT") = "0" Then
            For Each tr In o("DATA")
                Set newTransaction = new transaction
                newTransaction.Initialize tr("TransactionID"), tr("OperationID"), tr("TerminalID"), tr("NoticeType"), tr("NoticeCard"), tr("NoticeAmount"), tr("NoticePaymentNumber"), _
                                          tr("NoticeData1"), tr("NoticeData2"), tr("NoticeData3"), tr("NoticeData4"), tr("NoticeData5"), tr("ResponseCode"), tr("ResponseNumber"), _
                                          tr("ResponseText"), tr("ResponseData1"), tr("ResponseData2"), tr("ResponseData3"), tr("ResponseData4"), tr("ResponseData5"), tr("ResponseData6"), _
                                          tr("ResponseData7"), tr("ResponseData8"), tr("ResponseData9"), tr("ResponseData10"), tr("TimeT"), tr("NoticeDatetime"), tr("PackData")
                ReDim Preserve collection(UBound(collection) + 1)                                         
                    Set collection(UBound(collection)) = newTransaction
                    count = count + 1
            Next
        End If
    End Sub
    Public Sub allStartEnd(id, DateStart, DateEnd)
        param = "{""USER"":""tatsoc"",""PWD"":""" & md5.hash & """,""CMD"":""export"",""ID"":""" & id & """,""START"":""" & DateStart & """,""END"":""" & DateEnd & """}   "
			 xmlHTTP.Open "POST", adr, "false"
        xmlHTTP.SetRequestHeader "Content-Length", CStr(Len(param))
        xmlHTTP.SetRequestHeader "Cookie", Chr(34) & "console=0" & Chr(34)' PHPSESSID=" & cookie
        xmlHTTP.SetRequestHeader "Cookie", Chr(34) & "PHPSESSID=" & cookie & Chr(34)
        xmlHTTP.Send CStr(param)
			 Do While xmlHTTP.readystate <> 4
			 Loop
        rez = xmlHTTP.responseText 
        Set json = New VbsJson
        Set o = json.Decode(rez)
        If o("RESULT") = "0" Then
            For Each tr In o("DATA")
                Set newTransaction = new transaction
                newTransaction.Initialize tr("TransactionID"), tr("OperationID"), tr("TerminalID"), tr("NoticeType"), tr("NoticeCard"), tr("NoticeAmount"), tr("NoticePaymentNumber"), _
                                          tr("NoticeData1"), tr("NoticeData2"), tr("NoticeData3"), tr("NoticeData4"), tr("NoticeData5"), tr("ResponseCode"), tr("ResponseNumber"), _
                                          tr("ResponseText"), tr("ResponseData1"), tr("ResponseData2"), tr("ResponseData3"), tr("ResponseData4"), tr("ResponseData5"), tr("ResponseData6"), _
                                          tr("ResponseData7"), tr("ResponseData8"), tr("ResponseData9"), tr("ResponseData10"), tr("TimeT"), tr("NoticeDatetime"), tr("PackData")
                ReDim Preserve collection(UBound(collection) + 1)                                         
                    Set collection(UBound(collection)) = newTransaction
                    count = count + 1
            Next
        End If
    End Sub
    Public Sub deleteCollection
        ReDim collection(-1)
    End Sub
End Class
Class transaction
    'Public TransactionID, NoticeData1, NoticeData2, NoticeData3, NoticeData4, NoticeData5, Amount, Comission, Time_tr, ResponseNumber, ITHostProduct, ITHostProductID, ProductID, OperationID, NoticeType
    Public TransactionID, OperationID, TerminalID, NoticeType, NoticeCard, NoticeAmount, NoticePaymentNumber, NoticeData1, NoticeData2, NoticeData3, NoticeData4, NoticeData5, ResponseCode, ResponseNumber
    Public ResponseText, ResponseData1, ResponseData2, ResponseData3, ResponseData4, ResponseData5, ResponseData6, ResponseData7, ResponseData8, ResponseData9, ResponseData10, TimeT, NoticeDatetime, PackData
    Public Sub Initialize(id, opId, termID, nt, nCard, nAm, nPN, d1, d2, d3, d4, d5, rCode, rNum, rText, rd1, rd2, rd3, rd4, rd5, rd6, rd7, rd8, rd9, rd10, TimeTr, nDT, pD)
        TransactionID = id : OperationID = opID : TerminalID = termID : NoticeType = nt : NoticeCard = nCard : NoticeAmount = nAm : NoticePaymentNumber = nPN : NoticeData1 = d1 : NoticeData2 = d2 
        NoticeData3 = d3 : NoticeData4 = d4 : NoticeData5 = d5 : ResponseCode = rCode : ResponseNumber = rNum : ResponseText = rText : ResponseData1 = rd1 : ResponseData2 = rd2 : ResponseData3 = rd3
        ResponseData4 = rd4 : ResponseData5 = rd5 : ResponseData6 = rd6 : ResponseData7 = rd7 : ResponseData8 = rd8 : ResponseData9 = rd9 : ResponseData10 = rd10 : TimeT = TimeTr : NoticeDatetime = nDT : PackData = pD
    End Sub
End Class


operDate1 = "2017-07-21"
operDate2 = "2017-07-22"

allTrTest = AllForTheDays(operDate1, operDate2)

    Set fso = CreateObject("Scripting.FileSystemObject")
        fileName = "c:\\123\r.txt"

        If fso.FileExists(fileName) Then
            fso.deletefile fileName
        End If

        Set file = fso.OpenTextFile(fileName,8,True) 'G2C.GetSysParam("Р Р°Р±РѕС‡Р°СЏ РїР°РїРєР°") + "РѕР±РѕСЂРѕС‚С‹409.txt",8, True)
            Set tools = CreateObject("WFinTools.ComTools")
        file.WriteLine "Все операции за "
		file.WriteLine "идентификатор транзакции|состояние операции|внутренний идентификатор киоска|тип транзакции|маскированный номер карты|сумма операции|уникальный идентификатор" & _
						"NoticeData1|NoticeData2|NoticeData3|NoticeData4|NoticeData5|код обработки транзакции|уникальный идентификатор|текстовый результат обработки транзакции|" & _
						"ResponseData1|ResponseData2|ResponseData3|ResponseData4|ResponseData5|ResponseData6|ResponseData7|ResponseData8|ResponseData9|ResponseData10" & _
						"время фиксирования транзакции на сервере|время отправки транзакции с киоска|поле расширенных параметров"
       ' file.WriteLine Tools.sFormat("идентификатор транзакции",30, "l", 0, " ") & Tools.sFormat("состояние операции",30, "l", 0, " ") & Tools.sFormat("внутренний идентификатор киоска",30, "l", 0, " ") & _ 
        'Tools.sFormat("тип транзакции",30, "l", 0, " ") & Tools.sFormat("маскированный номер карты",30, "l", 0, " ") & Tools.sFormat("сумма операции",30, "l", 0, " ") & _
        'Tools.sFormat("уникальный идентификатор",30, "l", 0, " ") & Tools.sFormat("NoticeData1",30, "l", 0, " ") & Tools.sFormat("NoticeData2",30, "l", 0, " ") & Tools.sFormat("NoticeData3",30, "l", 0, " ") & _
        'Tools.sFormat("NoticeData4",30, "l", 0, " ") & Tools.sFormat("NoticeData5 ",30, "l", 0, " ") & Tools.sFormat(" код обработки транзакции",30, "l", 0, " ") & Tools.sFormat(" уникальный идентификатор",30, "l", 0, " ") & _
        'Tools.sFormat("текстовый результат обработки транзакции",30, "l", 0, " ") & Tools.sFormat("ResponseData1",30, "l", 0, " ") & Tools.sFormat("ResponseData2",30, "l", 0, " ") & _
        'Tools.sFormat("ResponseData3",30, "l", 0, " ") & Tools.sFormat("ResponseData4",30, "l", 0, " ") & Tools.sFormat("ResponseData5",30, "l", 0, " ") & Tools.sFormat("ResponseData6",30, "l", 0, " ") & _
        'Tools.sFormat("ResponseData7",30, "l", 0, " ") & Tools.sFormat("ResponseData8",30, "l", 0, " ") & Tools.sFormat("ResponseData9",30, "l", 0, " ") & Tools.sFormat("ResponseData10 ",30, "l", 0, " ") & _
        'Tools.sFormat(" время фиксирования транзакции на сервере",30, "l", 0, " ") & Tools.sFormat("время отправки транзакции с киоска",30, "l", 0, " ") & Tools.sFormat(" поле расширенных параметров",30, "l", 0, " ")



        For z=0 To UBound(allTrTest)
			file.WriteLine allTrTest(z).TransactionID & "|" & allTrTest(z).OperationID & "|" & allTrTest(z).TerminalID & "|" & allTrTest(z).NoticeType & "|" & allTrTest(z).NoticeCard & "|" & _ 
			allTrTest(z).NoticeAmount & "|" & allTrTest(z).NoticePaymentNumber & "|" & allTrTest(z).NoticeData1 & "|" & allTrTest(z).NoticeData2 & "|" & allTrTest(z).NoticeData3 & "|" & _ 
			allTrTest(z).NoticeData4 & "|" & allTrTest(z).NoticeData5 & "|" & allTrTest(z).ResponseCode & "|" & allTrTest(z).ResponseNumber & "|" & allTrTest(z).ResponseText & "|" & _
			allTrTest(z).ResponseData1 & "|" & allTrTest(z).ResponseData2 & "|" & allTrTest(z).ResponseData3 & "|" & allTrTest(z).ResponseData4 & "|" & allTrTest(z).ResponseData5 & "|" & _
			allTrTest(z).ResponseData6 & "|" & allTrTest(z).ResponseData7 & "|" & allTrTest(z).ResponseData8 & "|" & allTrTest(z).ResponseData9 & "|" & allTrTest(z).ResponseData10 & "|" & _
			allTrTest(z).TimeT & "|" & allTrTest(z).NoticeDatetime & "|" & allTrTest(z).PackData
		            'file.WriteLine Tools.sFormat(allTrTest(z).TransactionID,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).OperationID,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).TerminalID,30, "l", 0, " ") & _
                            'Tools.sFormat(allTrTest(z).NoticeType,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).NoticeCard,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).NoticeAmount,30, "l", 0, " ") & _
                            'Tools.sFormat(allTrTest(z).NoticePaymentNumber ,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).NoticeData1,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).NoticeData2,30, "l", 0, " ") & _
                            'Tools.sFormat(allTrTest(z).NoticeData3,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).NoticeData4,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).NoticeData5 ,30, "l", 0, " ") & _
                            'Tools.sFormat(allTrTest(z).ResponseCode ,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).ResponseNumber ,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).ResponseText ,30, "l", 0, " ") & _
                            'Tools.sFormat(allTrTest(z).ResponseData1,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).ResponseData2,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).ResponseData3,30, "l", 0, " ") & _
                            'Tools.sFormat(allTrTest(z).ResponseData4,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).ResponseData5,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).ResponseData6,30, "l", 0, " ") & _
                            'Tools.sFormat(allTrTest(z).ResponseData7,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).ResponseData8,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).ResponseData9,30, "l", 0, " ") & _
                            'Tools.sFormat(allTrTest(z).ResponseData10 ,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).TimeT,30, "l", 0, " ") & Tools.sFormat(allTrTest(z).NoticeDatetime,30, "l", 0, " ") & _
                            'Tools.sFormat(allTrTest(z).PackData ,30, "l", 0, " ")
            
        Next
    file.Close
tools.RunShell "c:\program files\wfininst\wpad.exe " & Chr(34) & fileName & Chr(34), 1