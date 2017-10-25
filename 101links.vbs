on error resume next
dim links, keys
Set FSO = CreateObject("Scripting.FileSystemObject")
Set F = FSO.GetFile(Wscript.ScriptFullName)

path = FSO.GetParentFolderName(F)
keysFile = path & "\101keys.txt"
linksFile = path & "\101links.txt"
sitemapFile = path & "\sitemap.txt"

call deleteFile(FSO, keysFile)
call deleteFile(FSO, linksFile)
call deleteFile(FSO, sitemapFile)
call deleteFile(FSO, path & "\sitemap.xml")

call write ("<?xml version=""1.0"" encoding=""UTF-8""?>", sitemapFile, "")
call write ("<urlset", sitemapFile, "")
call write ("xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9""", sitemapFile, "")
call write ("xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""", sitemapFile, "")
call write ("xsi:schemaLocation=""http://www.sitemaps.org/schemas/sitemap/0.9", sitemapFile, "")
call write ("http://www.sitemaps.org/schemas/sitemap/0.9/sitemap.xsd"">", sitemapFile, "")
links = Array()
ReDim Preserve links(UBound(links) + 1)
links(0) = "http://xn--101-8cdj6bhvjbaxmo.xn--p1ai/"
keys = Array()
	
	call test (0, links)
call write ("</urlset>", sitemapFile, "")
Fso.MoveFile sitemapFile, path & "\sitemap.xml"
sub test(ind, byref links)

	if ind<=UBound(links) then
		response = getPage(links(ind))
		
		WScript.sleep 100
		respTuLinks = response
		respTuKeys = response

		getLinks respTuLinks
		'msgbox("respTuLinks = " & respTuLinks)
		
		'msgbox("respTuKeys = " & respTuKeys)
		getKeys respTuKeys, links(ind)
		call test (ind+1, links)

	end if
end sub
sub write(str, file, adr)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set TextStream = FSO.OpenTextFile(file, 8, True)
	
	'for i=0 to Ubound(links)
		TextStream.WriteLine adr+""+str'links(i)
	'next
	TextStream.Close
end sub

function getPage(link)
	getPage = ""
	Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")
	xmlHTTP.Open "GET", link, "false"
	xmlHTTP.Send
	Do While xmlHTTP.readystate <> 4
		WScript.sleep 100
	Loop
	rez = xmlHTTP.responseText 
	'msgbox(rez)
	getPage = rez
end function

sub getKeys(str, adr)
'msgbox(str)
'<meta name="keywords" content="новостройка, новостройки, квартиры в новостройках, купить новостройку, сайт новостройки">
if instr(str, "keywords")>0 then
'msgbox(mid(str, instr(str, "<meta name=""keywords""")+1,100))

	workArr = split(str, "<meta name=""keywords"" content=")
	'msgbox(workArr(1))
	'msgbox(instr(workArr(1), """/>"))
	keywords = mid(workArr(1),1, instr(workArr(1), """ />"))
	'msgbox(keywords)
	keywords = split(keywords, ",")
	for i=0 to UBound(keywords)
		if isExist(keywords(i), keys) = false then
			ReDim Preserve keys(UBound(keys) + 1)
			keys(Ubound(keys)) = keywords(i)
			call write (keywords(i), keysFile, adr)
		end if
	next
end if
end sub

sub getLinks(str)

	workArr = split(str, "</a>")
	for i=0 to UBound(workArr)
		indIn = instr(workArr(i), "<a href=")
		str = mid(workArr(i), indIn+9)
		indOut = instr(str, """")
		if len(str) > indOut-1 then
			str = left(str, indOut-1)
			if left(str,1) = "/" and len(mid(str, 2))>0 then
				if isExist(links(0) + mid(str, 2), links) = false then
					ReDim Preserve links(UBound(links) + 1)
					links(Ubound(links)) = links(0) + mid(str, 2)
					call write (links(0) + mid(str, 2), linksFile, "")
					if instr(str, "&") = 0 and instr(str, ".pdf") = 0 and instr(str, ".doc") = 0 and instr(str, ".jpg") = 0 and instr(str, ".png") = 0   then
						call write("<url><loc>" + "http://101новостройка.рф/" + mid(str, 2) + "</loc></url>", sitemapFile, "")
					end if
				end if
				'writeLinks links(0) + mid(str, 1)
			end if
		end if
	next
	
end sub

sub writeLinks(link)
	if link<>"" then
	msgbox(UBound(links))
		if isExist(link, links) = false then
			ReDim Preserve links(UBound(links) + 1)
			links(Ubound(links)) = link
			call write(link)
			msgbox(link)
		end if
	end if
end sub

function deleteFile(fso, name)
if fso.FileExists(name) then
    fso.DeleteFile name
end if
end function

function isExist(str, arr)
	isExist = false
	if isArray(arr) then
		for i=0 to UBound(arr)
			if arr(i) = str then
				isExist = true
				exit for
			end if
		next
	end if
end function
msgbox("true")
