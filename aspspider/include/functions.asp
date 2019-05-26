<%
'#################################################################################
'##
'## THE ASP SPIDER 
'##
'## Copyright (C) 2006 Arathra Ltd
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or any later version.
'##
'## All copyright notices regarding the ASP SPider
'## must remain intact in the scripts and the "powered by" text
'## with a link back to http://aspspider.co.uk in the footer of the pages MUST
'## remain visible when the pages are viewed on the internet or intranet.
'##
'## This program is distributed in the hope that it will be useful,
'## but WITHOUT ANY WARRANTY; without even the implied warranty of
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'## GNU General Public License for more details.
'##
'## You should have received a copy of the GNU General Public License
'## along with this program; if not, write to the Free Software
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'##
'## Support can be obtained from support forums at:
'## http://aspspider.co.uk
'##
'#################################################################################
%>
<!--#include file="userfunctions.asp" -->
<%
function MySQLDate(theDate)
	MySQLDate = year(theDate) & "-" & month(theDate) & "-" & day(theDate) & " " & hour(theDate) & ":" & minute(theDate) & ":" & second(theDate)
end function
%>
<%
function getMetaTag (tag, ByVal pageContent)
	getMetaTag = ""
	'add the quotes to the outside to find the right meta tag
	tag = "name=" & chr(34) & tag & chr(34)
	'quick check to see if the tag exists
	if InStr(pageContent, tag) = 0 then exit function
	dim strTag, tStart, tStop
	'get the start of the tag
	tStart = InStr(lcase(pageContent), tag )
	tStart = tStart + len(tag)
	'get the end of the tag
	tStop = InStr(tStart, lcase(pageContent), ">" )
	strTag = trim(Mid( pageContent, tStart, (tStop - tStart)))
	'Remove stuff
	strTag = Replace(strTag, Chr(34), "")
	strTag = Replace(strTag, "content=", "")
	getMetaTag = strTag
end function
%>
<%
Dim StopWatch(19)
sub StartTimer(x)
	StopWatch(x) = timer
end sub
function StopTimer(x)
	dim EndTime
	EndTime = Timer
   	if EndTime < StopWatch(x) then
    	EndTime = EndTime + (86400)
   	end if
   	StopTimer = EndTime - StopWatch(x)
end function
%>
<%
function AddDomains ()

	Dim rstRecords, domName
	Set rstRecords = Server.CreateObject("ADODB.Recordset")
	rstRecords.ActiveConnection = MM_conn_STRING
	rstRecords.Source = "SELECT fIndex, fPage FROM " & tblpages & " WHERE fDomain = ''"
	rstRecords.CursorType = 0
	rstRecords.CursorLocation = 2
	rstRecords.LockType = 1
	rstRecords.Open()

 	WHILE NOT rstRecords.EOF
		domName = GetDomain(rstRecords("fPage"))
		response.write(domName)
		response.write("<br>")
		strSQL = "UPDATE " & tblpages & " SET fDomain = '" & Replace(domName, "'", "''") & "' WHERE fIndex = " & rstRecords("fIndex") & ""
		runCmd (strSQL)
		rstRecords.MoveNext
	WEND	  

	rstRecords.Close
	set rstRecords = Nothing		
end function
%>


<%
'# this removes the default file of a directory
function removeDefault ( BYVAL theURL )
	'take off the trailing slash
	if instr ( len ( theURL ), theURL, "/") <> 0 then
		theURL = left ( theURL, len ( theURL ) - 1 )			
	end if
	theURL = Replace(theURL, "default.htm", "")
	theURL = Replace(theURL, "default.asp", "")
	theURL = Replace(theURL, "default.html", "")
	theURL = Replace(theURL, "default.php", "")
	theURL = Replace(theURL, "default.shtml", "")
	theURL = Replace(theURL, "default.pl", "")
	theURL = Replace(theURL, "index.htm", "")
	theURL = Replace(theURL, "index.asp", "")
	theURL = Replace(theURL, "index.html", "")
	theURL = Replace(theURL, "index.php", "")
	theURL = Replace(theURL, "index.shtml", "")
	theURL = Replace(theURL, "index.pl", "")
	removeDefault = theURL
end function
%>
<%
function checkContent ( pageContent )
	'remove bad stuff
	strSQL = "SELECT * FROM " & tbl404 & " ORDER BY fIndex DESC"
	Dim rstRecords
	Set rstRecords = Server.CreateObject("ADODB.Recordset")
	rstRecords.ActiveConnection = MM_conn_STRING
	rstRecords.Source = strSQL
	rstRecords.CursorType = 0
	rstRecords.CursorLocation = 2
	rstRecords.LockType = 1
	rstRecords.Open() 

	dim Command1
	WHILE (NOT rstRecords.EOF) AND (pageContent <> "")
		if instr(pageContent, rstRecords("fBadContent")) <> 0 then 
			pageContent = ""
		end if
	rstRecords.MoveNext
	WEND	  

	rstRecords.Close
	set rstRecords = Nothing
	checkContent = pageContent
	
end function
%>
<%
function AddPageContent (theURL, sTitle, pageContent)

	if pageContent = "" then exit function
	if theURL = "" then exit function

	'get the domain name - this DOES NOT include the WWW or HTTP
	dim domName, i
	domName = theURL
	domName = Replace ( domName, "www.", "" )
	domName = Replace ( domName, "http://", "" )
	i = instr ( domName, "/" )
	if i <> 0 then domName = left ( domName, i - 1 )
	
	theURL = Replace ( theURL, "http://", "" )
	theURL = "http://" & theURL

	dim ID
	ID = 0
ON ERROR RESUME NEXT	
	'check if the page is already there
	Dim rstRecords
	Set rstRecords = Server.CreateObject("ADODB.Recordset")
	rstRecords.ActiveConnection = MM_conn_STRING
	rstRecords.Source = "SELECT fIndex FROM " & tblpages & " WHERE fPage = '" & theURL & "'"
	rstRecords.CursorType = 0
	rstRecords.CursorLocation = 2
	rstRecords.LockType = 1
	rstRecords.Open()
	if NOT rstRecords.EOF then ID = rstRecords("fIndex")
	rstRecords.Close
	set rstRecords = Nothing		

	if ID = 0 then
		'put in new record
		strSQL = "INSERT INTO " & tblpages & " (fDomain, fPage, fTitle, fContent, fIndexDate)  VALUES ("
		strSQL = strSQL & "'" + Replace(domName, "'", "''") + "', "
		strSQL = strSQL & "'" + Replace(theURL, "'", "''") + "', "
		strSQL = strSQL & "'" + Replace(sTitle, "'", "''") + "', "
		strSQL = strSQL & "'" + Replace(pageContent, "'", "''") + "', "
		strSQL = strSQL & "'" & MySQLDate(Date()) & "')"
	else
		strSQL = "UPDATE " & tblpages & " SET "
		strSQL = strSQL & "fDomain = '" & Replace(domName, "'", "''") & "', "
		strSQL = strSQL & "fTitle = '" & Replace(Request("sTitle"), "'", "''") & "', "
		strSQL = strSQL & "fContent = '" & Replace(Request("pageContent"), "'", "''") & "' "
		strSQL = strSQL & "fIndexDate = '" & MySQLDate(Date()) & "' "
		strSQL = strSQL & "WHERE fIndex = " & ID & ""
	end if	
	runCmd strSQL
ON ERROR GOTO 0
end function
%>
<%
Function GetDomain ( BYVAL xURL )
	GetDomain = ""
	dim i
	GetDomain = Replace ( xURL, "www.", "" )
	GetDomain = Replace ( GetDomain, "http://", "" )
	i = instr ( GetDomain, "/" )
	if i <> 0 then GetDomain = left ( GetDomain, i - 1 )
End Function
%>
<%
function CollectLinks (pageContent, theURL)
	'this collects links from a page and puts them in the DB ready for spidering
	
	AddRow "URL to collect from", theURL, "#ffffff"
	
	'these are characters which make a link illegal
	ReDim aNaughtyLinks(3)
	aNaughtyLinks(0) = "javascript"		'javascript
	aNaughtyLinks(1) = "../"			'up a directory
	aNaughtyLinks(2) = "@"				'email add
	aNaughtyLinks(3) = "#"				'anchor on (presumably same page)
	
	'split the HTML into an array, each element beginning with the required HTML
	Dim aHTML
	aHTML = split ( pageContent, "<a href=" )
		
	Dim theDirectory
	theDirectory = theURL
	'if the last char is a slash then it's a directory and we can leave it
	if Right(theDirectory, 1) = "/" then
		theDirectory = left(theDirectory, len(theDirectory) - 1)
	else
		'if it's a file then move back to the directory
		if instr(theDirectory, ".asp") = 0 then 'do nothing
		elseif instr(theDirectory, ".htm") = 0 then 'do nothing
		elseif instr(theDirectory, ".html") = 0 then 'do nothing
		elseif instr(theDirectory, ".php") = 0 then 'do nothing
		elseif instr(theDirectory, ".txt") = 0 then 'do nothing
		elseif instr(theDirectory, ".xhtml") = 0 then 'do nothing
		elseif instr(theDirectory, ".shtml") = 0 then 'do nothing
		elseif instr(theDirectory, ".cfm") = 0 then 'do nothing
		else
			if instr(theDirectory, "/") <> 0 then
				while not Right(theDirectory, 1) = "/"
					theDirectory = left(theDirectory, len(theDirectory) - 1)
				wend
				theDirectory = left(theDirectory, len(theDirectory) - 1)
			end if		
		end if
	end if

	
	dim theDomain
	theDomain = GetDomain(theURL)
	
	'Response.Write ( " DOM : " & theDomain & "<br><br><br>" )
	
	'go through each item and see what we've got; we need to start at 1 rather than 0 because the first 
	'item in the array is going to be irrelevant HTML
	Dim foundURL
	Dim i, x
	for i = 1 to UBOUND ( aHTML, 1 )		
		foundURL = aHTML ( i )	
		'the link is limited by > at end of the link
		on error resume next
		foundURL = trim(left ( foundURL, instr (foundURL, ">" ) -1 ))
		on error goto 0
		AddRow "Found URL Trimmed >", foundURL, "#ffffff"
		
		'if there's a space in it then we can chop off there
		if instr(foundURL, " ") <> 0 then foundURL = left ( foundURL, instr (foundURL, " " ) )
		AddRow "Found URL Trimmed SPACE", foundURL, "#ffffff"
		
		'remove quote marks
		foundURL = Replace ( foundURL, Chr(34), "")
		foundURL = Replace ( foundURL, Chr(39), "")
		AddRow "Found URL Trimmed QUOTES", foundURL, "#ffffff"
					
		'make sure it's a proper file
		if instr(foundURL, ".asp") = 0 then 'do nothing
		elseif instr(foundURL, ".htm") = 0 then 'do nothing
		elseif instr(foundURL, ".html") = 0 then 'do nothing
		elseif instr(foundURL, ".php") = 0 then 'do nothing
		elseif instr(foundURL, ".txt") = 0 then 'do nothing
		elseif instr(foundURL, ".xhtml") = 0 then 'do nothing
		elseif instr(foundURL, ".shtml") = 0 then 'do nothing
		elseif instr(foundURL, ".cfm") = 0 then 'do nothing
		else
			foundURL = ""
			AddRow "Found URL", "Invalid file type", "#ffffff"
		end if
					
		'see if it's allowed by checking against naughty links
		for x = 0 to UBOUND ( aNaughtyLinks, 1 )
			if instr ( foundURL, aNaughtyLinks ( x ) ) <> 0 then 
				foundURL = ""
				AddRow "Naughty URL", aNaughtyLinks ( x ), "#ffffff"
				exit for
			end if
		next 

		if ExternalLink ( foundURL, theDomain ) = true then 
			foundURL = ""
			AddRow "Found URL", "External", "#ffffff"
		end if 
	
		'add the domain if required
		if foundURL <> "" then
			
			foundURL = removeDefault (foundURL)
		
			if instr(foundURL, "//") = 0 then
				if instr(foundURL, "/") = 1 then
					foundURL = "http://" & theDirectory & foundURL
				else
					foundURL = "http://" & theDirectory & "/" & foundURL
				end if
			else
				'is there http?
				foundURL = Replace(foundURL, "http://", "")
				foundURL = "http://" & foundURL
			end if
			
			'add to table to spider
			ON ERROR RESUME NEXT
				strSQL = "INSERT INTO " & tblp2s & " (fPage)  VALUES ('" + Replace(foundURL, "'", "''") + "') "
				runCmd strSQL
			ON ERROR GOTO 0
			
		end if	
		if foundURL <> "" then
			AddRow i & " - foundURL", foundURL, "#ffffff"		
		end if
	next
	AddRow "Links Collected", i, "#ffffff"		
end function
%>
<%
Function ExternalLink ( BYVAL xURL, BYVAL xDomain )

	'assume it is
	ExternalLink = true
	
	'do a simple check to see if xURL contains http; if it doesn't then it's an internal link
	if instr ( xURL, "http" ) = false then
		ExternalLink = false
		exit function
	end if
	
	'remove the www by default before comparing
	xURL = Replace ( xURL, "http://", "" )
	xURL = Replace ( xURL, "www.", "" )
	
	if instr ( xURL, xDomain ) <> 0 then
		ExternalLink = false
		exit function
	end if

End function
%>
<%
function GetURLContent (theURL)

	'get rid of the http if it's there
	theURL = Replace(theURL, "http://", "")

	Const Request_POST = 1
	Const Request_GET = 2
	
	dim xObj
	Set xObj = Server.CreateObject("SOFTWING.AspTear")	
	
	'some error handling in case it doesn't find the page we want
	On Error Resume Next
	GetURLContent = xObj.Retrieve("http://" & theURL, Request_GET, "", "", "")
	If Err.Number <> 0 Then
		GetURLContent = ""
	End If
	set xobj = nothing
end function
%>
<%
function delURLfromPageToSpider ( URLIndex )
	strSQL =  "DELETE FROM " & tblp2s & " WHERE fIndex = '" & URLIndex & "'"
	runCmd (strSQL)
end function
%>
<%
function delURLfromSpider ( byval theURL )
	strSQL = "DELETE FROM " & tbldomadd & " WHERE fURL = '" & theURL & "'"
	runCmd (strSQL)
end function
%>
<%
function AddURLtoSpider (theURL)
	'this gets an domain input by a user, checks it and adds it to the list of domains to be spidered
	theURL = lcase ( trim ( theURL ) )

	'remove the first http
	theURL = Replace(theURL, "http://", "")

	'check if there's something there
	if theURL = "" then 
		AddURLtoSpider = 0
		Exit Function
	end if
	
	'remove everything after the final slash
	dim i
	i = instr(theURL, "/")
	if i <> 0 then
		theURL = left(theURL, i - 1)
	end if
	'now add the final slash
	theURL = theURL & "/"
	
	'see if it's a subdomain; if not add a www to the beginning
	if instr(theURL, "www.") = 0 then		
		'see if it's a subdomain by counting the dots
		i = UBound(split(theURL, "."))
		if i = 1 then 
			theURL = "www." & theURL
		end if
	end if

	ON ERROR RESUME NEXT	
		strSQL = "INSERT INTO " & tbldomadd & " (fURL)  VALUES ('" + Replace(theURL, "'", "''") + "') "
		runCmd (strSQL)
	ON ERROR GOTO 0
	
	AddURLtoSpider = 1

End Function
%>
<%
Function RunRegEx ( patrn, strng )
	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern = patrn
	regEx.IgnoreCase = true
	regEx.Global = True   
	RunRegEx = regEx.Replace(patrn, strng)
	Set regEx = Nothing
End Function
%>
<%
Function stripHTML ( strHTML )
	
	'Strips the HTML tags from strHTML
	
	Dim objRegExp, strOutput
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global = True	
	
	strHTML = Replace(strHTML, "<br>", Chr(10))
		
	'must get rid of scripts first
	objRegExp.Pattern = "<SCRIPT(.|\n)+?/SCRIPT>"
	strHTML = objRegExp.Replace(strHTML, "")

	objRegExp.Pattern = "<(.|\n)+?>"	
	
	'Replace all HTML tag matches with the empty string
	strHTML = objRegExp.Replace(strHTML, " ")
	  
	'non whitespace chars
	objRegExp.Pattern = "\s"
	strHTML = objRegExp.Replace(strHTML, " ") 
  
	'Replace all < and > with &lt; and &gt;
'	strHTML = Replace(strOutput, "<", "&lt;")
'	strHTML = Replace(strOutput, ">", "&gt;")
 
	strHTML = Trim(strHTML)
    strHTML = Replace(strHTML, "&nbsp;", " ")
	strHTML = Replace(strHTML, "&amp;", Chr(38))
	strHTML = Replace(strHTML, "&middot;", ".")
	strHTML = Replace(strHTML, "&copy; ", " (c) ")
	strHTML = Replace(strHTML, "&reg;", " (r) ")
	strHTML = Replace(strHTML, "&#8482;", " (tm) ")
	strHTML = Replace(strHTML, "&pound;", Chr(156))
	strHTML = Replace(strHTML, "&#8364;", " E ")
	strHTML = Replace(strHTML, "mmLoadMenus();", " ")
	strHTML = Replace(strHTML, "&acute;", Chr(39))
	
	while instr(strHTML, "  ") > 0
	    strHTML = Replace(strHTML, "  ", " ")
	wend 

  	stripHTML = trim ( strHTML )   'Return the value of strHTML

  	Set objRegExp = Nothing
  
End Function
%>	
