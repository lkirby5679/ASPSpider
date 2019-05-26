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
<%
	dim domGroup
	domGroup = trim(Request("dom"))

	dim sTerm
	sTerm = Request("fSearch")
	sTerm = lcase(trim(Replace(sTerm, "'", "''")))
	if sTerm = "" then
		Response.Redirect("results.asp?flag=1")	
	end if
	
	'update the search term table
	strSQL = "SELECT fCount  FROM " & tblsterm & "  WHERE fSearchTerm = '" & sTerm & "'"
	Dim rstST
	Set rstST = Server.CreateObject("ADODB.Recordset")
	rstST.ActiveConnection = MM_conn_STRING
	rstST.Source = strSQL
	rstST.CursorType = 0
	rstST.CursorLocation = 2
	rstST.LockType = 1
	rstST.Open()
	'if it's new add it in
	if rstST.EOF = true then
		strSQL = "INSERT INTO " & tblsterm & " (fSearchTerm, fCount)  VALUES ('" & sTerm & "', 1 ) "
		runCmd (strSQL)
	'it's already been done before, so update the count
	else
		stRSQL = "UPDATE " & tblsterm & "  SET fCount = fCount +1 WHERE fSearchTerm = '" & sTerm & "'"
		runCmd (strSQL)
	end if
	rstST.Close()
	Set rstST = Nothing

	'if the length of ANY word is less than 3 letters we must do a slow search
	dim slowsearch, iStart, iStop
	slowsearch = false
	iStart = 1
	iStop = 1
	do while instr(iStart, sTerm, " ") > iStart
		iStop = instr(iStart, sTerm, " ")
		if iStop - iStart < 4 then slowsearch = true
		iStart = iStop
	loop
	if len(sTerm) < 4 then slowsearch = true

	'the main search block
	Dim rstResults
	Dim rstResults_numRows
	
	'are there any quotes there already? If so, don't add them on artificially but take the ss as a literal string
	dim hasQuotes
	hasQuotes = false
	if instr(sTerm, Chr(34)) <> 0 then
		hasQuotes = true
	else
		'artificially put on quotes to see if we can get the whole phrase first
		sTerm = Chr(34) & sTerm & Chr(34)
	end if			
	'get the source
	strSQL = "SELECT * FROM " & tblpages & " "	
	if slowsearch = true then				
		if domGroup = "" then 'std search, no dom specified so group by domains
			strSQL = strSQL & " WHERE fContent LIKE '%" & sTerm & "%' GROUP BY fDomain ORDER BY fHits DESC LIMIT 100"
		else	'domain specified
			strSQL = strSQL & " WHERE fDomain='" & domGroup & "' AND fContent LIKE '%" & sTerm & "%' ORDER BY fHits DESC LIMIT 100"
		end if		
	elseif slowsearch = false then
		if domGroup = "" then 'std search, no dom specified so group by domains
			strSQL = strSQL & " WHERE match (fContent) against ('" & sTerm & "' IN BOOLEAN MODE ) GROUP BY fDomain ORDER BY fHits DESC LIMIT 100"
		else	'domain specified
			strSQL = strSQL & " WHERE fDomain='" & domGroup & "' AND match (fContent) against ('" & sTerm & "' IN BOOLEAN MODE ) ORDER BY fHits DESC LIMIT 100"
		end if
	else
		'no records
		strSQL = "SELECT fIndex FROM tblpages WHERE fIndex = 0"
	end if
	
	'Try a search
	Set rstResults = Server.CreateObject("ADODB.Recordset")
	rstResults.ActiveConnection = MM_conn_STRING
	rstResults.Source = strSQL
	rstResults.CursorType = 0
	rstResults.CursorLocation = 2
	rstResults.LockType = 1
	rstResults.Open()
	rstResults_numRows = 0
	
	'if there's no result and we artificially added quotes, take them off and try again
	if rstResults.EOF AND hasQuotes = false then
		rstResults.Close()
		'Set rstResults = Nothing
		sTerm = Replace(sTerm, Chr(34), "")
		'Set rstResults = Server.CreateObject("ADODB.Recordset")
		'get the source
		strSQL = "SELECT * FROM " & tblpages & " "	
		if slowsearch = true then
			if domGroup = "" then
				strSQL = strSQL & " WHERE fContent LIKE '%" & sTerm & "%' GROUP BY fDomain ORDER BY fHits DESC LIMIT 100"
			else
				strSQL = strSQL & " WHERE fDomain='" & domGroup & "' AND fContent LIKE '%" & sTerm & "%' ORDER BY fHits DESC LIMIT 100"
			end if
		elseif slowsearch = false then
			if domGroup = "" then
				strSQL = strSQL & " WHERE match (fContent) against ('" & sTerm & "' IN BOOLEAN MODE ) GROUP BY fDomain ORDER BY fHits DESC LIMIT 100"
			else
				strSQL = strSQL & " WHERE fDomain='" & domGroup & "' AND match (fContent) against ('" & sTerm & "' IN BOOLEAN MODE ) ORDER BY fHits DESC LIMIT 100"
			end if
		else
			'no records
			strSQL = "SELECT fIndex FROM tblpages WHERE fIndex = 0"
		end if
		rstResults.Source = strSQL
		rstResults.CursorType = 0
		rstResults.CursorLocation = 2
		rstResults.LockType = 1
		rstResults.Open()
		rstResults_numRows = 0
	end if	

	'get rid of the bits and pieces
	dim tempS
	tempS = sTerm
	tempS = Replace ( tempS, "*", "")
	tempS = Replace ( tempS, "+", "")
	tempS = Replace ( tempS, "-", "")
	tempS = Replace ( tempS, Chr(34), "" )

	'get an array of all the words in the search string	  
	'sTerm = Replace(sTerm, Chr(34), "")
	Dim sSearch, sSnippet, sMarkedString
	sSearch = tempS

	dim aWords
	aWords = Split(sSearch, " ")
	'create the marked string
	sMarkedString = ""
	for i = 0 to Ubound(aWords)
		sMarkedString = sMarkedString  & "<b>" & aWords(i) & "</b> "		
	next
	sMarkedString = trim(sMarkedString)
	
%>
          
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rstResults_numRows = rstResults_numRows + Repeat1__numRows
%>

<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rstResults_total
Dim rstResults_first
Dim rstResults_last

' set the record count
rstResults_total = rstResults.RecordCount

' set the number of rows displayed on this page
If (rstResults_numRows < 0) Then
  rstResults_numRows = rstResults_total
Elseif (rstResults_numRows = 0) Then
  rstResults_numRows = 1
End If

' set the first and last displayed record
rstResults_first = 1
rstResults_last  = rstResults_first + rstResults_numRows - 1

' if we have the correct record count, check the other stats
If (rstResults_total <> -1) Then
  If (rstResults_first > rstResults_total) Then
    rstResults_first = rstResults_total
  End If
  If (rstResults_last > rstResults_total) Then
    rstResults_last = rstResults_total
  End If
  If (rstResults_numRows > rstResults_total) Then
    rstResults_numRows = rstResults_total
  End If
End If
%>
          <%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rstResults_total = -1) Then

  ' count the total records by iterating through the recordset
  rstResults_total=0
  While (Not rstResults.EOF)
    rstResults_total = rstResults_total + 1
    rstResults.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rstResults.CursorType > 0) Then
    rstResults.MoveFirst
  Else
    rstResults.Requery
  End If

  ' set the number of rows displayed on this page
  If (rstResults_numRows < 0 Or rstResults_numRows > rstResults_total) Then
    rstResults_numRows = rstResults_total
  End If

  ' set the first and last displayed record
  rstResults_first = 1
  rstResults_last = rstResults_first + rstResults_numRows - 1
  
  If (rstResults_first > rstResults_total) Then
    rstResults_first = rstResults_total
  End If
  If (rstResults_last > rstResults_total) Then
    rstResults_last = rstResults_total
  End If

End If
%>
          <%
Dim MM_paramName 
%>
          <%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rstResults
MM_rsCount   = rstResults_total
MM_size      = rstResults_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
          <%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
          <%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
          <%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rstResults_first = MM_offset + 1
rstResults_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rstResults_first > MM_rsCount) Then
    rstResults_first = MM_rsCount
  End If
  If (rstResults_last > MM_rsCount) Then
    rstResults_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
          <%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
          <%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>

