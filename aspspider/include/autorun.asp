<%
'#################################################################################
'##
'## THE ASP SPIDER 
'##
'## this page v1.00.20
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
<!--#include file="../config/common.asp" -->
<!--#include file="functions.asp" -->
<%
	Dim rstRecords
	Set rstRecords = Server.CreateObject("ADODB.Recordset")
	rstRecords.ActiveConnection = MM_conn_STRING
	'this does a random selection on some days and a linear one on others
	if Day(Date) mod 2 = 1 Then
		rstRecords.Source = "SELECT * FROM " & tblp2s & " ORDER BY Rand() LIMIT " & autoSpider
	else
		rstRecords.Source = "SELECT * FROM " & tblp2s & " ORDER BY fIndex LIMIT " & autoSpider
	end if
	rstRecords.CursorType = 0
	rstRecords.CursorLocation = 2
	rstRecords.LockType = 1
	rstRecords.Open()
		
 	WHILE NOT rstRecords.EOF	
		SortThePage rstRecords("fPage"), 0
		delURLfromPageToSpider rstRecords("fIndex")
		rstRecords.MoveNext
	WEND	  

	rstRecords.Close
	set rstRecords = Nothing
%>
<%
'# DO NOT ALTER THIS SUB. This is essentially the SORTTHEPAGE sub from SPIDERCCODE.ASP with the
'# comments removed and the write to page removed
sub SortThePage (theURL, pCount)
	theURL = removeDefault ( theURL )
	pageContent = GetURLContent (theURL)
	if pageContent <> "" then	
		if InStr(lcase(getMetaTag ("robots", pageContent)), "nofollow") <> 0 then 
		else
			CollectLinks pageContent, theURL
		end if
		if InStr(lcase(getMetaTag ("robots", pageContent)), "noindex") <> 0 then 
		else
			dim sTitle, first, last
			first = InStr( lcase(pageContent), "<title>" ) + 7
			last = InStr( first, lcase(pageContent), "</title>" )
			on error resume next
			sTitle = Mid( pageContent, first, last-first )
			on error goto 0
			if sTitle = "" then sTitle = theURL
			pageContent = stripHTML(pageContent)
			pageContent = checkContent(pageContent)
			AddPageContent theURL, sTitle, pageContent
		end if
	end if
end sub
%>
