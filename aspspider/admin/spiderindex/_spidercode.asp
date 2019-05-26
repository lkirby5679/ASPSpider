<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../config/common.asp" -->
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
Response.Buffer = False
Server.ScriptTimeout = 2500	'5 minutes
%>
<!--#include file="../../include/functions.asp" -->
<%
Response.Expires=-2500
Response.CacheControl="no-cache"
if Session("pkey") <> pkey then Response.Redirect("../default.asp") 
%>
<style type="text/css">
<!--
.disabledTextBox {
	border: thin none #FFFFFF;
	color: #FFFFFF;
	background-color: #FF0000;
}-->
</style>

<div id="ProgBar"><TABLE width="800" Border=1><TR><TD><TABLE HEIGHT="16" Border=0><TR><TD BGCOLOR=RED ID=statuspic><input width="3" name="pc" type="text" class="disabledTextBox" id="pc" value="0%"></TD></TR></TABLE></table><BR></div>
<div id="outWrite"><table border="1" width=800 id="tblOutput"></table></div>
<script language="Javascript">var progBarWidth=800;</script>
<script language="Javascript"> 
function WriteLayer(ID,parentID,sText) { 
 if (document.layers) { 
   var oLayer; 
   if(parentID){ 
     oLayer = eval('document.' + parentID + '.document.' + ID + '.document'); 
   }else{ 
     oLayer = document.layers[ID].document; 
   } 
   oLayer.open(); 
   sText = oLayer.value + sText;
   oLayer.write(sText); 
   oLayer.close(); 
 } 
 else if (parseInt(navigator.appVersion)>=5&&navigator. 
appName=="Netscape") { 
	sText = document.getElementById(ID).innerHTML + sText;
   document.getElementById(ID).innerHTML = sText; 
 } 
 else if (document.all) {
 	sText = document.all[ID].innerHTML + sText;
	document.all[ID].innerHTML = sText;
	}
} 
</script>
<%
Sub ShowProgress(nPctComplete)
	Response.Write "<SCR" & "IPT LANGUAGE=""JavaScript"">" & vbCrlf
	Response.Write "pc.value = (" & nPctComplete & " * 100) + ""%"";" & vbCrlf
	Response.Write "statuspic.width = Math.ceil(" & nPctComplete & " * progBarWidth);" & vbCrlf
	Response.Write "</SCR" & "IPT>"
End Sub
Sub AddRow (s1, s2, colour)
	if detailSpider = false then exit sub
	Response.Write "<SCR" & "IPT LANGUAGE=""JavaScript"">" & vbCrlf
	Response.Write "addRowToTable('" & s1 & "','" & s2 & "','" & colour & "');" & vbCrLf
	Response.Write "</SCR" & "IPT>"
end sub
%>
<script language="Javascript"> 
function addRowToTable(s1, s2, colour)
{
  	var tbl = document.getElementById('tblOutput');
  	var lastRow = tbl.rows.length;
  	var iteration = lastRow;
  	var row = tbl.insertRow(lastRow);
  	var cellLeft = row.insertCell(0);
  	var textNode = document.createTextNode(s1);
  	cellLeft.appendChild(textNode);
  	cellLeft.style.background = colour;
  	var cellRight = row.insertCell(1);
  	var textNode = document.createTextNode(s2);
  	cellRight.appendChild(textNode);
	cellRight.style.background = colour;
}
</script>
<%
	dim pageContent, pCount
	dim nProcessedSoFar, nTotalRecords, pctComplete
	'if this is just a single URL from the dom to add then skip collecting records and go straight to 
	'the processing of that single URL
	if Request("theURL") <> "" then
		StartTimer 1
		pCount = 1
		StartTimer 2
		SortThePage Request("theURL"), pCount
		delURLfromSpider (Request("theURL"))
		pCount = 2
		nProcessedSoFar = 0
		nTotalRecords = 1
	    nProcessedSoFar = nProcessedSoFar + 1
    	pctComplete = (nProcessedSoFar / nTotalRecords)
		ShowProgress pctComplete
	else
		Dim rstRecords
		Dim nToProcess 
		nToProcess = Request("fAmount")
		Set rstRecords = Server.CreateObject("ADODB.Recordset")
		rstRecords.ActiveConnection = MM_conn_STRING
		'this does a random selection on some days and a linear one on others
		if Day(Date) mod 2 = 1 Then
			rstRecords.Source = "SELECT * FROM " & tblp2s & " ORDER BY Rand() LIMIT " & nToProcess
		else
			rstRecords.Source = "SELECT * FROM " & tblp2s & " ORDER BY fIndex LIMIT " & nToProcess
		end if
		rstRecords.CursorType = 0
		rstRecords.CursorLocation = 2
		rstRecords.LockType = 1
		rstRecords.Open()
		'get the content of the URL
		StartTimer 1
		pCount = 1
		nTotalRecords = nToProcess
		WHILE NOT rstRecords.EOF	
			StartTimer 2
			SortThePage rstRecords("fPage"), pCount
			pCount = pCount + 1
			delURLfromPageToSpider rstRecords("fIndex")
			rstRecords.MoveNext
			nProcessedSoFar = nProcessedSoFar + 1
			pctComplete = (nProcessedSoFar / nTotalRecords)
			ShowProgress pctComplete
		WEND	  
		rstRecords.Close
		set rstRecords = Nothing
	end if
	if detailSpider = false then detailSpider = true
	AddRow "Processing time for " & pCount - 1 & " pages:", StopTimer(1), "#ffff00"
	AddRow "Status:","All Done!","#ffff00"
	Response.Write("<p><a href=../main.asp>Return to Admin Home</a></p>")
%>
<% 
sub SortThePage (theURL, pCount)
	'some output
	AddRow pCount & " - URL to spider:", theURL, "#ccff00"
	if theURL = "" then exit sub
	
	'see if the URL has been spidered within the alloted time, if so bomb out
	strSQL = "SELECT fIndexDate FROM " & tblpages & " WHERE fPage = '" & Replace(theURL, "'", "") & "'"
	Dim rstRecords
	Set rstRecords = Server.CreateObject("ADODB.Recordset")
	rstRecords.ActiveConnection = MM_conn_STRING
	rstRecords.Source = strSQL
	rstRecords.CursorType = 0
	rstRecords.CursorLocation = 2
	rstRecords.LockType = 1
	rstRecords.Open()	
	if not rstRecords.EOF then		
		AddRow "Spider Date:",rstRecords("fIndexDate"), "#ffffff"
		if  DateDiff("d", rstRecords("fIndexDate"), Date()) < daysToSpider then
			theURL = ""
			AddRow "Spider Date:","NOT SPIDERED", "#ffffff"
		end if
	else
		AddRow "Spider Date:","Not previously spidered", "#ffffff"
	end if
	rstRecords.Close
	set rstRecords = Nothing		
	if theURL = "" then exit sub
	

	theURL = removeDefault ( theURL )
	pageContent = GetURLContent (theURL)
	'only work if there's something to do
	if pageContent <> "" then	
		'see if we can get the links from this page
		if InStr(lcase(getMetaTag ("robots", pageContent)), "nofollow") <> 0 then 
			AddRow "ROBOT Meta Tag:","NOFOLLOW set - page not spidered", "#ffffff"
		elseif InStr(lcase(getMetaTag ("robots", pageContent)), "no follow") <> 0 then 
			AddRow "ROBOT Meta Tag:","NOFOLLOW set - page not spidered", "#ffffff"
		else
			CollectLinks pageContent, theURL
		end if
		'can we index this page?
		if InStr(lcase(getMetaTag ("robots", pageContent)), "noindex") <> 0 then 
			AddRow "ROBOT Meta Tag:","NOINDEX set - page not indexed", "#ffffff"
		elseif InStr(lcase(getMetaTag ("robots", pageContent)), "no index") <> 0 then 
			AddRow "ROBOT Meta Tag:","NOFOLLOW set - page not spidered", "#ffffff"
		else
			'get the title of the page
			dim sTitle, first, last
			first = InStr( lcase(pageContent), "<title>" ) + 7
			last = InStr( first, lcase(pageContent), "</title>" )
			on error resume next
			sTitle = trim(Mid( pageContent, first, last-first ))
			on error goto 0
			if sTitle = "" then sTitle = theURL
			'now remove the HTML
			pageContent = stripHTML(pageContent)
			'make sure the content is ok
			pageContent = checkContent(pageContent)
			'add to db
			AddPageContent theURL, sTitle, pageContent
			'display the results
			AddRow "Page title:",sTitle,"#ffffff"
			'AddRow "Page content:",pageContent,"#ffffff"			
		end if
	end if
	AddRow "Time to process URL:",StopTimer(2), "#ffffff"
end sub
%>
