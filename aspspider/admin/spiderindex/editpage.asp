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
Response.Expires=-2500
Response.CacheControl="no-cache"
if Session("pkey") <> pkey then Response.Redirect("default.asp") 
%>
<%
	dim idx
	idx = trim(Request("fIndex"))
	if idx = "" then Response.Redirect("main.asp?flag=11")
	
	strSQL = "SELECT * FROM " & tblpages & " WHERE fIndex = " & idx

	Dim rstRecords
	Set rstRecords = Server.CreateObject("ADODB.Recordset")
	rstRecords.ActiveConnection = MM_conn_STRING
	rstRecords.Source = strSQL
	rstRecords.CursorType = 0
	rstRecords.CursorLocation = 2
	rstRecords.LockType = 1
	rstRecords.Open() 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/aspspider_admin.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title><%= pagetitle %></title>
<!-- InstanceEndEditable --><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
<!--#include file="../../config/meta.asp"-->
<link href="../../config/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<table class="admin">
  <tr>
    <td colspan="2"><p align="center"><a href="../../../default.asp"><img name="logo" src="../../media/logo.gif" width="850" height="120" border="0" alt=""></a></p>
    <p align="center"><strong>Powered by the <a href="http://www.transworldinteractive.net" target="_blank">Transworld Interactive ASP Spider</a></strong> v <%= ver %></p></td>
  </tr>
  <tr valign="top">
    <td width="1%" nowrap><p><strong>Admin Menu</strong></p>
      <p><a href="../main.asp">Admin Home</a></p>
      <p><a href="../stats/statistics.asp">Statistics</a> </p>
      <p><a href="../stats/searches.asp">All Searches</a> </p>
      <p><a href="deletesearches.asp">Delete Searches</a></p>
      <p><a href="../default.asp">Log Out</a> </p>
      <hr>
      <p><a href="../../../default.asp" target="_blank">Site Home</a> </p>
    </td>
    <td><!-- InstanceBeginEditable name="Body" -->

<p align="center">&nbsp;</p>
<form name="form1" method="post" action="_editpage.asp">
<table width="0%"  border="0" cellpadding="0" cellspacing="0" class="admin">
  <tr class="txtSmall">
    <td valign="top">Index</td>
    <td colspan="2" valign="top"><%= rstRecords("fIndex") %>
      <input name="fIndex" type="hidden" id="fIndex" value="<%= rstRecords("fIndex") %>"></td>
  </tr>
  <tr class="txtSmall">
    <td valign="top">Domain</td>
    <td colspan="2" valign="top">
        <input name="fDomain" type="text" id="fDomain" value="<%= rstRecords("fDomain") %>" size="70">
      </td>
  </tr>
  <tr class="txtSmall">
    <td valign="top">Page - <a href="<%= rstRecords("fPage") %>" target="_blank">visit</a></td>
    <td colspan="2" valign="top">
      <input name="fPage" type="text" id="fPage" value="<%= rstRecords("fPage") %>" size="70"></td>
  </tr>
  <tr class="txtSmall">
    <td valign="top">Title</td>
    <td colspan="2" valign="top">
      <input name="fTitle" type="text" id="fTitle" value="<%= rstRecords("fTitle") %>" size="70"></td>
  </tr>
  <tr class="txtSmall">
    <td valign="top">Contents</td>
    <td colspan="2" valign="top">
      <textarea name="fContent" cols="70" rows="30" wrap="VIRTUAL" id="fContent"><%= rstRecords("fContent") %></textarea></td>
  </tr>
  <tr class="txtSmall">
    <td valign="top">Total Hits</td>
    <td valign="top">
      <input name="fHits" type="text" id="fHits" value="<%= rstRecords("fHits") %>" size="20"></td>
    <td valign="top"><input type="submit" name="Submit" value="Submit"></td>
  </tr>
</table>
</form>
<p align="center"><em></em></p>
    <!-- InstanceEndEditable --></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
<!-- InstanceEnd --></html>
<%
	rstRecords.Close()
	Set rstRecords = Nothing
%>


