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
<!--#include file="../../include/functions.asp" -->
<%
Response.Expires=-2500
Response.CacheControl="no-cache"
if Session("pkey") <> pkey then Response.Redirect("../default.asp") 
%>
<%
	strSQL = "SELECT * FROM " & tbl404 & " ORDER BY fBadContent"

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
<title><%= pageTitle %></title>
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
<p align="center">Are you sure you want to delete all the pages containing the following strings?</p>
<table width="0%"  border="0" cellspacing="0" cellpadding="10">
<%  
 	dim i
	i = 0
	WHILE NOT rstRecords.EOF
	i = i + 1
	
%>
	<tr>
		<td><%= i %>:&nbsp;<%= rstRecords("fBadContent") %></td>
		<td><a href="_deletebadcontent.asp?fBC=<%= rstRecords("fBadContent") %>"><span class="bodyTextSmall">delete this string</span></a></td>
	</tr>
<%
	rstRecords.MoveNext
	WEND	  
%>
</table>


<p align="center"><a href="../main.asp?flag=7">no</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="_deletebadcontent.asp">yes</a></p>
<!-- InstanceEndEditable --></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
<!-- InstanceEnd --></html>
<%
	rstRecords.Close
	set rstRecords = Nothing
%>