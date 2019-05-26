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
Dim rstSearches
Dim rstSearches_numRows

Set rstSearches = Server.CreateObject("ADODB.Recordset")
rstSearches.ActiveConnection = MM_conn_STRING
rstSearches.Source = "SELECT SUM(fCount) AS SUMCOUNT  FROM " & tblsterm & ""
rstSearches.CursorType = 0
rstSearches.CursorLocation = 2
rstSearches.LockType = 1
rstSearches.Open()

rstSearches_numRows = 0
%>
<%
Dim rstNumSpidered
Dim rstNumSpidered_numRows

Set rstNumSpidered = Server.CreateObject("ADODB.Recordset")
rstNumSpidered.ActiveConnection = MM_conn_STRING
rstNumSpidered.Source = "SELECT COUNT(fIndex) as NumSpidered FROM " & tblpages & ""
rstNumSpidered.CursorType = 0
rstNumSpidered.CursorLocation = 2
rstNumSpidered.LockType = 1
rstNumSpidered.Open()

rstNumSpidered_numRows = 0
%>

<%
Dim rstNumIndexed
Dim rstNumIndexed_numRows

Set rstNumIndexed = Server.CreateObject("ADODB.Recordset")
rstNumIndexed.ActiveConnection = MM_conn_STRING
rstNumIndexed.Source = "SELECT COUNT(fIndex) as NumIndexed FROM " & tblp2s & ""
rstNumIndexed.CursorType = 0
rstNumIndexed.CursorLocation = 2
rstNumIndexed.LockType = 1
rstNumIndexed.Open()

rstNumIndexed_numRows = 0
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
      <p><a href="statistics.asp">Statistics</a> </p>
      <p><a href="searches.asp">All Searches</a> </p>
      <p><a href="../spiderindex/deletesearches.asp">Delete Searches</a></p>
      <p><a href="../default.asp">Log Out</a> </p>
      <hr>
      <p><a href="../../../default.asp" target="_blank">Site Home</a> </p>
    </td>
    <td><!-- InstanceBeginEditable name="Body" -->

                <table width="0%"  border="0" cellpadding="10" cellspacing="0" class="admin">
                  <tr bgcolor="#E6E6E6">
                    <td align="left" nowrap>searches made so far </td>
                    <td align="right">
						<% If rstSearches("SUMCOUNT") < 1 OR ISNULL(rstSearches("SUMCOUNT")) then
							Response.Write("0")
						Else
							Response.Write(FormatNumber(rstSearches("SUMCOUNT"), 0, 0, 0, TRUE))
						End If %>
					</td>
                  </tr>
                  <tr bgcolor="#F5F5F5">
                    <td align="left" nowrap>pages indexed </td>
                    <td align="right"><%= FormatNumber(rstNumSpidered("NumSpidered"), 0, 0, 0, TRUE) %></td>
                  </tr>
                  <tr bgcolor="#E6E6E6">
                    <td align="left" nowrap>pages queued for indexing </td>
                    <td align="right"><%= FormatNumber(rstNumIndexed("NumIndexed"), 0, 0, 0, TRUE) %></td>
                  </tr>
                  <tr>
                    <td align="left" nowrap>&nbsp;</td>
                    <td align="right">&nbsp;</td>
                  </tr>
                  <tr>
                    <td align="left" nowrap>&nbsp;</td>
                    <td align="right">&nbsp;</td>
                  </tr>
                  <tr>
                    <td align="left" nowrap>&nbsp;</td>
                    <td align="right">&nbsp;</td>
                  </tr>
                  <tr>
                    <td align="left" nowrap>&nbsp;</td>
                    <td align="right">&nbsp;</td>
                  </tr>
                  <tr>
                    <td align="left" nowrap>&nbsp;</td>
                    <td align="right">&nbsp;</td>
                  </tr>
                  <tr>
                    <td align="left" nowrap>&nbsp;</td>
                    <td align="right">&nbsp;</td>
                  </tr>
                  <tr>
                    <td align="left" nowrap>&nbsp;</td>
                    <td align="right">&nbsp;</td>
                  </tr>
      </table>
                <!-- InstanceEndEditable --></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
<!-- InstanceEnd --></html>
<%
rstSearches.Close()
Set rstSearches = Nothing
%>
<%
rstNumSpidered.Close()
Set rstNumSpidered = Nothing
%>

