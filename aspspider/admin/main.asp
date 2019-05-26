<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../config/common.asp" -->
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
Dim rstDuePages
Dim rstDuePages_numRows

Set rstDuePages = Server.CreateObject("ADODB.Recordset")
rstDuePages.ActiveConnection = MM_conn_STRING
rstDuePages.Source = "SELECT * FROM " & tbldomadd & ""
rstDuePages.CursorType = 0
rstDuePages.CursorLocation = 2
rstDuePages.LockType = 1
rstDuePages.Open()

rstDuePages_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rstDuePages_numRows = rstDuePages_numRows + Repeat1__numRows
%>
<%
	Dim MSG
	MSG = ""
	select case Request("flag")
		case 1
			MSG = "<br><p align='center' class='txtMSG'>Search History has been Deleted.</p><hr><br>"
		case 2
			MSG = "<br><p align='center' class='txtMSG'>Search History has NOT been Deleted.</p><hr><br>"
		case 3
			MSG = "<br><p align='center' class='txtMSG'>The URL was added successfully to the Spider/Index queue.</p><hr><br>"
		case 4
			MSG = "<br><p align='center' class='txtMSG'>The URL was NOT added successfully to the Spider/Index queue.</p><hr><br>"
		case 5
			MSG = "<br><p class='txtMSG'>404 code was added successfully.</p><hr><br>"
		case 6
			MSG = "<br><p class='txtMSG'>There was no 404 content to add.</p><hr><br>"
		case 7
			MSG = "<br><p class='txtMSG'>No pages were deleted for 404 content.</p><hr><br>"
		case 8
			MSG = "<br><p class='txtMSG'>Pages were deleted for 404 content.</p><hr><br>"
		case 9
			MSG = "<br><p class='txtMSG'>Pages were deleted for URL content.</p><hr><br>"
		case 10
			MSG = "<br><p class='txtMSG'>Duplicate pages removed.</p><hr><br>"
		case 11
			MSG = "<br><p class='txtMSG'>No page index specified.</p><hr><br>"
	end select
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/aspspider_admin.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title><%= pageTitle %></title>
<!-- InstanceEndEditable --><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" -->



<link href="../config/style.css" rel="stylesheet" type="text/css">
<!-- InstanceEndEditable -->
<!--#include file="../config/meta.asp"-->
<link href="../config/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<table class="admin">
  <tr>
    <td colspan="2"><p align="center"><a href="../../default.asp"><img name="logo" src="../media/logo.gif" width="850" height="120" border="0" alt=""></a></p>
    <p align="center"><strong>Powered by the <a href="http://www.transworldinteractive.net" target="_blank">Transworld Interactive ASP Spider</a></strong> v <%= ver %></p></td>
  </tr>
  <tr valign="top">
    <td width="1%" nowrap><p><strong>Admin Menu</strong></p>
      <p><a href="main.asp">Admin Home</a></p>
      <p><a href="stats/statistics.asp">Statistics</a> </p>
      <p><a href="stats/searches.asp">All Searches</a> </p>
      <p><a href="spiderindex/deletesearches.asp">Delete Searches</a></p>
      <p><a href="default.asp">Log Out</a> </p>
      <hr>
      <p><a href="../../default.asp" target="_blank">Site Home</a> </p>
    </td>
    <td><!-- InstanceBeginEditable name="Body" -->
                <p><%= MSG %>
                </p>
                <table width="0%"  border="0" cellspacing="0" cellpadding="20">
                  <tr valign="top">
                    <td width="1%"><table width="25%"  border="0" cellpadding="20" cellspacing="0" class="searchTable">
                      <tr>
                        <td nowrap><p><strong>Spider/Index Pages</strong></p>
                            <form name="form2" method="post" action="spiderindex/_spidercode.asp">
                              <select name="fAmount" id="fAmount">
                                <option value="1">1</option>
                                <option value="10">10</option>
                                <option value="50">50</option>
                                <option value="100">100</option>
                                <option value="200">200</option>
                                <option value="500">500</option>
                              </select>
                              <input name="Submit" type="submit" class="bodyTextSmall" value="Submit">
                            </form>
                            </td>
                      </tr>
                    </table></td>
                    <td><table border="0" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF" class="searchTable">
                      <tr>
                        <td nowrap><form name="form1" method="post" action="spiderindex/_addurl.asp">
                            <p><strong>Add URL due for Spidering </strong></p>
                            <p><span class="bodyTextMedium">Site URL </span>
                                <input name="fURL" type="text" id="fURL" value="http://" size="60">
                                <input name="Submit" type="submit" class="bodyTextSmall" value="Submit">
                            </p>
                        </form></td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td><table width="25%"  border="0" cellpadding="20" cellspacing="0" class="searchTable">
                      <tr>
                        <td nowrap><p><strong>Edit Page</strong></p>
                            <form name="form2" method="post" action="spiderindex/editpage.asp">
                              <input name="fIndex" type="text" id="fIndex" size="5" maxlength="10">
&nbsp;                              
<input name="Submit" type="submit" class="bodyTextSmall" value="Submit">
                          </form></td>
                      </tr>
                    </table></td>
                    <td><table width="100%" border="0" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF" class="searchTable">
                      <tr>
                        <td valign="top"><p align="left"><strong>Pages due for spidering: </strong></p>
                            <p align="left" class="bodyTextMedium">click on VIEW URL to check the domain is valid for this search engine</p>
                            <p align="left" class="bodyTextMedium">if the domain is not to be added to the engine, then click DELETE</p>
                            <p align="left" class="bodyTextMedium">to spider &amp; index the first page of the domain and prepare the domain for spidering/indexing click on the DOMAIN NAME </p>
                            <hr align="left">
                            <% 
While ((Repeat1__numRows <> 0) AND (NOT rstDuePages.EOF)) 
%>
                            <p align="left"><a href="spiderindex/_spidercode.asp?theURL=<%=(rstDuePages.Fields.Item("fURL").Value)%>"><%=(rstDuePages.Fields.Item("fURL").Value)%></a> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- &nbsp;&nbsp;&nbsp;&nbsp;<a href="spiderindex/_deleteurl.asp?fURL=<%=(rstDuePages.Fields.Item("fURL").Value)%>">delete</a> &nbsp;&nbsp;&nbsp;&nbsp;- &nbsp;&nbsp;&nbsp;&nbsp;<a href="http://<%=(rstDuePages.Fields.Item("fURL").Value)%>" target="_blank">view URL</a> </p>
                            <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rstDuePages.MoveNext()
Wend
%></td>
                      </tr>
                    </table></td>
                  </tr>
                </table>
                <p>&nbsp;</p>
                <table border="0" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF" class="searchTable">
                  <tr>
                    <td nowrap><form name="form1" method="post" action="spiderindex/_404code.asp">
                        <p><strong>Bad Page Content (404) </strong></p>
                        <p><span class="bodyTextMediumFaded">Delete pages with the following in their content </span></p>
                        <p>
                          <input name="fBadContent" type="text" id="fBadContent" size="60">
                          <input name="Submit" type="submit" class="bodyTextSmall" value="Submit">
                        </p>
                    </form>
                      <p><span class="bodyTextMediumFaded">This next operation can take some time depending on the amount of 404 content and the number of pages.</span> </p>
                      <p><a href="spiderindex/badcontent.asp">Clean up Bad Content!</a> </p></td>
                  </tr>
                </table>
                <p>&nbsp;</p>
                <table border="0" cellpadding="10" cellspacing="0" bgcolor="#FFFFFF" class="searchTable">
                  <tr>
                    <td nowrap><form name="form1" method="post" action="spiderindex/_deletebadurl.asp">
                        <p><strong>Bad URL</strong></p>
                        <p><span class="bodyTextMediumFaded">Delete pages with the following in their url </span></p>
                        <p>
                          <input name="fBadURL" type="text" id="fBadURL" size="60">
                          <input name="Submit" type="submit" class="bodyTextSmall" value="Submit">
                        </p>
                      </form>
                      </td>
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
rstDuePages.Close()
Set rstDuePages = Nothing
%>
