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
<!--#include file="../include/functions.asp" -->
<%
	'update the hit counter
	strSQL = "UPDATE " & tblpages & " SET fHits = (fHits + 1) WHERE fIndex = " & Request.QueryString("fIndex") & ""
	runCmd (strSQL)

	Dim rstURL
	Dim rstURL_numRows
	Dim URLMovement
	Set rstURL = Server.CreateObject("ADODB.Recordset")
	rstURL.ActiveConnection = MM_conn_STRING
	rstURL.Source = "SELECT fPage  FROM " & tblpages & " WHERE fIndex = " & Request.QueryString("fIndex") & ""
	rstURL.CursorType = 0
	rstURL.CursorLocation = 2
	rstURL.LockType = 1
	rstURL.Open()
	rstURL_numRows = 0

	URLMovement = Replace(rstURL.Fields.Item("fPage").Value, "http://", "")
	URLMovement = "http://" & URLMovement

	rstURL.Close()
	Set rstURL = Nothing
	
	Response.Redirect(URLMovement)
%>
