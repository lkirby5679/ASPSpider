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
	'update the info
	strSQL = "UPDATE " & tblpages & " SET"
	strSQL =  strSQL & " fDomain = '" & Replace(Request("fDomain"), "'", "''") & "',"
	strSQL =  strSQL & " fPage = '" & Replace(Request("fPage"), "'", "''") & "',"
	strSQL =  strSQL & " fTitle = '" & Replace(Request("fTitle"), "'", "''") & "',"
	strSQL =  strSQL & " fContent = '" & Replace(Request("fContent"), "'", "''") & "',"
	strSQL =  strSQL & " fHits = '" & Replace(Request("fHits"), "'", "''") & "'"
	strSQL =  strSQL & " WHERE fIndex = " & Request("fIndex") & ""

	dim Command1
	set Command1 = Server.CreateObject("ADODB.Command")
	Command1.ActiveConnection = MM_conn_STRING
	Command1.CommandText = strSQL
	Command1.CommandType = 1
	Command1.CommandTimeout = 0
	Command1.Prepared = true
	Command1.Execute()
	set Command1 = nothing
	
	Response.Redirect("editpage.asp?fIndex=" & Request("fIndex"))
%>

