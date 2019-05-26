<%@LANGUAGE="VBSCRIPT"%>
<!--#INCLUDE file="../config/common.asp" -->
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
<!--#INCLUDE file="../include/functions.asp" -->
<%
	dim sURL
	sURL = trim (Request.form("fURL"))
	'this does some simple filtering to avoid bad URLs being presented
	if instr(sURL, "Content-Type:") <> 0 then Response.Redirect("../default.asp")
	if instr(sURL, "MIME-Version:") <> 0 then Response.Redirect("../default.asp")
	if instr(sURL, "multipart/mixed:") <> 0 then Response.Redirect("../default.asp")
	if instr(sURL, "@") <> 0 then Response.Redirect("../default.asp")
	dim ret
	ret = AddURLtoSpider (sURL)
	Response.Redirect("../addurl.asp?flag=" & ret)
%>

