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
	'this sets up the tables and db as required
	strSql = "CREATE TABLE " & tbldomadd & " ( "
	strSql = strSql & "fURL varchar (255) NOT NULL default '', "	
	strSql = strSql & "PRIMARY KEY (fURL)"
	strSql = strSql & ") TYPE=MyISAM;"
	runCmd (strSQL)	

	strSql = "CREATE TABLE " & tbl404 & "( "
	strSql = strSql & "fIndex int(10) unsigned NOT NULL auto_increment,"
	strSql = strSql & "fBadContent varchar(225) NOT NULL default '',"
	strSql = strSql & "PRIMARY KEY (fIndex)"
	strSql = strSql & ") TYPE=MyISAM;"
	runCmd (strSQL)	

	strSql = "CREATE TABLE " & tblpages & " ( "
	strSql = strSql & "fIndex bigint(20) unsigned NOT NULL auto_increment,"
	strSql = strSql & "fDomain varchar(100) NOT NULL default '',"
	strSql = strSql & "fPage varchar(255) NOT NULL default '',"
	strSql = strSql & "fTitle varchar(100) NOT NULL default '',"
	strSql = strSql & "fContent text NOT NULL,"
	strSql = strSql & "fHits int(11) NOT NULL default '0',"
	strSql = strSql & "fIndexDate date NOT NULL default '0000-00-00',"
	strSql = strSql & "PRIMARY KEY  (fIndex),"
	strSql = strSql & "UNIQUE KEY page (fPage),"
	strSql = strSql & "FULLTEXT KEY content (fContent)"
	strSql = strSql & ") TYPE=MyISAM;"
	runCmd (strSQL)	
	
	strSql = "CREATE TABLE " & tblp2s & " ("
	strSql = strSql & "fIndex bigint(20) unsigned NOT NULL auto_increment,"
	strSql = strSql & "fPage varchar(100) NOT NULL default '',"
	strSql = strSql & "PRIMARY KEY  (fIndex),"
	strSql = strSql & "UNIQUE KEY page (fPage)"
	strSql = strSql & ") TYPE=MyISAM;"
	runCmd (strSQL)	

	strSql = "CREATE TABLE " & tblsterm & " ("
	strSql = strSql & "fIndex int(10) unsigned NOT NULL auto_increment,"
	strSql = strSql & "fSearchTerm varchar(50) NOT NULL default '',"
	strSql = strSql & "fCount int(10) unsigned NOT NULL default '0',"
	strSql = strSql & "PRIMARY KEY  (fIndex),"
	strSql = strSql & "UNIQUE KEY fSearchTerm (fSearchTerm)"
	strSql = strSql & ") TYPE=MyISAM;"
	runCmd (strSQL)
%>
The database has been set up. You should now delete this file -- admin/setup/setup.asp -- from the server. 
