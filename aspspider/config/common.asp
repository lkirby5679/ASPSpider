<% OPTION EXPLICIT %>
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
	'This file contains the variables and settings for the site. 

	'####### DO NOT ALTER ANY OF THE CODE IN THIS SECTION!!
	dim online, pword, pkey, strSQL
	dim db_server, db_name, db_uid, db_pass 
	dim MM_conn_STRING, ver
	dim sitename, domain, singleDom
	dim pagetitle
	dim emadd
	dim iCharsBefore, iCharsExample	
	dim ex1, ex2, ex3
	dim showTopAd, showSideAd
	dim txtWelcome, txtSearch
	dim maxReturns, grpReturns
	dim add1, autoSpider
	dim tbl404, tbldomadd, tblpages, tblp2s, tblsterm	
	dim daysToSpider, detailSpider
	dim sboxcolour
	'####### END OF SECTION
	

	'####### YOU WILL NEED TO SET THE VARIABLES IN THE SECTION BELOW TO REFLECT YOUR OWN SITE	
	'# you can specify two sets of connection details to allow for a local copy (for testing) and a
	'# remote copy for the final version; to switch between the two turn online into true or false
	online		 	= 	true
	if online = true then
		'# details for the online copy of the db
		db_server 	= 	"vr1mysqlserver"				'the MYSQL server name
		db_name 	= 	"aspspider"				'the MYSQL database name
		db_uid 		= 	"root"				'the username for the MYSQL database
		db_pass 	=	"227oct36"			'the password for the MYSQL database
		pword		=	"227oct36"					'the password used to get into the admin site
	else
		'# details for the local copy of the db
		db_server 	= 	"localhost"				'the MYSQL server name
		db_name 	= 	"aspspider"				'the MYSQL database name
		db_uid 		= 	"aspspider"					'the username for the MYSQL database
		db_pass 	=	"227oct36"						'the password for the MYSQL database	
		pword		=	"227oct36"				'the password used to get into the admin site
	end if

	'# site variables
	sitename 	=	"Transworld Interactive Search"						'the name/title of the site
	domain 		= 	"www.searchthree.net"				'the site domain name
	pagetitle 	= 	"Transworld Interactive Search Engine"		'the page title (appears in the browser)
	emadd		=	"USE THE FORUMS FOR CONTACT"			'email address where people can contact the site 
	sboxcolour	=	"#99CCFF"								'the colour of the addsearchbox box
	
	'a 2 word phrase which is used in the examples for different kinds of searches
	ex1			=	"asp"
	ex2			=	"code"
	'a string demonstraing variations of a * wildcard at the end of a word; this must be related to ex1 above
	ex3			=	"asp, aspic, aspen"
	
	'various text strings
	txtWelcome	= 	"Welcome to Transworld Interactive Search Engine!"
	txtSearch	=	"Search for an word or phrase."
	add1		=	"All sites will be considered for inclusion in this database."
	'####### END OF SECTION
	
	
	'####### DO NOT ALTER THE VARIABLES IN THIS SECTION UNLESS YOU KNOW WHAT YOU ARE DOING!!!
	iCharsBefore 	= 	200		
	iCharsExample 	= 	400
	showTopAd		=	false
	showSideAd		=	true
	maxReturns		=	100
	grpReturns		=	10
	pkey			=	"asp2006"
	autoSpider		=	1
	tbl404			=	"badcontent"
	tbldomadd		=	"domtoadd"
	tblpages		=	"pages"
	tblp2s			=	"pagetospider"
	tblsterm		=	"searchterm"	
	daysToSpider	=	365
	detailSpider	=	false
	singleDom		=	domain
	'####### END OF SECTION
	
	
	'####### DO NOT ALTER ANY OF THE CODE IN THIS SECTION!!
	MM_conn_STRING 	= 	"Driver={MySQL ODBC 3.51 Driver};Server=" & db_server & ";Port=3306;Option=131072;Stmt=;Database=" & db_name & ";Uid=" & db_uid & ";Pwd=" & db_pass & ";"
	ver 			= 	"1.10.00"
	'####### END OF SECTION
%>