<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="aspspider/config/common.asp" -->
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
	showSideAd = false

	Dim MSG
	MSG = ""
	if Request("flag") = 1 then
		MSG = "<p align='center' class='txtMSG'>No Search String specified.</p>"
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/aspspider_main.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title><%= pageTitle %></title>
<!-- InstanceEndEditable --><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style1 {font-style: italic}
-->
</style>
<!-- InstanceEndEditable -->
<!--#include file="aspspider/config/meta.asp"-->
<link href="aspspider/config/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<table>
  <tr>
    <td><a href="default.asp"><img name="logo" src="aspspider/media/logo.gif" width="850" height="120" border="0" alt=""></a></td>
  </tr>
  <% If showTopAd = true Then %>
  <tr>
    <td align="center">            		
        <!--#include file="aspspider/include/inc_ad_top.asp"-->        
	</td>
  </tr>
  <% End If %>
  <tr>
    <td align="center"><table class="middle">
      <tr valign="top">
        <td><!-- InstanceBeginEditable name="Body" -->	 
			  
            <table class="mainsearch">
              <tr>
                <td class="txtWelcome">
                  <%= txtWelcome %>
                </td>
              </tr>
              <tr class="trWhite">
                <td><form name="form1" method="post" action="results.asp">
                    <p class="txtInstruct"><%= txtSearch %></p>
                    <p><%= MSG %></p>
                    <table class="searchbox">
                      <tr>
                        <td><input name="fSearch" type="text" id="fSearch2" size="60"></td>
                      </tr>
                    </table>
                    <p>
                      <input type="submit" name="Submit" value="Search">
                    </p>
                </form></td>
              </tr>
              <tr>
                <td class="txtNB">NB: a search will return the top <%= maxReturns %> results only.</td>
              </tr>
            </table>
            <!-- InstanceEndEditable --></td>
		<% If showSideAd = true Then %>
        <td>
           
            <!--#include file="aspspider/include/inc_ad_side.asp"-->
           
        </td>
		<% End If %>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td align="center">
	<p> To place your ad here, contact Transworld Interactive at sales@transworldinteractive.net
</p>
</td>
  </tr>
  <tr>
    <td align="center"><table class="footer">
              <tr>
                <td><div align="center"><a href="default.asp"> Home</a>&nbsp;&nbsp;&middot;&nbsp;&nbsp;<a href="aspspider/addurl.asp"><strong>Add your URL</strong></a>&nbsp;&nbsp;&middot;&nbsp;&nbsp;<a href="aspspider/addsearchbox.asp">Add <%= sitename %> toYour Site!</a>&nbsp;&nbsp;&middot;&nbsp;<a href="mailto:<%= emadd %>">Contact</a>
  </div>
                  <div align="center" class="txtPoweredby">Powered by the <a href="http://www.transworldinteractive.net" target="_blank">Transworld Interactive ASP Spider</a> v <%= ver %></div></td>
              </tr>
            </table>
			
	</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>