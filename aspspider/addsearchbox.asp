<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="config/common.asp" -->
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/aspspider_main.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title><%= pagetitle %></title>
<!-- InstanceEndEditable --><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
<!--#include file="config/meta.asp"-->
<link href="config/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<table>
  <tr>
    <td><a href="../default.asp"><img name="logo" src="media/logo.gif" width="850" height="120" border="0" alt=""></a></td>
  </tr>
  <% If showTopAd = true Then %>
  <tr>
    <td align="center">            		
        <!--#include file="include/inc_ad_top.asp"-->        
	</td>
  </tr>
  <% End If %>
  <tr>
    <td align="center"><table class="middle">
      <tr valign="top">
        <td><!-- InstanceBeginEditable name="Body" -->
            <table>
              <tr>
                <td align="left"><p align="center" class="txtTitle">Add <%= sitename %> to your site!</p>
                  <p>If you would like to add the powerful search facilities of <%= sitename %> to your site, you can do so absolutely free.</p>
                  <p>Just copy and paste the following code into your site HTML and the search is ready. Results open in a new window.</p>
                  <table border="1">
                    <tr>
                      <td align="left" class="bodyTextMedium">&lt;table border=&quot;1&quot; align=&quot;center&quot; cellpadding=&quot;10&quot; cellspacing=&quot;0&quot; bgcolor=&quot;<%= sboxcolour %>&quot;&gt;&lt;tr&gt;&lt;td&gt;&lt;div align=&quot;center&quot;&gt;&lt;p&gt;&lt;a href=&quot;http://<%= domain %>&quot; target=&quot;_blank&quot;&gt;&lt;img src=&quot;http://<%= domain %>/media/linklogo.gif&quot; width=&quot;66&quot; height=&quot;60&quot; hspace=&quot;3&quot; vspace=&quot;0&quot; border=&quot;0&quot; align=&quot;absmiddle&quot;&gt;&lt;/a&gt;&amp;nbsp;&amp;nbsp;<%= txtSearch %>&lt;/p&gt;&lt;form action=&quot;http://<%= domain %>/results.asp&quot; method=&quot;post&quot; name=&quot;form1&quot; target=&quot;_blank&quot;&gt;&lt;div align=&quot;center&quot;&gt;&lt;p&gt;&lt;input name=&quot;fSearch&quot; type=&quot;text&quot;&gt;
&amp;nbsp;&amp;nbsp;&lt;input type=&quot;submit&quot; name=&quot;Submit&quot; value=&quot;Search&quot;&gt;&lt;/p&gt;&lt;/div&gt;&lt;/form&gt;&lt;a href=&quot;http://<%= domain %>&quot; target=&quot;_blank&quot;&gt;<%= domain %>&lt;/a&gt;&lt;/div&gt;&lt;/td&gt;&lt;/tr&gt;&lt;/table&gt;</td>
                    </tr>
                  </table>                  
                  <p>The above code on your site will display as...</p>
                  <table class="addsb">
                    <tr>
                      <td nowrap bgcolor="<%= sboxcolour %>"><div align="center">
                          <p><a href="http://<%= domain %>" target="_blank"><img src="media/linklogo.gif" width="66" height="60" hspace="3" vspace="0" border="0" align="absmiddle"></a>&nbsp;&nbsp;<%= txtSearch %></p>
                          <form action="http://<%= domain %>/results.asp" method="post" name="form1" target="_blank">
                            <div align="center">
                              <p>
                                <input name="fSearch" type="text">
&nbsp;&nbsp;
              <input type="submit" name="Submit" value="Search">
                              </p>
                            </div>
                          </form>
                      <a href="http://<%= domain %>" target="_blank"><%= domain %></a></div></td>
                    </tr>
                  </table>
                  <p>When someone uses the box to do a search, the results will open in a new window (so your site will remain open also).</p></td>
              </tr>
            </table>
            <!-- InstanceEndEditable --></td>
		<% If showSideAd = true Then %>
        <td>
           
            <!--#include file="include/inc_ad_side.asp"-->
           
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
                <td><div align="center"><a href="../default.asp"> Home</a>&nbsp;&nbsp;&middot;&nbsp;&nbsp;<a href="addurl.asp"><strong>Add your URL</strong></a>&nbsp;&nbsp;&middot;&nbsp;&nbsp;<a href="addsearchbox.asp">Add <%= sitename %> toYour Site!</a>&nbsp;&nbsp;&middot;&nbsp;<a href="mailto:<%= emadd %>">Contact</a>
  </div>
                  <div align="center" class="txtPoweredby">Powered by the <a href="http://www.transworldinteractive.net" target="_blank">Transworld Interactive ASP Spider</a> v <%= ver %></div></td>
              </tr>
            </table>
			
	</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
