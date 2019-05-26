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
<% 
	' URL added successfully
	dim MSG
	MSG = ""
	If Request("flag") = 1 Then 
		MSG = "<p align='center' class='txtMSG'><b>Your website has been registered for spidering. Sites usually appear online within a few days.</b></p> <p align='center' class='txtSmall'><a href='default.asp'>Click here</a> to return to the main index page or use the form below to add another URL.</p>"   	
	End If 
%> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/aspspider_main.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title><%= pageTitle %></title>
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
            <table class="resultbox">
              <tr>
                <td class="trWhite"><p align="center" class="txtTitle">Add URL</p>                 
                    <% If Request("flag") = 1 Then %>                  
                    <p align='center'><b>Your website has been registered for spidering. Sites usually appear online within 24-48 hours.</b></p>                  
                    <p align='center'>Our service is free. However, we would appreciate a linkto help spread the word. </p>
                    <p align='center'>If you can, please include the following HTML somewhere on your site:</p>
                    <table width="0%"  border="0" align="center" cellpadding="10" cellspacing="0" bgcolor="#FFFFCC">
                      <tr>
                        <td nowrap>&lt;a href=&quot;http://<%= domain %>&quot; target=&quot;_blank&quot;&gt;<%= txtSearch %>&lt;/a&gt;</td>
                      </tr>
                    </table>
                    <p align='center'>You can also add our search facility to your own site and offer your visitors an extra  service. <br>
                        <a href="addsearchbox.asp"> Click here</a> to read more about this. </p>                    <p align='center'><a href='../default.asp'>Click here</a> to return to the main index page.</p>
                    <hr noshade class="bodyTextMedium">                  <% End If %>
                  <p align="center" class="txtTitle">Use this form to ADD a site to <%= sitename %></p>                  
                  <ul>
                    <li><%= add1 %><br>
&nbsp;                    </li>
                    <li> Your site will be checked by one of our moderators and, if suitable, added to this database.<br>
&nbsp;</li>
                    <li>Only <strong>top level domains</strong> will be included.                       This means we will include <em>www.yourdomain.com</em> but not sites like <em>www.genericdomain.com/yourdomain</em> <br>
&nbsp;                    </li>
                    <li>Each site is spidered. This means we collect information from every page on your site and add it to our database. <strong><br>
&nbsp;&nbsp;                    </strong></li>
                    <li>The index page of your site shows up within a few hours; other pages will appear later.</li>
                  </ul>                  
                  <form name="form1" method="post" action="code/_addurl.asp">
                    <table width="0"  border="0" align="center" cellpadding="5" cellspacing="5">
                      <tr>
                        <td><div align="right"><strong>Site URL </strong></div></td>
                        <td><input name="fURL" type="text" id="fURL" value="http://" size="60"></td>
                      </tr>
                      <tr>
                        <td colspan="2"><div align="center">
                            <input type="submit" name="Submit" value="Submit">
                        </div></td>
                      </tr>
                    </table>
                </form>                  </td>
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
