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
<!--#include file="aspspider/include/userfunctions.asp" -->
<!--#include file="aspspider/include/results_header.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/aspspider_main.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title><%= pageTitle %></title>
<!-- InstanceEndEditable --><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
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

  <table class="resultbox">
              <tr>
                <td>Search term: <strong><%= sTerm %></strong>
                  <% if (domGroup <> "") then Response.Write(" within domain: <strong>" & domGroup & "</strong>") %></td>
              </tr>
              <tr>
                <td><form name="form1" method="post" action="results.asp">
                  <div align="right">
                        <input name="fSearch" type="text" id="fSearch">
                        <input type="submit" name="Submit2" value="Search">
                      </div>
                </form></td>
              </tr>
              <tr class="trWhite">
                <td>
				      <p class="txtMSG" align="center">
				        <% If rstResults.EOF Then %>
				    Sorry, we could find no sites matching your search.
</p>
				  <table class="examples">
                          <tr bgcolor="#FFFFFF">
                            <td colspan="3"><strong><em>How to Search</em> -------------------------------------------------------------------------</strong></td>
                          </tr>
                          <tr bgcolor="#FFFFFF" class="txtSmallFaded">
                            <td>Format </td>
                            <td nowrap>Example</td>
                            <td>Results</td>
                      </tr>
                          <tr bgcolor="#FFFFFF">
                            <td class="txtRed">&nbsp;</td>
                            <td nowrap class="txtRed"><%= ex1 %></td>
                            <td class="bodyTextMedium">the word <span class="txtRed"><%= ex1 %></span></td>
                          </tr>
                          <tr bgcolor="#FFFFFF">
                            <td class="txtRed">&nbsp;</td>
                            <td nowrap class="txtRed"><%= ex1 %>&nbsp;<%= ex2 %></td>
                            <td class="bodyTextMedium">if the phrase does not exist, at least one of the words</td>
                          </tr>
                          <tr bgcolor="#FFFFFF">
                            <td>&quot;...&quot;</td>
                            <td nowrap>&quot;<span class="txtRed"><%= ex1 %></span>&nbsp;<span class="txtRed"><%= ex2 %></span>&quot;</td>
                            <td class="bodyTextMedium">the whole phrase <span class="txtRed"><%= ex1 %><%= ex2 %></span>; if the phrase cannot be found then at least one of the words</td>
                          </tr>
                          <tr bgcolor="#FFFFFF">
                            <td>+</td>
                            <td nowrap>+<span class="txtRed"><%= ex1 %></span>&nbsp;<span class="txtRed"><%= ex2 %></span></td>
                            <td class="bodyTextMedium">must contain <span class="txtRed"><%= ex1 %></span>, optionally with <span class="txtRed"><%= ex2 %></span></td>
                          </tr>
                          <tr bgcolor="#FFFFFF">
                            <td>-</td>
                            <td nowrap>+<span class="txtRed"><%= ex1 %></span> -<span class="txtRed"><%= ex2 %></span></td>
                            <td class="bodyTextMedium">must contain <span class="txtRed"><%= ex1 %></span>, must <strong>not</strong> have the word <span class="txtRed"><%= ex2 %></span></td>
                          </tr>
                          <tr bgcolor="#FFFFFF">
                            <td>*</td>
                            <td nowrap><span class="txtRed"><%= ex1 %></span>*</td>
                            <td class="bodyTextMedium"><span class="txtRed"><%= ex3 %></span>...</td>
                          </tr>
                          <tr bgcolor="#FFFFFF">
                            <td colspan="3" class="txtNB">Note, common words like <strong>the</strong>, <strong>and</strong> and <strong>but</strong> are ignored, i.e. content words only are searched</td>
                          </tr>
                  </table>
						
						
                        <% Else %>
						
                        <p align="center">Showing <%=(rstResults_first)%> to <%=(rstResults_last)%> of <%=(rstResults_total)%> </p>
                        <table class="resultpaging">
                          <tr>
                            <td width="23%" align="center"><% If MM_offset <> 0 Then %>
                                <span class="style3"><a href="<%=MM_moveFirst%>">First</a>
                              <% End If ' end MM_offset <> 0 %>                              </td>
                            <td width="31%" align="center"><% If MM_offset <> 0 Then %>
                                <span class="style3"><a href="<%=MM_movePrev%>">Previous</a>
                              <% End If ' end MM_offset <> 0 %>                              </td>
                            <td width="23%" align="center"><% If Not MM_atTotal Then %>
                                <span class="style3"><a href="<%=MM_moveNext%>">Next</a>
                              <% End If ' end Not MM_atTotal %>                              </td>
                            <td width="23%" align="center"><% If Not MM_atTotal Then %>
                                <span class="style3"><a href="<%=MM_moveLast%>">Last</a>
                              <% End If ' end Not MM_atTotal %>                              </td>
                          </tr>
                    </table>
                        <hr>
                        <% 


	While ((Repeat1__numRows <> 0) AND (NOT rstResults.EOF)) 

%>
                        <p><a href="aspspider/code/_forward.asp?fIndex=<%=(rstResults.Fields.Item("fIndex").Value)%>"><strong><%=(rstResults.Fields.Item("fTitle").Value)%></strong></a>&nbsp;&nbsp;<span class="txtSmall">open page in <a href="aspspider/code/_forward.asp?fIndex=<%=(rstResults.Fields.Item("fIndex").Value)%>" target="_blank">new window</a></span>
                            <!--#include file="aspspider/include/results_snippet.asp" -->
                            <br>
                            <%= trim(showStr1) %><br>
                            <span class="txtSmallFaded"><%=(rstResults.Fields.Item("fPage").Value)%> - <a href="results.asp?fSearch=<%= Request("fSearch") %>&dom=<%=(rstResults.Fields.Item("fDomain").Value)%>">similar pages</a></span></p>
                        <hr noshade>
                        <% 
  	Repeat1__index=Repeat1__index+1
  	Repeat1__numRows=Repeat1__numRows-1
  	rstResults.MoveNext()
	Wend
%>
                        <p>
                          <% End If %>
                        </p>
                        <table class="resultpaging">
                          <tr>
                            <td width="23%" align="center"><% If MM_offset <> 0 Then %>
                                <span class="style3"><a href="<%=MM_moveFirst%>">First</a>
                              <% End If ' end MM_offset <> 0 %>                              </td>
                            <td width="31%" align="center"><% If MM_offset <> 0 Then %>
                                <span class="style3"><a href="<%=MM_movePrev%>">Previous</a>
                              <% End If ' end MM_offset <> 0 %>                              </td>
                            <td width="23%" align="center"><% If Not MM_atTotal Then %>
                                <span class="style3"><a href="<%=MM_moveNext%>">Next</a>
                              <% End If ' end Not MM_atTotal %>                              </td>
                            <td width="23%" align="center"><% If Not MM_atTotal Then %>
                                <span class="style3"><a href="<%=MM_moveLast%>">Last</a>
                              <% End If ' end Not MM_atTotal %>                              </td>
                          </tr>
                </table></td>
              </tr>
              <tr>
                <td><form name="form1" method="post" action="results.asp">
                  <div align="right">
                        <input name="fSearch" type="text" id="fSearch">
                        <input type="submit" name="Submit" value="Search">
                      </div>
                </form>
				<IFRAME src="aspspider/include/autorun.asp" WIDTH=0 HEIGHT=0></IFRAME> 
				</td>
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
<!--#include file="aspspider/include/results_footer.asp" -->

