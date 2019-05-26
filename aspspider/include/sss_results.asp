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
<!--#include file="results_header.asp" -->
<table width="100%"  border="0" cellpadding="10" cellspacing="0" bgcolor="#6666FF" class="searchTable">
  <tr valign="top">
    <td>Search term: <strong><%= sTerm %></strong>
        <!--#include file="sss.asp" -->
        <table width="100%"  border="0" cellspacing="0" cellpadding="10">
          <tr>
            <td bgcolor="#FFFFFF"><% If rstResults.EOF Then %>
                <p align="center" class="title"><br>
              Sorry, we could find no sites matching your search.</p>
                <% Else %>
                <p></p>
                <p align="center">Showing <%=(rstResults_first)%> to <%=(rstResults_last)%> of <%=(rstResults_total)%> </p>
                <table border="0" width="50%" align="center">
                  <tr>
                    <td width="23%" align="center"><% If MM_offset <> 0 Then %>
                        <span class="style3"><a href="<%=MM_moveFirst%>">First</a>
                        <% End If ' end MM_offset <> 0 %>                      </td>
                    <td width="31%" align="center"><% If MM_offset <> 0 Then %>
                        <span class="style3"><a href="<%=MM_movePrev%>">Previous</a>
                        <% End If ' end MM_offset <> 0 %>                      </td>
                    <td width="23%" align="center"><% If Not MM_atTotal Then %>
                        <span class="style3"><a href="<%=MM_moveNext%>">Next</a>
                        <% End If ' end Not MM_atTotal %>                      </td>
                    <td width="23%" align="center"><% If Not MM_atTotal Then %>
                        <span class="style3"><a href="<%=MM_moveLast%>">Last</a>
                        <% End If ' end Not MM_atTotal %>                      </td>
                  </tr>
                </table>
                <hr>
                <% 


	While ((Repeat1__numRows <> 0) AND (NOT rstResults.EOF)) 

%>
                <p><a href="aspspider/code/_forward.asp?fIndex=<%=(rstResults.Fields.Item("fIndex").Value)%>"><strong><%=(rstResults.Fields.Item("fTitle").Value)%></strong></a>&nbsp;&nbsp;<span class="bodyTextSmall">open page in <a href="aspspider/code/_forward.asp?fIndex=<%=(rstResults.Fields.Item("fIndex").Value)%>" target="_blank">new window</a></span>
                    <!--#include file="results_snippet.asp" -->
                    <br>
                    <%= trim(showStr1) %><br>
                    <span class="bodyTextMediumFaded"><%=(rstResults.Fields.Item("fPage").Value)%></span></p>
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
                <table border="0" width="50%" align="center">
                  <tr>
                    <td width="23%" align="center"><% If MM_offset <> 0 Then %>
                        <span class="style3"><a href="<%=MM_moveFirst%>">First</a>
                        <% End If ' end MM_offset <> 0 %>                      </td>
                    <td width="31%" align="center"><% If MM_offset <> 0 Then %>
                        <span class="style3"><a href="<%=MM_movePrev%>">Previous</a>
                        <% End If ' end MM_offset <> 0 %>                      </td>
                    <td width="23%" align="center"><% If Not MM_atTotal Then %>
                        <span class="style3"><a href="<%=MM_moveNext%>">Next</a>
                        <% End If ' end Not MM_atTotal %>                      </td>
                    <td width="23%" align="center"><% If Not MM_atTotal Then %>
                        <span class="style3"><a href="<%=MM_moveLast%>">Last</a>
                        <% End If ' end Not MM_atTotal %>                      </td>
                  </tr>
              </table></td>
          </tr>
        </table>
       <!--#include file="sss.asp" -->
      </td>
  </tr>
</table>
<!--#include file="results_footer.asp" -->

