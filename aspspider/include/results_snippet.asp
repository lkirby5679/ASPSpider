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
	'mark up the snippet
	dim i, regEx
	sSnippet = rstResults.Fields.Item("fContent").Value
	For i = 0 to UBound(aWords)
		Set regEx = New RegExp
		if instr(aWords(i), "*") <> 0 then
			aWords(i) = Replace(aWords(i), "*", "")
			regEx.Pattern = "\b" & aWords(i)
		else
			regEx.Pattern = "\b" & aWords(i) & "\b"
		end if
		regEx.IgnoreCase = true
		regEx.Global = True   
		sSnippet = regEx.Replace(sSnippet, "<b>" & aWords(i) & "</b>")
		Set regEx = Nothing
	Next
		
	'try to find the whole phrase	
	dim showStr1, showStr2
	tempS = ""
	tempS = lcase(sSnippet)
	iStart = instr(tempS, sMarkedString) 
	if iStart > iCharsBefore then
		showStr1 = mid(sSnippet, iStart - iCharsBefore, iCharsExample)
	elseif iStart > 0 then
		showStr1 = left(sSnippet, iStart + iCharsExample)
	else 
		'it's not there as a whole phrase so get the first word only
		iStart = instr(tempS, "<b>") 
		if iStart > iCharsBefore then
			showStr1 = mid(sSnippet, iStart - iCharsBefore, iCharsExample)
		else
			showStr1 = left(sSnippet, iCharsExample)
		end if
	end if
%>
