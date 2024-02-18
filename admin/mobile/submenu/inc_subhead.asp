<%
'###########################################################
' Description :  메인페이지
' History : 2015-09-22 이종화
'###########################################################
%>
<% Dim current_url  : current_url = Request.ServerVariables("url") %>
<%
	'response.write current_url
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr>
	<td colspan="7" align="center" bgcolor="#FFFFFF" height="35"> 모바일 서브 관리 &nbsp;&nbsp;&nbsp;&nbsp;<a href="/admin/mobile/topcatecode/index.asp?menupos=<%=menupos%>"><strong>GNB_CODE관리<strong></a></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td width="14.2%" <% If inStr(current_url,"topcatebanner") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/topcatebanner/main_manager.asp?menupos=<%=menupos%>">카테고리별-ONE 배너</a></td>
    <td width="14.2%" <% If inStr(current_url,"topevtbanner") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/topevtbanner/index.asp?menupos=<%=menupos%>">TOP 2 EVENT</a></td>
    <td width="14.2%" <% If inStr(current_url,"topkeyword") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/topkeyword/index.asp?menupos=<%=menupos%>">TOP Keyword</a></td>
    <td width="14.2%" <% If inStr(current_url,"topmdpick") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/topmdpick/index.asp?menupos=<%=menupos%>">MD`S PICK</a></td>
    <td width="14.2%" >&nbsp;</td>
    <td width="14.2%" >&nbsp;</td>
	<!--<td width="14.2%" <% If inStr(current_url,"showbanner") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/showbanner/index.asp?menupos=<%=menupos%>">Show Banner</a></td> -->
</tr>
</table>
<br>