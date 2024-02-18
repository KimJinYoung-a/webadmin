<%
'###########################################################
' Description :  메인페이지
' History : 2014.06.03w 이종화
'###########################################################
%>
<% Dim current_url  : current_url = Request.ServerVariables("url") %>
<%
	'response.write current_url
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr>
	<td colspan="7" align="center" bgcolor="#FFFFFF" height="35"> 앱 Today 메인 관리 </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="14.2%" <% If inStr(current_url,"todaywish") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/appmanage/today/todaywish/?menupos=<%=menupos%>">TODAY WISH(APP)</a></td>
    <!-- <td width="14.2%" <% If inStr(current_url,"enjoyevent") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/appmanage/today/enjoyevent/?menupos=<%=menupos%>">ENJOY EVENT(APP)-구버전</a></td> -->
	<td width="14.2%" <% If inStr(current_url,"enjoy") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/appmanage/today/enjoy/?menupos=<%=menupos%>">ENJOY EVENT(APP)</a></td>
    <td width="14.2%" <% If inStr(current_url,"deal") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/appmanage/today/todaydeal/?menupos=<%=menupos%>">TODAY DEAL(APP)</a></td>
	<td width="14.2%" <% If inStr(current_url,"hotkeyword") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/hotkeyword/?menupos=<%=menupos%>">HOT KEYWORD(2014)</a></td>
    <td width="14.2%" ></td>
    <td width="14.2%" ></td>
    <td width="14.2%" ></td>
</tr>
</table>
<br>