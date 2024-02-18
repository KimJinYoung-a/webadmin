<%
'###########################################################
' Description :  메인페이지
' History : 2014.04.01 김진영
'###########################################################
%>
<% Dim current_url  : current_url = Request.ServerVariables("url") %>
<%
	'response.write current_url
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr>
	<td colspan="7" align="center" bgcolor="#FFFFFF" height="35"> 비트윈 메인 관리 </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="33%" <% If inStr(current_url,"topbanner") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/etc/between/main/topbanner/index.asp?menupos=<%=menupos%>">최상ONE배너</a></td>
    <td width="33%" <% If inStr(current_url,"3banner") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/etc/between/main/3banner/index.asp?menupos=<%=menupos%>">3Banner</a></td>
    <td width="33%" <% If inStr(current_url,"mdpick") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/etc/between/main/mdpick/index.asp?menupos=<%=menupos%>">MDPICK</a></td>
</tr>
</table>
<br>