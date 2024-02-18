<%
'###########################################################
' Description :  웨딩바이블 관리 헤더
' History : 2018-04-10 정태훈 생성
'###########################################################
%>
<% Dim current_url  : current_url = Request.ServerVariables("url") %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr>
	<td align="center" colspan="9" bgcolor="#FFFFFF" height="35">
		<span style="font-weight:900;"><a href="/admin/sitemaster/wedding/index.asp?menupos=<%=menupos%>">웨딩바이블 관리</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="20%" <% If inStr(current_url,"plan_event_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/wedding/plan_event_manager.asp?menupos=<%=menupos%>">기획전 관리</a></td>
	<td width="20%" <% If inStr(current_url,"shopping_list_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/wedding/shopping_list_manager.asp?menupos=<%=menupos%>">쇼핑리스트 관리(PC)</a></td>
	<td width="20%" <% If inStr(current_url,"shopping_list_mo_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/wedding/shopping_list_mo_manager.asp?menupos=<%=menupos%>">쇼핑 리스트 관리(모바일)</a></td>
	<td width="20%" <% If inStr(current_url,"md_pick_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/wedding/md_pick_manager.asp?menupos=<%=menupos%>">MD's Pick</a></td>
	<td width="20%" <% If inStr(current_url,"kit_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/wedding/kit_manager.asp?menupos=<%=menupos%>">Kit</a></td>
</tr>
</table>
<br>