<%
'###########################################################
' Description :  �������̺� ���� ���
' History : 2018-04-10 ������ ����
'###########################################################
%>
<% Dim current_url  : current_url = Request.ServerVariables("url") %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr>
	<td align="center" colspan="9" bgcolor="#FFFFFF" height="35">
		<span style="font-weight:900;"><a href="/admin/sitemaster/wedding/index.asp?menupos=<%=menupos%>">�������̺� ����</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="20%" <% If inStr(current_url,"plan_event_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/wedding/plan_event_manager.asp?menupos=<%=menupos%>">��ȹ�� ����</a></td>
	<td width="20%" <% If inStr(current_url,"shopping_list_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/wedding/shopping_list_manager.asp?menupos=<%=menupos%>">���θ���Ʈ ����(PC)</a></td>
	<td width="20%" <% If inStr(current_url,"shopping_list_mo_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/wedding/shopping_list_mo_manager.asp?menupos=<%=menupos%>">���� ����Ʈ ����(�����)</a></td>
	<td width="20%" <% If inStr(current_url,"md_pick_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/wedding/md_pick_manager.asp?menupos=<%=menupos%>">MD's Pick</a></td>
	<td width="20%" <% If inStr(current_url,"kit_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/wedding/kit_manager.asp?menupos=<%=menupos%>">Kit</a></td>
</tr>
</table>
<br>