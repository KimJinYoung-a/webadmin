<%
'###########################################################
' Description :  ����������
' History : 2013.12.14 ����ȭ
' History : 2014-09-12 ����ȭ ������ ���� �߰�
'###########################################################
%>
<% Dim current_url  : current_url = Request.ServerVariables("url") %>
<script>
function todaymore(){
    var popwin = window.open('/admin/mobile/todaymore/index.asp','mainposcodeedit','width=600,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr>
	<td align="center" colspan="9" bgcolor="#FFFFFF" height="35"> 
		<span id="mtab1" style="font-weight:900;"><a href="/admin/mobile/main/index.asp?menupos=<%=menupos%>">TODAY ����</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<span style="font-weight:900;"><a href="" onclick="todaymore();return false;"><font color="red">ī�װ� ������</font></a></span>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="11%" <% If inStr(current_url,"main_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/main_manager.asp?menupos=<%=menupos%>">��� ���� - ONE Banner</a></td>
	<!--td width="10%" <% If inStr(current_url,"chance") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/chance/index.asp?menupos=<%=menupos%>">JUST1DAY</a></td-->
	<td width="11%" <% If inStr(current_url,"todaykeyword") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/todaykeyword/index.asp?menupos=<%=menupos%>">Keyword(2017ver)</a></td>
	<td width="11%" <% If inStr(current_url,"twinitems") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/twinitems/?menupos=<%=menupos%>">��ǰ ���</a></td>
	<td width="11%" <% If inStr(current_url,"todaybrand") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/todaybrand/?menupos=<%=menupos%>">�귣�� ���</a></td>
	<td width="11%" <% If inStr(current_url,"enjoy") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/appmanage/today/enjoy/?menupos=<%=menupos%>">TREND EVENT</a></td>
	<td width="11%" <% If inStr(current_url,"today_mdpick") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/today_mdpick/index.asp?menupos=<%=menupos%>">MDPICK</a></td>
	<td width="11%" <% If inStr(current_url,"exhibition") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/exhibition/index.asp?menupos=<%=menupos%>">���� ��ȹ�� ��ũ</a></td>
	<td width="11%" <% If inStr(current_url,"sitemaster/just1daymobile2018/") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/just1daymobile2018/index.asp?menupos=<%=menupos%>">JUST1DAY2018</a></td>
	<td width="11%" <% If inStr(current_url,"sitemaster/roundbanner/index.asp") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/roundbanner/index.asp?menupos=<%=menupos%>">������</a></td>
</tr>
</table>
<br>