<%
'###########################################################
' Description :  메인페이지
' History : 2018.06.21 정태훈 생성
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
		<span id="mtab1" style="font-weight:900;"><a href="/admin/mobile/gnbmanager/index.asp?menupos=<%=menupos%>">GNB 컨텐츠 관리</a></span>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="12.5%" <% If inStr(current_url,"event") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/gnbmanager/event/gnb_main_event_manager.asp?menupos=<%=menupos%>">GNB 메인 이벤트</a></td>
	<td width="12.5%" <% If inStr(current_url,"brand") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/gnbmanager/brand/gnb_brand_manager.asp?menupos=<%=menupos%>">브랜드 배너</a></td>
	<td width="12.5%" <% If inStr(current_url,"1") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>></td>
	<td width="12.5%" <% If inStr(current_url,"2") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>></td>
	<td width="12.5%" <% If inStr(current_url,"3") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>></td>
	<td width="12.5%" <% If inStr(current_url,"4") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>></td>
	<td width="12.5%" <% If inStr(current_url,"5") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>></td>
	<td width="12.5%" <% If inStr(current_url,"6") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>></td>
</tr>
</table>
<br>