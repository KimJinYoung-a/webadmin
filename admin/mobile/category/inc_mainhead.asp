<%
'###########################################################
' Description :  모바일 카테고리 메인 관리페이지 메뉴
' History : 2020.11.30 정태훈 생성
'###########################################################
%>
<% Dim current_url  : current_url = Request.ServerVariables("url") %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr>
	<td align="center" colspan="9" bgcolor="#FFFFFF" height="35"> 
		<span id="mtab1" style="font-weight:900;"><a href="/admin/mobile/category/index.asp?menupos=<%=menupos%>">모바일 카테고리 메인 관리</a></span>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="11%" <% If inStr(current_url,"main_rolling") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/category/main_rolling.asp?menupos=<%=menupos%>">롤링 배너 관리</a></td>

	<td width="11%" <% If inStr(current_url,"main_brand") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/category/main_brand.asp?menupos=<%=menupos%>">브랜드 관리</a></td>

	<td width="11%" <% If inStr(current_url,"main_event") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/mobile/category/main_event.asp?menupos=<%=menupos%>">기획전 관리</a></td>
</tr>
</table>
<br>