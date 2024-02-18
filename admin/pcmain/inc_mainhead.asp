<%
'###########################################################
' Description :  메인페이지 관리 - PCMAIN
' History : 2018-03-05 이종화
'###########################################################
%>
<% Dim current_url  : current_url = Request.ServerVariables("url") %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr>
	<td align="center" colspan="12" bgcolor="#FFFFFF" height="35">
		<span style="font-weight:900;"><a href="/admin/pcmain/index.asp?menupos=<%=menupos%>">PC메인 관리</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="10%" <% If inStr(current_url,"main_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/pcmain_manager.asp?menupos=<%=menupos%>">배너 관리</a></td>
	<td width="10%" <% If inStr(current_url,"enjoy_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/enjoy_manager.asp?menupos=<%=menupos%>">엔조이 기획전</a></td>
	<td width="10%" <% If inStr(current_url,"main_md_recommend_flash") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/main_md_recommend_flash.asp?menupos=<%=menupos%>">MDPICK</a></td>
	<!--<td width="10%" <% If inStr(current_url,"gather_event_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/gather_event_manager.asp?menupos=<%=menupos%>">기획전 모음</a></td>-->
	<!--td width="9.1%" <% If inStr(current_url,"chance/index") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/chance/index.asp?menupos=<%=menupos%>">저스트원데이&주말특가</a></td-->
	<td width="10%" <% If inStr(current_url,"multievent") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/pcmain/multievent/index.asp?menupos=<%=menupos%>">이벤트1~16배너</a></td>
	<td width="10%" <% If inStr(current_url,"look/index") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/look/index.asp?menupos=<%=menupos%>">Look시즌2</a></td>
	<td width="10%" <% If inStr(current_url,"brandbig/index") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/brandbig/index.asp?menupos=<%=menupos%>">브랜드빅</a></td>
	<td width="10%" <% If inStr(LCase(current_url),"onlybrand/index") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/onlybrand/index.asp?menupos=<%=menupos%>">온리브랜드</a></td>
	<td width="10%" <% If inStr(LCase(current_url),"wishbest/index") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/wishbest/index.asp?menupos=<%=menupos%>">위시베스트</a></td>
	<!--td width="9.1%" <% If inStr(LCase(current_url),"tempjust1day/index") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/tempjust1day/index.asp?menupos=<%=menupos%>">저스트원데이3개 테스트(임시)</a></td-->
	<td width="10%" <% If inStr(current_url,"new_brand_manager") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/new_brand_manager.asp?menupos=<%=menupos%>">New Brand</td>
	<td width="10%" <% If inStr(current_url,"just1day2018") > 0 Then %>bgcolor="#ddddff"<% Else %>bgcolor="#FFFFFF"<% End If %>><a href="/admin/sitemaster/just1day2018/index.asp?menupos=<%=menupos%>">Just1Day2018</td>
</tr>
</table>
<br>