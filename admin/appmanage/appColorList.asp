<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  AppColor 리스트
' History : 2013.12.16 김진영 생성
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/appmanage/appColorCls.asp" -->
<%
Dim lColor, page, i
page = request("page")

If page = "" Then page = 1

set lColor = new AppColorList
	lColor.FPageSize = 35
	lColor.FCurrPage = page
	lColor.sbColorList
%>
<script language="javascript">
function gotoColor(code){
	var regColor;
	regColor = window.open('/admin/appmanage/appColorModify.asp?wcolorCode='+code+'&menupos='+<%=menupos%>,'regColor','width=1600,height=600, resizable=1, scrollbars=1');
	regColor.focus();
}
function gosubmit(page){
    var frm = document.fsearch;
    frm.page.value=page;
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="등록" onclick="gotoColor('');">
	</td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=lColor.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="50">번호</td>
	<td width="70">색상코드</td>
	<td>색상명</td>
	<td width="90">iconImageUrl1</td>
	<td width="90">iconImageUrl2</td>
	<td width="70">색상String</td>
	<td width="90">글자 RGB코드</td>
	<td width="70">사용유무</td>
	<td width="150">등록일</td>
	<td width="70">순서</td>
</tr>
<%
	For i = 0 to lColor.fresultcount -1
%>
<tr height="25" bgcolor="FFFFFF" onClick="gotoColor('<%=lColor.FColorList(i).FColorCode%>')" style="cursor:pointer" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" >
	<td width="50" align="center"><%= lColor.FColorList(i).FIdx %></td>
	<td width="70" align="center"><%= lColor.FColorList(i).FColorCode %></td>
	<td><%= lColor.FColorList(i).FColorName %></td>
	<td width="90" align="center"><img src="<%=lColor.FColorList(i).FIconImageUrl1%>" border="0" width="50" height="50"></td>
	<td width="90" align="center"><img src="<%=lColor.FColorList(i).FIconImageUrl2%>" border="0" width="50" height="50"></td>
	<td width="70" align="center"><%= lColor.FColorList(i).FColor_str %></td>
	<td width="90" align="center"><%= lColor.FColorList(i).FWord_rgbCode %></td>
	<td width="70" align="center"><%= lColor.FColorList(i).FIsusing %></td>
	<td width="150" align="center"><%= lColor.FColorList(i).FRegdate %></td>
	<td width="70" align="center"><%= lColor.FColorList(i).FSortNo %></td>
</tr>
<%
	Next
%>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If lColor.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= lColor.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + lColor.StartScrollPage to lColor.StartScrollPage + lColor.FScrollCount - 1 %>
			<% If (i > lColor.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(lColor.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If lColor.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
</table>
<form name="fsearch" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page">
</form>
<% Set lColor = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->