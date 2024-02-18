<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  일별Color 리스트
' History : 2013.12.17 김진영 생성
'			2014.02.13 한용민 수정
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
	lColor.FPageSize = 20
	lColor.FCurrPage = page
	lColor.sbDailyColorList
%>

<script language="javascript">

function gotoColor(yyyymmdd){
	location.href='/admin/appmanage/appDailyColorModify.asp?yyyymmdd='+yyyymmdd+'&menupos=<%=menupos%>';
}
function gosubmit(page){
    var frm = document.fsearch;
    frm.page.value=page;
	frm.submit();
}

</script>

<form name="fsearch" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page">
</form>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" class="button" value="등록" onclick="gotoColor('');">
	</td>
</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%=lColor.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="100" align="center">오픈할 날짜</td>
	<td width="250" align="center">대표이미지 [2048x1536 (4:3)]</td>
	<td width="250" align="center">대표이미지2 [1920x1080 (16:9)]</td>
	<td width="250" align="center">오픈색상명</td>
	<td width="200" align="center">등록일</td>
	<td width="200" align="center">최종수정일</td>
	<td align="center">수정</td>
</tr>
<%
	For i = 0 to lColor.fresultcount -1
%>
<tr height="25" bgcolor="FFFFFF">
	<td width="100" align="center"><%= lColor.FColorList(i).FYyyymmdd %></td>
	<td width="250" align="center">
		<img src='<%= "http://thumbnail.10x10.co.kr/imgstatic" & mid(lColor.FColorList(i).FImageUrl,29,65) & "?cmd=thumb&width=50&height=50" %>' border="0" width="50" height="50">
	</td>
	<td width="250" align="center">
		<img src='<%= "http://thumbnail.10x10.co.kr/imgstatic" & mid(lColor.FColorList(i).FImageUrl2,29,65) & "?cmd=thumb&width=50&height=50" %>' border="0" width="50" height="50">
	</td>
	<td width="250" align="center"><%= lColor.FColorList(i).FColorName %></td>
	<td width="200" align="center"><%= lColor.FColorList(i).FRegdate %></td>
	<td width="200" align="center"><%= lColor.FColorList(i).FLastupdate %></td>
	<td align="center">
		<input type="button" class="button" value="수정[상세상품:<%= lColor.FColorList(i).FRegedItemCnt %>개]" onclick="gotoColor('<%= lColor.FColorList(i).FYyyymmdd %>');">
	</td>
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

<% Set lColor = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->