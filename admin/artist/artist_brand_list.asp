<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 아티스트 브랜드 관리 페이지   
' History : 2012.03.27 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/artist/artist_class.asp"-->
<%
'// 변수 선언
Dim page, isusing, designerid, i
	page = request("page")
	isusing = request("isusing")
	designerid = request("designerid")
	
	if page="" then page=1
	if isusing="" then isusing=""

'// 목록 접수
Dim oGallery
	set oGallery = New cposcode_list
	oGallery.FCurrPage = page
	oGallery.FPageSize=20
	oGallery.FRectIsusing = isusing
	oGallery.FDesignerID = designerid
	oGallery.FArtistBrandList

%>
<script>
function goView(ii){
	location.href = "artist_brand_write.asp?mode=edit&idx="+ii;
}
function gosubmit(page){
    frm.page.value=page;
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="등록" onclick="javascript:location.href='artist_brand_write.asp';">
	</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30"><td><img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>아티스트 브랜드 리스트</b></td></tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="page">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="50">번호</td>
	<td width="190">브랜드</td>
	<td>이미지</td>
	<td width="60">등록일</td>
	<td width="60">사용</td>
</tr>

<% If oGallery.FTotalCount = 0 Then %>
<tr height="25" bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'">
	<td align="center" colspan="6">[데이터가 없습니다.]</td>
</tr>
<% End If %>

<% For i=0 to oGallery.FResultCount-1 %>
<tr height="25" bgcolor="FFFFFF" onClick="goView('<%=oGallery.FItemList(i).fidx%>')" style="cursor:pointer" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" >
	<td align="center" width="50"><%=oGallery.FItemList(i).fidx%></td>
	<td align="center" width="190"><%=oGallery.FItemList(i).fdesignerid%></td>
	<td><img src="<%=uploadUrl%>/artist/brandbanner/<%=oGallery.FItemList(i).ffile2%>" height="50"></td>
	<td align="center" width="160"><%=oGallery.FItemList(i).fregdate%></td>
	<td align="center" width="60"><%=oGallery.FItemList(i).fisusing%></td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="6" align="center">
       	<% If oGallery.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= ohistory.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + oGallery.StartScrollPage to oGallery.StartScrollPage + oGallery.FScrollCount - 1 %>
			<% If (i > oGallery.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(oGallery.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If oGallery.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->