<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.10 한용민 생성
'	Description : artist gallery
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/artistGalleryCls.asp" -->

<%
dim page ,i , idx
	page = request("page")
	idx = request("idx")
	if page="" then page=1

dim oip
	set oip = New Cinquiry_list
	oip.frectidx = idx
	oip.finquiry_oneitem()
%>

<table width="700" align="center" cellpadding="4" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">작가명</td>
		<td align="center"><%= nl2br(oip.foneitem.fartist_name) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">성함</td>
		<td align="center"><%= oip.foneitem.fartist_name %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">주소</td>
		<td align="center"><%= nl2br(oip.foneitem.faddress) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">연락처</td>
		<td align="center"><%= oip.foneitem.fhp %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">이메일</td>
		<td align="center"><%= oip.foneitem.fmail %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">사업자등록번호</td>
		<td align="center"><%= oip.foneitem.flicense %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">홈페이지</td>
		<td align="center"><a href="<%= oip.foneitem.fhomepage %>" class="a" target="_blank"><%= oip.foneitem.fhomepage %></a></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">아티스트소개</td>
		<td align="center"><%= nl2br(oip.foneitem.fuser_info) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">상품수</td>
		<td align="center"><%= oip.foneitem.fsell_count %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">현재상품 판매처</td>
		<td align="center"><%= oip.foneitem.fon_off_isusing %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">상품소개</td>
		<td align="center"><%= nl2br(oip.foneitem.fitem_info) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("gray")%>">첨부파일</td>
		<td align="center"><a href="<%=staticImgUrl%>/<%= oip.foneitem.ffile1 %>">다운받기</a></td>
	</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
	set oip = nothing
%>	