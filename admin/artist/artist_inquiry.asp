<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.04.14 한용민 생성
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
dim page ,i
	page = request("page")
	if page="" then page=1

dim oip
	set oip = New Cinquiry_list
	oip.FPageSize = 30
	oip.FCurrPage = page
	oip.finquiry_list()
%>

<script language="javascript">
	function view(id){
		document.location.href="/admin/artist/artist_view.asp?idx="+id
	}

	window.resizeTo(850,700);
</script>

<table width="800" align="center" cellpadding="4" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oip.fresultcount >0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" width="50">번호</td>
		<td align="center" width="100">작가명</td>
		<td align="center" width="60">성함</td>
		<td align="center" width="100">연락처</td>
		<td align="center" width="190">이메일</td>
		<td align="center" width="80">상품수</td>
		<td align="center" width="100">현재상품 판매처</td>
		<!--<td align="center" width="50">상세보기</td>-->
	</tr>
	<% for i = 0 to oip.fresultcount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center"><a href="javascript:view(<%= oip.fitemlist(i).fidx %>);"><%= oip.fitemlist(i).fidx %></a></td>
		<td align="center"><a href="javascript:view(<%= oip.fitemlist(i).fidx %>);"><%= oip.fitemlist(i).fartist_name %></a></td>
		<td align="center"><a href="javascript:view(<%= oip.fitemlist(i).fidx %>);"><%= oip.fitemlist(i).fuser_name %></a></td>
		<td align="center"><%= oip.fitemlist(i).fhp %></td>
		<td align="center"><%= oip.fitemlist(i).fmail %></td>
		<td align="center" width="80"><%= oip.fitemlist(i).fsell_count %></td>
		<td align="center" width="100"><%= oip.fitemlist(i).fon_off_isusing %></td>
		<!--<td align="center" width="50"><a href="javascript:view(<%= oip.fitemlist(i).fidx %>);">보기</a></td>-->

	</tr>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oip.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->