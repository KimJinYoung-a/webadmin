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
dim page ,i
	page = request("page")
	if page="" then page=1

dim oip
	set oip = New Cinquiry_list
	oip.FPageSize = 30
	oip.FCurrPage = page
	oip.frecommend_list()
%>

<script language="javascript">

	function view(artist_idx){
		document.location.href="/admin/artist/artist_recommendview.asp?artist_idx="+artist_idx
	}

	function del(artist_idx){
		document.location.href="/admin/artist/artist_recommend_process.asp?artist_idx="+artist_idx+"&mode=del"
	}
	
	window.resizeTo(1024,768);
</script>

<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oip.fresultcount >0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" width="30">번호</td>
		<td align="center">아티스트</td>
		<td align="center">아티스트<br>태그</td>
		<td align="center">홈페이지</td>
		<td align="center">블로그</td>
		<td align="center">추천이유</td>
		<td align="center" width="100">작성고객</td>
		<td align="center" width="100">비고</td>
	</tr>
	<% for i = 0 to oip.fresultcount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center"><%= oip.fitemlist(i).fartist_idx %></td>
		<td align="center"><%= oip.fitemlist(i).fartist_name %></td>
		<td align="center"><%= chrbyte(oip.fitemlist(i).ftag,20,"Y") %></td>
		<td align="center"><%= chrbyte(oip.fitemlist(i).fhomepage,20,"Y") %></td>
		<td align="center"><%= chrbyte(oip.fitemlist(i).fblog,20,"Y") %></td>
		<td align="center" width="80"><%= chrbyte(nl2br(oip.fitemlist(i).fwhyrecommend),20,"Y") %></td>
		<td align="center" width="100"><%= oip.fitemlist(i).fuserid %></td>
		<td align="center" width="100">
			<input type="button" value="상세보기" class="button" onclick="view(<%= oip.fitemlist(i).fartist_idx %>);"><br>
			<input type="button" value="삭제" class="button" onclick="del(<%= oip.fitemlist(i).fartist_idx %>)">
		</td>
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
<%
	set oip = nothing
%>	