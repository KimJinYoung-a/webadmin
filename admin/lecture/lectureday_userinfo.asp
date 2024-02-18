<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/lectureday_userinfocls.asp"-->
<%
dim page
dim i, ix, olec

page = request("page")
if page="" then page=1

set olec = new CLecture
olec.FPageSize=20
olec.FCurrPage = page
olec.GetLectureList
%>
<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}

function AddLec(iitemid,iitemoption,iitemea){
	document.lecadd.itemid.value=iitemid;
	document.lecadd.itemoption.value=iitemoption;
	document.lecadd.itemea.value=iitemea;
	document.lecadd.submit();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>
<table width="800" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td ><a href="lectureday_userreg.asp?mode=add">[강사등록]</a></td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<form name="frm" method="post" action="/admin/lecture/lectureday_userimg.asp">
<input type="hidden" name="page" value="">
<tr bgcolor="#DDDDFF">
	<td align="center" width="30">Idx</td>
	<td align="center" width="100">이미지</td>
	<td align="center" width="100">강사ID</td>
	<td align="center" width="100">강사명</td>
	<td align="center" width="150">강좌명</td>
	<td align="center" width="50">사용유무</td>
	<td align="center" width="120">등록일</td>
	<td align="center" width="120">작품이미지등록</td>
	<td align="center" width="120">공방이미지등록</td>
</tr>
<% for i=0 to olec.FResultCount - 1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><% = olec.FItemList(i).Fidx %></td>
	<td align="center"><a href="lectureday_userreg.asp?idx=<% = olec.FItemList(i).Fidx %>&mode=edit"><img src="<% = olec.FItemList(i).Flecimg %>" width="50" height="50" border="0"></a></td>
	<td align="center"><% = olec.FItemList(i).Flectureid %></td>
	<td align="center"><% = olec.FItemList(i).Flecturer %></td>
	<td align="center"><% = olec.FItemList(i).Flecturename %></td>
	<td align="center"><%= olec.FItemList(i).FIsUsing %></td>
	<td align="center"><%= FormatDate(olec.FItemList(i).Fregdate,"0000.00.00") %></td>
	<td align="center"><a href="lectureday_userimg.asp?masteridx=<% = olec.FItemList(i).Fidx %>&lectureid=<% = olec.FItemList(i).Flectureid %>">등록</a></td>
	<td align="center"><a href="lectureday_workroom_reg.asp?masteridx=<% = olec.FItemList(i).Fidx %>&lectureid=<% = olec.FItemList(i).Flectureid %>">등록(공방)</a><br>
							   <a href="lectureday_workroom.asp?masteridx=<% = olec.FItemList(i).Fidx %>&lectureid=<% = olec.FItemList(i).Flectureid %>">수정</a></td>
</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="14" height="30" align="center">
		<% if olec.HasPreScroll then %>
			<a href="javascript:NextPage('<%= olec.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for ix=0 + olec.StarScrollPage to olec.FScrollCount + olec.StarScrollPage - 1 %>
			<% if ix>olec.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if olec.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</form>
</table>
<%
set olec = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->