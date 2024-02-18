<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lecturecls.asp"-->
<%
dim olec
set olec = new CLecture
olec.getNotRegItemList

dim i
%>
<script language='javascript'>
function InPutItem(frm){
	opener.lecform.linkitemid.value = frm.itemid.value;
	opener.lecform.lectitle.value = frm.itemname.value;
	opener.lecform.lecturerid.value = frm.makerid.value;
	opener.lecform.lecturer.value = frm.makername.value;
	opener.lecform.lecsum.value = frm.sellcash.value;
	opener.lecform.lecspace.value = '텐바이텐 컬리지 (대학로)';
	window.close();
}
</script>
<table width=600 border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td width=100>상품번호</td>
	<td width=100>상품명</td>
	<td width=100>강사ID</td>
	<td width=100>강사명</td>
	<td width=100>강의비</td>
	<td width=100>선택</td>
</tr>
<% for i=0 to olec.FResultCount -1 %>
<form name=buffrm<%= i %> >
<input type=hidden name=itemid value="<%= olec.FItemList(i).FItemID %>">
<input type=hidden name=itemname value="<%= olec.FItemList(i).FItemName  %>">
<input type=hidden name=makerid value="<%= olec.FItemList(i).FMakerid %>">
<input type=hidden name=makername value="<%= olec.FItemList(i).FMakerName %>">
<input type=hidden name=sellcash value="<%= olec.FItemList(i).FSellcash %>">
<tr bgcolor="#FFFFFF">
	<td ><%= olec.FItemList(i).FItemID %></td>
	<td ><%= olec.FItemList(i).FItemName %></td>
	<td ><%= olec.FItemList(i).FMakerid %></td>
	<td ><%= olec.FItemList(i).FMakerName %></td>
	<td ><%= olec.FItemList(i).FSellcash %></td>
	<td ><input type=button value="선택" onclick="InPutItem(buffrm<%= i %>)"></td>
</tr>
</form>
<% next %>
</table>
<%
set olec = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->