<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/cafecategorycls.asp"-->
<%
dim ocafecategorylist, i
set ocafecategorylist = new CCafeCategorySell
ocafecategorylist.GetCafeCategoryList
%>
<script language=javascript>
function DelCatemaster(frm,icatecode){
	var ret = confirm('카테고리를 삭제하면 기존 매칭된 데이터도 전부 삭제 됩니다. 삭제하시겠습니까?');
	if (ret){
		frm.mode.value = "delcate";
		frm.catecode.value = icatecode;
		frm.submit();
	}
}

function parentRefresh(){
	//opener.document.location.reload();
}

function SaveCate(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.submit();
	}
}
</script>
<body onunload="parentRefresh();" >
<table width="500" border="0" cellspacing="1" cellpadding="3" bgcolor="#3d3d3d" class=a>
<form name=frm method=post action=docafecategory.asp>
<input type=hidden name=mode value="inputcate">
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>카테고리코드</td>
	<td><input type="text" name=catecode value="" size=2 maxlength=2> (숫자 2, 중복불가)</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">카테고리명</td>
	<td><input type="text" name=catename value="" size=10 maxlength=16> (문자 16max)</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=2 align=center><input type=button value="저장" onClick="SaveCate(frm);"></td>
</tr>
</form>
</table>
<br>
<span class=a>* 등록된 카테고리</span>
<table width="500" border="0" cellspacing="1" cellpadding="3" bgcolor="#3d3d3d" class=a>
<tr bgcolor="#DDDDFF">
	<td width=100>카테고리코드</td>
	<td>카테고리명</td>
	<td width=50>삭제</td>
</tr>
<% for i=0 to ocafecategorylist.FResultCount -1 %>
<tr bgcolor="#FFFFFF">
	<td width=100><%= ocafecategorylist.FItemList(i).FCateCode %></td>
	<td><%= ocafecategorylist.FItemList(i).FCateName %></td>
	<td><a href="javascript:DelCatemaster(frm,'<%= ocafecategorylist.FItemList(i).FCateCode %>');">x</a></td>
</tr>
<% next %>
</table>
<%
set ocafecategorylist = Nothing
%>
</body>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->