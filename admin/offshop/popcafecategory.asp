<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 카페 카테고리
' History : 최초생성자모름
'			2017.04.12 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/cafecategorycls.asp"-->

<%
dim shopid, itemid, itemname
shopid = requestCheckVar(request("shopid"),32)
itemid = requestCheckVar(request("itemid"),10)
itemname = replace(request("itemname"),"||","&")

dim ocafecategory,i
set ocafecategory = new CCafeCategorySell
ocafecategory.FRectShopid = shopid
ocafecategory.FRectItemID = itemid
ocafecategory.FRectItemName = itemname
ocafecategory.GetCafeCategoyLink

dim catecode

catecode = ocafecategory.FItemList(0).Fcatecode
if catecode<>"" then
	itemname = ocafecategory.FItemList(0).FitemName
end if

dim ocafecategorylist
set ocafecategorylist = new CCafeCategorySell
ocafecategorylist.GetCafeCategoryList
%>
<script type='text/javascript'>

function popCategoryMaster(){
	var popwin = window.open("popcafecategorymaster.asp","popcafecategorymaster","width=640 height=580 scrollbars=yes");
}

function SaveCate(frm){
	if (frm.catecode.value.length<1){
		alert('카테고리를 지정하세요.');
		frm.catecode.focus();
		return;
	}
	var selidx = frm.catecode.selectedIndex;
	frm.catename.value = frm.catecode.options[selidx].text;
	//alert(frm.catename.value);
	var ret = confirm('저장하시겠습니까?');
	if (ret){
		frm.submit();
	}
}
</script>
<table width="500" border="0" cellspacing="1" cellpadding="3" bgcolor="#3d3d3d" class=a>
<form name=frm method=post action=docafecategory.asp>
<input type=hidden name=mode value="linkitem">
<input type=hidden name=shopid value="<%= shopid %>">
<input type=hidden name=itemid value="<%= itemid %>">
<input type=hidden name=itemname value="<%= itemname %>">
<input type=hidden name=catename value="">

<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>샾ID</td>
	<td><%= shopid %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">상품ID</td>
	<td><%= itemid %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">상품명</td>
	<td><%= itemname %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">카테고리</td>
	<td>
		<select name=catecode>
		<option value="">선택하세요</option>
		<% for i=0 to ocafecategorylist.FResultCount - 1 %>
		<option value="<%= ocafecategorylist.FItemList(i).FcateCode %>" <% if catecode=ocafecategorylist.FItemList(i).FcateCode then response.write "selected" %> ><%= ocafecategorylist.FItemList(i).FcateName %></option>
		<% next %>
		</select>
		<!-- <input type=button value="카테고리관리" onClick="popCategoryMaster();"> -->
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=2 align=center><input type=button value="저장" onClick="SaveCate(frm);"></td>
</tr>
</form>
</table>
<%
set ocafecategorylist = Nothing
set ocafecategory = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->