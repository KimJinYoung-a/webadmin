<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ī�� ī�װ�
' History : ���ʻ����ڸ�
'			2017.04.12 �ѿ�� ����(���Ȱ���ó��)
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
		alert('ī�װ��� �����ϼ���.');
		frm.catecode.focus();
		return;
	}
	var selidx = frm.catecode.selectedIndex;
	frm.catename.value = frm.catecode.options[selidx].text;
	//alert(frm.catename.value);
	var ret = confirm('�����Ͻðڽ��ϱ�?');
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
	<td bgcolor="#DDDDFF" width=100>��ID</td>
	<td><%= shopid %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">��ǰID</td>
	<td><%= itemid %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">��ǰ��</td>
	<td><%= itemname %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">ī�װ�</td>
	<td>
		<select name=catecode>
		<option value="">�����ϼ���</option>
		<% for i=0 to ocafecategorylist.FResultCount - 1 %>
		<option value="<%= ocafecategorylist.FItemList(i).FcateCode %>" <% if catecode=ocafecategorylist.FItemList(i).FcateCode then response.write "selected" %> ><%= ocafecategorylist.FItemList(i).FcateName %></option>
		<% next %>
		</select>
		<!-- <input type=button value="ī�װ�����" onClick="popCategoryMaster();"> -->
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=2 align=center><input type=button value="����" onClick="SaveCate(frm);"></td>
</tr>
</form>
</table>
<%
set ocafecategorylist = Nothing
set ocafecategory = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->