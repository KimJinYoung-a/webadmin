<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/limit_item_cls.asp"-->

<%
dim itemid, itemoption, mayreallimtno
dim mode, addno
itemid		= request("itemid")
itemoption	= request("itemoption")
mayreallimtno	= request("mayreallimtno")
mode	= request("mode")
addno	= request("addno")

dim sqlStr, result

if mode="addmaylimit" then
	result = UpdateItemLimitCount(itemid, itemoption, mayreallimtno, 0)
	response.write result
	
	sqlStr = "exec db_summary.dbo.sp_Ten_SellYnSetByLimitNo " & CStr(itemid)
	dbget.Execute sqlStr
	
elseif mode="addlimitno" then
	result = AddItemLimitNo(itemid, itemoption, addno)
end if

%>

<script language='javascript'>
alert('수정 되었습니다.');
// opener.location.reload();
opener.location.href=opener.document.location;
window.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->

