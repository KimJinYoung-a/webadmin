<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<%
Dim itemid : itemid = requestCheckvar(request("itemid"),10)
Dim cdl,cdm,cdn

dim oitem
set oitem = new CItemInfo
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItemInfo
	
	if oitem.FREsultCount>0 then
    	cdl = oitem.FOneItem.Fcate_large
    	cdm = oitem.FOneItem.Fcate_mid
    	cdn = oitem.FOneItem.Fcate_small
    end if
	
end if

set oitem = Nothing
%>
<script language='javascript'>
function getOnload(){
    alert('해당 상품이 포함된 카테고리 매핑을 수정합니다.');
    
    location.href='InterParkMatcheDispCate.asp?cdl=<%= cdl %>&cdm=<%= cdm %>&cdn=<%= cdn %>';
}

window.onload=getOnload;

</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->