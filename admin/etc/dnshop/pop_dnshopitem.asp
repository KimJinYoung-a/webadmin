<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->

<%
dim itemid, itemname, eventid, mode, i, makerid, delitemcode, vAction, deljaehyu

vAction = NullFillWith(request("action"),"")
If vAction = "del" Then
	Call Proc()
End IF

itemid  = request("itemid")
itemname= request("itemname")
eventid = request("eventid")
makerid= request("makerid")
deljaehyu = request("deljaehyu")

dim odnshopitem
set odnshopitem = new CExtSiteItem
odnshopitem.FPageSize       = 10000
odnshopitem.FCurrPage       = 1
odnshopitem.FRectItemID     = itemid
odnshopitem.FRectItemName   = itemname
odnshopitem.FRectEventid    = eventid
odnshopitem.FRectMakerid    = makerid
odnshopitem.FDelJaeHyu		= deljaehyu
odnshopitem.GetDnshopRegedItemList

%>

<script language="javascript">
function deleteitem()
{
	if(confirm("다음 어드민에서 먼저 삭제하셨습니까?") == true) {
		if(confirm("총 <%=FormatNumber(odnshopitem.FTotalCount,0)%> 개 상품이 맞습니까?") == true) {
			frm.submit();
			return true;
	     } else {
	     	return false;
	     }
	} else {
		return false;
	}
}
</script>

<form name="frm" method="post">
<input type="hidden" name="action" value="del">
<table border="1">
<tr>
	<td>
		<textarea name="delitem" cols="20" rows="23"><%
			For i=0 To odnshopitem.FResultCount - 1

				If i <> 0 Then
					Response.Write vbCrLf
					delitemcode = delitemcode & ","
				End If
				Response.Write "B540_" & odnshopitem.FItemList(i).FItemID
				delitemcode = delitemcode & odnshopitem.FItemList(i).FItemID

			Next
		%></textarea>
	</td>
	<td valign="top">
		<input class="button" type="button" value="좌측 상품 삭제" onclick="deleteitem();">
	</td>
</tr>
</table>
<input type="hidden" name="delitemcode" value="<%=delitemcode%>">
</form>
<%
set odnshopitem = Nothing


Function Proc()
	Dim itemid
	itemid  = request("delitemcode")
	If itemid <> "" Then
		dbget.Execute "Delete [db_item].[dbo].tbl_dnshop_reg_item Where itemid IN(" & itemid & ")"
		Response.Write "<script>alert('삭제되었습니다.');opener.location.href='dnshopitem.asp?menupos=974';window.close();</script>"
	End If
End Function

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->