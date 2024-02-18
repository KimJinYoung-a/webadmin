<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/giftCls.asp"-->

<%
	Dim vQuery, vIdx, vGubun, vItemID, vTotSellcash, vSellcash, vDiliItemcost, vUseYN, vDel
	vDel			= Request("del")
	vIdx 			= Request("idx")
	vGubun 			= Request("gubun")
	vItemID 		= Request("itemid")
	vTotSellcash 	= Request("tot_sellcash")
	vSellcash 		= Request("sellcash")
	vDiliItemcost 	= Request("dili_itemcost")
	vUseYN 			= Request("useyn")
	
	
	If vIdx = "" Then
		vQuery = "SELECT COUNT(idx) FROM [db_order].[dbo].[tbl_mobile_gift_item] WHERE itemid = '" & vItemID & "' AND gubun = '" & vGubun & "'"
		rsget.Open vQuery, dbget ,1

		If rsget(0) > 0 Then
			Response.Write "<script language='javascript'>alert('등록된 상품입니다.');window.close();</script>"
			dbget.close()
			Response.End
		End If
		rsget.close()

		vQuery = "INSERT INTO [db_order].[dbo].[tbl_mobile_gift_item](gubun, itemid, tot_sellcash, sellcash, dili_itemcost, useyn) " & _
				 "VALUES('" & vGubun & "', '" & vItemID & "', '" & vTotSellcash & "', '" & vSellcash & "', '" & vDiliItemcost & "', '" & vUseYN & "')"
		dbget.execute vQuery
	Else
		If vDel = "o" Then
			vQuery = "DELETE [db_order].[dbo].[tbl_mobile_gift_item] WHERE idx = '" & vIdx & "'"
			dbget.execute vQuery
		Else
			vQuery = "UPDATE [db_order].[dbo].[tbl_mobile_gift_item] SET " & _
					 "		tot_sellcash = '" & vTotSellcash & "', sellcash = '" & vSellcash & "', dili_itemcost = '" & vDiliItemcost & "', useyn = '" & vUseYN & "' " & _
					 "WHERE idx = '" & vIdx & "'"
			dbget.execute vQuery
		End IF
	End IF
%>

<script language='javascript'>
alert('저장되었습니다.');
opener.document.location.reload();
window.close();
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->