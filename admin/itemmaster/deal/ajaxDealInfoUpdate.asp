<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<%
	dim sqlStr, mode, itemid

	mode = request("mode")
	itemid = request("itemid")

	if mode = "all" then
		sqlStr = "EXEC [db_event].[dbo].[usp_WWW_Item_DealItemSaleUsingStateInfo_Upd]"
		dbget.execute sqlStr
	else
		if itemid <> "" then
			sqlStr = "EXEC [db_event].[dbo].[usp_WWW_Item_Deal_OneItemInfo_Upd] " & itemid
			dbget.execute sqlStr
		end if
	end if

	response.Write "OK"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->