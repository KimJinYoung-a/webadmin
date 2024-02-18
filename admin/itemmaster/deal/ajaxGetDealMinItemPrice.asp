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
	dim idx, oDealItem, arrList
	idx = requestCheckVar(Request("idx"),10)

	set oDealItem = new ClsDeal	
		oDealItem.FRectMasterIDX = idx	
 		arrList = oDealItem.fnGetDealItemMinPrice 		
	set oDealItem = Nothing

	IF isArray(arrList) THEN
		response.Write arrList(0,0) & "|" & arrList(1,0)
	end IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->