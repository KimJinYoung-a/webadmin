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
	dim idx, oDealItem, arrList, arrList2, SalePer
	idx = requestCheckVar(Request("idx"),10)

	Set oDealItem = New ClsDeal	
		oDealItem.FRectMasterIDX = idx	
 		arrList = oDealItem.fnGetMAXDealSalePer
		arrList2 = oDealItem.fnGetMAXDealCouponSalePer 
	Set oDealItem = Nothing

	If isArray(arrList) Then
		If arrList(2,0)="Y" Then
			SalePer = Cint(((arrList(0,0)-arrList(1,0))/arrList(0,0))*100) & "|" & arrList(4,0)
			If SalePer>"0" Then
				response.Write SalePer
			End If
		End If
	End If

	If isArray(arrList2) And SalePer<"1" Then
		response.Write Cint(arrList2(0,0)) & "|" & arrList2(1,0)
	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->