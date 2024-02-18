<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V2.asp"-->
<%
	dim eCode, cEvtItem, arrList
	eCode = requestCheckVar(Request.form("eC"),10)

	set cEvtItem = new ClsEvent	
		cEvtItem.FPSize = 1	
		cEvtItem.FECode = eCode	
        cEvtItem.FCPage = 1
		cEvtItem.FESSort = 6		'할인율순
		cEvtItem.FRectIsUsing = "Y"
		cEvtItem.FRectSellYN = "Y"

 		arrList = cEvtItem.fnGetEventItem 		
	set cEvtItem = Nothing

	IF isArray(arrList) THEN
		if arrList(18,0)="Y" then
			response.Write formatnumber(((arrList(7,0)-arrList(9,0))/arrList(7,0))*100,0) & "%"
		end IF
	end IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->