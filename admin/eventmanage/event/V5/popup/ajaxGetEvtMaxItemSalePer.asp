<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
	dim eCode, cEvtItem, arrList, eSailDiv
	eCode = requestCheckVar(Request("eC"),10)
	eSailDiv = requestCheckVar(Request("saildiv"),1)

	set cEvtItem = new ClsEvent	
		cEvtItem.FPSize = 1	
		cEvtItem.FECode = eCode	
        cEvtItem.FCPage = 1
		cEvtItem.FESSort = 6		'할인율순
		cEvtItem.FRectIsUsing = "Y"
		cEvtItem.FRectSellYN = "Y"
		If eSailDiv="S" Then
		cEvtItem.FRectSailYn="Y"
		arrList = cEvtItem.fnGetEventItem 
		Else
		cEvtItem.FRectCouponYn="Y"
		arrList = cEvtItem.fnGetEventItemCouponMax 
		End If
	set cEvtItem = Nothing

	
	If eSailDiv="S" Then
		IF isArray(arrList) THEN
			if arrList(18,0)="Y" then
				response.Write formatnumber(((arrList(7,0)-arrList(9,0))/arrList(7,0))*100,0)
			end IF
		end IF
	Else
			response.Write formatnumber(arrList,0)
	end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->