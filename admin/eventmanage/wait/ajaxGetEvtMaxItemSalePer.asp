<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/eventPartnerWaitCls.asp"-->
<%

dim evtCode
evtCode =    requestCheckVar(Request("eC"),10)

if evtCode = "" then
		Call Alert_return ("유입경로에 문제가 생겼습니다.   ")
end if

dim cEvtItem, arrList, eSailDiv,maxno
	 
	eSailDiv = requestCheckVar(Request("saildiv"),1)
	
	set cEvtItem = new CEvent		
		cEvtItem.FevtCode = evtCode  
		If eSailDiv="S" Then	
		cEvtItem.FRectSailYn="Y"
		Else
		cEvtItem.FRectCouponYn="Y"
		End If
 		maxno = cEvtItem.fnGetItemtMaxSale 		
	set cEvtItem = Nothing
 if maxno ="" or isNull(maxno) then maxno = 0
 response.write Cint(maxno)&"%"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->