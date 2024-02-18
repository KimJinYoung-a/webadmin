<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "euc-kr"
%>
<%
'####################################################
' Description : 
' History : 최초생성자모름
'			2017.04.10 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/partner/lib/function/partnerItemFunction.asp" --> 
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
  
	if request("isWt")="W" then
		'등록 대기 상품
		response.Write getDispCategoryWait2016(requestCheckVar(request("itemid"),10)) 
	else
		'실등록 상품
		response.Write getDispCategory2016(requestCheckVar(request("itemid"),10))
	end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->