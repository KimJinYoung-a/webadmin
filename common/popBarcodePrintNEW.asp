<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품 바코드 출력
' Hieditor : 2016-11-02,  skyer9 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim masteridx, orderType

masteridx = requestCheckVar(request("masteridx"),32)
orderType = requestCheckVar(request("orderType"),32)

response.write masteridx

'// http://webadmin.10x10.co.kr/common/popBarcodePrintNEW.asp?masteridx=401237&ordertype=offlineorder

'// (C_IS_SHOP = true)


'// aaaa

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
