<%@Language="VBScript" CODEPAGE="65001" %>
<% option explicit %>
<%
Response.CharSet="utf-8"
Session.codepage="65001"
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"
%>
<%
'###########################################################
' Description : 제휴몰 API 주문입력
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin_UTF8.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/db/dbTPLHelper.asp" -->
<!-- #include virtual="/lib/util/aspJSON3.8.1.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/3pl/tempOrder/3PLSiteOrderLib.asp"-->
<%

dim mode, companyid, partnercompanyid

mode				= requestCheckVar(html2db(request("mode")),32)
companyid			= requestCheckVar(html2db(request("companyid")),32)
partnercompanyid	= requestCheckVar(html2db(request("partnercompanyid")),32)

'// for aspJSON 3.8.1
Response.LCID = 1042

select case partnercompanyid
	case "12"
		Call GetOrderFromExtSite(companyid, partnercompanyid)
		response.write "<script>alert('저장되었습니다.');</script>"
	case else
		response.write "undefined"
end select

%>
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<%
Session.codepage="949"
%>
