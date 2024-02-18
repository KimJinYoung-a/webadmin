<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/admin/etc/order/lib/xSiteOrderLib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<script language='javascript'>
// 선택된 상품 일괄 등록
function orderClick(sellsite, gubun) {
	if (confirm("주문 Step" + gubun + " 호출?")){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.sellsite.value = sellsite;
		document.frmSvArr.gubunCode.value = gubun;
		document.frmSvArr.action = "/admin/etc/order/xSiteOrder_Ins_Process.asp";
		document.frmSvArr.submit();
    }
}
</script>
<%
Dim sellsite
sellsite = request("sellsite")
%>
<form name="frmSvArr" method="GET" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="sellsite" value="">
<input type="hidden" name="gubunCode" value="">
</form>

<input type="button" value="Step1" onclick="orderClick('<%= sellsite %>', 1);">
<input type="button" value="Step2" onclick="orderClick('<%= sellsite %>', 2);">
<input type="button" value="Step3" onclick="orderClick('<%= sellsite %>', 3);">
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300">sdsdsd</iframe>
<!-- #include virtual="/lib/db/dbclose.asp" -->
