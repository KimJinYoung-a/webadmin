<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_cashreceiptcls.asp"-->

<%
Dim orderSerial	: orderSerial	= req("orderSerial","")

Dim oreceipt
set oreceipt = new CCashReceipt
oreceipt.FRectorderSerial = orderSerial
oreceipt.GetReceiptByOrderSerial


%>
<html>
<head>
<title>현금영수증</title>
<script>
function jsCancel()
{
	var f = document.frmWrite;
	if (f.chkPrint.value)
	{
		if (confirm("현금영수증을 취소하시겠습니까?"))
		{
			f.submit();
		}
	}
}
</script>
</head>
<body style="margin:0;">
<iframe src="https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?noTid=<%=oreceipt.FOneItem.Ftid%>&noMethod=1" style="width:400px; height:340px; border:1px solid #000000; " frameborder="1"></iframe>

<form name="frmWrite" method="post" action="/cscenter/taxSheet/receipt_process.asp">
<input type="hidden" name="chkPrint" value="<%=oreceipt.FOneItem.Fidx%>">
<input type="hidden" name="Atype" value="C2">
<div align="center">
	<input type="button" value="현금영수증취소" onclick="jsCancel();">
</div>
</form>
</body>
</html>

<%
set oreceipt = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->