<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 배송요청사항 수정
' Hieditor : 2019-11-05 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim mode, strSQL, xDeliverymemo, deliverymemo
Dim xOutMallorderSerial : xOutMallorderSerial = requestCheckvar(request("outMallorderSerial"),30)
deliverymemo = request("deliverymemo")
mode = request("mode")

If mode = "U" Then
	strSQL = "UPDATE db_temp.dbo.tbl_XSite_TMporder SET "
	strSQL = strSQL & " deliverymemo = '"&deliverymemo&"' "
	strSQL = strSQL & " WHERE outMallorderSerial = '"&xOutMallorderSerial&"' "
	dbget.Execute strSQL
	response.write "<script>alert('변경되었습니다');opener.location.reload();window.close();</script>"
	response.end
End If

strSQL = "SELECT TOP 1 deliverymemo "
strSQL = strSQL & " FROM db_temp.dbo.tbl_XSite_TMporder "
strSQL = strSQL & " WHERE outMallorderSerial = '"&xOutMallorderSerial&"' "
rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
If Not(rsget.EOF or rsget.BOF) Then
	xDeliverymemo	= rsget("deliverymemo")
End If
rsget.Close
%>

<script type="text/javascript">

function memoUpdate(){
	var frm;
	frm = document.frm;

	if(frm.deliverymemo.value==""){
		alert("배송요청사항을 입력하세요");
		frm.deliverymemo.focus();
		return false;
	}

	if (confirm('배송요청사항을 변경하시겠습니까?')){
		frm.submit();
	}
}
</script>

<form name="frm" method="post" action="popMemoEdit.asp">
<input type="hidden" name="mode" value="U">
<input type="hidden" name="outMallorderSerial" value="<%= xOutMallorderSerial %>">
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="50">
    <td align="center" bgcolor="#E8E8FF">배송요청사항</td>
    <td bgcolor="#FFFFFF">
    	기존 배송요청사항 : <%= xDeliverymemo %><br>
    	변경 배송요청사항 : <input type="text" class="text" name="deliverymemo" size="80">
    </td>
</tr>
<tr align="center" height="25" bgcolor="#FFFFFF">
    <td colspan="2" align="center">
    <input type="button" value="변경" class="button" onClick="memoUpdate();">
    </td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->