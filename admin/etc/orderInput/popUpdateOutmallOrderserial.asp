<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim mode, strSQL
Dim outmallOrderserial : outmallOrderserial = requestCheckvar(Trim(request("outMallorderSerial")),30)
mode = request("mode")

If mode = "U" Then
    strSQL = ""
	strSQL = strSQL & " UPDATE db_temp.dbo.tbl_XSite_TMporder "
	strSQL = strSQL & " SET outmallOrderserial = '"& outmallOrderserial &"_1' "
	strSQL = strSQL & " , ref_outmallOrderserial = '"& outmallOrderserial &"' "
	strSQL = strSQL & " WHERE outmallOrderserial = '"& outmallOrderserial &"' "
    strSQL = strSQL & " and sellsite = 'gseshop' "
    strSQL = strSQL & " and matchState = 'I' "
	dbget.Execute strSQL
	response.write "<script>alert('변경되었습니다');window.close();</script>"
	response.write "<script>opener.location.reload();</script>"
End If
%>

<script type="text/javascript">
function zipUpdate(){
	var frm;
	frm = document.frm;

	if(frm.outmallOrderserial.value==""){
		alert("제휴주문번호를 입력하세요");
		frm.outmallOrderserial.focus();
		return false;
	}

	if (confirm('주문번호를 입력하면 입력 안 된 주문번호에 _1이 붙습니다.\n\n변경하시겠습니까?')){
		frm.submit();
	}
}
</script>

<form name="frm" method="post" action="popUpdateOutmallOrderserial.asp">
<input type="hidden" name="mode" value="U">
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="50">
    <td align="center" bgcolor="#E8E8FF">변경할 제휴주문번호</td>
    <td bgcolor="#FFFFFF">
        <input type="text" name="outmallOrderserial" value="">
    </td>
</tr>

<tr align="center" height="25" bgcolor="#FFFFFF">
    <td colspan="2" align="center">
    <input type="button" value="변경" class="button" onClick="zipUpdate();">
    </td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->