<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 주문제작문구 수정
' Hieditor : 2016.04.15 김진영 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incsessionadmin.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim orderseq, mode, insReqDetail
orderseq		= requestcheckvar(request("outmallorderseq"),10)
mode			= requestcheckvar(request("mode"),1)
insReqDetail	= Trim(db2html(request("reqdetail")))

Dim sqlStr, reqDetail
sqlStr = ""
sqlStr = sqlStr & " SELECT TOP 1 requireDetail "
sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_TMPOrder "
sqlStr = sqlStr & " WHERE outmallorderseq = '"&orderseq&"' "
rsget.Open sqlStr,dbget,1
If not rsget.EOF  then
	reqDetail = rsget("requireDetail")
Else
	reqDetail = ""
End If
rsget.close

If mode = "I" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_temp.dbo.tbl_xSite_TMPOrder SET " & vbcrlf
	sqlStr = sqlStr & " requireDetail = '"&html2db(insReqDetail)&"' " & vbcrlf
	sqlStr = sqlStr & " ,requireDetail11stYN = 'Y' " & vbcrlf
	sqlStr = sqlStr & " WHERE outmallorderseq = '"&orderseq&"' "
	dbget.Execute sqlStr
	
	response.write "<script>alert('저장되었습니다.');opener.location.reload();window.close();</script>"
End If
%>
<script>
function frmCheck(){
	var frm;
	frm = document.frm;
	if(frm.reqdetail.value==''){
		alert('주문제작문구를 입력하세요');
		frm.reqdetail.focus();
		return;
	}
	frm.mode.value = 'I'
	frm.submit();
}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post" action="/admin/etc/orderinput/xSiteReqDetailedit.asp">
<input type="hidden" name="mode">
<input type="hidden" name="outmallorderseq" value="<%=orderseq%>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="text" size="70" value="<%=reqDetail%>" name="reqdetail"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="button" value="저장" class="button" onclick="frmCheck();"></td>
</tr>
</form>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->