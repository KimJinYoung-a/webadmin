<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/shintvshopping/shintvshoppingCls.asp"-->
<%
Dim iisql, iniVal, mode
mode = request("mode")
iisql = ""
iisql = iisql & " SELECT TOP 1 isnull(iniVal, '') as iniVal "
iisql = iisql & " FROM db_etcmall.dbo.tbl_outmall_ini " & VbCRLF
iisql = iisql & " where mallid='shintvshopping' " & VbCRLF
iisql = iisql & " and inikey='pass'"
rsget.CursorLocation = adUseClient
rsget.Open iisql, dbget, adOpenForwardOnly, adLockReadOnly
if not rsget.Eof then
    iniVal	= rsget("iniVal")
end if
rsget.close

If mode = "I" Then
	Dim ipass
	ipass = Trim(request("pass"))
	iisql = ""
	iisql = iisql & " IF EXISTS (SELECT iniVal FROM db_etcmall.dbo.tbl_outmall_ini WHERE mallid='shintvshopping' and inikey='pass') "
	iisql = iisql & " 	BEGIN "
	iisql = iisql & "		UPDATE db_etcmall.dbo.tbl_outmall_ini "
	iisql = iisql & "		SET iniVal = '"& ipass &"' "
	iisql = iisql & "		WHERE mallid='shintvshopping'  "
	iisql = iisql & "		and inikey='pass' "
	iisql = iisql & " 	END "
	iisql = iisql & " ELSE "
	iisql = iisql & " 	BEGIN "
	iisql = iisql & "		INSERT INTO db_etcmall.dbo.tbl_outmall_ini (mallid, inikey, iniVal, lastupdate) VALUES "
	iisql = iisql & "		('shintvshopping', 'pass', '"& ipass &"', GETDATE()) "
	iisql = iisql & " 	END "
	dbget.execute iisql
	Response.Write "<script>parent.location.reload();</script>"
	Response.End
End If
%>
<script language="javascript">
function saveProcess(){
	if(confirm("신세계TV쇼핑 비밀번호와 같아야 합니다.\n동일하지 않을 시 모든 API에 오류가 발생합니다.\n\n저장하겠습니까?")){
		document.frm.target = "xLink";
		document.frm.mode.value = "I";
		document.frm.submit();
	}
}
</script>
<form name="frm" method="post" action="popPassword.asp" onSubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td with="50">업체코드</td>
	<td>419803</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td with="50">ID</td>
	<td>E419803</td>
</tr>
<% If (session("ssBctID")="kjy8517") or (session("ssBctID")="as2304") or (session("ssBctID")="sj100") or (session("ssBctID")="nys1006") Then %>
<tr height="25" bgcolor="FFFFFF">
	<td with="50">Password</td>
	<td>
		<input type="text" class="text" name="pass" value="<%= iniVal %>">
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF" align="center">
	<td colspan="2">
		<input type="button" class="button" value="저장" onclick="saveProcess();">
	</td>
</tr>
<% Else %>
<tr height="25" bgcolor="FFFFFF" align="center">
	<td colspan="2">
		변장혁, 백소정, 나예슬, 김진영 외에 수정 불가
	</td>
</tr>
<% End If %>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->