<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
Dim mallid, makerid, gubun, strSQL, adminText, idx, isReg
mallid		= request("mallid")
makerid		= request("makerid")
gubun		= request("gubun")
adminText	= request("adminText")
idx			= request("idx")
isReg		= False

If gubun = "I" Then
	strSQL = ""
	strSQL = ""
	strSQL = strSQL & " SELECT COUNT(*) as cnt FROM db_etcmall.dbo.tbl_jaehumall_hopeSell WHERE idx = '"&idx&"' and currstat = 2 "	
	rsget.Open strSQL,dbget,1
	If rsget("cnt") > 0 Then
		isReg = True
	End If
	rsget.Close

	strSQL = ""
	strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_jaehumall_hopeSell SET " & vbcrlf
	strSQL = strSQL & " currstat = 2 " & vbcrlf
	strSQL = strSQL & " ,adminText = '"&html2db(adminText)&"' " & vbcrlf
	strSQL = strSQL & " ,adminregdate = getdate() " & vbcrlf
	strSQL = strSQL & " WHERE idx= '"&idx&"' " & vbcrlf
	dbget.Execute strSQL

	If isReg = False Then
		strSQL = ""
		strSQL = strSQL & " INSERT INTO db_etcmall.dbo.tbl_jaehumall_hopeSell_Log (mallgubun, makerid, hopeStr, useYN, reguserid, regdate) " & vbcrlf
		strSQL = strSQL & " SELECT TOP 1 mallgubun, makerid, '[관리자] "&html2db(adminText)&"', hopesellstat, '"&session("ssBctID")&"', getdate() "
		strSQL = strSQL & " FROM db_etcmall.dbo.tbl_jaehumall_hopeSell "
		strSQL = strSQL & " WHERE idx= '"&idx&"' " & vbcrlf
		dbget.Execute strSQL
	End If
	response.write "<script language='javascript'>opener.location.reload();window.close();</script>"	
Else
	strSQL = ""
	strSQL = strSQL & " SELECT makerid, mallgubun FROM db_etcmall.dbo.tbl_jaehumall_hopeSell WHERE idx = '"&idx&"' "
	rsget.Open strSQL,dbget,1
	If not rsget.EOF Then
		mallid = rsget("mallgubun")
		makerid = rsget("makerid")
	End If
	rsget.Close
End If
%>
<script language='javascript'>
function frmsubmit(){
	var frm = document.frm;
	if(frm.adminText.value == ''){
		alert('사유를 입력하세요');
		frm.adminText.focus();
		return;
	}
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="POST" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="gubun" value="I">
<input type="hidden" name="idx" value="<%=idx%>">
<tr bgcolor="#FFFFFF" height="30">
	<td>
		[ 브랜드ID : <%= makerid %> ]<br><br>
		업체에게 전달할 코멘트 입력 후 저장 버튼<br>
		<input type="text" class="text" name="adminText" size="100">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30" align="center">
	<td>
		<input type="button" class="button" value="저장" onclick="frmsubmit();">&nbsp;&nbsp;
		<input type="button" class="button" value="취소" onclick="self.close();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
