<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/outmallSellCls.asp"-->
<%
Dim vMallid, hopeSell,  gubun, strSQL, isAllRegedHope, isRegCnt, vMakerID
Dim hopeSellstat, whyhope, mallgubun
vMallid			= requestCheckvar(Request("mallid"),16)
hopeSell		= requestCheckvar(Request("sellyn"),1)
gubun			= requestCheckvar(Request("gubun"),1)
vMakerID		= session("ssBctID")

hopeSellstat	= requestCheckvar(Request("hopeSellstat"),1)
mallgubun		= requestCheckvar(Request("mallgubun"),16)
whyhope			= request("whyhope")

'####### �̹� ���� ������ ��ϵǾ��� �� �˻� #######
isRegCnt = fnIsRegedHopeCnt(vMallid, vMakerID)	
If isRegCnt > 0 Then 
	response.write "<script language='javascript'>alert('�̹� �����û �ϼ̽��ϴ�');window.close();</script>" 
	response.end
End If
'###### �̹� ���� ������ ��ϵǾ��� �� �˻� �� #####

If gubun = "I" Then

	If InStr(whyhope, "[������]") > 0 Then
		response.write "<script language='javascript'>alert('���ڿ��� [������]�� �Է��� �� �����ϴ�.');document.location.replace('/designer/itemmaster/popHopeSell.asp?mallid="&mallgubun&"&sellyn="&hopeSellstat&"');</script>" 
		response.end
	End If
	
	If Len(whyhope) < 10 Then
		response.write "<script language='javascript'>alert('���ڿ��� �������� 10�� �̻� �Է��ϼž� �մϴ�');document.location.replace('/designer/itemmaster/popHopeSell.asp?mallid="&mallgubun&"&sellyn="&hopeSellstat&"');</script>" 
		response.end
	End If

	If mallgubun = "all" Then
		strSQL = ""
		strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_jaehumall_hopeSell SET " & vbcrlf
		strSQL = strSQL & " isComplete = 'X' " & vbcrlf
		strSQL = strSQL & " WHERE makerid = '"&vMakerID&"' and mallgubun <> 'all' and mallgubun <> 'daumep' and mallgubun <> 'naverep'  " & vbcrlf
		dbget.Execute strSQL
	End If
	strSQL = ""
	strSQL = strSQL & " IF EXISTS(SELECT TOP 1 * FROM db_etcmall.dbo.tbl_jaehumall_hopeSell WHERE makerid='"&vMakerID&"' and mallgubun='"&mallgubun&"' and currstat=2 and iscomplete <> 'X' )" & vbcrlf
	strSQL = strSQL & " 	BEGIN " & vbcrlf
	strSQL = strSQL & " 		UPDATE db_etcmall.dbo.tbl_jaehumall_hopeSell SET " & vbcrlf
	strSQL = strSQL & " 		whyhope = '"&html2db(whyhope)&"' " & vbcrlf
	strSQL = strSQL & " 		,currstat=1 " & vbcrlf
	strSQL = strSQL & " 		,hoperegdate = getdate() " & vbcrlf
	strSQL = strSQL & " 		WHERE makerid='"&vMakerID&"' and mallgubun='"&mallgubun&"' and currstat=2  " & vbcrlf
	strSQL = strSQL & " 	END " & vbcrlf
	strSQL = strSQL & " ELSE " & vbcrlf
	strSQL = strSQL & " 	BEGIN " & vbcrlf
	strSQL = strSQL & " 		INSERT INTO db_etcmall.dbo.tbl_jaehumall_hopeSell (makerid, mallgubun, currstat, hopesellstat, whyhope, hoperegdate, isComplete) " & vbcrlf
	strSQL = strSQL & " 		VALUES ('"&vMakerID&"', '"&mallgubun&"', '1', '"&hopeSellstat&"', '"&html2db(whyhope)&"', getdate(), 'N') " & vbcrlf
	strSQL = strSQL & " 	END " & vbcrlf
	dbget.Execute strSQL

	strSQL = ""
	strSQL = strSQL & " INSERT INTO db_etcmall.dbo.tbl_jaehumall_hopeSell_Log (mallgubun, makerid, hopeStr, useYN, reguserid, regdate) " & vbcrlf
	strSQL = strSQL & " VALUES ('"&mallgubun&"', '"&vMakerID&"', '"&whyhope&"', '"&hopeSellstat&"', '"&vMakerID&"', getdate()) " & vbcrlf
	dbget.Execute strSQL
	response.write "<script language='javascript'>opener.location.reload();window.close();</script>"
Else
	If vMallid = "all" Then
		isAllRegedHope = fnHoperegConfirm(vMakerID)
	End If
End If
%>
<script language='javascript'>
<% If isAllRegedHope Then %>
if (confirm('���޻� ��ü�� �����û�ϼ̽��ϴ�.\n���� ���޸��� ���ΰǵ��� ���õ˴ϴ�.\n�����Ͻðڽ��ϱ�?')){

}else{
	self.close();
}
<% End If %>
function frmsubmit(){
	var frm = document.frm;
	if(frm.whyhope.value == ''){
		alert('������ �Է��ϼ���');
		frm.whyhope.focus();
		return;
	}
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr><td width="90%"></td></tr>
		<tr>
			<td><strong>�Ǹż��� ���� <font color="RED">* �Ⱓ���� �Ұ�</font></strong></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="POST" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="mallgubun" value="<%=vMallid%>">
<input type="hidden" name="hopeSellstat" value="<%=hopeSell%>">
<input type="hidden" name="gubun" value="I">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="70">����</td>
	<td width="70">�����û</td>
	<td>���� <strong>(�ּ� 10byte�̻� �Է��ϼ���)</strong></td>
	<td width="70">�������</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>
	<%
		SELECT Case vMallid
			Case "all"			response.write "���޻� ��ü"
			Case "daumep" 		response.write "����"
			Case "naverep" 		response.write "���̹�"
			Case Else			response.write vMallid
		End Select
	%>
	</td>
	<td><%= Chkiif(hopeSell="Y", "�Ǹ�", "�Ǹž���") %></td>
	<td><input type="text" name="whyhope" size="80" class="text"></td>
	<td><input type="button" class="button" value="����" onclick="frmsubmit();"></td>
</tr>
</form>
</table>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
