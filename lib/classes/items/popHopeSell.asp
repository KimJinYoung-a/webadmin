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

'####### 이미 변경 사유가 등록되었는 지 검사 #######
isRegCnt = fnIsRegedHopeCnt(vMallid, vMakerID)	
If isRegCnt > 0 Then 
	response.write "<script language='javascript'>alert('이미 변경요청 하셨습니다');window.close();</script>" 
	response.end
End If
'###### 이미 변경 사유가 등록되었는 지 검사 끝 #####

If gubun = "I" Then

	If InStr(whyhope, "[관리자]") > 0 Then
		response.write "<script language='javascript'>alert('문자열에 [관리자]를 입력할 수 없습니다.');document.location.replace('/designer/itemmaster/popHopeSell.asp?mallid="&mallgubun&"&sellyn="&hopeSellstat&"');</script>" 
		response.end
	End If
	
	If Len(whyhope) < 10 Then
		response.write "<script language='javascript'>alert('문자열은 영문기준 10자 이상 입력하셔야 합니다');document.location.replace('/designer/itemmaster/popHopeSell.asp?mallid="&mallgubun&"&sellyn="&hopeSellstat&"');</script>" 
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
if (confirm('제휴사 전체를 변경요청하셨습니다.\n이하 제휴몰의 승인건들은 무시됩니다.\n진행하시겠습니까?')){

}else{
	self.close();
}
<% End If %>
function frmsubmit(){
	var frm = document.frm;
	if(frm.whyhope.value == ''){
		alert('사유를 입력하세요');
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
			<td><strong>판매설정 변경 <font color="RED">* 기간설정 불가</font></strong></td>
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
	<td width="70">구분</td>
	<td width="70">변경요청</td>
	<td>사유 <strong>(최소 10byte이상 입력하세요)</strong></td>
	<td width="70">진행상태</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>
	<%
		SELECT Case vMallid
			Case "all"			response.write "제휴사 전체"
			Case "daumep" 		response.write "다음"
			Case "naverep" 		response.write "네이버"
			Case Else			response.write vMallid
		End Select
	%>
	</td>
	<td><%= Chkiif(hopeSell="Y", "판매", "판매안함") %></td>
	<td><input type="text" name="whyhope" size="80" class="text"></td>
	<td><input type="button" class="button" value="저장" onclick="frmsubmit();"></td>
</tr>
</form>
</table>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
