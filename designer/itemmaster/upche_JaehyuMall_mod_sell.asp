<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/outmallSellCls.asp"-->
<%
Dim vMakerID, oOutMall, page, i, gubun, strSQL
Dim cisextusing, allHopeInsert, currstat, whyhope, adminText, adminRegdate, idx
vMakerID	= session("ssBctID")
If page = "" Then page = 1
gubun = requestCheckVar(request("gubun"),20)

If gubun = "D" Then
	Dim mallgubun, hopeidx
	mallgubun = requestCheckVar(request("mallgubun"),50)
	hopeidx	= requestCheckVar(request("hopeidx"),20)
	strSQL = ""
	strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_jaehumall_hopeSell SET " & vbcrlf
	strSQL = strSQL & " isComplete = 'X' " & vbcrlf
	strSQL = strSQL & " WHERE idx = '"&hopeidx&"' " & vbcrlf
	dbget.Execute strSQL

	strSQL = ""
	strSQL = strSQL & " INSERT INTO db_etcmall.dbo.tbl_jaehumall_hopeSell_Log (mallgubun, makerid, hopeStr, useYN, reguserid, regdate) " & vbcrlf
	strSQL = strSQL & " VALUES ('"&mallgubun&"', '"&vMakerID&"', '���� ��û ���', 'X', '"&vMakerID&"', getdate()) " & vbcrlf
	dbget.Execute strSQL
	response.redirect("/designer/itemmaster/upche_JaehyuMall_mod_sell.asp?menupos="&requestCheckVar(request("menupos"),10)&"")
End If

SET oOutMall = new cOutmall
	Call oOutMall.fnGetIsExtusing(vMakerID, cisextusing, allHopeInsert, currstat, whyhope, adminText, adminRegdate, idx)
	oOutMall.FCurrPage			= page
	oOutMall.FPageSize			= 1000
	oOutMall.FRectMakerid		= vMakerID
	oOutMall.getOutmallList
%>
<script language='javascript'>
function popHopeSell(imallid, isellyn){
    var pwin = window.open("/designer/itemmaster/popHopeSell.asp?mallid="+imallid+"&sellyn="+isellyn,"popHopeSell","width=800,height=300,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popLogView(imallid, imakerid){
    var pwin = window.open("/designer/itemmaster/popHopeLog.asp?mallid="+imallid+"&makerid="+imakerid,"popHopeLog","width=850,height=700,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function cancelHope(imallid, ihopeidx){
	var frm = document.frm;
	frm.mallgubun.value = imallid;
	frm.hopeidx.value = ihopeidx;
	frm.submit();
}
</script>
<form name="frm" method="POST" action="<%=CurrURL()%>" style="margin:0px;">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="gubun" value="D">
	<input type="hidden" name="mallgubun">
	<input type="hidden" name="hopeidx">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr><td width="90%"></td></tr>
		<tr>
			<td>�귣��ID : <%= vMakerID %></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<br>
<!-- ################################################## ���޸� ���� ���� ################################################## -->
<strong>���޸�</strong>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">����</td>
	<td width="100"><font color="BLUE"><strong>�����Ǹż���</strong></font></td>
	<td width="100">�����</td>
	<td width="300">����������</td>
	<td width="130"><font color="RED">�����û</font></td>
	<td>����</td>
	<td width="100">����Ȯ��</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>���޻� ��ü</td>
	<td><strong><%= Chkiif(cisextusing="N", "�Ǹž���", "�Ǹ�") %></strong></td>
	<td width="400" colspan="2">���޻� ��ü [�Ǹž���] �ΰ�� �Ʒ� ���� ������ ������� �Ǹž��� </td>
	<td>
	<%
		Dim disableChk : disableChk = false
		If allHopeInsert = "Y" and currstat = 0 Then
			disableChk = true
		End If

		If cisextusing = "Y" Then
			response.write "<input type='button' value='�Ǹž���' "& Chkiif(disableChk=true,"disabled","") &" class='button' onclick=""popHopeSell('all', 'N');"">"
		Else
			response.write "<input type='button' value='�Ǹ�' "& Chkiif(disableChk=true,"disabled","") &" class='button' onclick=""popHopeSell('all', 'Y');"">"
		End If
	%>
	</td>
	<td>
	<%
		If allHopeInsert = "Y" Then
			Select Case currstat
				Case "1"	response.write "<font title='"& whyhope &"'>���δ��</font>"
							response.write "<br><span style='cursor:pointer;' onclick=""cancelHope('all', "&idx&");""><font color='gray'>[���� ��û ���]</font></span>"
				Case "2"	response.write "<font title='"& whyhope &"'>�ݷ�</font>"
							response.write "<br><font color='RED'>("&adminText&")</font>"
							response.write "<br>"&adminRegdate
			End Select
		End If
	%>
	</td>
	<td><input type='button' value='����' class='button' onclick="popLogView('all', '<%=vMakerID%>');"></td>
</tr>
<tr height="2" bgcolor="#FFFFFF" >
	<td colspan="7"></td>
</tr>
<% For i = 0 To oOutMall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= oOutMall.FItemList(i).FMallid %></td>
	<td><strong><%= Chkiif(oOutMall.FItemList(i).FIdx = "0", "�Ǹ�", "�Ǹž���") %></strong></td>
	<td><%= oOutMall.FItemList(i).FReguserid %></td>
	<td><%= oOutMall.FItemList(i).FRegdate %></td>
	<td>
	<%
		If cisextusing = "Y" Then
			If oOutMall.FItemList(i).FIdx = "0" Then
				response.write "<input type='button' value='�Ǹž���' "& Chkiif(allHopeInsert="Y","disabled","") &" class='button' onclick=""popHopeSell('"&oOutMall.FItemList(i).FMallid&"', 'N');"">"
			Else
				response.write "<input type='button' value='�Ǹ�' "& Chkiif(allHopeInsert="Y","disabled","") &" class='button' onclick=""popHopeSell('"&oOutMall.FItemList(i).FMallid&"', 'Y');"">"
			End If
		Else
			response.write "�̿�Ұ�"
		End If
	%>
	</td>
	<td>
	<%
		Select Case oOutMall.FItemList(i).FCurrstat
			Case "1"	response.write "<font title='"& oOutMall.FItemList(i).FWhyhope &"'>���δ��</font>"
						response.write "<br><span style='cursor:pointer;' onclick=cancelHope('"&oOutMall.FItemList(i).FMallid&"','"&oOutMall.FItemList(i).FHopeidx&"');><font color='gray'>[���� ��û ���]</font></span>"
			Case "2"	response.write "<font color='RED' title='"& oOutMall.FItemList(i).FWhyhope &"'>�ݷ�</font>"
						response.write "<br><font color='RED'>("&oOutMall.FItemList(i).FadminText&")</font>"
						response.write "<br>"&oOutMall.FItemList(i).FadminRegdate
		End Select
	%>
	</td>
	<td><input type='button' value='����' class='button' onclick="popLogView('<%= oOutMall.FItemList(i).FMallid %>', '<%= vMakerID %>');"></td>
</tr>
<% Next %>
</table>
<% SET oOutMall = nothing %>
<!-- ################################################### ���޸� ���� �� ###################################################	-->
<%
If False Then
	Dim oPotalsite
	SET oPotalsite = new cOutmall
		oPotalsite.FCurrPage		= page
		oPotalsite.FPageSize		= 100
		oPotalsite.FRectMakerid		= vMakerID
		oPotalsite.getPotalSiteList
%>
<br><br>
<strong>���л���Ʈ</strong>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">����</td>
	<td width="100"><font color="BLUE"><strong>�����Ǹż���</strong></font></td>
	<td width="100">�����</td>
	<td width="300">����������</td>
	<td width="130"><font color="RED">�����û</font></td>
	<td>����</td>
	<td width="100">����Ȯ��</td>
</tr>
<% For i = 0 To oPotalsite.FResultCount - 1 %>
<%		If oPotalsite.FItemList(i).FMallID <> "shodocep" Then %>
<tr align="center" bgcolor="#FFFFFF">
	<td>
	<% 
		Select Case oPotalsite.FItemList(i).FMallID 
			Case "naverep" response.write "���̹�"
			Case "daumep" response.write "����"
			Case "shodocep" response.write "���"
		End Select
	%>
	</td>
	<td><strong><%= Chkiif(oPotalsite.FItemList(i).FIsusing = "Y", "�Ǹ�", "�Ǹž���") %></strong></td>
	<td><%= oPotalsite.FItemList(i).FReguserid %></td>
	<td><%= oPotalsite.FItemList(i).FLastupdate %></td>
	<td>
	<%
		If oPotalsite.FItemList(i).FIsusing = "Y" Then
			response.write "<input type='button' value='�Ǹž���' class='button' onclick=""popHopeSell('"&oPotalsite.FItemList(i).FMallid&"', 'N');""> "
		Else
			response.write "<input type='button' value='�Ǹ�' class='button' onclick=""popHopeSell('"&oPotalsite.FItemList(i).FMallid&"', 'Y');""> "
		End If
	%>
	</td>
	<td>
	<%
		Select Case oPotalsite.FItemList(i).FCurrstat
			Case "1"	response.write "<font title='"& oPotalsite.FItemList(i).FWhyhope &"'>���δ��</font>"
						response.write "<br><span style='cursor:pointer;' onclick=cancelHope('"&oPotalsite.FItemList(i).FMallid&"','"&oPotalsite.FItemList(i).FHopeidx&"');><font color='gray'>[���� ��û ���]</font></span>"
			Case "2"	response.write "<font color='RED' title='"& oPotalsite.FItemList(i).FWhyhope &"'>�ݷ�</font>"
						response.write "<br><font color='RED'>("&oPotalsite.FItemList(i).FadminText&")</font>"
						response.write "<br>"&oPotalsite.FItemList(i).FadminRegdate
		End Select
	%>
	</td>
	<td><input type='button' value='����' class='button' onclick="popLogView('<%= oPotalsite.FItemList(i).FMallid %>', '<%= vMakerID %>');"></td>
</tr>
<%		End If %>
<% Next %>
</table>
<% SET oPotalsite = nothing %>
<% End If %>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->