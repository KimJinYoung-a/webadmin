<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ ������������ �ϰ����� Excel �ٿ�ε� + ��ǰ���
' Hieditor : 2015.05.26 ������ ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	'/// �������¸� Ms-Excel�� ���� ///
	Response.ContentType = "application/vnd.ms-excel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition", "attachment;filename=�ٹ�����_��ǰ�����������_��ǰ_" & Date() & ".xls"
%>
<%
dim eCode, Sdate, Edate, limitLevel
dim strSql

	'// DB���� �������
	strSql = "select i.itemid, i.itemname, isNull(c.safetyYn,'N') as safetyYn, isnull(c.safetyDiv,0) as safetyDiv, c.safetyNum " &_
			"from db_item.dbo.tbl_item as i " &_
			"	join db_item.dbo.tbl_item_Contents as c " &_
			"		on i.itemid=c.itemid " &_
			"where i.makerid='" & session("ssBctId") & "' " &_
			"	and (isnull(c.safetyYn,'N')='N' " &_
			"		or (c.safetyYn='Y' and safetyDiv<>'10') " &_
			"	) " &_
			"	and i.isusing='Y' "

		rsget.Open strSql, dbget, 1
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr style="background-color:#66FFFF;">
	<td colspan="5">�� ������ 4��° �ٺ��� �ۼ����ּ��� (�������� �ʼ� �Է�). �ۼ� �� �ٸ��̸����� ���� > "Excel 97-2003���� ����"�� ���� �� ���ε����ּ���.</td>
</tr>
<tr style="background-color:#FFCCCC; display:none;">
	<td>code</td>
	<td>name</td>
	<td>sYn</td>
	<td>Div</td>
	<td>sNum</td>
</tr>
<tr style="background-color:#D8D8D8; color:#5A5A5A;">
	<td align="center">��ǰ�ڵ�</td>
	<td align="center">��ǰ��</td>
	<td align="center" colspan="2">�������� ��� ����</td>
	<td align="center">������������:KC��ũ ��ȣ</td>
</tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
<tr>
	<td align="center" style="background-color:#FFFF00;"><%=rsget("itemid")%></td>
	<td align="center" style='mso-number-format:"\@"'><%=rsget("itemname")%></td>
	<td align="center" style='background-color:#FFFF00; mso-number-format:"\@"'><%=rsget("safetyYn")%></td>
	<td align="center" style='mso-number-format:"\@"'><%=getSaftyDivName(rsget("safetyDiv"))%></td>
	<td align="center" style='background-color:#FFFF00; mso-number-format:"\@"'><%=rsget("safetyNum")%></td>
</tr>
<%
		rsget.MoveNext
		loop
	else
%>
<%	end if %>
</table>
</body>
</html>
<%
 function getSaftyDivName(sdiv)
 	Select Case cStr(sdiv)
 		Case "10"
 			getSaftyDivName = "10:������������(KC��ũ)"
 		Case "20"
 			getSaftyDivName = "20:�����ǰ ��������"
 		Case "30"
 			getSaftyDivName = "30:KPS �������� ǥ��"
 		Case "40"
 			getSaftyDivName = "40:KPS �������� Ȯ�� ǥ��"
 		Case "50"
 			getSaftyDivName = "50:KPS ��� ��ȣ���� ǥ��"
 		Case Else
 			getSaftyDivName = ""
 	End Select
 end Function
 rsget.close
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
