<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ�ؿܹ������ �ϰ����� Excel ���ε�
' Hieditor : 2016.06.03 ������ ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	'/// �������¸� Ms-Excel�� ���� ///
	Response.ContentType = "application/vnd.ms-excel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition", "attachment;filename=�ٹ�����_�ؿܹ��_��ǰ_" & Date() & ".xls"
%>
<%
dim eCode, Sdate, Edate, limitLevel
dim strSql

	'// DB���� �������
	strSql = "select i.itemid, i.itemname, isNull(i.deliverOverseas,'N') as deliverOverseas,isNull(i.itemWeight,0) as itemWeight " &_
			"from db_item.dbo.tbl_item as i "&_
			"where i.makerid='" & session("ssBctId") & "' " &_ 
			"	and i.isusing='Y' "

		rsget.Open strSql, dbget, 1
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr style="background-color:#66FFFF;">
	<td colspan="4">�� ������ 4��° �ٺ��� �ۼ����ּ��� (�������� �ʼ� �Է�). �ۼ� �� �ٸ��̸����� ���� > "Excel 97-2003���� ����"�� ���� �� ���ε����ּ���.<font color="red"><B>���Դ� ���ڸ�</b></font> �Է����ּ���(����,�޸� �ԷºҰ���)</td>
</tr>
<tr style="background-color:#FFCCCC; display:none;">
	<td>code</td>
	<td>name</td>
	<td>sYn</td>
	<td>iW</td> 
</tr>
<tr style="background-color:#D8D8D8; color:#5A5A5A;">
	<td align="center">��ǰ�ڵ�</td>
	<td align="center">��ǰ��</td>
	<td align="center">�ؿܹ�ۿ���</td>
	<td align="center">��ǰ����(g)</td>
</tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
<tr>
	<td align="center" style="background-color:#FFFF00;"><%=rsget("itemid")%></td>
	<td align="center" style='mso-number-format:"\@"'><%=rsget("itemname")%></td>
	<td align="center" style='background-color:#FFFF00; mso-number-format:"\@"'><%=rsget("deliverOverseas")%></td> 
	<td align="center" style='background-color:#FFFF00; mso-number-format:"\@"'><%=rsget("itemWeight")%></td>
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
 rsget.close
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
