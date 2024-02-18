<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  �̺�Ʈ �ڸ�Ʈ ������ Excel �ٿ�ε�
' History : 2007.10.12 ������ ����
'           2014.03.03 ������ ; �������� ������ ����
'			2014.03.10 �ѿ�� ����
'			2014.08.13 ����ȭ ��ȸ�� �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
response.write "������� �Ŵ��Դϴ�. �ŸŴ��� ��� ��Ź�帳�ϴ�."
response.end

dim eCode, Sdate, Edate, limitLevel
dim strSql

eCode = Request("eC")	'�̺�Ʈ�ڵ�
Sdate = Request("Sdate")	'���������
Edate = Request("Edate")	'����������

	'// DB���� �������
	strSql = "select " &_
			"	t1.sub_idx " &_
			"	, t1.regdate " &_
			"	, t1.sub_opt1, t1.sub_opt2, t1.sub_opt3 " &_
			" from [db_event].[dbo].[tbl_event_subscript] as t1 " &_
			" left join db_user.dbo.tbl_invalid_user iu" &_
			" 	on t1.userid=iu.invaliduserid" &_
			" 	and iu.isusing='Y'" &_
			" 	and iu.gubun='ONEVT'" &_			
			" where iu.idx is null and t1.evt_code=" & eCode &_
			"	and t1.regdate between '" & Sdate & "' and dateadd(d,1,'" & Edate & "') "

		rsget.Open strSql, dbget, 1
%>
<%	'���� ��½���
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=event" & eCode & "_" & Date() & ".xls"
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td>��ȣ</td>
	<td>������</td>
	<td>���� �� �Է¶� 1</td>
	<td>���� �� �Է¶� 2</td>
	<td>���� �� �Է¶� 3</td>
</tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
<tr>
	<td><%=rsget("sub_idx")%></td>
	<td><%=rsget("regdate")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("sub_opt1")%></td>
	<td><%=rsget("sub_opt2")%></td>
	<td><%=rsget("sub_opt3")%></td>
</tr>
<%
		rsget.MoveNext
		loop
	else
%>
<tr><td colspan="13" align="center">���ǿ� �´� �����ڰ� �����ϴ�</td></tr>
<%	end if %>
</table>
</body>
</html>
<% rsget.close %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
