<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  �̺�Ʈ �ڸ�Ʈ ������ Excel �ٿ�ε�
' History : 2007.10.12 ������ ����
'           2014.03.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
response.write "������� �Ŵ��Դϴ�. �ŸŴ��� ��� ��Ź�帳�ϴ�."
response.end

dim eCode, Sdate, Edate, limitLevel, strSql
	eCode = Request("eC")	'�̺�Ʈ�ڵ�
	Sdate = Request("Sdate")	'���������
	Edate = Request("Edate")	'����������
	limitLevel = Request("limitLevel")	'ȸ���������

'// DB���� �������
strSql = "select " &_
		"	t1.evtcom_idx, t1.userid " &_
		"	, Case When left(t2.juminno,2) < 20 Then (year(getdate())-('20' + left(t2.juminno,2))+1) " &_
		"		When left(t2.juminno,2) >= 20 Then (year(getdate())-('19' + left(t2.juminno,2))+1) " &_
		"		end as userAge " &_
		"	, t3.userlevel " &_
		"	, t1.evtcom_regdate, t2.regdate as joindate " &_
		"	, t1.evtcom_txt, t1.evtcom_point, t1.blogurl " &_
		"	,(select count(*) FROM [db_event].[dbo].[tbl_event_prize] as p WHERE p.evt_winner = t2.userid) as wincnt  " &_
		"	,(select top 1 evt_regdate FROM [db_event].[dbo].[tbl_event_prize] as p WHERE p.evt_winner = t2.userid order by evt_regdate desc) as windate " &_
		"	,t2.username " &_
		" from db_event.dbo.tbl_event_comment as t1 " &_
		"	Join db_user.[dbo].tbl_user_n as t2 " &_
		"		on t1.userid=t2.userid " &_
		"	Join db_user.[dbo].tbl_logindata as t3 " &_
		"		on t2.userid=t3.userid " &_
		" left join db_user.dbo.tbl_invalid_user iu" &_
		" 	on t1.userid=iu.invaliduserid" &_
		" 	and iu.isusing='Y'" &_
		" 	and iu.gubun='ONEVT'" &_
		" where iu.idx is null and t1.evt_code=" & eCode &_
		"	and t1.evtcom_using='Y' " &_
		"	and t1.evtcom_regdate between '" & Sdate & "' and dateadd(d,1,'" & Edate & "') "

	Select Case limitLevel
		Case "orange"
			strSql = strSql & "	and t3.userlevel not in ('5') "
		Case "yellow"
			strSql = strSql & "	and t3.userlevel not in ('0','5') "
		Case "white"
			strSql = strSql & "	and t3.userlevel not in ('0') "
	end Select
'	response.write strsql
'	response.end
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
	<td>ȸ��ID</td>
	<td>�̸�</td>
	<td>����</td>
	<td>ȸ�����</td>
	<td>�ۼ���</td>
	<td>ȸ��������</td>
	<td>�ڸ�Ʈ ����</td>
	<td>���ù�ȣ</td>
	<td>��α��ּ�</td>
	<td>�ֱٴ�÷��</td>
	<td>�̺�Ʈ��÷Ƚ��</td>
</tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
<tr>
	<td><%=rsget("evtcom_idx")%></td>
	<td><%=rsget("userid")%></td>
	<td><%=rsget("username")%></td>
	<td><%=rsget("userAge")%></td>
	<td><%= getUserLevelStrByDate(rsget("userlevel"), left(rsget("evtcom_regdate"),10)) %></td>
	<td><%=rsget("evtcom_regdate")%></td>
	<td><%=rsget("joindate")%></td>
	<td><%=rsget("evtcom_txt")%></td>
	<td><%=rsget("evtcom_point")%></td>
	<td><%=rsget("blogurl")%></td>
	<td><%=rsget("windate")%></td>
	<td><%=rsget("wincnt")%></td>
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
