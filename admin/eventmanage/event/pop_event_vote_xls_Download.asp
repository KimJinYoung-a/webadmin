<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  �̺�Ʈ �ڸ�Ʈ ������ Excel �ٿ�ε�
' History : 2007.10.12 ������ ����
'			2014.03.10 �ѿ�� ����
'			2016.03.02 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
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
	"	t1.sub_idx, t1.userid " &_
	"	, Case When left(t2.juminno,2) < 20 Then (year(getdate())-('20' + left(t2.juminno,2))+1) " &_
	"		When left(t2.juminno,2) >= 20 Then (year(getdate())-('19' + left(t2.juminno,2))+1) " &_
	"		end as userAge " &_
	"	, t3.userlevel " &_
	"	, t1.regdate, t2.regdate as joindate " &_
	"	, t1.sub_opt1, t1.sub_opt2, t1.sub_opt3  " &_
	"	, case t1.device  " &_
	"	When 'W' then 'pc��'  " &_
	"	When 'M' then '�������'  " &_
	"	When 'A' then '�ٹ����پ�'  " &_
	"	 End as sitegubun " &_
	" from [db_event].[dbo].[tbl_event_subscript] as t1 " &_
	"	Join db_user.[dbo].tbl_user_n as t2 " &_
	"		on t1.userid=t2.userid " &_
	"	Join db_user.[dbo].tbl_logindata as t3 " &_
	"		on t2.userid=t3.userid " &_
	" left join db_user.dbo.tbl_invalid_user iu" &_
	" 	on t1.userid=iu.invaliduserid" &_
	" 	and iu.isusing='Y'" &_
	" 	and iu.gubun='ONEVT'" &_			
	" where iu.idx is null and t1.evt_code=" & eCode &_
	"	and t1.regdate between '" & Sdate & "' and dateadd(d,1,'" & Edate & "') "

Select Case limitLevel
	Case "orange"
		strSql = strSql & "	and t3.userlevel not in ('5') "
	Case "yellow"
		strSql = strSql & "	and t3.userlevel not in ('0','5') "
	Case "white"
		strSql = strSql & "	and t3.userlevel not in ('0') "
end Select

rsget.Open strSql, dbget, 1

'���� ��½���
'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=event " & eCode & "_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td>��ȣ</td>
	<td>ȸ��ID</td>
	<td>����</td>
	<td>ȸ�����</td>
	<td>������</td>
	<td>ȸ��������</td>
	<td>���� �� �Է¶� 1</td>
	<td>���� �� �Է¶� 2</td>
	<td>���� �� �Է¶� 3</td>
	<td>����Ʈ����</td>
</tr>

<%
if Not(rsget.EOF or rsget.BOF) then
	do Until rsget.EOF
%>
<tr>
	<td><%=rsget("sub_idx")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("userid")%></td>
	<td><%=rsget("userAge")%></td>
	<td><%= getUserLevelStrByDate(rsget("userlevel"), left(rsget("regdate"),10)) %></td>
	<td><%=rsget("regdate")%></td>
	<td><%=rsget("joindate")%></td>
	<td style='mso-number-format:"\@"'><%=rsget("sub_opt1")%></td>
	<td><%=rsget("sub_opt2")%></td>
	<td><%=rsget("sub_opt3")%></td>
	<td><%=rsget("sitegubun")%></td>
</tr>
<%
	rsget.MoveNext
	loop
else
%>
	<tr>
		<td colspan="13" align="center">���ǿ� �´� �����ڰ� �����ϴ�</td>
	</tr>
<%	end if %>

</table>
</body>
</html>

<% rsget.close %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
