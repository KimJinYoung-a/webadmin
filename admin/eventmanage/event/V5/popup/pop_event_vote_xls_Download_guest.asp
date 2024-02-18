<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  �̺�Ʈ �ڸ�Ʈ ������ Excel �ٿ�ε�
' History : 2007.10.12 ������ ����
'			2014.03.10 �ѿ�� ����
'			2014.08.13 ����ȭ ��ȸ�� �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->

<%
dim eCode, Sdate, Edate, limitLevel, oevent
dim strSql
dim intLoop : intLoop = 0

eCode = Request("eC")	'�̺�Ʈ�ڵ�
Sdate = Request("Sdate")	'���������
Edate = Request("Edate")	'����������

set oevent = new ClsEventbbs
	oevent.frecteCode = eCode
	oevent.frectSdate = Sdate
	oevent.frectEdate = Edate
	oevent.fevent_subscriptguest_notpaging()

downPersonalInformation_rowcnt=oevent.ftotalcount

%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
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
<% if oevent.FresultCount>0 then %>
<% for intLoop=0 to oevent.FresultCount-1 %>
<tr>
	<td><%= oevent.FItemList(intLoop).fsub_idx %></td>
	<td><%= oevent.FItemList(intLoop).fregdate %></td>
	<td style='mso-number-format:"\@"'><%= oevent.FItemList(intLoop).fsub_opt1 %></td>
	<td><%= oevent.FItemList(intLoop).fsub_opt2 %></td>
	<td><%= oevent.FItemList(intLoop).fsub_opt3 %></td>
</tr>
<%
	intLoop = intLoop + 1
	if intLoop mod 1000 = 0 then
		Response.Flush		' ���۸��÷���
	end if
next
%>
<% else %>
<tr><td colspan="13" align="center">���ǿ� �´� �����ڰ� �����ϴ�</td></tr>
<% end if %>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
