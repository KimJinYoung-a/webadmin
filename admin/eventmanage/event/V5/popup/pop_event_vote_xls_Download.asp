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
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->

<%
dim eCode, Sdate, Edate, limitLevel, strSql, oevent
dim intLoop : intLoop = 0
	eCode = Request("eC")	'�̺�Ʈ�ڵ�
	Sdate = Request("Sdate")	'���������
	Edate = Request("Edate")	'����������
	limitLevel = Request("limitLevel")	'ȸ���������

set oevent = new ClsEventbbs
	oevent.frecteCode = eCode
	oevent.frectSdate = Sdate
	oevent.frectEdate = Edate
	oevent.frectlimitLevel = limitLevel
	oevent.fevent_subscript_notpaging()

downPersonalInformation_rowcnt=oevent.ftotalcount

%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
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

<% if oevent.FresultCount>0 then %>
<% for intLoop=0 to oevent.FresultCount-1 %>
<tr>
	<td><%= oevent.FItemList(intLoop).fsub_idx %></td>
	<td style='mso-number-format:"\@"'><%= oevent.FItemList(intLoop).fuserid %></td>
	<td><%= oevent.FItemList(intLoop).fuserAge %></td>
	<td><%= getUserLevelStrByDate(oevent.FItemList(intLoop).fuserlevel, left(oevent.FItemList(intLoop).fregdate,10)) %></td>
	<td><%= oevent.FItemList(intLoop).fregdate %></td>
	<td><%= oevent.FItemList(intLoop).fjoindate %></td>
	<td style='mso-number-format:"\@"'><%= oevent.FItemList(intLoop).fsub_opt1 %></td>
	<td><%= oevent.FItemList(intLoop).fsub_opt2 %></td>
	<td><%= oevent.FItemList(intLoop).fsub_opt3 %></td>
	<td><%= oevent.FItemList(intLoop).fsitegubun %></td>
</tr>
<%
	if (intLoop+1) mod 500 = 0 then
		Response.Flush		' ���۸��÷���
	end if
next
%>
<% else %>
	<tr>
		<td colspan="13" align="center">���ǿ� �´� �����ڰ� �����ϴ�</td>
	</tr>
<% end if %>

</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
