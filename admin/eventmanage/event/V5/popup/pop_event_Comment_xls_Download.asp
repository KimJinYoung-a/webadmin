<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls_Download.asp
' Description :  �̺�Ʈ �ڸ�Ʈ ������ Excel �ٿ�ε�
' History : 2007.10.12 ������ ����
'           2014.03.10 �ѿ�� ����
'			2019.11.14 ������ ���� (�޴��� ���� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
dim eCode, Sdate, Edate, limitLevel, strSql, oevent, intLoop
	eCode = Request("eC")	'�̺�Ʈ�ڵ�
	Sdate = Request("Sdate")	'���������
	Edate = Request("Edate")	'����������
	limitLevel = Request("limitLevel")	'ȸ���������

set oevent = new ClsEventbbs
	oevent.frecteCode = eCode
	oevent.frectSdate = Sdate
	oevent.frectEdate = Edate
	oevent.frectlimitLevel = limitLevel
	oevent.fevent_comment_notpaging()

downPersonalInformation_rowcnt=oevent.ftotalcount

%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
'���� ��½���
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
	<td>�̸���</td>
	<td>�޴���</td>
	<td>�ۼ���</td>
	<td>ȸ��������</td>
	<td>�ڸ�Ʈ ����</td>
	<td>���ù�ȣ</td>
	<td>��α��ּ�</td>
	<td>�ֱٴ�÷��</td>
	<td>�̺�Ʈ��÷Ƚ��</td>
</tr>
<% if oevent.FresultCount>0 then %>
<% for intLoop=0 to oevent.FresultCount-1 %>
<tr>
	<td><%= oevent.FItemList(intLoop).fevtcom_idx %></td>
	<td><%= oevent.FItemList(intLoop).fuserid %></td>
	<td><%= oevent.FItemList(intLoop).fusername %></td>
	<td><%= oevent.FItemList(intLoop).fuserAge %></td>
	<td><%= getUserLevelStrByDate(oevent.FItemList(intLoop).fuserlevel, left(oevent.FItemList(intLoop).fevtcom_regdate,10)) %></td>
	<td><%= oevent.FItemList(intLoop).fusermail %></td>
	<td><%= oevent.FItemList(intLoop).fusercell %></td>
	<td><%= oevent.FItemList(intLoop).fevtcom_regdate %></td>
	<td><%= oevent.FItemList(intLoop).fjoindate %></td>
	<td><%= oevent.FItemList(intLoop).fevtcom_txt %></td>
	<td><%= oevent.FItemList(intLoop).fevtcom_point %></td>
	<td><%= oevent.FItemList(intLoop).fblogurl %></td>
	<td><%= oevent.FItemList(intLoop).fwindate %></td>
	<td><%= oevent.FItemList(intLoop).fwincnt %></td>
</tr>
<%
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