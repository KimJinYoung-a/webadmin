<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Board_xls_Download.asp
' Description :  �̺�Ʈ �Խ��� ������ Excel �ٿ�ε�
' History : 2009.05.06 ������ ����
'           2014.03.10 �ѿ�� ����
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
	oevent.fevent_bbs_notpaging()

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
	<td>����</td>
	<td>ȸ�����</td>
	<td>�ۼ���</td>
	<td>ȸ��������</td>
	<td>����</td>
	<td>����</td>
	<td>�����</td>
	<td>�̹���1</td>
	<td>�̹���2</td>
</tr>
<% if oevent.FresultCount>0 then %>
<% for intLoop=0 to oevent.FresultCount-1 %>
	<tr>
		<td><%= oevent.FItemList(intLoop).fevtbbs_idx %></td>
		<td><%= oevent.FItemList(intLoop).fuserid %></td>
		<td><%= oevent.FItemList(intLoop).fuserAge %></td>
		<td><%= getUserLevelStrByDate(oevent.FItemList(intLoop).fuserlevel, left(oevent.FItemList(intLoop).fevtbbs_regdate,10)) %></td>
		<td><%= oevent.FItemList(intLoop).fevtbbs_regdate %></td>
		<td><%= oevent.FItemList(intLoop).fjoindate %></td>
		<td><%= oevent.FItemList(intLoop).fevtbbs_subject %></td>
		<td><%= oevent.FItemList(intLoop).fevtbbs_content %></td>
		<td><% if Not(oevent.FItemList(intLoop).fevtbbs_icon="" or isNull(oevent.FItemList(intLoop).fevtbbs_icon)) then %><img src="<%= staticImgUrl & "/contents/photo_event/" & eCode & "/icon1/" & oevent.FItemList(intLoop).fevtbbs_icon %>"><% end if %></td>
		<td><% if Not(oevent.FItemList(intLoop).fevtbbs_img1="" or isNull(oevent.FItemList(intLoop).fevtbbs_img1)) then %><%= staticImgUrl & "/contents/photo_event/" & eCode & "/" & oevent.FItemList(intLoop).fevtbbs_img1 %><% end if %></td>
		<td><% if Not(oevent.FItemList(intLoop).fevtbbs_img2="" or isNull(oevent.FItemList(intLoop).fevtbbs_img2)) then %><%= staticImgUrl & "/contents/photo_event/" & eCode & "/" & oevent.FItemList(intLoop).fevtbbs_img2 %><% end if %></td>
	</tr>
<%
	intLoop = intLoop + 1
	if intLoop mod 1000 = 0 then
	end if
next
%>
<% else %>
	<tr><td colspan="13" align="center">���ǿ� �´� �����ڰ� �����ϴ�</td></tr>
<%	end if %>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
