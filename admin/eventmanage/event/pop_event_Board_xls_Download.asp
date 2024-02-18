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
		"	t1.evtbbs_idx, t1.userid " &_
		"	, Case When left(t2.juminno,2) < 20 Then (year(getdate())-('20' + left(t2.juminno,2))+1) " &_
		"		When left(t2.juminno,2) >= 20 Then (year(getdate())-('19' + left(t2.juminno,2))+1) " &_
		"		end as userAge " &_
		"	, t3.userlevel " &_
		"	, t1.evtbbs_regdate, t2.regdate as joindate " &_
		"	, t1.evtbbs_subject, t1.evtbbs_content, t1.evtbbs_img1, t1.evtbbs_img2, t1.evtbbs_icon " &_
		" from db_event.dbo.tbl_event_bbs as t1 " &_
		"	Join db_user.[dbo].tbl_user_n as t2 " &_
		"		on t1.userid=t2.userid " &_
		"	Join db_user.[dbo].tbl_logindata as t3 " &_
		"		on t2.userid=t3.userid " &_
		" left join db_user.dbo.tbl_invalid_user iu" &_
		" 	on t1.userid=iu.invaliduserid" &_
		" 	and iu.isusing='Y'" &_
		" 	and iu.gubun='ONEVT'" &_			
		" where iu.idx is null and t1.evt_code=" & eCode &_
		"	and t1.evtbbs_using='Y' " &_
		"	and t1.evtbbs_regdate between '" & Sdate & "' and dateadd(d,1,'" & Edate & "') "
	
	Select Case limitLevel
		Case "orange"
			strSql = strSql & "	and t3.userlevel not in ('5') "
		Case "yellow"
			strSql = strSql & "	and t3.userlevel not in ('0','5') "
		Case "white"
			strSql = strSql & "	and t3.userlevel not in ('0') "
	end Select

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
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
	<tr>
		<td><%=rsget("evtbbs_idx")%></td>
		<td><%=rsget("userid")%></td>
		<td><%=rsget("userAge")%></td>
		<td><%= getUserLevelStrByDate(rsget("userlevel"), left(rsget("evtbbs_regdate"),10)) %></td>
		<td><%=rsget("evtbbs_regdate")%></td>
		<td><%=rsget("joindate")%></td>
		<td><%=rsget("evtbbs_subject")%></td>
		<td><%=rsget("evtbbs_content")%></td>
		<td><% if Not(rsget("evtbbs_icon")="" or isNull(rsget("evtbbs_icon"))) then %><img src="<%= staticImgUrl & "/contents/photo_event/" & eCode & "/icon1/" & rsget("evtbbs_icon")%>"><% end if %></td>
		<td><% if Not(rsget("evtbbs_img1")="" or isNull(rsget("evtbbs_img1"))) then %><%= staticImgUrl & "/contents/photo_event/" & eCode & "/" & rsget("evtbbs_img1")%><% end if %></td>
		<td><% if Not(rsget("evtbbs_img2")="" or isNull(rsget("evtbbs_img2"))) then %><%= staticImgUrl & "/contents/photo_event/" & eCode & "/" & rsget("evtbbs_img2")%><% end if %></td>
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
