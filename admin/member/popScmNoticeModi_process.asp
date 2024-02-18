<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �系��������
' Hieditor : �̻� ����
'			 2022.07.12 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/admin_keyclass.asp" -->
<%

dim mode
dim idx, scheduleDate, title, contents, reguserid, modiuserid, isusing, dispno, regdate, lastupdate

mode = requestCheckVar(request("mode"), 32)

idx = requestCheckVar(request("idx"), 32)
scheduleDate = requestCheckVar(request("scheduleDate"), 32)
title = requestCheckVar(request("title"), 32)
contents = requestCheckVar(request("contents"), 320)
reguserid = session("ssBctId")
modiuserid = session("ssBctId")
isusing = requestCheckVar(request("isusing"), 1)
dispno = requestCheckVar(request("dispno"), 2)

dim strSql

IF Not(C_OP Or C_PSMngPart Or C_SYSTEM_Part or C_ADMIN_AUTH) Then
	response.write "<script type='text/javascript'>alert('�系�������� ���/������ �λ��ѹ����� �������� �����մϴ�.'); history.back();</script>"
	dbget.close : response.end
End If

if scheduleDate <> "" and not(isnull(scheduleDate)) then
	scheduleDate = ReplaceBracket(scheduleDate)
end If
if title <> "" and not(isnull(title)) then
	title = ReplaceBracket(title)
end If
if contents <> "" and not(isnull(contents)) then
	contents = ReplaceBracket(contents)
end If

if checkNotValidHTML(scheduleDate) then
	response.write "<script type='text/javascript'>alert('������ ��ũ��Ʈ�� �Է��� �� �����ϴ�.'); history.back();</script>"
	dbget.close : response.end
end if

if checkNotValidHTML(title) then
	response.write "<script type='text/javascript'>alert('���� ��ũ��Ʈ�� �Է��� �� �����ϴ�.'); history.back();</script>"
	dbget.close : response.end
end if

if checkNotValidHTML(contents) then
	response.write "<script type='text/javascript'>alert('���뿡 ��ũ��Ʈ�� �Է��� �� �����ϴ�.'); history.back();</script>"
	dbget.close : response.end
end if

'// �űԵ��
if mode = "add" then
	strSql = "insert into [db_board].[dbo].[tbl_scm_notice](scheduleDate, title, contents, reguserid, modiuserid, isusing, dispno, regdate, lastupdate) "
	strSql = strSql + " values('" & html2db(scheduleDate) & "', '" & html2db(title) & "', '" & html2db(contents) & "', '" & reguserid & "', '" & modiuserid & "', 'Y', '" & html2db(dispno) & "', getdate(), getdate()) "
	'response.write strSql
	dbget.execute strSql

elseif mode = "modi" then
	strSql = "update [db_board].[dbo].[tbl_scm_notice] "
	strSql = strSql + " set modiuserid = '" & modiuserid & "', lastupdate = getdate() "
	strSql = strSql + " , scheduleDate = '" & html2db(scheduleDate) & "' "
	strSql = strSql + " , title = '" & html2db(title) & "' "
	strSql = strSql + " , contents = '" & html2db(contents) & "' "
	strSql = strSql + " , dispno = '" & html2db(dispno) & "' "
	strSql = strSql + " where idx = " & idx
	'response.write strSql
	dbget.execute strSql

elseif mode = "del" then
	strSql = ""
	strSql = strSql + " update [db_board].[dbo].[tbl_scm_notice] "
	strSql = strSql + " set modiuserid = '" & modiuserid & "', lastupdate = getdate() "
	strSql = strSql + " , isusing = 'N' "
	strSql = strSql + " where idx = " & idx
	'response.write strSql
	dbget.execute strSql
end if

%>
<script type='text/javascript'>
location.href="popScmNoticeModi.asp";
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
