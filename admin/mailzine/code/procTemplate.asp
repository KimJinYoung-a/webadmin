<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ������
' History : 2018.04.27 �̻� ����(���Ϸ� ���� ���� ���Ϸ��� �߼� ���� ����. ���� �������� ����.)
'			2019.06.24 ������ ����(���ø� ��� �ű� �߰�)
'			2020.05.28 �ѿ�� ����(TMS ���Ϸ� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
 dim iResult, sMode
 Dim mailzineKind, contentsKind, contentsEA, idx
 Dim strSql, codevalue, sSortNo, i
 
mailzineKind = requestCheckVar(Request("mailzineKind"),10)
contentsKind = requestCheckVar(Request("contentsKind"),10)
contentsEA = requestCheckVar(Request("contentsEA"),2)
idx = requestCheckVar(Request("idx"),10)
sMode = requestCheckVar(Request("mode"),10)

IF sMode = "I" THEN
	strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_mailzine_contents_manage] (kindCode, contentsCode, contentsEa, sortidx)"&_
			" Values('"&mailzineKind&"',"&contentsKind&",'"&contentsEA&"',0)"
	dbget.execute strSql
ELSEIF sMode="U" THEN
	strSql =" UPDATE [db_sitemaster].[dbo].[tbl_mailzine_contents_manage] Set kindCode = '"&mailzineKind&"', contentsCode = "&contentsKind&" , contentsEa ='"&contentsEA&"'"&_
			" WHERE idx =" & idx
	dbget.execute strSql
ELSEIF sMode="D" THEN
	strSql =" DELETE FROM [db_sitemaster].[dbo].[tbl_mailzine_contents_manage] WHERE idx =" & idx
	dbget.execute strSql
elseif sMode = "S" THEN
	'//����Ʈ��������
	for i=1 to request.form("idx").count
		codevalue = request.form("idx")(i)
		sSortNo = request.form("viewidx")(i)
		if sSortNo="" then sSortNo="0"

		strSql = strSql & "Update [db_sitemaster].[dbo].[tbl_mailzine_contents_manage] Set " & vbCrLf
		strSql = strSql & " sortidx=" & sSortNo & "" & vbCrLf
		strSql = strSql & " Where idx='" & codevalue & "';"
	Next
	If strSql <> "" then
		dbget.Execute strSql
	end if
END IF
	Call sbAlertMsg ("ó���Ǿ����ϴ�.", "/admin/mailzine/code/popManageTemplate.asp?mailzineKind="&mailzineKind, "self") 
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->