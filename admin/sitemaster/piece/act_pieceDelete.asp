<% Option Explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : �ǽ� ����
' Hieditor : 2017.09.01 ������ ����
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	dim sqlstr, mode, loginuserid, occupation, nickname, adminid, idx, updchkadminid
	idx = requestcheckvar(request("frmDelidx"), 20)
	adminid	=	requestcheckvar(request("frmDeladminid"),200)
	loginuserid		=	session("ssBctId")	'���ε����id


	if loginuserid="" or isNull(loginuserid) then
		Response.Write "ERR||�α����� ���ּ���."
		dbget.close() : Response.End
	End If

	'// �Ѿ�� adminid���� ���� ���ǿ� �ִ� id���� ���Ѵ�.
	if Trim(loginuserid)<>trim(adminid) then
		Response.Write "ERR||�������� ��η� �������ּ���."
		dbget.close() : Response.End
	End If

	'// idx���� ������ ƨ���.
	if Trim(idx)="" then
		Response.Write "ERR||�������� ��η� �������ּ���."
		dbget.close() : Response.End
	End If

	sqlstr = " update db_sitemaster.dbo.tbl_piece set deleteyn='Y', lastupdate = getdate(), deladminid='"&adminid&"' Where idx = '"&idx&"' "
	dbget.execute sqlstr
	Response.Write "OK||1"
	dbget.close() : Response.End

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
