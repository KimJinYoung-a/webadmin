<% Option Explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : �ǽ� �������� �Է�/����
' Hieditor : 2017.09.01 ������ ����
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	dim sqlstr, mode, loginuserid, occupation, nickname, adminid, idx, updchkadminid
	mode	=	requestcheckvar(request("frmmode"),5)
	idx = requestcheckvar(request("frmidx"), 20)
	occupation	=	requestcheckvar(request("frmoccupation"),200)
	nickname	=	requestcheckvar(unescape(request("frmnickname")),200)
	adminid	=	requestcheckvar(request("frmadminid"),200)
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

	'// ������Ʈ�� ��쿣 �� ���� ������ ���Ͽ� Ʋ���� ƨ���.
	If Trim(mode)="upd" Then
		sqlstr = " Select * From db_sitemaster.dbo.tbl_piece_nickname Where idx='"&idx&"' "
		rsget.Open SqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.bof Or rsget.eof) Then
			updchkadminid = rsget("adminid")
		Else
			Response.Write "ERR||�������� ��η� �������ּ���."
			Response.End
		End If

		If Trim(updchkadminid) <> loginuserid Then
			Response.Write "ERR||���������� ���θ� �����մϴ�."
			Response.End
		End If
		rsget.close
	End If

	if mode="ins" Then
		sqlstr = " insert into db_sitemaster.dbo.tbl_piece_nickname (adminid, occupation, nickname, lastupdate) "
		sqlstr = sqlstr & " values ('"&loginuserid&"', '"&occupation&"', '"&nickname&"', getdate()) "
		dbget.execute sqlstr
		Response.Write "OK||1"
		dbget.close() : Response.End
	ElseIf mode="upd" Then
		sqlstr = " update db_sitemaster.dbo.tbl_piece_nickname set occupation='"&occupation&"', nickname='"&nickname&"', lastupdate = getdate() Where idx = '"&idx&"' "
		dbget.execute sqlstr
		Response.Write "OK||2"
		dbget.close() : Response.End
	else
		Response.Write "ERR||�������� ��η� �������ּ���."
		dbget.close() : Response.End
	end If
	

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
