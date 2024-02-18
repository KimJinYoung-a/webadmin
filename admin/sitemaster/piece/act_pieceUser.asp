<% Option Explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : 피스 유저정보 입력/수정
' Hieditor : 2017.09.01 원승현 생성
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
	loginuserid		=	session("ssBctId")	'어드민등록자id


	if loginuserid="" or isNull(loginuserid) then
		Response.Write "ERR||로그인을 해주세요."
		dbget.close() : Response.End
	End If

	'// 넘어온 adminid값과 현재 세션에 있는 id값을 비교한다.
	if Trim(loginuserid)<>trim(adminid) then
		Response.Write "ERR||정상적인 경로로 접근해주세요."
		dbget.close() : Response.End
	End If

	'// 업데이트일 경우엔 원 유저 정보와 비교하여 틀리면 튕긴다.
	If Trim(mode)="upd" Then
		sqlstr = " Select * From db_sitemaster.dbo.tbl_piece_nickname Where idx='"&idx&"' "
		rsget.Open SqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.bof Or rsget.eof) Then
			updchkadminid = rsget("adminid")
		Else
			Response.Write "ERR||정상적인 경로로 접근해주세요."
			Response.End
		End If

		If Trim(updchkadminid) <> loginuserid Then
			Response.Write "ERR||정보수정은 본인만 가능합니다."
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
		Response.Write "ERR||정상적인 경로로 접속해주세요."
		dbget.close() : Response.End
	end If
	

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
