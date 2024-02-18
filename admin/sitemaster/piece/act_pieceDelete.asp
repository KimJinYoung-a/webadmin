<% Option Explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : 피스 삭제
' Hieditor : 2017.09.01 원승현 생성
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	dim sqlstr, mode, loginuserid, occupation, nickname, adminid, idx, updchkadminid
	idx = requestcheckvar(request("frmDelidx"), 20)
	adminid	=	requestcheckvar(request("frmDeladminid"),200)
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

	'// idx값이 없으면 튕긴다.
	if Trim(idx)="" then
		Response.Write "ERR||정상적인 경로로 접근해주세요."
		dbget.close() : Response.End
	End If

	sqlstr = " update db_sitemaster.dbo.tbl_piece set deleteyn='Y', lastupdate = getdate(), deladminid='"&adminid&"' Where idx = '"&idx&"' "
	dbget.execute sqlstr
	Response.Write "OK||1"
	dbget.close() : Response.End

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
