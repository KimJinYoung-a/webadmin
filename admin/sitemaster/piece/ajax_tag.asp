<%@ language=vbscript %>
<% option explicit %>
<%
' Session.CodePage  = 65001
' Response.CharSet  = "UTF-8"
' Response.AddHeader "Pragma","no-cache"
' Response.AddHeader "cache-control", "no-staff"
' Response.Expires  = -1
%>
<%
'###########################################################
' Description : 피스 태그 자동완성 데이터 가져오기
' Hieditor : 2017.08.14 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	dim sqlstr, ajaxtagtext, mode, loginuserid, idx
	mode	=	requestcheckvar(request("mode"),5)
	loginuserid		=	session("ssBctId")	'어드민등록자id

	if loginuserid="" or isNull(loginuserid) then
		Response.Write "ERR||로그인을 해주세요"
		dbget.close() : Response.End
	End If

	if mode="admin" then
		sqlstr = " SELECT STUFF(( " & vbCrlf
        sqlstr = sqlstr & "     SELECT ',' + tagtext " & vbCrlf
        sqlstr = sqlstr & "       FROM  [db_sitemaster].[dbo].[tbl_piece_tag] " & vbCrlf
        sqlstr = sqlstr & "        FOR XML PATH('') " & vbCrlf
		sqlstr = sqlstr & " ),1,1,'') AS tagtext "
	
		rsget.Open sqlStr,dbget,1
		IF Not rsget.Eof Then
			ajaxtagtext = rsget(0)
		End IF
		rsget.close

		Response.Write "OK||"&escape(ajaxtagtext)
		
		dbget.close() : Response.End
	else
		Response.Write "ERR||정상적인 경로로 접속해주세요."
		dbget.close() : Response.End
	end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
