<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/eventmanage/event_process.asp
' Description :  이벤트 개요 데이터처리 - 등록, 수정, 삭제
' History : 2007.02.12 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim sMode, sfimg, strSql , menupos,idx
sMode = requestCheckVar(Request.Form("hidM"),1) 
menupos = requestCheckVar(Request.Form("menupos"),10) 
sfimg= requestCheckVar(Request.Form("sfimg"),256) 
 idx = requestCheckVar(Request("idx"),10) 
SELECT Case sMode
Case "I"

	strSql = "INSERT INTO [db_sitemaster].[dbo].[tbl_scm_loginBackImg] (imgUrl, userid,lastupdate) "&vbCrlf&_
			"		VALUES ('"&sfimg&"','"&session("ssBctId")&"',getdate())"
	dbget.execute strSql
	Application("scmBGdiv") ="0"
	Call Alert_move ("배경화면 설정이 저장되었습니다.","loginImgMng.asp?menupos="&menupos)
Case "U"

	strSql = "UPDATE [db_sitemaster].[dbo].[tbl_scm_loginBackImg] SET imgUrl = '"&sfimg&"', userid='"&session("ssBctId")&"',lastupdate=getdate() "&vbCrlf&_
			" WHERE idx = "& idx
	dbget.execute strSql
	
	Application("scmBGdiv") ="0"
	Call Alert_move ("배경화면 설정이 저장되었습니다.","loginImgMng.asp?menupos="&menupos)
Case "D"

	strSql =  "UPDATE [db_sitemaster].[dbo].[tbl_scm_loginBackImg] SET isUsing = 0, userid='"&session("ssBctId")&"',lastupdate=getdate() "&vbCrlf&_
				" WHERE idx = "& idx
	dbget.execute strSql
	
	Application("scmBGdiv") ="0"
	Call Alert_move ("배경화면 설정이 삭제되었습니다.","loginImgMng.asp?menupos="&menupos)		
CASE Else
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.")
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->