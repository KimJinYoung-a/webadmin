<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/eventmanage/event_process.asp
' Description :  �̺�Ʈ ���� ������ó�� - ���, ����, ����
' History : 2007.02.12 ������ ����
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
	Call Alert_move ("���ȭ�� ������ ����Ǿ����ϴ�.","loginImgMng.asp?menupos="&menupos)
Case "U"

	strSql = "UPDATE [db_sitemaster].[dbo].[tbl_scm_loginBackImg] SET imgUrl = '"&sfimg&"', userid='"&session("ssBctId")&"',lastupdate=getdate() "&vbCrlf&_
			" WHERE idx = "& idx
	dbget.execute strSql
	
	Application("scmBGdiv") ="0"
	Call Alert_move ("���ȭ�� ������ ����Ǿ����ϴ�.","loginImgMng.asp?menupos="&menupos)
Case "D"

	strSql =  "UPDATE [db_sitemaster].[dbo].[tbl_scm_loginBackImg] SET isUsing = 0, userid='"&session("ssBctId")&"',lastupdate=getdate() "&vbCrlf&_
				" WHERE idx = "& idx
	dbget.execute strSql
	
	Application("scmBGdiv") ="0"
	Call Alert_move ("���ȭ�� ������ �����Ǿ����ϴ�.","loginImgMng.asp?menupos="&menupos)		
CASE Else
	Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.")
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->