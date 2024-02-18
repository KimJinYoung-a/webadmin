<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim vECode, vQuery, vFB_appid, vFB_content
	vECode = Request("ecode")
	If vECode = "" Then
		Response.End
	End If
	
	vFB_appid	= Request("fb_appid")
	vFB_content	= html2db(Request("fb_content"))
	
	vQuery = "UPDATE [db_event].[dbo].[tbl_event_display] SET fb_appid = '" & vFB_appid & "', fb_content = '" & vFB_content & "' WHERE evt_code = '" & vECode & "'"
	dbget.execute vQuery
	
	Response.Write "<Script>alert('저장되었습니다.');location.href='pop_event_facebookapp.asp?ecode="&vECode&"';</script>"
	Response.End
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->