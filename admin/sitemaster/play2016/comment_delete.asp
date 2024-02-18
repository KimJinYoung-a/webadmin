<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/play2016Cls.asp" -->
<%
	'### 기본정보 ###
	Dim i, vQuery, vCate, vDidx, vIdx, vReferer
	vCate		= requestCheckVar(Request("cate"),10)
	vIdx 		= requestCheckVar(Request("idx"),10)
	vDidx 		= requestCheckVar(Request("didx"),10)
	
	vReferer = Request.ServerVariables("HTTP_REFERER")
	
	If vCate = "42" Then
		vQuery = "Delete [db_giftplus].[dbo].[tbl_play_thingthing_entry] WHERE didx = '" & vDidx & "' and idx = '" & vIdx & "' "
		dbget.Execute vQuery
	ElseIf vCate = "1" Then
		vQuery = "Delete [db_giftplus].[dbo].[tbl_play_playlist_comment] WHERE didx = '" & vDidx & "' and idx = '" & vIdx & "' "
		dbget.Execute vQuery
	End If

	Response.Write "<script>alert('처리되었습니다.');location.href='" & vReferer & "';</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->