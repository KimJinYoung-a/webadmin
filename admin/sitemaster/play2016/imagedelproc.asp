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
	Dim i, vQuery, vIdx
	vIdx = requestCheckVar(Request("idx"),10)

	vQuery = "DELETE [db_giftplus].[dbo].[tbl_play_image] WHERE idx = '" & vIdx & "' "
	dbget.Execute vQuery

	Response.Write "<script>alert('삭제되었습니다.');opener.location.reload();window.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->