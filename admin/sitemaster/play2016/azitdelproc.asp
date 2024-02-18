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
	Dim i, vQuery, vDidx, vGroupNum
	vGroupNum 	= requestCheckVar(Request("groupnum"),10)
	vDidx 		= requestCheckVar(Request("didx"),10)

	vQuery = "DELETE [db_giftplus].[dbo].[tbl_play_azit] WHERE didx = '" & vDidx & "' and groupnum = '" & vGroupNum & "' "
	vQuery = vQuery & "UPDATE [db_giftplus].[dbo].[tbl_play_image] SET imgurl = '' WHERE didx = '" & vDidx & "' and groupnum = '" & vGroupNum & "' and gubun = '7'"
	dbget.Execute vQuery

	Response.Write "<script>alert('처리되었습니다.');opener.location.reload();window.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->