<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->

<%
	'######################################################################################
	'	반드시 [db_board].[dbo].[tbl_scm_commonBoard_list] 에 boardgubun 을 추가해야함.
	'######################################################################################
	''
	Dim vParam, parentidx, registuserid
	parentidx = "1"
	registuserid = ""
	vParam = "pidx="&parentidx&"&registid="&registuserid&"&boardtype=c&boardgubun=testcomment&cols=95&rows=3&btnwidth=80&btnheight=50"
%>
<iframe src="comment.asp?<%=vParam%>" name="iframeComment" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)" style="width:800px;"></iframe>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->