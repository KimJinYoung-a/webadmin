<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/onelineCls.asp"-->

<%
	Dim vEvtCode, vGubun, vQuery, vIdx
	vEvtCode 	= requestCheckVar(Request("eC"),10)
	vIdx		= requestCheckVar(Request("idx"),10)
	vGubun		= Request("gubun")
	
	If vGubun = "0" then
		dbget.execute "UPDATE [db_contents].[dbo].[tbl_one_comment] SET isusing = 'N' WHERE evt_code = '" & vEvtCode & "' AND idx = '" & vIdx & "'"
	ElseIf vGubun = "1" then
		dbget.execute "UPDATE [db_contents].[dbo].[tbl_one_comment] SET isusing = 'Y' WHERE evt_code = '" & vEvtCode & "' AND idx = '" & vIdx & "'"
	End IF
%>

	<script language="javascript">top.location.reload();</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->