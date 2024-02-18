<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/sitemaster/shoppingtalk/classes/shoppingtalkCls.asp" -->

<%
	Dim vQuery, i, vGubun, vAction, vTalkIdx, vIdx, vUseYN
	vGubun = Request("gubun")
	vAction = Request("action")
	vTalkIdx = Request("talkidx")
	vUseYN = Request("useyn")
	vIdx = Request("idx")
	
	If vAction = "update" Then
		If vGubun = "talk" Then
			vQuery = "UPDATE [db_board].[dbo].[tbl_shopping_talk] SET useyn = '" & vUseYN & "' WHERE talk_idx = '" & vTalkIdx & "'"
			dbget.execute vQuery
		ElseIf vGubun = "comment" Then
			vQuery = "UPDATE [db_board].[dbo].[tbl_shopping_talk_comment] SET useyn = '" & vUseYN & "' WHERE talk_idx = '" & vTalkIdx & "' AND idx = '" & vIdx & "'"
			dbget.execute vQuery
		End If
	End If
%>

<script>
parent.location.reload();
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->