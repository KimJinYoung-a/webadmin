<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
	Dim idx, vQuery, i, itemid, imgsize, vAction
	idx  = requestCheckVar(request("idx"),10)
	itemid = requestCheckVar(request("itemid"),200)
	imgsize = requestCheckVar(request("imgsize"),3)
	vAction = requestCheckVar(request("action"),10)
	
	IF idx = "" THEN
		Response.Write "<script>alert('잘못된 경로입니다.\nNo. 번호가 있어야 합니다.');</script>"
		dbget.close()
		Response.End
	END IF	
	IF IsNumeric(idx) = False THEN
		Response.Write "<script>alert('잘못된 경로입니다.\nNo. 번호가 있어야 합니다.');</script>"
		dbget.close()
		Response.End
	END IF
	
	if imgsize = "" Then
		imgsize = "100"
	End If

	If vAction = "insert" Then
		For i = LBound(Split(itemid,",")) To UBound(Split(itemid,","))
			vQuery = vQuery & " IF NOT EXISTS(select itemidx from db_giftplus.dbo.tbl_stylelife_weekly_item where idx = '" & idx & "' and itemid = '" & Trim(Split(itemid,",")(i)) & "') " & vbCrLf
			vQuery = vQuery & " 	BEGIN " & vbCrLf
			vQuery = vQuery & " 		insert into db_giftplus.dbo.tbl_stylelife_weekly_item(idx, itemid, imgsize) values('" & idx & "', '" & Trim(Split(itemid,",")(i)) & "', '" & imgsize & "') " & vbCrLf
			vQuery = vQuery & " 	END " & vbCrLf
		Next
		dbget.execute vQuery
	ElseIf vAction = "delete" Then
		vQuery = "delete db_giftplus.dbo.tbl_stylelife_weekly_item where idx = '" & idx & "' and itemid IN(" & itemid & ")"
		dbget.execute vQuery
	ElseIf vAction = "update" Then
		vQuery = "update db_giftplus.dbo.tbl_stylelife_weekly_item set imgsize = '" & imgsize & "' where idx = '" & idx & "' and itemid IN(" & itemid & ")"
		dbget.execute vQuery
	End If
%>

<script language="javascript">
opener.document.location.reload();
document.location.href = "stylelife_weekly_item.asp?idx=<%=idx%>";
</script>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	