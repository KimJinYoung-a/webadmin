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
<!-- #include virtual="/lib/classes/play/play_moCls.asp" -->
<%
	Dim vQuery, vIdx, vPlayType, vPlayTypeName, vIsUsing
	vPlayType = requestCheckVar(Request("playtype"),3)
	vPlayTypeName = requestCheckVar(Request("playtypename"),100)
	vIsUsing = requestCheckVar(Request("isusing"),10)
	
	If vPlayType = "" Then
		vQuery = "INSERT INTO [db_sitemaster].[dbo].[tbl_play_mo_stylecode](typename, isusing) VALUES('" & vPlayTypeName & "', '" & vIsUsing & "')"
		dbget.Execute vQuery
	ElseIf vPlayType <> "" Then
		vQuery = "UPDATE [db_sitemaster].[dbo].[tbl_play_mo_stylecode] SET typename = '" & vPlayTypeName & "', isusing = '" & vIsUsing & "' WHERE type = '" & vPlayType & "'"
		dbget.Execute vQuery
	End IF
	Response.Write "<script>alert('처리되었습니다.');location.href='pop_style.asp';</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->