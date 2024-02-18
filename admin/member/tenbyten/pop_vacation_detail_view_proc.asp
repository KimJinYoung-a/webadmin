<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim vQuery, vDetailIdx, vHalfGubun
	vDetailIdx	= requestCheckVar(Request("detailidx"),12)
	vHalfGubun	= requestCheckVar(Request("halfgubun"),2)
	
	IF vDetailIdx = "" Then
		Response.Write "<script>alert('잘못된 경로입니다.1');</script>"
		dbget.close()
		Response.End
	End IF
	
	If vHalfGubun <> "am" AND vHalfGubun <> "pm" Then
		Response.Write "<script>alert('잘못된 경로입니다.2');</script>"
		dbget.close()
		Response.End
	End IF
	
	vQuery = "UPDATE [db_partner].[dbo].[tbl_vacation_detail] SET halfgubun = '" & vHalfGubun & "' WHERE idx = '" & vDetailIdx & "' "
	dbget.Execute vQuery
	
	Response.Write "<script>top.opener.document.location.reload();top.document.location.reload();</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->