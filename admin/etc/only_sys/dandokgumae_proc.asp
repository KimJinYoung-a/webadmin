<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/only_sys_cls.asp"-->

<%
	Dim vQuery, vItemID, vIsDandok, vIsSunChak
	vItemID = requestCheckVar(Request("itemid"),300)
	vIsDandok = requestCheckVar(Request("dandok"),1)
	vIsSunChak = requestCheckVar(Request("sunchak"),1)
	
	
	If vItemID = "" Then
		dbget.close()
		Response.Write "<script>alert('잘못된접근');location.href='/admin/etc/only_sys/dandokgumae.asp';</script>"
		Response.End
	End If
	
	vQuery = vQuery & "update [db_item].[dbo].[tbl_Item] set" & vbCrLf
	If vIsDandok = "o" Then
		vQuery = vQuery & "reserveItemTp = '1'" & vbCrLf
	End If
	If vIsSunChak = "o" Then
		vQuery = vQuery & ",availPayType = '9'" & vbCrLf
	End If
	vQuery = vQuery & "where itemid in(" & vItemID & ")"
	dbget.Execute vQuery
%>

<script language="javascript">
document.location.href = "/admin/etc/only_sys/dandokgumae.asp?itemid=<%=vItemID%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->