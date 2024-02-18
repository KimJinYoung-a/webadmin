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
	Dim vQuery, cGoodUsing, vUserID, vItemID, vIsUsing
	vUserID = requestCheckVar(Request("userid"),100)
	vItemID = requestCheckVar(Request("itemid"),10)
	vIsUsing = requestCheckVar(Request("isusing"),1)
	
	If vUserID = "" AND vItemID = "" Then
		dbget.close()
		Response.Write "<script>alert('잘못된접근');location.href='/admin/etc/only_sys/brand_ordercomment.asp';</script>"
		Response.End
	End If
	
	vQuery = vQuery & "update db_board.dbo.tbl_Item_Evaluate" & vbCrLf
	vQuery = vQuery & "set IsUsing = '" & vIsUsing & "'" & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "' and itemid = '" & vItemID & "'"
	dbget.Execute vQuery
%>

<script language="javascript">
document.location.href = "/admin/etc/only_sys/goodusing.asp?userid=<%=vUserID%>&itemid=<%=vItemID%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->