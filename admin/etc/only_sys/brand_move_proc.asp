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
	Dim cBrand, vQuery, vMakerID, vItemID, vNewMakerID, vBrandName
	vMakerID = requestCheckVar(Request("makerid"),100)
	vItemID = Request("itemid")
	vNewMakerID = requestCheckVar(Request("newmakerid"),100)
	vBrandName = requestCheckVar(Request("brandname"),100)
	
	If vNewMakerID = "" Then
		dbget.close()
		Response.Write "<script>alert('肋给等立辟');location.href='/admin/etc/only_sys/brand_move.asp';</script>"
		Response.End
	End If
	If vMakerID = "" AND vItemID = "" Then
		dbget.close()
		Response.Write "<script>alert('肋给等立辟');location.href='/admin/etc/only_sys/brand_move.asp';</script>"
		Response.End
	End If
	If vMakerID <> "" AND vItemID <> "" Then
		dbget.close()
		Response.Write "<script>alert('肋给等立辟');location.href='/admin/etc/only_sys/brand_move.asp';</script>"
		Response.End
	End If
	
	
	vQuery = ""
	vQuery = vQuery & "update db_item.dbo.tbl_item" & vbCrLf
	vQuery = vQuery & "set" & vbCrLf
	vQuery = vQuery & "makerid = '" & vNewMakerID & "', brandname = '" & vBrandName & "', lastupdate = getdate()" & vbCrLf
	vQuery = vQuery & "where itemid in" & vbCrLf
	vQuery = vQuery & "(" & vbCrLf
	If vMakerID <> "" Then
		vQuery = vQuery & "	select itemid from db_item.dbo.tbl_item" & vbCrLf
		vQuery = vQuery & "	where" & vbCrLf
		vQuery = vQuery & "	makerid = '" & vMakerID & "'" & vbCrLf
	End IF
	If vItemID <> "" Then
		vQuery = vQuery & vItemID & vbCrLf
	End IF
	vQuery = vQuery & ")" & vbCrLf
	dbget.Execute vQuery
	
%>

<script language="javascript">
document.location.href = "/admin/etc/only_sys/brand_move.asp?change=o&makerid=<%=vMakerID%>&itemid=<%=vItemID%>&newmakerid=<%=vNewMakerID%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->