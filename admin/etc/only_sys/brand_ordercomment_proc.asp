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
	Dim cBrandOrder, vQuery, vMakerID, vItemID, vMoveItemCnt, vChange, vComment
	vMakerID = requestCheckVar(Request("makerid"),100)
	vItemID = Request("itemid")
	vComment = html2db(Request("comment"))
	
	If vMakerID = "" AND vItemID = "" Then
		dbget.close()
		Response.Write "<script>alert('잘못된접근');location.href='/admin/etc/only_sys/brand_ordercomment.asp';</script>"
		Response.End
	End If
	
	vQuery = ""
	vQuery = vQuery & "update db_item.dbo.tbl_item_Contents set" & vbCrLf
	vQuery = vQuery & "ordercomment = '" & vComment & "'" & vbCrLf
	vQuery = vQuery & "where itemid in (" & vbCrLf
	If vMakerID <> "" Then
		vQuery = vQuery & "	select itemid" & vbCrLf
		vQuery = vQuery & "	from db_item.dbo.tbl_item" & vbCrLf
		vQuery = vQuery & "	where makerid = '" & vMakerID & "'" & vbCrLf
	End If
	If vItemID <> "" Then
		vQuery = vQuery & "" & vItemID & "" & vbCrLf
	End If
	vQuery = vQuery & ")" & vbCrLf
	dbget.Execute vQuery
%>

<script language="javascript">
document.location.href = "/admin/etc/only_sys/brand_ordercomment.asp?makerid=<%=vMakerID%>&itemid=<%=vItemID%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->