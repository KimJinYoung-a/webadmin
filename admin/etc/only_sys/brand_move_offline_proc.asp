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
		Response.Write "<script>alert('肋给等立辟');location.href='/admin/etc/only_sys/brand_move_offline.asp';</script>"
		Response.End
	End If
	If vMakerID = "" AND vItemID = "" Then
		dbget.close()
		Response.Write "<script>alert('肋给等立辟');location.href='/admin/etc/only_sys/brand_move_offline.asp';</script>"
		Response.End
	End If
	
	vMakerID = Replace(vMakerID, " ", "")
	vMakerID = Replace(vMakerID, ",", "','")
	vMakerID = "'" & vMakerID & "'"
	
	vQuery = ""
	vQuery = vQuery & "update db_shop.dbo.tbl_shop_item " & vbCrLf
	vQuery = vQuery & "set " & vbCrLf
	vQuery = vQuery & "makerid = '" & vNewMakerID & "', updt = getdate() " & vbCrLf
	vQuery = vQuery & "where 1=1 " &vbCrLf
	If vMakerID <> "" Then
		vQuery = vQuery & "and makerid in (" & vMakerID & ") " &vbCrLf
	End If
	If vItemID <> "" Then
		vQuery = vQuery & "and shopitemid in(" & vItemID & ") " & vbCrLf
	End IF
	vQuery = vQuery & "and itemgubun = '90' and itemoption = '0000'"
	dbget.Execute vQuery
	
%>

<script language="javascript">
document.location.href = "/admin/etc/only_sys/brand_move_offline.asp?change=o&makerid=<%=vMakerID%>&itemid=<%=vItemID%>&newmakerid=<%=vNewMakerID%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->