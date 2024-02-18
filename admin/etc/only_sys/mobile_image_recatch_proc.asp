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
	
	
	If vItemID = "" Then
		dbget.close()
		Response.Write "<script>alert('잘못된접근');location.href='/admin/etc/only_sys/mobile_image_recatch.asp';</script>"
		Response.End
	End If
	
	vQuery = vQuery & "insert into db_etcmall.dbo.tbl_outmall_API_Que" & vbCrLf
	vQuery = vQuery & "select 'appDTL','EDIT',itemid,1100,GETDATE(),NULL,NULL,NULL,NULL,'" & session("ssBctId") & "'" & vbCrLf
	vQuery = vQuery & "from db_item.dbo.tbl_item" & vbCrLf
	vQuery = vQuery & "where itemid in (" & vItemID & ")" & vbCrLf
	dbget.Execute vQuery
%>

<script language="javascript">
document.location.href = "/admin/etc/only_sys/mobile_image_recatch.asp?itemid=<%=vItemID%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->