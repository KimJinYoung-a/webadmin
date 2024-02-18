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
	Dim vQuery, cOrderList, vUserID, vOrderSerial, vDB, vDBt, vDisplayYN, arrList, intLoop, vCount
	vUserID = requestCheckVar(Request("userid"),100)
	vOrderSerial = requestCheckVar(Request("orderserial"),11)
	vDB = NullFillWith(requestCheckVar(Request("db"),50),"1")
	vDBt = vDB
	If vDB = "1" Then
		vDB = "[db_order].[dbo].[tbl_order_master]"
	ElseIf vDB = "2" Then
		vDB = "[db_log].[dbo].[tbl_old_order_master_2003]"
	End If
	vDisplayYN = requestCheckVar(Request("displayyn"),1)
	
	vQuery = vQuery & "update " & vDB & vbCrLf
	vQuery = vQuery & "set userDisplayYn = '" & vDisplayYN & "'" & vbCrLf
	vQuery = vQuery & "where userid = '" & vUserID & "'" & vbCrLf
	If vOrderSerial <> "" Then
		vQuery = vQuery & "and orderserial = '" & vOrderSerial & "'" & vbCrLf
	End IF
	dbget.Execute vQuery
%>

<script language="javascript">
document.location.href = "/admin/etc/only_sys/orderlist.asp?userid=<%=vUserID%>&orderserial=<%=vOrderSerial%>&db=<%=vDBt%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->