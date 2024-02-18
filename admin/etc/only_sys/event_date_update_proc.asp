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
	Dim vQuery, vEvtCode, vEvtStatus, vEvtEDate
	vEvtCode = requestCheckVar(Request("evt_code"),6)
	vEvtStatus = requestCheckVar(Request("evt_status"),2)
	vEvtEDate = requestCheckVar(Request("evt_edate"),10)
	
	If vEvtCode = "" OR vEvtEDate = "" Then
		dbget.close()
		Response.Write "<script>alert('잘못된접근');location.href='/admin/etc/only_sys/event_date_update.asp';</script>"
		Response.End
	End If
	
	vQuery = "UPDATE [db_event].[dbo].[tbl_event] SET evt_enddate = '" & vEvtEDate & "'"
	If vEvtStatus <> "" Then
		vQuery = vQuery & ", evt_state = '" & vEvtStatus & "'"
	End IF
	vQuery = vQuery & "WHERE evt_code = '" & vEvtCode & "'"
	dbget.Execute vQuery
	
%>

<script language="javascript">
document.location.href = "/admin/etc/only_sys/event_date_update.asp?evt_code=<%=vEvtCode%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->