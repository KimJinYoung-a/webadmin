<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/offshop/staff/offshop_employee_managementCls.asp"-->

<%
	Dim i, vQuery, vAction, vEmpNO, vWorkDate, vWorkCode
	vAction = Request("action")
	vEmpNO = Request("empno")
	vWorkDate = Request("workdate")
	vWorkCode = Request("workcode")
	
	If vAction = "oneupdate" Then
		vQuery = "UPDATE [db_partner].[dbo].[tbl_offshop_employee_workschedule] SET workcode = '" & vWorkCode & "', reguserid = '" & session("ssBctId") & "' "
		vQuery = vQuery & "WHERE empno = '" & vEmpNO & "' AND convert(varchar(10),workdate,120) = '" & vWorkDate & "'"
		dbget.Execute vQuery
	End IF
%>

<script type="text/javascript">
<!--
document.domain = "10x10.co.kr";

<% If vAction = "oneupdate" Then %>
	parent.jsReload();
<% End IF %>
//-->
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->