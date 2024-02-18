<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/staff/offshop_employee_managementCls.asp"-->

<%
	Dim vQuery, vEmpNo, vYYYYMMDD, vWorkDate, vWorkType, vPlaceID
	vEmpNo = Request("empno")
	vYYYYMMDD = Request("yyyymmdd")
	vWorkDate = Request("inoutdate") & " " & Request("inouttime")
	vWorkType = Request("inouttype")
	vPlaceID = Request("placeid")
	
	vQuery = "INSERT INTO [db_partner].[dbo].[tbl_user_inouttime_log](placeid, empno, YYYYMMDD, inoutType, inoutTime, posIdx, posDate) "
	vQuery = vQuery & "VALUES('" & vPlaceID & "', '" & vEmpNo & "', '" & vYYYYMMDD & "', '" & vWorkType & "', '" & vWorkDate & "', 0, getdate())"
	dbget.Execute vQuery
%>

<script type="text/javascript">
document.domain = "10x10.co.kr";
opener.location.reload();
window.close();
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->