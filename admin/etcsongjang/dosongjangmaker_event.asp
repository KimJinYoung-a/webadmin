<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim txsongjang, iid, gubuncd, mode
txsongjang = request("txsongjang")
iid = request("id")
mode = request("mode")

dim sqlStr

if (mode="michulgo") then
	sqlStr = "update [db_sitemaster].[dbo].tbl_etc_songjang"
	sqlStr = sqlStr + " set issended='N'"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	
	dbget.Execute sqlStr
else
    sqlStr = "update [db_sitemaster].[dbo].tbl_etc_songjang"
	sqlStr = sqlStr + " set songjangno = '" + txsongjang + "',"
	sqlStr = sqlStr + " senddate=getdate(),"
	sqlStr = sqlStr + " issended='Y'"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	
	dbget.Execute sqlStr
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
alert('저장되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->