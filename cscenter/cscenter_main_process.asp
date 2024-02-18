<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim SQL
dim menupos
dim mode
dim orderdetailidx
dim csTime

menupos = requestCheckVar(request("menupos"), 32)
mode = requestCheckVar(request("mode"), 32)
orderdetailidx = requestCheckVar(request("orderdetailidx"), 32)
csTime = requestCheckVar(request("csTime"), 32)

if (mode = "setstockoutchargeuser") then
	'==============================================================================
	SQL = " exec db_cs.[dbo].[sp_Ten_MichulgoStockout_SetChargeID] " + CStr(orderdetailidx) + " "
	rsget.Open SQL, dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('분배되었습니다.'); location.href = 'cscenter_main.asp?menupos=" & menupos & "'; " &_
					"</script>"

elseif (mode = "RefreshData") then

	application(csTime) = DateAdd("d", -1, now)

	response.write	"<script language='javascript'>" &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

elseif (mode = "updChulgoAbleOrdr") then

	SQL = " exec [db_cs].[dbo].[usp_Ten_MakeChulgoAbleOrderList] "
	rsget.Open SQL, dbget, 1

	response.write	"<script language='javascript'>" &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
