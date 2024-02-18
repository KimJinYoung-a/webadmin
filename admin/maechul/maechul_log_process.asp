<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr
dim mode
dim startdate, enddate, orderserial

mode = requestCheckVar(request("mode"), 32)
startdate = requestCheckVar(request("startdate"), 32)
enddate = requestCheckVar(request("enddate"), 32)
orderserial = requestCheckVar(request("orderserial"), 32)

if (mode = "reorgorder") then

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeOrgOrderLog_ON] '" + CStr(startdate) + "', '" + CStr(enddate) + "' "
	db3_rsget.Open sqlStr, db3_dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

elseif (mode = "recsorder") then

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeCSOrderLog_ON] '" + CStr(startdate) + "', '" + CStr(enddate) + "' "
	db3_rsget.Open sqlStr, db3_dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

elseif (mode = "reorgorderfingers") then

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeOrgOrderLog_ACA] '" + CStr(startdate) + "', '" + CStr(enddate) + "' "
	db3_rsget.Open sqlStr, db3_dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

elseif (mode = "recsorderfingers") then

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeCSOrderLog_ACA] '" + CStr(startdate) + "', '" + CStr(enddate) + "' "
	db3_rsget.Open sqlStr, db3_dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

elseif (mode = "reorgorderone") then

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeOrgOrderLog_ON] '" + CStr(startdate) + "', '" + CStr(enddate) + "', '" + CStr(orderserial) + "' "
	db3_rsget.Open sqlStr, db3_dbget, 1

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeOrgOrderLog_ACA] '" + CStr(startdate) + "', '" + CStr(enddate) + "', '" + CStr(orderserial) + "' "
	db3_rsget.Open sqlStr, db3_dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

elseif (mode = "recsorderone") then

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeCSOrderLog_ON] '" + CStr(startdate) + "', '" + CStr(enddate) + "', '" + CStr(orderserial) + "' "
	''response.write sqlStr
	db3_rsget.Open sqlStr, db3_dbget, 1

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeCSOrderLog_ACA] '" + CStr(startdate) + "', '" + CStr(enddate) + "', '" + CStr(orderserial) + "' "
	db3_rsget.Open sqlStr, db3_dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"
elseif (mode = "reOrgorderCSone") then

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeOrgOrderLog_ByOrderserial_ON] '" + CStr(orderserial) + "' "
	db3_rsget.Open sqlStr, db3_dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

elseif (mode = "reOrgorderCSoneOFF") then

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeOrgOrderLog_ByOrderserial_OFF] '" + CStr(orderserial) + "' "

	'response.write sqlStr & "<br>"
	'response.end
	db3_rsget.Open sqlStr, db3_dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

elseif (mode = "reOrgorderCSoneACA") then

	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeOrgOrderLog_ByOrderserial_ACA] '" + CStr(orderserial) + "' "
	db3_rsget.Open sqlStr, db3_dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('작성되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"
elseif (mode="relogorderserialwithque") then
	sqlStr = " exec [db_datamart].[dbo].[usp_Ten_MakeOrderLog_BYOrderserialWithQueFin] '" + CStr(orderserial) + "' "
	db3_rsget.Open sqlStr, db3_dbget, 1

	response.write	"OK"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
