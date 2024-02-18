<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
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
dim songjangdiv

mode = requestCheckVar(request("mode"), 32)
songjangdiv = requestCheckVar(request("songjangdiv"), 32)

if (mode = "retry") then
	sqlStr = " update "
	sqlStr = sqlStr + " [db_datamart].[dbo].[tbl_DeliveryTrackingList] "
	sqlStr = sqlStr + " set checkCnt = checkCnt - 1 "
	sqlStr = sqlStr + " where 1=1 "
	sqlStr = sqlStr + " and beasongdate >= convert(varchar(10), DateAdd(d, -14, getdate()), 121) "
	sqlStr = sqlStr + " and songjangdiv = " & songjangdiv
	sqlStr = sqlStr + " and realDeliveryDate is NULL "
	sqlStr = sqlStr + " and checkCnt >= 3 "
	db3_dbget.Execute sqlStr

	response.write	"<script language='javascript'>" &_
					"	alert('저장되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"

elseif (mode = "receivedata") then

	sqlStr = " exec [db_cs].[dbo].[usp_Ten_DeliveryTrackingList_Get] "
	db3_dbget.Execute sqlStr

	response.write	"<script language='javascript'>" &_
					"	alert('저장되었습니다.'); " &_
					"	location.replace('" + CStr(refer) + "'); " &_
					"</script>"
else

	response.write "잘못된 접근입니다."

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
