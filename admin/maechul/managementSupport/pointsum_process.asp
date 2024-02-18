<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode
dim costpricepercent, yyyymm
dim idx, yyyymmdd

mode = requestCheckvar(request("mode"),32)
costpricepercent = requestCheckvar(request("costpricepercent"),32)
yyyymm = requestCheckvar(request("yyyymm"),32)
idx = requestCheckvar(request("idx"),32)
yyyymmdd = requestCheckvar(request("yyyymmdd"),32)

dim sqlStr
dim i, j, k

if (mode="modOnCostpricepercent") then

	sqlStr = " update db_summary.dbo.tbl_on_off_point_monthly_summary "
	sqlStr = sqlStr + " set costpricepercent = " + CStr(costpricepercent) + " "
	sqlStr = sqlStr + " where yyyymm = '" + CStr(yyyymm) + "' and onoffgubun = 'ON' "
	dbget.Execute sqlStr

elseif (mode="modOffCostpricepercent") then

	sqlStr = " update db_summary.dbo.tbl_on_off_point_monthly_summary "
	sqlStr = sqlStr + " set costpricepercent = " + CStr(costpricepercent) + " "
	sqlStr = sqlStr + " where yyyymm = '" + CStr(yyyymm) + "' and onoffgubun = 'OFF' "
	dbget.Execute sqlStr
elseif (mode="refreshpointDepositSummary") then
	if (yyyymm = "AUTO") then
		if (Day(Now()) = 1) then
			'// 전달
			yyyymm = Left(DateAdd("m", -1, Now()), 7)
		else
			'// 전일까지
			yyyymm = Left(DateAdd("m", 0, Now()), 7)
		end if
	end if

    sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_PointDeposit_summary_Make] '"&yyyymm&"'"
    dbget.Execute sqlStr

elseif (mode="modiDepositDate") then
    sqlStr = " update db_user.dbo.tbl_depositlog "
	sqlStr = sqlStr + " set fixyyyymmdd = '" & yyyymmdd & "' "
	sqlStr = sqlStr + " where idx = " & idx
    ''response.write sqlStr
    dbget.Execute sqlStr

    response.write "<script>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
    dbget.close() : response.end
elseif (mode="modiGiftDate") then
    sqlStr = " update [db_user].[dbo].[tbl_giftcard_log] "
	sqlStr = sqlStr + " set fixyyyymmdd = '" & yyyymmdd & "' "
	sqlStr = sqlStr + " where idx = " & idx
    ''response.write sqlStr
    dbget.Execute sqlStr

    response.write "<script>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
    dbget.close() : response.end
else

end if

%>
<% if (IsAutoScript) then  %>
OK
<% else %>
<script language='javascript'>
alert('저장되었습니다.');
location.replace('<%= refer %>');
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
