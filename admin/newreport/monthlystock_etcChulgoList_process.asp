<%@ language=vbscript %>
<% option explicit %>
<%
''Server.ScriptTimeOut = 60
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode, yyyymm, yyyymmdd

mode = request("mode")
yyyymm = request("yyyymm")

dim sqlStr, resultrows
if mode="etcavgprc" then

	yyyymmdd = yyyymm + "-01"
	if (DateDiff("m", yyyymmdd, Now()) >1) then
		response.write "지난달까지만 적용가능합니다."
		dbget.close()	:	response.End
	end if

    sqlStr = " exec [db_summary].[dbo].[sp_Ten_monthly_EtcChulgoList_Apply_avgBuyPrice] '" & yyyymm & "' "
    dbget.execute sqlStr

	response.write "<script>alert('적용 되었습니다.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End
else
	response.write "mode=" + mode
	dbget.close()	:	response.End
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
