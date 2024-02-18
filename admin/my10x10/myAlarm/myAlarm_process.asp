<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode, sqlStr
dim levelAlarmIdx, yyyymmdd, msgdiv, title, subtitle, contents, userlevel, wwwTargetURL, openYN, useYN, regdate, reguserid, lastupdate


mode = requestCheckvar(request("mode"),32)
levelAlarmIdx = requestCheckvar(request("levelAlarmIdx"),32)
yyyymmdd = requestCheckvar(request("yyyymmdd"),32)
title = requestCheckvar(request("title"),32)
subtitle = requestCheckvar(request("subtitle"),32)
contents = requestCheckvar(request("contents"),32)
userlevel = requestCheckvar(request("userlevel"),32)
wwwTargetURL = requestCheckvar(request("wwwTargetURL"),50)
openYN = requestCheckvar(request("openYN"),32)
useYN = requestCheckvar(request("useYN"),32)

reguserid = session("ssBctId")

Select Case mode
	Case "regalarm"

		sqlStr = " insert into db_my10x10.dbo.tbl_myAlarm_by_level(yyyymmdd, msgdiv, title, subtitle, contents, userlevel, wwwTargetURL, openYN, useYN, regdate, reguserid, lastupdate) "
		sqlStr = sqlStr + " values('" + CStr(yyyymmdd) + "', '', '" + CStr(html2db(title)) + "', '" + CStr(html2db(subtitle)) + "', '" + CStr(html2db(contents)) + "', " + CStr(userlevel) + ", '" + CStr(wwwTargetURL) + "', '" + CStr(openYN) + "', '" + CStr(useYN) + "', getdate(), '" + CStr(reguserid) + "', getdate()) "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		response.write "<script language='javascript'>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()
		response.end
	Case "modialarm"

		sqlStr = " update db_my10x10.dbo.tbl_myAlarm_by_level "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " 	yyyymmdd = '" + CStr(yyyymmdd) + "', "
		sqlStr = sqlStr + " 	title = '" + CStr(html2db(title)) + "',"
		sqlStr = sqlStr + " 	subtitle = '" + CStr(html2db(subtitle)) + "',"
		sqlStr = sqlStr + " 	contents = '" + CStr(html2db(contents)) + "',"
		sqlStr = sqlStr + " 	userlevel = '" + CStr(userlevel) + "', "
		sqlStr = sqlStr + " 	wwwTargetURL = '" + CStr(wwwTargetURL) + "',"
		sqlStr = sqlStr + " 	openYN = '" + CStr(openYN) + "', "
		sqlStr = sqlStr + " 	useYN = '" + CStr(useYN) + "',"
		sqlStr = sqlStr + " 	lastupdate = getdate() "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	levelAlarmIdx = " + CStr(levelAlarmIdx) + " "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		response.write "<script language='javascript'>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
		dbget.close()
		response.end
	Case Else
		''
End Select

%>
<script language='javascript'>
// alert('저장되었습니다.');
// location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
