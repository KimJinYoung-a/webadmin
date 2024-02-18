<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%

dim empno
empno = session("ssBctSn")

dim mode
dim divcd, startday, endday, totalvacationday, ishalfvacation, vHalfGubun, workAgent, callNum
Dim authstate
dim totalday
dim masteridx
dim detailidx
dim registerid, registerempno

dim oVacation

dim i, sql
Dim sReturnUrl

mode 				= requestCheckvar(request("mode"),32)
divcd 				= requestCheckvar(request("divcd"),1)
startday 			= requestCheckvar(request("startday"),10)
endday 				= requestCheckvar(request("endday"),10)
totalvacationday	= requestCheckvar(request("totalvacationday"),4)
ishalfvacation		= requestCheckvar(request("ishalfvacation"),1)
vHalfGubun 			= requestCheckvar(request("halfgubun"),2)
totalday 			= requestCheckvar(request("totalday"),12)
masteridx 			= requestCheckvar(Request("masteridx"),8)
detailidx 			= requestCheckvar(Request("detailidx"),8)

authstate 			= requestCheckvar(Request("ias"),10)
sReturnUrl 			= requestCheckvar(request("hidRU"),100)

workAgent 			= requestCheckvar(request("workAgent"),20)
callNum 			= requestCheckvar(request("callNum"),30)

if authstate = "5" then mode = "denydetail"  '전자결재를 통해서 반려처리된 경우 추가 (2011.05.12 정윤정)
empno = Replace(empno, " ", "")



'==============================================================================
dim oMember


if (mode = "adddetail") then

	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx
	oVacation.FRectsearchKey = " t.empno "
	oVacation.FRectsearchString = empno

	oVacation.GetMasterOne

	if (oVacation.FItemOne.IsAvailableVacation <> "Y") then
		response.write "<script>alert('사용할 수 없는 휴가입니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	if ((ishalfvacation = "Y") and (CLng(totalday) > 1)) then
		response.write "<script>alert('잘못된 반차등록입니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	'' if (ishalfvacation = "Y") then
	'' 	totalday = 0.5
	'' end if

	i = (oVacation.FItemOne.Ftotalvacationday - (oVacation.FItemOne.Fusedvacationday + oVacation.FItemOne.Frequestedday))
	if (CDbl(totalday) > i) then
		response.write "<script>alert('잔여일 수 이상의 휴가를 신청하셨습니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	if ((Left(oVacation.FItemOne.Fstartday,10) > startday) or (Left(oVacation.FItemOne.Fendday,10) < endday)) then
		'response.write "<script>alert('" & (oVacation.FItemOne.Fstartday < startday) & "');</script>"
		response.write "<script>alert('신청할 수 없는 휴가 기간입니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	'휴가는 해를 넘어 신청할 수 없다
	if (Left(startday,4) <> Left(endday,4)) then
		response.write "<script>alert('시작연도과 종료연도는 같아야 합니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	sql = "insert into [db_partner].[dbo].tbl_vacation_detail(masteridx, startday, endday, totalday, statedivcd, deleteyn, registerid, halfgubun, registerempno, workAgent, callNum) " & vbCrlf
	sql = sql & "values(" & CStr(masteridx) & ", '" & startday & " 00:00:01', '" & endday & " 23:59:59', " & CStr(totalday) & ", 'R', 'N', '" & session("ssBctId") & "', '" & vHalfGubun & "', '" & session("ssBctSn") & "', '"& html2db(workAgent) &"','"& html2db(callNum) &"') " & vbCrlf
	dbget.Execute(sql)

	sql = "update [db_partner].[dbo].tbl_vacation_master " & vbCrlf
	sql = sql & "set requestedday = requestedday + " & CStr(totalday) & " " & vbCrlf
	sql = sql & "where empno = '" & empno & "' " & vbCrlf
	sql = sql & "and idx = " & CStr(masteridx) & " " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('등록 되었습니다.');</script>"
	response.write "<script>opener.location.reload(); opener.focus();window.close()</script>"
elseif (mode = "deletedetail") then

	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx
	oVacation.FRectsearchKey = " t.empno "
	oVacation.FRectsearchString = empno

	oVacation.GetMasterOne

	if (oVacation.FItemOne.Fdeleteyn = "Y") then
		response.write "<script>alert('삭제된 데이타는 수정할 수 없습니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	Dim objCmd,returnValue
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppReport_chkExist]( 1, " & detailidx & ")}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
	End With
	returnValue = objCmd(0).Value
	Set objCmd = nothing
	IF 	returnValue = 1 and session("ssBctId") = "" THEN
		response.write "<script>alert('품의서가 등록되었습니다. 아이디로 로그인한 이후에만 삭제할 수 있습니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	END IF

	sql = "update [db_partner].[dbo].tbl_vacation_detail " & vbCrlf
	sql = sql & "set deleteyn = 'Y' " & vbCrlf
	sql = sql & ", approverid = '" & session("ssBctId") & "' " & vbCrlf
	sql = sql & ", approverempno = '" & session("ssBctSn") & "' " & vbCrlf
	sql = sql & ", approveday = getdate() " & vbCrlf
	sql = sql & "where idx = " & CStr(detailidx) & " " & vbCrlf
	sql = sql & "and deleteyn <> 'Y' " & vbCrlf
	dbget.Execute(sql)

	sql = "update m " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	m.requestedday = IsNull((select sum(totalday) from [db_partner].[dbo].tbl_vacation_detail d where d.deleteyn <> 'Y' and d.statedivcd = 'R' and d.masteridx = m.idx), 0) " & vbCrlf
	sql = sql & "	, m.usedvacationday = IsNull((select sum(totalday) from [db_partner].[dbo].tbl_vacation_detail d where d.deleteyn <> 'Y' and d.statedivcd = 'A' and d.masteridx = m.idx), 0) " & vbCrlf
	sql = sql & "from [db_partner].[dbo].tbl_vacation_master m " & vbCrlf
	sql = sql & "where 1 = 1 " & vbCrlf
	sql = sql & "and m.deleteyn <> 'Y' " & vbCrlf
	sql = sql & "and m.idx = " & CStr(masteridx) & " " & vbCrlf
	dbget.Execute(sql)

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppReport_DeleteWith]( 1,"&detailidx&",'"& session("ssBctId") &"')}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
	End With
	returnValue = objCmd(0).Value
	Set objCmd = nothing
	IF 	returnValue =0 THEN
		response.write "<script>alert('품의서 삭제에 문제가 발생했습니다.관리자에게 문의해주세요');</script>"
	END IF

	response.write "<script>alert('삭제되었습니다.');</script>"
	response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
elseif (mode = "deletemaster") then

	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx

	oVacation.GetMasterOne

	if (oVacation.FItemOne.Fdeleteyn = "Y") then
		response.write "<script>alert('삭제된 데이타는 수정할 수 없습니다.');</script>"
		response.write "<script>history.back();</script>"
		response.end
	end if

	sql = "update [db_partner].[dbo].tbl_vacation_master " & vbCrlf
	sql = sql & "set deleteyn = 'Y' " & vbCrlf
	sql = sql & "where idx = " & CStr(masteridx) & " " & vbCrlf
	sql = sql & "and deleteyn <> 'Y' " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('삭제되었습니다.');</script>"
	response.write "<script>location.href = '" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
