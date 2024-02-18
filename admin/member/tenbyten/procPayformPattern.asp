<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim objCmd, returnValue, sMode, sMsg
Dim sEmpno,patternseq,part_sn,patternname,defaultpay, foodpay, jobpay,  inBreakTime,holidaywd, holidaywdtime,overtime, totPaySum
Dim StartHour(8), StartMinute(8), EndHour(8), EndMinute(8), BreakSHour(8), BreakSMinute(8) , BreakEHour(8), BreakEMinute(8),Worktype(8)
Dim adminid, intLoop,iDefaultPaySeq
iDefaultPaySeq =requestCheckvar(request("iDPS"),10)

 sMode	= requestCheckvar(request("hidM"),1)
 sEmpno	= requestCheckvar(request("hidEN"),14)
patternseq = requestCheckvar(request("hidPS"),10)
part_sn 	= requestCheckvar(request("part_sn"),10)
patternname= requestCheckvar(request("sPN"),60)
defaultpay= requestCheckvar(request("iHP"),15)
foodpay	= requestCheckvar(request("iEP"),15)
jobpay	= requestCheckvar(request("iJP"),15)
IF jobpay =""  THEN jobpay = 0
totPaySum = requestCheckvar(request("itotp"),15)
IF totPaySum ="" THEN totPaySum = 0

IF requestCheckvar(request("selWH1"),1) = "3" THEN
	holidaywd = 1
ELSEIF	requestCheckvar(request("selWH2"),1) = "3" THEN
	holidaywd = 2
ELSEIF	requestCheckvar(request("selWH3"),1) = "3" THEN
	holidaywd = 3
ELSEIF	requestCheckvar(request("selWH4"),1) = "3" THEN
	holidaywd = 4
ELSEIF	requestCheckvar(request("selWH5"),1) = "3" THEN
	holidaywd = 5
ELSEIF	requestCheckvar(request("selWH6"),1) = "3" THEN
	holidaywd =6
ELSEIF requestCheckvar(request("selWH7"),1) = "3" THEN
	holidaywd = 7
END IF
holidaywdtime =  requestCheckvar(request("iWHT"&holidaywd),10)
IF holidaywdtime <> "" THEN
holidaywdtime = split(holidaywdtime,":")(0)*60+split(holidaywdtime,":")(1)
ELSE
holidaywdtime = 0
END IF
inBreakTime	=requestCheckvar(request("blnBT"),10)
IF inBreakTime = "" THEN inBreakTime = 0
overtime		= requestCheckvar(request("iOT"),10)
IF overtime = "" THEN overtime = 0
overtime		= overtime*60

	For intLoop = 1 To 7
	StartHour(intLoop) 		= requestCheckvar(request("iSH"&intLoop),2)
	StartMinute(intLoop)  	= requestCheckvar(request("iSM"&intLoop),2)
	EndHour(intLoop)       	= requestCheckvar(request("iEH"&intLoop),2)
	EndMinute(intLoop)      = requestCheckvar(request("iEM"&intLoop),2)
	BreakSHour(intLoop)     = requestCheckvar(request("iBSH"&intLoop),2)
	BreakSMinute(intLoop)   = requestCheckvar(request("iBSM"&intLoop),2)
	BreakEHour(intLoop)     = requestCheckvar(request("iBEH"&intLoop),2)
	BreakEMinute(intLoop)   = requestCheckvar(request("iBEM"&intLoop),2)
	Worktype(intLoop)		= requestCheckvar(request("selWH"&intLoop),1)

	IF 	StartHour(intLoop)  = "" THEN 	StartHour(intLoop)  = 0
	IF 	StartMinute(intLoop)  = "" THEN 	StartMinute(intLoop)  = 0
	IF 	EndHour(intLoop)  = "" THEN 	EndHour(intLoop)  = 0
	IF 	EndMinute(intLoop)  = "" THEN 	EndMinute(intLoop)  = 0
	IF 	BreakSHour(intLoop)  = "" THEN 	BreakSHour(intLoop)  = 0
	IF 	BreakSMinute(intLoop)  = "" THEN 	BreakSMinute(intLoop)  = 0
	IF 	BreakEHour(intLoop)  = "" THEN 	BreakEHour(intLoop)  = 0
	IF 	BreakEMinute(intLoop)  = "" THEN 	BreakEMinute(intLoop)  = 0
	Next

adminid	= session("ssBctId")

'계약정보 생성
Set objCmd = Server.CreateObject("ADODB.COMMAND")
SELECT CASE sMode
CASE "I"
	With objCmd
	.ActiveConnection = dbget
	.CommandType = adCmdText
	.CommandText = "{?= call  db_partner.[dbo].[sp_Ten_user_defaultpay_pattern_Insert]("&part_sn&",'"&patternname&"','"&defaultpay&"','"&foodpay&"','"&jobpay&"','"&inBreakTime&"',"&overtime&""&_
				","&(StartHour(1)*60+StartMinute(1))&","&(EndHour(1)*60+EndMinute(1))&","&(BreakSHour(1)*60+BreakSMinute(1))&","&(BreakEHour(1)*60+BreakEMinute(1))&","&(StartHour(2)*60+StartMinute(2))&","&(EndHour(2)*60+EndMinute(2))&","&(BreakSHour(2)*60+BreakSMinute(2))&","&(BreakEHour(2)*60+BreakEMinute(2))&_
					","&(StartHour(3)*60+StartMinute(3))&","&(EndHour(3)*60+EndMinute(3))&","&(BreakSHour(3)*60+BreakSMinute(3))&","&(BreakEHour(3)*60+BreakEMinute(3))&","&(StartHour(4)*60+StartMinute(4))&","&(EndHour(4)*60+EndMinute(4))&","&(BreakSHour(4)*60+BreakSMinute(4))&","&(BreakEHour(4)*60+BreakEMinute(4))&_
					","&(StartHour(5)*60+StartMinute(5))&","&(EndHour(5)*60+EndMinute(5))&","&(BreakSHour(5)*60+BreakSMinute(5))&","&(BreakEHour(5)*60+BreakEMinute(5))&","&(StartHour(6)*60+StartMinute(6))&","&(EndHour(6)*60+EndMinute(6))&","&(BreakSHour(6)*60+BreakSMinute(6))&","&(BreakEHour(6)*60+BreakEMinute(6))&_
					","&(StartHour(7)*60+StartMinute(7))&","&(EndHour(7)*60+EndMinute(7))&","&(BreakSHour(7)*60+BreakSMinute(7))&","&(BreakEHour(7)*60+BreakEMinute(7))&_
					","&worktype(1)&","&worktype(2)&","&worktype(3)&","&worktype(4)&","&worktype(5)&","&worktype(6)&","&worktype(7)&_
					","&holidaywdtime&",'"&totPaySum&"','"&adminid&"')}"
	.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
	.Execute, , adExecuteNoRecords
	End With
    returnValue = objCmd(0).Value
    sMsg = "등록되었습니다."
 Case   "U"
 	With objCmd
	.ActiveConnection = dbget
	.CommandType = adCmdText
	.CommandText = "{?= call  db_partner.[dbo].[sp_Ten_user_defaultpay_pattern_Update]("&patternSeq&","&part_sn&",'"&patternname&"','"&defaultpay&"','"&foodpay&"','"&jobpay&"','"&inBreakTime&"',"&overtime&""&_
					","&(StartHour(1)*60+StartMinute(1))&","&(EndHour(1)*60+EndMinute(1))&","&(BreakSHour(1)*60+BreakSMinute(1))&","&(BreakEHour(1)*60+BreakEMinute(1))&","&(StartHour(2)*60+StartMinute(2))&","&(EndHour(2)*60+EndMinute(2))&","&(BreakSHour(2)*60+BreakSMinute(2))&","&(BreakEHour(2)*60+BreakEMinute(2))&_
					","&(StartHour(3)*60+StartMinute(3))&","&(EndHour(3)*60+EndMinute(3))&","&(BreakSHour(3)*60+BreakSMinute(3))&","&(BreakEHour(3)*60+BreakEMinute(3))&","&(StartHour(4)*60+StartMinute(4))&","&(EndHour(4)*60+EndMinute(4))&","&(BreakSHour(4)*60+BreakSMinute(4))&","&(BreakEHour(4)*60+BreakEMinute(4))&_
					","&(StartHour(5)*60+StartMinute(5))&","&(EndHour(5)*60+EndMinute(5))&","&(BreakSHour(5)*60+BreakSMinute(5))&","&(BreakEHour(5)*60+BreakEMinute(5))&","&(StartHour(6)*60+StartMinute(6))&","&(EndHour(6)*60+EndMinute(6))&","&(BreakSHour(6)*60+BreakSMinute(6))&","&(BreakEHour(6)*60+BreakEMinute(6))&_
					","&(StartHour(7)*60+StartMinute(7))&","&(EndHour(7)*60+EndMinute(7))&","&(BreakSHour(7)*60+BreakSMinute(7))&","&(BreakEHour(7)*60+BreakEMinute(7))&_
					","&worktype(1)&","&worktype(2)&","&worktype(3)&","&worktype(4)&","&worktype(5)&","&worktype(6)&","&worktype(7)&_
					","&holidaywdtime&",'"&totPaySum&"','"&adminid&"')}"
	.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
	.Execute, , adExecuteNoRecords
	End With
    returnValue = objCmd(0).Value
    sMsg = "수정되었습니다."
 Case   "D"
 	With objCmd
	.ActiveConnection = dbget
	.CommandType = adCmdText
	.CommandText = "{?= call  db_partner.[dbo].[sp_Ten_user_defaultpay_pattern_Delete]("&patternSeq&",'"&adminid&"')}"
	.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
	.Execute, , adExecuteNoRecords
	End With
    returnValue = objCmd(0).Value
    sMsg = "삭제되었습니다."
Case ELSE
 	returnValue = "0"
END SELECT

Set objCmd = nothing

IF returnValue ="1" THEN
		Call Alert_move (sMsg, "pop_payform_pattern.asp?sEN="&sEmpno&"&iPS="&patternseq&"&iDPS="&iDefaultPaySeq)
ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.")
END IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->