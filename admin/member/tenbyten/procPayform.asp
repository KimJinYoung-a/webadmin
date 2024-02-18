<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim objCmd, returnValue, mode
Dim empno,iposit_sn,defaultpay, foodpay, jobpay, startdate,enddate,inBreakTime,holidaywd, holidaywdtime,overtime,totPaySum
Dim StartHour(8), StartMinute(8), EndHour(8), EndMinute(8), BreakSHour(8), BreakSMinute(8) , BreakEHour(8), BreakEMinute(8) ,Worktype(8)
Dim adminid, intLoop
 Dim ino, placekind, jobkind
 dim department_id, posit_name, departmentFullName

mode = requestCheckvar(request("hidMode"),6)
ino =requestCheckvar(request("ino"),10)
IF ino = "" THEN ino =0
empno 	= requestCheckvar(request("hidEN"),14)
iposit_sn 	= requestCheckvar(request("hidPSN"),10)
department_id =  requestCheckvar(request("hidDid"),10)
posit_name 	= requestCheckvar(request("hidPSNm"),120)
departmentFullName =  requestCheckvar(request("hidDPNm"),128)

if department_id ="" then department_id = 0
defaultpay= requestCheckvar(request("iHP"),15)
foodpay	= requestCheckvar(request("iEP"),15)
jobpay	= requestCheckvar(request("iJP"),15)
IF jobpay =""  THEN jobpay = 0
totPaySum = requestCheckvar(request("itotp"),15)
IF totPaySum ="" THEN totPaySum = 0

 startdate = requestCheckvar(request("dSD"),10)
 enddate = requestCheckvar(request("dED"),10)

jobkind = requestCheckvar(request("jobkind"),10)
placekind = requestCheckvar(request("placekind"),10)

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
overtime		= overtime*60

	For intLoop = 1 To 7
	StartHour(intLoop) 		= requestCheckvar(request("iSH"&intLoop),2)
	StartMinute(intLoop)  	= requestCheckvar(request("iSM"&intLoop),2)
	EndHour(intLoop)       	= requestCheckvar(request("iEH"&intLoop),2)
	EndMinute(intLoop)      = requestCheckvar(request("iEM"&intLoop),2)
	BreakSHour(intLoop)     = requestCheckvar(request("iBSH"&intLoop),2)
	BreakSMinute(intLoop)   = requestCheckvar(request("iBSM"&intLoop),2)
	BreakEHour(intLoop)     = requestCheckvar(request("iBEH"&intLoop),2)
	BreakEMinute(intLoop)    = requestCheckvar(request("iBEM"&intLoop),2)
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

'// 모드별 분기
Select Case mode
 	Case "modify"
		'계약정보 생성
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call  db_partner.[dbo].[sp_Ten_user_defaultpay_SetData]("&ino&",'"&empno&"','"&startdate&"','"&enddate&"','"&defaultpay&"','"&foodpay&"','"&jobpay&"','"&inBreakTime&"',"&overtime&""&_
							","&(StartHour(1)*60+StartMinute(1))&","&(EndHour(1)*60+EndMinute(1))&","&(BreakSHour(1)*60+BreakSMinute(1))&","&(BreakEHour(1)*60+BreakEMinute(1))&","&(StartHour(2)*60+StartMinute(2))&","&(EndHour(2)*60+EndMinute(2))&","&(BreakSHour(2)*60+BreakSMinute(2))&","&(BreakEHour(2)*60+BreakEMinute(2))&_
							","&(StartHour(3)*60+StartMinute(3))&","&(EndHour(3)*60+EndMinute(3))&","&(BreakSHour(3)*60+BreakSMinute(3))&","&(BreakEHour(3)*60+BreakEMinute(3))&","&(StartHour(4)*60+StartMinute(4))&","&(EndHour(4)*60+EndMinute(4))&","&(BreakSHour(4)*60+BreakSMinute(4))&","&(BreakEHour(4)*60+BreakEMinute(4))&_
							","&(StartHour(5)*60+StartMinute(5))&","&(EndHour(5)*60+EndMinute(5))&","&(BreakSHour(5)*60+BreakSMinute(5))&","&(BreakEHour(5)*60+BreakEMinute(5))&","&(StartHour(6)*60+StartMinute(6))&","&(EndHour(6)*60+EndMinute(6))&","&(BreakSHour(6)*60+BreakSMinute(6))&","&(BreakEHour(6)*60+BreakEMinute(6))&_
							","&(StartHour(7)*60+StartMinute(7))&","&(EndHour(7)*60+EndMinute(7))&","&(BreakSHour(7)*60+BreakSMinute(7))&","&(BreakEHour(7)*60+BreakEMinute(7))&_
							","&worktype(1)&","&worktype(2)&","&worktype(3)&","&worktype(4)&","&worktype(5)&","&worktype(6)&","&worktype(7)&_
							","&holidaywdtime&",'"&totPaySum&"','"&adminid&"','"&iposit_sn&"','"&department_id&"','"&posit_name&"','"&departmentFullName&"','"&jobkind&"','"&placekind&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing

		IF returnValue ="2" THEN
			Call Alert_return ("중복된 계약일이 존재합니다. 계약일을 다시 설정해주세요")
		ELSEIF returnValue ="1" THEN
				Call Alert_move ("등록되었습니다.", "pop_payform.asp?sEN="&empno&"&ino="&ino)
		ELSE
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.")
		END IF

	Case "delete"
		'계약정보 삭제
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call  db_partner.[dbo].[sp_Ten_user_defaultpay_DelData]("&ino&",'"&empno&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing

		IF returnValue ="2" THEN
			Call Alert_return ("선택된 계약서가 없습니다. 확인 후 다시 시도해주세요.")
		ELSEIF returnValue ="1" THEN
				Call Alert_Close ("삭제되었습니다.")
		ELSE
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.")
		END IF

End Select


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->