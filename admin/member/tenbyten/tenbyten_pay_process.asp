<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenPayCls.asp" -->
<%
Dim objCmd, returnValue
Dim sEmpNo, sYear, sMonth, sDay,sEndDay, defaultpay,iposit_sn, predefaultpay
Dim iStartWork(36), iEndWork(36), iBreakTimeS(36),iBreakTimeE(36), iworktype(36),iouttime(36)
Dim iSH(36), iSM(36), iEH(36), iEM(36), iBSH(36), iBSM(36), iBEH(36), iBEM(36) , iOH(36), iOM(36)
Dim iWT(36), ieWT(36), inWT(36), ihWT(36), iwhWT(40), iVT(36)
DIm iworktime,iextendtime,inighttime,iholidaytime,iwholidaytime 
Dim mtimepay,mextendpay,mnightpay,mholidaypay,mwholidaypay,mfoodpay,mjobpay,moutstandingpay,mlongtimepay,maddpay
Dim mtotpay,mnpensionpay,mhealthinspay ,mrecuinspay,munempinspay,mtaxtotpay,mrealtotpay,adminid, myearpay, mbonuspay
Dim intLoop, sMode, ino, ipaystate,menupos
dim workday
dim dSPaydate, dEPaydate, dSPayDay, dEPayDay, dDay, iVer,dFullPayDate
dim ireworktime,ireextime,irenttime,irehdtime,mretimepay,mreextendpay,mrenightpay,mreholidaypay,mrefoodpay,mretotpay,ireworkday 
dim ireextendtime,irenighttime,ireholidaytime
sMode		= requestCheckvar(request("hidM"),1)
sEmpNo 		= requestCheckvar(request("hidEN"),14)
ino			= requestCheckvar(request("ino"),10)
sYear 		= requestCheckvar(request("hidYear"),4)
sMonth		= requestCheckvar(request("hidMonth"),2) 
sEndDay 	= requestCheckvar(request("hidEday"),2)
ipaystate	= requestCheckvar(request("hidS"),1)
menupos		= requestCheckvar(request("menupos"),10)
defaultpay	= requestCheckvar(request("hidDP"),20)
predefaultpay= requestCheckvar(request("hidPDP"),20)
iposit_sn	= requestCheckvar(request("hidPSN"),4)

dSPaydate = requestCheckvar(request("hidSPdate"),10)
dEPaydate = requestCheckvar(request("hidEPdate"),10)
dSPayDay= requestCheckvar(request("hidSPday"),2)
dEPayDay= requestCheckvar(request("hidEPday"),2)
iVer= requestCheckvar(request("hidVer"),2)
			
		
			
mnpensionpay=0
mhealthinspay = 0
mrecuinspay= 0
munempinspay= 0
mtaxtotpay= 0
mrealtotpay= 0
adminid= session("ssBctId")

SELECT CASE sMode
CASE "D"  '일별근무시간 입력
Dim dberr
dbget.beginTrans
dberr =0
dFullPayDate =  dSPaydate
For intLoop = 0 To sEndDay
  iSH(intLoop) = requestCheckvar(request("iSH"&intLoop),2)
  iSM(intLoop) = requestCheckvar(request("iSM"&intLoop),2)
  iEH(intLoop) = requestCheckvar(request("iEH"&intLoop),2)
  iEM(intLoop) = requestCheckvar(request("iEM"&intLoop),2)
  iBSH(intLoop) = requestCheckvar(request("iBSH"&intLoop),2)
  iBSM(intLoop) = requestCheckvar(request("iBSM"&intLoop),2)
  iBEH(intLoop) = requestCheckvar(request("iBEH"&intLoop),2)
  iBEM(intLoop) = requestCheckvar(request("iBEM"&intLoop),2)
  iOH(intLoop)	= requestCheckvar(request("iOH"&intLoop),2)
  iOM(intLoop)	= requestCheckvar(request("iOM"&intLoop),2)
  iWT(intLoop) = requestCheckvar(request("iWT"&intLoop),5)
  ieWT(intLoop) = requestCheckvar(request("ieWT"&intLoop),5)
  inWT(intLoop) = requestCheckvar(request("inWT"&intLoop),5)
  ihWT(intLoop) = requestCheckvar(request("ihWT"&intLoop),5)
  iwhWT(intLoop) = requestCheckvar(request("iwhWT"&intLoop),5)
  iVT(intLoop) = requestCheckvar(request("iVT"&intLoop),5)
  
  IF iSH(intLoop) = "" THEN iSH(intLoop) = 0
  IF iSM(intLoop) = "" THEN iSM(intLoop) = 0
  IF iEH(intLoop) = "" THEN iEH(intLoop) = 0
  IF iEM(intLoop) = "" THEN iEM(intLoop) = 0
  IF iBSH(intLoop) = "" THEN iBSH(intLoop) = 0
  IF iBSM(intLoop) = "" THEN iBSM(intLoop) = 0
  IF iBEH(intLoop) = "" THEN iBEH(intLoop) = 0
  IF iBEM(intLoop) = "" THEN iBEM(intLoop) = 0
  IF iOH(intLoop) = "" THEN iOH(intLoop) = 0
  IF iOM(intLoop) = "" THEN iOM(intLoop) = 0
  IF iWT(intLoop) 	= "" THEN iWT(intLoop) = 0
  IF ieWT(intLoop)  = "" THEN ieWT(intLoop) = 0
  IF inWT(intLoop)  = "" THEN inWT(intLoop) = 0
  IF ihWT(intLoop)  = "" THEN ihWT(intLoop) = 0
  IF iwhWT(intLoop) = "" THEN iwhWT(intLoop) = 0
  IF iVT(intLoop) 	= "" THEN iVT(intLoop) = 0
 		
	iStartWork(intLoop) =  iSH(intLoop)*60 + iSM(intLoop)
	iEndWork(intLoop)  =  iEH(intLoop)*60 + iEM(intLoop)
	iBreakTimeS(intLoop) =  iBSH(intLoop)*60 + iBSM(intLoop)
	iBreakTimeE(intLoop) =  iBEH(intLoop)*60 + iBEM(intLoop)
	iouttime(intLoop)	= iOH(intLoop)*60 + iOM(intLoop)
	iworktype(intLoop) =   requestCheckvar(request("selWH"&intLoop),1)
	
			
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call db_partner.[dbo].[sp_Ten_user_dailypay_SetData]('"&sEmpNo&"','"&dFullPayDate&"', "&iStartWork(intLoop)&" ,"&	iEndWork(intLoop) &_
					" , "&iBreakTimeS(intLoop)&", "&iBreakTimeE(intLoop)&" ,"&iouttime(intLoop)&", "&iworktype(intLoop)&" , "&fnSetMinuteFromTimeForm(iWT(intLoop))&", "&fnSetMinuteFromTimeForm(ieWT(intLoop))&_
					", "&fnSetMinuteFromTimeForm(inWT(intLoop))&", "&fnSetMinuteFromTimeForm(ihWT(intLoop))&", "&fnSetMinuteFromTimeForm(iwhWT(intLoop))&",'"&adminid&"',"&ipaystate&", "&fnSetMinuteFromTimeForm(iVT(intLoop))&")}"

		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue <> "1" THEN	dberr = dberr + 1 
		dFullPayDate =  dateadd("d",1,dFullPayDate)  
Next

'추가 주휴수당 있을경우
iwhWT(40) = requestCheckvar(request("iwhWT40"),5)
 IF iwhWT(40) = "" THEN iwhWT(40) = 0

  IF iwhWT(40)<> "0" THEN
  	iwhWT(40) = fnSetMinuteFromTimeForm(iwhWT(40))

  	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call db_partner.[dbo].[sp_Ten_user_dailypay_SetData]('"&sEmpNo&"','"&sYear&"-"&Format00(2,sMonth)&"-32', 0 ,0" &_
					" ,0, 0,0, 0, 0,0,0,0, "&iwhWT(40)&",'"&adminid&"',"&ipaystate&",0)}"

		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue <> "1" THEN	dberr = dberr + 1
  END IF


	iworktime	= fnSetMinuteFromTimeForm(requestCheckvar(request("totWT"),20)) + fnSetMinuteFromTimeForm(requestCheckvar(request("totwhWT"),20))+fnSetMinuteFromTimeForm(requestCheckvar(request("totVT"),20))
	iextendtime	= fnSetMinuteFromTimeForm(requestCheckvar(request("toteWT"),20))
	inighttime	= fnSetMinuteFromTimeForm(requestCheckvar(request("totnWT"),20))
	iholidaytime= fnSetMinuteFromTimeForm(requestCheckvar(request("tothWT"),20))
	mtimepay     = round(defaultpay*(iworktime/60) ,0)
	mextendpay   = round((defaultpay*(iextendtime/60))*1.5,0)
	mnightpay    = round((defaultpay*(inighttime/60))*0.5,0)
	mholidaypay	 = round((defaultpay*(iholidaytime/60))*0.5,0)
 
 	ireworktime	= fnSetMinuteFromTimeForm(requestCheckvar(request("totSumWT"),20)) + fnSetMinuteFromTimeForm(requestCheckvar(request("totSumwhWT"),20))+fnSetMinuteFromTimeForm(requestCheckvar(request("totSumVT"),20))
 	
 	 
	ireextendtime	= fnSetMinuteFromTimeForm(requestCheckvar(request("totSumeWT"),20))
	irenighttime	= fnSetMinuteFromTimeForm(requestCheckvar(request("totSumnWT"),20))
	ireholidaytime= fnSetMinuteFromTimeForm(requestCheckvar(request("totSumhWT"),20))
	
	mretimepay     = round(predefaultpay*(ireworktime/60) ,0)
	mreextendpay   = round((predefaultpay*(ireextendtime/60))*1.5,0)
	mrenightpay    = round((predefaultpay*(irenighttime/60))*0.5,0)
	mreholidaypay	 = round((predefaultpay*(ireholidaytime/60))*0.5,0)
		 
 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_user_monthlypay_SetDailyData]('"&sEmpNo&"',"&ino&",'"&sYear&"-"&format00(2,sMonth)&"', "&iworktime&" ,"&iextendtime&" , "&inighttime&", "&iholidaytime&_
						" ,'"&mtimepay&"' ,'"&mextendpay&"' ,'"&mnightpay&"' ,'"&mholidaypay&"','"&adminid&"','"&ireworktime&"','"&ireextendtime&"','"&irenighttime&"','"&ireholidaytime&"','"&mretimepay&"','"&mreextendpay&"','"&mrenightpay&"','"&mreholidaypay&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue <> "1" THEN	dberr = dberr + 1

IF dberr = "0" THEN
		dbget.CommitTrans
	
		%>
	<script language="javascript">
	<!--
	window.opener.location.reload();
	//-->
	</script>
<% 
		Call Alert_move ("등록되었습니다.","pop_worktime.asp?sEN="&sempno&"&selY="&sYear&"&selM="&sMonth&"&ino="&ino)

ELSE
		dbget.RollBackTrans
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
END IF
response.end

CASE "M"
	iworktime= requestCheckvar(request("iWT"),20)
	iextendtime= requestCheckvar(request("iEWT"),20)
	inighttime= requestCheckvar(request("iNWT"),20)
	iholidaytime= requestCheckvar(request("iHDT"),20)
	mtimepay= replace(requestCheckvar(request("iTP"),20),",","")
	mextendpay= replace(requestCheckvar(request("iETP"),20),",","")
	mnightpay= replace(requestCheckvar(request("iNTP"),20),",","")
	mholidaypay= replace(requestCheckvar(request("iHDP"),20),",","")
	mfoodpay= replace(requestCheckvar(request("iFP"),20),",","")
	mjobpay= replace(requestCheckvar(request("iJP"),20),",","")
	moutstandingpay= replace(requestCheckvar(request("iOP"),20),",","")
	mlongtimepay= replace(requestCheckvar(request("iLP"),20),",","")
	maddpay	= replace(requestCheckvar(request("iAP"),20),",","")
	mtotpay= replace(requestCheckvar(request("itotP"),20),",","")
	mrealtotpay= replace(requestCheckvar(request("itotPS"),20),",","")
 
	myearpay = replace(requestCheckvar(request("iYP"),20),",","")
	mbonuspay	= replace(requestCheckvar(request("iBP"),20),",","")

	workday = requestCheckvar(request("totWorkDay"),20)

 
	ireworktime	= requestCheckvar(request("iRWT"),20)
	ireextime		= requestCheckvar(request("iREWT"),20)
	irenttime		= requestCheckvar(request("iRNWT"),20)
	irehdtime		= requestCheckvar(request("iRHDT"),20)
	 
	mretimepay 	= replace(requestCheckvar(request("iRTP"),20),",","")
	mreextendpay= replace(requestCheckvar(request("iRETP"),20),",","")
	mrenightpay = replace(requestCheckvar(request("iRNTP"),20),",","")
	mreholidaypay = replace(requestCheckvar(request("iRHDP"),20),",","")
	mrefoodpay 	= replace(requestCheckvar(request("iRFP"),20),",","")
	mretotpay 	= replace(requestCheckvar(request("iRtotP"),20),",","")
	ireworkday	= requestCheckvar(request("totReWorkDay"),20)

 	IF iworktime = "" THEN iworktime = 0
 	IF iextendtime = "" THEN iextendtime = 0
 	IF inighttime = "" THEN inighttime = 0
 	IF iholidaytime = "" THEN iholidaytime = 0

 	IF mtimepay = "" THEN mtimepay = 0
 	IF mextendpay = "" THEN mextendpay = 0
 	IF mnightpay = "" THEN mnightpay = 0
 	IF mholidaypay = "" THEN mholidaypay = 0
 	IF mfoodpay = "" THEN mfoodpay = 0
 	IF mjobpay = "" THEN mjobpay = 0
 	IF moutstandingpay = "" THEN moutstandingpay = 0
 	IF mlongtimepay = "" THEN mlongtimepay = 0
 	IF maddpay	= "" THEN maddpay = 0
 	IF mtotpay = "" THEN mtotpay = 0

 	IF myearpay	= "" THEN myearpay = 0
 	IF mbonuspay = "" THEN mbonuspay = 0

	IF workday = "" THEN workday = 0
 
	IF ireworktime	=  "" THEN ireworktime	= 0
	IF ireextime		=  "" THEN      ireextime		=0
	IF irenttime		=  "" THEN      irenttime		=0
	IF irehdtime		=  "" THEN      irehdtime		=0
                                    
	IF mretimepay 	=  "" THEN      mretimepay 	=0
	IF mreextendpay=   "" THEN     mreextendpay= 0
	IF mrenightpay =    "" THEN    mrenightpay = 0
	IF mreholidaypay ="" THEN       mreholidaypay =0
	IF mrefoodpay 	=  "" THEN      mrefoodpay 	=0
	IF mretotpay 	=  "" THEN      mretotpay 	=  0
	IF ireworkday	=  "" THEN      ireworkday	=  0


	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_user_monthlypay_SetData]('"&sEmpNo&"',"&ino&",'"&dSPaydate&"','"&dEPaydate&"','"&sYear&"-"&format00(2,sMonth)&"', "&iworktime&" ,"&iextendtime&" , "&inighttime&", "&iholidaytime&_
						" ,'"&mtimepay&"' ,'"&mextendpay&"' ,'"&mnightpay&"' ,'"&mholidaypay&"' ,'"&mfoodpay&"' ,'"&mjobpay&"' ,'"&moutstandingpay&"' ,'"&mlongtimepay&"','"&maddpay&"','"&mtotpay&"'"&_
						" ,'0' ,'0','0','0','0' ,'"&mrealtotpay&"' ,'"&adminid&"',"&ipaystate&",'"&myearpay&"','"&mbonuspay&"', '" & CStr(workday) & "'"&_
						","&ireworktime&","&ireextime&","&irenttime&","&irehdtime&",'"&mretimepay&"','"&mreextendpay&"','"&mrenightpay&"','"&mreholidaypay&"','"&mrefoodpay&"','"&mretotpay&"',"&ireworkday&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue = "1" THEN
			Call Alert_move ("등록되었습니다.","tenbyten_pay_reg.asp?sEN="&sempno&"&selY="&sYear&"&selM="&sMonth&"&ino="&ino&"&menupos="&menupos)
	ELSE
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
	END IF
	response.end
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")

END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
