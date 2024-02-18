<%@ language=vbscript %>
<% option explicit  %>
<%
'###########################################################
' Description : 운영비관리    리스트
' History : 2011.06.03 정윤정 생성
'			2020.07.27 한용민 수정(삭제기능추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/expenses/OpExpCls.asp"-->
<% 
Dim objCmd, returnValue, sMode, dYear,dMonth,dDay, dYYYYMM ,dyyyymmDD,doldyyyymm, menupos, iPartTypeIdx,iCurrpage
Dim iOpExpDailyIdx,iOpExpPartIdx,iarap_cd,mExp, minExp,mOutExp,sOpExpObj,sDetailConts,sadminID, arrArap, blnInOut
Dim sbizsection_cd, msupExp, mvatExp, sAuthNo, iOpExpIdx   ,state, clsOpExp, iAuthValue, dSYear,dSMonth, sReturnURL, sqlstr
	sMode		= requestCheckvar(Request("hidM"),10) 
	dDay		= requestCheckvar(request("iD"),2) 
	dYear		= requestCheckvar(request("selY"),4) 
	dMonth		= requestCheckvar(request("selM"),2)  
	dyyyymm		= dYear&"-"&Format00(2,dMonth)
	dyyyymmDD	= dyyyymm&"-"&Format00(2,dDay)
	doldyyyymm = requestCheckvar(request("dOYM"),7)
	state			= requestCheckvar(Request("hidS"),4)	
	iOpExpDailyIdx= requestCheckvar(Request("hidOED"),10)	
	iOpExpIdx	= requestCheckvar(Request("hidOE"),10)
	iPartTypeIdx= requestCheckvar(request("selPT"),10) 
	iOpExpPartIdx = requestCheckvar(Request("selP"),10)	
	sbizsection_cd = requestCheckvar(Request("sBcd"),10) 
	arrArap= split(requestCheckvar(Request("selA"),20),"^") 
 
IF ubound(arrArap)>=0 THEN 
	iarap_cd = arrArap(0)
	blnInOut = arrArap(1)
END IF
mExp		= requestCheckvar(Request("mExp"),10)
IF blnInOut   THEN '사용금액일때 
	mOutExp = mExp	 
ELSE	
	minExp = mExp
END IF
msupExp= requestCheckvar(Request("msupExp"),10)
mvatExp= requestCheckvar(Request("mvatExp"),10)
sAuthNo= requestCheckvar(Request("sAN"),30)
sOpExpObj	= requestCheckvar(Request("sO"),30)	
sDetailConts= requestCheckvar(Request("sDC"),200)	
sadminID	=  session("ssBctId")
iCurrpage=requestCheckvar(Request("iCP"),10)
menupos= requestCheckvar(Request("menupos"),10)
sReturnURL =    requestCheckvar(Request("hidRU"),100)
IF minExp = "" THEN	 minExp = 0 
IF mOutExp = "" THEN mOutExp = 0
IF msupExp = "" THEN msupExp = 0
dSYear		= requestCheckvar(request("selSY"),4) 
dSMonth		= requestCheckvar(request("selSM"),2)

'--권한체크---------------------------------------------------------- 
Function fnCheckAuth 
	Dim blnAdmin,blnWorker, blnReg
	
	set clsOpExp = new OpExp 
	clsOpExp.Fyyyymm  = dyyyymm
	clsOpExp.FMode  = sMode
	clsOpExp.FOpExpPartIdx  = iOpExpPartIdx
	clsOpExp.FadminID 		= session("ssBctId") 
	blnWorker = clsOpExp.fnGetOpExpAuth 
	set clsOpExp = nothing	
	
	blnAdmin = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))
	IF ( blnWorker  = 1 OR blnAdmin )  THEN
	blnReg =  1
	ELSEIF blnWorker = 2 THEN
	blnReg = 2
	ELSE
	blnReg =  0
	END IF
	fnCheckAuth = blnReg
	
End Function 
'--------------------------------------------------------------------
Function fnGetOpExpState(iOpExpIdx,dyyyymm,iOpExpPartIdx)
    Dim clsOpExp
    set clsOpExp = new OpExp
    clsOpExp.FOpExpidx = iOpExpIdx
    clsOpExp.Fyyyymm   = dyyyymm
    clsOpExp.FOpExpPartIdx = iOpExpPartIdx
    clsOpExp.fnGetOpExpMonthlyData
    fnGetOpExpState = clsOpExp.FState
    set clsOpExp = Nothing
End Function

''2012/02/21서동석 추가 // Erp 전송 후 수정 불가
Dim istate : istate = fnGetOpExpState(iOpExpIdx,dyyyymm,iOpExpPartIdx)

if (sMode<>"C") and (blnInOut) and (istate>9) then
    Call Alert_return ("erp 전송된 자료는 [사용 내역]을 입력/수정 할 수 없습니다.") 
response.end
end if
'--------------------------------------------------------------------
     	 
SELECT CASE sMode
Case "I"    
	'//권한체크
 	IF  fnCheckAuth = 2 THEN  
 		Call Alert_return ("선택하신 달의 이전달운영비가 작성중입니다. 이전달 운영비 작성완료 후 작성해주세요") 
	response.end
 	ELSEIF  fnCheckAuth = 0 THEN 
 		Call Alert_return ("등록권한이 없습니다. 확인 후 다시 시도해주세요") 
	response.end
 	END IF
 	
	dbget.beginTrans
 	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpDaily_Insert]('"&dyyyymmDD&"',"&iOpExpPartIdx&","&iarap_Cd&",'"&minExp&"','"&mOutExp&"','"&msupExp&"','"&mvatExp&"', '"&sAuthNo&"' ,'"&sOpExpObj&"','"&sDetailConts&"','"&sbizsection_Cd&"','"&sadminID&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	 
	IF returnValue <>  1  THEN  
		 dbget.RollBackTrans
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthly_setData]('"&dyyyymm&"','',"&iOpExpPartIdx&",'"&sadminID&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	 
	IF returnValue <>  1  THEN  
		 dbget.RollBackTrans
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	dim i
	IF dyyyymm < year(date())&"-"&Format00(2,month(date())) THEN
		For i= 1 To datediff("m",dyyyymm,date())
				Set objCmd = Server.CreateObject("ADODB.COMMAND")   
				With objCmd
					.ActiveConnection = dbget
					.CommandType = adCmdText  		
					.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthly_updateNextMonth]('"&dateadd("m",i,dyyyymm)&"',"&iOpExpPartIdx&",'"&sadminID&"')}"							 
					.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
					.Execute, , adExecuteNoRecords
					End With	
				    returnValue = objCmd(0).Value	  
			Set objCmd = nothing	 
			IF returnValue <>  1  THEN  
				 dbget.RollBackTrans
				Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
			END IF
		Next
	END IF
	
	dbget.CommitTrans
	 
	response.redirect "regOpExp.asp?selY="&year(dyyyymm)&"&selM="&month(dyyyymm)&"&selPT="&iPartTypeIdx&"&selP="&iOpExpPartIdx&"&iCP="&iCurrpage&"&menupos="&menupos
	response.end 
Case "U"
	'//권한체크
 	IF  fnCheckAuth = 0 THEN
 		Call Alert_return ("수정권한이 없습니다. 확인 후 다시 시도해주세요") 
	response.end
 	END IF 
    
dbget.beginTrans
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpDaily_Update]("&iOpExpDailyIdx&",'"&dyyyymmDD&"',"&iOpExpPartIdx&","&iarap_cd&",'"&minExp&"','"&mOutExp&"','"&msupExp&"','"&mvatExp&"', '"&sAuthNo&"','"&sOpExpObj&"','"&sDetailConts&"','"&sbizsection_cd&"','"&sadminID&"')}"							 					 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	IF returnValue <>  1  THEN  
		 dbget.RollBackTrans
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthly_setData]('"&dyyyymm&"','"&doldyyyymm&"',"&iOpExpPartIdx&",'"&sadminID&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	 
	IF returnValue <>  1  THEN  
		 dbget.RollBackTrans
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.2") 
	END IF
	
		IF dyyyymm < year(date())&"-"&Format00(2,month(date())) THEN
		For i= 1 To datediff("m",dyyyymm,date())
				Set objCmd = Server.CreateObject("ADODB.COMMAND")   
				With objCmd
					.ActiveConnection = dbget
					.CommandType = adCmdText  		
					.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthly_updateNextMonth]('"&dateadd("m",i,dyyyymm)&"',"&iOpExpPartIdx&",'"&sadminID&"')}"							 
					.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
					.Execute, , adExecuteNoRecords
					End With	
				    returnValue = objCmd(0).Value	  
			Set objCmd = nothing	 
			IF returnValue <>  1  THEN  
				 dbget.RollBackTrans
				Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
			END IF
		Next
	END IF
	dbget.CommitTrans
	response.redirect "regOpExp.asp?selY="&year(dyyyymm)&"&selM="&month(dyyyymm)&"&selPT="&iPartTypeIdx&"&selP="&iOpExpPartIdx&"&iCP="&iCurrpage&"&menupos="&menupos 
	response.end 
Case "D" 
'//권한체크
 	IF  fnCheckAuth = 0 THEN
 		Call Alert_return ("삭제권한이 없습니다. 확인 후 다시 시도해주세요") 
	response.end
 	END IF
dbget.beginTrans
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpDaily_Delete]("&iOpExpDailyIdx&",'"&sadminID&"')}"							 					 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	IF returnValue <>  1  THEN  
		 dbget.RollBackTrans
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthly_setData]('"&dyyyymm&"','"&doldyyyymm&"',"&iOpExpPartIdx&",'"&sadminID&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	 
	IF returnValue <>  1  THEN  
		 dbget.RollBackTrans
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.2") 
	END IF
	
		IF dyyyymm < year(date())&"-"&Format00(2,month(date())) THEN
		For i= 1 To datediff("m",dyyyymm,date())
				Set objCmd = Server.CreateObject("ADODB.COMMAND")   
				With objCmd
					.ActiveConnection = dbget
					.CommandType = adCmdText  		
					.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthly_updateNextMonth]('"&dateadd("m",i,dyyyymm)&"',"&iOpExpPartIdx&",'"&sadminID&"')}"							 
					.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
					.Execute, , adExecuteNoRecords
					End With	
				    returnValue = objCmd(0).Value	  
			Set objCmd = nothing	 
			IF returnValue <>  1  THEN  
				 dbget.RollBackTrans
				Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
			END IF
		Next
	END IF
	dbget.CommitTrans
	response.redirect "regOpExp.asp?selY="&year(dyyyymm)&"&selM="&month(dyyyymm)&"&selPT="&iPartTypeIdx&"&selP="&iOpExpPartIdx&"&menupos="&menupos 
	response.end 	
CASE "C"	'상태 변경처리

'//권한체크
	IF sReturnURL = "" THEN
	 	IF  fnCheckAuth = 0 THEN
	 		Call Alert_return ("처리권한이 없습니다. 확인 후 다시 시도해주세요") 
		response.end
	 	END IF
	END IF
	
	Dim strMsg
	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthly_setConfirm]("&iOpExpIdx&","&state&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	
 
 
	IF returnValue= 1  THEN   
		IF sReturnURL <> "" THEN
			response.redirect(sReturnURL)
		response.end
		ELSE
		Call Alert_move("처리되었습니다","index.asp?selSY="&dSYear&"&selSM="&dSMonth&"&selPT="&iPartTypeIdx&"&selP="&iOpExpPartIdx&"&menupos="&menupos)	
		END IF
	ELSE	
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.2") 
	END IF

CASE "monthdel"		' 삭제처리
	if iOpExpIdx="" then
		Call Alert_return ("삭제키가 없습니다.")
		response.end
	end if
	if not(C_MngPart or C_ADMIN_AUTH or C_PSMngPart) then
		Call Alert_return ("삭제권한이 없습니다.")
		response.end
	end if

	sqlstr="delete from db_partner.dbo.tbl_OpExpMonthly where opexpidx="& iOpExpIdx &""

	'response.write sqlstr & "<br>"
	dbget.execute sqlstr

	Call Alert_move("처리되었습니다","index.asp?selSY="&dSYear&"&selSM="&dSMonth&"&selPT="&iPartTypeIdx&"&selP="&iOpExpPartIdx&"&menupos="&menupos)	
	response.end

CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>