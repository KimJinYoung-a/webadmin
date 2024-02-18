<%@ language=vbscript %>
<% option explicit  %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/expenses/OpExpCls.asp"-->
<% 
Dim objCmd, returnValue, sMode
Dim dYYYYMM 
Dim iOpExpDailyIdx,iOpExpPartIdx,iarap_cd,mOutExp,sOpExpObj,sDetailConts,sadminID ,blndeducttype
Dim mVatExp, msupExp, msevExp
Dim sbizsection_cd,  sAuthNo
Dim menupos, iCurrpage
Dim iOpExpIdx   ,state
Dim clsOpExp, iAuthValue
Dim dSYear,dSMonth, sReturnURL, sNoSet
Dim arrArap, blnInOut 
Dim chk : chk = request("chk")  
Dim arrIdx 
sMode					= requestCheckvar(Request("hidM"),1)  
state					= requestCheckvar(Request("hidS"),4)	
iOpExpDailyIdx= requestCheckvar(Request("hidOED"),10)	
iOpExpIdx			= requestCheckvar(Request("hidOE"),10) 
iOpExpPartIdx = requestCheckvar(Request("selP"),10)	
sbizsection_cd= requestCheckvar(Request("sBcd"),10) 
iarap_cd			= requestCheckvar(Request("selA"),20)    
sDetailConts	= requestCheckvar(Request("sDC"),200)	
mOutExp	= requestCheckvar(Request("mO"),20)	
msupExp	= requestCheckvar(Request("mSP"),20)	
mVatExp	= requestCheckvar(Request("mV"),20)	
msevExp	= requestCheckvar(Request("mSV"),20)	
blndeducttype	= requestCheckvar(Request("rdoD"),1)	
sadminID			= session("ssBctId")
iCurrpage			= requestCheckvar(Request("iCP"),10)
menupos				= requestCheckvar(Request("menupos"),10)
sReturnURL 		= requestCheckvar(Request("hidRU"),100)
   
dSYear		= requestCheckvar(request("selY"),4) 
dSMonth		= requestCheckvar(request("selM"),2)
dYYYYMM 	= dSYear&"-"&Format00(2,dSMonth) 
sNoSet		= requestCheckvar(Request("hidNS"),1) 

'--권한체크---------------------------------------------------------- 
Function fnCheckAuth 
	Dim blnAdmin,blnWorker, blnReg
	
	set clsOpExp = new OpExp  
	clsOpExp.FOpExpPartIdx  = iOpExpPartIdx
	clsOpExp.FadminID 		= session("ssBctId") 
	blnWorker = clsOpExp.fnGetOpExpPartAuth 
	set clsOpExp = nothing	
	
	blnAdmin = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))
	IF ( blnWorker  = 1 OR blnAdmin )  THEN
	blnReg =  1 
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
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpDailyCard_Update]("&iOpExpDailyIdx&", "&iarap_cd&",'"&sDetailConts&"','"&mOutExp&"','"&mSupExp&"','"&mVatExp&"','"&mSevExp&"','"&sbizsection_cd&"','"&blndeducttype&"','"&sadminID&"')}"							 					 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	IF returnValue <>  1  THEN  
		 dbget.RollBackTrans
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	
IF dYYYYMM <> "" THEN
    Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpDailyCard_SetDate]('"&iOpExpDailyIdx&"','"&dyyyymm&"')}"							 					 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing 
	IF returnValue <>  1  THEN  
		 dbget.RollBackTrans
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
END IF	
IF sNoSet <> "Y" THEN 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthlyCard_setData]('"&dyyyymm&"',"&iOpExpPartIdx&",'"&sadminID&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	 
	IF returnValue <>  1  THEN  
		 dbget.RollBackTrans
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.2") 
	END IF
END IF
	dbget.CommitTrans
	response.redirect sReturnURL&"?selY="&dSYear&"&selM="&dSMonth&"&selP="&iOpExpPartIdx&"&iCP="&iCurrpage&"&menupos="&menupos 
	response.end 
Case "D" 
'//권한체크
 	IF  fnCheckAuth = 0 THEN
 		Call Alert_return ("삭제권한이 없습니다. 확인 후 다시 시도해주세요") 
	response.end
 	END IF
 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpDailyCard_Delete]("&iOpExpDailyIdx&",'"&sadminID&"')}"							 					 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	 
	IF returnValue <>  1  THEN   
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
 
	response.redirect sReturnURL&"?selY="&year(dyyyymm)&"&selM="&month(dyyyymm)&"&selP="&iOpExpPartIdx&"&menupos="&menupos 
	response.end 	
Case "R" 
'//권한체크
 	IF  fnCheckAuth = 0 THEN
 		Call Alert_return ("복구 권한이 없습니다. 확인 후 다시 시도해주세요") 
	response.end
 	END IF
 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpDailyCard_aLive]("&iOpExpDailyIdx&",'"&sadminID&"')}"							 					 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	 
	IF returnValue <>  1  THEN   
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
 
	response.redirect sReturnURL&"?selY="&year(dyyyymm)&"&selM="&month(dyyyymm)&"&selP="&iOpExpPartIdx&"&menupos="&menupos 
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
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthlyCard_setConfirm]("&iOpExpIdx&","&state&")}"							 
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
		Call Alert_move("처리되었습니다","index.asp?selSY="&dSYear&"&selSM="&dSMonth&"&selP="&iOpExpPartIdx&"&menupos="&menupos)	
		END IF
	ELSE	
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.2") 
	END IF
response.end
CASE "T"	'공제여부 변경 
 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthlyCard_setDeduct]("&iOpExpDailyIdx&",'"&blndeducttype&"')}"							 					 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	IF returnValue <>  1  THEN   
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF 
	'response.redirect sReturnURL&"?selY="&dSYear&"&selM="&dSMonth&"&selP="&iOpExpPartIdx&"&menupos="&menupos  ''2015/10/21 서동석 수정
	response.redirect sReturnURL&"?selSY="&request("selSY")&"&selSM="&request("selSM")&"&selEY="&request("selEY")&"&selEM="&request("selEM")&"&dedTp="&request("dedTp")&"&bizNo="&request("bizNo")&"&selP="&iOpExpPartIdx&"&menupos="&menupos 
	response.end 	
	
CASE "S"	'청구일등록
Dim i,AssignedRow
	arrIdx = split(chk,",")
	AssignedRow = 0
dbget.beginTrans 	 

For i = 0  To UBound(arrIdx) 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpDailyCard_SetDate]('"&trim(arrIdx(i))&"','"&dyyyymm&"')}"							 					 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing 

	IF returnValue =  1  THEN   
		AssignedRow=AssignedRow+1
	END IF 
Next
 
	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthlyCard_setData]('"&dyyyymm&"',0,'"&sadminID&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	
	response.write "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthlyCard_setData]('"&dyyyymm&"',0,'"&sadminID&"')}"	
	IF returnValue <>  1  THEN  
		 dbget.RollBackTrans
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.2") 
	END IF
	
	dbget.CommitTrans  
	response.redirect sReturnURL&"?selSY="&dSYear&"&selSM="&dSMonth&"&selP="&iOpExpPartIdx&"&menupos="&menupos 
	response.end 	

CASE "M"	'월별 서머리 생성  2015/12/14 청구일등록 프로세스 없나?

'//권한체크
	IF sReturnURL = "" THEN
	 	IF  fnCheckAuth = 0 THEN
	 		Call Alert_return ("처리권한이 없습니다. 확인 후 다시 시도해주세요") 
		response.end
	 	END IF
	END IF
	
	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthlyCard_setData]('"&dyyyymm&"',0,'"&sadminID&"')}"					 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	
 
 ''rw "db_partner.[dbo].[sp_Ten_OpExpMonthlyCard_setData]('"&dyyyymm&"',0,'"&sadminID&"')"
	IF returnValue= 1  THEN   
		IF sReturnURL <> "" THEN
			response.redirect(sReturnURL)
		response.end
		ELSE
		Call Alert_move("처리되었습니다","index.asp?selSY="&dSYear&"&selSM="&dSMonth&"&selP="&iOpExpPartIdx&"&menupos="&menupos)	
		END IF
	ELSE	
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.2") 
	END IF
response.end
		
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
