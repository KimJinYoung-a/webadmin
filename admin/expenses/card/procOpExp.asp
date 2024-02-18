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

'--����üũ---------------------------------------------------------- 
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

''2012/02/21������ �߰� // Erp ���� �� ���� �Ұ�
Dim istate : istate = fnGetOpExpState(iOpExpIdx,dyyyymm,iOpExpPartIdx)

if (sMode<>"C") and (blnInOut) and (istate>9) then
    Call Alert_return ("erp ���۵� �ڷ�� [��� ����]�� �Է�/���� �� �� �����ϴ�.") 
response.end
end if
'--------------------------------------------------------------------
     	 
SELECT CASE sMode 
Case "U"
	'//����üũ
 	IF  fnCheckAuth = 0 THEN
 		Call Alert_return ("���������� �����ϴ�. Ȯ�� �� �ٽ� �õ����ּ���") 
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
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1") 
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
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1") 
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
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.2") 
	END IF
END IF
	dbget.CommitTrans
	response.redirect sReturnURL&"?selY="&dSYear&"&selM="&dSMonth&"&selP="&iOpExpPartIdx&"&iCP="&iCurrpage&"&menupos="&menupos 
	response.end 
Case "D" 
'//����üũ
 	IF  fnCheckAuth = 0 THEN
 		Call Alert_return ("���������� �����ϴ�. Ȯ�� �� �ٽ� �õ����ּ���") 
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
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1") 
	END IF
 
	response.redirect sReturnURL&"?selY="&year(dyyyymm)&"&selM="&month(dyyyymm)&"&selP="&iOpExpPartIdx&"&menupos="&menupos 
	response.end 	
Case "R" 
'//����üũ
 	IF  fnCheckAuth = 0 THEN
 		Call Alert_return ("���� ������ �����ϴ�. Ȯ�� �� �ٽ� �õ����ּ���") 
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
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1") 
	END IF
 
	response.redirect sReturnURL&"?selY="&year(dyyyymm)&"&selM="&month(dyyyymm)&"&selP="&iOpExpPartIdx&"&menupos="&menupos 
	response.end 	
CASE "C"	'���� ����ó��

'//����üũ
	IF sReturnURL = "" THEN
	 	IF  fnCheckAuth = 0 THEN
	 		Call Alert_return ("ó�������� �����ϴ�. Ȯ�� �� �ٽ� �õ����ּ���") 
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
		Call Alert_move("ó���Ǿ����ϴ�","index.asp?selSY="&dSYear&"&selSM="&dSMonth&"&selP="&iOpExpPartIdx&"&menupos="&menupos)	
		END IF
	ELSE	
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.2") 
	END IF
response.end
CASE "T"	'�������� ���� 
 
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
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1") 
	END IF 
	'response.redirect sReturnURL&"?selY="&dSYear&"&selM="&dSMonth&"&selP="&iOpExpPartIdx&"&menupos="&menupos  ''2015/10/21 ������ ����
	response.redirect sReturnURL&"?selSY="&request("selSY")&"&selSM="&request("selSM")&"&selEY="&request("selEY")&"&selEM="&request("selEM")&"&dedTp="&request("dedTp")&"&bizNo="&request("bizNo")&"&selP="&iOpExpPartIdx&"&menupos="&menupos 
	response.end 	
	
CASE "S"	'û���ϵ��
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
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.2") 
	END IF
	
	dbget.CommitTrans  
	response.redirect sReturnURL&"?selSY="&dSYear&"&selSM="&dSMonth&"&selP="&iOpExpPartIdx&"&menupos="&menupos 
	response.end 	

CASE "M"	'���� ���Ӹ� ����  2015/12/14 û���ϵ�� ���μ��� ����?

'//����üũ
	IF sReturnURL = "" THEN
	 	IF  fnCheckAuth = 0 THEN
	 		Call Alert_return ("ó�������� �����ϴ�. Ȯ�� �� �ٽ� �õ����ּ���") 
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
		Call Alert_move("ó���Ǿ����ϴ�","index.asp?selSY="&dSYear&"&selSM="&dSMonth&"&selP="&iOpExpPartIdx&"&menupos="&menupos)	
		END IF
	ELSE	
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.2") 
	END IF
response.end
		
CASE ELSE
	Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.0")
END SELECT
%>
