<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<% 
Dim objCmd, returnValue, sMode
Dim sarap_cd,arrarap_cd, iarap_cd, iOpExpAccountIdx,InOutType
Dim intLoop,menupos
   
sMode		= requestCheckvar(Request("hidM"),1)
arrarap_cd		= ReplaceRequestSpecialChar(request("hidccd")) 
iOpExpAccountIdx		= requestCheckvar(Request("hidOEA"),10)	 
InOutType	= requestCheckvar(Request("hidInOut"),1)
menupos= requestCheckvar(Request("menupos"),10)


SELECT CASE sMode
Case "I"  
iarap_cd = split(arrarap_cd,",")  
 	For intLoop = 0 To UBound(iarap_cd)	  
 	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpAccount_Insert]("&trim(iarap_cd(intLoop))&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	 
		     
		IF  returnValue <> 1 THEN 
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")  
			Set objCmd = nothing
			Exit For
			response.end
		END IF 
	Set objCmd = nothing		
	Next	
 
	IF returnValue = "1" THEN 
		call Alert_closenreload("등록되었습니다.")
	ELSE	
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	response.end 
Case "U" 
	IF InOutType = "" THEN InOutType = 0
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpAccount_Update]("&iOpExpAccountIdx&","&InOutType&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue = "1" THEN
		Call Alert_move ("처리되었습니다.","/admin/expenses/account/?menupos="&menupos) 
	ELSE	
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	response.end 
	
Case "D"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpAccount_Delete]("&iOpExpAccountIdx&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue = "1" THEN
		Call Alert_move ("삭제되었습니다.","/admin/expenses/account/?menupos="&menupos) 
	ELSE	
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	response.end 
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
