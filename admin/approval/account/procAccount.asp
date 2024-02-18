<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<% 
Dim objCmd, returnValue, sMode
Dim iaccountidx , iaccountkind, iedmsidx, saccountname, blnusing
 

sMode		= requestCheckvar(Request("hidM"),1)
iaccountidx	= requestCheckvar(Request("iaidx"),10)
iaccountkind	= requestCheckvar(Request("selAK"),10)
iedmsidx	= requestCheckvar(Request("selC3"),10)
saccountname	= requestCheckvar(Request("sAN"),64) 
blnusing	= requestCheckvar(Request("blnU"),1)
 
IF blnusing = "" THEN blnusing = 1
	
SELECT CASE sMode
Case "I"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAccount_Insert]( "&iaccountkind&" ,"&iedmsidx&", '"&saccountname&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue = "1" THEN 
		call Alert_closenreload("등록되었습니다.")
	ELSE	
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	response.end 
Case "U"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAccount_Update]("&iaccountidx&","&iaccountkind&" ,"&iedmsidx&", '"&saccountname&"' ,"&blnusing&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue = "1" THEN
		Call Alert_closenreload ("수정되었습니다.") 
	ELSE	
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	response.end 
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
