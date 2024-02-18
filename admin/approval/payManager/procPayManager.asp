<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<% 
Dim objCmd, returnValue, sMode
Dim ipaymanageridx, ipaymanagertype, suserid, blnusing, blnDef 
Dim menupos

sMode			= requestCheckvar(Request("hidM"),1)
ipaymanageridx	= requestCheckvar(Request("ipm"),10)
ipaymanagertype	= requestCheckvar(Request("selPMT"),4)
suserid			= requestCheckvar(Request("hidAI"),32) 
blnusing			= requestCheckvar(Request("rdoU"),1) 
menupos		= requestCheckvar(Request("menupos"),10)
blnDef 		= requestCheckvar(Request("chkD"),1)
	
IF  blnDef = "" THEN blnDef = 0
		
SELECT CASE sMode
Case "I"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPayManager_insert]( '"&suserid&" ',"&ipaymanagertype&","&blnDef&")}"							 
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
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPayManager_Update]("&ipaymanageridx&",'"&suserid&"' ,"&ipaymanagertype&","&blnusing&","&blnDef&")}"							 
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
