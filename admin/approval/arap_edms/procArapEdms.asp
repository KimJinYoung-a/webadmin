<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<% 
Dim objCmd, returnValue, sMode
Dim dARAPCode,iedmsidx ,sUsing
Dim iARAPLinkedmsIdx

sMode		= requestCheckvar(Request("hidM"),1)
dARAPCode	= requestCheckvar(Request("dAC"),13)
iedmsidx	= requestCheckvar(Request("ieidx"),10)  
iARAPLinkedmsIdx = requestCheckvar(Request("idx"),10)  
sUsing	= requestCheckvar(Request("rdoU"),1)  
SELECT CASE sMode
Case "I"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_ARAPLinkedms_Insert]( '"&dARAPCode&"' ,"&iedmsidx&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue = "2" THEN 
		Call Alert_return ("기존데이터가 존재합니다.확인 후 다시 등록해주세요") 
	ELSEIF returnValue = "1" THEN 
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
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_ARAPLinkedms_Update]('"&iARAPLinkedmsIdx&"',"&iedmsidx&",'"&sUsing&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords	
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	

	IF returnValue = "1" THEN 
		call Alert_closenreload("수정되었습니다.")
	ELSE	
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	response.end	 
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
