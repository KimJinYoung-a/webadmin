<%@ language=vbscript %>
<% option explicit  %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache" 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
Dim objCmd, returnValue
Dim iType,isValue,BIZSECTION_CD
Dim sMode
iType =  requestCheckvar(Request("iT"),1)
isValue=  requestCheckvar(Request("blnV"),10)
BIZSECTION_CD=  requestCheckvar(Request("sBCD"),10)
sMode =  requestCheckvar(Request("sM"),1)
IF sMode = "U" THEN
	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			'.CommandText = "{?= call db_partner.[dbo].[sp_TMS_get_BA_BIZSECTION]}"							 
			.CommandText = "{?= call db_partner.[dbo].[sp_TMS_get_BA_BIZSECTION_sERP]}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	 	
	IF returnValue <>  1  THEN   
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	response.redirect "/admin/linkedERP/biz/" 
	response.end 	
ELSE	 
 	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_TMS_BA_BIZSECTION_UPDATE]('"&iType&"',"&isValue&",'"&BIZSECTION_CD&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	 	
	IF returnValue <>  1  THEN   
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	response.redirect "/admin/linkedERP/biz/" 
	response.end 	
END IF	
 
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->