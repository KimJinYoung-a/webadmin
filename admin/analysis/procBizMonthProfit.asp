<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ������ update
' History : 2012.08.13 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
Dim sMode, dYYYYMM ,dSDate,dEDate
Dim objCmd, returnValue
sMode		= requestCheckvar(Request("hidM"),2)
dYYYYMM= requestCheckvar(Request("hidYM"),7) 
dSDate = dateserial(year(dYYYYMM),month(dYYYYMM),"1")
dEDate = dateadd("d",-1,dateadd("m",1,dSDate))
SELECT CASE sMode
CASE "R" 
		Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_BizMonthProfit_Insert]('"&dSDate&"', '"&dEDate&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing  
	
	IF returnValue  = 0 THEN
	 Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.")	
	END IF
		Call alert_move ("ó���Ǿ����ϴ�.","bizMonthProfitReport.asp?selY="&year(dYYYYMM)&"&selM="&month(dYYYYMM))	
CASE "B"
		Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_BIZMonthProfit_Bizsection_Insert]('"&dYYYYMM&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing  
	IF returnValue  = 0 THEN
	 Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.")	
	END IF
		Call alert_move ("ó���Ǿ����ϴ�.","bizMonthProfitBiz.asp?selY="&year(dYYYYMM)&"&selM="&month(dYYYYMM))	
CASE ELSE
	Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.")	
END SELECT
<!-- #include virtual="/lib/db/dbclose.asp" -->
%>