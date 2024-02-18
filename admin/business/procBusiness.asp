<%@ language=vbscript %>
<% option explicit  %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/expenses/OpExpCls.asp"-->
<%
Dim sMode,busiIdx
Dim userid,busiNo,busiName,busiCEOName,busiAddr,busiType,busiItem,repName,repEmail,repTel,confirmYn,regdate,delYn,guestOrderserial,useType	 
Dim objCmd,returnValue

 sMode		= requestCheckvar(Request("hidM"),1) 
 busiIdx	=requestCheckvar(Request("hidBI"),10) 
 userid		= requestCheckvar(Request("sUID"),32) 
 busiNo		=requestCheckvar(Request("sBN1"), 3) &"-"&requestCheckvar(Request("sBN2"), 2) &"-"&requestCheckvar(Request("sBN3"), 5)
 busiName	=requestCheckvar(Request("sBNa"),60) 
 busiCEOName=requestCheckvar(Request("sCeo"),32) 
 busiAddr	=requestCheckvar(Request("sBA"),125) 
 busiType	=requestCheckvar(Request("sBT"),25) 
 busiItem	=requestCheckvar(Request("sBI"),25) 
 repName	=requestCheckvar(Request("sRN"),32) 
 repEmail	=requestCheckvar(Request("sRE"),125) 
 repTel		=requestCheckvar(Request("sRT"),18)  
 delYn		=requestCheckvar(Request("rdoD"),1) 
 guestOrderserial=requestCheckvar(Request("sGO"),11) 
 useType	=requestCheckvar(Request("sUT"),1) 	 
SELECT CASE sMode
Case "I"   
 	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText   
			.CommandText = "{?= call db_order.[dbo].[sp_Ten_busiInfo_Insert]('"&userid&"','"&busiNo&"','"&busiName&"','"&busiCEOName&"','"&busiAddr&"','"&busiType&"','"&busiItem&"', '"&repName&"' ,'"&repEmail&"','"&repTel&"','"&guestOrderserial&"','"&useType&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	 
	IF returnValue <>  1  THEN   
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	
response.redirect "popBusiness.asp?sBNo="&busiNo&"&sBNa="&busiName 
	response.end 
Case "U"  
	Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText   
			.CommandText = "{?= call db_order.[dbo].[sp_Ten_busiInfo_Update]("&busiIdx&",'"&userid&"','"&busiNo&"','"&busiName&"','"&busiCEOName&"','"&busiAddr&"','"&busiType&"','"&busiItem&"', '"&repName&"' ,'"&repEmail&"','"&repTel&"','"&delYn&"','"&guestOrderserial&"','"&useType&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
	Set objCmd = nothing	 
	IF returnValue <>  1  THEN   
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	
response.redirect "popBusiness.asp?sBNo="&busiNo&"&sBNa="&busiName 
	response.end 
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT	
%>