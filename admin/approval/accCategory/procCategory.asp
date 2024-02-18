<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<% 
Dim objCmd, returnValue, sMode
Dim icatedepth,scatename,iaccOrder,ipcateidx,icategoryidx,blnUsing ,chk
Dim iCcateidx, iCpcateidx,iCateDIdx
Dim menupos
dim isS10, isSP, sdivide, sdividedesc,sacccd
sMode		= requestCheckvar(Request("hidM"),1)
icategoryidx= requestCheckvar(Request("icidx"),10)
icatedepth	= requestCheckvar(Request("icd"),10)
ipcateidx	= requestCheckvar(Request("selCL"),10)
scatename	= requestCheckvar(Request("scn"),64)
iaccOrder	= requestCheckvar(Request("iAO"),5)
blnUsing	= requestCheckvar(Request("blnU"),1)
menupos		= requestCheckvar(Request("menupos"),10)
iCcateidx = requestCheckvar(Request("iccidx"),10)
iCpcateidx= requestCheckvar(Request("selCCL"),10)
iCateDIdx = requestCheckvar(Request("hidCDIdx"),10)
isS10		= requestCheckvar(Request("isS10"),1)
isSP		= requestCheckvar(Request("isSP"),1)
sdivide		= requestCheckvar(Request("sdivide"),128)
sdividedesc		= requestCheckvar(Request("sdividedesc"),500)
sacccd		 = requestCheckvar(Request("hidacc"),15)

chk = request("chk")  
SELECT CASE sMode
Case "I"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_ACC_CD_Category_Insert]( '"&scatename&"',"&icatedepth&" ,"&ipcateidx&",  "&iaccOrder&"  )}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue<>  "1" THEN 
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF	
	
		call Alert_closenmove("등록되었습니다.","categorylist.asp?selCL="&ipcateidx&"&menupos="&menupos)  
	response.end 
Case "U"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_ACC_CD_Category_Update]("&icategoryidx&",'"&scatename&"', "&icatedepth&" ,"&ipcateidx&",  "&iaccOrder&" ,"&blnUsing&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue <> "1" THEN
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
		END IF
		Call Alert_closenmove ("수정되었습니다.","categorylist.asp?selCL="&ipcateidx&"&menupos="&menupos)  
	response.end 
CASE "S"
Dim i,AssignedRow,arrIdx
	arrIdx = split(chk,",")
	AssignedRow = 0 

For i = 0  To UBound(arrIdx)  
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_ACC_CD_CategoryDetail_Insert]("&iCcateidx&",'"&trim(arrIdx(i))&"')}"							 					 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing 

	IF returnValue =  1  THEN   
		AssignedRow=AssignedRow+1
	END IF 
Next 
 
		Call Alert_move (AssignedRow&"건 등록되었습니다.","index.asp?selCCL="&iCpcateidx&"&selCC="&iCcateidx&"&menupos="&menupos)  
	response.end 	
CASE "D"
		Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_ACC_CD_CategoryDetail_Delete]("&iCateDIdx&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue <> "1" THEN
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
		END IF
		Call Alert_move ("삭제되었습니다.","index.asp?selCL="&ipcateidx&"&selC="&icategoryidx&"&menupos="&menupos)  
	response.end 
CASE "C" '안분기준 등록
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_ACC_CD_CategoryDetail_DivideInsert]('"&sacccd&"','"&isS10&"','"&isSP&"','"&sdivide&"','"&sdividedesc&"')}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue <> "1" THEN
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
		END IF
		Call Alert_closenmove ("안분기준이 등록되었습니다.","index.asp?selCL="&ipcateidx&"&selC="&icategoryidx&"&menupos="&menupos)  
	response.end 

CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
