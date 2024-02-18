<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<% 
Dim objCmd, returnValue, sMode
Dim  eappDepth,eappPartIdx,eappPartName, blnusing , istep1partidx, istep2partidx,partSort
Dim menupos

sMode			= requestCheckvar(Request("hidM"),1)
eappPartIdx		= requestCheckvar(Request("iepidx"),10)
istep1partidx	= requestCheckvar(Request("selp1"),10)
istep2partidx	= requestCheckvar(Request("selp2"),10) 
IF istep1partidx = "" THEN istep1partidx = 0
IF istep2partidx = "" THEN istep2partidx = 0
	
IF istep1partidx = 0 THEN
	eappDepth = 1 
ELSEIF istep2partidx = 0 THEN
	eappDepth = 2 
ELSE
	eappDepth = 3 
END IF

eappPartName	= requestCheckvar(Request("sPN"),32) 
partSort		= requestCheckvar(Request("iPS"),10) 
blnusing		= requestCheckvar(Request("rdoU"),1) 
menupos			= requestCheckvar(Request("menupos"),10)
 
SELECT CASE sMode
Case "I"
	IF partSort = "" THEN partSort = 0
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPart_Insert]( "&eappDepth&","&istep1partidx&","&istep2partidx&",'"&eappPartName&"',"&partSort&")}"							 
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
 IF partSort = "" THEN partSort = 0
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPart_Update]("&eappPartIdx&","&eappDepth&","&istep1partidx&","&istep2partidx&",'"&eappPartName&"',"&partSort&","&blnUsing&")}"							 
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
