<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 전자결재 공통코드 등록 
' History : 2011.03.09 정윤정  생성
'			2022.07.11 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<% 
Dim objCmd, returnValue, sMode
Dim icomm_cd, iparentkey, scomm_name, scomm_desc, idispnum, ierpcode,blnactiveYn
Dim menupos

sMode		= requestCheckvar(Request("hidM"),1)
icomm_cd	= requestCheckvar(Request("icc"),10)
iparentkey	= requestCheckvar(Request("selpk"),10)
scomm_name	= requestCheckvar(Request("scn"),10)
scomm_desc	= requestCheckvar(Request("scd"),64)
ierpcode	= requestCheckvar(Request("iEC"),10)
idispnum	= requestCheckvar(Request("idn"),5)
blnactiveYn	= requestCheckvar(Request("blnayn"),1)
menupos		= requestCheckvar(Request("menupos"),10)

IF iparentkey = "" THEN iparentkey = 0
IF ierpcode = "" THEN ierpcode = 0
IF idispnum = "" THEN idispnum = 0
IF blnactiveYn = "" THEN blnactiveYn = 1
	
SELECT CASE sMode
Case "I"
	if scomm_name <> "" and not(isnull(scomm_name)) then
		scomm_name = ReplaceBracket(scomm_name)
	end If
	if scomm_desc <> "" and not(isnull(scomm_desc)) then
		scomm_desc = ReplaceBracket(scomm_desc)
	end If

	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppCommCD_Insert]( "&iparentkey&" ,'"&scomm_name&"', '"&scomm_desc&"' ,"&ierpcode&","&idispnum&")}"							 
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
	if scomm_name <> "" and not(isnull(scomm_name)) then
		scomm_name = ReplaceBracket(scomm_name)
	end If
	if scomm_desc <> "" and not(isnull(scomm_desc)) then
		scomm_desc = ReplaceBracket(scomm_desc)
	end If

	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppCommCD_Update]("&icomm_cd&","&iparentkey&" ,'"&scomm_name&"', '"&scomm_desc&"' ,"&ierpcode&","&idispnum&","&blnactiveYn&")}"							 
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
