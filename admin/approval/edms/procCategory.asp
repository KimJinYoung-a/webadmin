<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<% 
Dim objCmd, returnValue, sMode
Dim icatedepth,scatename,scatecode,ipcateidx,icategoryidx,blnUsing 
Dim menupos

sMode		= requestCheckvar(Request("hidM"),1)
icategoryidx= requestCheckvar(Request("icidx"),10)
icatedepth	= requestCheckvar(Request("icd"),10)
ipcateidx	= requestCheckvar(Request("selCL"),10)
scatename	= requestCheckvar(Request("scn"),64)
scatecode	= requestCheckvar(Request("scc"),5)
blnUsing	= requestCheckvar(Request("blnU"),1)
menupos		= requestCheckvar(Request("menupos"),10)

if (checkNotValidHTML(scatename) = true) Then
	response.write "<script>alert('카테고리명에는 HTML을 사용하실 수 없습니다.');history.back();</script>"
	dbget.Close
	response.End
End If

SELECT CASE sMode
Case "I"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_edms_category_insert]( "&icatedepth&" ,'"&scatename&"', '"&scatecode&"' ,"&ipcateidx&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue = "1" THEN 
		call Alert_closenmove("등록되었습니다.","categorylist.asp?selCL="&ipcateidx&"&menupos="&menupos)
	ELSEIF 	returnValue = "2" THEN 
			Call Alert_move ("입력하신 카테고리코드값은 기존에 사용중입니다.다시 입력해주세요","popcategorydata.asp?icidx="&icategoryidx&"&menupos="&menupos)	
	ELSE	
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	response.end 
Case "U"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")  					
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_edms_category_update]("&icategoryidx&", "&icatedepth&" ,'"&scatename&"', '"&scatecode&"' ,"&ipcateidx&","&blnUsing&")}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	
	Set objCmd = nothing
	
	IF returnValue = "1" THEN
		Call Alert_closenmove ("수정되었습니다.","categorylist.asp?selCL="&ipcateidx&"&menupos="&menupos) 
	ELSEIF 	returnValue = "2" THEN 
		Call Alert_move ("입력하신 카테고리코드값은 기존에 사용중입니다.다시 입력해주세요","popcategorydata.asp?icidx="&icategoryidx&"&menupos="&menupos)		
	ELSE	
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1") 
	END IF
	response.end 
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
