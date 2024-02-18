<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim sqlStr, sellsite, yyyymm, ret, retval
dim mindate

sellsite = requestCheckVar(request("sellsite"), 32)
yyyymm = requestCheckVar(request("jsdate"), 10)

If sellsite = "" OR yyyymm = "" Then
	response.write	"<script language='javascript'>" &_
		"	alert('제휴몰 또는 매출월이 선택되지 않았습니다. 확인 후 재시도 해주세요'); " &_
			"	location.href='about:blank;' " &_
		"</script>"
	response.end
Else
	rw "시작!"
	response.flush
	response.clear
end If
dim objCmd, returnValue
Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping1]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 추가
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

if returnValue = 0 Then
	call Alert_return("매칭1단계 오류- 다시 시도해주세요")
	response.end
Else
	rw "1"
	response.flush
	response.clear
end if

Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping1_D]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 추가
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

if returnValue = 0 Then
	call Alert_return("매칭1-D단계 오류- 다시 시도해주세요")
	response.end
Else
	rw "1-D"
	response.flush
	response.clear
end if

Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping2]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 추가
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

 if returnValue = 0 Then
 		call Alert_return("매칭2단계 오류- 다시 시도해주세요")
 	response.end
Else
	rw "2"
	response.flush
	response.clear
end if

Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping3]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 추가
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

 if returnValue = 0 Then
 		call Alert_return("매칭3단계 오류- 다시 시도해주세요")
 	response.end
Else
	rw "3"
	response.flush
	response.clear
end if

Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping4]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 추가
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

 if returnValue = 0 Then
 		call Alert_return("매칭4단계 오류- 다시 시도해주세요")
 	response.end
Else
	rw "4"
	response.flush
	response.clear
end if


Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderItem_mapping5]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 추가
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

 if returnValue = 0 Then
 		call Alert_return("매칭5단계 오류- 다시 시도해주세요")
 	response.end
Else
	rw "5"
	response.flush
	response.clear
end if

Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbSTSget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_statistics].[dbo].[usp_Ten_extJungsan_orderMaster_mapping]('"&sellsite&"', '"&yyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.CommandTimeout = 300 ''2019/01/16 추가
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

 if returnValue = 0 Then
 		call Alert_return("매칭M단계 오류- 다시 시도해주세요")
 	response.end
Else
	rw "M끝..완료!"
	response.flush
	response.clear
end if

%>
<script type="text/javascript">
alert("작성되었습니다.");
//location.href = "<%= manageUrl %>/common/popReloadOpener.asp";
opener.location.reload();
self.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->

