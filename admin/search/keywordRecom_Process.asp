<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim i, sqlStr

dim mode : mode=requestCheckvar(request("mode"),32)
dim keyword : keyword=requestCheckvar(request("keyword"),50)
dim group_no : group_no=requestCheckvar(request("group_no"),10)
dim itemid : itemid=requestCheckvar(request("itemid"),10)

dim reguserid : reguserid = session("ssBctId")
dim returnValue
dim retErrStr, succMsg

Dim objCmd

if (mode="addmaster") then
    succMsg ="저장되었습니다."
    Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdStoredProc
		.CommandText = "[db_temp].[dbo].[usp_TEN_ksearch_keyword_recom_master_insert]"
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@keyword", adVarchar, adParamInput,100 , keyword)
        .Parameters.Append .CreateParameter("@retErrStr", adVarchar,adParamOutput,100, "")
        
		.Execute , , adExecuteNoRecords
		End With
        returnValue = objCmd.Parameters("RETURN_VALUE").Value
        retErrStr  = objCmd.Parameters("@retErrStr").Value

	Set objCmd = Nothing
	
elseif (mode="delmaster") then
    succMsg ="삭제되었습니다."
    Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdStoredProc
		.CommandText = "[db_temp].[dbo].[usp_TEN_ksearch_keyword_recom_master_delete]"
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@keyword", adVarchar, adParamInput,100 , keyword)
        .Parameters.Append .CreateParameter("@retErrStr", adVarchar,adParamOutput,100, "")
        
		.Execute , , adExecuteNoRecords
		End With
        returnValue = objCmd.Parameters("RETURN_VALUE").Value
        retErrStr  = objCmd.Parameters("@retErrStr").Value

	Set objCmd = Nothing
elseif (mode="additem") then
    succMsg ="추가되었습니다."
    Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdStoredProc
		.CommandText = "[db_temp].[dbo].[usp_TEN_ksearch_keyword_recom_detail_insert]"
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@group_no", adInteger, adParamInput, , group_no)
        .Parameters.Append .CreateParameter("@itemid", adInteger, adParamInput, , itemid)
        .Parameters.Append .CreateParameter("@retErrStr", adVarchar,adParamOutput,100, "")
        
		.Execute , , adExecuteNoRecords
		End With
        returnValue = objCmd.Parameters("RETURN_VALUE").Value
        retErrStr  = objCmd.Parameters("@retErrStr").Value

	Set objCmd = Nothing
elseif (mode="delitem") then
    succMsg ="삭제되었습니다."
    Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdStoredProc
		.CommandText = "[db_temp].[dbo].[usp_TEN_ksearch_keyword_recom_detail_delete]"
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@group_no", adInteger, adParamInput, , group_no)
        .Parameters.Append .CreateParameter("@itemid", adInteger, adParamInput, , itemid)
        .Parameters.Append .CreateParameter("@retErrStr", adVarchar,adParamOutput,100, "")
        
		.Execute , , adExecuteNoRecords
		End With
        returnValue = objCmd.Parameters("RETURN_VALUE").Value
        retErrStr  = objCmd.Parameters("@retErrStr").Value

	Set objCmd = Nothing
    

else
    rw "ERR:"&mode

end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")

if (mode="addmaster") or (mode="delmaster") or (mode="additem") or (mode="delitem") then
    if (returnValue<0) then
        response.write "<script language='javascript'>alert('"&retErrStr&"');</script>"

        response.write retErrStr&"<br>"
        response.write "<input type='button' value='BACK' onClick='location.href="""&refer&"""'>"
    else
        response.write "<script language='javascript'>alert('"&succMsg&"');</script>"
        response.write "<script language='javascript'>location.replace('" + refer + "');</script>"
    end if


else
    response.write "<script language='javascript'>alert('수정되었습니다.');</script>"
    response.write "<script language='javascript'>location.replace('" + refer + "');</script>"
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
