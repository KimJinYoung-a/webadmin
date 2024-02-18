<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim catecode 
dim metakeywords 
dim searchkeywords 

dim i,cnt : cnt =  request("cksel").Count
dim idx, sqlStr

dim mode : mode=requestCheckvar(request("mode"),32)
dim addkeyword : addkeyword=requestCheckvar(request("addkeyword"),32)
dim addcatecode : addcatecode=requestCheckvar(request("addcatecode"),18)
dim edtcateusing : edtcateusing=requestCheckvar(request("edtcateusing"),1)

dim addmakerid : addmakerid=requestCheckvar(request("addmakerid"),32)
dim edtbrandusing : edtbrandusing=requestCheckvar(request("edtbrandusing"),1)

dim reguserid : reguserid = session("ssBctId")
dim assignedRow : assignedRow = 0

Dim objCmd

if (mode="addcateboostkey") then
    
    Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_const].[dbo].[usp_Ten_Const_display_cate_BoostKeyWord_Add]('"&addkeyword&"','"&addcatecode&"','"&reguserid&"')}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    assignedRow = objCmd(0).Value
	Set objCmd = Nothing
	
	if assignedRow<1 then assignedRow=0
elseif (mode="cateboostkeychg") then
    
    Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_const].[dbo].[usp_Ten_Const_display_cate_BoostKeyWord_Edit]('"&addkeyword&"','"&addcatecode&"','"&edtcateusing&"')}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    assignedRow = objCmd(0).Value
	Set objCmd = Nothing
	
    if assignedRow<1 then assignedRow=0
elseif (mode="addbrandboostkey") then
    Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_const].[dbo].[usp_Ten_Const_Brand_BoostKeyWord_Add]('"&addkeyword&"','"&addmakerid&"','"&reguserid&"')}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    assignedRow = objCmd(0).Value
	Set objCmd = Nothing
	
	if assignedRow<1 then assignedRow=0
elseif (mode="brandboostkeychg") then
    
    Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_const].[dbo].[usp_Ten_Const_Brand_BoostKeyWord_Edit]('"&addkeyword&"','"&addmakerid&"','"&edtbrandusing&"')}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    assignedRow = objCmd(0).Value
	Set objCmd = Nothing
	
    if assignedRow<1 then assignedRow=0
else
    for i=1 to cnt
        idx = Trim(request("cksel")(i))
        catecode = Trim(request("catecode")(idx+1))
        metakeywords = Trim(request("metakeywords")(idx+1))
        searchkeywords = Trim(request("searchkeywords")(idx+1))
        
        ''response.write catecode&"|"&metakeywords&"|"&searchkeywords&"<br>"
        
        sqlStr = "EXEC [db_const].[dbo].[usp_Ten_Const_display_cate_Edit_keyWords] '"&catecode&"','"&metakeywords&"','"&searchkeywords&"'"
        dbget.Execute sqlStr
    next
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")

if (mode="addcateboostkey") or (mode="cateboostkeychg") or (mode="addbrandboostkey") or (mode="brandboostkeychg") then
    response.write "<script language='javascript'>alert('"&assignedRow&"건 반영되었습니다..');</script>"
    ''response.write "<script language='javascript'>alert('수정되었습니다.');</script>"
    response.write "<script language='javascript'>location.replace('" + refer + "');</script>"
else
    response.write "<script language='javascript'>alert('수정되었습니다.');</script>"
    response.write "<script language='javascript'>location.replace('" + refer + "');</script>"
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
