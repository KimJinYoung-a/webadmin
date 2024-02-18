<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim strSql
Dim mode, itemid, makerid, excepttp, iAasignMxDt
mode	    = requestCheckVar(request("mode"), 32)
itemid	    = requestCheckVar(request("itemid"), 10)
makerid	    = requestCheckVar(request("makerid"), 32)
excepttp	= requestCheckVar(request("excepttp"), 1)
iAasignMxDt = requestCheckVar(request("iAasignMxDt"), 10)

Dim cmd, intResult
Dim iErrMsg

If (excepttp="B") and ((mode = "D") or (mode = "R") or (mode = "I")) Then
	strSql = "db_temp.[dbo].[usp_Ten_NV_ItemCPn_Except_Brand]"

    set cmd = server.CreateObject("ADODB.Command")

    cmd.ActiveConnection = dbget
    cmd.CommandText = strSql
    cmd.CommandType = adCmdStoredProc

    cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append cmd.CreateParameter("@makerid", adVarchar, adParamInput, 32, makerid)
    cmd.Parameters.Append cmd.CreateParameter("@act", adVarchar, adParamInput, 32, mode)
    cmd.Parameters.Append cmd.CreateParameter("@reguser", adVarchar, adParamInput, 32, session("ssBctID"))
    cmd.Parameters.Append cmd.CreateParameter("@AasignMxDt", adVarchar, adParamInput, 10, iAasignMxDt)
    
    cmd.Execute 
    intResult = cmd.Parameters("returnValue").Value
    set cmd = Nothing

    if (intResult=-2) Then
        iErrMsg = "해당 브랜드가 없습니다."
    elseif (intResult=-3) Then
        iErrMsg = "이미 등록된 브랜드입니다."
    end if


    if (intResult<1) Then
        response.write	"<script language='javascript'>" &_
		    			"	alert('데이터 처리에 문제가 발생하였습니다.("&iErrMsg&")');  " &_
		    			"</script>"
    else
       
        response.write	"<script language='javascript'>" &_
                        "	alert('수정 되었습니다.'); opener.location.reload(); window.close();" &_
                        "</script>"
    end if
ElseIf (excepttp="I") and ((mode = "D") or (mode = "R") or (mode = "I")) Then
	strSql = "db_temp.[dbo].[usp_Ten_NV_ItemCPn_Except_Item]"

    set cmd = server.CreateObject("ADODB.Command")

    cmd.ActiveConnection = dbget
    cmd.CommandText = strSql
    cmd.CommandType = adCmdStoredProc

    cmd.Parameters.Append cmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
    cmd.Parameters.Append cmd.CreateParameter("@itemid", adInteger, adParamInput, 32, itemid)
    cmd.Parameters.Append cmd.CreateParameter("@act", adVarchar, adParamInput, 32, mode)
    cmd.Parameters.Append cmd.CreateParameter("@reguser", adVarchar, adParamInput, 32, session("ssBctID"))
    cmd.Parameters.Append cmd.CreateParameter("@AasignMxDt", adVarchar, adParamInput, 10, iAasignMxDt)
    
    cmd.Execute 
    intResult = cmd.Parameters("returnValue").Value
    set cmd = Nothing

    if (intResult=-2) Then
        iErrMsg = "해당 상품이 없습니다."
    elseif (intResult=-3) Then
        iErrMsg = "이미 등록된 상품입니다."
    end if

    if (intResult<1) Then
        response.write	"<script language='javascript'>" &_
		    			"	alert('데이터 처리에 문제가 발생하였습니다.("&iErrMsg&")'); " &_
		    			"</script>"
    else
        response.write	"<script language='javascript'>" &_
                        "	alert('수정 되었습니다.'); opener.location.reload(); window.close();" &_
                        "</script>"
    end if
Else
	response.write	"<script language='javascript'>" &_
					"	alert('잘못 된 접근입니다. invalid Param');  " &_
					"</script>"
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->