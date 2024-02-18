<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	dim sqlStr, getdate
    getdate = request("getdate")

    sqlStr = "EXEC [db_sitemaster].[dbo].[usp_SCM_MDPick_ItemCopy_Add] '" & getdate & "'" & VbCrlf
    dbget.Execute sqlStr

    if Err.Number <> 0 then
        response.Write "Error"
    else
        response.Write "OK"
    end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->