<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
Dim sqlStr
Dim vType, vUserinputdata
Dim useq, userid, userdiv, lastlogin, counter, userlevel
dim oJson
Set oJson = jsObject()


vType			= requestCheckVar(request("ttype"),200)
vUserinputdata	= requestCheckVar(request("tinputdata"),800)


Set oJson("data") = jsArray()
Set oJson("data")(null) = jsObject() 


If vType="" Then
	oJson("data")(null)("result") = "ERR|정상적인 경로로 접근해주세요."
	oJson.flush
	Set oJson = Nothing	
	dbget.close() : Response.End
End IF

If vUserinputdata="" Then
	oJson("data")(null)("result") = "ERR|useq값이나 userid를 입력해주세요."
	oJson.flush
	Set oJson = Nothing	
	dbget.close() : Response.End
End If

If vType="userid" Then
    sqlStr = "select useq*3 as useq, userid, userdiv, CONVERT(VARCHAR(20),lastlogin,120) AS lastlogin, counter, userlevel from db_user.dbo.tbl_logindata with(nolock) where userid = '" & vUserinputdata & "' "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    IF Not rsget.Eof Then
        useq = rsget("useq")
        userid = rsget("userid")
        userdiv = rsget("userdiv")
        lastlogin = rsget("lastlogin")
        counter = rsget("counter")
        userlevel = rsget("userlevel")
    Else
        useq = ""
        userid = ""
        userdiv = ""
        lastlogin = ""
        counter = ""
        userlevel = ""
    End IF
    rsget.close
End If

If vType="useq" Then
    sqlStr = "select useq*3 as useq, userid, userdiv, CONVERT(VARCHAR(20),lastlogin,120) AS lastlogin, counter, userlevel from db_user.dbo.tbl_logindata with(nolock) where useq = " & vUserinputdata/3 & " "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    IF Not rsget.Eof Then
        useq = rsget("useq")
        userid = rsget("userid")
        userdiv = rsget("userdiv")
        lastlogin = rsget("lastlogin")
        counter = rsget("counter")
        userlevel = rsget("userlevel")
    Else
        useq = ""
        userid = ""
        userdiv = ""
        lastlogin = ""
        counter = ""
        userlevel = ""
    End IF
    rsget.close
End If

If userid = "" Then
	oJson("data")(null)("result") = "ERR|사용자 데이터가 없습니다."
	oJson.flush
	Set oJson = Nothing	
	dbget.close() : Response.End
End If

oJson("data")(null)("result") = "OK|entry"
oJson("data")(null)("useq") = useq
oJson("data")(null)("userid") = userid
oJson("data")(null)("userdiv") = userdiv
oJson("data")(null)("lastlogin") = lastlogin
oJson("data")(null)("counter") = counter
oJson("data")(null)("userlevel") = userlevel
oJson.flush
Set oJson = Nothing		
dbget.close() : Response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->