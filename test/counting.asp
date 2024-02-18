<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/classes/appmanage/JSON_2.0.4.asp"-->

<%

'//헤더 출력
Response.ContentType = "text/json"

dim sToken, oJson, clientIp

'// json객체 선언
Set oJson = jsObject()

oJson("cc") = request.cookies

oJson.flush

Set oJson = Nothing
%>