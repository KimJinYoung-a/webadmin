<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incTenRedisSession.asp"-->
<%
    dim pssBctDiv : pssBctDiv = session("ssBctDiv")
    session.abandon
    
    Response.Cookies("partner").domain = "10x10.co.kr"
    Response.Cookies("partner") = ""
    Response.Cookies("partner").Expires = Date - 1

    Response.Cookies("ThreePL").Domain	= "10x10.co.kr"
    Response.Cookies("ThreePL") = ""
    Response.Cookies("ThreePL").Expires = Date - 1
    
    Response.Cookies("wapi").Domain	= "10x10.co.kr"
    Response.Cookies("wapi") = ""
    Response.Cookies("wapi").Expires = Date - 1
    
    Dim cookieDomain
    If Application("Svr_Info") = "Dev" And InStr(Request.ServerVariables("HTTP_REFERER"), "localhost") > 0 Then
        cookieDomain = "localhost"
    Else
        cookieDomain = "10x10.co.kr"
    End If
    Response.Cookies("pinfo").Domain	= cookieDomain
    Response.Cookies("pinfo") = ""
    Response.Cookies("pinfo").Expires = Date - 1
    
    ''2018/12/28
    Call fn_RDS_SSN_Expire() 

	Response.Write	"<html><body><script type=""text/javascript"">"
    Response.Write	"alert('로그아웃되었습니다.');"
    Response.Write	"top.location = '/index.asp';"
	Response.Write	"</script></body></html>"
%>
