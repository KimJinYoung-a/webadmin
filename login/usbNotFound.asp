<%@ language=vbscript %>
<% option explicit %>
<%
    'session("ssBctId") = ""
    'session("ssBctCompanyName") = ""
    'session("ssBctTel") = ""
    'session("ssBctFax") = ""
    'session("ssBctUrl") = ""
    'session("ssBctEmail") = ""
    'session("ssGroupid") = ""
	session.abandon

	Response.Write	"<html><body><script language=javascript>"
	Response.Write	"alert('텐바이텐 웹어드민 USB키가 없습니다.\n\n※로그인하시려면 USB키를 꼽아주세요.');"
	Response.Write	"top.location = '/';"
	Response.Write	"</script></body></html>"
%>
