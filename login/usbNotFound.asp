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
	Response.Write	"alert('�ٹ����� ������ USBŰ�� �����ϴ�.\n\n�طα����Ͻ÷��� USBŰ�� �ž��ּ���.');"
	Response.Write	"top.location = '/';"
	Response.Write	"</script></body></html>"
%>
