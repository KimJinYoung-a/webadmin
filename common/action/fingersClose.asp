<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2011.03.29 ������ ����
' Description : �̹������� ���� domain�� �ٸ������ iframe �θ�â �̵� ó�� ������
'####################################################

		response.write "<script>alert('�����Ǿ����ϴ�.');</script>"
		response.write "<script>"
		response.write "function selfClose() {"
		response.write "	if (/MSIE/.test(navigator.userAgent)) { "
		response.write "		if(navigator.appVersion.indexOf('MSIE 8.0')>=0) {"
		response.write "			window.opener='Self';"
		response.write "			window.open('','_parent','');"
		response.write "			window.close();"
		response.write "		} else if(navigator.appVersion.indexOf('MSIE 7.0')>=0) {"
		response.write "			window.open('about:blank','_self').close();"
		response.write "		} else { "
		response.write "			window.opener = self;"
		response.write "			self.close();"
		response.write "		}"
		response.write "	} else {"
		response.write "		self.close();"
		response.write "	}"
		response.write "}"
		response.write "selfClose();"
		response.write "</script>"
%>
