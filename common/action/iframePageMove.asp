<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2016.11.24 ������ ����
' Description : �̹������� ������ ���������� ����ó��
'####################################################

	dim refURL
	refURL = request("url")
	refURL = Replace(refURL,"|","&")

	Response.Write "<script type=""text/javascript"">" & vbCrLf &_
			"self.location.href='about:blank;'"  & vbCrLf &_
			"</script>"
	response.End
%>
