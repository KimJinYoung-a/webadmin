<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2011.03.29 ������ ����
' Description : �̹������� ���� domain�� �ٸ������ iframe �θ�â �̵� ó�� ������
'####################################################

	dim refURL
	refURL = request("url")
	refURL = Replace(refURL,"|","&")

	Response.Write "<script type=""text/javascript"">" & vbCrLf &_
			"parent.location.replace(""" & refURL & """);" & vbCrLf &_
			"</script>"
	response.End
%>
