<%@ language=vbscript %>
<% option explicit %>
<%
	Response.Write "<script type='text/javascript'>" & vbCrLf &_
			"alert('����Ǿ����ϴ�.');"& vbCrLf &_
			"opener.location.reload();" & vbCrLf &_
			"window.close();"& vbCrLf &_
			"</script>"
	response.End
%>