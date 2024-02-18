<%@ language=vbscript %>
<% option explicit %>
<%
	Response.Write "<script type='text/javascript'>" & vbCrLf &_
			"alert('저장되었습니다.');"& vbCrLf &_
			"opener.location.reload();" & vbCrLf &_
			"window.close();"& vbCrLf &_
			"</script>"
	response.End
%>