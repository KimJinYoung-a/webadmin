<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2011.03.29 허진원 생성
' Description : 이미지서버 등의 domain이 다른경우의 iframe 부모창 이동 처리 페이지
'####################################################

	dim refURL
	refURL = request("url")
	refURL = Replace(refURL,"|","&")

	Response.Write "<script type=""text/javascript"">" & vbCrLf &_
			"parent.location.replace(""" & refURL & """);" & vbCrLf &_
			"</script>"
	response.End
%>
