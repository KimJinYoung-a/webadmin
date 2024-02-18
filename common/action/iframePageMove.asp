<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2016.11.24 정윤정 생성
' Description : 이미지서버 오류시 아이프레임 리셋처리
'####################################################

	dim refURL
	refURL = request("url")
	refURL = Replace(refURL,"|","&")

	Response.Write "<script type=""text/javascript"">" & vbCrLf &_
			"self.location.href='about:blank;'"  & vbCrLf &_
			"</script>"
	response.End
%>
