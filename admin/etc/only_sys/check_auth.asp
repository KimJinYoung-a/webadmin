<%
	If not(C_SYSTEM_Part) Then
		Response.Write "<script>alert('시스템팀만의 접근 메뉴입니다.');location.href='http://scm.10x10.co.kr/admin/index.asp';</script>"
		Response.End
	End IF
%>
<title>유지보수업무</title>