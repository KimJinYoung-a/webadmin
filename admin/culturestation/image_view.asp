<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이미지 미리보기 
' History : 2008.04.22 한용민 생성
'###########################################################
%>

<%
dim image
	image = request("image")
%>
<img src="<%=image%>">