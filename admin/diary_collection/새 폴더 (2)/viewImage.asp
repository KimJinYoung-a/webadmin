<html>
<head>
<title>이미지 원본 보기 </title></head>
<body>
<%
dim imageUrl

imageUrl = request("imageUrl")
%>
<img src="<%= imageUrl %>" border="0">
</body>
<html>
