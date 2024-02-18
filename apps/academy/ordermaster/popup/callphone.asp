<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"

Dim phonenum
phonenum=request("phonenum")
%>
<!doctype html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="Generator" content="EditPlus®">
<meta name="Author" content="">
<meta name="Keywords" content="">
<meta name="Description" content="">
<title>전화걸기</title>
<script type="text/javascript">
<!--
	location.href="tel:<%=phonenum%>";
	setTimeout(function(){
		window.close();
	}, 100);
//-->
</script>
</head>
<body>
</body>
</html>