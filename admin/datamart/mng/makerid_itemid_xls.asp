<%@ language=vbscript %>
<%
'Response.Buffer = true
Response.AddHeader "Content-Disposition","attachment;filename=텐바이텐 신규입점/상품등록 현황.xls"
Response.ContentType = "application/vnd.ms-excel"
Response.CacheControl = "public"
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
br { mso-data-placement:same-cell; }
</style>
</head>
<body>
<%=fnCSS(request("excel_val"))%>
</body>
</html>
<%
'### 엑셀 저장엔 class 가 안먹어서 일일이 style로 바꿔줌.
Function fnCSS(body)
	body = Replace(body, "<table class=""tbType1 listTb"">", "<table border=1>")
	body = Replace(body, "<div>", "")
	body = Replace(body, "</div>", "")
	body = Replace(body, "class=", "style=")
	body = Replace(body, "ct", "text-align:center !important;")
	body = Replace(body, "fontstrong", "font-weight:bold;")
	body = Replace(body, "fontred", "color:#FF0000 !important;")
	body = Replace(body, "fontblue", "color:#0000FF !important;")
	body = Replace(body, "bgbluett", "background-color:#FF5F5F;")
	body = Replace(body, "bgredtt", "background-color:#39A5FD;")
	body = Replace(body, "bggraytt", "background-color:#F3F3F3;")
	body = Replace(body, "bgitemtt", "background-color:#FaFaFa;")
	body = Replace(body, "bgred", "background-color:#FFD6D6;")
	body = Replace(body, "bgblue", "background-color:#BFE9FF;")
	body = Replace(body, "bgcolor2", "")
	body = Replace(body, "bgcolor3", "")
	body = Replace(body, "bgcolor4", "")
	body = Replace(body, "bgcolor5", "")
	body = Replace(body, "bgcolor6", "")
	body = Replace(body, "bgcolor7", "")
	body = Replace(body, "bgcolor8", "")
	body = Replace(body, "bgcolor9", "")
	body = Replace(body, "bgcolor10", "")
	body = Replace(body, "bgcolor11", "")
	body = Replace(body, "bgcolor12", "")
	body = Replace(body, "bgcolor13", "")
	body = Replace(body, "bgcolor14", "")
	body = Replace(body, "bgcolor15", "")
	body = Replace(body, "bgcolor16", "")
	body = Replace(body, "bgcolor17", "")
	body = Replace(body, "cGy1", "color:#999 !important;")
	body = Replace(body, "fs11", "font-size:11px;")
	fnCSS = body
End Function
%>