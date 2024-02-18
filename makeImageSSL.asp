<%@ language="vbscript" %><% option explicit %><%

dim imageFileName : imageFileName = request("f")

imageFileName = "http://webimage.10x10.co.kr" + imageFileName

response.end
'' 사용안함

Response.ContentType = "image/jpeg"

dim xmlhttp, status
set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.3.0")
xmlhttp.open "GET", imageFileName, false
xmlhttp.send ""
Response.BinaryWrite xmlhttp.ResponseBody
set xmlhttp = nothing

%>
