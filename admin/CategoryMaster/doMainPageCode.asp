<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim divCd, divName, divType
dim imgWidth, imgHeight
dim isusing    


divCd   = request.Form("divCd")
divName   = Replace(request.Form("divName"),"'","''")
divType   = request.Form("divType")
imgWidth= request.Form("imgWidth")
imgHeight	= request.Form("imgHeight")
isusing   = request.Form("isusing")


dim sqlStr, ItemExists

if Not(divCd="" or isNull(divCd)) then
    sqlStr = " update [db_sitemaster].[dbo].tbl_category_mainItem_div" + VbCrlf
    sqlStr = sqlStr + " set divName='" + divName + "'" + VbCrlf
    sqlStr = sqlStr + " ,imgWidth='" + imgWidth + "'" + VbCrlf
    sqlStr = sqlStr + " ,imgHeight='" + imgHeight + "'" + VbCrLf
    sqlStr = sqlStr + " ,divType='" + divType + "'" + VbCrlf
    sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " where divCd=" + CStr(divCd) + VbCrlf
    
    dbget.Execute sqlStr
else
    sqlStr = " insert into [db_sitemaster].[dbo].tbl_category_mainItem_div" + VbCrlf
    sqlStr = sqlStr + " (divName,imgWidth,imgHeight,divType,isusing)"+ VbCrlf
    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " '" + divName + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + imgWidth + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + imgHeight + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + divType + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " )" + VbCrlf
    
    dbget.Execute sqlStr
end if


dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->