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
dim gnbcode 
dim gnbname
dim isusing    

gnbcode   = request.Form("gcode")
gnbname   = html2Db(request.Form("gnbname"))
isusing   = request.Form("isusing")

dim sqlStr, ItemExists

sqlStr = "select top 1 * from db_sitemaster.[dbo].[tbl_mobile_main_topcatecode]"
sqlStr = sqlStr + " where gnbcode=" + CStr(gnbcode)



rsget.Open sqlStr,dbget,1
    ItemExists = Not rsget.Eof
rsget.Close

if (ItemExists) then
    sqlStr = " update db_sitemaster.[dbo].[tbl_mobile_main_topcatecode]" + VbCrlf
    sqlStr = sqlStr + " set gnbname='" + gnbname + "'" + VbCrlf
    sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " where gnbcode=" + CStr(gnbcode) + VbCrlf
    
    dbget.Execute sqlStr
else
    sqlStr = " insert into db_sitemaster.[dbo].[tbl_mobile_main_topcatecode]" + VbCrlf
    sqlStr = sqlStr + " (gnbcode,gnbname,isusing)"+ VbCrlf
    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " " + CStr(gnbcode) + VbCrlf
    sqlStr = sqlStr + " ,'" + gnbname + "'" + VbCrlf
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