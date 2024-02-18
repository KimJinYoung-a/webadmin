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
dim poscode 
dim posname
dim posVarname 
dim linktype   
dim fixtype   
dim imagewidth 
dim isusing    
dim imageheight
dim useSet

poscode   = request.Form("poscode")
posname   = html2Db(request.Form("posname"))
posVarname= request.Form("posVarname")
linktype  = request.Form("linktype")
fixtype   = request.Form("fixtype")
imagewidth= request.Form("imagewidth")
isusing   = request.Form("isusing")
imageheight= request.Form("imageheight")
useSet= request.Form("useSet")

dim sqlStr, ItemExists

sqlStr = "select top 1 * from [db_sitemaster].[dbo].tbl_StreetMain_poscode"
sqlStr = sqlStr + " where poscode=" + CStr(poscode)

rsget.Open sqlStr,dbget,1
    ItemExists = Not rsget.Eof
rsget.Close

if (ItemExists) then
    sqlStr = " update [db_sitemaster].[dbo].tbl_StreetMain_poscode" + VbCrlf
    sqlStr = sqlStr + " set posname='" + posname + "'" + VbCrlf
    sqlStr = sqlStr + " ,posVarname='" + posVarname + "'" + VbCrlf
    sqlStr = sqlStr + " ,linktype='" + linktype + "'" + VbCrlf
    sqlStr = sqlStr + " ,fixtype='" + fixtype + "'" + VbCrlf
    sqlStr = sqlStr + " ,imagewidth='" + imagewidth + "'" + VbCrlf
    sqlStr = sqlStr + " ,imageheight='" + imageheight + "'" + VbCrlf
    sqlStr = sqlStr + " ,useSet=" + useSet + VbCrlf
    sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " where poscode=" + CStr(poscode) + VbCrlf
    
    dbget.Execute sqlStr
else
    sqlStr = " insert into [db_sitemaster].[dbo].tbl_StreetMain_poscode" + VbCrlf
    sqlStr = sqlStr + " (poscode,posname,posVarname,linktype,fixtype,imagewidth,imageheight,useSet,isusing)"+ VbCrlf
    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " " + CStr(poscode) + VbCrlf
    sqlStr = sqlStr + " ,'" + posname + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + posVarname + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + linktype + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + fixtype + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + imagewidth + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + imageheight + "'" + VbCrlf
    sqlStr = sqlStr + " ," + useSet + VbCrlf
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