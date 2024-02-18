<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2009.04.01 한용민 개발
' culturestation 포스코드 저장 리스트
'###########################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim poscode 
dim posname
dim imagetype     
dim imagewidth 
dim isusing    
dim imageheight,imagecount

poscode   = RequestCheckvar(request.Form("poscode"),10)
posname   = html2Db(RequestCheckvar(request.Form("posname"),64))
imagetype  = RequestCheckvar(request.Form("imagetype"),10)
'imagetype   = RequestCheckvar(request.Form("imagetype"),10)
imagewidth= RequestCheckvar(request.Form("imagewidth"),10)
isusing   = RequestCheckvar(request.Form("isusing"),1)
imageheight= RequestCheckvar(request.Form("imageheight"),10)
imagecount= RequestCheckvar(request.Form("imagecount"),10)

dim sqlStr, ItemExists

sqlStr = "select top 1 poscode,posname,imagetype,imagewidth,imageheight,isusing,imagecount from db_academy.dbo.tbl_corner_poscode"
sqlStr = sqlStr + " where poscode=" + CStr(poscode)

rsACADEMYget.Open sqlStr,dbACADEMYget,1
    ItemExists = Not rsACADEMYget.Eof
rsACADEMYget.Close

if (ItemExists) then
    sqlStr = " update db_academy.dbo.tbl_corner_poscode" + VbCrlf
    sqlStr = sqlStr + " set posname='" + posname + "'" + VbCrlf
    sqlStr = sqlStr + " ,imagetype='" + imagetype + "'" + VbCrlf
    sqlStr = sqlStr + " ,imagewidth=" + imagewidth + "" + VbCrlf
    sqlStr = sqlStr + " ,imageheight=" + imageheight + "" + VbCrlf
    sqlStr = sqlStr + " ,imagecount=" + imagecount + "" + VbCrlf
    sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " where poscode=" + CStr(poscode) + VbCrlf
    
    response.write sqlStr
    dbACADEMYget.Execute sqlStr
else
    sqlStr = " insert into db_academy.dbo.tbl_corner_poscode" + VbCrlf
    sqlStr = sqlStr + " (poscode,posname,imagetype,imagewidth,imageheight,isusing,imagecount)"+ VbCrlf
    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " " + CStr(poscode) + VbCrlf
    sqlStr = sqlStr + " ,'" + posname + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + imagetype + "'" + VbCrlf
    sqlStr = sqlStr + " ," + imagewidth + "" + VbCrlf
    sqlStr = sqlStr + " ," + imageheight + "" + VbCrlf
    sqlStr = sqlStr + " ,'" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " ," + imagecount + "" + VbCrlf
    sqlStr = sqlStr + " )" + VbCrlf
    
    response.write sqlStr    
    dbACADEMYget.Execute sqlStr
end if


dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->