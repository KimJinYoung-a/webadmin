<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.23 한용민 생성
'	Description : 오거나이저
'#######################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->
<%
dim poscode 
dim posname
dim imagetype     
dim imagewidth 
dim isusing    
dim imageheight,imagecount

poscode   = request.Form("poscode")
posname   = html2Db(request.Form("posname"))
imagetype  = request.Form("imagetype")
imagetype   = request.Form("imagetype")
imagewidth= request.Form("imagewidth")
isusing   = request.Form("isusing")
imageheight= request.Form("imageheight")
imagecount= request.Form("imagecount")

dim sqlStr, ItemExists

sqlStr = "select top 1 poscode,posname,imagetype,imagewidth,imageheight,isusing,imagecount from db_diary2010.dbo.tbl_organizer_poscode"
sqlStr = sqlStr + " where poscode=" + CStr(poscode)

rsget.Open sqlStr,dbget,1
    ItemExists = Not rsget.Eof
rsget.Close

if (ItemExists) then
    sqlStr = " update db_diary2010.dbo.tbl_organizer_poscode" + VbCrlf
    sqlStr = sqlStr + " set posname='" + posname + "'" + VbCrlf
    sqlStr = sqlStr + " ,imagetype='" + imagetype + "'" + VbCrlf
    sqlStr = sqlStr + " ,imagewidth=" + imagewidth + "" + VbCrlf
    sqlStr = sqlStr + " ,imageheight=" + imageheight + "" + VbCrlf
    sqlStr = sqlStr + " ,imagecount=" + imagecount + "" + VbCrlf
    sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " where poscode=" + CStr(poscode) + VbCrlf
    
    response.write sqlStr
    dbget.Execute sqlStr
else
    sqlStr = " insert into db_diary2010.dbo.tbl_organizer_poscode" + VbCrlf
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
    dbget.Execute sqlStr
end if


dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
