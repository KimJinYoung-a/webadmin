<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2010.04.02 한용민 생성
' culturestation 레드리본 저장 리스트
'###########################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/giftplus/giftplus_cls.asp"-->
<%
dim poscode ,posname, imagetype ,imagewidth ,isusing ,imageheight, imagecount ,sqlStr, ItemExists
	poscode   = request.Form("poscode")
	posname   = html2Db(request.Form("posname"))
	imagetype  = request.Form("imagetype")
	imagetype   = request.Form("imagetype")
	imagewidth= request.Form("imagewidth")
	isusing   = request.Form("isusing")
	imageheight= request.Form("imageheight")
	imagecount= request.Form("imagecount")

sqlStr = "select top 1 poscode,posname,imagetype,imagewidth,imageheight,isusing,imagecount from db_giftplus.dbo.tbl_giftplus_poscode"
sqlStr = sqlStr + " where poscode=" + CStr(poscode)

rsget.Open sqlStr,dbget,1
    ItemExists = Not rsget.Eof
rsget.Close

if (ItemExists) then
    sqlStr = " update db_giftplus.dbo.tbl_giftplus_poscode" + VbCrlf
    sqlStr = sqlStr + " set posname='" + posname + "'" + VbCrlf
    sqlStr = sqlStr + " ,imagetype='" + imagetype + "'" + VbCrlf
    sqlStr = sqlStr + " ,imagewidth=" + imagewidth + "" + VbCrlf
    sqlStr = sqlStr + " ,imageheight=" + imageheight + "" + VbCrlf
    sqlStr = sqlStr + " ,imagecount=" + imagecount + "" + VbCrlf
    sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " where poscode=" + CStr(poscode) + VbCrlf
    
    'response.write sqlStr
    dbget.Execute sqlStr
else
    sqlStr = " insert into db_giftplus.dbo.tbl_giftplus_poscode" + VbCrlf
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
    
    'response.write sqlStr    
    dbget.Execute sqlStr
end if


dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
