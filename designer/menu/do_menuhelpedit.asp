<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->

<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim id, menuname, viewidx, linkurl, menucolor, isusing, menuposnotice, menuposhelp

id	= request.form("id")
menuname	= html2db(request.form("menuname"))
viewidx	= request.form("viewidx")
linkurl	= html2db(request.form("linkurl"))
menucolor	= html2db(request.form("menucolor"))
isusing	= request.form("isusing")
menuposnotice = html2db(request.form("menuposnotice"))
menuposhelp = html2db(request.form("menuposhelp"))


dim sqlStr

sqlStr = " update [db_partner].[dbo].tbl_partner_menu"	+ VbCrlf
sqlStr = sqlStr + " set menuname='" + menuname + "'"	+ VbCrlf
sqlStr = sqlStr + " , viewidx=" + viewidx + ""	+ VbCrlf
sqlStr = sqlStr + " , linkurl='" + linkurl + "'"	+ VbCrlf
sqlStr = sqlStr + " , menucolor='" + menucolor + "'"	+ VbCrlf
sqlStr = sqlStr + " , isusing='" + isusing + "'"	+ VbCrlf
sqlStr = sqlStr + " , menuposnotice='" + menuposnotice + "'"	+ VbCrlf
sqlStr = sqlStr + " , menuposhelp='" + menuposhelp + "'"	+ VbCrlf
sqlStr = sqlStr + " where id=" + CStr(id)

dbget.Execute sqlStr
%>

<script>alert('저장되었습니다.');</script>
<script>location.replace('<%= refer %>');</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->