<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<%
'####################################################
' Description : scm 매뉴
' History : 최초생성자모름
'			2017.04.10 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->

<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim id, menuname, viewidx, linkurl, menucolor, isusing, menuposnotice, menuposhelp

id	= requestCheckVar(request.form("id"),10)
menuname	= requestCheckVar(html2db(request.form("menuname")),32)
viewidx	= requestCheckVar(request.form("viewidx"),10)
linkurl	= requestCheckVar(html2db(request.form("linkurl")),128)
menucolor	= requestCheckVar(html2db(request.form("menucolor")),16)
isusing	= requestCheckVar(request.form("isusing"),1)
menuposnotice = request.form("menuposnotice")
menuposhelp = request.form("menuposhelp")

if menuposnotice <> "" then
	if checkNotValidHTML(menuposnotice) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "	location.replace('"& refer &"');"
	response.write "</script>"
	dbget.close()	:	response.End
	end if
end if
if menuposhelp <> "" then
	if checkNotValidHTML(menuposhelp) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "	location.replace('"& refer &"');"
	response.write "</script>"
	dbget.close()	:	response.End
	end if
end if

dim sqlStr
sqlStr = " update [db_partner].[dbo].tbl_partner_menu"	+ VbCrlf
sqlStr = sqlStr + " set menuname='" + menuname + "'"	+ VbCrlf
sqlStr = sqlStr + " , viewidx=" + viewidx + ""	+ VbCrlf
sqlStr = sqlStr + " , linkurl='" + linkurl + "'"	+ VbCrlf
sqlStr = sqlStr + " , menucolor='" + menucolor + "'"	+ VbCrlf
sqlStr = sqlStr + " , isusing='" + isusing + "'"	+ VbCrlf
sqlStr = sqlStr + " , menuposnotice='" + html2db(menuposnotice) + "'"	+ VbCrlf
sqlStr = sqlStr + " , menuposhelp='" + html2db(menuposhelp) + "'"	+ VbCrlf
sqlStr = sqlStr + " where id=" + CStr(id)

dbget.Execute sqlStr
%>

<script type='text/javascript'>alert('저장되었습니다.');</script>
<script type='text/javascript'>location.replace('<%= refer %>');</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
	session.codePage = 949
%>
