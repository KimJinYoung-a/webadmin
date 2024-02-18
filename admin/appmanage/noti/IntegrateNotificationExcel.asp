<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 통합알림신청자
' Hieditor : 2022.12.26 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/noti/IntegrateNotificationCls.asp" -->
<%
Dim page, i, oNoti, linkCode, isusing, reload, notiType, sendType, userId, menupos, arrList
	menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)
	page = requestcheckvar(getNumeric(trim(request("page"))),10)
	linkCode = requestcheckvar(getNumeric(trim(request("linkCode"))),10)
	isusing = requestcheckvar(trim(request("isusing")),1)
	reload = requestcheckvar(request("reload"),2)
    notiType=requestcheckvar(trim(request("notiType")),32)
	sendType=requestcheckvar(trim(request("sendType")),16)
	userId=requestcheckvar(trim(request("userId")),32)

if page = "" then page = 1
if reload="" and isusing="" then isusing="Y"

set oNoti = new cNotiList
	oNoti.FPageSize = 1000000
	oNoti.FCurrPage = page
	oNoti.frectlinkCode = linkCode
	oNoti.frectisusing = isusing
	oNoti.frectnotiType = notiType
	oNoti.frectsendType = sendType
	oNoti.frectuserId = userId
	oNoti.fIntegrateNotificationListNotPaging()

if oNoti.FTotalCount>0 then
    arrLIst=oNoti.fArrLIst
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENIntegrateNotificationLIST" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>
<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%= oNoti.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>신청<br>번호</td>
	<td>구분</td>
	<td>관련코드</td>
	<td>발송구분</td>
	<td>고객아이디</td>
	<td>신청채널</td>
	<td>등록일</td>
	<td>비고</td>
</tr>
<% if isarray(arrLIst) then %>
    <% for i=0 to ubound(arrLIst,2) %>

    <tr bgcolor="#FFFFFF" align="center">
    	<td><%= arrLIst(0,i) %></td>
    	<td>
			<%= getNotiType(arrLIst(1,i)) %>
    	</td>
    	<td>
			<%= arrLIst(2,i) %>
    	</td>
    	<td>
			<%= getSendType(arrLIst(3,i)) %>
    	</td>
    	<td>
			<%= arrLIst(4,i) %>
    	</td>
    	<td>
			<%= getIntegrateNotificationDevice(arrLIst(5,i)) %>
    	</td>
    	<td>
			<%= arrLIst(6,i) %>
    	</td>
    	<td></td>
    </tr>
    <%
    if i mod 1000 = 0 then
        Response.Flush		' 버퍼리플래쉬
    end if
    next
    %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="8" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</table>

</body>
</html>
<%
session.codePage = 949
set oNoti = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->