<%@ language=vbscript %>
<% option Explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"

Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'###########################################################
' Description : 매주 화요일 휴면 전환 예정고객
' History : 2023.06.30 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/crm/DormantSleepCls.asp"-->
<%
dim page, research, yyyy1,mm1,dd1, fromDate,i, arrLIst
    page = RequestCheckVar(getNumeric(request("page")),10)
    research = RequestCheckVar(request("research"),2)
    yyyy1 = RequestCheckVar(request("yyyy1"),4)
    mm1   = RequestCheckVar(request("mm1"),2)
    dd1   = RequestCheckVar(request("dd1"),2)

if (yyyy1="") then yyyy1 = Cstr(Year(dateadd("d",-1,date())))
if (mm1="") then mm1 = Cstr(Month(dateadd("d",-1,date())))
if (dd1="") then dd1 = Cstr(day(dateadd("d",-1,date())))
if (page="") then page=1
fromDate = CStr(DateSerial(yyyy1, mm1, dd1))

dim odormantsleep
set odormantsleep = new CDormantSleepList
    odormantsleep.FCurrPage = page
    odormantsleep.FPageSize = 200000
    odormantsleep.FRectStartDate = fromDate
    odormantsleep.GetDormantSleepNotPaging

if odormantsleep.FTotalCount>0 then
    arrLIst=odormantsleep.fArrLIst
end if

downPersonalInformation_rowcnt=odormantsleep.ftotalcount

%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENDormantSleepLIST" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
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
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td align="left" colspan="8">
        ※ <%= CStr(DateSerial(yyyy1, mm1, dd1)) %>일에 휴면으로 전환될 예정인 고객 리스트 입니다.
    </td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%= odormantsleep.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>고객아이디</td>
    <td>고객명</td>
    <td>회원등급</td>
    <td>푸시수신</td>
    <td>문자수신</td>
    <td>이메일수신</td>
    <td>마지막로그인</td>
    <td>현재마일리지</td>
    <td>비고</td>
</tr>
<% if isarray(arrLIst) then %>
<% for i=0 to ubound(arrLIst,2) %>
<tr bgcolor="#FFFFFF" align="center">
    <td class="txt">
        <% if C_CriticInfoUserLV1 then %>
            <%= arrLIst(0,i) %>
        <% else %>
            <%= printUserId(arrLIst(0,i),2,"*") %>
        <% end if %>
    </td>
    <td class="txt">
        <% if C_CriticInfoUserLV1 then %>
            <%= arrLIst(1,i) %>
        <% else %>
            <%= printUserId(arrLIst(1,i),2,"*") %>
        <% end if %>
    </td>
    <td><%= arrLIst(2,i) %></td>
    <td><%= arrLIst(3,i) %></td>
    <td><%= arrLIst(4,i) %></td>
    <td><%= arrLIst(5,i) %></td>
    <td class="txt"><%= arrLIst(6,i) %></td>
    <td><%= FormatNumber(arrLIst(7,i), 0) %></td>
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
		<td colspan="9" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>
<%
set odormantsleep = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->