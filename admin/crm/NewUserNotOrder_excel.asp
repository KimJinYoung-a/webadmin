<%@ language=vbscript %>
<% option Explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"

Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'###########################################################
' Description : 신규가입후 구매이력 없는고객
' History : 2023.06.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/crm/NewUserNotOrderCls.asp"-->
<%
dim page, research, yyyy1,mm1,dd1,yyyy2,mm2,dd2, fromDate,toDate, i, sixmonthago, arrLIst
    page = RequestCheckVar(getNumeric(request("page")),10)
    research = RequestCheckVar(request("research"),2)
    yyyy1 = RequestCheckVar(request("yyyy1"),4)
    mm1   = RequestCheckVar(request("mm1"),2)
    dd1   = RequestCheckVar(request("dd1"),2)
    yyyy2 = RequestCheckVar(request("yyyy2"),4)
    mm2   = RequestCheckVar(request("mm2"),2)
    dd2   = RequestCheckVar(request("dd2"),2)
    sixmonthago   = RequestCheckVar(request("sixmonthago"),2)

if (yyyy1="") then yyyy1 = Cstr(Year(dateadd("d",-1,date())))
if (mm1="") then mm1 = Cstr(Month(dateadd("d",-1,date())))
if (dd1="") then dd1 = Cstr(day(dateadd("d",-1,date())))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (page="") then page=1
fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim oNewUser
set oNewUser = new CNewUserNotOrderList
    oNewUser.FCurrPage = page
    oNewUser.FPageSize = 200000
    oNewUser.FRectStartDate = fromDate
    oNewUser.FRectEndDate   = toDate
    oNewUser.FRectsixmonthago   = sixmonthago
    oNewUser.GetNewUserNotOrderNotPaging

if oNewUser.FTotalCount>0 then
    arrLIst=oNewUser.fArrLIst
end if

downPersonalInformation_rowcnt=oNewUser.ftotalcount

%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENNewUserNotOrderLIST" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
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
        ※ <%= CStr(DateSerial(yyyy1, mm1, dd1)) %>~<%= CStr(DateSerial(yyyy2, mm2, dd2)) %>에 신규가입후 현재까지 구매이력 없는 고객 리스트 입니다.<br>느린 매뉴 입니다. 클릭후 기다려 주세요.
    </td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%= oNewUser.FTotalCount %></b>
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
</body>
</html>
<%
set oNewUser = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->