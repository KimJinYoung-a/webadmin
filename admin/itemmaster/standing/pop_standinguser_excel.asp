<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 정기구독 대상자 발송 엑셀 다운로드
' History : 2016.06.16 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/standing/item_standing_cls.asp"-->
<%
dim itemid, itemoption, i, menupos, page, orderserial, userid, sendstatus, arrlist
dim reserveDlvDate, reserveidx, reserveItemID, reserveItemOption, reserveItemName, regadminid, regdate
dim lastadminid, lastupdate, username, isusing, reloading, jukyogubun
	itemid = getNumeric(requestcheckvar(request("itemid"),10))
	reserveitemid = getNumeric(requestcheckvar(request("reserveitemid"),10))
	menupos = getNumeric(requestcheckvar(request("menupos"),10))
	itemoption = requestcheckvar(request("itemoption"),4)
	page = getNumeric(requestcheckvar(request("page"),10))
	reserveidx = getNumeric(requestcheckvar(request("reserveidx"),10))
	orderserial = requestcheckvar(request("orderserial"),11)
	username = requestcheckvar(request("username"),32)
	userid = requestcheckvar(request("userid"),32)
	isusing = requestcheckvar(request("isusing"),1)
	reloading = requestcheckvar(request("reloading"),2)
	sendstatus = requestcheckvar(request("sendstatus"),10)
	jukyogubun = requestcheckvar(request("jukyogubun"),16)

if reloading="" and isusing="" then isusing="Y"
if page="" then page=1

dim ouser
set ouser = new Citemstanding
	ouser.FPageSize = 100000
	ouser.FCurrPage = 1
	ouser.FRectItemID = itemid
	ouser.FRectreserveitemid = reserveitemid
	ouser.FRectitemoption = itemoption
	ouser.FRectreserveidx = reserveidx
	ouser.FRectorderserial = orderserial
	ouser.FRectusername = username
	ouser.FRectuserid = userid
	ouser.FRectisusing = isusing
	ouser.FRectsendstatus = sendstatus
	ouser.FRectjukyogubun = jukyogubun
	ouser.fitemstanding_user_getrows

if ouser.ftotalcount >0 then
	arrlist = ouser.fstandingarr
end if

downPersonalInformation_rowcnt=ouser.ftotalcount

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_정기구독_배송리스트_" & Left(CStr(now()),10) & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=euc-kr" />
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#DDDDDD" border=1>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= ouser.ftotalcount %></b>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
    <td>발행회차Vol.(번호)</td>
    <td>배송상품코드</td>
    <td>배송옵션코드</td>
    <td>배송상품명</td>
	<td>적요</td>
    <td>주문번호</td>
    <td>수량</td>
    <td>아이디</td>
    <td>이름</td>
	<td>상태</td>
	<td>발송일</td>
	<td>사용여부</td>
    <td>판매용상품코드</td>
    <td>판매용옵션코드</td>
	<td>우편번호</td>
	<td>주소</td>
	<td>상세주소</td>
	<td>전화번호</td>
	<td>핸드폰</td>
</tr>

<% if isarray(arrlist) then %>
	<%
	for i=0 to ubound(arrlist,2)
	%>
	<tr bgcolor="<%=chkIIF(arrlist(15,i)="Y","#FFFFFF","#DDDDDD")%>" align="center">
	    <td><%= arrlist(3,i) %></td>
	    <td class='txt'><%= arrlist(21,i) %></td>
		<td class='txt'><%= arrlist(22,i) %></td>
		<td class='txt' align="left"><%= arrlist(23,i) %></td>
		<td><%= getjukyoname(arrlist(4,i)) %></td>
		<td class='txt'><%= arrlist(5,i) %></td>
		<td><%= arrlist(7,i) %></td>
		<td class='txt'><%= arrlist(6,i) %></td>
		<td><%= arrlist(10,i) %></td>
		<td><%= getsendstatusname(arrlist(8,i)) %></td>
		<td>
	    	<%= left(arrlist(9,i),10) %>
	    	<Br><%= mid(arrlist(9,i),12,11) %>
		</td>
		<td><%= arrlist(16,i) %></td>
		<td><%= arrlist(1,i) %></td>
		<td class='txt'><%= arrlist(2,i) %></td>
		<td class='txt' align="left"><%= arrlist(11,i) %></td>
		<td class='txt' align="left"><%= arrlist(12,i) %></td>
		<td class='txt' align="left"><%= arrlist(13,i) %></td>
		<td class='txt'><%= arrlist(14,i) %></td>
		<td class='txt'><%= arrlist(15,i) %></td>
	</tr>
	<%
	if i mod 3000 = 0 then
		Response.Flush		' 버퍼리플래쉬
	end if
	Next
	%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="19" align="center">검색결과가 없습니다.</td>
	</tr>
<% end if %>
</table>
</html>

<%
set ouser=nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->