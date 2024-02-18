<%@ language=vbscript %>
<%
option explicit
Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'#######################################################
' Description : 기프티콘/기프팅 금액권내역
' History	:  강준구 생성
'              2023.05.23 한용민 수정(엑셀다운로드 전체 다운로드 가능하게 변경함. 소스 표준소스로 변경.)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/giftCls.asp"-->

<%
Dim iCurrentpage, intLoop, arrList, GiftStatisticlist, GiftStatisticshortlist, i, iTotCnt1, iTotCnt, vSDate, vEDate, page
dim vGubun, vOrderSerial, vUserID, vUserName, vReqHP, vReqHP1, vReqHP2, vReqHP3, vTotalSum
dim vSumTemp
	vTotalSum = "x"
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	page = requestCheckVar(getNumeric(request("page")),10)
	vGubun			= NullFillWith(requestCheckVar(request("gubun"),10),"")
	vOrderSerial	= NullFillWith(requestCheckVar(request("orderserial"),30),"")
	vUserID			= NullFillWith(requestCheckVar(request("userid"),50),"")
	vUserName		= NullFillWith(requestCheckVar(request("username"),100),"")
	vReqHP1			= NullFillWith(requestCheckVar(request("reqhp1"),3),"")
	vReqHP2			= NullFillWith(requestCheckVar(request("reqhp2"),4),"")
	vReqHP3			= NullFillWith(requestCheckVar(request("reqhp3"),4),"")
	If vReqHP1 <> "" AND vReqHP2 <> "" AND vReqHP3 <> "" Then
		vReqHP = vReqHP1 & "-" & vReqHP2 & "-" & vReqHP3
	End If
	vSDate			= NullFillWith(requestCheckVar(request("sdate"),10),"")
	vEDate			= NullFillWith(requestCheckVar(request("edate"),10),"")

if page = "" then page = 1

	Set GiftStatisticlist = new ClsGift
	GiftStatisticlist.FPageSize = "200000"
	GiftStatisticlist.FCurrPage = page
	GiftStatisticlist.FGubun = vGubun
	GiftStatisticlist.FOrderSerial = vOrderSerial
	GiftStatisticlist.FUserID = vUserID
	GiftStatisticlist.FUSerName = vUserName
	GiftStatisticlist.FReqHP = vReqHP
	GiftStatisticlist.FSDate = vSDate
	GiftStatisticlist.FEDate = vEDate
	GiftStatisticlist.FGiftStatisticList_notpaging
	if GiftStatisticlist.ftotalcount>0 then
		arrList=GiftStatisticlist.fArrList
	end if
	iTotCnt = GiftStatisticlist.ftotalcount
	
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
Response.AddHeader "Content-Disposition", "attachment;filename=기프티콘_기프팅_금액권내역_" & Left(CStr(now()),10) & ".xls"

%>

<html>
<head></head>
<body>
<table cellpadding="3" cellspacing="1" border="1" width="100%" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="9">
		Total Count : <b><%= iTotCnt %></b>
	</td>
</tr>
<tr bgcolor="#E6E6E6" align="center">
	<td>결제방법</td>
	<td>티콘/팅 쿠폰번호</td>
	<td>UserID</td>
	<td>수령인</td>
	<td>카드명</td>
	<td>판매가</td>
	<td>실결제액</td>
	<td>카드상태</td>
	<td>등록일</td>
</tr>
<%
	If isarray(arrList) Then
		vSumTemp = 0
		For i = 0 To ubound(arrList,2)
%>
		<tr bgcolor="FFFFFF">
			<td align="center">
				<% if arrList(1,i)="550" then %>
					기프팅
				<% elseif arrList(1,i)="560" then %>
					기프티콘
				<% end if %>
			</td>
			<td align="center" style="mso-number-format:'\@'"><%= arrList(1,i) %></td>
			<td align="center"><%= arrList(2,i) %></td>
			<td align="center"><%= arrList(3,i) %></td>
			<td align="center"><%= GetCardName(arrList(4,i)) %></td>
			<td align="center"><%=FormatNumber(arrList(4,i),0) %></td>
			<td align="center"><%=FormatNumber(arrList(5,i),0) %></td>
			<td align="center"><font color="<%= GetCardStatusColor(arrList(6,i)) %>"><%= GetCardStatusName(arrList(6,i)) %></font></td>
			<td> <%= arrList(7,i) %></td>
		</tr>
<%
			vSumTemp = vSumTemp + arrList(5,i)

			if i mod 500 = 0 then
				Response.Flush		' 버퍼리플래쉬
			end if
		Next
	Else
%>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="20" align="center" class="page_link">[데이터가 없습니다.]</td>
		</tr>
<%
	End If
%>
</table>

<%
set GiftStatisticlist = nothing
%>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->