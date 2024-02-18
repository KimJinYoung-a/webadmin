<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  멤버십카드 일일데이터
' History : 2015.06.18 원승현 개발
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/report/MemberShipCardDailyCls.asp"-->

<%
	Dim defaultdate1, yyyy1, mm1, dd1, yyyy2, mm2, dd2, MemberShipCardDailylist, i, strTemp, strXML, ChartViDi, strDay, strWeb, strMobile, strApp, strWebLen, strMobileLen, strAppLen, strDate, strDateLen, striOs, striOsLen, strAnd, strAndLen


	defaultdate1 = dateadd("d",-10,year(now) & "-" &month(now) & "-" & day(now))		'날짜값이 없을때 기본값으로 10이전까지 검색	
	yyyy1 = request("yyyy1")
	if yyyy1 = "" then yyyy1 = left(defaultdate1,4)
	mm1 = request("mm1")
	if mm1 = "" then mm1 = mid(defaultdate1,6,2)
	dd1 = request("dd1")
	if dd1 = "" then dd1 = right(defaultdate1,2)	
	yyyy2 = request("yyyy2")
	if yyyy2 = "" then yyyy2 = year(now)
	mm2 = request("mm2")
	if mm2 = "" then 
		mm2 = month(now)
	end if
	dd2 = request("dd2")
	if dd2 = "" then dd2 = day(now)


	set MemberShipCardDailylist = new CMemberShipCardDaily
	MemberShipCardDailylist.FRectFromDate = dateserial(yyyy1,mm1,dd1)
	MemberShipCardDailylist.FRectToDate = dateserial(yyyy2,mm2,dd2)
	MemberShipCardDailylist.GetMemberShipCardDailyReport()


%>



<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form action="" name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">			
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red">- 하루전 데이터까지만 검색가능합니다.<br>- 데이터는 2015년 5월1일부터 검색 가능합니다.</font>	
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<% if MemberShipCardDailylist.ftotalcount > 0 then %>			
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan="2">날짜</td>
		<td rowspan="2">카드발급수</td>
		<td colspan="4">멤버십카드 적립/사용률</td>
		<td colspan="2">온라인 마일리지 전환(모바일+PC)</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>적립건수</td>
		<td>적립된 포인트</td>
		<td>사용건수</td>
		<td>사용된 포인트</td>
		<td>전환건수</td>
		<td>전환된 포인트</td>
	</tr>
	<% for i = 0 to MemberShipCardDailylist.ftotalcount -1 %>
		<tr bgcolor="#FFFFFF" align="center">
			<td><%=MemberShipCardDailylist.FItemList(i).Fregdate%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FCardRegCnt,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FCardSavingCnt,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FCardSavingPoint,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FCardUsingCnt,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FCardUsingPoint,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FChangeOnlineCnt,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FChangeOnlinePoint,0)%></td>
		</tr>
	<% next %>
	</table>
<% else %>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#FFFFFF">
		<td >검색 결과가 없습니다.</td>
	</tr>
	</table>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->