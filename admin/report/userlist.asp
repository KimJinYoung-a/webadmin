<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : [통계]회원관련>>회원가입현황
' History : 최초생성자모름
'			2017.04.14 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/report/userjoincls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr, page
dim eventinclude, research
dim rpttype, totalSum, totalMobile
dim iTotal, iOnline, iMobile

yyyy1 = requestCheckVar(request("yyyy1"),4)
mm1 = requestCheckVar(request("mm1"),2)
dd1 = requestCheckVar(request("dd1"),2)
yyyy2 = requestCheckVar(request("yyyy2"),4)
mm2 = requestCheckVar(request("mm2"),2)
dd2 = requestCheckVar(request("dd2"),2)
rpttype = requestCheckVar(request("rpttype"),32)
page = requestCheckVar(request("page"),10)
eventinclude = requestCheckVar(request("eventinclude"),2)
research = requestCheckVar(request("research"),2)

if (research="") then eventinclude="on"
if page="" then page=1

nowdateStr = CStr(now())


if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

if (rpttype="") then rpttype="day"

If (IsDate( yyyy1 + "-" + mm1 + "-" + dd1)) = "True" and (IsDate( yyyy2 + "-" + mm2 + "-" + dd2)) = "True" Then
	startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
	nextdateStr = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)
Else	
	response.write "<script>alert('해당월에 해당하는 날짜가 없습니다.\n다시 선택하세요')</script>"
End If


dim oneuserjoin
set oneuserjoin = new UserJoinClass
oneuserjoin.FPageSize=500
oneuserjoin.FCurrPage = page
oneuserjoin.FRectStart = startdateStr
oneuserjoin.FRectEnd =  nextdateStr
oneuserjoin.FRectGroup = rpttype
oneuserjoin.FRectEventInclude = eventinclude
oneuserjoin.getdayReport

dim ix,p1, p2, p3
%>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		Report :
		<select name="rpttype" >
	     <option value='year' <%if rpttype="year" then response.write " selected"%>>year</option>
		 <option value='month' <%if rpttype="month" then response.write " selected"%>>month</option>
	     <option value='day' <%if rpttype="day" then response.write " selected"%>>day</option>
	     <option value='all' <%if rpttype="all" then response.write " selected"%>>all</option>
	   </select>
		기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="checkbox" name="eventinclude" <% if eventinclude="on" then response.write "checked" %> >이벤트 경로 유입 가입 포함
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="800" cellspacing="1" cellpadding="3" bgcolor="#EFBE00" class="a">
    <tr align="center">
      <td class="a"><font color="#FFFFFF">기간</font></td>
      <td class="a" colspan="3"><font color="#FFFFFF">내용</font></td>
    </tr>
<%
	totalSum = 0: totalMobile = 0
	for ix=0 to oneuserjoin.FResultCount-1
		iTotal	= oneuserjoin.FItemList(ix).Fcount
		iMobile	= oneuserjoin.FItemList(ix).FcountMobile
		iOnline	= iTotal-iMobile

		if oneuserjoin.maxt<>0 then
			p1 = Clng(iTotal/oneuserjoin.maxt*100)
			p2 = Clng(iOnline/oneuserjoin.maxt*100)
			p3 = Clng(iMobile/oneuserjoin.maxt*100)
		end if
		totalSum = totalSum + iTotal
		totalMobile = totalMobile + iMobile
%>
<tr bgcolor="#FFFFFF" height="10">
	<td width="80" rowspan="3">
		<%= oneuserjoin.FItemList(ix).Fdatestr %>
	</td>
	<td width="80">전체</td>
	<td width="560">
		<div align="left"> <img src="/images/dot1.gif" height="4" width="<%= p1 %>%" title="<%= p1 %>%"></div>
	</td>
	<td width="80">
		<%= FormatNumber(iTotal,0) %>명
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="10">
	<td>온라인</td>
	<td width="600">
		<div align="left"> <img src="/images/dot2.gif" height="4" width="<%= p2 %>%" title="<%= cLng(iOnline/iTotal*100) %>%"></div>
	</td>
	<td>
		<%= FormatNumber(iOnline,0) %>명
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="10">
	<td>모바일</td>
	<td width="600">
		<div align="left"> <img src="/images/dot3.gif" height="4" width="<%= p3 %>%" title="<%= cLng(iMobile/iTotal*100) %>%"></div>
	</td>
	<td>
		<%= FormatNumber(iMobile,0) %>명
	</td>
</tr>
<%
	next

	if totalSum>0 then
%>
<tr bgcolor="#F8F8F8" height="10">
	<td>총계</td>
	<td colspan="2" align="right">
		전체<br />
		온라인<br />
		모바일
	</td>
	<td>
		<%= FormatNumber(totalSum,0) %>명<br />
		<%= FormatNumber(totalSum-totalMobile,0) %>명<br />
		<%= FormatNumber(totalMobile,0) %>명
	</td>
</tr>
<%
	end if
%>
</table>

<%
set oneuserjoin = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->