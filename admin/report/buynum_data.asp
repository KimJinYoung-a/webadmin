<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->

사용중지
<%

dbget.close()	:	response.End


dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, nextdateStr
Dim fromDate,toDate,oreport
dim buynum,reyyyy

buynum = request("buynum")
reyyyy = request("reyyyy")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

nowdateStr = CStr(now())

if (buynum="") then buynum = "1"
if (reyyyy="") then reyyyy = "2003"

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)

set oreport = new CBuyNumReport
oreport.FRectRegStart = fromDate
oreport.FRectRegEnd = toDate
oreport.FRectBuyNum = buynum
oreport.FRectYYYY = reyyyy
oreport.FirstBuySellReport


dim buysellavg,buycntavg

if IsNull(oreport.Fsubtotalprice) then oreport.Fsubtotalprice=0
if IsNull(oreport.Fitemno) then oreport.Fitemno=0

if IsNull(oreport.Fsubtotalprice) or (oreport.Fitemno=0) then
	buysellavg = 0
else
	buysellavg = CLng(oreport.Fsubtotalprice / oreport.Fitemno)
end if

	buycntavg = Round((oreport.Ftotalcnt / oreport.Fcnt),2)

dim i
%>

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;구매 횟수 : <input type="text" name="buynum" value="<% = buynum %>" size="4">
		&nbsp;년간 :
		<select name="reyyyy">
			<option value="2003" <% if reyyyy = "2003" then response.write "selected" %>>2003</option>
			<option value="2004" <% if reyyyy = "2004" then response.write "selected" %>>2004</option>
			<option value="2005" <% if reyyyy = "2005" then response.write "selected" %>>2005</option>
			<option value="2006" <% if reyyyy = "2006" then response.write "selected" %>>2006</option>
		</select>
		</td>
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
		</td>
	</tr>
	</form>
</table>
<br>
<div class="a">회원 구매 비율</div>
<table width="800" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td align="center" height="25">비고</td>
	<td align="center">건 수</td>
	<td align="center">총 액</td>
	<td align="center">평균</td>
</tr>
<tr class="a" bgcolor="#FFFFFF">
		<td align="center" height="25"><% = buynum %>번째 구매</td>
		<td align="center"><%= FormatNumber(CLng(oreport.Fitemno),0) %>건</td>
		<td align="center"><%= FormatNumber(oreport.Fsubtotalprice,0) %>원</td>
		<td align="center"><% = FormatNumber(buysellavg,0) %>원</td>
</tr>
<tr class="a" bgcolor="#FFFFFF">
		<td align="center" height="25"><% = reyyyy %>년도</td>
		<td align="center"><%= FormatNumber(CLng(oreport.Fcnt),0) %>건</td>
		<td align="center"><%= FormatNumber(oreport.Ftotalcnt,0) %>건</td>
		<td align="center">평균<% = buycntavg %>건</td>
</tr>
</table><br><br>
<table class="a" >
<tr>
	<td>
* 구매횟수별평균객단가 : 구매횟수에 숫자를 기입한 다음 조회<br>
		건수(총구매횟수), 총액(총구매액), 평균(평균객단가)
	</td>
</tr>
<tr>
	<td>
**	년간 평균구매횟수 : 년간에 년을 선택한후 조회<br>
		건수(구매자총수), 총액(총구매횟수), 평균(평균구매횟수)
	</td>
</tr>
</table>
<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->