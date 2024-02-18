<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->

<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, nextdateStr
Dim fromDate,toDate,ojumun

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

nowdateStr = CStr(now())

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)

set ojumun = new CJumunMaster

ojumun.FRectFromDate = fromDate
ojumun.FRectToDate = toDate
ojumun.MemberBuyPercent2


dim ps_all, ps_user, ps_biuser, ps_newuser
dim ps_cnt_all, ps_cnt_user, ps_cnt_biuser, ps_cnt_newuser
dim ps_man, ps_woman
dim ps_cnt_man, ps_cnt_woman

if IsNull(ojumun.FMtotalmoney) then ojumun.FMtotalmoney=0
if IsNull(ojumun.FNtotalmoney) then ojumun.FNtotalmoney=0
if IsNull(ojumun.FBtotalmoney) then ojumun.FBtotalmoney=0
if IsNull(ojumun.FMTotalsellcnt) then ojumun.FMTotalsellcnt=0
if IsNull(ojumun.FNTotalsellcnt) then ojumun.FNTotalsellcnt=0
if IsNull(ojumun.FBTotalsellcnt) then ojumun.FBTotalsellcnt=0

if IsNull(ojumun.Ftotalmoney) or (ojumun.Ftotalmoney=0) then
	ps_user = 0
	ps_biuser = 0
	ps_newuser = 0
else
	ps_user = CLng((ojumun.FMtotalmoney / ojumun.Ftotalmoney) * 100)
	ps_biuser = CLng((ojumun.FNtotalmoney / ojumun.Ftotalmoney) * 100)
	ps_newuser = CLng((ojumun.FBtotalmoney / ojumun.Ftotalmoney) * 100)
end if

if IsNull(ojumun.FTotalsellcnt) or (ojumun.FTotalsellcnt=0) then
	ps_cnt_user = 0
	ps_cnt_biuser = 0
	ps_cnt_newuser = 0
else
	ps_cnt_user = CLng((ojumun.FMTotalsellcnt / ojumun.FTotalsellcnt) * 100)
	ps_cnt_biuser = CLng((ojumun.FNTotalsellcnt / ojumun.FTotalsellcnt) * 100)
	ps_cnt_newuser = CLng((ojumun.FBTotalsellcnt / ojumun.FTotalsellcnt) * 100)
end if

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
	<td align="center">백분율</td>
	<td align="center">총 액</td>
	<td align="center">백분율</td>
</tr>
<tr class="a" bgcolor="#FFFFFF">
		<td align="center" height="25">총 구매건수</td>
		<td align="center"><%= ojumun.FTotalsellcnt %></td>
		<td align="center">100%</td>
		<td align="center"><%= FormatNumber(ojumun.Ftotalmoney,0) %></td>
		<td align="center">100%</td>
</tr>
<tr class="a" bgcolor="#FFFFFF">
		<td align="center" height="25">기존 회원 구매건수</td>
		<td align="center"><%= ojumun.FMTotalsellcnt %></td>
		<td align="center"><%= ps_cnt_user %>%</td>
		<td align="center"><%= FormatNumber(ojumun.FMtotalmoney,0) %></td>
		<td align="center"><%= ps_user %>%</td>
</tr>
<tr class="a" bgcolor="#FFFFFF">
		<td align="center" height="25">신규회원 구매건수</td>
		<td align="center"><%= ojumun.FNTotalsellcnt %></td>
		<td align="center"><%= ps_cnt_biuser %>%</td>
		<td align="center"><%= FormatNumber(ojumun.FNtotalmoney,0) %></td>
		<td align="center"><%= ps_biuser %>%</td>
</tr>
<tr class="a" bgcolor="#FFFFFF">
		<td align="center" height="25">비회원 구매건수</td>
		<td align="center"><%= ojumun.FBTotalsellcnt %></td>
		<td align="center"><%= ps_cnt_newuser %>%</td>
		<td align="center"><%= FormatNumber(ojumun.FBtotalmoney,0) %></td>
		<td align="center"><%= ps_newuser %>%</td>
</tr>

</table>
<br>
<%
ojumun.MemberBuySex

if (ojumun.FManTotalCount + ojumun.FWoManTotalCount)=0 then
	ps_cnt_man   = 0
	ps_cnt_woman = 0
else
	ps_cnt_man   = CLng(ojumun.FManTotalCount/(ojumun.FManTotalCount + ojumun.FWoManTotalCount)*100)
	ps_cnt_woman = CLng(ojumun.FWoManTotalCount/(ojumun.FManTotalCount + ojumun.FWoManTotalCount)*100)
end if


if (ojumun.FManTotalMoney + ojumun.FWoManTotalMoney)=0 then
	ps_man   = 0
	ps_woman = 0
else
	ps_man   = CLng(ojumun.FManTotalMoney/(ojumun.FManTotalMoney + ojumun.FWoManTotalMoney)*100)
	ps_woman = CLng(ojumun.FWoManTotalMoney/(ojumun.FManTotalMoney + ojumun.FWoManTotalMoney)*100)
end if


%>

<div class="a">성별 구매 비율</div>
<table width="800" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td align="center" height="25">비고</td>
	<td align="center">건 수</td>
	<td align="center">백분율</td>
	<td align="center">총 액</td>
	<td align="center">백분율</td>
</tr>
<tr class="a" bgcolor="#FFFFFF">
	<td align="center" height="25">회원 구매건수</td>
	<td align="center"><%= ojumun.FManTotalCount + ojumun.FWoManTotalCount %></td>
	<td align="center">100%</td>
	<td align="center"><%= FormatNumber(ojumun.FManTotalMoney + ojumun.FWoManTotalMoney,0) %></td>
	<td align="center">100%</td>
</tr>
<tr class="a" bgcolor="#FFFFFF">
	<td align="center" height="25">남성 회원 구매건수</td>
	<td align="center"><%= ojumun.FManTotalCount %></td>
	<td align="center"><%= ps_cnt_man %>%</td>
	<td align="center"><%= FormatNumber(ojumun.FManTotalMoney,0) %></td>
	<td align="center"><%= ps_man %>%</td>
</tr>
<tr class="a" bgcolor="#FFFFFF">
	<td align="center" height="25">여성 회원 구매건수</td>
	<td align="center"><%= ojumun.FWoManTotalCount %></td>
	<td align="center"><%= ps_cnt_woman %>%</td>
	<td align="center"><%= FormatNumber(ojumun.FWoManTotalMoney,0) %></td>
	<td align="center"><%= ps_woman %>%</td>
</tr>
</table>
<br>

<%
dim sqlStr
dim m_naiStr(5), m_naiStart(5), m_naiEnd(5)
dim m_naiCnt(5), m_naiTot(5), m_sex(5)
dim m_man_naiCnt(5), m_man_naiTot(5)
dim m_woman_naiCnt(5), m_woman_naiTot(5)

dim m_naicnttot, m_naisumtot
dim m_man_naicnttot, m_man_naisumtot
dim m_woman_naicnttot, m_woman_naisumtot
dim m_naisum

m_naiStr(0) ="0~18 세"
m_naiStr(1) ="19~22 미만"
m_naiStr(2) ="23~26세"
m_naiStr(3) ="27~29세"
m_naiStr(4) ="30~32세"
m_naiStr(5) ="33세이상"

m_naiStart(0) =0
m_naiStart(1) =18
m_naiStart(2) =22
m_naiStart(3) =26
m_naiStart(4) =29
m_naiStart(5) =32

m_naiEnd(0) =18
m_naiEnd(1) =22
m_naiEnd(2) =26
m_naiEnd(3) =29
m_naiEnd(4) =32
m_naiEnd(5) =110

sqlStr = "select count(m.orderserial) as cnt, sum(subtotalprice) as sumprice,"
sqlStr = sqlStr + " (year(getdate())-Left(u.juminno,2)-1900) as nai,"
sqlStr = sqlStr + " Left(Right(u.juminno,7),1) as sex"
sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m, [db_user].[dbo].tbl_user_n u"
sqlStr = sqlStr + " where m.regdate > '" & fromDate & "'"
sqlStr = sqlStr + " and m.regdate < '" & ToDate & "'"
sqlStr = sqlStr + " and m.sitename='10x10'"
sqlStr = sqlStr + " and m.cancelyn='N'"
sqlStr = sqlStr + " and m.userid=u.userid"
sqlStr = sqlStr + " and m.userid <> ''"
sqlStr = sqlStr + " and m.ipkumdiv>=4"
sqlStr = sqlStr + " and m.jumundiv<9"
sqlStr = sqlStr + " group by (year(getdate())-Left(u.juminno,2)-1900), Left(Right(u.juminno,7),1)"

dim nai, naicnt, naisum, naisex
dim isexistsRect

rsget.Open sqlStr,dbget,1
	do until rsget.Eof
		nai = rsget("nai")
		naicnt = rsget("cnt")
		naisum = rsget("sumprice")
		naisex = rsget("sex")

		m_naisum = m_naisum +  nai
		m_naicnttot = m_naicnttot + naicnt
		m_naisumtot = m_naisumtot + naisum

		if naisex="1" then
			m_man_naicnttot = m_man_naicnttot + naicnt
			m_man_naisumtot = m_man_naisumtot + naisum
		end if

		if naisex="2" then
			m_woman_naicnttot = m_woman_naicnttot + naicnt
			m_woman_naisumtot = m_woman_naisumtot + naicnt
		end if

		isexistsRect = false


		for i=0 to 17
			if (nai>=m_naiStart(i)) and (nai<m_naiEnd(i)) then
				m_naiCnt(i)=m_naiCnt(i) + naicnt
				m_naiTot(i)=m_naiTot(i) + naisum

				if naisex="1" then
					m_man_naiCnt(i) = m_man_naiCnt(i) + naicnt
					m_man_naiTot(i) = m_man_naiTot(i) + naisum
				end if

				if naisex="2" then
					m_woman_naiCnt(i) = m_woman_naiCnt(i) + naicnt
					m_woman_naiTot(i) = m_woman_naiTot(i) + naisum
				end if

				isexistsRect = true
				exit for
			end if
		next

		if not isexistsRect then
			response.write CStr(nai) + "<br>"
			m_naiCnt(5)=m_naiCnt(5) + naicnt
			m_naiTot(5)=m_naiTot(5) + naisum
		end if

		rsget.MoveNext
	loop
rsget.close
dim tmoney
tmoney = 0
%>
<div class="a">연령별 구매 비율</div>
<table width="800" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td align="center" height="25">비고</td>
	<td align="center">총건수</td>
	<td align="center">백분율</td>
	<td align="center">남성</td>
	<td align="center">여성</td>
	<td align="center">총 액</td>
	<td align="center">백분율</td>
	<td align="center">남성</td>
	<td align="center">여성</td>
</tr>
<% for i=0 to 5 %>
<tr class="a" bgcolor="#FFFFFF">
	<td align="center" height="25"><%= m_naiStr(i) %></td>
	<td align="center"><%= FormatNumber(m_naiCnt(i),0) %></td>
	<td align="center">
	<% if m_naicnttot<>0 then %>
	<%= CLng(m_naiCnt(i)/m_naicnttot*100) %> %
	<% else %>
	0%
	<% end if %>
	</td>
	<td align="center">
		<%= FormatNumber(m_man_naiCnt(i),0) %>
		<% if m_naicnttot<>0 then %>
		(<%= CLng(m_man_naiCnt(i)/m_naicnttot*100) %> %)
		<% else %>
		(0%)
		<% end if %>
	</td>
	<td align="center">
		<%= FormatNumber(m_woman_naiCnt(i),0) %>
		<% if m_naicnttot<>0 then %>
		(<%= CLng(m_woman_naiCnt(i)/m_naicnttot*100) %> %)
		<% else %>
		(0%)
		<% end if %>
	</td>

	<td align="center"><%= FormatNumber(m_naiTot(i),0) %></td>
	<td align="center">
	<% if m_naisumtot<>0 then %>
	<%= CLng(m_naiTot(i)/m_naisumtot*100) %> %
	<% else %>
	0%
	<% end if %>
	</td>
	<td align="center">
		<%= FormatNumber(m_man_naiTot(i),0) %>
		<% if m_naisumtot<>0 then %>
		(<%= CLng(m_man_naiTot(i)/m_naisumtot*100) %> %)
		<% else %>
		(0%)
		<% end if %>
	</td>
	<td align="center">
		<%= FormatNumber(m_woman_naiTot(i),0) %>
		<% if m_naisumtot<>0 then %>
		(<%= CLng(m_woman_naiTot(i)/m_naisumtot*100) %> %)
		<% else %>
		(0%)
		<% end if %>
	</td>
</tr>
<% tmoney = tmoney + m_naiTot(i) %>
<% next %>
</table>
<%= FormatNumber(tmoney,0) %>
<br><br>
<table width="800" cellspacing="1" class="a" bgcolor=#FFFFFF>
<tr bgcolor="#FFFFFF">
	<td colspan="5">
	* 텐바이텐만 검색<br>
	** 취소삭제 검색안함<br>
	*** 결제 완료 이상만 검색<br>
	**** 마이너스 주문 검색안함<br>
	</td>
</tr>
</table>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->