<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/diysell_reportcls.asp"-->

<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim research, oldlist
dim period

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)
period = RequestCheckvar(request("period"),16)

research = RequestCheckvar(request("research"),2)

if period="" then period="day"

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CDiyReportMaster
oreport.FRectFromDate = fromDate
oreport.FRectToDate = toDate

if (period="month") then
	oreport.GetDiyMonthlyReport
else
	oreport.GetDiyDailyReport
end if

dim i,p1,p2


Dim SumTtl, SumBuy, SumCnt
%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %> (주문일)
		&nbsp;&nbsp;
		<input type="radio" name="period" value="day" <% if period="day" then response.write "checked" %> > 일별
		<input type="radio" name="period" value="month" <% if period="month" then response.write "checked" %> > 월별
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">기간</font></td>
          <td class="a"><font color="#FFFFFF"></font></td>
          <td class="a" width="100"><font color="#FFFFFF">구매금액</font></td>
          <td class="a" width="100"><font color="#FFFFFF">객단가<br>(구매금액)</font></td>
          <td class="a" width="100"><font color="#FFFFFF">결재금액</font></td>
          <td class="a" width="50"><font color="#FFFFFF">결재건수</font></td>
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%

            SumTtl = SumTtl + oreport.FItemList(i).Fselltotal
            SumCnt = SumCnt + oreport.FItemList(i).Fsellcnt

			if oreport.maxt<>0 then
				p1 = Clng(oreport.FItemList(i).Fselltotal/oreport.maxt*100)
			end if

			if oreport.maxc<>0 then
				p2 = Clng(oreport.FItemList(i).Fsellcnt/oreport.maxc*100)
			end if
		%>
				<tr bgcolor="#FFFFFF" height="35" class="a">
					<td width="120">
						<%= oreport.FItemList(i).Fyyyymmdd %>
          </td>
          <td>
						<div align="left"> <img src="/images/dot1.gif" height="3" width="<%= p1 %>%"></div><br>
          	<div align="left"> <img src="/images/dot2.gif" height="3" width="<%= p2 %>%"></div>
         	</td>
			<td class="a" width="100" align="right">
				<%= FormatNumber(oreport.FItemList(i).Forgtotal,0) %><br>
			</td>
			<td class="a" width="100" align="right">
				<%= FormatNumber(oreport.FItemList(i).Fsellavg,0) %><br>
			</td>
			<td class="a" width="100" align="right">
				<%= FormatNumber(oreport.FItemList(i).Fselltotal,0) %><br>
			</td>
					<td class="a" width="50" align="right">
				   	<%= FormatNumber(oreport.FItemList(i).Fsellcnt,0) %>
					</td>

        </tr>
        <% next %>
        <tr bgcolor="#FFFFFF" height="35" class="a">
        <td>총계</td>
        <td></td>
        <td></td>
        <td></td>
        <td align="right"><%= FormatNumber(SumTtl,0) %></td>
        <td align="right"><%= FormatNumber(SumCnt,0) %></td>
        </tr>
</table>
<%
set oreport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->