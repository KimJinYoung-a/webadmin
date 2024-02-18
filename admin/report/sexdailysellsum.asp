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
dim fromDate,toDate

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CJumunMaster
oreport.FRectFromDate = fromDate
oreport.FRectToDate = toDate
oreport.MailGuMaeReport

dim ototal
set ototal = new CJumunMaster
ototal.FRectFromDate = fromDate
ototal.FRectToDate = toDate
ototal.MailGuMaeDayTotalReport

dim i,ix,p1,p2,totalsum,totalea
totalsum = 0
totalea = 0
%>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
	남자매출 : <img src="/images/dot1.gif" height="4" width="20" align="absmiddle"><br>
	여자매출 : <img src="/images/dot2.gif" height="4" width="20" align="absmiddle">
	</td>
</tr>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">기간</font></td>
          <td class="a" width="600"><font color="#FFFFFF"></font></td>
          <td class="a" width="120"><font color="#FFFFFF">내용</font></td>
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%
			if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*100)
			end if

			if oreport.maxc<>0 then
				p2 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
			end if
				'총합계금액구하기
				for ix=0 to ototal.FResultCount-1
					if oreport.FMasterItemList(i).Fsitename = ototal.FMasterItemList(ix).FDate then
						totalsum = ototal.FMasterItemList(ix).FDayselltotal
						totalea = ototal.FMasterItemList(ix).FDaysellcnt
					end if
				next
		%>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
		  <td width="120" height="10" align="center">
          	<%= oreport.FMasterItemList(i).Fsitename %>(<%= oreport.FMasterItemList(i).GetDpartName %>)
          </td>
          <td  height="10" width="600">
			<div align="left">
			<% if oreport.FMasterItemList(i).Fsex = 1 then %>
			<img src="/images/dot1.gif" height="4" width="<%= p1 %>%">
			<% else %>
			<img src="/images/dot2.gif" height="4" width="<%= p1 %>%">
			<% end if %>
			</div><br>
          	<div align="left"> <img src="/images/dot4.gif" height="4" width="<%= p2 %>%"></div>
          </td>
		  <td class="a" width="160" align="right">
			<% if oreport.FMasterItemList(i).Fsex = 1 then %>
			<%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %>원 <font color="#0080FF">(<% = round(oreport.FMasterItemList(i).Fselltotal / totalsum,3) * 100 %>%)</font><br>
			<% else %>
			<%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %>원 <font color="#FF6666">(<% = round(oreport.FMasterItemList(i).Fselltotal / totalsum,3) * 100 %>%)</font><br>
          	<% end if %>
			<%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>건 <font color="#808080">(<% = round(oreport.FMasterItemList(i).Fsellcnt / totalea,3) * 100 %>%)</font>
		  </td>
        </tr>
        <% next %>
</table>
<%
set oreport = Nothing
set ototal = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->