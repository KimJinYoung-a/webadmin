<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/academy_reportcls.asp"-->

<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim ck_joinmall,ck_ipjummall,ck_pointmall,research
dim oldlist
dim includematrial

includematrial = RequestCheckvar(request("includematrial"),1)

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)

research = RequestCheckvar(request("research"),2)
ck_joinmall = RequestCheckvar(request("ck_joinmall"),2)
ck_ipjummall = RequestCheckvar(request("ck_ipjummall"),2)
ck_pointmall = RequestCheckvar(request("ck_pointmall"),2)
oldlist = RequestCheckvar(request("oldlist"),1)

if research<>"on" then
	if ck_joinmall="" then ck_joinmall="on"
	if ck_ipjummall="" then ck_ipjummall="on"
end if

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)



'==============================================================================
dim oreport

set oreport = new CJumunMaster

oreport.FRectFromDate = fromDate
oreport.FRectToDate = toDate
oreport.FRectOldJumun = oldlist
oreport.FRectIncludeMatrial = includematrial

oreport.SearchSellReportByRegdate



'==============================================================================
dim oreportByLecday

set oreportByLecday = new CJumunMaster

oreportByLecday.FRectFromDate = fromDate
oreportByLecday.FRectToDate = toDate
oreportByLecday.FRectOldJumun = oldlist
oreportByLecday.FRectIncludeMatrial = includematrial

oreportByLecday.SearchMallSellrePort5_1



'==============================================================================
dim oreportByLecdaySum

set oreportByLecdaySum = new CJumunMaster

oreportByLecdaySum.FRectFromDate = fromDate
oreportByLecdaySum.FRectToDate = toDate
oreportByLecdaySum.FRectOldJumun = oldlist
oreportByLecdaySum.FRectIncludeMatrial = includematrial

oreportByLecdaySum.SearchMallSellrePort5_2


dim i,p1,p2
dim plussum,pluscount
dim miletotal,coupontotal

dim j, dashflag
dashflag = false
%>


<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %> (주문일)
		<input type="radio" name="includematrial" value=""  <% if (includematrial = "") then %>checked<% end if %>>전체
		<input type="radio" name="includematrial" value="N" <% if (includematrial = "N") then %>checked<% end if %>>강좌료만
		<input type="radio" name="includematrial" value="M" <% if (includematrial = "M") then %>checked<% end if %>>재료비만
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" class="a">
<tr>
	<td>결제금액 = 구매금액 합계 - 쿠폰/ 마일리지사용 </td>
</tr>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">기간</font></td>
          <td class="a"><font color="#FFFFFF"></font></td>
          <% if (NOT C_InspectorUser) then %>
          <td class="a" width="220">
	          <table width="220" border=0 cellspacing=0 cellpadding=1 class="a">
			  <tr align="center">
			  	<td width="60"><font color="#FFFFFF">강좌월</font></td>
			  	<td width="60"><font color="#FFFFFF">구매금액</font></td>
			  	<td width="40"><font color="#FFFFFF">결제<br>건수</font></td>
			  	<td width="60"><font color="#FFFFFF">객단가</font></td>
			  </tr>
			  </table>
          </td>
          
          <td class="a" width="70"><font color="#FFFFFF">마일리지<br>/ 쿠폰 사용</font></td>
          <% end if %>
          <td class="a" width="80"><font color="#FFFFFF"> 결제금액(원)</font></td>
          <td class="a" width="50"><font color="#FFFFFF">결제건수</font></td>
         
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%
			if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*100)
			end if

			if oreport.maxc<>0 then
				p2 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
			end if

			plussum		=	plussum + oreport.FMasterItemList(i).Fselltotal
			pluscount	=	pluscount + oreport.FMasterItemList(i).Fsellcnt

			miletotal	=	miletotal + oreport.FMasterItemList(i).Fmiletotal
			coupontotal	=	coupontotal + oreport.FMasterItemList(i).Fcoupontotal
		%>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
		<td width="120" height="10">
          	<%= oreport.FMasterItemList(i).FYYYYMMDD %>(<%= oreport.FMasterItemList(i).GetDpartName %>)
          	</td>
          	<td  height="35">
			<div align="left"> <img src="/images/dot1.gif" height="3" width="<%= p1 %>%"></div><br>
          		<div align="left"> <img src="/images/dot2.gif" height="3" width="<%= p2 %>%"></div>
        </td>
        <% if (NOT C_InspectorUser) then %>
		<td align="right" width="220" >
			<table width="220" border=0 cellspacing=0 cellpadding=1 class="a">

			<% for j=0 to oreportByLecday.FResultcount-1 %>
				<% if oreport.FMasterItemList(i).FYYYYMMDD=oreportByLecday.FMasterItemList(j).FYYYYMMDD then %>
				<% if (dashflag) then %>
				<tr height="1" bgcolor="#EFBE00">
					<td colspan="4"></td>
				</tr>
				<% end if %>
				<tr>
					<td width="60" align="center"><%= oreportByLecday.FMasterItemList(j).FLecYYYYMM %></td>
					<td width="60" align="right"><%= FormatNumber(oreportByLecday.FMasterItemList(j).Fselltotal,0) %></td>
					<td width="40" align="center"><%= oreportByLecday.FMasterItemList(j).Fsellcnt %></td>
					<td width="60" align="center">
					<% if oreportByLecday.FMasterItemList(j).Fsellcnt<>0 then %>
					<%= FormatNumber(CLng(oreportByLecday.FMasterItemList(j).Fselltotal/oreportByLecday.FMasterItemList(j).Fsellcnt),0) %>
					<% end if %>
					</td>
				</tr>
				<% dashflag = true %>
				<% end if %>
			<% next %>

			<% dashflag = false %>
			</table>
		</td>
		
		<td width="70" align="right">
			<%= FormatNumber(oreport.FMasterItemList(i).Fmiletotal*-1,0) %><br>
			<%= FormatNumber(oreport.FMasterItemList(i).Fcoupontotal*-1,0) %>
		</td>
		<% end if %>
		<td align="right">
			<%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %>
		</td>
		<td width="50" align="right">
			<%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>
		</td>
	    
        </tr>
        <% next %>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
        	<td>Total</td>
        	<td></td>
        	<% if (NOT C_InspectorUser) then %>
        	<td align="right" width="220" >
				<table width="220" border=0 cellspacing=0 cellpadding=1 class="a">

				<% for j=0 to oreportByLecdaySum.FResultcount-1 %>
					<% if (dashflag) then %>
					<tr height="1" bgcolor="#EFBE00">
						<td colspan="4"></td>
					</tr>
					<% end if %>
					<tr>
						<td width="60" align="center"><%= oreportByLecdaySum.FMasterItemList(j).FLecYYYYMM %></td>
						<td width="60" align="right"><%= FormatNumber(oreportByLecdaySum.FMasterItemList(j).Fselltotal,0) %></td>
						<td width="40" align="center"><%= oreportByLecdaySum.FMasterItemList(j).Fsellcnt %></td>
						<td width="60" align="center">
						<% if oreportByLecdaySum.FMasterItemList(j).Fsellcnt<>0 then %>
						<%= FormatNumber(CLng(oreportByLecdaySum.FMasterItemList(j).Fselltotal/oreportByLecdaySum.FMasterItemList(j).Fsellcnt),0) %>
						<% end if %>
						</td>
					</tr>
					<% dashflag = true %>
				<% next %>

				</table>
			</td>
			
        	<td align="right">
        		<%= Formatnumber(miletotal*-1,0) %><br>
        		<%= Formatnumber(coupontotal*-1,0) %>
        	</td>
            <% end if %>
        	<td align="right"><%= Formatnumber(plussum,0) %></td>
        	<td align="right"><%= Formatnumber(pluscount,0) %></td>
           
        </tr>
</table>
<%
set oreport = Nothing
set oreportByLecday = Nothing
set oreportByLecdaySum = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->