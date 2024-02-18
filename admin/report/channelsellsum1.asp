<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls1.asp"-->

<%
'### 2011.01.14
'### 일단 작업이 완료되어 변경됨. 기존의 소스에 1을 붙인 파일. 총 3개 수정.
'### /admin/report/channelsellsum1.asp, /admin/report/channelsellsummonthly1.asp, /lib/classes/report/reportcls1.asp

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2,Param
dim fromDate,toDate,cdL,cdM
dim ck_joinmall,ck_ipjummall,ck_pointmall,research
dim rectoldjumun,dategubun

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
rectoldjumun = request("rectoldjumun")
dategubun = request("dategubun")
research = request("research")
ck_joinmall = request("ck_joinmall")
ck_ipjummall = request("ck_ipjummall")
ck_pointmall = request("ck_pointmall")

cdL = request("cd1")
cdM = request("cd2")

if research<>"on" then
	if ck_joinmall="" then ck_joinmall="on"
	if ck_ipjummall="" then ck_ipjummall="on"
	if dategubun="" then dategubun="D"
end if

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2)


Param = "&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&dategubun="&dategubun&"&research="&research&"&ck_joinmall="&ck_joinmall&"&ck_ipjummall="&ck_ipjummall&"&ck_pointmall="&ck_pointmall&"&rectoldjumun="&rectoldjumun


dim oReport
set oReport = new CJumunMaster
oReport.FRectFromDate = fromDate
oReport.FRectToDate = toDate
oReport.FRectToDateGubun = dategubun

oReport.FRectCD1 = cdL
oReport.FRectCD2 = cdM
'oReport.FRectJoinMallNotInclude = ck_joinmall
oReport.FRectExtMallNotInclude = ck_ipjummall
'oReport.FRectPointNotInclude = ck_pointmall
oReport.FRectOldJumun = rectoldjumun

oReport.SearchMallSellrePortChannel
'oreport.SearchMallSellrePortMonthlyChannel
dim i,p1,p2
dim prename, nextname
dim buftext, bufname, bufimage
dim sumtotal
dim ch1,ch2,ch3,ch4,ch5,ch6,ch7,ch8,ch9,ch10,ch11


dim sellcnt, selltotal, buytotal
dim TTLsellcnt, TTLselltotal, TTLbuytotal
%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="cd1" value="<%= cdL %>">
	<input type="hidden" name="cd2" value="<%= cdM %>">
	<tr>
		<td class="a" >
		<!--<input type="checkbox" name="rectoldjumun" <% if rectoldjumun="on" then response.write "checked" %> >6개월이전자료&nbsp;&nbsp;//-->
		<input type="radio" name="dategubun" value="D" <% If dategubun<>"M" Then response.write "checked" %>>일별 <input type="radio" name="dategubun" value="M" <% If dategubun="M" Then response.write "checked" %>>월별
			
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<!--<input type="checkbox" name="ck_joinmall" <% if ck_joinmall="on" then response.write "checked" %> >제휴몰 포함//-->
		<input type="checkbox" name="ck_ipjummall" <% if ck_ipjummall="on" then response.write "checked" %> >입점몰 포함
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" class="a" >
<tr>
	<td>* 검색되어 나오는 카테고리는 <b>판매당시의 카테고리</b> 입니다.<br>* 실시간 데이터와는 <b>약 1시간 내외의 차이</b>가 있습니다.<br>* 마이너스금액, 쿠폰, 산정하지 않음.</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00" class="a" >
    <tr align="center">
      <td width="90" class="a"><font color="#FFFFFF">기간</font></td>
      <td width="300" ><font color="#FFFFFF">&nbsp;</font></td>
      <td width="60" class="a"><font color="#FFFFFF">건수</font></td>
      <td width="100" class="a"><font color="#FFFFFF">매출액</font></td>
      <td width="100" class="a"><font color="#FFFFFF">매입액</font></td>
      <td width="100" class="a"><font color="#FFFFFF">수익</font></td>
      <td width="60" class="a"><font color="#FFFFFF">수익율</font></td>
    </tr>
	<% for i=0 to oreport.FResultCount-1 %>
	<%
		p1 = 0
		if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*90)
		end if

		sellcnt		=	sellcnt + oreport.FMasterItemList(i).Fsellcnt
		selltotal	=	selltotal + oreport.FMasterItemList(i).Fselltotal
		buytotal	=	buytotal + oreport.FMasterItemList(i).Fbuytotal
	%>
	<tr bgcolor="#FFFFFF">
	<td align="center">
			<% IF cdL<>""  and cdM<>"" Then %>
				<%= oReport.FMasterItemList(i).FItemgubunNm %>
			<% ElseIF cdL<>"" Then %>
				<a href="?cd1=<%= cdL %>&cd2=<%=oReport.FMasterItemList(i).Fitemgubun&Param %>"><%= oReport.FMasterItemList(i).FItemgubunNm %></a>
			<% Else %>
				<a href="?cd1=<%=oReport.FMasterItemList(i).Fitemgubun&Param %>"><%= oReport.FMasterItemList(i).FItemgubunNm %></a>			
			<% End IF %>
			<!--	  	<a href=<%= oreport.FMasterItemList(i).GetChannelName_Kor %>-->
	</td>
	  <td >
			<table border="0" class="a" width='<%= CStr(p1) %>%' >
			  <tr>
			  	<% if trim(oreport.FMasterItemList(i).Fitemgubun)="" then %>
			  	<td height='20' background='/images/dot030.gif'>
			  	<% else %>
			  	<td background='/images/dot<%= "0"&right(oreport.FMasterItemList(i).Fitemgubun,2) %>.gif' height='20' >
			  	<% end if %>
			  	<% if oreport.FTotalPrice<>0 then %>
			  	<%= CLng(oreport.FMasterItemList(i).Fselltotal/oreport.FTotalPrice*10000)/100 %>%
			  	<% end if %>
			  	</td>
			  </tr>
			</table>
	  </td>
	  <td align="right"><%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>건</td>
	  <td align="right"><%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %> 원</td>
	  <td align="right"><%= FormatNumber(oreport.FMasterItemList(i).Fbuytotal,0) %> 원</td>
	  <td align="right"><%= FormatNumber(oreport.FMasterItemList(i).Fselltotal-oreport.FMasterItemList(i).Fbuytotal,0) %> 원</td>
	  <td align="center">
	  <% if oreport.FMasterItemList(i).Fselltotal<>0 then %>
	  	<%= 100-CLng(oreport.FMasterItemList(i).Fbuytotal/oreport.FMasterItemList(i).Fselltotal*100*100)/100 %>%
	  <% end if %>
	  </td>
	</tr>
	<%
	prename = oreport.FMasterItemList(i).Fsitename
	if oreport.FResultCount>i+1 then nextname = oreport.FMasterItemList(i+1).Fsitename else nextname=""
	%>
	<% if (prename<>"") and (prename<>nextname) or (nextname="") then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="6"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td align="center"><%= prename  %></td>
	  <td></td>
	  <td align="right"><%= FormatNumber(sellcnt,0) %>건</td>
	  <td align="right"><%= FormatNumber(selltotal,0) %> 원</td>
	  <td align="right"><%= FormatNumber(buytotal,0) %> 원</td>
	  <td align="right"><%= FormatNumber(selltotal-buytotal,0) %> 원</td>
	  <td align="center">
	   <% if selltotal<>0 then %>
	  	<%= 100-CLng(buytotal/selltotal*100*100)/100 %>%
	  <% end if %>
	  </td>
	</tr>
	<tr bgcolor="#EFBE00">
		<td colspan="6"></td>
	</tr>
	<%
		TTLsellcnt	= TTLsellcnt + sellcnt
		TTLselltotal= TTLselltotal + selltotal
		TTLbuytotal = TTLbuytotal + buytotal

		sellcnt = 0
		selltotal = 0
		buytotal = 0
	%>
	<% end if %>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="6"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td align="center">Total</td>
	  <td></td>
	  <td align="right"><%= FormatNumber(TTLsellcnt,0) %>건</td>
	  <td align="right"><%= FormatNumber(TTLselltotal,0) %> 원</td>
	  <td align="right"><%= FormatNumber(TTLbuytotal,0) %> 원</td>
	  <td align="right"><%= FormatNumber(TTLselltotal-TTLbuytotal,0) %> 원</td>
	  <td align="center">
	  <% if TTLselltotal<>0 then %>
	  	<%= 100-CLng(TTLbuytotal/TTLselltotal*100*100)/100 %>%
	  <% end if %>
	  </td>
	</tr>
	<tr bgcolor="#EFBE00">
		<td colspan="6"></td>
	</tr>
</table>


<%
set oreport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->