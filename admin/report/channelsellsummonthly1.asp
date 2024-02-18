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

dim ck_joinmall,ck_ipjummall,ck_pointmall,research
dim yyyy1,mm1,sDate,cdL,cdM
dim rectoldjumun
dim Param

yyyy1 = request("yyyy1")
mm1 = request("mm1")
cdL = request("cd1")
cdM = request("cd2")

research = request("research")
ck_joinmall = request("ck_joinmall")
ck_ipjummall = request("ck_ipjummall")
ck_pointmall = request("ck_pointmall")
rectoldjumun = request("rectoldjumun")

if research<>"on" then
	if ck_joinmall="" then ck_joinmall="on"
	if ck_ipjummall="" then ck_ipjummall="on"
	'if ck_pointmall="" then ck_pointmall="on"
end if

if yyyy1="" then
	yyyy1 = LefT(Now(),4)
	mm1 = mid(Now(),6,2)
end if

Param = "&yyyy1="&yyyy1&"&mm1="&mm1&"&research="&research&"&ck_joinmall="&ck_joinmall&"&ck_ipjummall="&ck_ipjummall&"&ck_pointmall="&ck_pointmall&"&rectoldjumun="&rectoldjumun

sDate = yyyy1 & "-" & mm1
dim oReport
set oReport = new CJumunMaster
'oReport.FRectJoinMallNotInclude = ck_joinmall
oReport.FRectExtMallNotInclude = ck_ipjummall
'oReport.FRectPointNotInclude = ck_pointmall
oReport.FRectSearchDate = sDate
oReport.FRectCD1 = cdL
oReport.FRectCD2 = cdM
oReport.FRectOldJumun = rectoldjumun

oReport.SearchMallSellrePortMonthlyChannel

dim i,p1,p2
dim prename
dim buftext, bufname, bufimage
dim sumtotal, counttotal, buytotal
dim ch1,ch2,ch3,ch4,ch5,ch6,ch7,ch8,ch9,ch10,ch11

%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		<!--<input type="checkbox" name="rectoldjumun" <% if rectoldjumun="on" then response.write "checked" %> >6개월이전자료//-->

		<% DrawYMBox yyyy1,mm1 %>
		<!--<input type="checkbox" name="ck_joinmall" <% if ck_joinmall="on" then response.write "checked" %> >제휴몰 포함 //-->
		<!--<input type="checkbox" name="ck_pointmall" <% if ck_pointmall="on" then response.write "checked" %> >포인트몰 포함//-->
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
        	<td class="a" width="120"><font color="#FFFFFF">카테고리</font></td>
        	<td class="a"><font color="#FFFFFF"></font></td>
        	<td class="a" width="50"><font color="#FFFFFF">건수</font></td>
       		<td class="a" width="80"><font color="#FFFFFF">매출액(원)</font></td>
       		<td class="a" width="80"><font color="#FFFFFF">매입액(원)</font></td>
          	<td class="a" width="80"><font color="#FFFFFF">수익율(%)</font></td>
        </tr>

		<% for i=0 to oReport.FResultCount-1 %>
		<%
			if oReport.maxt<>0 then
				p1 = Clng(oReport.FMasterItemList(i).Fselltotal/oReport.maxt*90)
			end if
        	sumtotal = sumtotal + oReport.FMasterItemList(i).Fselltotal
        	buytotal = buytotal + oReport.FMasterItemList(i).Fbuytotal
			counttotal = counttotal + oReport.FMasterItemList(i).Fsellcnt
        	%>

        <tr bgcolor="#FFFFFF" height="10"  class="a">
		<td width="120" height="10">
			<% IF cdL<>""  and cdM<>"" Then %>
				<%= oReport.FMasterItemList(i).FItemgubunNm %>
			<% ElseIF cdL<>"" Then %>
				<a href="?cd1=<%= cdL %>&cd2=<%=oReport.FMasterItemList(i).Fitemgubun&Param %>"><%= oReport.FMasterItemList(i).FItemgubunNm %></a>
			<% Else %>
				<a href="?cd1=<%=oReport.FMasterItemList(i).Fitemgubun&Param %>"><%= oReport.FMasterItemList(i).FItemgubunNm %></a>			
			<% End IF %>
		</td>

		<td>
	        <table border="0" class="a" width='<%= CStr(p1) %>%' >
			  <tr>
			  	<td background='/images/dot<%= CStr((oReport.FMasterItemList(i).Fitemgubun)) %>.gif' height='20' >
			  	<% if oReport.FTotalPrice<>0 then %>
			  	<%= CLng(oReport.FMasterItemList(i).Fselltotal/oReport.FTotalPrice*10000)/100 %>%
			  	<% end if %>
			  	</td>
			  </tr>
			</table>
	    </td>
	    <td class="a" align="right">
			<%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt,0) %>
		</td>
		<td class="a" align="right">
			<%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %>
		</td>
		<td class="a" align="right">
			<%= FormatNumber(oReport.FMasterItemList(i).Fbuytotal,0) %>
		</td>
		<td class="a" align="center">
			<% if oReport.FMasterItemList(i).Fselltotal<>0 then %>
			<%= 100-CLng(oReport.FMasterItemList(i).Fbuytotal/oReport.FMasterItemList(i).Fselltotal*100*100)/100 %> %
			<% end if %>
		</td>
        </tr>
        <% next %>
        <tr bgcolor="#FFFFFF">
        	<td>Total</td>
        	<td ></td>
        <td class="a" align="right">
			<%= FormatNumber(counttotal,0) %>
		</td>
		<td class="a" align="right">
			<%= FormatNumber(sumtotal,0) %>
		</td>
		<td class="a" align="right">
			<%= FormatNumber(buytotal,0) %>
		</td>
		<td class="a" align="center">
			<% if sumtotal<>0 then %>
			<%= 100-CLng(buytotal/sumtotal*100*100)/100 %> %
			<% end if %>
		</td>
        </tr>
</table>
<%
set oReport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->