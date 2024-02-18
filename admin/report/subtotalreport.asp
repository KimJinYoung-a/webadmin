<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 구매금액별건수
' Hieditor : 최초생성자모름
'			 2019.09.11 한용민 수정(파라메타길이체크)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2, fromDate,toDate, i,p1,p2,pro,pro2, totcnt, totsum
dim ck_joinmall,ck_ipjummall,research, ck_tendeliverExists, oldlist
	yyyy1 = RequestCheckVar(request("yyyy1"),4)
	mm1 = RequestCheckVar(request("mm1"),2)
	dd1 = RequestCheckVar(request("dd1"),2)
	yyyy2 = RequestCheckVar(request("yyyy2"),4)
	mm2 = RequestCheckVar(request("mm2"),2)
	dd2 = RequestCheckVar(request("dd2"),2)
	research = RequestCheckVar(request("research"),2)
	ck_joinmall = RequestCheckVar(request("ck_joinmall"),2)
	ck_ipjummall = RequestCheckVar(request("ck_ipjummall"),2)
	ck_tendeliverExists = RequestCheckVar(request("ck_tendeliverExists"),2)
	oldlist = RequestCheckVar(request("oldlist"),2)

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

dim oreport
set oreport = new CJumunMaster
	oreport.FRectFromDate = fromDate
	oreport.FRectToDate = toDate
	oreport.FRectJoinMallNotInclude = ck_joinmall
	oreport.FRectExtMallNotInclude = ck_ipjummall
	oreport.FRectOldJumun = oldlist
	oreport.FRectTenDeliverExists = ck_tendeliverExists
	oreport.SearchMallSellrePort6

%>

<script type="text/javascript">

function ReSearch(){
	frm.submit();
}

function downloadexcel(){
	document.frm.target = "view";
	document.frm.action = "/admin/report/subtotalreport_excel.asp";
	document.frm.submit();
	document.frm.target = "";
	document.frm.action = "";
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="ReSearch();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="ck_joinmall" <% if ck_joinmall="on" then response.write "checked" %> >제휴몰 포함
		<input type="checkbox" name="ck_ipjummall" <% if ck_ipjummall="on" then response.write "checked" %> >입점몰 포함
		<input type="checkbox" name="ck_tendeliverExists" <% if ck_tendeliverExists="on" then response.write "checked" %> >텐바이텐 배송포함
	</td>
</tr>
</table>
</form>

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<% if oreport.FResultCount > 0 then %>
			<input type="button" onclick="downloadexcel();" value="엑셀다운로드" class="button">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="120">기간</td>
	<td width="600"</td>
	<td width="120">내용</td>
</tr>
<% if oreport.FResultCount > 0 then %>
<% for i=0 to oreport.FResultCount-1 %>
<%
	pro = 0
	if oreport.maxc<>0 then
		p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*100)
		p2 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
		if oreport.FTotalsellcnt<>0 then
			pro = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.FTotalsellcnt*100)
		end if

		if oreport.Ftotalmoney<>0 then
			pro2 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.Ftotalmoney*100)
		end if
	end if
	totcnt = totcnt + oreport.FMasterItemList(i).Fsellcnt
	totsum = totsum + oreport.FMasterItemList(i).Fselltotal
%>
<tr bgcolor="#FFFFFF">
	<td width="120" height="10">
		<%= oreport.FMasterItemList(i).Fsitename %>
	</td>
	<td  height="10" width="600">
		<div align="left"> <img src="/images/dot2.gif" height="4" width="<%= p2 %>%"></div><br>
		<div align="left"> <img src="/images/dot1.gif" height="4" width="<%= p1 %>%"></div>
	</td>
	<td class="a" width="160" align="right">
		<%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>건(<%= pro %>%)<br>
		<%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %>원(<%= pro2 %>%)
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td class="a" colspan="3" align="right">
		총건수 : <%= FormatNumber(totcnt,0) %>
		총금액 : <%= FormatNumber(totsum,0) %>
		객단가 :
		<% if totcnt<>0 then %>
		<%= FormatNumber(CLng(totsum/totcnt),0) %>
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width=100% height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>
<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
