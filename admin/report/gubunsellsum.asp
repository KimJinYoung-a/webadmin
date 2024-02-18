<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
const Maxlines = 10

dim totalpage, totalnum, q, i
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim ojumun
dim fromDate,toDate,jnx,tmpStr,siteId
dim searchId

dim ck_noextsite,ck_nopoint

ck_noextsite= request("ck_noextsite")
ck_nopoint  = request("ck_nopoint")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")


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

ojumun.FRectSearchType = 1
ojumun.SearchMallSellrePort3

dim arr_site(4), arr_totalmoney(4), arr_totalcount(4), arr_percentmoney(4), arr_percentcount(4)
arr_totalmoney(0) = 0
arr_totalcount(0) = 0
arr_site(0) = "전체"
arr_site(1) = "10x10"
arr_site(2) = "제휴몰"
arr_site(3) = "외부몰"
arr_site(4) = "포인트결제"

arr_totalmoney(1) = ojumun.FMtotalmoney
arr_totalcount(1) = ojumun.FMtotalsellcnt

ojumun.FRectSearchType = 2
ojumun.SearchMallSellrePort3

arr_totalmoney(2) = ojumun.FMtotalmoney
arr_totalcount(2) = ojumun.FMtotalsellcnt

ojumun.FRectSearchType = 3
ojumun.SearchMallSellrePort3

arr_totalmoney(3) = ojumun.FMtotalmoney
arr_totalcount(3) = ojumun.FMtotalsellcnt

ojumun.FRectSearchType = 4
ojumun.SearchMallSellrePort3

arr_totalmoney(4) = ojumun.FMtotalmoney
arr_totalcount(4) = ojumun.FMtotalsellcnt

arr_totalmoney(0) = arr_totalmoney(1) + arr_totalmoney(2) + arr_totalmoney(3) + arr_totalmoney(4)
arr_totalcount(0) = arr_totalcount(1) + arr_totalcount(2) + arr_totalcount(3) + arr_totalcount(4)

if arr_totalmoney(0)<>0 then
	arr_percentmoney(0) = 100
	arr_percentmoney(1) = CLng(arr_totalmoney(1)/arr_totalmoney(0)*100)
	arr_percentmoney(2) = CLng(arr_totalmoney(2)/arr_totalmoney(0)*100)
	arr_percentmoney(3) = CLng(arr_totalmoney(3)/arr_totalmoney(0)*100)
	arr_percentmoney(4) = CLng(arr_totalmoney(4)/arr_totalmoney(0)*100)
end if

if arr_totalcount(0)<>0 then
	arr_percentcount(0) = 100
	arr_percentcount(1) = CLng(arr_totalcount(1)/arr_totalcount(0)*100)
	arr_percentcount(2) = CLng(arr_totalcount(2)/arr_totalcount(0)*100)
	arr_percentcount(3) = CLng(arr_totalcount(3)/arr_totalcount(0)*100)
	arr_percentcount(4) = CLng(arr_totalcount(4)/arr_totalcount(0)*100)
end if
%>
<script language='javascript'>
function Check(){
   if(document.frm.elements[1].checked == true){
       document.frm.ckipkumdiv4.value="on";
   }
   document.frm.submit();
}
</script>


<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		(마이너스주문은 포함하지 않음)
		<td class="a" align="right">
			<a href="javascript:Check();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
	<tr align="center">
        	<td class="a" width="120"><font color="#FFFFFF">사이트명</font></td>
          	<td class="a"><font color="#FFFFFF"></font></td>
          	<td class="a" width="80"><font color="#FFFFFF">금액(원)</font></td>
          	<td class="a" width="50"><font color="#FFFFFF">비율(%)</font></td>
          	<td class="a" width="50"><font color="#FFFFFF">건수</font></td>
          	<td class="a" width="50"><font color="#FFFFFF">비율(%)</font></td>
        </tr>

		<% for i=0 to ubound(arr_site) %>
        <tr bgcolor="#FFFFFF" height="10" class="a">
		<td height="10">
          	<%= arr_site(i) %>
          	</td>
          	<td height="35">
			<div align="left"> <img src="/images/dot1.gif" height="3" width="<%= CLng(arr_percentmoney(i)) %>%"></div><br>
          		<div align="left"> <img src="/images/dot2.gif" height="3" width="<%= CLng(arr_percentcount(i)) %>%"></div>
          	</td>
		<td class="a" align="right">
			<%= FormatNumber(arr_totalmoney(i),0) %>
		</td>
		<td class="a" align="right">
			<%= arr_percentmoney(i) %>
		</td>
		<td class="a" align="right">
          		<%= FormatNumber(arr_totalcount(i),0) %>
		</td>
		<td class="a" align="right">
          		<%= arr_percentcount(i) %>
		</td>
		
        </tr>
        <% next %>
</table>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
