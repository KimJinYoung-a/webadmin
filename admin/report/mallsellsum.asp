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
dim ojumun ,oldOrder
dim fromDate,toDate,jnx,tmpStr,siteId
dim searchId

dim ck_noextsite,ck_nopoint,ck_minus

ck_noextsite= request("ck_noextsite")
ck_nopoint  = request("ck_nopoint")
ck_minus  = request("ck_minus")
oldOrder = request("oldOrder")

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
ojumun.FRectExtMallNotInclude = ck_noextsite
ojumun.FRectPointNotInclude   = ck_nopoint
ojumun.FRectMinusNotInclude   = ck_minus
ojumun.FRectOldJumun = oldOrder
ojumun.SearchMallSellrePort2

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
		<input type="checkbox" name="ck_noextsite" <% if ck_noextsite="on" then response.write "checked" %> >외부몰포함안함
		<input type="checkbox" name="ck_nopoint" <% if ck_nopoint="on" then response.write "checked" %> >포인트결제포함안함
		<input type="checkbox" name="ck_minus" <% if ck_minus="on" then response.write "checked" %> >취소주문포함안함
		<input type="checkbox" name="oldOrder" <% if oldOrder="on" then response.write "checked" %> >6개월이전 주문
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
          	<td class="a" width="50"><font color="#FFFFFF">건수</font></td>
        </tr>
	<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="4" align="center"  class="a">[검색결과가 없습니다.]</td>
	</tr>
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="4" align="center"  class="a">
			총금액:&nbsp;<%= FormatNumber(ojumun.FMtotalmoney,0) %>
			&nbsp;&nbsp;&nbsp;&nbsp;
			총건수:&nbsp;<%= FormatNumber(ojumun.FMtotalsellcnt,0) %>
		</td>
	</tr>
	<% for i=0 to ojumun.FResultCount - 1 %>
        <tr bgcolor="#FFFFFF" height="10" class="a">
		<td width="120" height="10">
          		<%= ojumun.FMasterItemList(i).Fsitename %>
          	</td>
          	<td  height="35">
          		<% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
          		<div align="left"> <img src="/images/dot1.gif" height="3" width="<%= CLng((ojumun.FMasterItemList(i).Fselltotal/ojumun.maxt)*600) %>"></div><br>
          		<div align="left"> <img src="/images/dot2.gif" height="3" width="<%= CLng((ojumun.FMasterItemList(i).Fsellcnt/ojumun.maxc)*600) %>"></div>
          		<% end if %>
          	</td>
		<td class="a">
			<% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
			<div align="right"> <%= FormatNumber(FormatCurrency(ojumun.FMasterItemList(i).Fselltotal),0) %> </div>
		  	<% end if %>
		</td>
		<td>
		  	<% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
		  	<div align="right"> <%= ojumun.FMasterItemList(i).Fsellcnt %> </div>
		  	<% end if %>
		</td>
        </tr>
        <% next %>
	<% end if %>
</table>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
</body>
</html>
