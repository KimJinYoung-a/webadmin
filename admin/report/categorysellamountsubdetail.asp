<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/category_reportcls.asp"-->
<%
const Maxlines = 10

dim i
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim oreport
dim fromDate,toDate
dim order_desum,oldlist
dim codelarge, codemid
Dim totalsum,totalper,totalcnt

codelarge = request("codelarge")
codemid = request("codemid")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
order_desum = request("order_desum")
oldlist=request("oldlist")


if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))


fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

set oreport = new CCategoryReport
oreport.FRectFromDate = fromDate
oreport.FRectToDate = toDate
oreport.FRectCD1 = codelarge
oreport.FRectCD2 = codemid
oreport.FRectOldJumun=oldlist
oreport.SearchCategorySellrePortSubDetail

dim selltotal
selltotal = 0

for i=0 to oreport.FResultCount - 1
	if not IsNULL(oreport.FItemList(i).Fselltotal) then
		selltotal = selltotal + oreport.FItemList(i).Fselltotal
	end if
next
%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
      <input type="hidden" name="showtype" value="showtype">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="codelarge" value="<%= codelarge %>">
	<input type="hidden" name="codemid" value="<%= codemid %>">
	<tr>
		<td class="a" ><input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역<br>
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
		  <td width="140" class="a"><font color="#FFFFFF">카테고리</font></td>
          <td class="a" width="500"></td>
          <td class="a" width="100"><font color="#FFFFFF">매출</font></td>
          <td class="a" width="50"><font color="#FFFFFF">%</font></td>
        </tr>
		<% if oreport.FresultCount<1 then %>
		<tr bgcolor="#FFFFFF">
			<td colspan="4" align="center"  class="a">[검색결과가 없습니다.]</td>
		</tr>
		<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="4" align="center"  class="a"><%= oreport.FRectFromDate & " ~ " & oreport.FRectToDate%></td>
		</tr>
        <% for i=0 to oreport.FResultCount - 1 %>
        <tr bgcolor="#FFFFFF">
		  <td width="140" height="10" class="a">
          	 <a href="categorysellsubbestseller.asp?menupos=221&cdl=<%= codelarge %>&cdm=<%= codemid %>&cds=<%= oreport.FItemList(i).FCsmall %>&yyyy1=<%= yyyy1 %>&mm1=<%= Format00(2,mm1) %>&dd1=<%= Format00(2,dd1) %>&yyyy2=<%= yyyy2 %>&mm2=<%= Format00(2,mm2) %>&dd2=<%= Format00(2,dd2) %>&oldlist=<%= oldlist %>"><%= oreport.FItemList(i).FCSName %></a> (<%= oreport.FItemList(i).FCsmall %>)
          </td>
          <td  height="10" width="500">
          <% if Not (IsNull(oreport.FItemList(i).Fselltotal)) then %>
          	<div align="left"> <img src="/images/dot1.gif" height="4" width="<%= CLng((oreport.FItemList(i).Fselltotal/oreport.maxt)*500) %>"></div><br>
          	<div align="left"> <img src="/images/dot2.gif" height="4" width="<%= CLng((oreport.FItemList(i).Fsellcnt/oreport.maxc)*500) %>"></div>
          <% end if %>
          </td>
		  <td class="a" width="100">
		  <% if Not (IsNull(oreport.FItemList(i).Fselltotal)) then %>
		  	<div align="right"> <%= FormatNumber(FormatCurrency(oreport.FItemList(i).Fselltotal),0) & "원"%> </div>
		  	<div align="right"> <%= oreport.FItemList(i).Fsellcnt & "건"%> </div>
		  <% end if %>
		  </td>
		  <td class="a" width="50">
		  <% if (oreport.maxt<>0) and Not (IsNull(oreport.FItemList(i).Fselltotal)) then %>
		  <%= CLng(oreport.FItemList(i).Fselltotal/selltotal * 100 *100)/100 %> %
		  <% end if %>
		  </td>
        </tr>
		  <% totalsum = totalsum + oreport.FItemList(i).Fselltotal %>
		  <% totalcnt = totalcnt + oreport.FItemList(i).Fsellcnt %>
		  <% totalper = totalper + CLng(oreport.FItemList(i).Fselltotal/selltotal * 100 *100)/100 %>
        <% next %>
		<% end if %>
        <tr align="center" bgcolor="#FFFFFF">
			 <td width="140" class="a">총계</td>
          <td class="a" width="500"></td>
          <td class="a" width="100" align="right"><% = FormatNumber(totalsum,0) %>원<br><% = FormatNumber(totalcnt,0) %>건</td>
          <td class="a" width="50" align="left"><% = totalper %>%</td>
        </tr>
</table>
<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
</body>
</html>
