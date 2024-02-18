<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim page,shopid
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim oldlist

shopid = request("shopid")
page = request("page")
if page="" then page=1


shopid = "cafe002"


yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
oldlist = request("oldlist")

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-14)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)


dim ooffsell
set ooffsell = new COffShopSellReport
ooffsell.FRectShopID = shopid
ooffsell.FPageSize=20
ooffsell.FCurrPage=page
ooffsell.FRectNormalOnly = "on"
ooffsell.FRectStartDay = fromDate
ooffsell.FRectEndDay = toDate
ooffsell.FRectOldData = oldlist

ooffsell.GetDaylySumList3TimeBojung

dim i
%>

<table width="900" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
			매출일 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >과거내역조회(2008년1월1일이전)
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="900" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF" class="a" >
<tr>
	<td>
	* 야간매출(새벽 <font color=red>5시</font>까지)는 전날 날짜 매출로 표시됩니다.(일별매출통계만 적용)
	</td>
</tr>
</table>
<table width="900" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align="center">
	<td width="80">샾구분</td>
	<td width="80">매출일</td>
	<td width="80">매출건수</td>
	<td width="80">총금액</td>
	<td width="80">결제액</td>
	<td width="80">마일리지사용</td>
	<td width="60">현금</td>
	<td width="60">카드</td>
	<td width="60">기타</td>
	<td width="60">아이템목록</td>
	<td width="60">주문별목록</td>
</tr>
<% for i=0 to ooffsell.FresultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><%= ooffsell.FItemList(i).FShopName %></td>
	<td><%= ooffsell.FItemList(i).FTerm %></td>
	<td align="center"><%= ooffsell.FItemList(i).FCount %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSum+ooffsell.FItemList(i).FSpendMile,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSum,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSpendMile,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FCashSum,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FCardSum,0) %></td>
	<td align="right">
	<% if (ooffsell.FItemList(i).FSum<>ooffsell.FItemList(i).FCashSum+ooffsell.FItemList(i).FCardSum+ooffsell.FItemList(i).FGiftCardPaysum) then %>
	   <font color="#CCCCCC">(<%= ooffsell.FItemList(i).FSum-(ooffsell.FItemList(i).FCashSum+ooffsell.FItemList(i).FCardSum+ooffsell.FItemList(i).FGiftCardPaysum) %>)</font>
	<% end if %>
	<%= FormatNumber(ooffsell.FItemList(i).FGiftCardPaysum,0) %></td>
	<td align="center"><a href="todayselldetail.asp?menupos=<%= menupos %>&terms=<%= ooffsell.FItemList(i).FTerm %>&shopid=<%= ooffsell.FItemList(i).FShopid %>&oldlist=<%=oldlist%>">보기</a></td>
	<td align="center"><a href="todaysellmaster.asp?menupos=<%= menupos %>&terms=<%= ooffsell.FItemList(i).FTerm %>&shopid=<%= ooffsell.FItemList(i).FShopid %>&oldlist=<%=oldlist%>">보기</a></td>
</tr>
<% next %>

</table>
<%
set ooffsell= Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->