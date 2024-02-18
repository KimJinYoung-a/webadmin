<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
rw "사용중지메뉴-관리자문의요망"
response.end
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nextdateStr,searchnextdate
dim shopid, rectorder, makerid
dim offgubun
dim oldlist

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
shopid = request("shopid")
rectorder = request("rectorder")
makerid = request("makerid")
offgubun = request("offgubun")
oldlist = request("oldlist")


if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)
if rectorder="" then rectorder="bysum"

dim ooffsell,i,ix

if shopid<>"" then offgubun=""

set ooffsell = new COffShopSellReport
ooffsell.FRectStartDay = yyyy1 + "-" + mm1 + "-" + dd1
ooffsell.FRectEndDay = searchnextdate
ooffsell.FRectShopID = "cafe002"
ooffsell.FPageSize = 1000
ooffsell.FRectOrder = rectorder
ooffsell.FRectOffgubun = "CAF"
ooffsell.FRectOldData = oldlist

ooffsell.ShopJumunListBybestseller

%>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
      <input type="hidden" name="showtype" value="showtype">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3개월이전내역
		&nbsp;
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<br>정렬기준 :
		<input type="radio" name="rectorder" value="bysum" <% if rectorder="bysum" then response.write "checked" %> > 매출액
		<input type="radio" name="rectorder" value="bycnt" <% if rectorder="bycnt" then response.write "checked" %> > 매출건수

		</td>
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
	<td colspan="3">(매출액은 상품원 판매가 기준입니다.)</td>
	<td colspan="7" height="25" align="right">검색결과 : 총 <font color="red"><% = ooffsell.FResultCount %></font>개 (최대 <%= ooffsell.FPageSize %> 건 검색)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="90" align="center">상품번호</td>
	<td  align="center">상품</td>
	<td width="80" align="center">옵션</td>
	<td width="100" align="center">브랜드</td>
	<td width="80" align="center">가격</td>
	<td width="65" align="center">매출건수</td>
	<td width="65" align="center">%</td>
	<td width="80" align="center">매출액</td>
	<td width="65" align="center">%</td>
</tr>
<% if ooffsell.FResultCount<1 then %>
<tr  bgcolor="#FFFFFF">
	<td colspan="12" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to ooffsell.FResultCount -1 %>
<%
Dim sumprice,totalsumprice
sumprice = ooffsell.FItemList(ix).FItemCost * ooffsell.FItemList(ix).FItemNo
%>
	<tr class="a"  bgcolor="#FFFFFF">
		<td align="center" height="25"><%= ooffsell.FItemList(ix).FItemGubun %>-<%= Format00(6,ooffsell.FItemList(ix).FItemID)  %>-<%= ooffsell.FItemList(ix).FItemOption %></td>
		<td align="left"><%= ooffsell.FItemList(ix).FItemName %></td>
		<% if (ooffsell.FItemList(ix).FItemOptionStr="") then %>
			<td align="center">&nbsp;</td>
		<% else %>
			<td align="center"><%= ooffsell.FItemList(ix).FItemOptionStr %></td>
		<% end if %>
		<td ><%= ooffsell.FItemList(ix).FMakerid %></td>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(ix).FItemCost,0)  %></td>
		<td align="center"><%= ooffsell.FItemList(ix).FItemNo %></td>
		<td align="center">
		<% if ooffsell.maxc<>0 then %>
			<%= Clng(ooffsell.FItemList(ix).FItemNo/ooffsell.maxc*10*100)/10 %> %
		<% end if %>
		</td>
		<td align="right"><%= FormatNumber(sumprice,0) %></td>
		<td align="center">
		<% if ooffsell.maxt<>0 then %>
			<%= Clng(ooffsell.FItemList(ix).FItemNo*ooffsell.FItemList(ix).FItemCost/ooffsell.maxt*10*100)/10 %> %
		<% end if %>
		</td>
	</tr>
	 <% totalsumprice =  totalsumprice + sumprice %>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="10" height="25" align="right">현재 페이지 합계 금액 : <font color="red"><% =FormatNumber(totalsumprice,0) %></font>원&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	</tr>
<% end if %>
</table>
<%
set ooffsell = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->