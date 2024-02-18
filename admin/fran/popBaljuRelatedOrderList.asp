<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/order/baljuofflinecls.asp"-->
<%

dim itemgubun, itemid, itemoption
dim makerid

itemgubun 	= requestCheckVar(request("itemgubun"), 32)
itemid 		= requestCheckVar(request("itemid"), 32)
itemoption 	= requestCheckVar(request("itemoption"), 32)

dim oshopjumun, oupchejumun
set oshopjumun = new CTenBaljuOffline

oshopjumun.FRectItemGubun = itemgubun
oshopjumun.FRectItemId = itemid
oshopjumun.FRectItemOption = itemoption

oshopjumun.GetShopOrderList()


if oshopjumun.FResultCount > 0 then
	makerid = oshopjumun.FItemList(0).Fmakerid
end if

set oupchejumun = new CTenBaljuOffline

oupchejumun.FRectItemGubun = itemgubun
oupchejumun.FRectItemId = itemid
oupchejumun.FRectItemOption = itemoption
oupchejumun.FRectMakerID = makerid

oupchejumun.GetUpcheOrderList()

dim i
dim totShopJumunCnt, totUpcheJumunCnt, totUpcheRealCnt
totShopJumunCnt = 0
totUpcheJumunCnt = 0
totUpcheRealCnt = 0

%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" >
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			* 최근 2개월 주문만 표시됩니다.(해외배송 제외)
		</td>

		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();" disabled>
		</td>
	</tr>
	</form>
</table>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmview" method="get">
	<input type="hidden" name="iid" value="">
	</form>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="8">
			<b>샵주문</b> 검색결과 : <b><%= oshopjumun.FResultCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="70">주문코드</td>
		<td width="100">물류코드</td>
		<td width="140">샵이름</td>
		<td width="140">브랜드</td>
		<td width="60">주문상태</td>
		<td width="70">입고요청일</td>
		<td width="50">주문수량</td>
		<td>비고</td>
	</tr>
	<% if oshopjumun.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF">
	  	<td colspan="8" align="center">검색결과가 없습니다.</td>
	</tr>
	<% else %>
	<%
	for i=0 to oshopjumun.FResultCount -1
		totShopJumunCnt = totShopJumunCnt + oshopjumun.FItemList(i).Fbaljuitemno
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oshopjumun.FItemList(i).Fbaljucode %></td>
		<td><%= oshopjumun.FItemList(i).Fprdcode %></td>
		<td><%= oshopjumun.FItemList(i).Fbaljuid %></td>
		<td><%= oshopjumun.FItemList(i).Fmakerid %></td>
		<td><%= oshopjumun.FItemList(i).GetStateName %></td>
		<td><%= oshopjumun.FItemList(i).Fscheduledate %></td>
		<td><%= oshopjumun.FItemList(i).Fbaljuitemno %></td>
		<td></td>
	</tr>
	<% next %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="6">합계</td>
		<td><b><%= totShopJumunCnt %></b></td>
		<td></td>
	</tr>
  <% end if %>
</table>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmview" method="get">
	<input type="hidden" name="iid" value="">
	</form>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="10">
			<b>업체주문</b> 검색결과 : <b><%= oupchejumun.FResultCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="70">주문코드</td>
		<td width="100">물류코드</td>
		<td width="140">공급처</td>
		<td width="60">주문상태</td>
		<td width="70">입고요청일</td>
		<td width="70">배송일</td>
		<td width="50">주문수량</td>
		<td width="50">확정수량</td>
		<td width="100">송장번호</td>
		<td>비고</td>
	</tr>
	<% if oupchejumun.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF">
	  	<td colspan="10" align="center">검색결과가 없습니다.</td>
	</tr>
	<% else %>
	<%
	for i=0 to oupchejumun.FResultCount -1
		totUpcheJumunCnt = totUpcheJumunCnt + oupchejumun.FItemList(i).Fbaljuitemno
		totUpcheRealCnt = totUpcheRealCnt + oupchejumun.FItemList(i).Frealitemno
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oupchejumun.FItemList(i).Fbaljucode %></td>
		<td><%= oupchejumun.FItemList(i).Fprdcode %></td>
		<td><%= oupchejumun.FItemList(i).Ftargetid %></td>
		<td><%= oupchejumun.FItemList(i).GetStateName %></td>
		<td><%= oupchejumun.FItemList(i).Fscheduledate %></td>
		<td><%= oupchejumun.FItemList(i).Fbeasongdate %></td>
		<td><%= oupchejumun.FItemList(i).Fbaljuitemno %></td>
		<td><%= oupchejumun.FItemList(i).Frealitemno %></td>
		<td><%= oupchejumun.FItemList(i).Fsongjangno %></td>
		<td></td>
	</tr>
	<% next %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="6">합계</td>
		<td><font color="<%= CHKIIF(totShopJumunCnt>totUpcheJumunCnt, "red", "black") %>"><b><%= totUpcheJumunCnt %></b></font></td>
		<td><%= totUpcheRealCnt %></td>
		<td colspan="2">
			<%= CHKIIF(totShopJumunCnt>totUpcheJumunCnt, "<font color='red'><b>업체 주문부족 상태입니다.</b></font>", "") %>
		</td>
	</tr>
  <% end if %>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
