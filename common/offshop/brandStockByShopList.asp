<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  매장별 브랜드별 상품별 재고현황
' History : 2018-04-11 이상구 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%

dim research
dim shopid, makerid
dim itembarcode, usingyn, NoZeroStock, showminusOnly
dim itemgubun, itemid, itemoption

''매장
if (C_IS_SHOP) then
    dbget.close : response.end
end if

''업체
if (C_IS_Maker_Upche) then
    dbget.close : response.end
end if

shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
itembarcode  = RequestCheckVar(request("itembarcode"),32)
usingyn      = RequestCheckVar(request("usingyn"),1)
research     = RequestCheckVar(request("research"),2)
NoZeroStock  = RequestCheckVar(request("NoZeroStock"),32)
showminusOnly  = RequestCheckVar(request("showminusOnly"),32)

if (itembarcode <> "") then
    if Not (fnGetItemCodeByPublicBarcode(itembarcode,itemgubun,itemid,itemoption)) then
        if Len(itembarcode)=12 then
            itemgubun   = Left(itembarcode, 2)
            itemid      = CStr(Mid(itembarcode, 3, 6) + 0)
            itemoption  = Right(itembarcode, 4)
        elseif Len(itembarcode)=14 then
            itemgubun   = Left(itembarcode, 2)
            itemid      = CStr(Mid(itembarcode, 3, 8) + 0)
            itemoption  = Right(itembarcode, 4)
        else
            itemgubun   = Left(itembarcode, 2)
            itemid      = CStr(0)
            itemoption  = Right(itembarcode, 4)
        end if
    end if
end if

dim i, j, k

dim oOffStock
set oOffStock = new CShopItemSummary
if (shopid <> "") then
	oOffStock.FRectShopID       = "[" & shopid & "]"
end if
oOffStock.FRectMakerID      = makerid
''oOffStock.FRectIsUsing      = usingyn
''oOffStock.FRectNoZeroStock  = NoZeroStock
''oOffStock.FRectShowMinusOnly  = showminusOnly
if (itembarcode <> "") then
    oOffStock.FRectItemGubun    = itemgubun
    oOffStock.FRectItemId       = itemid
    oOffStock.FRectItemOption   = itemoption
end if

''if (makerid = "") or (makerid <> "" and shopid <> "") then
	oOffStock.GetDirectShopList
''end if


dim oOffStockBrand
set oOffStockBrand = new CShopItemSummary
if (shopid <> "") then
	oOffStockBrand.FRectShopID       = "[" & shopid & "]"
else
	for i = 0 to oOffStock.FResultCount - 1
		oOffStockBrand.FRectShopID = oOffStockBrand.FRectShopID & "," & "[" & oOffStock.FItemList(i).Fshopid & "]"
	next
	oOffStockBrand.FRectShopID = Mid(oOffStockBrand.FRectShopID, 2, 4000)
end if
oOffStockBrand.FRectMakerID      = makerid
''oOffStockBrand.FRectIsUsing      = usingyn
''oOffStockBrand.FRectNoZeroStock  = NoZeroStock
''oOffStockBrand.FRectShowMinusOnly  = showminusOnly
if (itembarcode <> "") then
    oOffStockBrand.FRectItemGubun    = itemgubun
    oOffStockBrand.FRectItemId       = itemid
    oOffStockBrand.FRectItemOption   = itemoption
end if

dim rs
dim rowCnt, item, val
''if (makerid = "") or (makerid <> "" and shopid <> "") then
	rs = oOffStockBrand.GetDirectShopBrandList
''end if

dim totCnt, totPrice

%>
<script language='javascript'>

function jsViewBrandDetail(shopid, makerid) {
	var frm = document.frm;
	/*
	if ((shopid == '') && (makerid != '')) {
		alert('매장과 브랜드를 모두 선택해야만 상세내역을 조회할 수 있습니다.');
		return;
	}
	*/
	frm.shopid.value = shopid;
	frm.makerid.value = makerid;

	frm.submit();
}

</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	    매장 :
		<% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp;

		브랜드 :
		<% drawSelectBoxDesignerwithName "makerid", makerid %> &nbsp;&nbsp;

		<!--
		상품바코드 :
		<input type="text" class="text" name="itembarcode" value="<%= itembarcode %>" size="20" maxlength="32">
		-->
		<br>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!--
		상품 사용구분 : <% drawSelectBoxUsingYN "usingyn", usingyn %> &nbsp;&nbsp;
		-->

		<input type="checkbox" name="NoZeroStock" <%= CHKIIF(NoZeroStock="on","checked","") %> checked disabled> <font color="red">재고0인 상품 검색 안함.</font>
		<!--
		&nbsp;
		<input type="checkbox" name="showminusOnly" <%= CHKIIF(showminusOnly="on","checked","") %> > 마이너스 재고만.</font>
		-->
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p />

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td rowspan="2">브랜드</td>
		<% if (makerid <> "") then %>
		<td width="30" rowspan="2">구분</td>
		<td width="60" rowspan="2">상품ID</td>
		<td width="40" rowspan="2">옵션</td>
		<td rowspan="2">상품명</td>
		<td rowspan="2">옵션명</td>
		<td width="80" rowspan="2">판매가</td>
		<td width="80" rowspan="2">
			본사<br>매입가
		</td>
		<td width="80" rowspan="2">
			물류<br>재고
		</td>
		<% end if %>
		<% for i = 0 to oOffStock.FResultCount - 1 %>
		<td width="120" colspan="2"><a href="javascript:jsViewBrandDetail('<%= oOffStock.FItemList(i).Fshopid %>', '')"><%= oOffStock.FItemList(i).Fshopid %></a></td>
		<% next %>
		<td rowspan="2">비고</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<% for i = 0 to oOffStock.FResultCount - 1 %>
		<td width="40" bgcolor="F4F4F4">현재고</td>
		<td width="60">재고금액</td>
		<% next %>
	</tr>
	<% if (makerid <> "") then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td>합계</td>
		<td colspan="8"></td>
		<% for i = 0 to oOffStock.FResultCount - 1 %>
		<%
		totCnt = 0
		totPrice = 0
		For j = 0 To UBound(rs,2)
			if Not IsNull(rs((i + 9),j)) then
				totCnt = totCnt + rs((i + 9),j)
			end if
			if Not IsNull(rs((i + 9),j)) and Not IsNull(rs((i + 9),j)) then
				totPrice = totPrice + rs(7,j)*rs((i + 9),j)
			end if
		next
		%>
		<td width="40" bgcolor="F4F4F4"><%= FormatNumber(totCnt, 0) %></td>
		<td width="60"><%= FormatNumber(totPrice, 0) %></td>
		<% next %>
		<td></td>
	</tr>
	<% end if %>
	<%
	If IsArray(rs) Then
		rowCnt = UBound(rs,2) + 1
		For j = 0 To UBound(rs,2)
	%>
	<tr align="center" bgcolor="#FFFFFF" height="25">
		<td><a href="javascript:jsViewBrandDetail('<%= shopid %>', '<%= rs(0,j) %>')"><%= rs(0,j) %></a></td>
		<% if (makerid = "") then %>
		<% for i = 0 to oOffStock.FResultCount - 1 %>
		<td><a href="javascript:jsViewBrandDetail('<%= oOffStock.FItemList(i).Fshopid %>', '<%= rs(0,j) %>')"><%= rs((i + 1),j) %></a></td>
		<td></td>
		<% next %>
		<% else %>
		<td><%= rs(1,j) %></td>
		<td><%= rs(2,j) %></td>
		<td><%= rs(3,j) %></td>
    	<td align="left">
          	<%= db2html(rs(4,j)) %>
        </td>
		<td align="left">
          	<%= db2html(rs(5,j)) %>
        </td>
		<td><%= FormatNumber(rs(6,j), 0) %></td>
		<td><%= FormatNumber(rs(7,j), 0) %></td>
		<td><%= FormatNumber(rs(8,j), 0) %></td>
		<%
		for i = 0 to oOffStock.FResultCount - 1
			totCnt = totCnt + rs((i + 9),j)
			totPrice = totPrice + rs(7,j)*rs((i + 9),j)
		%>
		<td>
			<%= CHKIIF(IsNull(rs((i + 9),j)), 0, rs((i + 9),j)) %>
		</td>
		<td>
			<% if IsNull(rs((i + 9),j)) then %>
				0
			<% elseif IsNull(rs(7,j)) then %>
				0
			<% else %>
				<%= CHKIIF(IsNull(rs(7,j)*rs((i + 9),j)), 0, FormatNumber(rs(7,j)*rs((i + 9),j),0)) %>
			<% end if %>
		</td>
		<% next %>
		<% end if %>
		<td></td>
	</tr>
	<%
		next
	%>
	<% if (makerid <> "") and False then %>
	<tr align="center" bgcolor="#FFFFFF" height="25">
		<td>합계</td>
		<td colspan="7"></td>
		<td><%= FormatNumber(totCnt, 0) %></td>
		<td><%= FormatNumber(totPrice, 0) %></td>
		<td></td>
	</tr>
	<% end if %>
	<%
	end if
	%>
</table>

<% if (makerid <> "" and shopid = "") and False then %><br />* <font color="red">매장과 브랜드를 모두 선택</font>해야만 상세내역을 조회할 수 있습니다.<% end if %>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
