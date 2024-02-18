<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_summary.asp"-->
<!-- #include virtual="/lib/classes/stock/realjaegocls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%

const C_STOCK_DAY=7

dim itemgubun, itemid, itemoption, shopid, barcode
itemgubun  = request("itemgubun")
itemid     = request("itemid")
itemoption = request("itemoption")
barcode     = request("barcode")
shopid     = request("shopid")

if (barcode <> "") then
    if BF_IsMaybeTenBarcode(barcode) then
        itemgubun 	= BF_GetItemGubun(barcode)
    	itemid 		= BF_GetItemId(barcode)
    	itemoption 	= BF_GetItemOption(barcode)
    end if
else
    IF (itemid>=1000000) THEN
        barcode = itemgubun + "" + Format00(8,itemid) + "" + itemoption
    ELSE
        barcode = itemgubun + "" + Format00(6,itemid) + "" + itemoption
    END IF
end if


if (shopid = "") then
        shopid = "-"
end if

dim nowyyyymmdd
nowyyyymmdd = Left(now(), 10)


'==============================================================================
'상품기본정보
if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim ojaegoitem
set ojaegoitem = new CRealJaeGo
ojaegoitem.FRectItemGubun = itemgubun
ojaegoitem.FRectItemID = itemid
ojaegoitem.FRectItemOption = itemoption
if itemid<>"" then
	ojaegoitem.GetOfflineItemDefaultData
end if

dim oitemoption
set oitemoption = new CItemOptionInfo
oitemoption.FRectItemID =  itemid
if itemid<>"" then
	oitemoption.getOptionList
end if

if (oitemoption.FResultCount<1) then
	itemoption = "0000"
end if


'==============================================================================
'상품요약정보(current)
dim ocursummary
set ocursummary = new CShopItemSummary

ocursummary.FRectShopID =  shopid
ocursummary.FRectItemGubun =  itemgubun
ocursummary.FRectItemId =  itemid
ocursummary.FRectItemOption =  itemoption

if itemid<>"" then
	ocursummary.GetShopItemCurrentSummary
end if


'==============================================================================
'상품요약정보(monthly)
dim omonsummary
set omonsummary = new CShopItemSummary

omonsummary.FRectShopID =  shopid
omonsummary.FRectItemGubun =  itemgubun
omonsummary.FRectItemId =  itemid
omonsummary.FRectItemOption =  itemoption

if itemid<>"" then
	omonsummary.GetShopItemMonthlySummaryList
end if


'==============================================================================
'상품요약정보(last month)
dim olastmonsummary
set olastmonsummary = new CShopItemSummary

olastmonsummary.FRectShopID =  shopid
olastmonsummary.FRectItemGubun =  itemgubun
olastmonsummary.FRectItemId =  itemid
olastmonsummary.FRectItemOption =  itemoption

if itemid<>"" then
	olastmonsummary.GetShopItemLastMonthSummary
end if


'==============================================================================
'상품요약정보(daily)
dim odaysummary
set odaysummary = new CShopItemSummary

odaysummary.FRectShopID =  shopid
odaysummary.FRectItemGubun =  itemgubun
odaysummary.FRectItemId =  itemid
odaysummary.FRectItemOption =  itemoption

if itemid<>"" then
	odaysummary.GetShopItemDailySummaryList
end if


dim i, buf
dim dstart, dend

%>

<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
}

function RefreshRecentStock(yyyymmdd,itemgubun,itemid,itemoption){
	if (confirm('최근 2달 내역을 새로고침 하시겠습니까?')){
		frmrefresh.mode.value="itemrecentipchulrefresh";
		frmrefresh.submit();
	}
}

function RefreshTodayStock(itemgubun,itemid,itemoption){
	if (confirm('금일 내역을 새로고침 하시겠습니까?')){
		frmrefresh.mode.value="itemtodayipchulrefresh";
		frmrefresh.submit();
	}
}


function RefreshALLStock(yyyymmdd,itemgubun,itemid,itemoption){
	if (confirm('전체 내역을 새로고침 하시겠습니까?')){
		frmrefresh.mode.value="itemallipchulrefresh";
		frmrefresh.submit();
	}
}

function PopStockBaditem(fromdate,todate,itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poperritemlist.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popbaditemlist','width=800,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popRealErrList(fromdate,todate,itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poperritemlist.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'poperritemlist','width=800,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemUpcheIpChulListOffLine(fromdate,todate,itemgubun,itemid,itemoption, ipchulflag, shopid){
	var popwin = window.open('/common/pop_upcheipgolist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemSellListOffLine(fromdate,todate,itemgubun,itemid,itemoption, ipchulflag, shopid){
	var popwin = window.open('/common/pop_selllist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}







</script>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr valign="bottom">
		<td width="10" height="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td height="10" valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" height="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top">
		<td height="20" background="/images/tbl_blue_round_04.gif"></td>
		<td height="20" background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>OFFLINE삽별상품별재고현황</strong></font></td>
		<td height="20" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			오프라인 샵별 상품별 재고 정보입니다.
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td height="10"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td height="10" background="/images/tbl_blue_round_08.gif"></td>
		<td height="10"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method=get>
	<input type=hidden name=menupos value="<%= menupos %>">
    <tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="25" valign="bottom">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
        	샾 : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
        	바코드: <input type=text name=barcode value="<%= barcode %>" size=14 maxlength=14>&nbsp;&nbsp;
			&nbsp;
        </td>
        <td valign="top" align="right">
        <a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- 표 상단바 끝-->

<% if ojaegoitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 5 + ojaegoitem.FResultCount -1 %> width="110" valign=top align=center><img src="<%= ojaegoitem.FItemList(0).FImageList %>" width="100" height="100"></td>
      	<td width="60"><b>*상품정보</b></td>
      	<td width="300">
      	<input type="button" value="수정" onclick="PopItemSellEdit('<%= itemid %>');">
      	</td>
      	<td width="60">거래방식 :</td>
      	<td colspan=2><%= ojaegoitem.FItemList(0).getChargeDivName %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>상품코드 :</td>
      	<td><%= itemgubun %> <b><%= CHKIIF(ojaegoitem.FItemList(0).FItemID>=1000000,Format00(8,ojaegoitem.FItemList(0).FItemID),Format00(6,ojaegoitem.FItemList(0).FItemID)) %></b> <%= itemoption %></td>
      	<td>소비자가 : </td>
      	<td colspan=2><%= FormatNumber(ojaegoitem.FItemList(0).FSellcash,0) %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>브랜드ID :</td>
      	<td><%= ojaegoitem.FItemList(0).FMakerid %></td>
      	<td>샆공급가 : </td>
      	<td colspan=2><%= FormatNumber(ojaegoitem.FItemList(0).FBuycash,0) %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>상품명 :</td>
      	<td><%= ojaegoitem.FItemList(0).FItemName %></td>
      	<td></td>
      	<td colspan=2></td>
    </tr>
    <% for i=0 to ojaegoitem.FResultCount -1 %>
	    <% if ojaegoitem.FItemList(i).Foptionusing<>"Y" then %>
	    <tr bgcolor="#FFFFFF">
	      	<td><font color="#AAAAAA">옵션명 :</font></td>
	      	<td><font color="#AAAAAA"><%= ojaegoitem.FItemList(i).FItemOptionName %></font></td>
	      	<td></td>
	      	<td></td>
	      	<td></td>
	    </tr>
	    <% else %>

	    <% if ojaegoitem.FItemList(i).FItemOption=itemoption then %>
	    <tr bgcolor="#EEEEEE">
	    <% else %>
	    <tr bgcolor="#FFFFFF">
	    <% end if %>
	      	<td>옵션명 :</td>
	      	<td><%= ojaegoitem.FItemList(i).FItemOptionName %></td>
	      	<td>한정여부 : </td>
	      	<td><%= ojaegoitem.FItemList(i).FLimitYn %> (<%= ojaegoitem.FItemList(i).GetLimitStr %>)</td>
	      	<td><%= ojaegoitem.FItemList(i).Fcurrno %></td>
	    </tr>
	    <% end if %>
    <% next %>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="20" valign="bottom">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<br>시스템 총재고 = 입고/반품합 + 업체입고/반품합 - 총OFF판매합 + 기타출고/반품합
        	<!--
	        <br>시스템 유효재고 = 시스템 총재고 - 불량
	        <br>실사 재고 = 시스템 유효재고 - 입력오차

	        <br>재고파악 재고 = 실사 재고 - ON상품준비 - OFF상품준비
		<br>업체주문 재고 = 실사 재고 - ON상품준비 - ON결제완료 - OFF상품준비
		-->
		<br><br><p>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->




<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="40" valign="bottom">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td><b>*예상재고</b></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->


<table width="100%" align="center" cellpadding="2" cellspacing="1" bgcolor="#BABABA" class="a">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="60">총입고<br>(텐바이텐)</td>
    	<td width="60">총반품<br>(텐바이텐)</td>
    	<td width="60">총입고<br>(업체)</td>
    	<td width="60">총반품<br>(업체)</td>
    	<td width="60">총판매</td>
    	<td width="60">총반품</td>
    	<td width="60" bgcolor="F4F4F4">시스템재고</td>
    	<td width="60">샘플</td>
    	<td width="60">불량</td>
    	<td width="60" bgcolor="F4F4F4">유효재고</td>
    	<td width="60">오차</td>
    	<td width="60" bgcolor="F4F4F4">예상재고</td>
    	<td>비고</td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td><%= ocursummary.FOneItem.Flogicsipgono %></td>
    	<td><%= ocursummary.FOneItem.Flogicsreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fsellno %></td>
    	<td><%= ocursummary.FOneItem.Fresellno %></td>
    	<td bgcolor="F4F4F4"><b><%= ocursummary.FOneItem.Fsysstockno %></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    </tr>
</table>












<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="40" valign="bottom">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td><b>*일별 입출내역</b></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->








<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="60">일시</td>
    	<td width="60">입고<br>(텐바이텐)</td>
    	<td width="60">반품<br>(텐바이텐)</td>
    	<td width="60">입고<br>(업체)</td>
    	<td width="60">반품<br>(업체)</td>
    	<td width="60">판매</td>
    	<td width="60">반품</td>
    	<td width="60" bgcolor="F4F4F4">시스템재고</td>
    	<td width="60">샘플</td>
    	<td width="60">불량</td>
    	<td width="60" bgcolor="F4F4F4">유효재고</td>
    	<td width="60">오차</td>
    	<td width="60" bgcolor="F4F4F4">예상재고</td>
    	<td>비고</td>
    </tr>
    <% for i=0 to omonsummary.FResultcount-1 %>
    <%
    dstart = omonsummary.FItemList(i).Fyyyymm + "-01"
    dend = Left(dateadd("m",1,dstart),7)+"-01"
    dend = Left(dateadd("d",-1,dend),10)
    %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td><%= omonsummary.FItemList(i).Fyyyymm %></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= omonsummary.FItemList(i).Fsysstockno %></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    </tr>
    <% next %>
    <% if (omonsummary.FResultcount < 1) then %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan="14" align="center">데이타없음</td>
    </tr>
    <% end if %>
    <%
    dstart = "2001-10-10"
    dend = Left(dateadd("m",-1,nowyyyymmdd),7)+"-01"
    dend = Left(dateadd("d",-1,dend),10)

    %>
    <tr bgcolor="#DDDDFF" height="25" align=center>
    	<td>합계<br>(2개월전)</td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= olastmonsummary.FOneItem.Fsysstockno %></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    </tr>
    <% for i=0 to odaysummary.FResultcount-1 %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td><%= odaysummary.FItemList(i).Fyyyymmdd %></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= odaysummary.FItemList(i).Fsysstockno %></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    </tr>
    <% next %>
    <% if (odaysummary.FResultcount < 1) then %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan="14" align="center">데이타없음</td>
    </tr>
    <% end if %>
    <tr bgcolor="#DDDDFF" height="25" align=center>
    	<td>합계</td>
    	<td><%= ocursummary.FOneItem.Flogicsipgono %></td>
    	<td><%= ocursummary.FOneItem.Flogicsreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fsellno %></td>
    	<td><%= ocursummary.FOneItem.Fresellno %></td>
    	<td bgcolor="F4F4F4"><b><%= ocursummary.FOneItem.Fsysstockno %></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    </tr>
</table>







<% else %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#000000">
    <tr align="center" bgcolor="#DDDDFF">
    <td align=center bgcolor="#FFFFFF">검색 결과가 없습니다.</td>
    </tr>
</table>
<% end if %>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->


<% if (oitemoption.FResultCount>0) and (itemoption="0000") then %>
<script language='javascript'>
alert('옵션 선택 후 재 검색하세요.');
</script>
<% elseif (oitemoption.FResultCount<1) and (itemoption<>"0000") then %>
<script language='javascript'>
alert('재 검색하세요.');
</script>
<% end if %>
<%
set oitemoption = Nothing
set ojaegoitem = Nothing
set ocursummary = Nothing
set omonsummary = Nothing
set ocursummary = Nothing
%>
<form name=frmrefresh method=post action="dostockrefresh.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->