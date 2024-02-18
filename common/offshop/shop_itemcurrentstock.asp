<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->

<%

const C_STOCK_DAY=7

dim itemgubun, itemid, itemoption, shopid, barcode, makerid
itemgubun  = requestCheckVar(request("itemgubun"),2)
itemid     = requestCheckVar(request("itemid"),9)
itemoption = requestCheckVar(request("itemoption"),4)
barcode    = requestCheckVar(request("barcode"),32)
shopid     = requestCheckVar(request("shopid"),32)

'/매장
if (C_IS_SHOP) then

	'//직영점일때
	if C_IS_OWN_SHOP then

		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")	'"7321"
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

if (barcode <> "") then
    if Not (fnGetItemCodeByPublicBarcode(barcode,itemgubun,itemid,itemoption)) then
        if (Len(barcode)=12) then
            itemgubun   = Left(barcode, 2)
            itemid      = CStr(Mid(barcode, 3, 6) + 0)
            itemoption  = Right(barcode, 4)
        elseif (Len(barcode)=14) then
            itemgubun   = Left(barcode, 2)
            itemid      = CStr(Mid(barcode, 3, 8) + 0)
            itemoption  = Right(barcode, 4)
        end if
    end if
elseif (itemid<>"") then
    if (itemid>=1000000) then
        barcode = itemgubun + "" + Format00(8,itemid) + "" + itemoption
    else
        barcode = itemgubun + "" + Format00(6,itemid) + "" + itemoption
    end if
end if


if (shopid = "") then
        shopid = ""
end if

dim nowyyyymmdd
nowyyyymmdd = Left(now(), 10)


'==============================================================================
'상품기본정보
if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim ojaegoitem
set ojaegoitem = new COffShopItem
ojaegoitem.FRectItemGubun   = itemgubun
ojaegoitem.FRectItemID      = itemid
ojaegoitem.FRectItemOption  = itemoption
ojaegoitem.FRectShopid      = shopid
if (itemid<>"") then
	ojaegoitem.GetOffOneItem
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


dim BasicMonth
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)

'==============================================================================
'상품요약정보(daily)
dim odaysummary
set odaysummary = new CShopItemSummary

odaysummary.FRectShopID =  shopid
odaysummary.FRectItemGubun =  itemgubun
odaysummary.FRectItemId =  itemid
odaysummary.FRectItemOption =  itemoption
odaysummary.FRectStartDate  =  BasicMonth + "-01"
if itemid<>"" then
	odaysummary.GetShopItemDailySummaryList
end if


dim i, buf
dim dstart, dend

dim sysstockSum
dim availstockSum
dim realstockSum

sysstockSum    =0
availstockSum  =0
realstockSum   =0

dim IsUpcheWitakItem
if (ojaegoitem.FResultCount>0) then
    IsUpcheWitakItem = (ojaegoitem.FOneItem.Fcomm_cd="B012")
else
    IsUpcheWitakItem = False
end if

%>

<script type='text/javascript'>

function popOffItemEdit(ibarcode){
	<% if (C_IS_SHOP) then %>

		//직영점일때
		<% if C_IS_OWN_SHOP then %>
			var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
			popwin.focus();
		<% else %>
			return;
		<% end if %>
	<% else %>
		<% if (C_IS_Maker_Upche) then %>
			var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
			popwin.focus();
		<% else %>
			<% if (Not C_ADMIN_USER) then %>
				return;
			<% else %>
				var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
				popwin.focus();
			<% end if %>
		<% end if %>
	<% end if %>
}

function refreshOffStockByItem(itemgubun,itemid,itemoption){
    if (frmrefresh.shopid.value.length<1){
        alert('매장을 선택후 검색 후 사용하세요.');
        return;
    }

    if (confirm('재고 내역을 전체 새로고침 하시겠습니까?')){
		frmrefresh.mode.value="OFFitemAllRefresh";
		frmrefresh.submit();
	}
}

function refreshOffStockByItemV2(itemgubun,itemid,itemoption){
    if (frmrefresh.shopid.value.length<1){
        alert('매장을 선택후 검색 후 사용하세요.');
        return;
    }

    if (confirm('재고 내역을 전체 새로고침 하시겠습니까?')){
		frmrefresh.mode.value="OFFStockitemRecentRefresh";
		frmrefresh.submit();
	}
}

function refreshAccStockShop(comp,yyyymm){
	var frm =document.frmrefresh;
	frm.mode.value = "itemAccStockShop";
	frm.yyyymm.value = yyyymm;
	

    var confirmstr = yyyymm+'월 매장전체 기말재고를 새로고침 하시겠습니까?'

    if (confirm(confirmstr)){
		comp.disabled=true;
		frm.submit();
	}
}




function popOffErrInput(shopid,itemgubun,itemid,itemoption){
    //입력창에서 체크 : 업체위탁 상품.
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
    alert('권한이 없습니다. - 업체위탁 상품만 재고 수정 가능.');
    return;
    <% else %>
    var popwin = window.open('/common/offshop/popOffrealerrinput.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popOffrealerrinput','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
	<% end if %>
}

function popOffStockBaditem(fromdate,todate,itembarcode,errType,shopid){
    <% if (C_ADMIN_USER) then %>
	var popwin = window.open('/admin/stock/off_baditem_list.asp?fromdate=' + fromdate + '&todate=' + todate + '&itembarcode=' + itembarcode +  '&errType=' + errType + '&shopid=' + shopid,'popoffbaditemlist','width=900,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
	<% end if %>
}

function PopItemUpcheIpChulListOffLine(fromdate,todate,itemgubun,itemid,itemoption, ipchulflag, shopid){
    <% if (C_ADMIN_USER) then %>
	var popwin = window.open('/common/pop_upcheipgolist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
	<% end if %>
}

function PopItemSellListOffLine(fromdate,todate,itemgubun,itemid,itemoption, ipchulflag, shopid){
    <% if (C_ADMIN_USER) then %>
	var popwin = window.open('/common/pop_selllist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
	<% end if %>
}


function popAsgnAccGbn(yyyymm,shopid,itemgubun,itemid,itemoption){
return;
    <% if (C_ADMIN_AUTH) then %>
    var popwin = window.open('/admin/newreport/popAssignMonthlyAccMwgubun.asp?stockPlace=S&yyyymm='+yyyymm+'&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
    <% end if %>
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
			<% if (C_IS_SHOP) then %>
				<% if C_IS_OWN_SHOP then %>
					매장 : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp;
				<% else %>
					매장 : <%= shopid %>
				<% end if %>
			<% else %>
				<% if (C_IS_Maker_Upche) then %>
					매장 : <% drawSelectBoxOpenOffShop "shopid",shopid %>
				<% else %>
					<% if (Not C_ADMIN_USER) then %>
					<% else %>
						매장 : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp;
					<% end if %>
				<% end if %>
			<% end if %>

        	<% if (C_IS_Maker_Upche) then %>
        	    <input type="hidden" name="barcode" value="<%= barcode %>">
        	<% else %>
        	바코드: <input type="text" class="text" name="barcode" value="<%= barcode %>" size=16 maxlength=20 <%= ChkIIF(C_ADMIN_USER,"","readonly") %> >&nbsp;&nbsp;
			&nbsp;
			<% end if %>
			<input type="button" class="button" value=" 검 색 " onClick="document.frm.submit();">
        </td>
        <td valign="top" align="right">
            <% if (C_ADMIN_USER) or (C_OFF_AUTH) then %>
            <input type="button" class="button" value="재고 새로 고침" onClick="refreshOffStockByItem('<%= itemgubun %>','<%= itemid %>','<%= itemoption %>')">
			&nbsp;
			<input type="button" class="button" value="재고 새로 고침 V2" onClick="refreshOffStockByItemV2('<%= itemgubun %>','<%= itemid %>','<%= itemoption %>')">
            <% end if %>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- 표 상단바 끝-->

<% if ojaegoitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan="5" width="110" valign=top align=center><img src="<%= ojaegoitem.FOneItem.GetImageList %>" width="100" height="100"></td>
      	<td width="60"><b>*상품정보</b></td>
      	<td width="300">
			<% if (C_IS_SHOP) then %>
				<% if C_IS_OWN_SHOP then %>
					<input type="button" class="button" value="수정" onclick="popOffItemEdit('<%= barcode %>');">
				<% else %>
					
				<% end if %>
			<% else %>
				<% if (C_IS_Maker_Upche) then %>
					<input type="button" class="button" value="수정" onclick="popOffItemEdit('<%= barcode %>');">
				<% else %>
					<% if (Not C_ADMIN_USER) then %>
					<% else %>
						<input type="button" class="button" value="수정" onclick="popOffItemEdit('<%= barcode %>');">
					<% end if %>
				<% end if %>
			<% end if %>
      	</td>
      	<td width="80">거래방식 </td>
      	<td colspan=2>
          	<% if Not ojaegoitem.FOneItem.IsShopContractExists then %>
          	<font color="red"><strong>미지정</strong></font>
          	<% else %>
          	<%= GetJungsanGubunName(ojaegoitem.FOneItem.FComm_cd) %>
          	    <% if (C_ADMIN_USER) then %>
          	    [<%= ojaegoitem.FOneItem.FMakerMargin %> -&gt; <%= ojaegoitem.FOneItem.FshopMargin %>]
          	    <% end if %>
          	<% end if %>
      	</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>상품코드</td>
      	<td><%= ojaegoitem.FOneItem.GetBarCode %></td>
      	<td>판매가</td>
      	<td colspan=2>
      	    <% if (ojaegoitem.FOneItem.IsOffSaleItem) then %>
      	    <strike><%= FormatNumber(ojaegoitem.FOneItem.FShopItemOrgprice,0) %></strike>
      	    &nbsp;&nbsp;
      	    <%= FormatNumber(ojaegoitem.FOneItem.Fshopitemprice,0) %>
      	    <% else %>
      	    <%= FormatNumber(ojaegoitem.FOneItem.Fshopitemprice,0) %>
      	    <% end if %>
      	</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>브랜드ID</td>
      	<td><%= ojaegoitem.FOneItem.FMakerid %></td>
      	<% if (C_IS_Maker_Upche) or (C_ADMIN_USER) then %>
      	<td>매입가(업체)</td>
      	<td colspan=2>
          	<% if ojaegoitem.FOneItem.IsShopContractExists then %>
          	    <%= FormatNumber(ojaegoitem.FOneItem.GetOfflineBuycash,0) %>
          	<% end if %>
      	</td>
      	<% elseif (C_IS_SHOP) then %>
      	<td>공급가(SHOP)</td>
      	<td colspan=2>
      	<% if ojaegoitem.FOneItem.IsShopContractExists then %>
      	    <%= FormatNumber(ojaegoitem.FOneItem.GetOfflineSuplycash,0) %>
      	<% end if %>
        </td>
      	<% else %>
      	<td></td>
      	<td colspan=2></td>
      	<% end if %>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>상품명</td>
      	<td>
      	    <%= ojaegoitem.FOneItem.FShopItemName %>
      	    <% if (ojaegoitem.FOneItem.FShopItemOptionName<>"") then %>
      	    <font color="blue">[<%= ojaegoitem.FOneItem.FShopItemOptionName %>]</font>
      	    <% end if %>
      	</td>
      	<% if (C_ADMIN_USER) then %>
      	<td>공급가(SHOP)</td>
      	<td colspan=2>
      	<% if ojaegoitem.FOneItem.IsShopContractExists then %>
      	    <%= FormatNumber(ojaegoitem.FOneItem.GetOfflineSuplycash,0) %>
      	<% end if %>
        </td>
        <% else %>
        <td></td>
      	<td colspan=2></td>
        <% end if %>
    </tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="20" valign="bottom">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<br>시스템 재고     = 입고/반품합 + 업체입고/반품합 + 총판매/반품
	        <br>실사 재고       = 시스템재고 + 입력오차
	        <br>유효재고        = 시스템 재고 + 입력오차 + 샘플 <!-- + 불량 -->

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
        <td width="60">&nbsp;</td>
    	<td width="60">총입고<br>(텐바이텐)</td>
    	<td width="60">총반품<br>(텐바이텐)</td>
    	<td width="60">총입고<br>(업체)</td>
    	<td width="60">총반품<br>(업체)</td>
    	<td width="60">총판매</td>
    	<td width="60">총반품</td>
    	<td width="60" bgcolor="F4F4F4">시스템재고<br>(누적)</td>
    	<td width="60">오차</td>
    	<td width="60" bgcolor="F4F4F4">실사재고<br>(누적)</td>
    	<td width="60">샘플</td>
    	<!-- <td width="60">불량</td> -->
    	<td width="60" bgcolor="F4F4F4">유효재고<br>(누적)</td>
		<td width="60">배송중</td>
		<td width="60">반품중</td>
		<td width="60" bgcolor="F4F4F4">매장재고<br>(현재)</td>
    	<td>비고</td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
        <td>&nbsp;</td>
    	<td><%= ocursummary.FOneItem.Flogicsipgono %></td>
    	<td><%= ocursummary.FOneItem.Flogicsreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fsellno %></td>
    	<td><%= ocursummary.FOneItem.Fresellno %></td>
    	<td bgcolor="F4F4F4"><b><%= ocursummary.FOneItem.Fsysstockno %></b></td>
    	<td><%= ocursummary.FOneItem.Ferrrealcheckno %></td>
    	<td><b><%= ocursummary.FOneItem.Frealstockno %></b></td>
    	<td><%= ocursummary.FOneItem.Ferrsampleitemno %></td>
    	<!-- <td><%= ocursummary.FOneItem.Ferrbaditemno %></td> -->
    	<td bgcolor="F4F4F4"><%= ocursummary.FOneItem.getAvailStock %></td>
		<td bgcolor="F4F4F4"><%= ocursummary.FOneItem.Flogischulgo %></td>
		<td bgcolor="F4F4F4"><%= ocursummary.FOneItem.Flogisreturn %></td>
		<td bgcolor="FFDDDD"><%= ocursummary.FOneItem.getShopRealStock %></td>
    	<td>
    	    <% if ocursummary.FOneItem.Fpreorderno>0 then %>
    	    기주문 : <%= ocursummary.FOneItem.Fpreorderno %>
    	        <% if (ocursummary.FOneItem.Fpreorderno<>ocursummary.FOneItem.FpreordernoFix) then %>
                    => <strong> <%= ocursummary.FOneItem.FpreordernoFix %></strong>
    	        <% end if %>
    	    <% end if %>
    	</td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
        <td colspan="9"></td>
        <td></td>
        <td><input type="button" class="button" value="샘플" onClick="popOffErrInput('<%= shopid %>','<%= itemgubun %>','<%= itemid %>','<%= itemoption %>');"></td>
        <!-- <td><input type="button" class="button" value="불량"></td> -->
        <td></td>
        <td></td>
		<td></td>
		<td bgcolor="#FFDDDD"><input type="button" class="button" value="실사" onClick="popOffErrInput('<%= shopid %>','<%= itemgubun %>','<%= itemid %>','<%= itemoption %>');"></td>
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
    	<td width="60" bgcolor="F4F4F4">시스템재고<br>(누적)</td>
    	<td width="60">오차</td>
    	<td width="60" bgcolor="F4F4F4">실사재고<br>(누적)</td>
    	<td width="60">샘플</td>
    	<!-- <td width="60">불량</td> -->
    	<td width="60" bgcolor="F4F4F4">유효재고<br>(누적)</td>
    	<td>비고</td>
    </tr>
    <% for i=0 to omonsummary.FResultcount-1 %>
    <%
    dstart = omonsummary.FItemList(i).Fyyyymm + "-01"
    dend = Left(dateadd("m",1,dstart),7)+"-01"
    dend = Left(dateadd("d",-1,dend),10)

    sysstockSum    = sysstockSum    + omonsummary.FItemList(i).Fsysstockno
    availstockSum  = availstockSum  + omonsummary.FItemList(i).getAvailStock
    realstockSum   = realstockSum   + omonsummary.FItemList(i).Frealstockno
    %>
    <tr bgcolor="#FFFFFF" height="20" align=center>
    	<td><%= omonsummary.FItemList(i).Fyyyymm %></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= sysstockSum %></b></td>
    	<td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','D','<%= shopid %>')"><%= omonsummary.FItemList(i).Ferrrealcheckno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= realstockSum %></b></td>
    	<td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','S','<%= shopid %>')"><%= omonsummary.FItemList(i).Ferrsampleitemno %></a></td>
    	<!-- <td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','B','<%= shopid %>')"><%= omonsummary.FItemList(i).Ferrbaditemno %></a></td> -->
    	<td bgcolor="F4F4F4"><b><%= availstockSum %></b></td>
    	<td>
    	    <a href="javascript:popAsgnAccGbn('<%=omonsummary.FItemList(i).Fyyyymm%>','<%=shopid%>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>');">
    	   <%=omonsummary.FItemList(i).Fcomm_cd%>
    	    /
    	    <%=omonsummary.FItemList(i).FCenterMwdiv%>
    	    /
    	    <% if Not isNULL(omonsummary.FItemList(i).FAccSysstockno)  THEN %>
    	       <% if sysstockSum<>omonsummary.FItemList(i).FAccSysstockno then %>
    	            <font color=red><%=omonsummary.FItemList(i).FAccSysstockno%></font>
    	       <% else %>
    	            <%=omonsummary.FItemList(i).FAccSysstockno%>
    	        <% end if %>
    	    <% end if %>
    	   /
    	    </a>
    	</td>
    </tr>
    <% next %>
    <% if (omonsummary.FResultcount < 1) then %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan="14" align="center">[월별 데이타 없음]</td>
    </tr>
    <% end if %>
    <%
    dstart = "2001-10-10"
    dend = Left(dateadd("m",-1,nowyyyymmdd),7)+"-01"
    dend = Left(dateadd("d",-1,dend),10)

    %>
    <tr bgcolor="#DDDDFF" height="20" align=center>
    	<td>합계<br>(2개월전)</td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= sysstockSum %></b><%= ChkIIF(sysstockSum<>olastmonsummary.FOneItem.Fsysstockno,"<font color=red>(" & (olastmonsummary.FOneItem.Fsysstockno) & ")</font>","") %></td>
    	<td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','D','<%= shopid %>')"><%= olastmonsummary.FOneItem.Ferrrealcheckno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= olastmonsummary.FOneItem.Frealstockno %></b><%= ChkIIF(realstockSum<>olastmonsummary.FOneItem.Frealstockno,"<font color=red>(" + CStr(realstockSum) + ")</font>","") %></td>
    	<td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','S','<%= shopid %>')"><%= olastmonsummary.FOneItem.Ferrsampleitemno %></a></td>
    	<!-- <td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','B','<%= shopid %>')"><%= olastmonsummary.FOneItem.Ferrbaditemno %></a></td> -->
    	<td bgcolor="F4F4F4"><b><%= availstockSum %></b><%= ChkIIF(availstockSum<>olastmonsummary.FOneItem.getAvailStock,"<font color=red>(" & (olastmonsummary.FOneItem.getAvailStock) & ")</font>","") %></td>
    	<td></td>
    </tr>
    <% for i=0 to odaysummary.FResultcount-1 %>
    <%
    sysstockSum    = sysstockSum    + odaysummary.FItemList(i).Fsysstockno
    availstockSum  = availstockSum  + odaysummary.FItemList(i).getAvailStock
    realstockSum   = realstockSum   + odaysummary.FItemList(i).Frealstockno
    %>
    <tr bgcolor="#FFFFFF" height="20" align=center>
    	<td><%= odaysummary.FItemList(i).Fyyyymmdd %></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= sysstockSum %></b></td>
    	<td><a href="javascript:popOffStockBaditem('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= ojaegoitem.FOneItem.GetBarCode %>','D','<%= shopid %>')"><%= odaysummary.FItemList(i).Ferrrealcheckno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= realstockSum %></b></td>
    	<td><a href="javascript:popOffStockBaditem('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= ojaegoitem.FOneItem.GetBarCode %>','S','<%= shopid %>')"><%= odaysummary.FItemList(i).Ferrsampleitemno %></a></td>
    	<!-- <td><a href="javascript:popOffStockBaditem('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= ojaegoitem.FOneItem.GetBarCode %>','B','<%= shopid %>')"><%= odaysummary.FItemList(i).Ferrbaditemno %></a></td> -->
    	<td bgcolor="F4F4F4"><b><%= availstockSum %></b></td>
    	<td></td>
    </tr>
    <% next %>
    <% if (odaysummary.FResultcount < 1) then %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan="14" align="center">[일별 서머리 데이타가 존재하지 않습니다.]</td>
    </tr>
    <% end if %>
    <tr bgcolor="#DDDDFF" height="20" align=center>
    	<td>합계</td>
    	<td><%= ocursummary.FOneItem.Flogicsipgono %></td>
    	<td><%= ocursummary.FOneItem.Flogicsreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fsellno %></td>
    	<td><%= ocursummary.FOneItem.Fresellno %></td>
    	<td bgcolor="F4F4F4"><b><%= sysstockSum %></b></td>
    	<td><%= ocursummary.FOneItem.Ferrrealcheckno %></td>
    	<td bgcolor="F4F4F4"><b><%= realstockSum %></b></td>
    	<td><%= ocursummary.FOneItem.Ferrsampleitemno %></td>
    	<!-- <td><%= ocursummary.FOneItem.Ferrbaditemno %></td> -->
    	<td bgcolor="F4F4F4"><b><%= availstockSum %></b></td>
    	<td>
		<% if (C_ADMIN_USER) or (C_OFF_AUTH) then %>
			<input type="button" value="기말재작성 <%=LEFT(dateadd("m",-1,now()),7)%>" onClick="refreshAccStockShop(this,'<%=LEFT(dateadd("m",-1,now()),7)%>')">
			<input type="button" value="기말재작성 <%=LEFT(dateadd("m",-0,now()),7)%>" onClick="refreshAccStockShop(this,'<%=LEFT(dateadd("m",-0,now()),7)%>')">
		<% end if %>
		</td>
    </tr>
</table>

<% else %>
<table width="100%" height="30" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
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



<%
set ojaegoitem = Nothing
set ocursummary = Nothing
set omonsummary = Nothing
set ocursummary = Nothing
%>
<form name=frmrefresh method=post action="/common/offshop/shop_stockrefresh_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
<input type="hidden" name="yyyymm" value="">
</form>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
