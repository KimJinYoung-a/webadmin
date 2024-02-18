<% option explicit %>
<%
'###########################################################
' Description :  오프라인매장 정리대상 상품 포함 브랜드
' History : 2011.08
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/brandStockTurnOverCls.asp"-->


<%
dim shopid, makerid, research, yyyy1, mm1
dim page
dim sMode

shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
research     = RequestCheckVar(request("research"),32)
sMode        = RequestCheckVar(request("sMode"),32)

dim usingyn, centermwdiv ,NoZeroStock, comm_cd
usingyn      = RequestCheckVar(request("usingyn"),32)
centermwdiv  = RequestCheckVar(request("centermwdiv"),32)
NoZeroStock  = RequestCheckVar(request("NoZeroStock"),32)
comm_cd      = RequestCheckVar(request("comm_cd"),32)

yyyy1         = RequestCheckVar(request("yyyy1"),4)
mm1           = RequestCheckVar(request("mm1"),2)
page          = RequestCheckVar(request("page"),10)
if (page="") then page=1
if (research="") and (sMode="") then sMode="X"

Dim PreMonth : PreMonth = DateAdd("m",-1,Now())
if (yyyy1="") then
    yyyy1 = Left(CStr(PreMonth),4)
    mm1   = Mid(CStr(PreMonth),6,2)
end if


''매장
if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if
if (research="") then NoZeroStock="on"

dim oOffOutItem
set oOffOutItem = new CBrandStockTurnOver
oOffOutItem.FCurrPage  = page
oOffOutItem.FPageSize  = 50 
oOffOutItem.FRectShopID       = shopid
oOffOutItem.FRectMakerID      = makerid
oOffOutItem.FRectComm_cd      = comm_cd
oOffOutItem.FRectYYYYMM       = yyyy1 + "-" + mm1
oOffOutItem.FRectSearchMode   = sMode

''oOffOutItem.FRectNoZeroStock  = NoZeroStock
if (shopid<>"") or (makerid<>"") then
    if ((shopid<>"") and (makerid<>"")) then oOffOutItem.FRectYYYYMM=""
        
    oOffOutItem.getOutItemList
end if

dim i
%>
<script language='javascript'>
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popBrandStock(shopid,makerid){
    var popUrl = "/common/offshop/shop_brandcurrentstock.asp?menupos=1074&shopid="+shopid+"&makerid="+makerid+"&research=on"+"&NoZeroStock=on";
    var popwin = window.open(popUrl,'popBrandStock','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popBrandStockTaking(shopid,makerid){
    var popUrl = "/common/offshop/shop_brandcurrentstock_taking.asp?menupos=1074&shopid="+shopid+"&makerid="+makerid+"&research=on"+"&NoZeroStock=on";
    var popwin = window.open(popUrl,'popBrandStock','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popBrandStockTakingInput(stIdx){
    var popUrl = "/common/offshop/shop_brandcurrentstock_byjobkey.asp?idx="+stIdx+"&sType=stTaking";
    var popwin = window.open(popUrl,'popBrandStockInput','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popOffOutItemList(makerid, shopid, yyyy, mm){
    var popUrl = "/common/offshop/stock/OutItemList.asp?makerid="+makerid+"&shopid="+shopid+"&yyyy="+yyyy+"&mm="+mm;
    var popwin = window.open(popUrl,'popItemStockTurnOver','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popOffErrInput(shopid,itemgubun,itemid,itemoption){
    
        var popwin = window.open('/common/offshop/popOffrealerrinput.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popAdmOffrealerrinput','width=900,height=460,scrollbars=yes,resizable=yes');
	    popwin.focus();
	
}

function popShopCurrentStock(shopid,itemgubun,itemid,itemoption){
    var popwin = window.open('/common/offshop/shop_itemcurrentstock.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popShopCurrentStock','width=900,height=600,resizable=yes,scrollbars=yes');
    popwin.focus();
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
		    대상 년/월 <% CALL DrawYMBox(yyyy1,mm1) %>
		    &nbsp;
		    <% if (C_IS_SHOP) then %>
    		    <input type="hidden" name="shopid" value="<%= shopid %>">
    		    매장 : <%= shopid %>
		    <% elseif (C_IS_Maker_Upche) then %>
    		    <!-- 계약된 업체 -->
    		    매장 : <% drawSelectBoxOpenOffShop "shopid",shopid %>
		    <% else %>
		        매장 : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
		    <% end if %>
		    
		    <% if (C_IS_Maker_Upche) then %>
		        <input type="hidden" name="makerid" value="<%= makerid %>">
		    <% else %>
    			브랜드 :
    			<% drawSelectBoxDesignerwithName "makerid", makerid %> &nbsp;&nbsp;
			<% end if %>
			<br>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<!-- 상품 사용구분 : <% drawSelectBoxUsingYN "usingyn", usingyn %> &nbsp;&nbsp; -->
			
             &nbsp;&nbsp;
             검색 구분 : 
             <input type="radio" name="sMode" <%= CHKIIF(sMode="X","checked","") %> value="X"> 정리대상 상품( 재고>0 이고 판매량 <1 인 상품 )
             <input type="radio" name="sMode" <%= CHKIIF(sMode="S","checked","") %> value="S"> 재고>0인 상품
             <input type="radio" name="sMode" <%= CHKIIF(sMode="","checked","") %> value=""> 전체
		</td>
	</tr>
	
	</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" > 
    <tr height="60">
        <td>
        * 마이너스 재고는 재고 0 으로 산정함.<br>
        </td>
    </tr>
	<tr height="30">
		<td align="left">
			검색결과 총 <%= oOffOutItem.FTotalCount %> 건
		</td>
	</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="60">년/월</td>
        <td width="110">이미지</td>
    	<td width="100">상품코드</td>
    	<td >상품명 <font color="blue">[옵션]</font></td>
    	<td width="60">재고수량</td>
    	<td width="70"><%= mm1 %>월 <br>판매수량</td>
    	<td width="70"><%= mm1 %>월 <br>매출액</td>
        <td >비고</td>
    </tr>
    <% if (shopid="") then %>
    <tr align="center" bgcolor="#FFFFFF" height="30">
        <td colspan="14">[먼저 <Strong>매장</Strong> 을 선택하세요.]</td>
    </tr>
    <% else %>
    <% for i=0 to oOffOutItem.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= oOffOutItem.FItemList(i).Fyyyymm %></td>
        <td height="50">
            <% if (oOffOutItem.FItemList(i).IsImageExists) then %>
            <img src="<%= oOffOutItem.FItemList(i).GetImageSmall %>" width="50">
            <% else %>
            <img src="http://webimage.10x10.co.kr/images/no_image.gif" width="50">
            <% end if %>
        </td>
        <td><a href="javascript:popShopCurrentStock('<%= shopid %>','<%= oOffOutItem.FItemList(i).Fitemgubun %>','<%= oOffOutItem.FItemList(i).Fitemid %>','<%= oOffOutItem.FItemList(i).Fitemoption %>')"><%= oOffOutItem.FItemList(i).getTenBarCode %></a></td>
        <td><%= oOffOutItem.FItemList(i).FItemName %>
        <% if oOffOutItem.FItemList(i).FItemOptionName<>"" then %>
        <font color="blue">[<%= oOffOutItem.FItemList(i).FItemOptionName %>]</font>
        <% end if %>
        </td>
        <td><%= oOffOutItem.FItemList(i).Fstockno %></td>
        <td><%= oOffOutItem.FItemList(i).FtotSellno %></td>
        <td><%= FormatNumber(oOffOutItem.FItemList(i).FtotRealSellPrice,0) %></td>
        <td>
        <input type="button" class="button" value="실사" onclick="popOffErrInput('<%= oOffOutItem.FItemList(i).Fshopid %>','<%= oOffOutItem.FItemList(i).FitemGubun %>','<%= oOffOutItem.FItemList(i).FitemID %>','<%= oOffOutItem.FItemList(i).FitemOption %>');">    
        </td>
    </tr>
    <% next %>
    <% end if %>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4">
			<% if oOffOutItem.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oOffOutItem.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + oOffOutItem.StartScrollPage to oOffOutItem.FScrollCount + oOffOutItem.StartScrollPage - 1 %>
				<% if i>oOffOutItem.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if oOffOutItem.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
<%
set oOffOutItem = Nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" --> 
