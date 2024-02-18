<%@ language=vbscript %>
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
<!-- #include virtual="/lib/classes/offshop/stock/shopstockClearCls.asp"-->

<%

'response.write "수정중"
'response.end

dim shopid, makerid, research, yyyy1, mm1
dim dispDiv

shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
research     = RequestCheckVar(request("research"),32)

dim usingyn, centermwdiv ,NoZeroStock, comm_cd
usingyn      = RequestCheckVar(request("usingyn"),32)
centermwdiv  = RequestCheckVar(request("centermwdiv"),32)
NoZeroStock  = RequestCheckVar(request("NoZeroStock"),32)
comm_cd      = RequestCheckVar(request("comm_cd"),32)

yyyy1         = RequestCheckVar(request("yyyy1"),4)
mm1           = RequestCheckVar(request("mm1"),2)

dispDiv           = RequestCheckVar(request("dispDiv"),2)


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
set oOffOutItem = new CShopStockClear
oOffOutItem.FRectShopID		= shopid
oOffOutItem.FRectMakerID	= makerid
oOffOutItem.FRectCommCD		= comm_cd
oOffOutItem.FRectDispDiv	= dispDiv
''oOffOutItem.FRectYYYYMM	= yyyy1 + "-" + mm1


if (shopid<>"") or (makerid<>"") then
    oOffOutItem.GetShopStockClearBrandList
end if

dim i
%>
<script language='javascript'>

function popOffOutItemList(makerid, shopid, cType){
    var popUrl = "/admin/offshop/stock/OutItemListByBrand.asp?makerid="+makerid+"&shopid="+shopid+"&cType="+cType;
    var popwin = window.open(popUrl,'OutItemListByBrand','width=1400,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popMonthlyStock(makerid, shopid, yyyymm){
    var popUrl = "/admin/newreport/monthlystockShop_detail.asp?menupos=1346&showminus=on&sysorreal=sys&makerid="+makerid+"&shopid="+shopid+"&yyyymm="+yyyymm;
    var popwin = window.open(popUrl,'popMonthlyStock','width=1100,height=800,scrollbars=yes,resizable=yes');
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
		    <!-- 현재고 기준으로 변경
		    대상 년/월
		    <input type="text" name="yyyy1" value="<%= yyyy1 %>" size="4" readOnly class="text_ro">
		    <input type="text" name="mm1" value="<%= mm1 %>" size="2" readOnly class="text_ro">
		    &nbsp;
		    -->
		    <% if (C_IS_SHOP) then %>
    		    <input type="hidden" name="shopid" value="<%= shopid %>">
    		    매장 : <%= shopid %>
		    <% elseif (C_IS_Maker_Upche) then %>
    		    <!-- 계약된 업체 -->
    		    매장 : <% drawSelectBoxOpenOffShop "shopid",shopid %>
		    <% else %>
		        매장 : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp; <!-- drawSelectBoxOffShop -->
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

		    매입구분 :
		    <% drawSelectBoxOFFJungsanCommCD "comm_cd",comm_cd %>

            &nbsp;&nbsp;
			표시구분 :
			<select class="select" name="dispDiv">
				<option value="SY" <%if (dispDiv = "SY") then %>selected<% end if %> >SYS재고</option>
				<option value="ER" <%if (dispDiv = "ER") then %>selected<% end if %> >오차재고</option>
				<option value="SM" <%if (dispDiv = "SM") then %>selected<% end if %> >샘플재고</option>
			</select>
			있는 브랜드만
            <!--
             <input type="checkbox" name="NoZeroStock" <%= CHKIIF(NoZeroStock="on","checked","") %> > 재고0인 브랜드 검색 안함.
             -->
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
        * 재고>0 이고 판매량 <1 인 상품
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
        <td width="150">브랜드ID</td>
    	<td width="90">(현)매입구분</td>
    	<td width="70">시스템재고<br>존재상품</td>
    	<td width="70">시스템재고</td>
		<td width="70">오차</td>
    	<td width="70"><b>실사재고</b></td>
		<td width="70">샘플</td>
		<td width="80"><b>유효재고<br>(실사+샘플)</b></td>
    	<td width="70">2개월 <br>판매수량</td>
    	<td width="70">2개월 <br>매출액</td>
    	<td width="70">최초 <br>입고일</td>
    	<td width="70">최종 <br>입고일</td>
    	<!--
    	<td width="90">정리대상<br>상품수</td>
    	<td width="90">정리대상<br>상품재고</td>
    	-->
        <td >비고</td>
    </tr>
    <% if (shopid="") then %>
    <tr align="center" bgcolor="#FFFFFF" height="30">
        <td colspan="13">[먼저 <Strong>매장</Strong> 을 선택하세요.]</td>
    </tr>
    <% else %>
    <% for i=0 to oOffOutItem.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= oOffOutItem.FItemList(i).Fmakerid %></td>
        <td><%= oOffOutItem.FItemList(i).Fcomm_name %></td>
        <td><a href="javascript:popMonthlyStock('<%= oOffOutItem.FItemList(i).Fmakerid %>','<%= shopid %>','<%= yyyy1 %>','<%= mm1 %>');"><%= oOffOutItem.FItemList(i).FItemCnt %></a></td>
        <td><%= oOffOutItem.FItemList(i).Ftotsysstockno %></td>
        <td><%= oOffOutItem.FItemList(i).Ftoterrrealcheckno %></td>
		<td><b><%= oOffOutItem.FItemList(i).Ftotrealstockno %></b></td>
		<td><%= oOffOutItem.FItemList(i).Ftoterrsampleitemno %></td>
		<td><b><%= (oOffOutItem.FItemList(i).Ftotrealstockno + oOffOutItem.FItemList(i).Ftoterrsampleitemno) %></b></td>
        <td><%= oOffOutItem.FItemList(i).FtotSellNo %></td>
        <td><%= FormatNumber(oOffOutItem.FItemList(i).FtotRealSellPrice,0) %></td>
        <td>
            <% if IsRecentIpchul(oOffOutItem.FItemList(i).Ffirstipgodate) then %>
            <b><%= oOffOutItem.FItemList(i).Ffirstipgodate %></b>
            <% else %>
            <%= oOffOutItem.FItemList(i).Ffirstipgodate %>
            <% end if %>
        </td>
        <td>
            <% if IsRecentIpchul(oOffOutItem.FItemList(i).Flastipgodate) then %>
            <b><%= oOffOutItem.FItemList(i).Flastipgodate %></b>
            <% else %>
            <%= oOffOutItem.FItemList(i).Flastipgodate %>
            <% end if %>
        </td>
        <td>
            <!--
            <input type="button" value="재고조정(임시)" class="button" onClick="popOffOutItemList('<%= oOffOutItem.FItemList(i).Fmakerid %>','<%= shopid %>','C')">
            &nbsp;

            <input type="button" value="오차조정" class="button" onClick="popOffOutItemList('<%= oOffOutItem.FItemList(i).Fmakerid %>','<%= shopid %>','M')">
            &nbsp;
            -->
            <input type="button" value="오차 로스처리" class="button" onClick="popOffOutItemList('<%= oOffOutItem.FItemList(i).Fmakerid %>','<%= shopid %>','L')">
			<input type="button" value="샘플 출고" class="button" onClick="popOffOutItemList('<%= oOffOutItem.FItemList(i).Fmakerid %>','<%= shopid %>','S')">
        </td>
    </tr>
    <% next %>
    <% end if %>
</table>
<%
set oOffOutItem = Nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
