<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인매장 브랜드별 재고 파악 (배치)
' History : 2011.08
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<!-- #include virtual="/lib/classes/stock/shopbatchstockcls.asp"-->

<%
dim idx : idx = RequestCheckVar(request("idx"),32)
dim oshopBatch, shopid, jobkey, joborderno, StockDate
set oshopBatch = new CShopOrder
	oshopBatch.FRectidx=idx
	
	if (idx<>"") then
	    oshopBatch.GetOneShopBatchOrder
	end if

if (oshopBatch.FResultCount>0) then
    shopid = oshopBatch.FOneItem.Fjobshopid
    jobkey = oshopBatch.FOneItem.Fjobkey
    joborderno = oshopBatch.FOneItem.Forderno
    StockDate = Left(oshopBatch.FOneItem.FShopRegDate,10)
end if 

set oshopBatch= Nothing

'', research,NoZeroStock,centermwdiv
''dim makerid : makerid      = RequestCheckVar(request("makerid"),32)
''research     = RequestCheckVar(request("research"),32)

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if
''if (research="") then NoZeroStock="on"

dim oOffStock
set oOffStock = new CShopItemSummary
oOffStock.FRectShopID = shopid
oOffStock.FRectBatchIdx = idx
if (shopid<>"") then
    oOffStock.GetShopBrandBatchCheckList
end if

dim i
%>
<script language='javascript'>
function popBrandStock(shopid,makerid){
    var popUrl = "/common/offshop/shop_brandcurrentstock_byJobKey.asp?menupos=1074&shopid="+shopid+"&makerid="+makerid+"&idx=<%= idx %>";
    var popwin = window.open(popUrl,'popBrandStock','scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="idx" value="<%= idx %>">
	
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
		    <% if (C_IS_SHOP) then %>
    		    <input type="hidden" name="shopid" value="<%= shopid %>">
    		    매장 : <%= shopid %>
		    <% elseif (C_IS_Maker_Upche) then %>
    		    <!-- 계약된 업체 -->
    		    매장 : <% drawSelectBoxOpenOffShop "shopid",shopid %>
		    <% else %>
		        매장 : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
		    <% end if %>
		    
			<br>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    작업번호 : <%= jobkey %> (<%= joborderno %>) 
		    재고일시 : <%= StockDate %>
		    
			
		</td>
	</tr>
	
	</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" > 
    <tr height="30">
        <td>
        * 마이너스 재고는 재고금액 0 으로 산정함.    
        </td>
    </tr>
	<tr height="30">
		<td align="left">
			검색결과 총 <%= oOffStock.FTotalCount %> 건
		</td>
	</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="20"></td>
        <td width="120">브랜드ID</td>
    	<td width="100">매입구분</td>
    	<td width="100">총 시스템상<br>실사 재고</td>
    	<td width="100">재고 파악 실사재고</td>
    	<td width="100">최초 입고일</td>
        <td >최근실사일</td>
        <td >브랜드실사</td>
    </tr>
    <% if (shopid="") then %>
    <tr align="center" bgcolor="#FFFFFF" height="30">
        <td colspan="10">[먼저 매장 을 선택하세요.]</td>
    </tr>
    <% else %>
    <% for i=0 to oOffStock.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td></td>
        <td><%= oOffStock.FItemList(i).Fmakerid %></td>
        <td><%= oOffStock.FItemList(i).Fcomm_name %></td>
        <td><%= FormatNumber(oOffStock.FItemList(i).FtotSysRealStockNo,0) %></td>
        <td><%= FormatNumber(oOffStock.FItemList(i).FtotRealStockNo,0) %></td>
        
        <td><%= oOffStock.FItemList(i).Ffirstipgodate %></td>
        <td><%= oOffStock.FItemList(i).FlastStockdate %></td>
        <td>
        <input type="button" class="button" value="브랜드 실사 입력" onClick="popBrandStock('<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>');">    
        </td>
    </tr>
    <% next %>
    <% end if %>
</table>
<%
set oOffStock = Nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" --> 