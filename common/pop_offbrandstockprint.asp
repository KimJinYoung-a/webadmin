<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<%
dim shopid, makerid, centermwdiv, itembarcode, usingyn, research, NoZeroStock
dim itemgubun, itemid, itemoption
dim ImgUsing, pagesize

shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
centermwdiv  = RequestCheckVar(request("centermwdiv"),10)
itembarcode  = RequestCheckVar(request("itembarcode"),32)
usingyn      = RequestCheckVar(request("usingyn"),1)
research     = RequestCheckVar(request("research"),2)
ImgUsing     = RequestCheckVar(request("ImgUsing"),1)
NoZeroStock  = RequestCheckVar(request("NoZeroStock"),32)
pagesize  		= RequestCheckVar(request("pagesize"),32)

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

if (itembarcode <> "") then
    if Not (fnGetItemCodeByPublicBarcode(itembarcode,itemgubun,itemid,itemoption)) then
        itemgubun   = Left(itembarcode, 2)
        itemid      = CStr(Mid(itembarcode, 3, 6) + 0)
        itemoption  = Right(itembarcode, 4)
    end if
end if

'''if (research="") and (usingyn="") then usingyn="Y"
if (research="") and (ImgUsing="") then ImgUsing="Y"
if (pagesize = "") then pagesize = "100"


dim oOffStock
set oOffStock = new CShopItemSummary
oOffStock.FCurrPage 		= 1
oOffStock.FPageSize 		= pagesize
oOffStock.FRectShopID       = shopid
oOffStock.FRectMakerID      = makerid
oOffStock.FRectCenterMwDiv  = centermwdiv
oOffStock.FRectIsUsing      = usingyn
oOffStock.FRectNoZeroStock  = NoZeroStock
if (itembarcode <> "") then
    oOffStock.FRectItemGubun    = itemgubun
    oOffStock.FRectItemId       = itemid
    oOffStock.FRectItemOption   = itemoption
end if

if ((shopid<>"") and (makerid<>"")) or ((shopid<>"") and (itembarcode<>"")) then
    oOffStock.GetShopItemCurrentSummaryList
end if

dim i
dim totsysstock, totavailstock, totrealstock
%>
<script language='javascript'>
function RefreshPageByImg(ImgUsing){
    document.frm.ImgUsing.value = ImgUsing;
    document.frm.submit();
}
function RefreshPage(){
    document.frm.submit();
}
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="ImgUsing" value="<%=ImgUsing%>">

	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
		    <% if (C_IS_SHOP) then %>
		    <input type="hidden" name="shopid" value="shopid">
            매장 : <%= shopid %>
            <% else %>
		    매장 : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
		    <% end if %>
			브랜드 :
			<% drawSelectBoxDesignerwithName "makerid", makerid %> &nbsp;&nbsp;
			&nbsp;
			표시갯수 :
			<select class="select" name="pagesize">
				<option value="100" <%= CHKIIF(pagesize = "100", "selected", "") %>>100</option>
				<option value="500" <%= CHKIIF(pagesize = "500", "selected", "") %>>500</option>
				<option value="1000" <%= CHKIIF(pagesize = "1000", "selected", "") %>>1000</option>
				<option value="2000" <%= CHKIIF(pagesize = "2000", "selected", "") %>>2000</option>
			</select>
		</td>

		<td rowspan="2" width="220" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" name="brandstockprint" value="재검색" onclick="RefreshPage();">
			<% if ImgUsing="Y" then %>
        		<input type="button" class="button" name="brandstockprint" value="이미지없애기" onclick="RefreshPageByImg('N');">
        	<% else %>
        		<input type="button" class="button" name="brandstockprint" value="이미지보이기" onclick="RefreshPageByImg('Y');">
        	<% end if %>
        	<input type="button" class="button" name="brandstockprint" value="출력하기" onclick="window.print();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			상품 사용구분 : <% drawSelectBoxUsingYN "usingyn", usingyn %> &nbsp;&nbsp;
			<input type="checkbox" name="NoZeroStock" <%= CHKIIF(NoZeroStock="on","checked","") %> > 재고0인 상품 검색 안함.
			<!--
			센터매입구분 :
			   <select class="select" name="centermwdiv">
               <option value="">전체</option>
               <option value="MW" <%= ChkIIF(centermwdiv="MW","selected","") %> >매입+위탁</option>
               <option value="W"  <%= ChkIIF(centermwdiv="W","selected","") %> >위탁</option>
               <option value="M"  <%= ChkIIF(centermwdiv="M","selected","") %> >매입</option>
               <option value="NULL" <%= ChkIIF(centermwdiv="NULL","selected","") %> >미지정</option>
               </select>
            &nbsp;&nbsp;
            -->
            [출력일 : <%= now() %>]
		</td>
	</tr>

	</form>
</table>
<!-- 검색 끝 -->
<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="30">구분</td>
    	<td width="40">상품ID</td>
    	<td width="40">옵션</td>
    	<% if (ImgUsing="Y") then %>
    	<td width="50">이미지</td>
    	<% end if %>
    	<td>상품명<br>[옵션명]</td>
    	<!-- td width="40">센터<br>매입<br>구분</td -->
    	<td width="40">센터<br>입고</td>
    	<td width="40">센터<br>반품</td>
    	<td width="40">브랜드<br>입고</td>
    	<td width="40">브랜드<br>반품</td>
        <td width="40">매장<br>판매</td>
        <td width="40">매장<br>반품</td>
        <td width="40" bgcolor="F4F4F4">시스템<br>총재고</td>
        <td width="40">총<br>실사<br>오차</td>
        <td width="40" bgcolor="F4F4F4">실사<br>재고</td>
		<td width="40">배송중</td>
		<td width="40">반품중</td>
		<td width="40" bgcolor="F4F4F4">매장<br>재고<br>(현재)</td>
        <!-- <td width="40">총<br>샘플</td>
        <td width="40">총<br>불량</td>-->
        <!-- <td width="40" bgcolor="F4F4F4">유효<br>재고</td-->

        <td width="30">사용<br>여부</td>
        <td width="100">비고</td>
    </tr>
<% if oOffStock.FResultCount<1 then %>
    <tr align="center" bgcolor="#FFFFFF" height="30">
        <% if (shopid="") and (makerid="") then %>
        <td colspan="20" >[ 매장 및 브랜드를 선택 하세요. ]</td>
        <% else %>
        <td colspan="20" >[ 검색 결과가 없습니다. ]</td>
        <% end if %>
    </tr>
<% else %>
    <% for i=0 to oOffStock.FResultCount - 1 %>
    <%
    totsysstock	    = totsysstock + oOffStock.FItemList(i).FsysstockNo
    totavailstock   = totavailstock + oOffStock.FItemList(i).getAvailStock
    totrealstock    = totrealstock + oOffStock.FItemList(i).FrealstockNo

    %>
    	<% if oOffStock.FItemList(i).Fisusing="Y" then %>
        <tr bgcolor="#FFFFFF" align="center" >
        <% else %>
        <tr bgcolor="#FFFFFF" align="center" >
        <% end if %>
            <td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).FItemGubun %></td>
        	<td style="border-bottom:1px solid black">
        	    <%= oOffStock.FItemList(i).Fitemid %>
        	</td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).FItemOption %></td>
        	<% if (ImgUsing="Y") then %>
        	<td style="border-bottom:1px solid black"><img src="<%= oOffStock.FItemList(i).GetImageSmall %>" width=50 height=50> </td>
        	<% end if %>
        	<td align="left" style="border-bottom:1px solid black">
              	<%= oOffStock.FItemList(i).FShopitemname %>
              	<% if oOffStock.FItemList(i).FShopitemoptionName <>"" then %>
              		<br>
              		<font color="blue">[<%= oOffStock.FItemList(i).FShopitemoptionName %>]</font>
              	<% end if %>
            </td>
            <!-- td><%= fnColor(oOffStock.FItemList(i).FCenterMwdiv,"mw") %></td -->
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Flogicsipgono %></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Flogicsreipgono %></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Fbrandipgono %></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Fbrandreipgono %></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Fsellno %></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Fresellno %></td>
        	<td bgcolor="F4F4F4" style="border-bottom:1px solid black"><b><%= oOffStock.FItemList(i).FsysstockNo %></b></td>
        	<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Ferrrealcheckno %></td>
        	<td bgcolor="F4F4F4" style="border-bottom:1px solid black"><b><%= oOffStock.FItemList(i).FrealstockNo %></b></td>
			<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Flogischulgo %></td>
			<td style="border-bottom:1px solid black"><%= oOffStock.FItemList(i).Flogisreturn %></td>
			<td bgcolor="F4F4F4" style="border-bottom:1px solid black"><b><%= oOffStock.FItemList(i).getShopRealStockNoExc %></b></td>
        	<!-- td><%= oOffStock.FItemList(i).Ferrsampleitemno %></td>
        	<td><%= oOffStock.FItemList(i).Ferrbaditemno %></td> -->
        	<!-- td bgcolor="F4F4F4"><b><%= oOffStock.FItemList(i).getAvailStock %></b></td -->

        	<td style="border-bottom:1px solid black">
        	    <% if oOffStock.FItemList(i).Fisusing="N" then %>
        	    <strong><%= oOffStock.FItemList(i).Fisusing %></strong>
        	    <% else %>
        	    <%= oOffStock.FItemList(i).Fisusing %>
        	    <% end if %>
        	</td>
        	<td valign="top" style="border-bottom:1px solid black">
        	<% if (oOffStock.FItemList(i).Ferrsampleitemno<>0) then %>
        	(샘플 <%= oOffStock.FItemList(i).Ferrsampleitemno*-1 %>)
        	<% else %>
        	&nbsp;
        	<% end if %>
        	</td>

        </tr>
    <% next %>
<% end if %>
</table>

<%
set oOffStock = Nothing
%>






<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
