<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itembarcodecls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim itembarcode
dim itemgubun,itemid,itemoption
dim BasicMonth

itembarcode = requestCheckVar(request("itembarcode"),20)
itemgubun 	= requestCheckVar(request("itemgubun"),2)
itemid		 = requestCheckVar(request("itemid"),10)
itemoption	 = requestCheckVar(request("itemoption"),4)
BasicMonth	 = requestCheckVar(request("BasicMonth"),10)

if (Len(itembarcode)=12) then
	itemgubun 	= left(itembarcode,2)
	itemid		= CLng(mid(itembarcode,3,6))
	if (itemoption="") then itemoption = right(itembarcode,4)
	itembarcode = itemgubun + Format00(6,itemid) + itemoption
elseif (Len(itembarcode)=14) then
	itemgubun 	= left(itembarcode,2)
	itemid		= CLng(mid(itembarcode,3,8))
	if (itemoption="") then itemoption = right(itembarcode,4)
	itembarcode = itemgubun + Format00(8,itemid) + itemoption
elseif (Len(itembarcode)<>0) and (itemid<>"") then
	if itemgubun="" then itemgubun = "10"
	itemid = itembarcode
	if (itemoption="") then itemoption  = "0000"
elseif (Len(itembarcode)>7) then
    '''바코드인경우 검색후 상품코드 가져옴.
    call fnGetItemCodeByPublicBarcode(itembarcode, itemgubun, itemid, itemoption)
else
    if itemgubun="" then itemgubun = "10"
    if (itemid="") then itemid = itembarcode
    if (itemoption="") then itemoption  = "0000"

    if (itemid>=1000000) then
        itembarcode = itemgubun + Format00(8,itemid) + itemoption
    else
        itembarcode = itemgubun + Format00(6,itemid) + itemoption
    end if
end if



dim oitem
set oitem = new CItem
oitem.FRectItemID = itemid

if (itemid<>"") and (itemgubun="10") then
	oitem.GetOneItem
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if (itemid<>"")  and (itemgubun="10") then
	oitemoption.GetItemOptionInfo
end if

''오프상품
dim ooffitem
set ooffitem = new COffShopItem
ooffitem.FRectItemGubun = itemgubun
ooffitem.FRectItemID    = itemid
ooffitem.FRectItemOption = itemoption
if (itemgubun<>"") and (itemid<>"") and (itemoption<>"") and (itemgubun<>"10") then
	ooffitem.GetOffOneItem
end if

dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectStartDate = BasicMonth + "-01"
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID    =  itemid
osummarystock.FRectItemOption =  itemoption
if itemid<>"" then
	osummarystock.GetCurrentItemStock
end if


dim otodayerritem
set otodayerritem = new CSummaryItemStock
otodayerritem.FRectItemGubun = itemgubun
otodayerritem.FRectItemID =  itemid
otodayerritem.FRectItemOption =  itemoption
if itemid<>"" then
    otodayerritem.GetTodayErrItem
end if

dim difftime
if (osummarystock.FResultcount>0) then
    difftime = ABS(datediff("h",osummarystock.FOneItem.Flastupdate,now()))
end if

dim i
dim IsVaildCode, IsStockExists
IsVaildCode = False
if (oitemoption.FResultCount>0) then
    for i=0 to oitemoption.FResultCount-1
        if (oitemoption.FITemList(i).FItemOption=itemoption) then
            IsVaildCode = (oitem.FResultCount>0)
            exit For
        end if
    next
else
    IsVaildCode = ((oitem.FResultCount>0) and (itemoption="0000")) or (ooffitem.FResultCount>0)
end if

IsStockExists = (osummarystock.FResultCount>0)
%>
<script language='javascript'>


function RecalcuErr(){
	var checkstock = calcufrm.checkstock.value;  // 재고파악재고.

	calcufrm.todayerrrealcheckno.value = checkstock-calcufrm.orgrealstock.value - calcufrm.todaybaljuno.value;
	calcufrm.errrealcheckno.value = checkstock - calcufrm.availsysstock.value - calcufrm.todaybaljuno.value;
}

function SaveErr(){
//	if (<%= difftime %>>=4){
//		alert('최종 업데이트시간이 4시간 이후 입니다. \n먼저 새로고침후 사용하세요.');
//		return;
//	}

	var realstock = calcufrm.checkstock.value;
	if (!IsInteger(realstock)){
		alert('숫자를 입력하세요.');
		calcufrm.checkstock.focus();
		return;
	}

	if (confirm('실사오차를 저장하시겠습니까?')){
		frmrefresh.mode.value ="errcheckupdate";
		frmrefresh.realstock.value = realstock;
		frmrefresh.submit();
	}
}

function GetOnLoad(){
	<% if Not IsVaildCode then %>
	alert('상품코드가 정확하지 않습니다. 재검색 하세요.');
	document.frm.itembarcode.select();
	document.frm.itembarcode.focus();
	<% else %>
	if (calcufrm.checkstock){
	    document.calcufrm.checkstock.select();
	    document.calcufrm.checkstock.focus();
	}
	<% end if %>
}
window.onload=GetOnLoad;

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type=hidden name=BasicMonth value="<%= BasicMonth %>">
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red">&nbsp;<strong>재고(오차)입력</strong></font>
				    </td>
				    <td align="right">
						상품코드:
						<input type=text class="text" name=itembarcode value="<%= itembarcode %>" size=16 maxlength=16 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ document.frm.submit(); return false;}">
						<!--
			        	<input type=text class="text_ro" name=itemgubun value="<%= itemgubun %>" size=2 maxlength=2 readonly>
			        	<input type=text class="text" name=itemid value="<%= itemid %>" size=9 maxlength=9>
			        	<input type="text" class="text_ro" value="<%= itemoption %>" size=4 maxlength=4 readonly>
			        	-->
						&nbsp;

						<% if oitemoption.FResultCount>0 then %>

						<select class="select" name="itemoption">
						<option value="0000">----
						<% for i=0 to oitemoption.FResultCount-1 %>
						<option value="<%= oitemoption.FITemList(i).FItemOption %>" <% if itemoption=oitemoption.FITemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FITemList(i).FOptionName %>
						<% next %>
						</select>
						<% end if %>

        				<input type="button" class="button" value="검색" onclick="document.frm.submit();">
        				<!-- 최근 내역 새로고침 후 입력됨
				        <%= BasicMonth %>-01 ~
				        <input type="button" value="새로고침" onclick="RefreshRecentStock();">
				        -->
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- 상단바 끝 -->
	</form>
</table>

<p>

<% if (oitem.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 6 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
      	<td width="60">상품코드</td>
      	<td width="300">
      		10<Strong><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></Strong><%= itemoption %>
      	</td>
      	<td width="60"></td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>브랜드ID</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td>판매여부</td>
      	<td colspan=2><%= fnColor(oitem.FOneItem.FSellyn,"yn") %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>상품명</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td>사용여부</td>
      	<td colspan=2><%= fnColor(oitem.FOneItem.FIsUsing,"yn") %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>판매가</td>
      	<td>
      		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
      		&nbsp;&nbsp;
      		<font color="<%= mwdivColor(oitem.FOneItem.FMwDiv) %>"><%= oitem.FOneItem.getMwDivName %></font>
      	    <% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
			<% end if %>
			&nbsp;&nbsp;
			<!-- 할인여부/쿠폰적용여부 -->
			<% if (oitem.FOneItem.FSailYn="Y") then %>
			    <font color=red>
			    <% if (oitem.FOneItem.Forgprice<>0) then %>
			        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
			    <% end if %>
			     할인
			    </font>
			<% end if %>

			<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

			    <font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> 쿠폰
			    (<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
			<% end if %>

      	</td>
      	<td>단종여부</td>
      	<td colspan=2>
      		<% if oitem.FOneItem.Fdanjongyn="Y" then %>
			<font color="#33CC33">단종</font>
			<% elseif oitem.FOneItem.Fdanjongyn="M" then %>
			<font color="#CC3333">MD품절</font>
			<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
			<font color="#33CC33">일시품절</font>
			<% else %>
			생산중
			<% end if %>
		</td>
    </tr>
     <% if oitemoption.FResultCount>1 then %>
	    <!-- 옵션이 있는경우 -->
	    <% for i=0 to oitemoption.FResultCount -1 %>
	    	<% if oitemoption.FITemList(i).Fitemoption=itemoption then %>
	    	<tr bgcolor="#FFFFFF">
	    		<td>옵션명</td>
		      	<td><%= oitemoption.FITemList(i).FOptionName %> (<%= fnColor(oitemoption.FITemList(i).FOptIsUsing,"yn") %>)</td>
		      	<td>한정여부</td>
		      	<td><%= fnColor(oitem.FOneItem.Flimityn,"yn") %> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
		      	<td>한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% end if %>
		<% next %>
	<% else %>
    	<tr bgcolor="#FFFFFF">
	      	<td>옵션명</td>
	      	<td>-</td>
	      	<td>한정여부</td>
	      	<td><%= fnColor(oitem.FOneItem.Flimityn,"yn") %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
	      	<td>한정 비교재고 (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
	    </tr>
    <% end if %>

</table>

<% elseif (ooffitem.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=5 width="110" valign=top align=center><img src="<%= ooffitem.FOneItem.FOffimgList %>" width="100" height="100"></td>
      	<td width="60">상품코드</td>
      	<td width="300">
      		<%= ooffitem.FOneItem.GetBarCode %>
      	</td>
      	<td width="60"></td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>브랜드ID</td>
      	<td><%= ooffitem.FOneItem.FMakerid %></td>
      	<td></td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>상품명</td>
      	<td><%= ooffitem.FOneItem.Fshopitemname %></td>
      	<td>사용여부</td>
      	<td colspan=2><%= fnColor(ooffitem.FOneItem.FIsUsing,"yn") %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>판매가</td>
      	<td>
      		<%= FormatNumber(ooffitem.FOneItem.Fshopitemprice,0) %>
      		<!--
      		/ <%= FormatNumber(ooffitem.FOneItem.Fshopsuplycash,0) %>
      		-->
      		&nbsp;&nbsp;
      		<!--
      	    <% if ooffitem.FOneItem.Fshopitemprice<>0 then %>
			<%= CLng((1- ooffitem.FOneItem.Fshopsuplycash/ooffitem.FOneItem.Fshopitemprice)*100) %> %
			<% end if %>

			&nbsp;&nbsp;
			-->
			<!-- 할인여부/쿠폰적용여부 -->
			<% if (ooffitem.FOneItem.FShopItemOrgprice>ooffitem.FOneItem.Fshopitemprice) then %>
			    <font color=red>
			    <% if (ooffitem.FOneItem.FShopItemOrgprice<>0) then %>
			        <%= CLng((ooffitem.FOneItem.FShopItemOrgprice-ooffitem.FOneItem.Fshopitemprice)/ooffitem.FOneItem.FShopItemOrgprice*100) %> %
			    <% end if %>
			     할인
			    </font>
			<% end if %>


      	</td>
      	<td>단종여부</td>
      	<td colspan=2>

		</td>
    </tr>

    	<tr bgcolor="#FFFFFF">
	      	<td>옵션명</td>
	      	<td><%= ooffitem.FOneItem.Fshopitemoptionname %></td>
	      	<td></td>
	      	<td></td>
	      	<td></td>
	    </tr>

</table>
<% else %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
        <td align="center">[검색 결과가 없습니다.]</td>
    </tr>
</table>
<% end if %>
<p>
<% if osummarystock.FResultCount>0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name=calcufrm >
	<input type="hidden" name="orgrealstock" value="<%= osummarystock.FOneItem.Frealstock %>">
	<input type="hidden" name="orgerrrealcheckno" value="<%= osummarystock.FOneItem.Ferrrealcheckno %>">
	<input type="hidden" name="availsysstock" value="<%= osummarystock.FOneItem.Favailsysstock %>">
	<input type="hidden" name="todaybaljuno" value="<%= osummarystock.FOneItem.GetTodayBaljuNo %>">
	<input type="hidden" name="todayinputedrealcheckerrno" value="<%= otodayerritem.FOneItem.Ferrrealcheckno %>">

<!-- 실시간 업데이트 됨
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td colspan="16" align=right>최종업데이트 : <%= osummarystock.FOneItem.Flastupdate %> </td>
    </tr>
-->
    <tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50">총입고/반품</td>
    	<td width="50">총판매/반품</td>
		<td width="50">샾출고/반품</td>
		<td width="50">기타출고/반품</td>
		<td width="50">CS<br>출고/반품</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">시스템<br>총재고</td>
		<td width="50">총실사<br>오차</td>
		<td width="50">총불량</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">실사<br>유효재고</td>
		<td width="50">ON상품<br>준비</td>
		<td width="50">OFF상품<br>준비</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">재고파악<br>재고</td>
		<td width="50">ON결제<br>완료</td>
		<td width="50">ON주문<br>접수</td>
		<td width="50">OFF주문<br>접수</td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ftotipgono %></td>
    	<td rowspan="2"><%= -1*osummarystock.FOneItem.Ftotsellno %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Foffchulgono + osummarystock.FOneItem.Foffrechulgono %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Fetcchulgono + osummarystock.FOneItem.Fetcrechulgono %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ferrcsno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.Ftotsysstock %></td>
    	<td><%= osummarystock.FOneItem.Ferrrealcheckno %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ferrbaditemno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.Frealstock %></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv5 %></td>
    	<td><%= osummarystock.FOneItem.Foffconfirmno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><input type="text" name="checkstock" value="<%= osummarystock.FOneItem.GetCheckStockNo %>" size="4" maxlength="7" style="border:1px #999999 solid; text-align=center" onKeyUp="RecalcuErr();"></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv4 %></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv2 %></td>
    	<td><%= osummarystock.FOneItem.Foffjupno %></td>

    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td ><input type="text" name="errrealcheckno" value="<%= osummarystock.FOneItem.Ferrrealcheckno  %>"  size="4" maxlength="7" readonly style="background:#CCCCCC; border:1px #999999 solid; text-align=center"></td>
    	<td colspan="2"><%= osummarystock.FOneItem.Fipkumdiv5 + osummarystock.FOneItem.Foffconfirmno %></td>
    	<td colspan="3"><%= osummarystock.FOneItem.Fipkumdiv4 + osummarystock.FOneItem.Fipkumdiv2 + osummarystock.FOneItem.Foffjupno %></td>

    </tr>
    <tr bgcolor="#FFFFFF">
    	<td colspan="6" align=right>금일 입력된 오차</td>
    	<td align="center" ><input type="text" name="todayerrrealcheckno" value="<%= otodayerritem.FOneItem.Ferrrealcheckno %>"  size="4" maxlength="7" readonly style="background:#CCCCCC; border:1px #999999 solid; text-align=center"></td>
		<td colspan="4"></td>
		<td align="left" colspan="4" >
		<input type="button" class="button" value="실사오차저장" onclick="SaveErr();" <%= ChkIIF((Not IsVaildCode) And (Not IsStockExists),"disabled","") %> >
		</td>
	</tr>
	</form>
</table>

<form name=frmrefresh method=post action="/admin/stock/stockrefresh_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="realstock" value="">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
</form>
<% end if %>
<%
set oitem = Nothing
set oitemoption = Nothing
set ooffitem = Nothing
set otodayerritem = Nothing
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->