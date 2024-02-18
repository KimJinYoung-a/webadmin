<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->

<%
dim page
dim makerid, itemid, sellyn, isusing, realstocknotzero


page    = request("page")
makerid = request("makerid")
itemid  = request("itemid")
sellyn  = request("sellyn")
isusing = request("isusing")
realstocknotzero = request("realstocknotzero")

if ((request("research") = "") and (isusing = "")) then
    isusing = "off"
end if

if ((request("research") = "") and (realstocknotzero = "")) then
    realstocknotzero = "on"
end if

if page="" then page=1

dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FCurrPage=page
osummarystock.FPageSize=100
osummarystock.FRectMakerid = makerid
osummarystock.FRectItemID = itemid
osummarystock.FRectOnlyIsUsing = isusing
osummarystock.FRectrealstocknotzero = realstocknotzero

if (makerid<>"") then
    osummarystock.GetCurrentStockByOnlineBrandDanjong
else
    osummarystock.FPageSize=1000
    osummarystock.GetCurrentStockByOnlineBrandDanjong_GroupBrand
end if

dim i, ttlitemno

%>


<script language='javascript'>

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function changecontent(){
	// nothing
}

function Research(page){
    var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.submit();
}

function SearchByBrand(makerid){
    var frm = document.frm;
    frm.makerid.value = makerid;
	frm.submit();
}

function PopReturnItemByBrand(imakerid){
    var params = "menupos=" + frm.menupos.value + "&makerid=" + imakerid
    if (frm.isusing.checked==true){
        params = params + "&isusing=" + frm.isusing.value;
    }

    if (frm.realstocknotzero.checked==true){
        params = params + "&realstocknotzero=" + frm.realstocknotzero.value;
    }

    var popwin = window.open('/admin/stock/return_item.asp?' + params,'PopReturnItemByBrand','width=900, height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
        	&nbsp;
        	상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">
        	&nbsp;
        	<input type=checkbox name="isusing" value="on" <% if isusing="on" then response.write "checked" %> >사용상품만
        	&nbsp;
        	<input type=checkbox name="realstocknotzero" value="on" <% if realstocknotzero="on" then response.write "checked" %> >실사재고가 0이 아닌 상품
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><%= FormatNumber(osummarystock.FTotalCount,0) %></b>
			&nbsp;
			페이지 :
			<% if osummarystock.FCurrPage > 1  then %>
				<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
			<% end if %>
			<b><%= page %> / <%= osummarystock.FTotalpage %></b>
			<% if (osummarystock.FTotalpage - osummarystock.FCurrPage)>0  then %>
				<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
			<% end if %>
		</td>
	</tr>
<% if makerid<>"" then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="40">상품<br>코드</td>
		<td width="50">이미지</td>
		<td width="70">브랜드</td>
		<td>상품명<br>[옵션명]</td>
		<td width="30">계약<br>구분</td>
        <td width="30">전체<br>입고<br>반품</td>
        <td width="30">전체<br>판매<br>반품</td>
        <td width="30">전체<br>출고<br>반품</td>
        <td width="30">기타<br>출고<br>반품</td>
<!--    <td width="35"><b>시스템<br>재고</b></td>	-->
		<td width="30">총<br>불량</td>
<!--    <td width="35"><b>유효<br>재고</b></td>	-->
        <td width="30">총<br>실사<br>오차</td>
        <td width="30"><b>실사<br>재고</b></td>
        <td width="30">총<br>상품<br>준비</td>
        <td width="30">총<br>주문<br>접수</td>
        <td width="30"><b>예상<br>재고</b></td>

		<td width="30">판매<br>여부</td>
		<td width="30">사용<br>여부</td>
		<td width="30">한정<br>여부</td>
		<td width="30">단종<br>여부</td>
<!--	<td width="30">품절<br>여부</td>	-->
    </tr>
<% for i=0 to osummarystock.FresultCount-1 %>
	<% if osummarystock.FItemList(i).Fisusing="Y" then %>
    <tr bgcolor="#FFFFFF" align="center">
    <% else %>
    <tr bgcolor="#EEEEEE" align="center">
    <% end if %>
    	<td><a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemID %>');"><%= osummarystock.FItemList(i).FItemID %></a></td>
		<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
		<td align="left"><a href="javascript:SearchByBrand('<%= osummarystock.FItemList(i).FMakerID %>');"><%= osummarystock.FItemList(i).FMakerID %></a></td>
		<td align="left">
			<a href="javascript:PopItemDetail('<%= osummarystock.FItemList(i).FItemID %>','<%= osummarystock.FItemList(i).FItemOption %>')"><%= osummarystock.FItemList(i).FItemName %></a>
			<% if (osummarystock.FItemList(i).FItemOptionName <> "") then %>
			<br><font color="blue">[<%= osummarystock.FItemList(i).FItemOptionName %>]</font>
			<% end if %>
        </td>
        <td><%= osummarystock.FItemList(i).GetMwDivName %></td>
		<td><%= osummarystock.FItemList(i).Ftotipgono %></td>
		<td><%= -1*osummarystock.FItemList(i).Ftotsellno %></td>
		<td><%= osummarystock.FItemList(i).Foffchulgono + osummarystock.FItemList(i).Foffrechulgono %></td>
        <td><%= osummarystock.FItemList(i).Fetcchulgono + osummarystock.FItemList(i).Fetcrechulgono %></td>
<!--    <td><%= osummarystock.FItemList(i).Ftotsysstock %></td>	-->
        <td><%= osummarystock.FItemList(i).Ferrbaditemno %></td>
<!--    <td><%= osummarystock.FItemList(i).Favailsysstock %></td>	-->
        <td><%= osummarystock.FItemList(i).Ferrrealcheckno %></td>
        <td><b><%= osummarystock.FItemList(i).Frealstock %></b></td>
        <td><%= osummarystock.FItemList(i).Fipkumdiv5 + osummarystock.FItemList(i).Foffconfirmno %></td>
        <td><%= osummarystock.FItemList(i).Fipkumdiv4 + osummarystock.FItemList(i).Fipkumdiv2 + osummarystock.FItemList(i).Foffjupno %></td>
        <td><b><%= osummarystock.FItemList(i).GetMaystock %></b></td>

        <td><font color="<%= ynColor(osummarystock.FItemList(i).Fsellyn) %>"><%= osummarystock.FItemList(i).Fsellyn %></font></td>
        <td><font color="<%= ynColor(osummarystock.FItemList(i).Fisusing) %>"><%= osummarystock.FItemList(i).Fisusing %></font></td>
        <td>
        	<font color="<%= ynColor(osummarystock.FItemList(i).Flimityn) %>"><%= osummarystock.FItemList(i).Flimityn %>
			<% if (osummarystock.FItemList(i).Flimityn = "Y") then %>
				<br>
				(<%= osummarystock.FItemList(i).GetLimitStr %>)
			<% end if %>
			</font>
        </td>
        <td>
            <% if osummarystock.FItemList(i).FDanjongyn="Y" then %>
            <font color="#33CC33">단종</font>
            <% elseif osummarystock.FItemList(i).FDanjongyn="M" then %>
            <font color="#33CC33">MD<br>품절</font>
            <% elseif osummarystock.FItemList(i).FDanjongyn="S" then %>
            <font color="#33CC33">일시<br>품절</font>
            <% else %>
            <% end if %>
        </td>
<!--    <td><% if osummarystock.FItemList(i).IsSoldOut  then %><font color="red">품절</font><% end if %></td>	-->
	</tr>
	</form>
<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if osummarystock.HasPreScroll then %>
					<a href="javascript:GotoPage(<%= osummarystock.StartScrollPage-1 %>)">[pre]</a>
			<% else %>
					[pre]
			<% end if %>

			<% for i=0 + osummarystock.StartScrollPage to osummarystock.FScrollCount + osummarystock.StartScrollPage - 1 %>
			        <% if i>osummarystock.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
					<font color="red">[<%= i %>]</font>
				<% else %>
					<a href="javascript:GotoPage(<%= i %>)">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if osummarystock.HasNextScroll then %>
					<a href="javascript:GotoPage(<%= i %>)">[next]</a>
			<% else %>
					[next]
			<% end if %>
		</td>
	</tr>
</table>


<% else %>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150">브랜드</td>
		<td width="100">대상상품수</td>
		<td >&nbsp;</td>
	</tr>
	<% for i= 0 to osummarystock.FResultCount-1 %>
	<%
	ttlitemno = ttlitemno + osummarystock.FItemList(i).FCnt
    %>
	<tr bgcolor="#FFFFFF" >
	    <td><a href="javascript:PopReturnItemByBrand('<%= osummarystock.FItemList(i).FMakerid %>');"><%= osummarystock.FItemList(i).FMakerid %></a></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).FCnt,0) %></td>
	    <td align="right"><!-- <img src="/images/icon_detail.gif"> --></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" >
	    <td></td>
	    <td align="center"><%= FormatNumber(ttlitemno,0) %></td>
	    <td></td>
	</tr>
</table>
<% end if %>


<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
