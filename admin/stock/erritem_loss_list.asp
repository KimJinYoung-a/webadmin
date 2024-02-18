<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/baditemcls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%

dim makerid,mode, searchtype
makerid = request("makerid")
mode = request("mode")
searchtype = request("searchtype")

searchtype = "err"

dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectmakerid = makerid
osummarystock.FRectSearchType = searchtype

if (makerid<>"") then
    osummarystock.GetDailyErrItemListByBrand
else
    osummarystock.GetDailyErrRealCheckItemListByBrandGroup
end if

dim i

%>
<script language='javascript'>

function PopErrItemLossInput(makerid){
	var popwin = window.open('/common/pop_erritem_re_input.asp?makerid=' + makerid + '&actType=actloss','pop_erritem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function SubmitSearchByBrandNew(makerid) {
	document.frm.makerid.value = makerid;
	document.frm.submit();
}

function ChangePage(v) {
	var frm = document.frm;

	if (v == "bad") {
		frm.action = "baditem_return_list.asp";
	} else {
		frm.action = "erritem_loss_list.asp";
	}

	frm.submit();
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
			<input type="radio" name="searchtype" value="bad" <% if (searchtype = "bad") then %>checked<% end if %> onClick="ChangePage('bad')" > 불량상품
			<input type="radio" name="searchtype" value="err" <% if (searchtype = "err") then %>checked<% end if %> onClick="ChangePage('err')"> 오차등록상품
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<% if makerid<>"" then %>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
        	<input type="button" class="button" value="오차등록상품 로스처리" onclick="PopErrItemLossInput('<%= makerid %>')" border="0">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= osummarystock.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30">구분</td>
		<td width="50">상품코드</td>
		<td width="40">옵션</td>
		<td width="50">이미지</td>
    	<td width="100">브랜드ID</td>

		<td>아이템명</td>
		<td>옵션명</td>
		<td width="40">계약<br>구분</td>

		<td width="50">소비자가</td>
		<td width="40">오차<br>수량</td>
    </tr>

	<% for i=0 to osummarystock.FResultCount - 1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= osummarystock.FItemList(i).FItemgubun %></td>
		<td><%= osummarystock.FItemList(i).FItemid %></td>
		<td><%= osummarystock.FItemList(i).FItemoption %></td>
		<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" ></td>
    	<td><%= osummarystock.FItemList(i).Fmakerid %></td>

		<td align="left"><%= osummarystock.FItemList(i).FItemname %></td>
		<td align="left"><%= osummarystock.FItemList(i).FItemOptionName %></td>
		<td><%= osummarystock.FItemList(i).GetMwDivName %></td>

		<td align="right"><%= formatnumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
		<td><%= osummarystock.FItemList(i).Ferrrealcheckno %></td>
    </tr>
    <% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
		</td>
	</tr>
</table>

<% else %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= osummarystock.FResultCount %></b>
		</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150">브랜드</td>
		<td width="100">오차상품수On</td>
		<td width="100">오차상품수Off</td>
		<td >&nbsp;</td>
	</tr>
	<% for i=0 to osummarystock.FResultCount-1 %>
	<tr bgcolor="#FFFFFF" height="30">
	    <td><a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>');"><%= osummarystock.FItemList(i).FMakerid %></a></td>
	    <td align="center"><%= osummarystock.FItemList(i).FOnCnt %></td>
	    <td align="center"><%= osummarystock.FItemList(i).FOffCnt %></td>
	    <td align="left">
        	<input type="button" class="button" value="오차등록상품 로스처리" onclick="PopErrItemLossInput('<%= osummarystock.FItemList(i).FMakerid %>')" border="0">
	    </td>
	</tr>
	<% next %>
</table>
<% end if %>

<p>




<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
