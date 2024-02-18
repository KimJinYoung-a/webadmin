<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%

dim makerid,mode, searchtype, purchasetype
makerid = request("makerid")
mode = request("mode")
searchtype = request("searchtype")
purchasetype = request("purchasetype")

if (searchtype = "") then
	searchtype = "bad"
end if


'// ===========================================================================
dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectmakerid = makerid
osummarystock.FRectSearchType = searchtype
osummarystock.FRectPurchaseType = purchasetype


if (makerid<>"") then
    osummarystock.GetBadOrErrItemListByBrand
else
    osummarystock.GetBadOrErrItemListByBrandGroup
end if

''response.end

'// ===========================================================================
dim BadOrErrText
if (searchtype="bad") then
    BadOrErrText = "불량"
else
    BadOrErrText = "오차등록"
end if


dim i

%>
<script language='javascript'>
function PopBadItemReInput(makerid){
	var popwin = window.open('/common/pop_baditem_re_input.asp?makerid=' + makerid,'pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function PopBadItemLossInput(makerid){
	var popwin = window.open('/common/pop_baditem_re_input.asp?makerid=' + makerid + '&actType=actloss','pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function SubmitSearchByBrandNew(makerid) {
	var frm = document.frm;

	frm.makerid.value = makerid;
	frm.submit();
}

function ChangePage(v) {
	var frm = document.frm;

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
			<input type="radio" name="searchtype" value="bad" <% if (searchtype = "bad") then %>checked<% end if %> onClick="ChangePage(this)" > 불량상품
			<input type="radio" name="searchtype" value="err" <% if (searchtype = "err") then %>checked<% end if %> onClick="ChangePage(this)"> 오차등록상품
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
			&nbsp;
			구매유형 : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<br><br>
<font size="8">작업중</font>
<br><br>

<% if makerid<>"" then %>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="불량상품반품" onclick="PopBadItemReInput('<%= makerid %>')" border="0">
			&nbsp;
        	<input type="button" class="button" value="불량상품로스처리" onclick="PopBadItemLossInput('<%= makerid %>')" border="0">
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
		<td>브랜드ID</td>
		<td width="50">이미지</td>
		<td width="50">거래<br>구분</td>
		<td width="30">상품<br>구분</td>
		<td width="50">상품코드</td>
		<td width="40">옵션</td>
		<td>상품명<br><font color="blue">[옵션명]</font></td>

		<td width="50">소비자가</td>
		<td width="30">판매<br>여부</td>
		<td width="30">사용<br>여부</td>
		<td width="60"><%= BadOrErrText %><br>수량</td>
		<td width="80">출고수량<br>(ON+OFF)</td>
    </tr>

	<% for i=0 to osummarystock.FResultCount - 1 %>
	<% if (osummarystock.FItemList(i).Fisusing = "Y") then %>
		<tr bgcolor="#FFFFFF" height="30">
	<% else %>
		<tr bgcolor="#BBBBBB" height="30">
	<% end if %>
    	<td><%= osummarystock.FItemList(i).Fmakerid %></td>
    	<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" ></td>
    	<td align="center"><font color="<%= osummarystock.FItemList(i).GetMwDivColor %>"><%= osummarystock.FItemList(i).GetMwDivName %></font></td>
    	<td align="center"><%= osummarystock.FItemList(i).FItemgubun %></td>
		<td align="center"><%= osummarystock.FItemList(i).FItemid %></td>
		<td align="center"><%= osummarystock.FItemList(i).FItemoption %></td>
		<td align="left"><%= osummarystock.FItemList(i).FItemname %><br><font color="blue">[<%= osummarystock.FItemList(i).FItemOptionName %>]</font></td>

		<td align="right"><%= FormatNumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fsellyn %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fisusing %></td>
		<td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fregitemno, 0) %></td>
		<td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fchulgoitemno, 0) %></td>
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
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td rowspan="3">브랜드</td>
		<td rowspan="3">브랜드명</td>
		<td width="40" rowspan="3">브랜드<br>사용<br>여부</td>
		<td colspan="8"><%= BadOrErrText %>상품수량</td>
		<td rowspan="3">비고</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td colspan="3">10</td>
		<td>70</td>
		<td>80</td>
		<td colspan="2">90</td>
		<td rowspan="2" width="80">소계</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="55">매입</td>
		<td width="55">특정</td>
		<td width="55">업배</td>
		<td width="55"></td>
		<td width="55"></td>
		<td width="55">매입</td>
		<td width="55">특정</td>
	</tr>
	<% for i=0 to osummarystock.FResultCount-1 %>
	<% if (osummarystock.FItemList(i).Fuseyn = "Y") then %>
		<tr bgcolor="#FFFFFF" height="30">
	<% else %>
		<tr bgcolor="#BBBBBB" height="30">
	<% end if %>
	    <td><a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>');"><%= osummarystock.FItemList(i).FMakerid %></a></td>
	    <td><%= osummarystock.FItemList(i).Fmakername %></td>
	    <td align="center"><%= osummarystock.FItemList(i).Fuseyn %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem10M, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem10W, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem10U, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem70, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem80, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem90M, 0) %></td>
	    <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Fitem90W, 0) %></td>
	    <td align="center"><%= FormatNumber((osummarystock.FItemList(i).FOnCnt + osummarystock.FItemList(i).FOffCnt), 0) %></td>
	    <td align="left">
			<input type="button" class="button" value="불량상품반품" onclick="PopBadItemReInput('<%= osummarystock.FItemList(i).FMakerid %>')" border="0">
			&nbsp;
        	<input type="button" class="button" value="불량상품로스처리" onclick="PopBadItemLossInput('<%= osummarystock.FItemList(i).FMakerid %>')" border="0">
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
