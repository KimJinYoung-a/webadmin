<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  브랜드별재고
' History : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%

dim research
dim makerid, onoffgubun, mwdiv, usingyn, centermwdiv, itemgubun, startMon, endMon, purchaseType
dim stocktype, limitrealstock
dim page, pagesize

dim i

research     	= RequestCheckVar(request("research"),32)
makerid     	= RequestCheckVar(request("makerid"),32)
onoffgubun     	= RequestCheckVar(request("onoffgubun"),32)
mwdiv     		= RequestCheckVar(request("mwdiv"),32)
usingyn     	= RequestCheckVar(request("usingyn"),32)
centermwdiv     = RequestCheckVar(request("centermwdiv"),32)
stocktype     	= RequestCheckVar(request("stocktype"),32)
limitrealstock 	= RequestCheckVar(request("limitrealstock"),32)
page     		= RequestCheckVar(request("page"),32)
pagesize     	= RequestCheckVar(request("pagesize"),32)
itemgubun     	= RequestCheckVar(request("itemgubun"),32)
startMon     	= RequestCheckVar(request("startMon"),32)
endMon     		= RequestCheckVar(request("endMon"),32)
purchaseType	= RequestCheckVar(request("purchaseType"),32)

if (research = "") then
	stocktype = "real"
end if

if (page = "") then
	page = "1"
end if
if (pagesize = "") then
	pagesize = "100"
end if

if itemgubun = "" then
	''itemgubun = "10"
end if


dim osummarystockbrand
set osummarystockbrand = new CSummaryItemStock

	osummarystockbrand.FPageSize = pagesize
	osummarystockbrand.FCurrPage = page

	osummarystockbrand.FRectMakerid = makerid
	osummarystockbrand.FRectOnlyIsUsing = usingyn
	osummarystockbrand.FRectMWDiv = mwdiv
	osummarystockbrand.FRectCenterMWDiv = centermwdiv
	osummarystockbrand.FRectStockType = stocktype
	osummarystockbrand.FRectlimitrealstock = limitrealstock
	osummarystockbrand.FRectItemGubun = itemgubun
	osummarystockbrand.FRectPurchaseType = purchaseType

	if IsNumeric(startMon) then
		osummarystockbrand.FRectStartDate = startMon
	elseif (startMon <> "") then
		response.write "<script>alert('월령은 숫자만 가능합니다. " & startMon & "')</script>"
	end if
	if IsNumeric(endMon) then
		osummarystockbrand.FRectEndDate = endMon
	elseif (endMon <> "") then
		response.write "<script>alert('월령은 숫자만 가능합니다. " & endMon & "')</script>"
	end if

	if itemgubun = "10" or itemgubun = "" then
		osummarystockbrand.GetRealStockByOnlineBrand
	else
		osummarystockbrand.GetRealStockByOfflineBrand
	end if

%>

<script language="JavaScript" src="/js/ttpbarcode.js"></script>
<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.submit();
}

function SubmitSearch() {
    document.frm.submit();
}

function jsPopDetail(makerid, mwdiv, centermwdiv, itemgubun) {
	var onoffgubun = itemgubun;
	if (onoffgubun == "10") {
		onoffgubun = "on";
	} else if (itemgubun == "exc10") {
		onoffgubun = "off";
	} else {
		onoffgubun = "off" + itemgubun;
	}
	var url = "/admin/stock/brandcurrentstock.asp?menupos=708&onoffgubun=" + onoffgubun + "&makerid=" + makerid + "&usingyn=&mwdiv=" + mwdiv + "&centermwdiv=" + centermwdiv + "&stocktype=<%= stocktype %>&limitrealstock=<%= limitrealstock %>&startMon=<%= startMon %>&endMon=<%= endMon %>";

	var popwin = window.open(url, "jsPopDetail", "width=1500,height=800,scrollbars=yes,resizable=yes");
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
		* 브랜드:	<% drawSelectBoxDesignerwithName "makerid", makerid %>
		&nbsp;&nbsp;
		* 사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;&nbsp;
		* 거래구분 :<% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;&nbsp;
		* 센터매입구분 :
		<select class="select" name="centermwdiv">
			<option value="">전체</option>
			<option value="M" <% if centermwdiv="M" then response.write "selected" %> >매입</option>
			<option value="W" <% if centermwdiv="W" then response.write "selected" %> >특정</option>
			<option value="N" <% if centermwdiv="N" then response.write "selected" %> >미지정</option>
		</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="SubmitSearch();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 상품구분: <% drawSelectBoxItemGubunForSearch "itemgubun", itemgubun %>
		&nbsp;&nbsp;
	    * 재고구분 :
		<select name="stocktype" class="select">
			<option value="sys" <% if (stocktype = "sys") then %>selected<% end if %> >시스템재고</option>
			<option value="real" <% if (stocktype = "real") then %>selected<% end if %> >유효재고</option>
		</select>
		: <% drawSelectBoxexistsstock "limitrealstock", limitrealstock, "" %>
		&nbsp;&nbsp;
		* 표시갯수 :
		<select class="select" name="pagesize">
			<option value="100" <% if (pagesize = "100") then %>selected<% end if %> >100 개</option>
			<option value="500" <% if (pagesize = "500") then %>selected<% end if %> >500 개</option>
			<option value="1000" <% if (pagesize = "1000") then %>selected<% end if %> >1000 개</option>
		</select>
		&nbsp;&nbsp;
		* 재고월령 :
		<input type="text" class="text" name="startMon" size="2" value="<%= startMon %>">
		~
		<input type="text" class="text" name="endMon" size="2" value="<%= endMon %>"> 개월
		&nbsp;&nbsp;
	    * 구매유형 :
		<% drawPartnerCommCodeBox true,"purchasetype","purchaseType",purchaseType,"" %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="29">
		검색결과 : <b><%= osummarystockbrand.FTotalCount %></b>
		&nbsp;
		페이지 :
		<% if osummarystockbrand.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>
		<b><%= page %> / <%= osummarystockbrand.FTotalPage %></b>
		<% if (osummarystockbrand.FTotalpage - osummarystockbrand.FCurrPage)>0  then %>
			<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"></td>
    <td width="200">브랜드ID</td>
	<td width="60">상품구분</td>
	<td width="60">매입구분</td>
	<td width="60">센터<br>매입구분</td>
	<td width="70">상품품목수</td>
	<td width="70">재고>0<br>상품수</td>
	<td width="70">총판매</td>

	<td width="70">시스템재고</td>
	<td width="70">실사오차</td>
	<td width="70">실사재고</td>
	<td width="70">불량</td>
	<td width="70">유효재고</td>

	<td >비고</td>
</tr>
<% if (osummarystockbrand.FResultCount = 0) then %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td colspan="14">내역이 없습니다.</td>
</tr>
<% else %>
<% for i=0 to osummarystockbrand.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="25">
    <td></td>
    <td><%= osummarystockbrand.FItemList(i).Fmakerid %></td>
    <td><%= osummarystockbrand.FItemList(i).Fitemgubun %></td>
	<td><%= osummarystockbrand.FItemList(i).Fmwdiv %></td>
	<td><%= osummarystockbrand.FItemList(i).FCentermwdiv %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).FitemCnt,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).FitemPlusCnt,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).Ftotsellno,0) %></td>

	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).Ftotsysstock,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).Ferrrealcheckno,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).getErrAssignStock,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).Ferrbaditemno,0) %></td>
	<td align="right" style="padding-right:5px"><%= FormatNumber(osummarystockbrand.FItemList(i).Frealstock,0) %></td>

	<td>
		<input type="button" class="button" value="상세" onClick="jsPopDetail('<%= osummarystockbrand.FItemList(i).Fmakerid %>', '<%= osummarystockbrand.FItemList(i).Fmwdiv %>', '<%= osummarystockbrand.FItemList(i).Fcentermwdiv %>', '<%= osummarystockbrand.FItemList(i).Fitemgubun %>')" >
	</td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="29" align="center">
		<% if osummarystockbrand.HasPreScroll then %>
		<a href="javascript:NextPage('<%= osummarystockbrand.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + osummarystockbrand.StartScrollPage to osummarystockbrand.FScrollCount + osummarystockbrand.StartScrollPage - 1 %>
			<% if i>osummarystockbrand.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if osummarystockbrand.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% end if %>
</table>

<%
set osummarystockbrand = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
