<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  오프라인 주문서
' History : 2016.09.05 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/classes/stock/ipchulbarcodecls.asp"-->
<%
dim chulgoyn, showdeleted, showmichulgo, michulgoreason, statecd, itemid, makerid, shopdiv, cartoonboxmasteridx
dim day5chulgo, shortchulgo, tempshort, danjong, etcshort, innerboxno, research, dateType
dim yyyy1,mm1 , dd1, yyyy2, mm2, dd2, fromDate, toDate, masteridx, baljucode, shopid, i
	masteridx = getNumeric(requestCheckVar(request("masteridx"),10))
	shopid = requestCheckVar(request("shopid"),32)
	baljucode = requestCheckVar(request("baljucode"),32)
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	chulgoyn = request("chulgoyn")
	showdeleted = request("showdel")		'웹서버 웹나이트가 파라미터중 delete 문구가 있는 경우 막는다.
	showmichulgo = request("showmichulgo")
	michulgoreason = request("michulgoreason")
	innerboxno = request("innerboxno")
	statecd = request("statecd")
	itemid = request("itemid")
	makerid = request("makerid")
	shopdiv = request("shopdiv")
	day5chulgo = request("day5chulgo")
	shortchulgo = request("shortchulgo")
	tempshort = request("tempshort")
	danjong = request("danjong")
	etcshort = request("etcshort")
	research = request("research")
	dateType = requestCheckVar(request("dateType"),1)
	cartoonboxmasteridx = getNumeric(requestCheckVar(request("cartoonboxmasteridx"),10))

if dateType="" then dateType="J"
if (research = "") then
	showdeleted = "N"
	michulgoreason = "all"
end if

michulgoreason = "|"
if (day5chulgo = "Y") then
	'5일내출고
	michulgoreason = michulgoreason + "5|"
end if
if (shortchulgo = "Y") then
	'재고부족
	michulgoreason = michulgoreason + "S|"
end if
if (tempshort = "Y") then
	'일시품절
	michulgoreason = michulgoreason + "T|"
end if
if (danjong = "Y") then
	'단종
	michulgoreason = michulgoreason + "D|"
end if
if (etcshort = "Y") then
	'기타
	michulgoreason = michulgoreason + "E|"
end if

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))
	dd1 = Cstr(day(now()))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oforeign_detail
set oforeign_detail = new CStorageDetail
	oforeign_detail.FPageSize = 500
	oforeign_detail.FCurrPage = 1
	oforeign_detail.FRectbaljucode = baljucode
	oforeign_detail.FRectMasterIdx = masteridx
	oforeign_detail.FRectshopid = shopid
	oforeign_detail.FRectmakerid = makerid
	oforeign_detail.FRectItemid = itemid
	oforeign_detail.FRectstartdate = fromDate
	oforeign_detail.FRectenddate = toDate
	oforeign_detail.FRectinnerboxno = innerboxno
	oforeign_detail.FRectShopdiv = shopdiv
	oforeign_detail.FRectShowDeleted = "N"
	oforeign_detail.FRectMichulgoReason = michulgoreason
	oforeign_detail.FRectDateType = dateType
	oforeign_detail.FRectcartoonboxmasteridx = cartoonboxmasteridx

	if (statecd = "A") then
		oforeign_detail.FRectChulgoYN = "N"
	else
		oforeign_detail.FRectStatecd = statecd
	end if

	oforeign_detail.Getordersheet_foreign_detail
%>

<script type="text/javascript">

function downloadOrder(masteridx, baljucode, shopid) {
	frm.masteridx.value=masteridx;
	frm.baljucode.value=baljucode;
	frm.shopid.value=shopid;
	frm.action='/common/popOrderSheet_foreign_excel.asp';
	frm.target='view';
	frm.submit();
	frm.action='';
	frm.target='';
	return false;
}

//기본상품명일괄적용
function autoiteminfo(tp) {
	if (tp==''){
		alert('구분자가 없습니다.');
		return;
	}
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	var frm;
	var itemname = '';
	var itemoptionname = '';
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			itemname = '';
			itemoptionname = '';
			if (frm.cksel.checked){
				if (tp=='1'){
					itemname = frm.itemname_10x10.value;
					frm.itemname.value = itemname
					itemoptionname = frm.itemoptionname_10x10.value;
					frm.itemoptionname.value = itemoptionname;
				}else if (tp=='2'){
					itemname = frm.itemname_en.value;
					frm.itemname.value = itemname
					itemoptionname = frm.itemoptionname_en.value;
					frm.itemoptionname.value = itemoptionname;
				}
			}
		}
	}
	return;
}

function ModiArr(upfrm){
    if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm1;
	var itemname = '';
	var itemoptionname = '';

	upfrm.detailidxarr.value = '';
	upfrm.baljucodearr.value = '';
	upfrm.itemgubunarr.value = '';
	upfrm.itemidarr.value = '';
	upfrm.itemoptionarr.value = '';
	upfrm.itemnamearr.value = '';
	upfrm.itemoptionnamearr.value = '';
	for (var i=0;i<document.forms.length;i++){
		frm1 = document.forms[i];
		if (frm1.name.substr(0,9)=='frmBuyPrc') {
			if (frm1.cksel.checked){
				if (frm1.itemname.value == ''){
					alert('상품명을 입력해주세요');
					frm1.itemname.focus();
					return;
				}

				upfrm.detailidxarr.value = upfrm.detailidxarr.value + frm1.detailidx.value + ',' ;
				upfrm.baljucodearr.value = upfrm.baljucodearr.value + frm1.baljucode.value + ',' ;
				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm1.itemgubun.value + ',' ;
				upfrm.itemidarr.value = upfrm.itemidarr.value + frm1.itemid.value + ',' ;
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm1.itemoption.value + ',' ;

				itemname = frm1.itemname.value;
				upfrm.itemnamearr.value = upfrm.itemnamearr.value + itemname.replace(',','!@#') + ',' ;

				itemoptionname = frm1.itemoptionname.value
				upfrm.itemoptionnamearr.value = upfrm.itemoptionnamearr.value + itemoptionname.replace(',','!@#') + ',' ;
			}
		}
	}

	upfrm.mode.value = 'itemedit';
	upfrm.target='view';
	upfrm.method='post';
	upfrm.action = '/common/offshop/order/orderitem_info_foreign_process.asp';
	upfrm.submit();
	upfrm.detailidxarr.value = '';
	upfrm.baljucodearr.value = '';
	upfrm.itemgubunarr.value = '';
	upfrm.itemidarr.value = '';
	upfrm.itemoptionarr.value = '';
	upfrm.itemnamearr.value = '';
	upfrm.itemoptionnamearr.value = '';
}

</script>

<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" ></iframe>	

<form name="actfrm" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="detailidxarr" value="">
<input type="hidden" name="baljucodearr" value="">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="itemnamearr" value="">
<input type="hidden" name="itemoptionnamearr" value="">
</form>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="masteridx" value="">
<input type="hidden" name="research" value="on">

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		ShopID : <% drawSelectBoxOffShop "shopid",shopid %>
		&nbsp;
		<select class="select" name="dateType">
			<option value="B" <%= CHKIIF(dateType="B", "selected", "") %> >발주일</option>
			<option value="C" <%= CHKIIF(dateType="C", "selected", "") %> >출고일</option>
			<option value="J" <%= CHKIIF(dateType="J", "selected", "") %> >주문일</option>
		</select> :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		주문상태 :
		<select name="statecd" class="select">
			<option value="">전체
			<option value=" " <% if statecd=" " then response.write "selected" %> >작성중
			<option value="0" <% if statecd="0" then response.write "selected" %> >주문접수
			<option value="1" <% if statecd="1" then response.write "selected" %> >주문확인
			<option value="2" <% if statecd="2" then response.write "selected" %> >입금대기
			<option value="5" <% if statecd="5" then response.write "selected" %> >배송준비
			<option value="6" <% if statecd="6" then response.write "selected" %> >출고대기
			<option value="7" <% if statecd="7" then response.write "selected" %> >출고완료
			<option value="8" <% if statecd="8" then response.write "selected" %> >입고대기
			<option value="9" <% if statecd="9" then response.write "selected" %> >입고완료
			<option value="">========
			<option value="A" <% if statecd="A" then response.write "selected" %> >출고이전전체
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		주문코드 : <input type="text" class="text" name="baljucode" value="<%= baljucode %>" size="10" maxlength="8">
		&nbsp;
		브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		&nbsp;
		상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="12">
		&nbsp;
		InnerBoxNO : <input type="text" class="text" name="innerboxno" value="<%= innerboxno %>" size="4" maxlength="12">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
     	현재 SHOP구분 :
     	<input type="radio" name="shopdiv" value="" <% if shopdiv="" then response.write "checked" %> >전체
		<input type="radio" name="shopdiv" value="direct" <% if shopdiv="direct" then response.write "checked" %> >직영
		<input type="radio" name="shopdiv" value="franchisee" <% if shopdiv="franchisee" then response.write "checked" %> >가맹점
		<input type="radio" name="shopdiv" value="foreign" <% if shopdiv="foreign" then response.write "checked" %> >해외
		<input type="radio" name="shopdiv" value="buy" <% if shopdiv="buy" then response.write "checked" %> >도매
		&nbsp;&nbsp;
		|
		&nbsp;&nbsp;
		미출고사유 :
		<input type="checkbox" name="day5chulgo" value="Y" <% if day5chulgo="Y" then response.write "checked" %> >5일내출고
		<input type="checkbox" name="shortchulgo" value="Y" <% if shortchulgo="Y" then response.write "checked" %> >재고부족
		<input type="checkbox" name="tempshort" value="Y" <% if tempshort="Y" then response.write "checked" %> >일시품절
		<input type="checkbox" name="danjong" value="Y" <% if danjong="Y" then response.write "checked" %> >단종
		<input type="checkbox" name="etcshort" value="Y" <% if etcshort="Y" then response.write "checked" %> >기타
	</td>
</tr>
</table>
<!-- 검색 끝 -->

</form>

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		<input type="button" onclick="autoiteminfo('1'); return false;" value="선택온라인상품정보적용" class="button">
		<input type="button" onclick="autoiteminfo('2'); return false;" value="선택EN상품정보적용" class="button">
		<input type="button" class="button" value="선택일괄저장" onclick="ModiArr(actfrm)">
    </td>
    <td align="right">
    	<input type="button" onclick="downloadOrder('<%= masteridx %>','<%= baljucode %>','<%= shopid %>'); return false;" value="엑셀다운로드(5000개제한)" class="button">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		검색결과 : <b><%= oforeign_detail.FTotalCount %></b> ※ 최대 500개 까지 검색 됩니다.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>Order NO</td>
	<td>Release NO</td>
	<td>Order Date</td>
	<td>Packing Date</td>
	<td>Inner Box NO</td>
	<td>Carton Box NO</td>
	<td>Brand</td>
	<td>Item Code</td>
	<td>Item Name</td>
	<td>Option Name</td>
	<td>Qty</td>
	<td>Retail Price</td>
	<td>Wholesale Price</td>
	<td>Discount Rate</td>
	<td>Total Price</td>
	<td>RRP</td>
	<td>exchangeRate</td>
	<td>multipleRate</td>
	<td></td>
	<td>Item Name<Br>[KR]</td>
	<td>Option Name<Br>[KR]</td>
	<td>Material<Br>[KR]</td>
	<td>Origin<Br>[KR]</td>
	<td></td>
	<td>Item Name<Br>[EN]</td>
	<td>Option Name<Br>[EN]</td>
	<td>Material<Br>[EN]</td>
	<td>Origin<Br>[EN]</td>
	<td>Item<Br>Weight(g)</td>
</tr>
<% if oforeign_detail.FresultCount > 0 then %>
<% for i=0 to oforeign_detail.FresultCount-1 %>
<form method="get" action="" name="frmBuyPrc<%=i%>" style="margin:0px;">
<input type="hidden" name="detailidx" value="<%= oforeign_detail.FItemList(i).fdetailidx %>">
<input type="hidden" name="baljucode" value="<%= oforeign_detail.FItemList(i).fbaljucode %>">
<input type="hidden" name="itemgubun" value="<%= oforeign_detail.FItemList(i).fitemgubun %>">
<input type="hidden" name="itemid" value="<%= oforeign_detail.FItemList(i).Fitemid %>">
<input type="hidden" name="itemoption" value="<%= oforeign_detail.FItemList(i).fitemoption %>">
<tr bgcolor="#FFFFFF">
	<td >
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	</td>
	<td><%= oforeign_detail.FItemList(i).fbaljucode %></td>
	<td><%= oforeign_detail.FItemList(i).falinkcode %></td>
	<td><%= oforeign_detail.FItemList(i).fregdate %></td>
	<td><%= oforeign_detail.FItemList(i).fbaljudate %></td>
	<td><%= oforeign_detail.FItemList(i).finnerboxno %></td>
	<td><%= oforeign_detail.FItemList(i).fcartoonboxno %></td>
	<td><%= oforeign_detail.FItemList(i).fmakerid %></td>
	<td>
		<%= oforeign_detail.FItemList(i).fitemgubun %>
		<%= CHKIIF(oforeign_detail.FItemList(i).Fitemid>=1000000,Format00(8,oforeign_detail.FItemList(i).Fitemid),Format00(6,oforeign_detail.FItemList(i).Fitemid)) %>
		<%= oforeign_detail.FItemList(i).fitemoption %>
	</td>
	<td align="left">
		<input type="text" name="itemname" value='<%= replace(oforeign_detail.FItemList(i).fitemname,"'","""") %>'>
	</td>
	<td align="left">
		<input type="text" name="itemoptionname" value='<%= replace(oforeign_detail.FItemList(i).fitemoptionname,"'","""") %>' <% 'if oforeign_detail.FItemList(i).fitemoption="0000" then response.write " readonly" %>>
	</td>
	<td><%= FormatNumber(oforeign_detail.FItemList(i).frealitemno, 0) %></td>
	<td><%= getdisp_price(oforeign_detail.FItemList(i).fsellcash,oforeign_detail.FItemList(i).fcurrencyunit) %></td>
	<td><%= getdisp_price(oforeign_detail.FItemList(i).fsuplycash,oforeign_detail.FItemList(i).fcurrencyunit) %></td>
	<td><%= oforeign_detail.FItemList(i).fdefaultsuplymargin %></td>
	<td><%= getdisp_price(oforeign_detail.FItemList(i).ftotalsuplycash,oforeign_detail.FItemList(i).fcurrencyunit) %></td>
	<td><%= FormatNumber(oforeign_detail.FItemList(i).flcprice, 0) %></td>
	<td><%= oforeign_detail.FItemList(i).fexchangeRate %></td>
	<td><%= oforeign_detail.FItemList(i).fmultipleRate %></td>
	<td></td>
	<td align="left">
		<%= oforeign_detail.FItemList(i).fitemname_10x10 %>
		<input type="hidden" name="itemname_10x10" value='<%= replace(oforeign_detail.FItemList(i).fitemname_10x10,"'","""") %>'>
	</td>
	<td align="left">
		<%= oforeign_detail.FItemList(i).foptionname_10x10 %>
		<input type="hidden" name="itemoptionname_10x10" value='<%= replace(oforeign_detail.FItemList(i).foptionname_10x10,"'","""") %>'>
	</td>
	<td align="left"><%= oforeign_detail.FItemList(i).fitemsource_10x10 %></td>
	<td align="left"><%= oforeign_detail.FItemList(i).fsourcearea_10x10 %></td>
	<td></td>
	<td align="left">
		<%= oforeign_detail.FItemList(i).fitemname_en %>
		<input type="hidden" name="itemname_en" value='<%= replace(oforeign_detail.FItemList(i).fitemname_en,"'","""") %>'>
	</td>
	<td align="left">
		<%= oforeign_detail.FItemList(i).foptionname_en %>
		<input type="hidden" name="itemoptionname_en" value='<%= replace(oforeign_detail.FItemList(i).foptionname_en,"'","""") %>'>
	</td>
	<td align="left"><%= oforeign_detail.FItemList(i).fitemsource_en %><input type="hidden" name="itemsource_en" value="<%= oforeign_detail.FItemList(i).fitemsource_en %>">
	</td>
	<td align="left"><%= oforeign_detail.FItemList(i).fsourcearea_en %></td>
	<td><%= oforeign_detail.FItemList(i).fitemweight %></td>
</tr>
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="29" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<%
set oforeign_detail = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
