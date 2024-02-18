<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 샵별패킹내역(박스별) 공용페이지
' History : 2011.01.18 이상구 생성
'			2012.08.14 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_balju.asp"-->
<%

if (Not C_IS_SHOP) and (Not C_ADMIN_USER) then
    %><html>
    <script type='text/javascript'>
    alert("세션이 종료되었습니다. \n재로그인후 사용하실수 있습니다.");
    top.location = "/index.asp";
    </script>
    </html><%
    response.End : dbget.close()
end if

dim page, shopid, chulgoyn, showdeleted, showmichulgo, michulgoreason ,statecd, itemid, brandid
dim day5chulgo, shortchulgo, tempshort, danjong, etcshort ,research, i, shopdiv, baljucode ,tmpcartonboxbarcode
dim innerboxno, innerboxsongjangno, cartoonboxno, cartonboxsongjangno ,innerboxbarcode , cartonboxbarcode
dim yyyy1,mm1 , dd1, yyyy2, mm2, dd2, fromDate, toDate ,siteSeq
dim dateType
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	chulgoyn = requestCheckVar(request("chulgoyn"),1)
	showdeleted = requestCheckVar(request("showdel"),1)		'웹서버 웹나이트가 파라미터중 delete 문구가 있는 경우 막는다.
	showmichulgo = requestCheckVar(request("showmichulgo"),10)
	michulgoreason = requestCheckVar(request("michulgoreason"),32)
	statecd = requestCheckVar(request("statecd"),10)
	itemid = requestCheckVar(request("itemid"),10)
	brandid = requestCheckVar(request("brandid"),32)
	shopdiv = requestCheckVar(request("shopdiv"),32)
	baljucode = requestCheckVar(request("baljucode"),32)
	day5chulgo = requestCheckVar(request("day5chulgo"),1)
	shortchulgo = requestCheckVar(request("shortchulgo"),1)
	tempshort = requestCheckVar(request("tempshort"),1)
	danjong = requestCheckVar(request("danjong"),1)
	etcshort = requestCheckVar(request("etcshort"),1)
	research = requestCheckVar(request("research"),2)
	innerboxno 			= requestCheckVar(request("innerboxno"),10)
	innerboxsongjangno 	= requestCheckVar(request("innerboxsongjangno"),32)
	innerboxbarcode = requestCheckVar(request("innerboxbarcode"),32)
	cartoonboxno 		= requestCheckVar(request("cartoonboxno"),10)
	cartonboxsongjangno = requestCheckVar(request("cartonboxsongjangno"),32)
	cartonboxbarcode = requestCheckVar(request("cartonboxbarcode"),32)
	dateType = requestCheckVar(request("dateType"),1)

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

siteSeq = "10"
if (page = "") then
	page = 1
end if

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

if dateType="" then dateType="B"

yyyy1 = requestCheckVar(request("yyyy1"),4)
mm1 = requestCheckVar(request("mm1"),2)
dd1 = requestCheckVar(request("dd1"),2)
yyyy2 = requestCheckVar(request("yyyy2"),4)
mm2 = requestCheckVar(request("mm2"),2)
dd2 = requestCheckVar(request("dd2"),2)

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))
	if (C_IS_SHOP) then
		'샵은 한달
	    mm1 = Cstr(Month(now()) - 1)
	else
		mm1 = Cstr(Month(now()))
	end if
	dd1 = Cstr(day(now()))

	fromDate = DateSerial(yyyy1, mm1, dd1)
	yyyy1 = CStr(Year(fromDate))
	mm1 = CStr(Month(fromDate))
	dd1 = CStr(Day(fromDate))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)
'prevMonth = DateAdd("m", -1, dtDate)

dim oshopbalju
set oshopbalju = new CShopBalju
	oshopbalju.FRectFromDate = fromDate
	oshopbalju.FRectToDate = toDate
	oshopbalju.FRectDateType = dateType
	oshopbalju.FRectBaljuId = shopid
	oshopbalju.FRectItemid = itemid
	oshopbalju.FRectBrandid = brandid
	oshopbalju.FRectShopdiv = shopdiv
	oshopbalju.FRectBaljucode = baljucode
	oshopbalju.FRectBoxno = innerboxno
	oshopbalju.FRectCartonBoxno = cartoonboxno
	oshopbalju.FRectBoxsongjangno = innerboxsongjangno
	oshopbalju.FRectCartonBoxsongjangno = cartonboxsongjangno
	oshopbalju.frectinnerboxbarcode = innerboxbarcode
	oshopbalju.frectcartonboxbarcode = cartonboxbarcode

	if (statecd = "A") then
		oshopbalju.FRectChulgoYN = "N"
	else
		oshopbalju.FRectStatecd = statecd
		if (C_IS_SHOP) then
			oshopbalju.FRectStatecd = ""
		    oshopbalju.FRectChulgoYN = "Y"
		end if
	end if

	oshopbalju.FRectShowDeleted = "N"
	'oshopbalju.FRectMichulgoReason = michulgoreason
	oshopbalju.FCurrPage = page
	oshopbalju.Fpagesize = 25
	oshopbalju.GetShopBaljuByBoxNEW
%>

<script type='text/javascript'>

function regsubmit(page){
	if (frm.innerboxbarcode.value != ''){
		if (frm.innerboxbarcode.value.length != 19){
			alert('<%= CTX_Type_Mismatch %> (<%= CTX_INNERBOX %> <%= CTX_Barcode %>)');
			return;
		}

		if (!IsDouble(frm.innerboxbarcode.value)){
			alert('<%= CTX_Only_numbers %> (<%= CTX_INNERBOX %> <%= CTX_Barcode %>)');
			frm.innerboxbarcode.focus();
			return;
		}
	}

	if (frm.cartonboxbarcode.value != ''){
		if (frm.cartonboxbarcode.value.length != 19){
			alert('<%= CTX_Type_Mismatch %> (<%= CTX_CARTONBOX %> <%= CTX_Barcode %>)');
			return;
		}

		if (!IsDouble(frm.cartonboxbarcode.value)){
			alert('<%= CTX_Only_numbers %> (<%= CTX_CARTONBOX %> <%= CTX_Barcode %>)');
			frm.cartonboxbarcode.focus();
			return;
		}
	}

	frm.page.value=page;
	frm.submit();
}

function GotoPage(pageno) {
	frm.page.value = pageno;
	frm.submit();
}

function ModifyBox(frm) {
	if (CheckBox(frm) == true) {
		/*
		if (frm.detailidx.value =="") {
			alert("로직스에서 입력된 내역에 대해서만 수정이 가능합니다.");
			return;
		}
		*/
		if (confirm("입력하시겠습니까?") == true) {
			frm.submit();
		}
	}
}

function SetRecv(frm) {
	if (confirm("도착확인하시겠습니까?") == true) {
		frm.mode.value = "setrecv";
		frm.submit();
	}
}

function CheckBox(frm) {
	if (frm.cartoonboxno.value == "") {
		alert("Carton박스번호를 입력하세요.");
		frm.cartoonboxno.focus();
		return false;
	}

	if (frm.cartoonboxno.value*0 != 0) {
		alert("Carton박스번호는 숫자만 가능합니다.");
		frm.cartoonboxno.focus();
		return false;
	}

	if (frm.innerboxweight.value == "") {
		frm.innerboxweight.value = 0;
	}

	if (frm.innerboxweight.value*0 != 0) {
		alert("Inner박스 무게는 숫자만 가능합니다.");
		frm.innerboxweight.focus();
		return false;
	}

	if (frm.cartoonboxweight.value == "") {
		frm.cartoonboxweight.value = 0;
	}

	if (frm.cartoonboxweight.value*0 != 0) {
		alert("Carton박스 무게는 숫자만 가능합니다.");
		frm.cartoonboxweight.focus();
		return false;
	}

	return true;
}

function DeleteBox(frm) {
	if (confirm("삭제하시겠습니까?") == true) {
		frm.mode.value = "deletedetail";
		frm.submit();
	}
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>"><%=CTX_SEARCH%><br><%= CTX_conditional %></td>
	<td align="left">
		ShopID :
		<% if (C_IS_SHOP) then %>
			<%= shopid %>
		<% else %>
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		<% end if %>
		&nbsp;
		<select class="select" name="dateType">
			<option value="B" <%= CHKIIF(dateType="B", "selected", "") %> >발주일</option>
			<option value="C" <%= CHKIIF(dateType="C", "selected", "") %> >출고일</option>
		</select> :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<%= CTX_Order_Status %> :
		<% if (C_IS_SHOP) then %>
			<%= CTX_Shipped %>&nbsp;<%= CTX_after %>&nbsp;ALL
		<% else %>
			<select name="statecd" class="select">
				<option value="">ALL
				<option value=" " <% if statecd=" " then response.write "selected" %> ><%= CTX_in_process %>
				<option value="0" <% if statecd="0" then response.write "selected" %> ><%= CTX_Register %>
				<option value="1" <% if statecd="1" then response.write "selected" %> ><%= CTX_Confirmed %>
				<option value="2" <% if statecd="2" then response.write "selected" %> ><%= CTX_Payment_waiting %>
				<option value="5" <% if statecd="5" then response.write "selected" %> ><%= CTX_Packing_in_Process %>
				<option value="6" <% if statecd="6" then response.write "selected" %> ><%= CTX_Shipment_Standby %>
				<option value="7" <% if statecd="7" then response.write "selected" %> ><%= CTX_Shipped %>
				<option value="8" <% if statecd="8" then response.write "selected" %> ><%= CTX_preparing %>
				<option value="9" <% if statecd="9" then response.write "selected" %> ><%= CTX_stocked %>
				<option value="">========
				<option value="A" <% if statecd="A" then response.write "selected" %> ><%= CTX_Shipped %>&nbsp;<%= CTX_before %>&nbsp;ALL
				<option value="C" <% if statecd="C" then response.write "selected" %> ><%= CTX_Shipped %>&nbsp;<%= CTX_after %>&nbsp;ALL
			</select>
		<% end if %>
	</td>
	<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="<%=CTX_SEARCH%>" onClick="regsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<%= CTX_Order_code %> : <input type="text" class="text" name="baljucode" value="<%= baljucode %>" size="10" maxlength="8">
		<%= CTX_Brand %> : <% drawSelectBoxDesignerwithName "brandid", brandid %>
		&nbsp;<%= CTX_Item_Code %> : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="12">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<%= CTX_INNERBOX_NO %> : <input type="text" class="text" name="innerboxno" value="<%= innerboxno %>" size="10" maxlength="10">
		<%= CTX_INNERBOX %>&nbsp;<%= CTX_Barcode %> : <input type="text" class="text" name="innerboxbarcode" value="<%= innerboxbarcode %>" onKeyPress="if(window.event.keyCode==13) regsubmit('');" size="23" maxlength="19">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<%= CTX_CARTONBOX_NO %> : <input type="text" class="text" name="cartoonboxno" value="<%= cartoonboxno %>" size="10" maxlength="10">
		<%= CTX_CARTONTRAC_NO %> : <input type="text" class="text" name="cartonboxsongjangno" value="<%= cartonboxsongjangno %>" size="20" maxlength="20">
		<%= CTX_CARTONBOX %>&nbsp;<%= CTX_Barcode %> : <input type="text" class="text" name="cartonboxbarcode" value="<%= cartonboxbarcode %>" onKeyPress="if(window.event.keyCode==13) regsubmit('');" size="23" maxlength="19">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
     	<%= ctx_shop %>&nbsp;<%= CTX_divide %> :
     	<input type="radio" name="shopdiv" value="" <% if shopdiv="" then response.write "checked" %> >ALL

     	<% if (Not C_IS_SHOP) then %>
			<input type="radio" name="shopdiv" value="direct" <% if shopdiv="direct" then response.write "checked" %> ><%= CTX_direct_store %>
			<input type="radio" name="shopdiv" value="franchisee" <% if shopdiv="franchisee" then response.write "checked" %> ><%= CTX_franchise %>
			<input type="radio" name="shopdiv" value="foreign" <% if shopdiv="foreign" then response.write "checked" %> ><%= CTX_Foreign_store %>
			<input type="radio" name="shopdiv" value="buy" <% if shopdiv="buy" then response.write "checked" %> ><%= CTX_wholesale %>
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<%= CTX_search_result %> : <b><%= oshopbalju.FTotalCount %></b>
		&nbsp;
		<%= CTX_page %> : <b><%= page %> / <%= oshopbalju.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=80><%= CTX_Real_Order_Date %></td>
	<td width=60><%= CTX_INNERBOX_NO %></td>
	<td width=60>Inner<br>박스중량</td>
	<td width=60><%= CTX_CARTONBOX_NO %></td>
	<td width=60>Carton<br>박스중량</td>
	<td><%= CTX_real_Order_code %></td>
	<td><%= CTX_Order_code %></td>
	<td><%= CTX_Status %></td>
	<td><%= CTX_Shipment_Date %></td>
	<td>Inner<br>운송장번호</td>
	<td>Carton<br>택배사</td>
	<td>Carton<br>운송장번호</td>

	<% if C_IS_SHOP or C_ADMIN_USER then %>
		<td>도착확인</td>
		<td>비고</td>
	<% end if %>

	<td><%= CTX_Barcode %></td>
</tr>
<% if oshopbalju.FResultCount >0 then %>
<% for i=0 to oshopbalju.FResultcount-1 %>
<form name="frmModiPrc_<%= i %>" method="post" action="/admin/fran/cartoonbox_process.asp">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mode" value="modifycartoondetail">
	<input type="hidden" name="masteridx" value="<%= oshopbalju.FItemList(i).Fcartoonmasteridx %>">
	<input type="hidden" name="detailidx" value="<%= oshopbalju.FItemList(i).Fcartoondetailidx %>">
	<input type="hidden" name="baljudate" value="<%= oshopbalju.FItemList(i).Fbaljudate %>">
	<input type="hidden" name="shopid" value="<%= oshopbalju.FItemList(i).Fbaljuid %>">
	<input type="hidden" name="innerboxno" value="<%= oshopbalju.FItemList(i).Fboxno %>">
	<input type="hidden" name="baljunum" value="<%= oshopbalju.FItemList(i).Fbaljunum %>">
	<input type="hidden" name="page" value="<%= page %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><%= oshopbalju.FItemList(i).Fbaljudate %></td>
	<td align="center">
		<%
		if (oshopbalju.FItemList(i).Fboxno <> "0") then
			response.write oshopbalju.FItemList(i).Fboxno
		end if
		%>
	</td>
	<td align="center">
		<%
		if (oshopbalju.FItemList(i).Finnerboxweight <> "") then
			oshopbalju.FItemList(i).Finnerboxweight = FormatNumber(oshopbalju.FItemList(i).Finnerboxweight, 2)
		end if
		%>
		<input type="text" class="text" name="innerboxweight" value="<%= oshopbalju.FItemList(i).Finnerboxweight %>" size="3" maxlength="6" style="text-align:right">
	</td>
	<td align="center">
		<input type="text" class="text" name="cartoonboxno" value="<%= oshopbalju.FItemList(i).Fcartoonboxno %>" size="3" maxlength="6" style="text-align:right">
	</td>
	<td align="center">
		<%
		if (oshopbalju.FItemList(i).Fcartoonboxweight <> "") then
			oshopbalju.FItemList(i).Fcartoonboxweight = FormatNumber(oshopbalju.FItemList(i).Fcartoonboxweight, 2)
		end if
		%>
		<input type="text" class="text" name="cartoonboxweight" value="<%= oshopbalju.FItemList(i).Fcartoonboxweight %>" size="3" maxlength="6" style="text-align:right">
	</td>
	<td align="center"><%= oshopbalju.FItemList(i).Fbaljunum %></td>
	<td align="center"><%= oshopbalju.FItemList(i).Fbaljucode %></td>
	<td align="center">
		<font color="<%= oshopbalju.FItemList(i).GetStateColor %>"><%= oshopbalju.FItemList(i).GetStateName %></font>
	</td>
	<td align="center"><%= oshopbalju.FItemList(i).Fchulgodate %></td>
	<td align="center">
		<input type="text" class="text" name="innerboxsongjangno" value="<%= oshopbalju.FItemList(i).Fboxsongjangno %>" size="16" maxlength="20" style="text-align:right">
	</td>
	<td align="center">
		<% drawSelectBoxDeliverCompany "cartonboxsongjangdiv", oshopbalju.FItemList(i).Fcartonboxsongjangdiv %>
	</td>
	<td align="center">
		<input type="text" class="text" name="cartonboxsongjangno" value="<%= oshopbalju.FItemList(i).Fcartonboxsongjangno %>" size="16" maxlength="20" style="text-align:right">
		<% if (oshopbalju.FItemList(i).Ffindurl <> "") then %>
		<input type="button" class="button" value="추적" onclick="window.open('<%= oshopbalju.FItemList(i).Ffindurl + oshopbalju.FItemList(i).Fcartonboxsongjangno %>');">
		<% end if %>
	</td>

	<% if C_IS_SHOP or C_ADMIN_USER then %>
		<td align="center">
			<% If Not IsNull(oshopbalju.FItemList(i).FshopReceive) Then %>
			<% If (oshopbalju.FItemList(i).FshopReceive = "N") Then %>
			<input type="button" class="button" value=" 확인 " onClick="SetRecv(frmModiPrc_<%= i %>)">
			<% Else %>
			<%= oshopbalju.FItemList(i).FshopReceiveUserID %>
			<% End If %>
			<% End If %>
		</td>
		<td align="center">
			<input type="button" class="button" value=" 수정 " onClick="ModifyBox(frmModiPrc_<%= i %>)">
			<% if (oshopbalju.FItemList(i).Fcartoondetailidx <> "") then %>
			&nbsp;
			<!--
					<input type="button" class="button" value=" 삭제 " onClick="DeleteBox(frmModiPrc_<%= i %>)">
				-->
			<% end if %>
		</td>
	<% end if %>

	<td align="center">
		<input type="button" class="button" value="출력" onclick="printbarcode_off('PACKING', '', '', '', '', '', '<%= oshopbalju.FItemList(i).Fordermasteridx %>', '<%= oshopbalju.FItemList(i).Fboxno %>', '');">
	</td>
</tr>
</form>
<% next %>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=15 align=center>[<%= CTX_search_returns_no_results %>]</td>
</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if oshopbalju.HasPreScroll then %>
			<a href="javascript:GotoPage(<%= oshopbalju.StartScrollPage-1 %>)">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oshopbalju.StartScrollPage to oshopbalju.FScrollCount + oshopbalju.StartScrollPage - 1 %>
			<% if i>oshopbalju.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:GotoPage(<%= i %>)">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oshopbalju.HasNextScroll then %>
			<a href="javascript:GotoPage(<%= i %>)">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set oshopbalju = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
