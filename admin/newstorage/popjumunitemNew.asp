<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품 추가 팝업
' History : 이상구 생성
'			2016.03.20 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/stock/newshortagestockcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
const C_STOCK_DAY=7

''아래 두 페이지는 검색조건을 동일하게 맞춰야 한다.
''/admin/stock/newshortagestock.asp
''/admin/newstorage/popjumunitemNew.asp

dim page, mode, makerid, shopid,itemid, research, onlynotmddanjong, includepreorder, skiplimitsoldout
dim onlynotupchebeasong, onlynotDealItem, onlyusingitem, onlyusingitemoption, onlynotdanjong, soldoutover7days, onlysell, onlynottempdanjong
dim onoffgubun, idx, shortagetype, onlystockminus, changemakerid, purchasetype, itemgubun, itemname
dim chkMinusStockGubun, minusStockGubun, useoffinfo, i, shopsuplycash, buycash, cdl, cdm, cds
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate, iStartDate, iEndDate, onlyrealstockexists
dim mwdiv, excmkr, priceGbn, itemoption, barcode, sqlStr
dim autoBarcode, skipChkItemNo
	shopid = requestCheckvar(request("shopid"),32)
	page = requestCheckvar(request("page"),32)
	mode = requestCheckvar(request("mode"),32)
	itemid = requestCheckvar(request("itemid"),512)
	useoffinfo = requestCheckvar(request("useoffinfo"),32)
	research = requestCheckvar(request("research"),32)
	onlynotupchebeasong = requestCheckvar(request("onlynotupchebeasong"),32)
	onlynotDealItem = requestCheckvar(request("onlynotDealItem"),32)
	onlyusingitem = requestCheckvar(request("onlyusingitem"),32)
	onlyusingitemoption = requestCheckvar(request("onlyusingitemoption"),32)
	onlynotdanjong = requestCheckvar(request("onlynotdanjong"),32)
	soldoutover7days = requestCheckvar(request("soldoutover7days"),32)
	onoffgubun = requestCheckvar(request("onoffgubun"),32)
	idx = requestCheckvar(request("idx"),32)
	shortagetype = requestCheckvar(request("shortagetype"),32)
	onlysell = requestCheckvar(request("onlysell"),32)
	onlynottempdanjong = requestCheckvar(request("onlynottempdanjong"),32)
	onlynotmddanjong = requestCheckvar(request("onlynotmddanjong"),32)
	includepreorder = requestCheckvar(request("includepreorder"),32)
	skiplimitsoldout = requestCheckvar(request("skiplimitsoldout"),32)
	onlystockminus = requestCheckvar(request("onlystockminus"),32)
	purchasetype = requestCheckvar(request("purchasetype"),32)
	itemgubun = requestCheckvar(request("itemgubun"),32)
	itemname = requestCheckvar(request("itemname"),320)
	chkMinusStockGubun = requestCheckvar(request("chkMinusStockGubun"),32)
	minusStockGubun = requestCheckvar(request("minusStockGubun"),32)
	changemakerid = requestCheckvar(request("changesuplyer"),32)
	makerid = requestCheckvar(request("makerid"),32)
	onlyrealstockexists = requestCheckvar(request("onlyrealstockexists"),32)
	cdl = requestCheckvar(request("cdl"),3)
	cdm = requestCheckvar(request("cdm"),3)
	cds = requestCheckvar(request("cds"),3)
	mwdiv = requestCheckVar(request("mwdiv"),32)
	excmkr = requestCheckVar(request("excmkr"),32)
	priceGbn = requestCheckVar(request("priceGbn"),32)
	barcode = requestCheckVar(request("barcode"),20)
	autoBarcode = requestCheckVar(request("autoBarcode"),20)
    skipChkItemNo = requestCheckVar(request("skipChkItemNo"),20)

	if (changemakerid = "") then
		changemakerid = requestCheckvar(request("changemakerid"),32)
	end if
	if (makerid = "") then
		makerid = requestCheckvar(request("suplyer"),32)
	end if

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	arrItemid = ""
	iA = 0
	do while iA <= ubound(arrTemp)
		if trim(arrTemp(iA))<>"" then
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.[0]');history.back();</script>"
				Response.Write "[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다."
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	if (arrItemid <> "") then
		itemid = left(arrItemid,len(arrItemid)-1)
	end if
end if

if (research<>"on") then
	excmkr = "Y"
    'shortagetype = "14day"
    'onlynotmddanjong = "on"
    'includepreorder = "on"
end if

if (research<>"on") and (onlyrealstockexists = "") then
'	onlyrealstockexists = "on"
end if
if (research<>"on") and (onlynotupchebeasong = "") then
	onlynotupchebeasong = "on"
end if
if (research<>"on") and (onlynotDealItem = "") then
	onlynotDealItem = "on"
end if
if (research<>"on") and (onlyusingitem = "") then
	onlyusingitem = "on"
end if
if (research<>"on") and (onlyusingitemoption="") then
	onlyusingitemoption = "on"
end if
if (research<>"on") and (onlynotdanjong = "") then
	onlynotdanjong = "on"
end if
if (research<>"on") and (onoffgubun="") then
	onoffgubun = "online"
end if
if (research<>"on") and (itemgubun="") then
	itemgubun = "10"
end if

if page="" then page=1
if mode="" then mode="bybrand"

'상품코드 유효성 검사(2008.07.31;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

if trim(barcode)<>"" then

	'//바코드가 있을경우, 범용바코드는 필수로 검색
	sqlStr = "select top 1"
	sqlStr = sqlStr + " itemgubun,shopitemid,itemoption"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
	sqlStr = sqlStr + " where extbarcode='" + trim(barcode) + "'"

	'response.write sqlStr & "<Br>"
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		itemgubun = rsget("itemgubun")
		itemid = rsget("shopitemid")
		itemoption = rsget("itemoption")
	end if
	rsget.Close

    if itemid = "" then
    	itemgubun 	= BF_GetItemGubun(barcode)
    	itemid 		= BF_GetItemId(barcode)
    	itemoption 	= BF_GetItemOption(barcode)
    end if
end if

dim oshortagestock
set oshortagestock  = new CShortageStock
	oshortagestock.FPageSize = 50
	oshortagestock.FCurrPage = page
	oshortagestock.FRectOnlySell			= onlysell
	oshortagestock.FRectOnlyUsingItem		= onlyusingitem
	oshortagestock.FRectOnlyUsingItemOption	= onlyusingitemoption
	oshortagestock.FRectOnlyNotUpcheBeasong	= onlynotupchebeasong
	oshortagestock.FRectOnlynotDealItem		= onlynotDealItem
	oshortagestock.FRectShortage7days		= chkIIF(shortagetype="7day","on","")
	oshortagestock.FRectShortage14days		= chkIIF(shortagetype="14day","on","")
	oshortagestock.FRectShortageRealStock	= chkIIF(shortagetype="5under","on","")
	oshortagestock.FRectOnlyNotDanjong		= onlynotdanjong
	oshortagestock.FRectOnlyNotTempDanjong	= onlynottempdanjong
	oshortagestock.FRectOnlyNotMDDanjong	= onlynotmddanjong
	oshortagestock.FRectIncludePreOrder		= includepreorder
	oshortagestock.FRectSkipLimitSoldOut	= skiplimitsoldout
	oshortagestock.FRectOnlyStockMinus		= onlystockminus
	oshortagestock.FRectPurchaseType		= purchasetype
	oshortagestock.FRectMakerid				= makerid
	oshortagestock.FRectItemId				= itemid
	oshortagestock.FRectItemOption			= itemoption
	oshortagestock.FRectItemGubun			= itemgubun
	oshortagestock.FRectItemName			= itemname
	oshortagestock.FRectonlyrealstockexists			= onlyrealstockexists
	oshortagestock.FRectCD1   = cdl
	oshortagestock.FRectCD2   = cdm
	oshortagestock.FRectCD3   = cds

	oshortagestock.FRectMWDiv				= mwdiv
	oshortagestock.FRectExcMkr				= excmkr

	''온라인상품 테이블에는 없고, 오프라인상품 테이블에만 옵션이 있는 경우(업체에서 옵션을 변경한 경우)
	''if (itemgubun = "10") and (itemid <> 538402) then

	if (chkMinusStockGubun = "Y") then
		oshortagestock.FRectMinusStockGubun			= minusStockGubun
	end if

	if (itemgubun = "10") and (useoffinfo = "") then
        if (research <> "") or (makerid <> "") then
		    oshortagestock.GetShortageItemListOnline
        end if
	else
		oshortagestock.GetShortageItemListOffline
	end if

if (yyyy1="") then
    nowdate = Left(CStr(DateAdd("d",now(),-2)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

    nowdate = Left(CStr(DateAdd("d",now(),+2)),10)
	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

iStartDate  = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
iEndDate    = Left(CStr(DateSerial(yyyy2,mm2,dd2)),10)

dim ogroup,opartner, IsSimpleVAT
IsSimpleVAT = False
if (makerid <> "") then
	set opartner = new CPartnerUser
	opartner.FRectDesignerID = makerid
	opartner.GetOnePartnerNUser

	if opartner.FResultCount > 0 then
		set ogroup = new CPartnerGroup
		ogroup.FRectGroupid = opartner.FOneItem.FGroupid
		ogroup.GetOneGroupInfo

		if (ogroup.FOneItem.Fjungsan_gubun = "간이과세") then
			IsSimpleVAT = True
			response.write "<script type='text/javascript'>"
			response.write "	alert('======================================================\n\n\n간이과세 업체입니다.\n\n\n매장에 출고할 수 없고\n해외출고도 안되며\n기타출고할 때에도 재무팀에 문의후 출고해야 합니다.\n\n\n======================================================')"
			response.write "</script>"
		end if
	end if
end if

dim OffMwMarginArr, centermwdiv

if trim(barcode)<>"" then
	'// 뭐지?
	''itemgubun 	= ""
	itemid 		= ""
	itemoption 	= ""
end if
%>

<script type="text/javascript">

function popOffItemEdit(ibarcode){
	var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'adminitemselledit','width=500,height=600,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function enablebrand(bool){
	//document.frm.designer.disabled = bool;
}

function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

function search(frm){
	/*
	if ((frm.makerid.value.length<1)){
		if ((frm.mode[0].checked)&&(frm.designer.value.length<1)){
			alert('브랜드를 선택 하세요.');
			frm.designer.focus();
			return;
		}
	}
	*/
	frm.page.value=1;
	frm.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function AddArr(){
	var upfrm = document.frmArrupdate;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	upfrm.jungsan_temp.value = "";
	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.itemnamearr.value = "";
	upfrm.itemoptionnamearr.value = "";
	upfrm.designerarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				// 2018.11.29 한용민
				if (frm.jungsan_gubun.value=="간이과세"){
					if (upfrm.jungsan_temp.value!='간이과세'){
						upfrm.jungsan_temp.value = frm.jungsan_gubun.value;
						alert('======================================================\n\n\n간이과세 업체입니다.\n\n\n매장에 출고할 수 없고\n해외출고도 안되며\n기타출고할 때에도 재무팀에 문의후 출고해야 합니다.\n\n\n======================================================');
						frm.jungsan_gubun.focus();
					}
				}

                <% if (skipChkItemNo <> "Y") then %>
				if (!IsInteger(frm.itemno.value)){
					alert('갯수는 정수만 가능합니다.');
					frm.itemno.focus();
					return;
				}

				if (frm.itemno.value=="0"){
					alert('수량을 입력하세요.');
					frm.itemno.focus();
					return;
				}
                <% end if %>

				//방화벽에서 막혀서 치환함. 실제 저장 할때 디비에서 옵션명 가져옴.	'/2016.05.26 한용민 추가
				if (frm.itemoptionname.value=='Script' || frm.itemoptionname.value=='script'){
					frm.itemoptionname.value = frm.itemoptionname.value.split("Script").join("s_cript")
					frm.itemoptionname.value = frm.itemoptionname.value.split("script").join("s_cript")
				}

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
				upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
				upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
				upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
				upfrm.itemnamearr.value = upfrm.itemnamearr.value + frm.itemname.value + "|";
				upfrm.itemoptionnamearr.value = upfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
				upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
				upfrm.mwdivarr.value = upfrm.mwdivarr.value + frm.mwdiv.value + "|";

				<% if False and (useoffinfo <> "on") then %>
				if (frm.itemmwdiv.value == "U") {
					alert("업체배송 상품은 검색조건에서 '오프상품정보(10) 사용' 을 체크 후 추가하세요.");
					return;
				}
				<% end if %>
			}
		}
	}

	opener.ReActItems('<%= idx %>', upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.buycasharr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value,upfrm.mwdivarr.value);

	//초기화
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				frm.cksel.checked = false;
				frm.itemno.value="0"


			}
		}
	}

}

window.onload = function() {
	var cksel, itemno;
	<%
	if (autoBarcode = "Y") and barcode <> "" then
		if (oshortagestock.FTotalCount < 1) then
			response.write "alert('===========================================================\n\n검색된 상품이 없습니다.\n\n===========================================================');"
		elseif (oshortagestock.FTotalCount > 1) then
			response.write "alert('===========================================================\n\n복수의 상품이 검색되었습니다.\n\n===========================================================');"
		else
	%>
	try {
		cksel = frmBuyPrc_0.cksel;
		itemno = frmBuyPrc_0.itemno;
		itemno.value = 1;
		cksel.checked = true;
		AnCheckClick(cksel);
		AddArr();
		frm.barcode.select();
	} catch(e) { }
	<%
		end if
	end if
	%>
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="priceGbn" value="<%= priceGbn %>">
<input type="hidden" name="skipChkItemNo" value="<%= skipChkItemNo %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if (changemakerid <> "Y") then %>
	<input type="hidden" name="makerid" value="<%= makerid %>" >
<% else %>
	<input type="hidden" name="changemakerid" value="Y" >
<% end if %>

<input type="hidden" name="shopid" value="<%= shopid %>" >
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<% if (changemakerid <> "Y") then %>
		브랜드 : <b><%= makerid %></b>
		<% else %>
		브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		<% end if %>
		&nbsp;
		|
		&nbsp;
		구분 :
		<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
		<!--
		<select class="select" name="itemgubun">
			<option value="10" <% if (itemgubun = "10") then %>selected<% end if %> >온라인(10)</option>
			<option value="90" <% if (itemgubun = "90") then %>selected<% end if %> >오프라인(90)</option>
			<option value="70" <% if (itemgubun = "70") then %>selected<% end if %> >사은품 등(70)</option>
			<option value="80" <% if (itemgubun = "80") then %>selected<% end if %> >사은품 등(80)</option>
			<option value="XX" <% if (itemgubun = "XX") then %>selected<% end if %> >기타</option>
		</select>
		-->
		&nbsp;
		|
		&nbsp;
		<input type=checkbox name="onlyusingitem" <% if onlyusingitem = "on" then response.write "checked" %> >사용상품만
		<input type=checkbox name="onlyusingitemoption" <% if onlyusingitemoption = "on" then response.write "checked" %> >사용옵션만
		<input type=checkbox name="onlysell" <% if onlysell = "on" then response.write "checked" %> >판매상품만
		<input type=checkbox name="onlynotupchebeasong" <% if onlynotupchebeasong = "on" then response.write "checked" %> >업체배송제외
		<input type=checkbox name="onlyrealstockexists" <% if onlyrealstockexists = "on" then response.write "checked" %> >재고0인상품제외
		<input type=checkbox name="onlynotDealItem" <% if onlynotDealItem = "on" then response.write "checked" %> >딜상품제외
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:search(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
        부족구분:
        <input type="radio" name="shortagetype" value="" <% if shortagetype="" then response.write "checked" %> >전체
        <input type="radio" name="shortagetype" value="7day" <% if shortagetype="7day" then response.write "checked" %> ><%= C_STOCK_DAY %>일후재고부족
		<input type="radio" name="shortagetype" value="14day" <% if shortagetype="14day" then response.write "checked" %> ><%= C_STOCK_DAY*2 %>일후재고부족
        <input type="radio" name="shortagetype" value="5under" <% if shortagetype="5under" then response.write "checked" %> >실사유효재고 5이하
		&nbsp;
		|
		&nbsp;
		<input type=checkbox name="onlynotdanjong" <% if onlynotdanjong = "on" then response.write "checked" %> >단종제외(옵션포함)
		<input type=checkbox name="onlynottempdanjong" <% if onlynottempdanjong = "on" then response.write "checked" %> >일시품절제외(옵션포함)
		<input type=checkbox name="onlynotmddanjong" <% if onlynotmddanjong = "on" then response.write "checked" %> >MD품절제외(옵션포함)
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		상품코드 : <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size=16 maxlength=16>
		&nbsp;
		|
		&nbsp;
        <input type=checkbox name="includepreorder" <% if includepreorder = "on" then response.write "checked" %> >기주문포함부족만
        <!--
        <input type=checkbox name="skiplimitsoldout" <% if skiplimitsoldout = "on" then response.write "checked" %> >한정&판매중지제외
        -->
        <input type=checkbox name="onlystockminus" <% if onlystockminus = "on" then response.write "checked" %> >실사유효재고마이너스만
		&nbsp;
		|
		&nbsp;
        구매유형 : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
		&nbsp;
        <input type=checkbox name="useoffinfo" <% if useoffinfo = "on" then response.write "checked" %> > 오프상품정보(10) 사용
		<br>
		<input type="checkbox" name="chkMinusStockGubun" value="Y" <%if (chkMinusStockGubun = "Y") then %>checked<% end if %> >
		재고구분 :
		<select class="select" name="minusStockGubun">
			<option value="real" <%if (minusStockGubun = "real") then %>selected<% end if %> >실사유효재고</option>
			<option value="check" <%if (minusStockGubun = "check") then %>selected<% end if %> >재고파악재고</option>
			<option value="may" <%if (minusStockGubun = "may") then %>selected<% end if %> >예상재고</option>
		</select>
		마이너스만
		&nbsp;
		바코드 :
		<input type="text" name="barcode" value="<%= barcode %>" size="16" maxlength="20" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) NextPage('1');">
		<input type="checkbox" name="autoBarcode" value="Y" <%if (autoBarcode = "Y") then %>checked<% end if %> > 바코드 자동입력
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		&nbsp;
		거래구분 :
		<select class="select" name="mwdiv">
			<option value="">-선택-</option>
			<option value="M" <% if (mwdiv = "M") then %>selected<% end if %> >매입</option>
			<option value="W" <% if (mwdiv = "W") then %>selected<% end if %> >위탁</option>
			<option value="U" <% if (mwdiv = "U") then %>selected<% end if %> >업체</option>
			<option value="Z" <% if (mwdiv = "Z") then %>selected<% end if %> >미지정</option>
		</select>
		&nbsp;
		<input type="checkbox" class="checkbox" name="excmkr" value="Y" <%= CHKIIF(excmkr="Y", "checked", "")%> > 아이띵소제외
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% if IsSimpleVAT then %>
			<strong><font color="red">간이과세 업체입니다. 출고 불가 브랜드 입니다.
			<br>매장에 출고할 수 없고, 해외출고도 안되며, 기타출고할 때에도 재무팀에 문의후 출고해야 합니다.</font></strong>
		<% end if %>
	</td>
	<td align="right">
		<input type="button" class="button" value="전체선택" onClick="AnSelectAllFrame(true)">
		<input type="button" class="button" value="선택 아이템 추가" onclick="AddArr()">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		검색결과 : <b><%= oshortagestock.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= Page %> / <%= oshortagestock.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="ckall" onClick="ckAll(this)"></td>
	<td width="50">이미지</td>
	<td width="80">브랜드ID</td>
	<td width="90">상품코드</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td width="35">할인<br>여부</td>
	<td width="90">할인기간</td>
	<td width="45">판매가</td>
	<td width="45">매입가<br />(마진X)</td>
    <td width="45">거래<br />구분</td>
	<td width="45">마진<br />(센터)</td>
	<td width="60">판매일수/<br />3개월판매<br />(오픈일)</td>
	<td width="45">수량</td>
	<td>비고</td>
</tr>
<% if oshortagestock.FResultCount > 0 then %>
<% for i=0 to oshortagestock.FResultCount -1 %>

<form name="frmBuyPrc_<%= i %>" style="margin:0px;">
<input type="hidden" name="jungsan_gubun" value="<%= oshortagestock.FItemList(i).Fjungsan_gubun %>">
<input type="hidden" name="itemgubun" value="<%= oshortagestock.FItemList(i).Fitemgubun %>">
<input type="hidden" name="itemid" value="<%= oshortagestock.FItemList(i).Fitemid %>">
<input type="hidden" name="itemoption" value="<%= oshortagestock.FItemList(i).Fitemoption %>">
<input type="hidden" name="itemname" value="<%= oshortagestock.FItemList(i).FItemName %>">
<input type="hidden" name="itemoptionname" value="<%= oshortagestock.FItemList(i).FItemOptionName %>">
<input type="hidden" name="desingerid" value="<%= oshortagestock.FItemList(i).FMakerid %>">
<% if (priceGbn = "saleprice") then %>
<input type="hidden" name="sellcash" value="<%= oshortagestock.FItemList(i).FSellcash %>">
<% else %>
<input type="hidden" name="sellcash" value="<%= oshortagestock.FItemList(i).Forgprice %>">
<% end if %>
<input type="hidden" name="suplycash" value="<%= chkIIF(oshortagestock.FItemList(i).IsOffContractExist, oshortagestock.FItemList(i).GetOffContractBuycash, oshortagestock.FItemList(i).FBuycash) %>">
<input type="hidden" name="buycash" value="<%= chkIIF(oshortagestock.FItemList(i).IsOffContractExist, oshortagestock.FItemList(i).GetOffContractBuycash, oshortagestock.FItemList(i).FBuycash) %>">
<input type="hidden" name="mwdiv" value="<%= chkIIF(oshortagestock.FItemList(i).IsOffContractExist, oshortagestock.FItemList(i).GetOffContractCenterMW, oshortagestock.FItemList(i).Fmwdiv) %>">
<input type="hidden" name="itemmwdiv" value="<%= oshortagestock.FItemList(i).Fmwdiv %>">

<% if (oshortagestock.FItemList(i).Foptionusing="N") or (oshortagestock.FItemList(i).Fisusing="N") then %>
<tr bgcolor="<%= adminColor("gray") %>">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
	<td rowspan=2 align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td rowspan=2>
		<% if oshortagestock.FItemList(i).Fitemgubun = "10" then %>
			<img src="<%= oshortagestock.FItemList(i).FimageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'">
		<% else %>
			<img src="<%= oshortagestock.FItemList(i).FOffimgSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'">
		<% end if %>
	</td>
	<td height="25"><%= oshortagestock.FItemList(i).FMakerid %></td>
	<% if oshortagestock.FItemList(i).FItemGubun<>"10" then %>
	<td ><a href="javascript:popOffItemEdit('<%= oshortagestock.FItemList(i).GetBarCode %>')"><%= oshortagestock.FItemList(i).GetBarCodeBoldStr %></a></td>
	<% else %>
	<td ><a href="javascript:PopItemSellEdit('<%= oshortagestock.FItemList(i).FItemID %>');"><%= oshortagestock.FItemList(i).GetBarCodeBoldStr %></a></td>
	<% end if %>
	<td ><a href="/admin/stock/itemcurrentstock.asp?itemid=<%= oshortagestock.FItemList(i).FItemID %>&itemoption=<%= oshortagestock.FItemList(i).FItemOption %>" target=_blank ><%= oshortagestock.FItemList(i).FItemName %></a></td>
	<td ><%= oshortagestock.FItemList(i).FItemOptionName %></td>
	<td rowspan=2 align="center">
		<% if (oshortagestock.FItemList(i).FSailYn="Y") then %>
		<font color=red>
			<% if (oshortagestock.FItemList(i).Forgprice<>0) then %>
			<%= CLng((oshortagestock.FItemList(i).Forgprice-oshortagestock.FItemList(i).Fsellcash)/oshortagestock.FItemList(i).Forgprice*100) %> %
			<% end if %>
		</font>
		<% end if %>
	</td>
	<td rowspan=2 align="center">
		<%= Replace(oshortagestock.FItemList(i).FsaleStr, "~", "~<br>") %>
	</td>
	<td rowspan=2 align=right>
		<%= CHKIIF(oshortagestock.FItemList(i).FSellcash<>oshortagestock.FItemList(i).Forgprice, FormatNumber(oshortagestock.FItemList(i).Forgprice,0) & "<br />" & "=&gt;", "")%>
		<%= FormatNumber(oshortagestock.FItemList(i).FSellcash,0) %>
	</td>
	<td rowspan=2 align=right>
		<%= FormatNumber(oshortagestock.FItemList(i).FBuycash,0) %>
		<% if oshortagestock.FItemList(i).IsOffContractExist then %>
		<br>-&gt;<font color="blue"><%= FormatNumber(oshortagestock.FItemList(i).GetOffContractBuycash,0) %></font>
		<% end if %>
	</td>
    <td rowspan=2 align=center>
        <font color="<%= oshortagestock.FItemList(i).getMwDivColor %>"><%= oshortagestock.FItemList(i).getMwDivName %></font>
    </td>
	<td rowspan=2 align=center>
	<font color="<%= oshortagestock.FItemList(i).getMwDivColor %>"><%= oshortagestock.FItemList(i).getMwDivName %></font><br>
	<% if oshortagestock.FItemList(i).Forgprice<>0 then %>
	<%= 100-(CLng(oshortagestock.FItemList(i).FBuycash/oshortagestock.FItemList(i).Forgprice*10000)/100) %> %
	<% end if %>
	<% if oshortagestock.FItemList(i).IsOffContractExist then %>
	<br>-&gt;<font color="blue"><%= oshortagestock.FItemList(i).GetOffContractMargin %>%</font>
	<% end if %>
	<%
	if (oshortagestock.FItemList(i).Fmwdiv <> "M") and (oshortagestock.FItemList(i).Fmwdiv <> "W") then
		centermwdiv = ""
		if Not IsNull(oshortagestock.FItemList(i).FOffMwMargin) then
			OffMwMarginArr = Split(oshortagestock.FItemList(i).FOffMwMargin, "_")
			if UBound(OffMwMarginArr) = 2 then
				centermwdiv = OffMwMarginArr(0)
				response.write "(" & centermwdiv & ")"
			end if
		end if
		if centermwdiv = "" then
			response.write "<font color='red'>(계약X)</font>"
		end if
	end if
	%>
	</td>
	<td rowspan=2 align="center">
		<% if (oshortagestock.FItemList(i).FsellSTDateStr <> "") then %>
		<%= oshortagestock.FItemList(i).FsellSTDateStr %>/<%= oshortagestock.FItemList(i).FthreeMonthSellNo %>
		<% if Not IsNull(oshortagestock.FItemList(i).FsellSTDate) then %>
		<% if DateDiff("m", oshortagestock.FItemList(i).FsellSTDate, Now) <= 3 then %>
		<br />(<%= Right(Left(oshortagestock.FItemList(i).FsellSTDate, 10),5) %>)
		<% else %>
		<br />(<%= Left(oshortagestock.FItemList(i).FsellSTDate, 7) %>)
		<% end if %>
		<% end if %>
		<% end if %>
	</td>
	<td rowspan=2>
		<% if oshortagestock.FItemList(i).Frealstock<0 and oshortagestock.FItemList(i).Fsell7days=0 then %>
		<input type="text" class="text" name="itemno" value="0" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);">
	    <% elseif oshortagestock.FItemList(i).GetNdayShortageNo(14) < 0 then %>
	    <input type="text" class="text" name="itemno" value="<%= (oshortagestock.FItemList(i).GetNdayShortageNo(14))*-1 %>" size="4" maxlength="5" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);">
	    <% else %>
	    <input type="text" class="text" name="itemno" value="0" size="4" maxlength="5" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);">
	    <% end if %>
	</td>
	<td rowspan=2 <%= CHKIIF(oshortagestock.FItemList(i).Fpreorderno<>0, "bgcolor='#DDFFDD'", "") %> align="center">
		<% if oshortagestock.FItemList(i).IsOffContractExist then %>
		<font color="blue">오프계약</font>
		<% end if %>

		<%= fnColor(oshortagestock.FItemList(i).Fdanjongyn,"dj") %>
        <% if oshortagestock.FItemList(i).Foptdanjongyn="S" then %>
		<font color="#3333CC">옵션부족</font>
		<% end if %>
        <% if oshortagestock.FItemList(i).Foptdanjongyn="Y" then %>
		<font color="#33CC33">옵션단종</font><br>
		<% end if %>
        <% if oshortagestock.FItemList(i).Foptdanjongyn="M" then %>
		<font color="#CC3333">옵션MD</font><br>
		<% end if %>
		<br>

		<!-- 재고부족의 경우 재입고예정일 표시 -->
		<% if (oshortagestock.FItemList(i).Fdanjongyn = "S") or (oshortagestock.FItemList(i).Foptdanjongyn = "S") then %>
			<% if ((Not IsNull(oshortagestock.FItemList(i).FreipgoMayDate)) and (Left(oshortagestock.FItemList(i).FreipgoMayDate, 10) >= iStartDate) and (Left(oshortagestock.FItemList(i).FreipgoMayDate, 10) <= iEndDate)) then %>
				<%= Left(oshortagestock.FItemList(i).FreipgoMayDate, 10) %><br>
			<% elseif (Not IsNull(oshortagestock.FItemList(i).FreipgoMayDate)) then %>
				<font color="gray"><%= Left(oshortagestock.FItemList(i).FreipgoMayDate, 10) %></font><br>
			<% end if %>
		<% end if %>

		<% if oshortagestock.FItemList(i).Foptionusing="N" then %>
		<font color="red">옵션x</font><br>
		<% end if %>
		<% if oshortagestock.FItemList(i).IsSoldOut then %>
		<font color="red">품절</font><br>
		<% end if %>
		<% if oshortagestock.FItemList(i).Flimityn="Y" then %>
		<font color="blue">한정(<%= oshortagestock.FItemList(i).getOptionLimitNo %>)</font><br>
		<% end if %>
		<% if oshortagestock.FItemList(i).Fpreorderno<>0 then %>
			<br />
            <font color="red">기주문:
			<% if oshortagestock.FItemList(i).Fpreorderno<>oshortagestock.FItemList(i).Fpreordernofix then response.write "</br>" + CStr(oshortagestock.FItemList(i).Fpreorderno) + "->" %>
				<%= oshortagestock.FItemList(i).Fpreordernofix %><br><br/>
            </font>
		<% end if %>
        <% if oshortagestock.FItemList(i).FlastIpgoDate <> "" then %>
        최종 : <%= oshortagestock.FItemList(i).FlastIpgoDate %><br />
        <% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=4>
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>입고</td>
				<td>판매</td>
				<td>출고</td>

				<!-- 필요없다 하여 숨김, skyer9, 2016-08-11
				<td>기타</td>
				<td>CS</td>
				<td>불량</td>
				<td>오차</td>
				-->

				<td>실사유효재고</td>
				<td bgcolor="<%= adminColor("green") %>"><b>출고이전[<%= oshortagestock.FItemList(i).GetReqNotChulgoNo %>]</b></td>
				<td>예상재고</td>

				<% if oshortagestock.FItemList(i).Fmaxsellday<>7 then %>
				<td bgcolor="<%= adminColor("green") %>">On<font color="#CC1111"><%= oshortagestock.FItemList(i).Fmaxsellday %></font>일</td>
				<td bgcolor="<%= adminColor("green") %>">Off<font color="#CC1111"><%= oshortagestock.FItemList(i).Fmaxsellday %></font>일</td>
				<% else %>
				<td bgcolor="<%= adminColor("green") %>">On<%= oshortagestock.FItemList(i).Fmaxsellday %>일</td>
				<td bgcolor="<%= adminColor("green") %>">Off<%= oshortagestock.FItemList(i).Fmaxsellday %>일</td>
				<% end if %>

				<td><%= C_STOCK_DAY %>일</td>
				<td><%= C_STOCK_DAY*2 %>일</td>
				<!--
				<td>OFF준비</td>
				-->
			</tr>
			<tr bgcolor="#FFFFFF" align=center>
				<td><%= oshortagestock.FItemList(i).Ftotipgono %></td>
				<td><%= oshortagestock.FItemList(i).Ftotsellno %></td>
				<td><%= oshortagestock.FItemList(i).Ftotchulgono %></td>
				<!--
				<td></td>
				<td></td>
				<td><%= oshortagestock.FItemList(i).Ferrbaditemno %></td>
				<td><%= oshortagestock.FItemList(i).Ferrrealcheckno %></td>
				-->

				<td>
					<b>
					<% if oshortagestock.FItemList(i).Frealstock<1 then %>
					<font color="#CC1111"><b><%= oshortagestock.FItemList(i).GetCheckStockNo %></b></font>
					<% else %>
					<%= oshortagestock.FItemList(i).Frealstock %>
					<% end if %>
					</b>
				</td>

				<td>
				    <!-- 출고이전 -->
					<!--
				    <%= oshortagestock.FItemList(i).GetReqNotChulgoNo %>
					-->
					<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
						<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
							<td>ON준비</td>
							<td>OFF준비</td>
							<td>ON결제</td>
							<td>ON접수</td>
							<td>OFF접수</td>
						</tr>
						<tr bgcolor="#FFFFFF" align=center>
							<td><%= CHKIIF(oshortagestock.FItemList(i).Fipkumdiv5<>"0", oshortagestock.FItemList(i).Fipkumdiv5, "") %></td>
							<td><%= CHKIIF(oshortagestock.FItemList(i).Foffconfirmno<>"0", oshortagestock.FItemList(i).Foffconfirmno, "") %></td>
							<td><%= CHKIIF(oshortagestock.FItemList(i).Fipkumdiv4<>"0", oshortagestock.FItemList(i).Fipkumdiv4, "") %></td>
							<td><%= CHKIIF(oshortagestock.FItemList(i).Fipkumdiv2<>"0", oshortagestock.FItemList(i).Fipkumdiv2, "") %></td>
							<td><%= CHKIIF(oshortagestock.FItemList(i).Foffjupno<>"0", oshortagestock.FItemList(i).Foffjupno, "") %></td>
						</tr>
					</table>
				</td>
				<td>
					<b>
					<% if oshortagestock.FItemList(i).Frealstock + oshortagestock.FItemList(i).GetReqNotChulgoNo < 1 then %>
					<font color="#CC1111"><%= oshortagestock.FItemList(i).Frealstock + oshortagestock.FItemList(i).GetReqNotChulgoNo %></b></font>
					<% else %>
					<%= oshortagestock.FItemList(i).Frealstock + oshortagestock.FItemList(i).GetReqNotChulgoNo %>
					<% end if %>
					</b>
				</td>
				<td><%= oshortagestock.FItemList(i).Fsell7days %></td>
				<td><%= oshortagestock.FItemList(i).Foffchulgo7days %></td>


				<td>
				    <!-- 7일 -->
					<% if oshortagestock.FItemList(i).Fshortageno< 1 then %>
					<font color="#CC1111"><b><%= oshortagestock.FItemList(i).Fshortageno %></b></font>
					<% else %>
					<%= oshortagestock.FItemList(i).Fshortageno %>
					<% end if %>
				</td>
				<td>
				    <!-- N일 필요 -->
					<% if (oshortagestock.FItemList(i).GetNdayShortageNo(14))< 1 then %>
					<font color="#CC1111"><b><%= oshortagestock.FItemList(i).GetNdayShortageNo(14) %></b></font>
					<% else %>
					<%= oshortagestock.FItemList(i).GetNdayShortageNo(14) %>
					<% end if %>
				</td>
				<!--
				<td><%= oshortagestock.FItemList(i).Foffconfirmno %></td>
				-->
			</tr>
		</table>
	</td>
</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">
		<% if oshortagestock.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oshortagestock.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oshortagestock.StartScrollPage to oshortagestock.FScrollCount + oshortagestock.StartScrollPage - 1 %>
			<% if i>oshortagestock.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oshortagestock.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<form name="frmArrupdate" method="post" action="">
<input type="hidden" name="mode" value="arrins">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="itemnamearr" value="">
<input type="hidden" name="itemoptionnamearr" value="">
<input type="hidden" name="designerarr" value="">
<input type="hidden" name="mwdivarr" value="">
<input type="hidden" name="jungsan_temp" value="">
</form>

<script type="text/javascript">
	//alert('수정중');
</script>

<%
set oshortagestock = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
