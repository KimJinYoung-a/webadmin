<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 지역별 상품 설정
' History : 2010.08.03 서동석 생성
'			2010.08.05 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopLocaleItemcls.asp"-->

<%
dim designer,page, pagesize, usingyn ,research,pricediff,imageview ,itemgubun, itemid, itemname , shopitemname , gubun , nameeng
dim cdl, cdm, cds ,shopid , i ,PriceDiffExists , arrexchangerate, currencyUnit_Pos ,multipleRate , exchangeRate, countrylangcd
dim decimalPointLen, decimalPointCut
dim prdcode, generalbarcode , shopdiv , adminok

	designer    	= RequestCheckVar(request("designer"),32)
	page        	= RequestCheckVar(request("page"),9)
	pagesize       	= RequestCheckVar(request("pagesize"),9)
	usingyn     	= RequestCheckVar(request("usingyn"),1)
	research    	= RequestCheckVar(request("research"),9)
	pricediff   	= RequestCheckVar(request("pricediff"),9)
	imageview   	= RequestCheckVar(request("imageview"),9)

	itemgubun   	= RequestCheckVar(request("itemgubun"),2)
	itemid      	= RequestCheckVar(request("itemid"),9)

	itemname    	= RequestCheckVar(request("itemname"),32)
	shopitemname	= RequestCheckVar(request("shopitemname"),32)

	cdl         	= RequestCheckVar(request("cdl"),3)
	cdm         	= RequestCheckVar(request("cdm"),3)
	cds         	= RequestCheckVar(request("cds"),3)
	shopid      	= RequestCheckVar(request("shopid"),32)
	gubun      		= RequestCheckVar(request("gubun"),10)
	nameeng 		= RequestCheckVar(request("nameeng"),10)

	prdcode 		= RequestCheckVar(request("prdcode"),32)
	generalbarcode 	= RequestCheckVar(request("generalbarcode"),32)

''차후 session("ssAdminPsn")="6" : 부서번호만 사용할것.
if session("ssBctDiv")="201" or session("ssAdminPsn")="6" then
	shopid = "cafe002"
elseif session("ssBctDiv")="301" or session("ssAdminPsn")="16" then
	shopid = "cafe003"
else
end if

if C_ADMIN_USER then

''유저구분이 가맹점인경우 박아 넣는다
elseif (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

if page="" then page=1
if pagesize="" then pagesize = 100
''if research<>"on" then usingyn="Y"

decimalPointLen = 0
dim oexchangerate
set oexchangerate = new COffShopLocale
	oexchangerate.frectuserid = shopid

if shopid = "" then
	response.write "<script>alert('매장을 선택하세요');</script>"
else
	oexchangerate.fexchangeratecheck()

	shopdiv = oexchangerate.foneitem.fshopdiv
	currencyUnit_Pos = oexchangerate.foneitem.fcurrencyUnit_Pos
	multipleRate = oexchangerate.foneitem.fmultipleRate
	exchangeRate = oexchangerate.foneitem.fexchangeRate
	decimalPointLen = oexchangerate.foneitem.fdecimalPointLen
	decimalPointCut = oexchangerate.foneitem.fdecimalPointCut
    countrylangcd   = oexchangerate.foneitem.fcountrylangcd

	'/해외매장이 아닐경우
	if shopdiv <> "7" then
		adminok = false

		response.write "<script>"
		response.write "	alert('선택하신 매장이 해외매장이 아닙니다');"
		response.write "</script>"
		response.write "<font color='red'>해외매장만 사용가능</font>"
		response.end

	'/해외매장일 경우
	else
		adminok	 = true
	end if

	if oexchangerate.foneitem.fcurrencyUnit_Pos = "" or isnull(oexchangerate.foneitem.fcurrencyUnit_Pos) then response.write "<script>alert('[필수]해당매장에 화폐단위가 등록되어 있지 않습니다\n\n매장별 화폐단위와 배수는 [OFF]오프_매장관리>>오프샾리스트 에서 입력해주세요.');</script>"
	if oexchangerate.foneitem.fmultipleRate = "" or isnull(oexchangerate.foneitem.fmultipleRate) then response.write "<script>alert('[필수]해당매장에 마진배수가 등록되어 있지 않습니다\n\n매장별 화폐단위와 배수는 [OFF]오프_매장관리>>오프샾리스트 에서 입력해주세요.');</script>"
end if

dim ioffitem
set ioffitem  = new COffShopLocale
	ioffitem.FPageSize = pagesize
	ioffitem.FCurrPage = page
	ioffitem.FRectShopId = shopid
	ioffitem.FRectDesigner = designer
	ioffitem.FRectOnlyUsing = usingyn
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemID = itemid
	ioffitem.FRectItemName = html2db(itemname)
	ioffitem.FRectShopItemName = html2db(shopitemname)
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.frectgubun = gubun
	ioffitem.frectnameeng = nameeng
	ioffitem.FRectPrdCode = prdcode
	ioffitem.FRectGeneralBarcode = generalbarcode
    ioffitem.FRectMultipleRate = MultipleRate
    ioffitem.FRectExchangeRate = exchangeRate

    ioffitem.FRectcountrylangcd = countrylangcd
	if (shopid<>"") then
	    ioffitem.GetLocaleItemList()
	end if


dim isShowMultiLang : isShowMultiLang = (NOT isNULL(countrylangcd) and (countrylangcd<>"") and (countrylangcd<>"KR"))
%>

<script language='javascript'>
function isMayEng(str){
    return (str.length==getbyteLength(str))
}

function getbyteLength (str){
    var retCode = 0;
    var strLength = 0;

    for (i = 0; i < str.length; i++){
        var code = str.charCodeAt(i)
        var ch = str.substr(i,1).toUpperCase()

        code = parseInt(code)

        if ((ch < "0" || ch > "9") && (ch < "A" || ch > "Z") && ((code > 255) || (code < 0)))
            strLength = strLength + 2;
        else
            strLength = strLength + 1;
    }
    return strLength;
}

//선택한 배수 역계산
function CheckThislcprice(frm){
	frm.mrate.value = Math.round(((frm.lcprice.value / frm.ShopItemprice.value) * frm.erate.value) * 100) / 100 ;
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

//선택한 판매가 계산
function CheckThismrate(frm){
    var upfrm = document.frm;

    var cutn = upfrm.decimalPointCut.value*1;
	var pown = upfrm.decimalPointLen.value*1;
    var cutnPow = Math.pow(10, cutn)*1;

	//frm.lcprice.value = Math.round(((frm.ShopItemprice.value / frm.erate.value)* frm.mrate.value) * 100) / 100;
	var cc = Math.round(((frm.ShopItemprice.value / frm.erate.value)* frm.mrate.value) * cutnPow) / cutnPow;
	frm.lcprice.value = cc.toFixed(pown);

	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

// 새상품 추가 팝업
function addnewItem(){
	var popup_item;
	popup_item = window.open("pop_localeItem_input.asp", "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popup_item.focus();
}

function popForeignPriceBase(shopid){
    var popwin = window.open('/common/offshop/exchangerate/popForeignPriceBase.asp?shopid='+shopid,'popForeignPriceBase','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//환율관리 등록 & 수정 - 사용중지
function popexchangerate(){
    var popexchangerate = window.open('/common/offshop/exchangerate/exchangerate.asp','popexchangerate','width=1024,height=768,scrollbars=yes,resizable=yes');
    popexchangerate.focus();
}

// 환율 배수 일괄적용
function automulti(upfrm){
    if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}


	var frm;
	var cutn = upfrm.decimalPointCut.value*1;
	var pown = upfrm.decimalPointLen.value*1;
    var cutnPow = Math.pow(10, cutn)*1;

    //3.1234 * 100

		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){

					//if (frm.lcprice.value==''){
					//	alert('현지판매가 설정되지 않았습니다');
					//	frm.lcprice.focus;
					//	return;
					//}

					frm.erate.value = upfrm.exchangeRate.value
					frm.mrate.value = upfrm.multipleRate.value
					//frm.lcprice.value = Math.round(((frm.ShopItemprice.value / upfrm.exchangeRate.value)* upfrm.multipleRate.value) * 100) / 100;

					var cc = Math.round(((frm.ShopItemprice.value / upfrm.exchangeRate.value)* upfrm.multipleRate.value) * cutnPow) / cutnPow;
					frm.lcprice.value = cc.toFixed(pown);
				}
			}
		}
}

//환율일괄적용
function autoexchangeRate(upfrm){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){

					if (frm.lcprice.value==''){
						alert('현지판매가 설정되지 않았습니다');
						frm.lcprice.focus;
						return;
					}

					frm.lcprice.value = frm.lcprice.value / frm.exchangeRate.value;


				}
			}
		}
}

//마진배수일괄적용
function automultipleRate(upfrm){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){

					if (frm.lcprice.value==''){
						alert('현지판매가 설정되지 않았습니다');
						frm.lcprice.focus;
						return;
					}

					frm.lcprice.value = frm.lcprice.value * upfrm.multipleRate.value;


				}
			}
		}
}

//기본판매가일괄적용
function autoShopItemprice(upfrm){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					frm.lcprice.value = frm.ShopItemprice.value

				}
			}
		}
}

function autoShopItemNameNOptionName(upfrm,tp){
    return; //이문재 이사요청
    autoShopItemName(upfrm,tp);
    autoshopitemoptionname(upfrm,tp);
}

//기본상품명일괄적용
function autoShopItemName(upfrm,tp){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
				    if (tp==0){
					    frm.lcitemname.value = frm.ShopItemName.value;
    				}else if (tp==1){
    				    if (frm.multiLang_itemname.value.length>0){
        				    frm.lcitemname.value = frm.multiLang_itemname.value;
        				}
    				}else if (tp==2){

    				    if ((frm.multiLang_itemname.value.length>0)&&(isMayEng(frm.multiLang_itemname.value))){
    				        frm.lcitemname.value = frm.multiLang_itemname.value;
    				    }else if (isMayEng(frm.ShopItemName.value)){
    				        frm.lcitemname.value = frm.ShopItemName.value;
    				    }

    				}
				}
			}
		}
}

//기본옵션명일괄적용
function autoshopitemoptionname(upfrm,tp){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
				    if (tp==0){
					    frm.lcitemoptionname.value = frm.shopitemoptionname.value;
					}else if (tp==1){
					    if (frm.multiLang_optionname.value.length>0){
    				        frm.lcitemoptionname.value = frm.multiLang_optionname.value;
    				    }
    				}else if (tp==2){

    				    if ((frm.multiLang_optionname.value.length>0)&&(isMayEng(frm.multiLang_optionname.value))){
    				        frm.lcitemoptionname.value = frm.multiLang_optionname.value;
    				    }else if (isMayEng(frm.shopitemoptionname.value)){
    				        frm.lcitemoptionname.value = frm.shopitemoptionname.value;
    				    }

    				}
				}
			}
		}
}

function ModiArr(upfrm){
    if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm1;
	var lina = '';
	var liona = '';

	upfrm.eratea.value = '';
	upfrm.mratea.value = '';
	upfrm.ia.value = '';
	upfrm.ioa.value = '';
	upfrm.iga.value = '';
	upfrm.lina.value = '';
	upfrm.liona.value = '';
	upfrm.lpa.value = '';
		for (var i=0;i<document.forms.length;i++){
			frm1 = document.forms[i];
			if (frm1.name.substr(0,9)=="frmBuyPrc") {
				if (frm1.cksel.checked){
/*
					if (frm1.lcitemname.value == ''){
						alert('상품명을 입력해주세요');
						frm1.lcitemname.focus();
						return;
					}
*/
					if (frm1.lcprice.value == ''){
						alert('판매가를 입력해주세요');
						frm1.lcprice.focus();
						return;
					}
					upfrm.eratea.value = upfrm.eratea.value + frm1.erate.value + "," ;
					upfrm.mratea.value = upfrm.mratea.value + frm1.mrate.value + "," ;
					upfrm.ia.value = upfrm.ia.value + frm1.itemid.value + "," ;
					upfrm.ioa.value = upfrm.ioa.value + frm1.itemoption.value + "," ;
					upfrm.iga.value = upfrm.iga.value + frm1.itemgubun.value + "," ;

					lina = ''; //frm1.lcitemname.value;
					upfrm.lina.value = upfrm.lina.value + lina.replace(",","") + "," ;
					lina = '';

					liona = ''; //frm1.lcitemoptionname.value
					upfrm.liona.value = upfrm.liona.value + liona.replace(",","") + "," ;
					liona = '';
					upfrm.lpa.value = upfrm.lpa.value + frm1.lcprice.value + "," ;
				}
			}
		}

		upfrm.mode.value = 'litemadd';
		upfrm.method="post";
		upfrm.action = 'localeitem_process.asp';
		upfrm.submit();
}

function reg(page){
    var frm = document.frm;
	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('상품코드는 숫자만 가능합니다.');
			frm.itemid.focus();
			return;
		}
	}

	frm.page.value=page;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="currencyUnit_Pos" value="<%= currencyUnit_Pos %>">

<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				매장 : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
			<% end if %>
		<% else %>
			매장 : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
		<% end if %>
	    상품사용구분:<% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>

	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:reg('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		브랜드 : <% drawSelectBoxDesignerwithName "designer",designer %>
		&nbsp;
		상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" onKeyPress="if (event.keyCode == 13) reg('');">
		&nbsp;
		상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
		&nbsp;
		샵별상품명 : <input type="text" class="text" name="shopitemname" value="<%= shopitemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		물류코드 :
		<input type="text" class="text" name="prdcode" value="<%= prdcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
		&nbsp;
		범용바코드 :
		<input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		설정구분 : <% drawlocaleitemgubun "gubun" , gubun , "" %>
		&nbsp;
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >이미지보기
		&nbsp;
		<input type="checkbox" name="nameeng" value="on" <% if nameeng="on" then response.write "checked" %> >영문(상품명,옵션명)만 보기
		&nbsp;
		표시갯수 :
		<select class="select" name="pagesize">
			<option value="100">100</option>
			<option value="250" <%= CHKIIF(CLng(pagesize) = 250, "selected", "") %> >250</option>
			<option value="500" <%= CHKIIF(CLng(pagesize) = 500, "selected", "") %> >500</option>
		</select>
	</td>
</tr>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<% if adminok then %>
<tr>
    <td height="30">
    ※ 매장별 화폐단위와 배수는 <input type="button" value="해외 가격 배수 관리" class="button" onClick="popForeignPriceBase('<%= shopid %>');"> 에서 수정하세요.
    </td>
</tr>
<tr>
	<td align="left">
	    <% if currencyUnit_Pos <> "" and multipleRate <> "" then %>
	    	판매가 X 환율<input type="text" name="exchangeRate" value="<%= exchangeRate %>" size=5 maxlength=6>
	    	X 배수<input type="text" name="multipleRate" value="<%= multipleRate %>" size=3 maxlength=4>


			&nbsp;&nbsp;
			(
			소수점<input type="text" class="text" name="decimalPointLen" value="<%= decimalPointLen %>" size=1 maxlength=2>자리표시
	    	소수점<input type="text" class="text" name="decimalPointCut" value="<%= decimalPointCut %>" size=1 maxlength=2>반올림
	    	)
			<!--<input type="button" class="button" value="기본판매가적용" onclick="autoShopItemprice(frm)">
			<input type="button" class="button" value="환율적용" onclick="autoexchangeRate(frm)">
			<input type="button" class="button" value="배수적용(X<%= multipleRate %>)" onclick="automultipleRate(frm)">-->
		<% end if %>
	</td>
	<td align="right">
		    <input type="button" class="button" value="해외 판매가 계산" onclick="automulti(frm)">
			&nbsp;
		<% if (FALSE) then %>
			<input type="button" class="button" value="기본상품명/옵션명 적용" onclick="autoShopItemNameNOptionName(frm,0)">
			<% if (isShowMultiLang) then %>
			&nbsp;<input type="button" class="button" value="<%=countrylangcd%> 상품명/옵션명 적용" onclick="autoShopItemNameNOptionName(frm,1)">
			&nbsp;<input type="button" class="button" value="영문 우선  상품명/옵션명 적용" onclick="autoShopItemNameNOptionName(frm,2)">
		    <% end if %>
		<% end if %>

			<!--
			<input type="button" class="button" value="기본상품명" onclick="autoShopItemName(frm)">
			<input type="button" class="button" value="기본옵션명" onclick="autoshopitemoptionname(frm)">
			-->

			&nbsp;<input type="button" class="button" value="선택일괄저장" onclick="ModiArr(actfrm)">
	</td>
</tr>
<% end if %>
</form>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= ioffitem.FTotalcount %></b>
		&nbsp;
		<% if ioffitem.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>
		<b><%= page %> / <%= ioffitem.FTotalpage %></b>
	</td>
</tr>
<% if ioffitem.FresultCount > 0 then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<% if (imageview<>"") then %>
	<td>이미지</td>
	<% end if %>
	<td>설정<br>구분</td>
	<td>브랜드ID<br>범용바코드</td>
	<td>물류코드<br>옵션추가금액</td>
	<td>상품명</font><br>샵별상품명</td>
	<td>옵션명</font><br>샵별옵션명</td>
	<td>소비자가(원)<br>판매가(원)</td>
	<td>해외판매가<br>(<%= currencyUnit_Pos %>)</td>
	<td>환율</td>
	<td>배수</td>
	<!-- <td>해외매장<br>판매가(<%= currencyUnit_Pos %>)</td> -->

</tr>

<% for i=0 to ioffitem.FresultCount -1 %>
<form method="get" action="" name="frmBuyPrc<%=i%>">

<% if ioffitem.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE" align="center">
<% else %>
<tr bgcolor="#FFFFFF" align="center">
<% end if %>
	<td >
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		<input type="hidden" name="shopid" value="<%=shopid%>">
		<input type="hidden" name="itemid" value="<%=ioffitem.FItemlist(i).FShopitemid%>">
		<input type="hidden" name="itemoption" value="<%=ioffitem.FItemlist(i).Fitemoption%>">
		<input type="hidden" name="itemgubun" value="<%=ioffitem.FItemlist(i).fitemgubun%>">
	</td>
	<% if (imageview<>"") then %>
	<td><img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0></td>
	<% end if %>
	<td>
		<%= ioffitem.FItemlist(i).fstatus %>
	</td>
	<td>
		<%= ioffitem.FItemlist(i).FMakerID %>
		<br><%= ioffitem.FItemlist(i).FextBarcode %>
	</td>
	<td>
		<%= ioffitem.FItemlist(i).Fitemgubun %><%=  FormatCode(ioffitem.FItemlist(i).Fshopitemid) %><%= ioffitem.FItemlist(i).Fitemoption %>
		<br>
		<% if ioffitem.FItemlist(i).FOnlineOptaddprice<>0 then %>
		    <%= FormatNumber(ioffitem.FItemlist(i).FOnlineOptaddprice,0) %>
		<% end if %>
	</td>
	<td>
		<%= ioffitem.FItemlist(i).FShopItemName %>
		<% if (isShowMultiLang) then %>
		<p><%= ioffitem.FItemlist(i).FmultiLang_itemname %></p>
	    <% end if %>
		<input type="hidden" name="ShopItemName" value="<%= ioffitem.FItemlist(i).FShopItemName %>">
		<input type="hidden" name="multiLang_itemname" value="<%= ioffitem.FItemlist(i).FmultiLang_itemname %>">
		<% if (FALSE) then %>
		<br><input type="text" name="lcitemname" value="<%= ioffitem.FItemlist(i).flcitemname %>" maxlength=123 size=30 readonly style="background-color:'#EEEEEE'">
		<% end if %>
	</td>
	<td>
	    <%= ioffitem.FItemlist(i).FShopitemOptionname %>
	    <% if (isShowMultiLang) then %>
		<p><%= ioffitem.FItemlist(i).FmultiLang_optionname %></p>
	    <% end if %>

		<input type="hidden" name="shopitemoptionname" value="<%= ioffitem.FItemlist(i).fshopitemoptionname %>">
		<input type="hidden" name="multiLang_optionname" value="<%= ioffitem.FItemlist(i).FmultiLang_optionname %>">
		<% if (FALSE) then %>
		<br><input type="text" name="lcitemoptionname" value="<%= ioffitem.FItemlist(i).flcitemoptionname %>" maxlength=95 size=15 readonly style="background-color:'#EEEEEE'">
	    <% end if %>
	</td>
    <% PriceDiffExists = false %>
    <td>
        <% if (FALSE) then %>
        <%= FormatNumber(ioffitem.FItemlist(i).FShopItemOrgprice,0) %>
		<br>
	    <% end if %>
	    <%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice,0) %>
	    <input type="hidden" name="ShopItemprice" value="<%=ioffitem.FItemlist(i).FShopItemprice%>">
    </td>
    <td>
		<input type="text" name="lcprice" value="<%= CHKIIF(IsNULL(ioffitem.FItemlist(i).flcprice),"",NULL2Zero(ioffitem.FItemlist(i).flcprice)) %>" size=5 maxlength=10 onKeyup="CheckThislcprice(frmBuyPrc<%= i %>)">
    </td>
	<td>환율<input type="text" name="erate" value="<%= ioffitem.FItemlist(i).fexchangeRate %>" size=5 maxlength=5 readonly></td>
	<td>X 배수<input type="text" name="mrate" value="<%= ioffitem.FItemlist(i).fmultipleRate %>" size=5 maxlength=4 onKeyup="CheckThismrate(frmBuyPrc<%= i %>)"></td>


</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if ioffitem.HasPreScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=ioffitem.StartScrollPage-1%>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ioffitem.StartScrollPage to ioffitem.StartScrollPage + ioffitem.FScrollCount - 1 %>
			<% if (i > ioffitem.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ioffitem.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ioffitem.HasNextScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=i%>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>

<% else %>

<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</table>
<form name="actfrm" method="post">
<input type="hidden" name="shopid" value="<%=shopid%>">
<input type="hidden" name="usingyn" value="<%=usingyn%>">
<input type="hidden" name="designer" value="<%=designer%>">
<input type="hidden" name="itemid" value="<%=itemid%>">
<input type="hidden" name="itemname" value="<%=itemname%>">
<input type="hidden" name="shopitemname" value="<%=shopitemname%>">
<input type="hidden" name="prdcode" value="<%=prdcode%>">
<input type="hidden" name="generalbarcode" value="<%=generalbarcode%>">
<input type="hidden" name="gubun" value="<%=gubun%>">
<input type="hidden" name="imageview" value="<%=imageview%>">
<input type="hidden" name="nameeng" value="<%=nameeng%>">
<input type="hidden" name="currencyUnit_Pos" value="<%= currencyUnit_Pos %>">
<input type="hidden" name="ia">
<input type="hidden" name="ioa">
<input type="hidden" name="iga">
<input type="hidden" name="lina">
<input type="hidden" name="liona">
<input type="hidden" name="lpa">
<input type="hidden" name="eratea">
<input type="hidden" name="mratea">
<input type="hidden" name="mode">
</form>
<%
	set ioffitem = nothing
	set oexchangerate = nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
