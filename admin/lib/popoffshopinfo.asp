<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  OFFSHOP 정보
' History : 2009.04.07 서동석 생성
'			2010.08.04 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbhelper.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/countryinfocls.asp" -->

<%
''가맹점 사업자 정보 변경.. partner , partner_group
dim shopid ,i ,mode ,oems ,IsForeignShop ,ochargeuser ,ogroup, menupos
dim userid ,userpass ,shopname ,shopdiv ,shopphone ,shopzipcode ,shopaddr1 ,shopaddr2 ,shopCountryCode
dim manname ,manphone ,manemail ,manhp ,currencyUnit ,exchangeRate ,decimalPointLen ,decimalPointCut, currencyUnit_POS
dim multipleRate ,pyeong ,stockbasedate ,IsUsing ,vieworder ,ismobileusing ,mobileshopname ,GetMobileShopImage50X50
dim GetMobileShopImage ,mobileshopimage ,mobileworkhour ,mobileclosedate ,mobiletel ,mobileaddr ,mobilemapimage
dim GetMobileMapImage ,mobilebysubway ,mobilebybus ,mobilelatitude ,mobilelongitude ,groupid
dim Company_name ,ceoname ,company_no ,jungsan_gubun ,company_zipcode ,company_address ,company_address2
dim company_uptae ,company_upjong ,jungsan_name ,jungsan_email ,jungsan_hp ,admindisplang, loginsite, countrylangcd
dim ctropen, ViewSort, engName, ShopFax, engAddress
	shopid = RequestCheckVar(request("shopid"),32)
	menupos = RequestCheckVar(request("menupos"),10)

if shopid <> "" then
	mode = "edit"
else
	mode = "new"
end if

set ochargeuser = new COffShopChargeUser
	ochargeuser.FRectShopID = shopid

	if shopid <> "" then
		ochargeuser.GetOffShopList
	end if

if ochargeuser.FResultCount > 0 then
	userid = ochargeuser.FItemList(0).Fuserid
	countrylangcd = ochargeuser.FItemList(0).fcountrylangcd
	userpass = ochargeuser.FItemList(0).Fuserpass
	shopname = ochargeuser.FItemList(0).Fshopname
	shopdiv = ochargeuser.FItemList(0).Fshopdiv
	shopphone = ochargeuser.FItemList(0).Fshopphone
	shopzipcode = ochargeuser.FItemList(0).Fshopzipcode
	shopaddr1 = ochargeuser.FItemList(0).Fshopaddr1
	shopaddr2 = ochargeuser.FItemList(0).Fshopaddr2
	shopCountryCode = ochargeuser.FItemList(0).FshopCountryCode
	manname = ochargeuser.FItemList(0).Fmanname
	manphone = ochargeuser.FItemList(0).Fmanphone
	manemail = ochargeuser.FItemList(0).Fmanemail
	manhp = ochargeuser.FItemList(0).Fmanhp
	currencyUnit = ochargeuser.FItemList(0).fcurrencyUnit
	currencyUnit_POS = ochargeuser.FItemList(0).fcurrencyUnit_POS
	exchangeRate = ochargeuser.FItemList(0).FexchangeRate
	decimalPointLen = ochargeuser.FItemList(0).FdecimalPointLen
	decimalPointCut = ochargeuser.FItemList(0).FdecimalPointCut
	multipleRate = ochargeuser.FItemList(0).fmultipleRate
	pyeong = ochargeuser.FItemList(0).fpyeong
	stockbasedate = ochargeuser.FItemList(0).Fstockbasedate
	IsUsing = ochargeuser.FItemList(0).FIsUsing
	vieworder = ochargeuser.FItemList(0).Fvieworder
	ismobileusing = ochargeuser.FItemList(0).Fismobileusing
	mobileshopname = ochargeuser.FItemList(0).Fmobileshopname
	GetMobileShopImage50X50 = ochargeuser.FItemList(0).GetMobileShopImage50X50
	GetMobileShopImage = ochargeuser.FItemList(0).GetMobileShopImage
	mobileshopimage = ochargeuser.FItemList(0).Fmobileshopimage
	mobileworkhour = ochargeuser.FItemList(0).Fmobileworkhour
	mobileclosedate = ochargeuser.FItemList(0).Fmobileclosedate
	mobiletel = ochargeuser.FItemList(0).Fmobiletel
	mobileaddr = ochargeuser.FItemList(0).Fmobileaddr
	mobilemapimage = ochargeuser.FItemList(0).Fmobilemapimage
	GetMobileMapImage = ochargeuser.FItemList(0).GetMobileMapImage
	mobilebysubway = ochargeuser.FItemList(0).Fmobilebysubway
	mobilebybus = ochargeuser.FItemList(0).Fmobilebybus
	mobilelatitude = ochargeuser.FItemList(0).Fmobilelatitude
	mobilelongitude = ochargeuser.FItemList(0).Fmobilelongitude
	groupid = ochargeuser.FItemList(0).Fgroupid
	admindisplang = ochargeuser.FItemList(0).Fadmindisplang
	loginsite = ochargeuser.FItemList(0).floginsite

	IsForeignShop = ochargeuser.FItemList(0).IsForeignShop
	ctropen= ochargeuser.FItemList(0).Fctropen
	ViewSort= ochargeuser.FItemList(0).FViewSort
	engName= ochargeuser.FItemList(0).FengName
	ShopFax= ochargeuser.FItemList(0).FShopFax
	engAddress= ochargeuser.FItemList(0).FengAddress
end if

set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = groupid

	if groupid <> "" then
		ogroup.GetOneGroupInfo
	end if

if ogroup.FTotalCount > 0 then
	Company_name = ogroup.FOneItem.FCompany_name
	ceoname = ogroup.FOneItem.Fceoname
	company_no = ogroup.FOneItem.Fcompany_no
	jungsan_gubun = ogroup.FOneItem.Fjungsan_gubun
	company_zipcode = ogroup.FOneItem.Fcompany_zipcode
	company_address = ogroup.FOneItem.Fcompany_address
	company_address2 = ogroup.FOneItem.Fcompany_address2
	company_uptae = ogroup.FOneItem.Fcompany_uptae
	company_upjong = ogroup.FOneItem.Fcompany_upjong
	jungsan_name = ogroup.FOneItem.Fjungsan_name
	jungsan_email = ogroup.FOneItem.Fjungsan_email
	jungsan_hp = ogroup.FOneItem.Fjungsan_hp
end if

SET oems = New CCountryInfo
    oems.FRectCurrPage = 1
    oems.FRectPageSize = 200
    oems.FRectisUsing  = ""
    oems.GetCountryInfoList

if isusing = "" then isusing = "Y"
if ctropen = "" then ctropen = "0"
if admindisplang = "" then admindisplang = "KOR"
if isnull(currencyUnit) or currencyUnit="" then currencyUnit="KRW"
if loginsite = "" then loginsite = "SCM"
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function CopyZip(flag,post1,post2,add,dong){
	frmedit.shopzipcode.value= post1 + "-" + post2;
	frmedit.shopaddr1.value= add;
	frmedit.shopaddr2.value= dong;
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function editShopInfo(frm,mode){

	<% if (Not IsForeignShop) then %>
	    if (frm.groupid.value.length<1){
			alert('사업자정보를 등록하세요.');
			frm.groupid.focus();
			return;
		}
	<% end if %>

	var errMsg = chkIsValidJungsanGubun(frm.company_no.value, frm.jungsan_gubun.value);
	if (errMsg != "OK") {
		alert(errMsg);
		return;
	}
    <% if ochargeuser.FResultCount > 0 then %>
    <% else %>
	if (frm.userpass.value.length<4){
		alert('패스워드는 4자 이상입니다.');
		frm.userpass.focus();
		return;
	}
    <% end if %>
	if (frm.shopid.value.length<1){
		alert('매장아이디를 입력하세요.');
		frm.shopid.focus();
		return;
	}

	if (frm.shopid.value.substr(0,10) != 'streetshop' &&
		frm.shopid.value.substr(0,8) != 'ithinkso' &&
		frm.shopid.value.substr(0,9) != 'wholesale' &&
		frm.shopid.value.substr(0,9) != 'ygentshop' &&
		frm.shopid.value.substr(0,8) != 'its_exp_' &&
		frm.shopid.value.substr(0,8) != '3pl_its_' &&
		frm.shopid.value.substr(0,4) != '3pl_' &&
		frm.shopid.value.substr(0,8) != 'notagb2b' &&
		frm.shopid.value.substr(0,8) != 'gsexport' &&
		frm.shopid.value.substr(0,10) != 'fotonotaes' &&
		frm.shopid.value.substr(0,11) != 'offgiftcard'
	) {
		alert('유효한 매장아이디를 입력하세요.');
		frm.shopid.focus();
		return;
	}

	if (frm.shopname.value.length<1){
		alert('매장 이름을 입력하세요.');
		frm.shopname.focus();
		return;
	}

	if (frm.shopdiv.value.length<1){
		alert('매장 구분을 입력하세요.');
		frm.shopdiv.focus();
		return;
	}

	if (frm.shopCountryCode.value == ''){
		alert('국가선택을 해주세요');
		frm.shopCountryCode.focus();
		return;
	}

	if (frm.shopCountryCode.value == 'KR'){
	    if (frm.shopzipcode.value==''){
	        alert('국가 대한민국을 선택하셨습니다. 우편번호를 입력하세요.');
			return;
	    }
	}

	if (frm.currencyUnit.value == ''){
		alert('대표화폐를 선택해주세요');
		frm.currencyUnit.focus();
		return;
	}

	if (frm.loginsite.value == ''){
		alert('로그인사이트를 선택해주세요');
		frm.loginsite.focus();
		return;
	}
	if (frm.countrylangcd.value == ''){
		alert('대표언어를 선택해주세요');
		return;
	}

 	if (frm.ismobileusing[0].checked == true) {
 		// 모바일 표시정보

	    if (frm.mobileshopname.value.length<1){
			alert('모바일샵명을 입력하세요.');
			frm.mobileshopname.focus();
			return;
		}

	    if (frm.mobileworkhour.value.length<1){
			alert('영업시간을 입력하세요.');
			frm.mobileworkhour.focus();
			return;
		}

	    if (frm.mobileclosedate.value.length<1){
			alert('휴점일을 입력하세요.');
			frm.mobileclosedate.focus();
			return;
		}

	    if (frm.mobiletel.value.length<1){
			alert('대표전화를 입력하세요.');
			frm.mobiletel.focus();
			return;
		}

	    if (frm.mobileaddr.value.length<1){
			alert('모바일주소를 입력하세요.');
			frm.mobileaddr.focus();
			return;
		}
 	}

    if (frm.mobilelatitude.value.length<1){
		frm.mobilelatitude.value.length = 0.0;
	} else {
		if (frm.mobilelatitude.value.length*0 != 0) {
			alert('위도는 숫자만 입력가능합니다.');
			frm.mobilelatitude.focus();
			return;
		}
	}

    if (frm.mobilelongitude.value.length<1){
		frm.mobilelongitude.value.length = 0.0;
	} else {
		if (frm.mobilelongitude.value.length*0 != 0) {
			alert('위도는 숫자만 입력가능합니다.');
			frm.mobilelongitude.focus();
			return;
		}
	}

	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value=mode;
		frm.action="/admin/lib/popoffshopinfo_process.asp";
		frm.submit();
	}
}

function emsBoxChange(obj) {
	var shopCountryCode = obj.value;

	if (shopCountryCode == "") {
		return;
	}

	if (shopCountryCode == "KR") {
		frmedit.btnsearchzipcode.disabled = false;
		frmedit.shopzipcode.readOnly = true;
		frmedit.shopaddr1.readOnly = true;
		return;
	} else {
		frmedit.btnsearchzipcode.disabled = true;
		frmedit.shopzipcode.readOnly = false;
		frmedit.shopaddr1.readOnly = false;

		frmedit.shopzipcode.value= '';
		return;
	}
}

function clearZipcode() {
	frmedit.shopzipcode.value = "";
	frmedit.shopaddr1.value = "";
}

function popUploadShopimage(frm) {
	var mode, imagekind, pk;

	if (frm.mobileshopimage.value == "") {
		mode = "addimage";
	} else {
		mode = "editimage";
	}

	imagekind = "mobileshopimage";
	pk = frm.shopid.value;


	var popwin = window.open("/common/pop_upload_image.asp?mode=" + mode + "&imagekind=" + imagekind + "&pk=" + pk + "&50X50=Y","popUploadShopimage","width=390 height=120 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popUploadShopmap(frm) {
	var mode, imagekind, pk;

	if (frm.mobilemapimage.value == "") {
		mode = "addimage";
	} else {
		mode = "editimage";
	}

	imagekind = "mobilemapimage";
	pk = frm.shopid.value;


	var popwin = window.open("/common/pop_upload_image.asp?mode=" + mode + "&imagekind=" + imagekind + "&pk=" + pk,"popUploadShopmap","width=390 height=120 scrollbars=yes resizable=yes");
	popwin.focus();
}

function chkIsValidJungsanGubun(company_no, jungsan_gubun) {
	// 000-00-00000
	// 가운데 두글자 : 구분코드
	// =========================================================================
	// 01-79 : 개인사업자+과세사업자
	// 90-99 : 개인사업자+면세사업자
	// 기타 : 과세 면세 모두 가능
	// 앞자리 888 = 영세(해외)
	// =========================================================================

	if (company_no.length != 12) {
		return "잘못된 사업자번호입니다.";
	}

	var soc_gubun = company_no.substring(4, 6)*1;
	var IsForeign = (company_no.substring(0, 3) == "888");

	if (IsForeign) {
		if (jungsan_gubun != "영세(해외)") {
			return "영세(해외) 사업자만 가능한 사업자번호입니다.";
		}

		return "OK";
	} else {
		if (jungsan_gubun == "영세(해외)") {
			return "영세(해외) 사업자로 변경 불가능한 사업자번호입니다.";
		}

		/*
		if ((soc_gubun >= 1) && (soc_gubun <= 79)) {
			if (jungsan_gubun == "면세") {
				return "면세로 등록할 수 없는 사업자번호입니다.";
			}

			return "OK";
		}
		*/

		if ((soc_gubun >= 90) && (soc_gubun <= 99)) {
			if (jungsan_gubun != "면세") {
				return "면세로만 등록가능한 사업자번호입니다.";
			}

			return "OK";
		}

		return "OK";
	}
}

function chcountrylangcd(loginsite, countrylangcd){
	var str = $.ajax({
		type: "GET",
		url: "/common/offshop/exchangeRate/ajax_countrylangcd.asp",
		data: "loginsite="+loginsite+"&countrylangcd="+countrylangcd,
		dataType: "text",
		async: false
	}).responseText;

	$('#divcountrylangcd').empty().html(str);
	frmedit.countrylangcd.value="";		// 초기화
}

function selectedcountrylangcd(countrylangcd){
	frmedit.countrylangcd.value=countrylangcd;
}

</script>

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
<form name="frmedit" method="post" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="countrylangcd" value="<%= countrylangcd %>">
<input type="hidden" name="mode" value="">
<tr  bgcolor="ffffff">
	<td colspan="4"><b>1.사업자정보</b></td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">*업체코드</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="groupid" value="<%= groupid %>" size="7" maxlength="5" readonly>

		<% if GroupId<>"" then %>
			<input type="button" class="button" value="사업자정보수정" onclick="PopUpcheInfoEdit('<%= groupid %>')">
		<% else %>
			<input type="button" class="button" value="업체선택" onClick="PopUpcheSelect_shop('frmedit','Y');">
		<% end if %>

		<% if C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH then %>
			(관리자모드 : <input type="button" class="button" value="업체신규등록" onClick="PopUpcheInfoEdit('');">)
		<% end if %>
	</td>
</tr>
<tr >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" width="120">회사명(상호)</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="company_name" id="company_name" value="<%= Company_name %>" size="30" maxlength="32" readonly>
	</td>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" width="120">대표자</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="ceoname" value="<%= ceoname %>" size="16" maxlength="16" readonly>
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">사업자번호</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="company_no" value="<%= company_no %>" size="16" maxlength="20" readonly>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">과세구분</td>
	<td bgcolor="#FFFFFF">

		<input type="text" class="text_ro" name="jungsan_gubun" value="<%= jungsan_gubun %>" size="16" maxlength="20" readonly>
		<!--<select name="jungsan_gubun" class="select">
			<option value="일반과세" <%' if jungsan_gubun = "일반과세" then response.write " selected" %>>일반과세</option>
			<option value="간이과세" <%' if jungsan_gubun = "간이과세" then response.write " selected" %>>간이과세</option>
			<option value="원천징수" <%' if jungsan_gubun = "원천징수" then response.write " selected" %>>원천징수</option>
			<option value="면세" <%' if jungsan_gubun = "면세" then response.write " selected" %>>면세</option>
			<option value="영세(해외)" <%' if jungsan_gubun = "영세(해외)" then response.write " selected" %>>영세(해외)</option>
		</select>-->
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">사업장소재지</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="company_zipcode" value="<%= company_zipcode %>" size="7" maxlength="7" readonly>
		<input type="text" class="text_ro" name="company_address" value="<%= company_address %>" size="30" maxlength="64" readonly>
		<input type="text" class="text_ro" name="company_address2" value="<%= company_address2 %>" size="46" maxlength="64" readonly>
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">업태</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="company_uptae" value="<%= company_uptae %>" size="30" maxlength="32" readonly>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">업종</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="company_upjong" value="<%= company_upjong %>" size="30" maxlength="32" readonly>
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">정산담당자</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="jungsan_name" value="<%= jungsan_name %>" size="30" maxlength="32" readonly>
		<input type="text" class="text_ro" name="jungsan_email" value="<%= jungsan_email %>" size="30" maxlength="64" readonly>
		<input type="text" class="text_ro" name="jungsan_hp" value="<%= jungsan_hp %>" size="16" maxlength="16" readonly>
	</td>
</tr>
<tr  bgcolor="ffffff">
	<td colspan="4"><b>2.Shop정보</b></td>
</tr>
<tr >
	<td width="90" bgcolor="<%= adminColor("tabletop") %>">*ShopID</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="shopid" value="<%= userid %>" size="16" maxlength="16" <% if mode <> "new" then response.write " readonly class='text_ro'" %>>
		<br>텐바이텐 매장 - streetshopxxx
		<br>아이띵소 매장 - ithinksoxxxxx, 3pl_its_xxxxx
		<br>도매 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;- wholesale1xxx
		<br>대행 매장 &nbsp; &nbsp; &nbsp; - ygentshop1xxx, 3pl_xxx_xxxxx
		<br>아이띵소 해외출고처 &nbsp; &nbsp; &nbsp; - its_exp_xxxxx
	</td>
	<td width="90" bgcolor="<%= adminColor("tabletop") %>">*Password</td>
	<td bgcolor="#FFFFFF">
	    <% if ochargeuser.FResultCount > 0 then %>
        <input type="button" value="수정" onClick="alert('관리자문의요망');">
	    <% else %>
		<input type="password" class="text" name="userpass" value="" size="16" maxlength="16">
		<br>(포스로그인시사용)
		<!-- 패스워드 팝업창으로 변경 -->
		<% end if %>
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">*Shop명</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="shopname" value="<%= shopname %>" size="20" maxlength="64">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">*Shop구분</td>
	<td bgcolor="#FFFFFF">
	    <% drawoffshop_commoncode "shopdiv", shopdiv, "shopdiv", "", "", "" %>
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">Shop전화번호</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="shopphone" value="<%= shopphone %>" size="16" maxlength="16"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">Shop Fax번호</td>
	<td bgcolor="#FFFFFF">
	    <input type="text" class="text" name="shopfax" value="<%= ShopFax %>" size="16" maxlength="16">
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">*Shop주소</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" name="shopzipcode" class="text" value="<%= shopzipcode %>" size="7" maxlength="7" <%= CHKIIF(IsForeignShop,"","ReadOnly") %> >
		<input type="button" class="button" value="검색" onClick="FnFindZipNew('frmedit','F')">
		<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frmedit','F')">
		<% '<input type="button" class="button" name="btnsearchzipcode" value="검색(구)" onclick="javascript:popZip('s');"> %>
		<input type="button" class="button" value="삭제" onclick="javascript:clearZipcode();">

		<select name="shopCountryCode" class="select" style="width:200px;height:20px;" onChange="emsBoxChange(this);">
			<option value="">국가선택</option>
			<option value="KR" <% if (shopCountryCode = "KR") then %>selected<% end if %>>대한민국</option>

			<% for i=0 to oems.FREsultCount-1 %>
				<option value="<%= oems.FItemList(i).FcountryCode %>" <% if (shopCountryCode = oems.FItemList(i).FcountryCode) then %>selected<% end if %>><%= oems.FItemList(i).FcountryNameKr %>(<%= oems.FItemList(i).FcountryNameEn %>)</option>
			<% next %>
		</select> ( <b>*</b> 매장이 위치한 국가)
		<br>
		<input type="text" class="text" name="shopaddr1" value="<%= shopaddr1 %>" size="60" maxlength="64">
		<input type="text" class="text" name="shopaddr2" value="<%= shopaddr2 %>" size="60" maxlength="64">

	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">매니저</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="manname" value="<%= manname %>" size="16" maxlength="32">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">매니저Phone</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="manphone" value="<%= manphone %>" size="16" maxlength="16">
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">매니저Email</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="manemail" value="<%= manemail %>" size="25" maxlength="128">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">매니저HP</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="manhp" value="<%= manhp %>" size="16" maxlength="16">
	</td>
</tr>
<tr  bgcolor="ffffff">
	<td colspan="4"><b>3.글로벌정보</b></td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td bgcolor="<%= adminColor("tabletop") %>">로그인사이트</td>
	<td bgcolor="#FFFFFF">
		<% drawoffshop_commoncode "loginsite", loginsite, "loginsite", "MAIN", "", " onchange=""chcountrylangcd(this.value,'"& countrylangcd &"');""" %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">대표화폐</td>
	<td bgcolor="#FFFFFF">
		<% DrawexchangeRate "currencyUnit",currencyUnit,"" %>

	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td bgcolor="<%= adminColor("tabletop") %>">화폐(LOCALE)<br>환율<br>해외마진배수 </td>
	<td bgcolor="#FFFFFF">
	    <% DrawexchangeRate "currencyUnit_POS",currencyUnit_POS,"" %><br>
	    <input type="text" class="text" name="exchangeRate" value="<%= exchangeRate %>" size=12 maxlength=12><br>
		<input type='text' name='multipleRate' value='<%=multipleRate%>' size=10 maxlength=10>
		ex) 판매가 x 마진배수(1.0) = 실제판매가
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">대표언어</td>
	<td bgcolor="#FFFFFF" id="divcountrylangcd">
		<% DrawexchangeRate_countrylangcd "tmpcountrylangcd",countrylangcd, loginsite, " onchange='selectedcountrylangcd(this.value);'" %>
		<br>※로그인사이트에 따라 해당되는 언어만 표기 됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td bgcolor="<%= adminColor("tabletop") %>">어드민표기언어</td>
	<td bgcolor="#FFFFFF">
		<% drawoffshop_commoncode "admindisplang", admindisplang, "admindisplang", "MAIN", "", "" %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">영문명</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="engName" value="<%= engName %>" size=20 maxlength=32>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td bgcolor="<%= adminColor("tabletop") %>">영문주소</td>
	<td colspan="3"  bgcolor="#FFFFFF">
		<input type="text" class="text" name="engAddress" value="<%= engAddress %>" size=70 maxlength=128>
	</td>
</tr>
<tr  bgcolor="ffffff">
	<td colspan="4"><b>4.기타정보</b></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">화면표시순서</td>
	<td bgcolor="#FFFFFF">
		프론트 노출 여부 : <input type="text" class="text" name="vieworder" value="<%= vieworder %>" size="2">	(0 일경우 화면표시안함.)<br>
		상품 상세 판매매장 정렬 순서 : <input type="text" class="text" name="viewsort" value="<%= ViewSort %>" size="2">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">화폐소수점</td>
	<td bgcolor="#FFFFFF">
	    표시 <input type="text" class="text" name="decimalPointLen" value="<%= decimalPointLen %>" size=2 maxlength=2> 자리
	    절삭 <input type="text" class="text" name="decimalPointCut" value="<%= decimalPointCut %>" size=2 maxlength=2> 자리
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td bgcolor="<%= adminColor("tabletop") %>">사용구분(매장종료여부)</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <% if IsUsing="Y" then response.write "checked" %> >사용함
		<input type="radio" name="isusing" value="N" <% if IsUsing="N" then response.write "checked" %> >사용안함
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">매장총평수</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="pyeong" value="<%= pyeong %>" size=5 maxlength=5>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">재고기준일</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="stockbasedate" value="<%= stockbasedate %>" size=10 maxlength=10> ex) 2012-01-01
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">업체계약정보</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="ctropen" value="1" <% if ctropen="1" then response.write "checked" %> >사용함
		<input type="radio" name="ctropen" value="0" <% if ctropen="0" then response.write "checked" %> >사용안함
	</td>
</tr>
<tr  bgcolor="ffffff">
	<td colspan="4"><b>5.모바일표시정보</b></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">모바일표시여부</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="ismobileusing" value="Y" <% if ismobileusing="Y" then response.write "checked" %> >표시함
		<input type="radio" name="ismobileusing" value="N" <% if ismobileusing<>"Y" then response.write "checked" %> >표시안함
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">모바일샵명</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="mobileshopname" value="<%= mobileshopname %>" size=32 maxlength=32>
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">모바일주소</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="mobileaddr" value="<%= mobileaddr %>" size=50 maxlength=50>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">대표전화</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="mobiletel" value="<%= mobiletel %>" size=16 maxlength=16>
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">영업시간</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="mobileworkhour" value="<%= mobileworkhour %>" size=50 maxlength=100>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">휴점일</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="mobileclosedate" value="<%= mobileclosedate %>" size=50 maxlength=100>
	</td>
</tr>

<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">샵이미지(400X400)</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% if (mobileshopimage <> "") then %>
			<img src="<%= GetMobileShopImage50X50 %>"><br>
			<img src="<%= GetMobileShopImage %>"><br>
			<input type="button" class="button" value="수정하기" onclick="popUploadShopimage(frmedit)">
		<% else %>
			<input type="button" class="button" value="등록하기" onclick="popUploadShopimage(frmedit)">
		<% end if %>
		<input type="hidden" name="mobileshopimage" value="<%= mobileshopimage %>">
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">약도(400X400)</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% if (mobilemapimage <> "") then %>
			<img src="<%= GetMobileMapImage %>"><br>
			<input type="button" class="button" value="수정하기" onclick="popUploadShopmap(frmedit)">
		<% else %>
			<input type="button" class="button" value="등록하기" onclick="popUploadShopmap(frmedit)">
		<% end if %>
		<input type="hidden" name="mobilemapimage" value="<%= mobilemapimage %>">
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">대중교통지하철</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<textarea class="textarea" cols="80" rows="3" name="mobilebysubway">
			<%= mobilebysubway %>
		</textarea>
	</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("tabletop") %>">대중교통버스</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<textarea class="textarea" cols="80" rows="3" name="mobilebybus">
			<%= mobilebybus %>
		</textarea>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">위도</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="mobilelatitude" value="<%= mobilelatitude %>" size=16 maxlength=16>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">경도</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="mobilelongitude" value="<%= mobilelongitude %>" size=16 maxlength=16>
	</td>
</tr>
<tr  bgcolor="#FFFFFF">
	<td colspan="4" align="center"><input type="button" class="button" value="정보수정" onclick="editShopInfo(frmedit,'<%=mode%>')"></td>
</tr>
</form>
</table>

<script type="text/javascript">
	<% if countrylangcd<>"" then %>
		selectedcountrylangcd('<%= countrylangcd %>');
	<% end if %>
	emsBoxChange(frmedit.shopCountryCode);
</script>

<%
set oems = Nothing
set ochargeuser = Nothing
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
