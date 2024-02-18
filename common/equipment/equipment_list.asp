<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 장비자산관리
' History : 2008년 06월 27일 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/common/equipment/equipment_cls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
dim page, equip_gubun, part_sn, idx, using_userid, using_username, usingIp , equip_code ,equip_name ,manufacture_company, manufacture_sn
dim ipgocheck, yyyy1, yyyy2, mm1, mm2, dd1, dd2, fromDate, toDate, useip ,property_gubun ,research, BIZSECTION_CD, BIZSECTION_NM
dim totalcurrsum,totaljasan, Alltotaljasan, part_code ,state , parameter, only1000, i, dispmainimageyn, accountassetcode, sorttype
dim onlyusing, accountGubun, department_id, outcheck, yyyy3, yyyy4, mm3, mm4, dd3, dd4, fromDate2, toDate2, paymentrequestidx, buyCompanyName
	accountassetcode = requestcheckvar(request("accountassetcode"),32)
	paymentrequestidx = requestcheckvar(request("paymentrequestidx"),10)
	page = requestcheckvar(request("page"),10)
	equip_gubun = requestcheckvar(Request("equip_gubun"),2)
	part_sn = requestcheckvar(Request("part_sn"),10)
	using_userid = requestcheckvar(Request("using_userid"),32)
	using_username = requestcheckvar(Request("using_username"),32)
	equip_code = requestcheckvar(request("equip_code"),20)
	ipgocheck = requestcheckvar(request("ipgocheck"),2)
	yyyy1 = requestcheckvar(request("yyyy1"),4)
	yyyy2 = requestcheckvar(request("yyyy2"),4)
	mm1	  = requestcheckvar(request("mm1"),2)
	mm2	  = requestcheckvar(request("mm2"),2)
	dd1	  = requestcheckvar(request("dd1"),2)
	dd2	  = requestcheckvar(request("dd2"),2)
	part_code = requestcheckvar(Request("part_code"),10)
	equip_name = requestcheckvar(Request("equip_name"),64)
	manufacture_company = requestcheckvar(Request("manufacture_company"),64)
	buyCompanyName = requestcheckvar(Request("buyCompanyName"),64)
	manufacture_sn = requestcheckvar(Request("manufacture_sn"),64)
	property_gubun = requestcheckvar(Request("property_gubun"),10)
	state = requestcheckvar(Request("state"),10)
	research = requestcheckvar(Request("research"),2)
	onlyusing = requestcheckvar(Request("onlyusing"),2)
	accountGubun = requestcheckvar(Request("accountGubun"),5)
	department_id = requestcheckvar(Request("department_id"),5)
	BIZSECTION_CD = requestcheckvar(Request("BIZSECTION_CD"),15)
	BIZSECTION_NM = requestcheckvar(Request("BIZSECTION_NM"),55)
	only1000 = requestcheckvar(Request("only1000"),55)
	outcheck = requestcheckvar(request("outcheck"),2)
	yyyy3 = requestcheckvar(request("yyyy3"),4)
	yyyy4 = requestcheckvar(request("yyyy4"),4)
	mm3	  = requestcheckvar(request("mm3"),2)
	mm4	  = requestcheckvar(request("mm4"),2)
	dd3	  = requestcheckvar(request("dd3"),2)
	dd4	  = requestcheckvar(request("dd4"),2)
	dispmainimageyn = requestcheckvar(request("dispmainimageyn"),1)
	sorttype = requestcheckvar(request("sorttype"),1)

if sorttype="" then sorttype="1"
if page="" then page=1
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if (yyyy3="") then yyyy3 = Cstr(Year(now()))
if (mm3="") then mm3 = Cstr(Month(now()))
if (dd3="") then dd3 = Cstr(day(now()))
if (yyyy4="") then yyyy4 = Cstr(Year(now()))
if (mm4="") then mm4 = Cstr(Month(now()))
if (dd4="") then dd4 = Cstr(day(now()))

fromDate2 = CStr(DateSerial(yyyy3, mm3, dd3))
toDate2 = CStr(DateSerial(yyyy4, mm4, dd4+1))

if (research = "") then
	onlyusing = "Y"
end if

if research = "" and dispmainimageyn="" then
	dispmainimageyn="Y"
end if

dim oequip
set oequip = new CEquipment
	oequip.FPageSize = 50
	oequip.FCurrPage = page
	oequip.FRectequip_gubun = equip_gubun
	oequip.FRectpart_sn = part_sn
	oequip.FRectusing_userid = using_userid
	oequip.FRectusing_username = using_username
	oequip.Frectequip_code = equip_code
	oequip.frectequip_name = equip_name
	oequip.frectmanufacture_company = manufacture_company
	oequip.fRectBuyCompanyName = buyCompanyName
	oequip.frectmanufacture_sn = manufacture_sn
	oequip.frectproperty_gubun = property_gubun
	oequip.frectstate = state
	oequip.FRectIsusing = onlyusing
	oequip.FRectAccountGubun = accountGubun
	oequip.FRectDepartmentID = department_id
	oequip.FRectBIZSECTION_CD = BIZSECTION_CD
	oequip.FRectOnly1000 = only1000
	oequip.frectaccountassetcode = accountassetcode
	oequip.frectpaymentrequestidx = paymentrequestidx
	oequip.frectsorttype = sorttype

	if ipgocheck = "on" then
		oequip.frectbuy_startdate = fromDate
		oequip.frectbuy_enddate = toDate
	end if

	if outcheck = "on" then
		oequip.frectout_startdate = fromDate2
		oequip.frectout_enddate = toDate2
	end if

	oequip.getEquipmentList

totalcurrsum = 0
totaljasan	 = 0
Alltotaljasan = 0

parameter = Request.ServerVariables("QUERY_STRING")
%>

<script type="text/javascript" src="/js/ttpbarcode.js"></script>
<script type="text/javascript">

//바코드 출력 팝업
function barcode(barcode){
	var barcode = window.open('/common/barcode/barcode_image.asp?barcode='+barcode+'&image=3&barcodetype=23&height=30&barwidth=1','barcode','width=600,height=400,scrollbars=yes,resizable=yes');
	barcode.focus();
}

//라벨프린터출력 시작
function ExcelSheet(idx1){
	var ExcelSheet = window.open('/common/equipment/popexcel_equipment.asp?idx=' + idx1,'ExcelSheet','width=400,height=300,scrollbars=yes,resizable=yes');
	ExcelSheet.focus();
}

//신규등록
function pop_Equipmentreg(idx){
	var pop_Equipmentreg = window.open('/common/equipment/pop_equipmentreg.asp?idx=' + idx,'pop_Equipmentreg','width=1280,height=960,scrollbars=yes,resizable=yes');
	pop_Equipmentreg.focus();
}

//현재페이지엑셀출력
function pageexcelsheet(page,jangbi,sayoug,user,idx,code){
	var pageexcelsheet = window.open('/common/equipment/equipment_excel.asp?<%=parameter%>','pageexcelsheet','width=400,height=300,scrollbars=yes,resizable=yes');
	pageexcelsheet.focus();
}

function NextPage(page){
	document.frm.page.value= page;
	document.frm.mode.value = "";
	document.frm.arridx.value = "";
	document.frm.method="GET";
	document.frm.action = "";
	document.frm.target = "";
	document.frm.submit();
}

//구매일 체크
function EnDisabledDateBox(comp){
	document.frm.yyyy1.disabled = !comp.checked;
	document.frm.yyyy2.disabled = !comp.checked;
	document.frm.mm1.disabled = !comp.checked;
	document.frm.mm2.disabled = !comp.checked;
	document.frm.dd1.disabled = !comp.checked;
	document.frm.dd2.disabled = !comp.checked;
}

//폐기일 체크
function EnDisabledDateBox2(comp){
	document.frm.yyyy3.disabled = !comp.checked;
	document.frm.yyyy4.disabled = !comp.checked;
	document.frm.mm3.disabled = !comp.checked;
	document.frm.mm4.disabled = !comp.checked;
	document.frm.dd3.disabled = !comp.checked;
	document.frm.dd4.disabled = !comp.checked;
}

//코드관리
function popcodemanager(){
	var popcodemanager = window.open('/common/equipment/popmanagecode.asp','popcodemanager','width=800,height=768,scrollbars=yes,resizable=yes');
	popcodemanager.focus();
}

function popEtcBar(){
    var popwin = window.open('popEtcBar.asp','popcodemanager','width=800,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//자금관리부서 선택
function jsGetPart(){
	var winP = window.open('/admin/linkedERP/Biz/popGetBizOne.asp','popGetBizOne','width=600, height=500, resizable=yes, scrollbars=yes');
	winP.focus();
}

//자금관리부서 등록
function jsSetPart(selUP, sPNM){
	document.frm.BIZSECTION_CD.value = selUP;
	document.frm.BIZSECTION_NM.value = sPNM;
}

function jsClearPart() {
	document.frm.BIZSECTION_CD.value = "";
	document.frm.BIZSECTION_NM.value = "";
}

// 장비 바코드 출력
function EquipBarcodePrint() {
	var arr = new Array();

	var ttptype, papermargin, heightoffset;
	var equip_code, AccountGubunName, equip_name, buy_date;

	if (!CheckSelected()){
		alert("선택아이템이 없습니다.");
		return;
	}

	ttptype			= "TTP-243_80x50";
	papermargin		= 3;
	heightoffset	= 0;

	var frm;
	for (var i = 0;i < document.forms.length; i++) {
		frm = document.forms[i];

		if (frm.name.substr(0, 10)=="frmBuyPrc_") {
			if (frm.cksel.checked == true) {
				equip_code			= frm.equip_code.value;
				AccountGubunName	= frm.AccountGubunName.value;
				equip_name			= frm.equip_name.value;
				buy_date			= frm.buy_date.value;

				// alert(equip_code);

				var v = new TTPEquipBarcodeDataClass(equip_code, AccountGubunName, equip_name, buy_date);
				arr.push(v);
			}
		}
	}

	if (arr.length < 1) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	// /js/barcode.js 참조
	if (initTTPprinter(ttptype, "", "", "", "", "", "", papermargin, heightoffset) != true) {
		alert("프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[4]");
		return;
	}

	printTTPMultiEquipBarcode(arr);
}

// 장비 바코드 출력
function EquipSmallBarcodePrint() {
	var arr = new Array();

	var ttptype, papermargin, heightoffset;
	var equip_code, AccountGubunName, equip_name, buy_date;

	if (!CheckSelected()){
		alert("선택아이템이 없습니다.");
		return;
	}

	ttptype			= "TTP-243_45x22";
	papermargin		= 3;
	heightoffset	= 0;

	var frm;
	for (var i = 0;i < document.forms.length; i++) {
		frm = document.forms[i];

		if (frm.name.substr(0, 10)=="frmBuyPrc_") {
			if (frm.cksel.checked == true) {
				equip_code			= frm.equip_code.value;
				AccountGubunName	= frm.AccountGubunName.value;
				equip_name			= frm.equip_name.value;
				buy_date			= frm.buy_date.value;

				// alert(equip_code);

				var v = new TTPEquipBarcodeDataClass(equip_code, AccountGubunName, equip_name, buy_date);
				arr.push(v);
			}
		}
	}

	if (arr.length < 1) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	// /js/barcode.js 참조
	if (initTTPprinter(ttptype, "", "", "", "", "", "", papermargin, heightoffset) != true) {
		alert("프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[4]");
		return;
	}

	printTTPMultiEquipSmallBarcode(arr);
}

// 선택 장비 삭제
function EquipDelete() {
	var arr = new Array();

	if (!CheckSelected()){
		alert("선택된 장비가 없습니다.");
		return;
	}

	var frm;
	for (var i = 0;i < document.forms.length; i++) {
		frm = document.forms[i];

		if (frm.name.substr(0, 10)=="frmBuyPrc_") {
			if (frm.cksel.checked == true) {
				arr.push(frm.idx.value);
			}
		}
	}

	if (arr.length < 1) {
		alert("선택된 장비자산이 없습니다.");
		return;
	}

	if(confirm("선택된 장비를 삭제하시겠습니까?")) {
		document.frm.arridx.value = arr.toString();
		document.frm.mode.value = "equipmentDelete";
		document.frm.method="POST";
		document.frm.action="do_equipment.asp";
		document.frm.target = "procFrame";
		document.frm.submit();
	}

}

//이미지 확대화면 새창으로 보여주기
function jsImgView(sImgUrl){
	var wImgView;

	wImgView = window.open('/common/equipment/pop_equipment_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

//결제요청서보기
function pop_paymentrequestidx(paymentrequestidx){
	var pop_paymentrequestidx = window.open('/admin/approval/payreqList/regPayRequest.asp?ipridx='+paymentrequestidx,'pop_paymentrequestidx','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_paymentrequestidx.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="POST" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="arridx" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<input type="checkbox" name="ipgocheck" value="on" <% if ipgocheck="on" then  response.write " checked" %> onclick="EnDisabledDateBox(this)">
		구매일자 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<input type="checkbox" name="outcheck" value="on" <% if outcheck="on" then  response.write " checked" %> onclick="EnDisabledDateBox2(this)">
		폐기일자 : <% DrawDateBoxdynamic yyyy3,"yyyy3",yyyy4,"yyyy4",mm3,"mm3",mm4,"mm4",dd3,"dd3",dd4,"dd4" %>
		&nbsp;
		장비구분 : <% DrawEquipMentGubun "10","equip_gubun",equip_gubun ," onchange='NextPage("""");'" %>
		&nbsp;
		상태 : <% DrawEquipMentGubun "50","state",state," onchange='NextPage("""");'" %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="NextPage('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		자산구분 : <% drawEquipmentAccountCode "accountGubun" ,accountGubun, "" %>
		<!--
		&nbsp;
		부서 : <%= drawChSelectBoxDepartment("department_id", department_id,"") %>
		-->
		&nbsp;
		손익부서 :
		<input type="text" name="BIZSECTION_CD" value="<%= BIZSECTION_CD %>" size="15"  class="text_ro"> <input type="text" name="BIZSECTION_NM" value="<%= BIZSECTION_NM %>" class="text_ro" size="15">
		<input type="button" class="button" value="X" onClick="jsClearPart()">
		<a href="javascript:jsGetPart();"> <img src="/images/icon_search.jpg" border="0"></a>
		&nbsp;
		사용자(사용처) :
		<input type="text" name="using_username" value="<%=using_username%>">
		&nbsp;
		회계자산관리코드 :
		<input type="text" name="accountassetcode" value="<%=accountassetcode%>">
		&nbsp;
		결제요청서IDX :
		<input type="text" name="paymentrequestidx" value="<%=paymentrequestidx%>">
		<!--
		<% drawpartneruser "using_userid", using_userid ," onchange='NextPage("""");'" %>
		-->
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		장비코드 : <input type="text" name="equip_code" value="<%=equip_code%>">
		&nbsp;
		제품명 : <input type="text" name="equip_name" value="<%=equip_name%>" onkeypress="if(event.keyCode==13) {NextPage(''); return false;}">
		&nbsp;
		<!--
		제조사 : <input type="text" name="manufacture_company" value="<%=manufacture_company%>">
		-->
		구매처 : <input type="text" name="buyCompanyName" value="<%=buyCompanyName%>" onkeypress="if(event.keyCode==13) {NextPage(''); return false;}">
		&nbsp;
		장비시리얼 : <input type="text" name="manufacture_sn" value="<%=manufacture_sn%>">
		&nbsp;
		<input type="checkbox" name="only1000" value="Y" <% if (only1000 = "Y") then %>checked<% end if %> > 잔존가치 1000원 자산만
		&nbsp;
		<input type="checkbox" name="onlyusing" value="Y" <% if (onlyusing = "Y") then %>checked<% end if %> > 삭제내역 제외
		&nbsp;
		<input type="checkbox" name="dispmainimageyn" value="Y" <% if dispmainimageyn = "Y" then %> checked<% end if %> > 이미지보기
	</td>
</tr>
</table>
<!-- 검색 끝 -->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" class="button" onclick="pageexcelsheet();" value="현재페이지엑셀출력">
		&nbsp;&nbsp;
		<input type="button" class="button" name="equip_barcode_print" value="선택 바코드출력(80x50)" onclick="EquipBarcodePrint();">
		&nbsp;&nbsp;
		<input type="button" class="button" name="equip_barcode_print" value="선택 바코드출력(45x22)" onclick="EquipSmallBarcodePrint();">
		<% if C_ADMIN_AUTH or C_MngPart or C_PSMngPart then %>
		&nbsp;&nbsp;
		<input type="button" class="button_auth" name="equip_delete" value="선택장비 삭제" onclick="EquipDelete();">
		<% end if %>
	</td>
	<td align="right">
		<input type="button" class="button" onclick="pop_Equipmentreg('');" value="신규등록">
		<input type="button" class="button" onclick="popEtcBar();" value="(임시)바코드 출력">
		<%
		'/관리자, 재경
		if C_ADMIN_AUTH or C_MngPart or C_PSMngPart then
		%>
			<input type="button" class="button" onclick="popcodemanager();" value="코드관리">
		<% end  if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25" align="left">
		검색결과 : <b><%= oequip.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oequip.FTotalPage %></b>
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25" align="right">
		<select name="sorttype">
			<option value="" <% if sorttype="" then response.write " selected" %>>선택</option>
			<option value="1" <% if sorttype="1" then response.write " selected" %>>IDX(최근순)</option>
			<option value="2" <% if sorttype="2" then response.write " selected" %>>구매일자(최근순)</option>
			<option value="3" <% if sorttype="3" then response.write " selected" %>>구매일자(과거순)</option>
		</select>
	</td>
</tr>
</form>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>IDX</td>

	<% if dispmainimageyn="Y" then %>
		<td>이미지</td>
	<% end if %>

	<td>장비코드</td>
	<td>회계자산<br>관리코드</td>
	<td>자산구분</td>
	<td>결제요청서<br>IDX</td>
	<td>손익부서</td>
	<td>위치</td>
	<td>사용자<br>(사용처)</td>
	<td>장비구분</td>
	<td>제품명</td>
	<!--<td>장비시리얼</td>
	<td>구매처</td>-->
	<td>구매일자</td>
	<td>구매원가<br>(공급가)</td>
	<td><%= year(dateadd("yyyy",-1,date)) %>년<br>잔존가치</td>
	<td>월별<br>감가상각</td>
	<!--<td>구매가</td>-->
	<td><%= month(date) %>월말<br>자산가치</td>
	<td>상태</td>
	<td>폐기일자</td>
	<td>사용<br>여부</td>
	<!--<td>라벨<br>출력</td>
	<td>바코드출력</td>-->
</tr>
<% if oequip.FResultCount > 0 then %>
	<% for i=0 to oequip.FResultCount - 1 %>
	<%
	totalcurrsum = totalcurrsum + oequip.FItemList(i).Fbuy_sum
	totaljasan	 = totaljasan + oequip.FItemList(i).GetCurrentvalue
	%>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="equip_code" value="<%= oequip.FItemList(i).Fequip_code %>">
	<input type="hidden" name="AccountGubunName" value="<%= oequip.FItemList(i).GetAccountGubunName %>">
	<input type="hidden" name="equip_name" value="<%= oequip.FItemList(i).Fequip_name %>">
	<input type="hidden" name="buy_date" value="<%= oequip.FItemList(i).Fbuy_date %>">
	<input type="hidden" name="idx" value="<%= oequip.FItemList(i).Fidx %>">
	<% if oequip.FItemList(i).Fisusing = "Y" then %>
		<tr align="center" bgcolor="#FFFFFF" height="25">
	<% else %>
		<tr align="center" bgcolor="#f1f1f1" height="25">
	<% end if %>

		<td width=20><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td width=50>
			<%= oequip.FItemList(i).Fidx %>
		</td>

		<% if dispmainimageyn="Y" then %>
			<td width=50>
				<% if oequip.FItemList(i).fequip_mainimage<>"" then %>
					<img src="<%= getThumbImgFromURL(oequip.FItemList(i).fequip_mainimage, 50, 50,"true","false") %>" border="0" width=50 height=50 onclick="jsImgView('<%= oequip.FItemList(i).fequip_mainimage %>');" alt="누르시면 확대 됩니다">
				<% end if %>
			</td>
		<% end if %>

		<td width=130>
			<a href="javascript:pop_Equipmentreg('<%= oequip.FItemList(i).Fidx %>');" onfocus="this.blur()">
			<%= oequip.FItemList(i).Fequip_code %></a>
		</td>
		<td width=100>
			<%= oequip.FItemList(i).faccountassetcode %>
		</td>
		<td width=80>
			<%= oequip.FItemList(i).GetAccountGubunName %>
		</td>
		<td width=70>
			<a href="#" onclick="pop_paymentrequestidx('<%= oequip.FItemList(i).fpaymentrequestidx %>'); return false;">
			<%= oequip.FItemList(i).fpaymentrequestidx %></a>
		</td>
		<td width=100>
			<%= oequip.FItemList(i).FBIZSECTION_NM %>
		</td>
		<td width=100>
			<%= oequip.FItemList(i).Flocate_gubun_name %>
		</td>
		<td width=100>
			<%= oequip.FItemList(i).fusingusername %>
			<% if oequip.FItemList(i).fstatediv <> "Y" then %>
				<font color="red">[퇴사]</font>
			<% end if %>
	
			<% if oequip.FItemList(i).Fusing_userid <> "" then %>
				<!-- <Br><%= oequip.FItemList(i).Fusing_userid %> -->
			<% end if %>
		</td>
		<td width=100>
			<%= oequip.FItemList(i).Fequip_gubun_name %>
		</td>
		<td align="left">
			<%= oequip.FItemList(i).Fequip_name %>
		</td>
		<!--<td>
			<%'= oequip.FItemList(i).fmanufacture_sn %>
		</td>
		<td>
			<%'= oequip.FItemList(i).fbuy_company_name %>
		</td>-->
		<td width=80>
			<%= oequip.FItemList(i).Fbuy_date %>
		</td>
		<td align="right" width=70>
			<%= formatNumber(oequip.FItemList(i).fwonga_cost,0) %>
		</td>
		<td align="right" width=70>
			<%= oequip.FItemList(i).getremainValue %>
		</td>
		<td align="right" width=70>
			<% if (oequip.FItemList(i).FmonthlyDeprice <> 0) then %>
				<%= formatNumber(oequip.FItemList(i).FmonthlyDeprice,0) %>
			<% end if %>
		</td>
		<!--<td align="right" width=70>
			<%= formatNumber(oequip.FItemList(i).Fbuy_sum,0) %>
		</td>-->
		<td align="right" width=70>
			<% if oequip.FItemList(i).getCurrentValue<>"" then %>
				<font color="#EE3333"><%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%></font>
			<% else %>
				<%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%>
			<% end if %>
		</td>
		<td width=80>
			<%= oequip.FItemList(i).fstate_name %>
		</td>
		<td width=80>
			<%= oequip.FItemList(i).fout_date %>
		</td>
		<td width=30>
			<%= oequip.FItemList(i).fisusing %>
		</td>
		<!--
		<td width=30>
			<a href="javascript:ExcelSheet('<%= oequip.FItemList(i).Fidx %>');">
			<img src="/images/iexcel.gif" border="0"></a>
		</td>
		<td align="center" width=250>
			<Br>
			<a href="javascript:barcode('<%= oequip.FItemList(i).Fequip_code %>');" onfocus="this.blur()">
			<img srcXXXX="http://company.10x10.co.kr/barcode/barcode.asp?image=3&type=23&data=<%= trim(oequip.FItemList(i).Fequip_code) %>&height=30&barwidth=1&Size=7" border=0></a>
			<Br>
		</td>
		-->
	</tr>
	</form>
	<% next %>
	
	<!--
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" colspan=10>총계</td>
		<td align="right"></td>
		<td align="right"><%= formatNumber(totalcurrsum,0) %></td>
		<td align="right"><font color="#EE3333"><%= formatNumber(totaljasan,0) %></font></td>
		<td align="right"></td>
		<td align="right"></td>
	</tr>
	-->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="25" align="center">
	    	<% if oequip.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oequip.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
	
			<% for i=0 + oequip.StartScrollPage to oequip.FScrollCount + oequip.StartScrollPage - 1 %>
				<% if i>oequip.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>
	
			<% if oequip.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>
<iframe name="procFrame" src="about:blank" width="0" height="0"></iframe>
<script type="text/javascript">
	EnDisabledDateBox(document.frm.ipgocheck);
	EnDisabledDateBox2(document.frm.outcheck);
</script>

<%
set oequip = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
