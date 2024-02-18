<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/etcmeachulcls.asp"-->
<%

'// 사용안함
1

'// 업무협조>>[공통]세금계산서발행요청 에서 신규발행하면 나오는 페이지
'// [OFF]오프_가맹관리>>가맹점정산관리(매출) 에서 발행요청 하면 나오는 페이지
'// http://webadmin.10x10.co.kr/cscenter/ordermaster/ordermaster_detail.asp?orderserial=12021576159 에서 증빙서류 발급 -> 세금계산서 발행에서 사용되는 페이지

dim socno, socname, ceoname, socaddr, socstatus, socevent, managername, managerphone, managermail
dim taxtype, totalpricesum, itemname, totalsuply

dim etcstring, billdiv
dim orderserial, issuetype, idx
dim i, strSql

dim taxwritedate, previssuecount, userid, orderidx
dim errMSG

dim sellBizCd, selltype, taxissuetype

function Is3PLShopid(shopid)
	dim sqlStr

	Is3PLShopid = False

	sqlStr = " select top 1 p.id as shopid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p "
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_partner_tpl t "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		p.tplcompanyid = t.tplcompanyid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p.id = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	and IsNull(p.tplcompanyid, '') <> '' "
	''response.write sqlStr
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF then
		Is3PLShopid = True
	end if
	rsget.close
end function

function Get3PLUpcheInfoByShopid(shopid, byRef tplcompanyid, byRef tplcompanyname, byRef tplgroupid, byRef tplbillUserID)
	dim sqlStr

	sqlStr = " select top 1 t.tplcompanyid, t.tplcompanyname, t.groupid as tplgroupid, billUserID as tplbillUserID "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner p "
	sqlStr = sqlStr + " 	join db_partner.dbo.tbl_partner_tpl t "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		p.tplcompanyid = t.tplcompanyid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and p.id = '" + CStr(shopid) + "' "
	sqlStr = sqlStr + " 	and IsNull(p.tplcompanyid, '') <> '' "
	'' response.write sqlStr
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF then
		tplcompanyid = rsget("tplcompanyid")
		tplcompanyname = db2html(rsget("tplcompanyname"))
		tplgroupid = rsget("tplgroupid")
		tplbillUserID = rsget("tplbillUserID")
	end if
	rsget.close
end function

orderserial = request("orderserial")
issuetype 	= request("issuetype")
idx 		= request("idx")


itemname = "XXXX 외 X 건"

sellBizCd = ""
selltype = "0"

if (orderserial <> "") or (issuetype = "orderserial") then
	'// 소비자 매출
	etcstring 		= orderserial
	billdiv 		= "01"
	issuetype 		= "orderserial"
	taxtype			= "Y"

	sellBizCd 		= "0000000101"		'// 온라인(공통)
	selltype 		= "20166"			'// B2C
	taxissuetype	= "C"

	'==============================================================================
	dim ojumun
	set ojumun = new COrderMaster

	if (orderserial <> "") then
	    ojumun.FRectOrderSerial = orderserial
	    ojumun.QuickSearchOrderMaster
	end if

	if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
	    ojumun.FRectOldOrder = "on"
	    ojumun.QuickSearchOrderMaster
	end if

	if (ojumun.FResultCount < 1) and (errMSG = "") then
		errMSG = "잘못된 주문번호 입니다."
	else
		orderidx = ojumun.FOneItem.Fidx
		taxwritedate = getMayTaxDate(ojumun.FOneItem.Fipkumdate)
	end if

	'==============================================================================
	strSql =	"select ( select " &_
			"			Case " &_
			"				When count(idx)>1 Then max(itemname) + '외 ' + Cast((count(idx)-1) as varchar) + '건' " &_
			"				Else max(itemname) " &_
			"			End " &_
			"		from db_order.[dbo].tbl_order_detail " &_
			"		where orderserial='" & orderserial & "' and itemid<>0 and cancelyn='N' group by orderserial " &_
			"	) as itemname " &_
			"	, subtotalprice, accountdiv, IsNull(sumPaymentEtc, 0) as sumPaymentEtc " &_
			"from db_order.[dbo].tbl_order_master " &_
			"Where orderserial = '" & orderserial & "'"
	rsget.Open strSql, dbget, 1

	if Not(rsget.EOF or rsget.BOF) then
		itemname = rsget("itemname")

		if (CStr(rsget("accountdiv")) = "7") or (CStr(rsget("accountdiv")) = "20") then
			'무통장, 실시간이체 : 전체금액
			totalpricesum = rsget("subtotalprice")
		else
			'보조결제금액만
			totalpricesum = rsget("sumPaymentEtc")
		end if
	end if
	rsget.close

	'==============================================================================
	dim oTax
	set oTax = new CTax

	if (errMSG = "") then
		oTax.FCurrPage = 1
		oTax.FPageSize = 100
		'oTax.FRectsearchDiv = "Y"					'발행된 내역만
		oTax.FRectsearchBilldiv = "01"				'소비자매출
		oTax.FRectsearchKey = "t1.userid"
		oTax.FRectDelYn = "N"

		if (ojumun.FOneItem.FUserID <> "") then
			oTax.FRectsearchString = ojumun.FOneItem.FUserID
			userid = ojumun.FOneItem.FUserID
		else
			oTax.FRectsearchString = "----"
		end if

		oTax.GetTaxList

		previssuecount = oTax.FTotalCount
	end if
end if

dim tplcompanyid, tplcompanyname, tplgroupid, tplbillUserID
dim tplsocno, tplsocname, tplceoname, tplsocaddr, tplsocstatus, tplsocevent, tplmanagername, tplmanagerphone, tplmanagermail

if (issuetype = "etcmeachul") then
	'// 기타매출
	billdiv = "51"
	orderidx = idx

	'==========================================================================
	'기타매출
	dim oetcmeachul
	set oetcmeachul = new CEtcMeachul
	oetcmeachul.FRectidx = idx
	oetcmeachul.getOneEtcMeachul

	sellBizCd 		= oetcmeachul.FOneItem.Fbizsection_cd
	selltype 		= oetcmeachul.FOneItem.Fselltype
	taxissuetype	= "E"

	'oetcmeachul.FOneItem.Ftotalsum '총 발행금액을 총공급가로 함.(부가세포함금액)


	'==========================================================================
	'삽아이디에서 그룹코드 추출
	dim opartner
	set opartner = new CPartnerUser

	opartner.FCurrpage = 1
	opartner.FPageSize = 100
	opartner.FRectDesignerID = oetcmeachul.FOneItem.Fshopid
	opartner.FRectUserDiv = "all"

	opartner.GetPartnerNUserCList

	'==========================================================================
	'그룹코드에서 세금계산서/정산담당자 정보 추출
	dim ogroup
	set ogroup = new CPartnerGroup

	if (opartner.FResultCount > 0) then
		ogroup.FRectGroupid = opartner.FPartnerList(0).FGroupID
		ogroup.GetOneGroupInfo
	else
		ogroup.FResultCount = 0
	end if

	if (opartner.FResultCount < 1) then
		errMSG = "잘못된 브랜드입니다."
	elseif (ogroup.FResultCount < 1) then
		errMSG = "그룹코드가 지정되어 있지 않은 업체입니다."
	else
		socno			= ogroup.FOneItem.Fcompany_no
		socname			= ogroup.FOneItem.FCompany_name
		ceoname			= ogroup.FOneItem.Fceoname
		socaddr			= ogroup.FOneItem.Fcompany_address & " " & ogroup.FOneItem.Fcompany_address2
		socstatus		= ogroup.FOneItem.Fcompany_uptae
		socevent		= ogroup.FOneItem.Fcompany_upjong
		managername		= ogroup.FOneItem.Fjungsan_name
		managerphone	= ogroup.FOneItem.Fjungsan_hp
		managermail		= ogroup.FOneItem.Fjungsan_email

		taxtype			= "Y"
		totalpricesum	= oetcmeachul.FOneItem.Ftotalsum
		itemname		= oetcmeachul.FOneItem.Ftitle
		etcstring		= idx
	end if

	'==========================================================================
	'' 삽아이디에서 3PL 업체인지 확인
	if (Is3PLShopid(oetcmeachul.FOneItem.Fshopid) = True) then
		Call Get3PLUpcheInfoByShopid(oetcmeachul.FOneItem.Fshopid, tplcompanyid, tplcompanyname, tplgroupid, tplbillUserID)

		dim otplgroup
		set otplgroup = new CPartnerGroup

		otplgroup.FRectGroupid = tplgroupid
		otplgroup.GetOneGroupInfo

		if (otplgroup.FResultCount < 1) then
			errMSG = "3PL그룹코드가 지정되어 있지 않은 업체입니다."
		else
			tplsocno			= otplgroup.FOneItem.Fcompany_no
			tplsocname			= otplgroup.FOneItem.FCompany_name
			tplceoname			= otplgroup.FOneItem.Fceoname
			tplsocaddr			= otplgroup.FOneItem.Fcompany_address & " " & otplgroup.FOneItem.Fcompany_address2
			tplsocstatus		= otplgroup.FOneItem.Fcompany_uptae
			tplsocevent			= otplgroup.FOneItem.Fcompany_upjong
			tplmanagername		= otplgroup.FOneItem.Fjungsan_name
			tplmanagerphone		= otplgroup.FOneItem.Fjungsan_hp
			tplmanagermail		= otplgroup.FOneItem.Fjungsan_email

			billdiv = "99"
		end if
	end if
end if

if (issuetype <> "") and (orderidx <> "") then
	'==========================================================================
	''기발행 세금계산서인지 체크

	set oTax = new CTax

	oTax.FRectsearchKey = " t1.orderidx "
	oTax.FRectsearchString = CStr(orderidx)
	oTax.FRectDelYn = "N"

	oTax.GetTaxList

	if oTax.FResultCount > 0 then
		if oTax.FTaxList(0).FisueYn="Y" then
			if (errMSG = "") then
				errMSG = "이미 발행된 세금계산서가 있습니다.\n\n재발행 하시려면 거래처에 [취소요청]후 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다"
			end if
		else
			if (errMSG = "") then
				errMSG = "발행대기중인 세금계산서가 있습니다.\n\n재발행 하시려면 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다"
			end if
		end if
	end if
end if

%>
<script language="javascript">
var errMSG = "<%= errMSG %>";

function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

// 사업자등록증 확인 처리
function chkSheetOk(){
	if (confirm('사업자등록증을 확인하셨습니까?')){
		document.frm_trans.mode.value="BusiOk";
		document.frm_trans.submit();
	}
}

// 요청서 출력 처리
function GotoTaxPrint(){
    alert('네오포트는 더이상 지원하지 않습니다.');
    return;
	if (confirm('세금계산서를 발행하시겠습니까?')){
		document.frm_trans.mode.value="sheetOk";
		document.frm_trans.submit();
	}
}

// 요청서 삭제
function GotoTaxDel(){
	if (confirm('요청서를 삭제 하시겠습니까?\n\n계산서가 발행된 경우 발행이 취소된후 삭제하시기 바랍니다.')){
		document.frm_trans.mode.value="sheetDel";
		document.frm_trans.submit();
	}
}

// 세금계산서 보기
function goView(tax_no, b_biz_no, s_biz_no)
{
	<% if (application("Svr_Info")="Dev") then %>
		// 테스트
		window.open("http://ifs.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
	<% else %>
		// 실서버
		window.open("http://web1.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + tax_no + "&cur_biz_no="+b_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
	<% end if %>
}

function goView2(tax_no, b_biz_no, s_biz_no){
	<% if (application("Svr_Info")="Dev") then %>
		// 테스트
		window.open("http://ifs.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&cur_biz_no="+s_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
	<% else %>
		// 실서버
		window.open("http://web1.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&cur_biz_no="+s_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
	<% end if %>
}

function goView_Bill36524(tax_no, b_biz_no)
{
		window.open("http://www.bill36524.com/popupBillTax.jsp?NO_TAX=" + tax_no + "&NO_BIZ_NO="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
}

function setRegisterInfo()
{
	var isUnitTax, isUnitTaxTypeChanged;

	// 2012-01-01 부터 단위과세 세금계산서를 발행한다.
	if (frm.yyyymmdd_register.value == "") {
		// 작성일이 없으면 단위과세 세금계산서
		frm.prev_yyyymmdd_register.value = "2012-01-01";
		isUnitTax = true;
		isUnitTaxTypeChanged = false;
	} else {
		if (frm.yyyymmdd_register.value >= "2012-01-01") {
			isUnitTax = true;
		} else {
			isUnitTax = false;
		}

		if (((frm.yyyymmdd_register.value >= "2012-01-01") && (frm.prev_yyyymmdd_register.value >= "2012-01-01")) || ((frm.yyyymmdd_register.value < "2012-01-01") && (frm.prev_yyyymmdd_register.value < "2012-01-01"))) {
			isUnitTaxTypeChanged = false;
		} else {
			isUnitTaxTypeChanged = true;
		}

		frm.prev_yyyymmdd_register.value = frm.yyyymmdd_register.value;
	}

	if (isUnitTaxTypeChanged == true) {
		alert("작성일이 변경되었습니다. 2012-01-01 부터 단위과세 세금계산서를 발행합니다.");
	}

	// ================================================================
	// cs_taxsheetcls.asp 에서 가져온다.
	// ================================================================
	frm.reg_socno.value = "<%= TENBYTEN_SOCNO %>";
	frm.reg_subsocno.value = "<%= TENBYTEN_SUBSOCNO %>";
	frm.reg_socname.value = "<%= TENBYTEN_SOCNAME %>";
	frm.reg_ceoname.value = "<%= TENBYTEN_CEONAME %>";
	frm.reg_socaddr.value = "<%= TENBYTEN_SOCADDR %>";
	frm.reg_socstatus.value = "<%= TENBYTEN_SOCSTATUS %>";
	frm.reg_socevent.value = "<%= TENBYTEN_SOCEVENT %>";
	frm.reg_managername.value = "<%= TENBYTEN_MANAGERNAME %>";
	frm.reg_managerphone.value = "<%= TENBYTEN_MANAGERPHONE %>";
	frm.reg_managermail.value = "<%= TENBYTEN_MANAGERMAIL %>";

	// ===================================================================
	if(frm.billdiv.value == "52") {
		// 공급자 (주)블루앤더블유
        alert('발행불가');
        return;

		frm.reg_socno.value = "101-85-29011";
		frm.reg_socname.value = "(주)블루앤더블유";
		frm.reg_ceoname.value = "이문재";
		frm.reg_socaddr.value = "서울 종로구 이화동 197-1 이엠씨빌딩 2층";
		frm.reg_socstatus.value = "서비스,도소매";
		frm.reg_socevent.value = "전자상거래 등";
		frm.reg_managername.value = "신희영";
		frm.reg_managerphone.value = "02-554-2033";
		frm.reg_managermail.value = "accounts@10x10.co.kr";
	}

	if(frm.billdiv.value == "55") {
		// 공급자 (주)에이플러스비
        alert('발행불가');
        return;

		frm.reg_socno.value = "101-86-64617";
		frm.reg_socname.value = "(주)에이플러스비";
		frm.reg_ceoname.value = "이창우";
		frm.reg_socaddr.value = "서울 종로구 이화동 197-1 이엠씨빌딩2층";
		frm.reg_socstatus.value = "도소매";
		frm.reg_socevent.value = "전자상거래";
		frm.reg_managername.value = "김민환";
		frm.reg_managerphone.value = "070-7515-5410"
		frm.reg_managermail.value = "gogo27@10x10.co.kr"
	}

	if (isUnitTax == true) {
		if(frm.billdiv.value == "53") {
			// 공급자 (주)아이띵소

			frm.reg_subsocno.value = "0001";
		}

		if(frm.billdiv.value == "54") {
			// 공급자 (주)텐바이텐 리빙

			frm.reg_subsocno.value = "0002";
		}

		if(frm.billdiv.value == "99") {
			// 공급자 3PL업체

			frm.reg_socno.value = "<%= tplsocno %>";
			frm.reg_socname.value = "<%= tplsocname %>";
			frm.reg_ceoname.value = "<%= tplceoname %>";
			frm.reg_socaddr.value = "<%= tplsocaddr %>";
			frm.reg_socstatus.value = "<%= tplsocstatus %>";
			frm.reg_socevent.value = "<%= tplsocevent %>";
			frm.reg_managername.value = "<%= tplmanagername %>";
			frm.reg_managerphone.value = "<%= tplmanagerphone %>";
			frm.reg_managermail.value = "<%= tplmanagermail %>";
		}
	} else {
		if(frm.billdiv.value == "53") {
			// 공급자 (주)아이띵소

			frm.reg_socno.value = "101-85-36109";
			frm.reg_socname.value = "(주)아이띵소";
			frm.reg_ceoname.value = "이문재";
			frm.reg_socaddr.value = "서울 종로구 동숭동 1-45 자유빌딩 4층";
			frm.reg_socstatus.value = "도소매";
			frm.reg_socevent.value = "팬시용품";
			frm.reg_managername.value = "김민환";
			frm.reg_managerphone.value = "02-554-2033";
			frm.reg_managermail.value = "accounts@10x10.co.kr";
		}

		if(frm.billdiv.value == "54") {
			// 공급자 (주)텐바이텐 리빙

			frm.reg_socno.value = "101-85-38408";
			frm.reg_socname.value = "(주)텐바이텐 리빙";
			frm.reg_ceoname.value = "이문재";
			frm.reg_socaddr.value = "서울 종로구 동숭동 1-45 자유빌딩 1층";
			frm.reg_socstatus.value = "도소매";
			frm.reg_socevent.value = "식료품,팬시용품";
			frm.reg_managername.value = "김민환";
			frm.reg_managerphone.value = "02-554-2033";
			frm.reg_managermail.value = "accounts@10x10.co.kr";
		}
	}
}

function SearchSocno() {
	if (frm.socno.value == "") {
		alert("사업자번호를 입력하세요.");
		return;
	}

	if (frm.socno.value.length != 12) {
		alert("사업자번호는 아래와 같은 형식으로 입력하세요.\n\n000-00-00000");
		return;
	}

	icheckframe.location.href="isearchframe.asp?socno=" + frm.socno.value;
	// location.href="isearchframe.asp?socno=" + frm.socno.value;
}

function popMeachulDetailList() {
	if (frm.taxissuetype.value != "E") {
		alert("기타매출인 경우에만 내역을 추가할 수 있습니다.");
		return;
	}

	var popwin = window.open('pop_etc_meachul_list.asp?idx=<%= etcstring %>','popMeachulDetailList','width=1000, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function setCompanyInfo(subsocno, socname, ceoname, socaddr, socstatus, socevent, managername, managerphone, managermail)
{
	frm.subsocno.value = subsocno;
	frm.socname.value = socname;
	frm.ceoname.value = ceoname;
	frm.socaddr.value = socaddr;
	frm.socstatus.value = socstatus;
	frm.socevent.value = socevent;
	frm.managername.value = managername;
	frm.managerphone.value = managerphone;
	frm.managermail.value = managermail;
}

function CalcPrice()
{
	if (frm.totalsuply.value == "") { return; }

	if (frm.taxtype.value.length<1){
		alert('과세구분을 입력하세요.');
		return;
	}

	if (frm.totalsuply.value*0 != 0) { alert("잘못된 값을 입력했습니다."); return; }

	frm.totalsuply2.value = frm.totalsuply.value;
	frm.totalsuplysum.value = frm.totalsuply.value;

    if (frm.ckHand.checked){
        frm.totaltaxsum.value = frm.totaltax.value;
    }else{
		if (frm.taxtype.value == "Y") {
			frm.totaltax.value = parseInt(frm.totalsuply.value*0.1);
			frm.totaltaxsum.value = parseInt(frm.totalsuply.value*0.1);
		} else {
			frm.totaltax.value = 0;
			frm.totaltaxsum.value = 0;
		}
	}

	frm.totalpricesum.value = parseInt(frm.totalsuply.value) + parseInt(frm.totaltaxsum.value);
	frm.totalpricesum2.value = parseInt(frm.totalsuply.value) + parseInt(frm.totaltaxsum.value);
	frm.totalpricesum3.value = parseInt(frm.totalsuply.value) + parseInt(frm.totaltaxsum.value);
}

function CalcPriceWithPrice()
{
	if (frm.totalpricesum.value == "") { return; }

	if (frm.taxtype.value.length<1){
		alert('과세구분을 입력하세요.');
		return;
	}

	if (frm.totalpricesum.value*0 != 0) { alert("잘못된 값을 입력했습니다."); return; }

	frm.totalpricesum2.value = frm.totalpricesum.value;
	frm.totalpricesum3.value = frm.totalpricesum.value;

	if (frm.taxtype.value == "Y") {
		// 세액은 공급가를 구하고 0.1 후 반올림 해주면 된다.
		frm.totaltax.value = Math.round(1.0 * frm.totalpricesum.value / 1.1 / 10.0);
		frm.totaltaxsum.value = frm.totaltax.value;
	} else {
		frm.totaltax.value = 0;
		frm.totaltaxsum.value = 0;
	}

	frm.totalsuply.value = frm.totalpricesum.value - frm.totaltax.value;
	frm.totalsuply2.value = frm.totalsuply.value;
	frm.totalsuplysum.value = frm.totalsuply.value;
}


function doRegisterSheet(){

	if (errMSG != "") {
		alert(errMSG);
		return;
	}

	if (frm.issuetype.value != "") {
		if ((frm.issuetype.value == "orderserial") && (frm.billdiv.value != "01")) {
			alert('소비자 매출만 작성할 수 있습니다.');
			frm.billdiv.focus();
			return;
		}

		if ((frm.issuetype.value == "etcmeachul") && (frm.billdiv.value != "51") && (frm.billdiv.value != "99")) {
			alert('기타매출만 작성할 수 있습니다.');
			frm.billdiv.focus();
			return;
		}

		if(frm.etcstring.value == "") {
			alert('주문번호 또는 기타매출 코드가 비고에 입력되어 있어야 합니다.');
			return;
		}
	}

	if ((frm.selltype.value == "20036") && (frm.taxtype.value != "0")) {
		alert('계정과목이 영세인 경우 영세계산서만 작성가능합니다.');
		return;
	} else if ((frm.selltype.value != "20036") && (frm.taxtype.value == "0")) {
		alert('계정과목이 영세가 아니면 영세계산서를 작성할 수 없습니다.');
		return;
	}

	if(frm.billdiv.value == "0") {
		alert('공급자를 선택하세요.');
		return;
	}

	if (frm.socname.value.length<1){
		alert('사업자 등록상의 회사명을 입력하세요.');
		frm.socname.focus();
		return;
	}

	if (frm.ceoname.value.length<1){
		alert('사업자 등록상의 대표자명을 입력하세요.');
		frm.ceoname.focus();
		return;
	}

	if (frm.socno.value.length<1){
		alert('사업자 등록 번호를 입력하세요.');
		frm.socno.focus();
		return;
	}

	if (frm.socaddr.value.length<1){
		alert('사업자 등록상의 주소를 입력하세요.');
		frm.socaddr.focus();
		return;
	}

	if (frm.socstatus.value.length<1){
		alert('사업자 등록상의 업태를 입력하세요.');
		frm.socstatus.focus();
		return;
	}

	if (frm.socevent.value.length<1){
		alert('사업자 등록상의 업종을 입력하세요.');
		frm.socevent.focus();
		return;
	}

	if (frm.managername.value.length<1){
		alert('담당자 성함을 입력하세요.');
		frm.managername.focus();
		return;
	}

	if (frm.managerphone.value.length<1){
		alert('담당자 전화번호를 입력하세요.');
		frm.managerphone.focus();
		return;
	}

	if (frm.managermail.value.length<1){
		alert('담당자 이메일주소를 입력하세요.');
		frm.managermail.focus();
		return;
	}

	if (frm.yyyymmdd_register.value.length<1){
		alert('작성일을 입력하세요.');
		return;
	}

	if (frm.itemname.value.length<1){
		alert('품목을 입력하세요.');
		return;
	}

	if (frm.totalsuply.value.length<1){
		alert('단가를 입력하세요.');
		return;
	}

	if (frm.taxtype.value.length<1){
		alert('과세구분을 입력하세요.');
		return;
	}

	if ((frm.subsocno.value.length != 0) && (frm.subsocno.value.length != 4)) {
		alert('종사업장번호를 4자리로 입력하세요');
		return;
	}

	if(frm.billdiv.value == "01") {
		if(frm.etcstring.value == "") {
			alert('비고에 주문번호 또는 출고코드를 입력하세요.');
			return;
		}
	} else if ((frm.etcstring.value != "") && (frm.billdiv.value != "03") && (frm.billdiv.value != "51") && (frm.billdiv.value != "99")) {
		alert('소비자매출/프로모션/기타매출에만 비고에 주문번호 또는 출고코드를 넣을 수 있습니다.');
		return;
	}

	if (frm.billdiv.value != "99") {
		if (frm.sellBizCd.value.length<1){
			alert('매출부서를 지정하세요.');
			return;
		}
	} else {
		if (frm.sellBizCd.value.length > 0){
			alert('3PL매출에는 부서를 지정할 수 없습니다.');
			return;
		}
	}

	if (frm.selltype.value.length<1){
		alert('매출계정을 지정하세요.');
		return;
	}

	if (frm.taxissuetype.value.length<1){
		alert('세부내역을 지정하세요.');
		return;
	}

	setRegisterInfo();

    if (confirm('세금계산서 발행신청을 하시겠습니까?')){
        document.frm.submit();
    }
}

function chgHandTax(comp){
    var txbox = comp.form.totaltax;

    if (comp.checked){
        txbox.readOnly = false;
        txbox.className = "writebox";
    }else{
        txbox.readOnly = true;
        txbox.className = "readonlybox";
    }
}

function popListPreviousCustomerTaxSheet(userid){
    var popwin=window.open("/cscenter/taxsheet/popListPreviousCustomerTaxSheet.asp?userid=" + userid,"popListPreviousCustomerTaxSheet","width=700,height=400,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function ReactMeachulDetailList(arrchk, tottaxsum) {
    var frm = document.frm;

    frm.totalpricesum.value = tottaxsum;
    frm.etcstring.value = arrchk;

    CalcPriceWithPrice();
}

</script>

<style type="text/css">
.readonlybox { border:0px; }
.writebox { border:0px; background:#E6E6E6; }
</style>



<table width="800" border="0" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td>

		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
			<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td colspan="2" align="left">
					<b>세금계산서 발행요청</b>
				</td>
			</tr>
			<tr height="25">
				<td align="center" width="120" bgcolor="#F0F0FD">요청자</td>
				<td bgcolor="#FFFFFF"><%= session("ssBctId") %></td>
			</tr>
			<tr height="25">
				<td align="center" width="120" bgcolor="#F0F0FD">설명</td>
				<td bgcolor="#FFFFFF">
					공급자는 <b>텐바이텐/아이띵소/유아러걸</b> 중에 하나를 선택하시면, 자동입력됩니다.<br>
					공급받는자는 등록번호에 하이픈(-)을 포함한 사업자번호를 입력하시면, 기존 데이타가 있을경우 자동입력됩니다.<br>
					<b>(검색 후, 담당자등의 내용이 상이할 경우, 수정입력하시면 됩니다.)</b><br>
					품목을 꼭 입력하시기 바랍니다.(현재 상품가액은 총액으로만 입력 가능합니다.)<br>
				</td>
			</tr>
		</table>

	</td>
</tr>
<tr height="70">
	<td>
    	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="post" onsubmit="return false;" action="doTaxOrder.asp">
			<input type=hidden name=mode value="tax_register_new">
			<input type=hidden name=issuetype value="<%= issuetype %>">
			<input type=hidden name=tplcompanyid value="<%= tplcompanyid %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    			<td height="25" width="10%"><b>부서</b></td>
    			<td align="left" bgcolor="#FFFFFF" width="40%">
    				<%= fndrawSaleBizSecCombo(true,"sellBizCd", sellBizCd,"") %>
    			</td>
    			<td height="25" width="10%"><b>계정</b></td>
    			<td align="left" bgcolor="#FFFFFF">
    				<% drawPartnerCommCodeBox true,"sellacccd","selltype", selltype,"" %>
    			</td>
    		</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    			<td height="25"><b>세부내역</b></td>
    			<td align="left" bgcolor="#FFFFFF">
    				<select class="select" name="taxissuetype">
    					<option value="">-선택-</option>
    					<option value="C" <% if (taxissuetype = "C") then %>selected<% end if %> >온라인주문</option>
    					<option value="E" <% if (taxissuetype = "E") then %>selected<% end if %> >기타매출</option>
    					<option value="S" <% if (taxissuetype = "S") then %>selected<% end if %> >출고리스트</option>
    					<option value="X" <% if (taxissuetype = "X") then %>selected<% end if %> >내역없음</option>
    				</select>
    			</td>
    			<td height="25"></td>
    			<td align="left" bgcolor="#FFFFFF">
    			</td>
    		</tr>
    	</table>
	</td>
</tr>
<tr>
	<td>

		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			<tr valign="top">
		        <td width="49%">
		        	<!-- 공급자정보 시작 -->
		        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		        			<td colspan="2" height="25"><b>공급자 정보</b></td>
		        			<td colspan="2">
		        				<select class="select" name="billdiv" onchange="setRegisterInfo()">
		        					<option value="0">공급자선택</option>
									<% if (billdiv <> "99") then %>
		        					<option value="01" <% if (billdiv = "01") then %>selected<% end if %>>소비자(customer)</option>
		        					<option value="02" <% if (billdiv = "02") then %>selected<% end if %>>가맹점(accounts)</option>
		        					<option value="03" <% if (billdiv = "03") then %>selected<% end if %>>프로모션(promotion)</option>
		        					<option value="51" <% if (billdiv = "51") then %>selected<% end if %>>기타매출(accounts)</option>
		        					<!-- option value="52">유아러걸(youareagirl)</option -->
									<!--
		        					<option value="53" <% if (billdiv = "53") then %>selected<% end if %>>아이띵소(ithinkso)</option>
									-->
		        					<option value="54" <% if (billdiv = "54") then %>selected<% end if %>>텐바이텐 리빙(living1010)</option>
		        					<!-- <option value="55">에이플러스비(aplusb)</option> -->
									<% else %>
									<option value="99" <% if (billdiv = "99") then %>selected<% end if %>><%= tplcompanyname %>(<%= tplbillUserID %>)</option>
									<% end if %>
		        				</select>
		        			</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">등록번호</td>
		        			<td colspan="3">
		        				<input type=text name="reg_socno" size=12 value="" class="readonlybox" readonly>
		        			</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td width="70" bgcolor="#F0F0FD" height="25">상호</td>
		        			<td><input type=text name="reg_socname" size=14 value="" border=0 class="readonlybox" readonly></td>
		        			<td width="70" bgcolor="#F0F0FD">대표자</td>
		        			<td><input type=text name="reg_ceoname" size=8 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">사업장주소</td>
		        			<td colspan="3"><input type=text name="reg_socaddr" size=40 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">업태</td>
		        			<td colspan=2><input type=text name="reg_socstatus" size=20 value="" class="readonlybox" readonly></td>
		        			<td bgcolor="#F0F0FD">종사업장번호</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">종목</td>
		        			<td colspan=2><input type=text name="reg_socevent" size=20 value="" class="readonlybox" readonly></td>
		        			<td><input type=text name="reg_subsocno" size=4 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">담당자</td>
		        			<td><input type=text name="reg_managername" size=14 value="" class="readonlybox" readonly></td>
		        			<td bgcolor="#F0F0FD">연락처</td>
		        			<td><input type=text name="reg_managerphone" size=14 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">이메일</td>
		        			<td colspan=3><input type=text name="reg_managermail" size=40 value="" class="readonlybox" readonly></td>
		        		</tr>
		        	</table>
		        	<!-- 공급자정보 끝 -->
		        </td>
		        <td>&nbsp;</td>
		        <td width="49%">
		        	<!-- 공급받는자정보 시작 -->
		        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		        			<td colspan="4" height="25"><b>공급받는자 정보</b></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">등록번호</td>
		        			<td colspan="3">
		        				<input type=text name="socno" size=12 value="<%= socno %>" class="writebox">
		        				<input type="button" class="button_s" value="검 색" onClick="SearchSocno()">
		        				<% if (userid <> "") then %>
		        					<input type="button" class="button_s" value="기존(<%= previssuecount %>)" onClick="popListPreviousCustomerTaxSheet('<%= userid %>')" <% if (previssuecount < 1) then %>disabled<% end if %>>
		        				<% end if %>
		        			</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td width="70" bgcolor="#F0F0FD" height="25">상호</td>
		        			<td align="left"><input type=text name="socname" size=14 value="<%= socname %>" border=0 class="writebox"></td>
		        			<td width="70" bgcolor="#F0F0FD">대표자</td>
		        			<td align="left"><input type=text name="ceoname" size=14 value="<%= ceoname %>" class="writebox"></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">사업장주소</td>
		        			<td align="left" colspan="3"><input type=text name="socaddr" size=40 value="<%= socaddr %>" class="writebox"></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">업태</td>
		        			<td colspan=2><input type=text name="socstatus" size=20 value="<%= socstatus %>" class="writebox"></td>
		        			<td bgcolor="#F0F0FD">종사업장번호</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">종목</td>
		        			<td colspan=2><input type=text name="socevent" size=20 value="<%= socevent %>" class="writebox"></td>
		        			<td><input type=text name="subsocno" size=4 value="" class="writebox"></td>
		        		</tr>
		        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">담당자</td>
		        			<td align="left"><input type=text name="managername" size=14 value="<%= managername %>" class="writebox"></td>
		        			<td bgcolor="#F0F0FD">연락처</td>
		        			<td align="left"><input type=text name="managerphone" size=14 value="<%= managerphone %>" class="writebox"></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">이메일</td>
		        			<td align="left" colspan="3"><input type=text name="managermail" size=40 value="<%= managermail %>" class="writebox"></td>
		        		</tr>
		        	</table>
		        	<!-- 공급받는자정보 끝 -->
		        </td>
			</tr>
		</table>

	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    <tr align="center" bgcolor="#F0F0FD">
				<td width="120" height="25">작성일</td>
				<td width="100">공급가액</td>
				<td width="100">과세구분</td>
				<td width="100">세액</td>
				<td width="100">합계금액</td>
				<td>비고</td>
			</tr>
		    <tr align="center" bgcolor="#FFFFFF">
				<td height="25">
					<input type="text" size="10" name="yyyymmdd_register" value="<%= taxwritedate %>" onClick="jsPopCal('frm','yyyymmdd_register');" style="cursor:hand;" class="writebox">
					<input type=hidden name=prev_yyyymmdd_register value="<%= taxwritedate %>">
				</td>
				<td><input type=text name="totalsuplysum" size=10 value="" class="readonlybox" readonly></td>
				<td>
					<select name=taxtype class="select" onchange="CalcPriceWithPrice()">
					<option value="">====</option>
					<option value="Y" <% if (taxtype = "Y") then %>selected<% end if %>>과세</option>
					<option value="N" <% if (taxtype = "N") then %>selected<% end if %>>면세</option>
					<option value="0" <% if (taxtype = "0") then %>selected<% end if %>>영세</option>
					</select>
				</td>
				<td><input type=text name="totaltaxsum" size=10 value="" class="readonlybox" readonly></td>
				<td><input type=text name="totalpricesum" size=10 value="<%= totalpricesum %>" class="writebox" onkeyup="CalcPriceWithPrice()"></td>
				<td>
					<input type=text name="etcstring" size=20 value="<%= etcstring %>" class="writebox">
					<input type=button class=button name="btnCombine" value="추가" onClick="popMeachulDetailList()">
				</td>
			</tr>
		</table>

	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    <tr align="center" bgcolor="#F0F0FD">
				<td width="30" height="25">월</td>
				<td width="30">일</td>
				<td>품목</td>
				<td width="50">규격</td>
				<td width="50">수량</td>
				<td width="100">단가</td>
				<td width="100">공급가액</td>
				<td width="100">세액</td>
			</tr>
		    <tr align="center" bgcolor="#FFFFFF">
				<td height="25"></td>
				<td></td>
				<td><input type=text name="itemname" size=40 value="<%= itemname %>" class="writebox"></td>
				<td></td>
				<td>1</td>
				<td><input type=text name="totalsuply" size=10 value="" class="writebox" onkeyup="CalcPrice()"></td>
				<td><input type=text name="totalsuply2" size=10 value="" class="readonlybox" readonly></td>
				<td><input type=text name="totaltax" size=10 value="" class="readonlybox" readonly  onKeyUp="CalcPrice();"></td>
			</tr>
		</table>

	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    <tr align="center" bgcolor="#F0F0FD">
				<td height="25"><b>합계금액</b></td>
				<td width="100">현금</td>
				<td width="100">수표</td>
				<td width="100">어음</td>
				<td width="100">외상미수금</td>
			</tr>
		    <tr align="center" bgcolor="#FFFFFF">
				<td height="25"><input type=text name="totalpricesum3" size=10 value="" class="readonlybox" readonly></td>
				<td>
				</td>
				<td></td>
				<td></td>
				<td>
					<input type=text name="totalpricesum2" size=10 value="" class="readonlybox" readonly>
				</td>

			</tr>

		</table>

	</td>
</tr>
<tr height="5">
	<td align="right">
		<input type="checkbox" name="ckHand" onClick="chgHandTax(this)">세액 수기입력
	</td>
</tr>
</form>
<tr>
	<td>

		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
		    <tr align="center">
				<td align="center" height="25">
				  <input type="button" class="button" value="작성" onClick="doRegisterSheet()">
				  &nbsp;
				  <input type="button" class="button" value="목록" onClick="self.location='Tax_list.asp'">
				</td>
			</tr>
		</table>

	</td>
</tr>
</table>

<p>

<iframe src="" name="icheckframe" width="0" height="0" frameborder="1" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>

<script>

// 페이지 시작시 작동하는 스크립트
function getOnload(){

	if (frm.billdiv.value != "0") {
		setRegisterInfo();
		CalcPriceWithPrice();
	}

	if (errMSG != "") {
		alert(errMSG);
	}
}

window.onload = getOnload;

</script>


<!-- 세금계산 요청서 정보 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
