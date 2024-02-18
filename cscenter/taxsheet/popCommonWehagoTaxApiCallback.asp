<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 공용세금계산서 발행 위하고 api 연동
' History : 2022.10.31 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderutf8.asp"-->
<!-- #include virtual="/lib/util/htmllib_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/common/jungsan/wehagoApiFunction.asp" -->
<%
dim taxIdx  : taxIdx  = requestCheckVar(getNumeric(request("taxIdx")),10)

if taxIdx="" then
    response.write "세금계산서번호가 없습니다."
    session.codePage = 949
    dbget.close() : response.end
end if

dim oTax, repEmail
set oTax = new CTax
    oTax.FRecttaxIdx = taxIdx
    oTax.GetTaxRead

dim sell_hp, sell_hp1, sell_hp2, sell_hp3
dim buy_hp, buy_hp1, buy_hp2, buy_hp3

sell_hp = Split(oTax.FOneItem.FsupplyRepTel, "-")
buy_hp = Split(oTax.FOneItem.FrepTel, "-")

if (UBound(sell_hp) >= 0) then
	sell_hp1 = sell_hp(0)
end if

if (UBound(sell_hp) >= 1) then
	sell_hp2 = sell_hp(1)
end if

if (UBound(sell_hp) >= 2) then
	sell_hp3 = sell_hp(2)
end if

if (UBound(buy_hp) >= 0) then
	buy_hp1 = buy_hp(0)
end if

if (UBound(buy_hp) >= 1) then
	buy_hp2 = buy_hp(1)
end if

if (UBound(buy_hp) >= 2) then
	buy_hp3 = buy_hp(2)
end if
repEmail = db2html(oTax.FOneItem.FrepEmail)

IF application("Svr_Info")="Dev" THEN
    sell_hp1 = "010"
    sell_hp2 = "9177"
    sell_hp3 = "8708"
    buy_hp1 = "010"
    buy_hp2 = "9177"
    buy_hp3 = "8708"
    repEmail = "tozzinet@10x10.co.kr"
end if

if (oTax.FOneItem.Fbilldiv = "52") or (oTax.FOneItem.Fbilldiv = "55") then
	response.write "텐바이텐 이외 사업자 발행불가"
    session.codePage = 949
    dbget.close() : response.end
end if

dim reg_socno
dim reg_subsocno
dim reg_socname
dim reg_ceoname
dim reg_socaddr
dim reg_socstatus
dim reg_socevent
dim reg_managername
dim reg_managerphone
dim reg_managermail

dim tplcompanyid, tplcompanyname, tplgroupid, tplbillUserID, tplbillUserPass
dim busiNo
reg_socno			= oTax.FOneItem.FsupplyBusiNo
reg_subsocno		= oTax.FOneItem.FsupplyBusiSubNo
reg_socname			= oTax.FOneItem.FsupplyBusiName
reg_ceoname			= oTax.FOneItem.FsupplyBusiCEOName
reg_socaddr			= oTax.FOneItem.FsupplyBusiAddr
reg_socstatus		= oTax.FOneItem.FsupplyBusiType
reg_socevent		= oTax.FOneItem.FsupplyBusiItem
reg_managername		= oTax.FOneItem.FsupplyRepName
reg_managerphone	= oTax.FOneItem.FsupplyRepTel
reg_managermail		= oTax.FOneItem.FsupplyRepEmail
busiNo = oTax.FOneItem.FbusiNo

dim FG_VAT : FG_VAT = "1"			'// 1과세, 3면세, 2영세(잘못된것 아님 : 빌365)

if IsNull(oTax.FOneItem.Ftaxtype) then
	oTax.FOneItem.Ftaxtype = ""
end if

'// Y : 과세 / N : 면세 / 0 : 영세
Select Case oTax.FOneItem.Ftaxtype
	Case "Y"
		FG_VAT = "1"
	Case "N"
		FG_VAT = "3"
	Case "0"
		FG_VAT = "2"
	Case Else
		response.write "과세구분 설정 에러"
        session.codePage = 949
        dbget.close() : response.end
End Select

dim isueDate

if IsNull(oTax.FOneItem.FisueDate) then
	oTax.FOneItem.FisueDate = ""
end if

if (oTax.FOneItem.FisueDate = "") then
	response.write "발행일자 설정 에러"
    session.codePage = 949
    dbget.close() : response.end
else
	isueDate = oTax.FOneItem.FisueDate
end if

dim ipkumdate : ipkumdate = ""

if IsNull(oTax.FOneItem.Fipkumdate) then
	oTax.FOneItem.Fipkumdate = ""
end if

'// 고객 주문의 경우 입금일자
ipkumdate = oTax.FOneItem.Fipkumdate

dim consignYN

if IsNull(oTax.FOneItem.FconsignYN) then
	oTax.FOneItem.FconsignYN = ""
end if

if (oTax.FOneItem.FconsignYN = "") then
	response.write "위수탁구분 설정 에러"
    session.codePage = 949
    dbget.close() : response.end
else
	consignYN = oTax.FOneItem.FconsignYN
end if

if (oTax.FOneItem.Fbilldiv = "99") then
	Call Get3PLUpcheInfoByTPLCompanyid(oTax.FOneItem.Ftplcompanyid, tplcompanyname, tplgroupid, tplbillUserID, tplbillUserPass)
end if

IF application("Svr_Info")="Dev" THEN
    reg_socno = "2222222227"
    busiNo = "1111111119"
    reg_managerphone	= "01091778708"
    reg_managermail		= "tozzinet@10x10.co.kr"

    isueDate = date()
end if

%>
<script type="text/javascript" src="https://static.wehago.com/support/jquery/jquery-3.2.1.min.js"></script>
<script type="text/javascript" src="https://static.wehago.com/support/jquery/jquery.cookie.js"></script>
<script type="text/javascript" src="https://static.wehago.com/support/wehago.0.1.2.js" charset="utf-8"></script>
<script type="text/javascript" src="https://static.wehago.com/support/wehagoLogin-1.1.6.min.js" charset="utf-8"></script>
<script type="text/javascript" src="https://static.wehago.com/support/service/common/wehagoCommon-0.2.8.min.js" charset="utf-8"></script>
<script type="text/javascript" src="https://static.wehago.com/support/service/invoice/wehagoInvoice-0.0.5.min.js" charset="utf-8"></script>
<script type="text/javascript" src="/js/gibberish-aesUTF8.js"></script>
<script type="text/javascript">
    <!-- #include virtual="/common/jungsan/wehago_globals_js.asp"-->

var NO_SENDER_PK='';
var last_NO_SENDER_PK='';

//발행후 저장
function saveTaxEvalResult(result,no_tax,result_msg,no_iss){
    var frm = document.taxSaveFrm;

    frm.action="/cscenter/taxsheet/saveTaxResult_utf8.asp";
    frm.result.value = result;
    frm.no_tax.value = no_tax;
    frm.result_msg.value = result_msg;
    frm.no_iss.value = no_iss;
	frm.target = "ipreSave";
	frm.submit();
}

function getAutoSendTax() {
    // 입력값 체크

    // cno조회
    getCompanyList();

    // 토큰조회
    setTimeout(function(){
        getServiceToken();
    },500);
}

// cno조회
function getCompanyList() {
    wehago_common.get_company_list_uncert(function (response) {
        for (var i=0; i<=response.resultData.length-1; i++) {
            var companyName = response.resultData[i].company_name_kr;
            var companyNo = response.resultData[i].company_no;
            $("#cno").val(companyNo);
            $("#disp_cno").text(companyNo);
        }
    }, callbackErrorcno);
}

// 토큰조회
function getServiceToken() {
    // 소속회사 리스트 조회한 cno값 입력 후, 토큰값 발급
    var service_hash_key = $.cookie(wehago_id_login.service_code + "_token");
    if($("#cno").val() == null || $("#cno").val() == undefined || $("#cno").val() == "") {
        alert("위하고쪽 실패 메세지4 [회사번호를 입력하세요]");
        return;
    }
    if (service_hash_key == undefined || $.cookie(wehago_id_login.service_code + "_selected_company_no") != $("#cno").val()) {
        wehago.getToken({
            "cno": $("#cno").val(),
            "thirdparty_a_token": wehago_id_login.getAccessToken()
        }, function (result) {
            if (result.resultCode == 200) {
                //alert("토큰발급 성공!");

                // 위하고 사용자 프로필 조회
                getWehagoUserProfile();
            } else {
                //alert("토큰발급 실패!");
                alert( "위하고쪽 실패 메세지5 [코드 : " + result.resultCode + "]" + result.resultMsg );
                return;
            }
        });
    } else {
        //alert("토큰발급 성공!");

        // 위하고 사용자 프로필 조회
        getWehagoUserProfile();
    }
}

// 위하고 사용자 프로필 조회
function getWehagoUserProfile() {
    wehago_id_login.get_wehago_userprofile("wehagoSignInCallback()");

    // 세금계산서 발행
    getInvoiceSendTax();
}

// 위하고 사용자 프로필 조회 이후 프로필 정보를 처리할 callback function
function wehagoSignInCallback() {
    $("#_wehago_id").text(wehago_id_login.getProfileData('wehago_id'));
    $("#_user_name").text(wehago_id_login.getProfileData('user_name'));
    $("#_user_contact").text(wehago_id_login.getProfileData('user_contact'));
    $("#_user_email").text(wehago_id_login.getProfileData('user_email'));
    $("#_user_no").text(wehago_id_login.getProfileData('user_no'));
}

// 세금계산서 발행
function getInvoiceSendTax() {
    //발행 상태
    var FG_FINAL = "1";
    var NO_SENDER_PK="SO_" + "<%= Trim(oTax.FOneItem.Forderserial) %>";

    // 위하고 서버와 통신 지연시 사용자가 버튼 두번 이상 클릭 방지
    if (NO_SENDER_PK==last_NO_SENDER_PK){
        alert('세금계산서발행키중복.');
        return;
    }

    last_NO_SENDER_PK = NO_SENDER_PK;
    var param = {
        "TB_TAX": {
            "DC_RMK": "",       // 비고1
            "DC_RMK2": "",      // 비고2
            "DC_RMK3": "",      // 비고3
            "FG_VAT": "<%= FG_VAT %>",      // ‘1’ : 과세, ‘2’ : 영세, ‘3’ : 면세
            "YN_TURN": "Y",     // 발행구분코드 ‘Y’:정발행, ‘N’:역발행
            "FG_IO": "1",       // 거래구분코드 , ‘1’ : 매출, ‘2’ : 매입
            "FG_PC": "1",       // 기업, 개인구분 ‘1’ : 기업, ‘2’ : 개인, ‘3’ : 외국인
            <% if (ipkumdate <> "") then %>
                "FG_BILL": "2",     // 청구유형 코드 ‘1’ : 청구, ‘2’ : 영수
            <% else %>
                "FG_BILL": "1",     // 청구유형 코드 ‘1’ : 청구, ‘2’ : 영수
            <% end if %>
            "YN_FX": "N",     // 수정 세금계산서 여부  Y:수정 세금 계산서, N: 정상 발행
            //"YN_DLV_ISS": "N",     // "Y" :지연교부 발행 , "N" : 지연교부 발행 않음 (과산세 부여되는 부분이므로 옵션 확인 후 발행할것)

            <%
            ' 1. 비고에 값이 있고 첫 두글자가 SO 로 되어 있으면 출고분계산서, 아니면 주문번호로를 PK 로 한다. (SO_주문번호, CUST_주문번호)
            ' 2. 비고에 값이 없고 orderidx 에 0 이 아닌 값이 있으면 가맹점계산서(FRAN_orderidx)
            ' 3. 비고에 값이 없고 orderidx 에 0 이면 추가발행계산서(TAX_taxIdx)
            %>
            <% if (Trim(oTax.FOneItem.Forderserial) <> "") and (Left(oTax.FOneItem.Forderserial, 2) = "SO") then %>
                // 출고코드
                "NO_SENDER_PK": NO_SENDER_PK,     // ERP내의 고유키 (PK). NO_SENDER_PK 로 세금계산서 조회 가능함.
            <% elseif (Trim(oTax.FOneItem.Forderserial) <> "") and (Left(oTax.FOneItem.Forderserial, 2) <> "SO") then %>
                <%
                dim osePK
                ''osePK = getOrderSerialPK(oTax.FOneItem.Forderserial)
                ''if (osePK="") then
                ''    response.write "alert('이미 발행 되었거나 올바른 주문번호가 아닙니다. - 관리자문의요망');return;"
                ''end if
                osePK = oTax.FOneItem.Forderserial & "_" & reg_socno
                %>
                // 주문번호
                "NO_SENDER_PK": "CUST_" + "<%= Trim(osePK) %>",     // ERP내의 고유키 (PK). NO_SENDER_PK 로 세금계산서 조회 가능함.
            <% else %>
                // 기타
                "NO_SENDER_PK": "TAX_" + "<%= Trim(CStr(oTax.FOneItem.FtaxIdx)) %>",     // ERP내의 고유키 (PK). NO_SENDER_PK 로 세금계산서 조회 가능함.
            <% end if %>

            "YN_CSMT": "<%= consignYN %>",      // 위수탁구분코드 ‘Y’:위수탁발행, ‘N’:정상발행

            // 공급자
            "SELL_NO_BIZ": "<%= Replace(reg_socno, "-", "") %>",    // 공급자 사업자등록번호
            "SELL_NM_CORP": "<%= reg_socname %>",     // 상호
            "SELL_NM_CEO": "<%= reg_ceoname %>",       // 대표자명
            "SELL_ADDR1": "<%= reg_socaddr %>",        // 주소
            "SELL_ADDR2": "",       // 상세주소
            //"SELL_DAM_DEPT": "",        // 담당자부서명
            "SELL_BIZ_STATUS": "<%= reg_socstatus %>",     // 업태
            "SELL_BIZ_TYPE": "<%= reg_socevent %>",      // 업종
            "SELL_DAM_NM": "<%= reg_managername %>",      // 공급자 담당자명
            "SELL_DAM_TEL1": "<%= sell_hp1 %>",      // 전화번호
            "SELL_DAM_TEL2": "<%= sell_hp2 %>",      // 전화번호
            "SELL_DAM_TEL3": "<%= sell_hp3 %>",      // 전화번호
            "SELL_DAM_MOBIL1": "<%= sell_hp1 %>",    // 휴대폰번호
            "SELL_DAM_MOBIL2": "<%= sell_hp2 %>",    // 휴대폰번호
            "SELL_DAM_MOBIL3": "<%= sell_hp3 %>",    // 휴대폰번호
            "BUY_DAM_TEL1": "<%= buy_hp1 %>",         // 전화번호
            "BUY_DAM_TEL2": "<%= buy_hp2 %>",         // 전화번호
            "BUY_DAM_TEL3": "<%= buy_hp3 %>",         // 전화번호
            "BUY_DAM_MOBIL1": "<%= buy_hp1 %>",     // 휴대폰번호
            "BUY_DAM_MOBIL2": "<%= buy_hp2 %>",    // 휴대폰번호
            "BUY_DAM_MOBIL3": "<%= buy_hp3 %>",   // 휴대폰번호
            "BUY_NO_BIZ": "<%= replace(replace(busiNo,"-","")," ","") %>",     // 공급받는자 사업자등록번호
            "BUY_NM_CORP": "<%= oTax.FOneItem.FbusiName %>",    // 상호
            "BUY_NM_CEO": "<%= oTax.FOneItem.FbusiCEOName %>",     // 공급받는자 대표자 성명
            "BUY_ADDR1": "<%= oTax.FOneItem.FbusiAddr %>",      // 공급받는자 주소
            "BUY_ADDR2": "",      // 공급받는자 상세주소
            "BUY_BIZ_STATUS": "<%= oTax.FOneItem.FbusiType %>",
            "BUY_BIZ_TYPE": "<%= oTax.FOneItem.FbusiItem %>",
            "BUY_DAM_DEPT": "",     // 담당자부서명
            "BUY_DAM_NM": "<%= db2html(oTax.FOneItem.FrepName) %>",       // 담당자명
            "SELL_REG_ID": "<%= reg_subsocno %>",      // 공급자 종사업장번호
            "BUY_REG_ID": "<%= Trim(CStr(NULL2Blank(oTax.FOneItem.FbusiSubNo))) %>",       // 공급받는자 종사업장번호
            "SELL_DAM_EMAIL": "<%= reg_managermail %>",     // 담당자이메일
            "BUY_DAM_EMAIL": "<%= repEmail %>",        // 담당자이메일
            "AMT": <%= oTax.FOneItem.FtotalPrice %>,      // 합계금액
            "AMT_CASH": <%= oTax.FOneItem.FtotalPrice %>,      // 현금
            "AMT_CHECK": 0,     // 수표
            "AMT_NOTE": 0,     // 어음
            "AMT_AR": <%= oTax.FOneItem.FtotalPrice %>,       // 외상미수금
            "AM": <%= oTax.FOneItem.FtotalPrice-oTax.FOneItem.FtotalTax %>,    // 공급가액
            "AM_VAT": <%= oTax.FOneItem.FtotalTax %>,      // 부가세액
            "YMD_WRITE": "<%= Replace(isueDate,"-","") %>",     // 작성일자 (YYYYMMDD)
            //"NO_ISSUE": "",       // 책번호 – 호
            //"NO_VOL": "",       // 책번호 – 권
            //"NO_SERIAL": pidx,      // 책번호 – 일련번호
            //"trade_connect_key": "",
            "CD_SVC": "100W",     // 과금창에서 전달된 서비스코드
            //"YN_LIQUOR": "",     // ‘Y’ : 품목란에 99개까지 보임. ‘N’ OR NULL : ‘품목이 4개이상 일 경우 ~외 처리’
            //"APP_NO_USER": "",     // 협력업체코드
            "YN_ISS": "0",     // 국세청신고여부 : FG_VAT:’3’(명세) 일경우 ‘0’으로 픽스 나머진 null
            //"YN_PAPER": "",     // 종이세금계산서 여부
            //"NO_TAX": "",     // 관리번호
            "SEND_RMK": "",     // 메일에 포함하는 메시지 내용
            "FG_TURN": "00",        // 과금창에서 전달된 값(과금구분:00.정발행,01.역발행-매입자과금,11.역발행-매출자과금)
            "FG_FINAL": FG_FINAL,    // 상태코드 : ‘1’ : 발행요청 FIX, 저장 : '0'
            "NM_SENDER_SYS": "WEHAGO"       // 시스템명(NM_SENDER_SYS)
        },
        "TB_TAX_LINE_LIST": [{
            "MM_WRITE": "<%= Mid(isueDate,6,2) %>",       // 작성월(MM) ‘05’
            "DD_WRITE": "<%= Mid(isueDate,9,2) %>",       // 작성일(DD) ‘25’
            "NM_ITEM": "<%= oTax.FOneItem.Fitemname %>",     // 품목명
            //"ITEM_STD": "",    // 규격
            "QTY": 1,       // 수량
            //"UM": ,     // 단가
            "AM": <%= oTax.FOneItem.FtotalPrice-oTax.FOneItem.FtotalTax %>,     // 공급가액
            "AM_VAT": <%= oTax.FOneItem.FtotalTax %>,      // 부가세액
            "AMT": <%= oTax.FOneItem.FtotalPrice %>     // 합계금액
        }],
        "TB_SERVICE": {
            "CD_SVC": "100W"        // 과금창에서 전달된 서비스코드
        },
        "YN_MGMT": "N",     // 자동책번호 사용여부
        "CERT_INFO": {},        // 인증서 정보( YN_CERT : N 일 경우 필수 ) / 솔루션 연동시 필수값 아니며, API에서 채워줍니다.
        "YN_CERT": "Y",     // 서버 인증서 사용여부 (Y : 사용, N : 미사용)
        "ccode": "biz201703300000011"       // 회사코드
    };
    wehago_invoice.get_invoice_sendtax(callbackSuccessSendTax, callbackErrorSendTax, param);
}

function callbackSuccessSendTax(response) {
    console.log("callback_success: ", response);

    var saveresult = "";
    var saveresult_msg = "";
    var saveno_tax = "";
    var saveno_iss = "";
    //alert( response.resultCode );
    //alert( response.resultMsg );

    if (response.resultData.RESULT=="00000"){
        //alert( response.resultData.TB_TAX.NO_TAX );
        //alert( response.resultData.TB_TAX.YN_ISS );
        //alert( response.resultData.TB_TAX.NO_ISS );
        //alert( response.resultData.TB_TAX.FG_FINAL );

        saveresult = response.resultData.RESULT;
        saveresult_msg  = response.resultData.RESULT_MSG;
        saveno_tax = response.resultData.TB_TAX.NO_TAX;
        saveno_iss = response.resultData.TB_TAX.NO_ISS;     //국세청승인번호

        //발행후 저장
        saveTaxEvalResult(saveresult,saveno_tax,saveresult_msg,saveno_iss);
        alert('발행 되었습니다.');
        setTimeout("opener.location.reload(); window.close();",2000)
    }else if (response.resultData.RESULT=="99991"){
        alert("이미 발행된 세금계산서 입니다.");
        return;
    }else{
        //alert( response.resultData.RESULT );
        alert( "위하고쪽 실패 메세지1 : " + response.resultData.RESULT_MSG );
        opener.location.reload(); window.close();
    }
}

function callbackErrorSendTax(response) {
    console.log("callback_error: ", response);
    alert( "위하고쪽 실패 메세지2 : " + response.resultMsg );
    opener.location.reload(); window.close();
}

function callbackErrorcno(response) {
    console.log("callback_error: ", response);
    alert( "위하고쪽 실패 메세지3 : " + response.resultMsg );
    opener.location.reload(); window.close();
}

function callbackSuccess(response) {
    console.log("callback_success: ", response);
}

function callbackError(response) {
    console.log("callback_error: ", response);
}

/*
// 세금계산서 발행취소
function getInvoiceSendTaxCancel() {
    var param = {
        "NO_TAXLIST": ["TX202001T036166"],
        "FG_IO": "1",
        "ccode": "biz201610170000001"
    };
    wehago_invoice.get_invoice_sendtaxcancel(callbackSuccess, callbackError, param);
}

// 세금계산서 리스트조회
function getInvoiceSearchTaxAccountlist() {
    var param = {
        "SC_TAXDATE_ST": "20210706",
        "SC_TAXDATE_ED": "20210806"
    }
    wehago_invoice.get_invoice_searchtaxaccountlist(callbackSuccess, callbackError, param);
}

// 세금계산서 단건조회
function getInvoiceSearchTaxAccount() {
    var param = {
        "NO_TAX": "TX202001T036167",
        "ccode": "biz201610170000001"
    };
    wehago_invoice.get_invoice_searchtaxaccount(callbackSuccess, callbackError, param);
}
*/

</script>

<div>
    <div style="width: 600px; margin: 0px auto 0; padding: 50px; border: 1px solid #9297a4;">
        <h1>위하고 로그인 정보</h1>
        <p>
            * wehago_id : <label id="_wehago_id"></label>
            <br />
            * access_token : <label id="access_token"></label>
            <br />
            * 회사 관리 시퀀스 : <label id="disp_cno"></label><input id="cno" name="cno" type="hidden">
            <br />
            * user_name : <label id="_user_name"></label>
            <br />
            * user_contact : <label id="_user_contact"></label>
            <br />
            * user_email : <label id="_user_email"></label>
            <br />
            * user_no : <label id="_user_no"></label>
        </p>
    </div>

    <form name="taxSaveFrm" method="post" style="margin:0px;" >
    <input type="hidden" name="taxIdx" value="<%= taxIdx %>">
    <input type="hidden" name="result" value="">
    <input type="hidden" name="no_tax" value="">
    <input type="hidden" name="billsiteCode" value="WE"> <!-- 더존B, 웹캐시W, 위하고WE -->
    <input type="hidden" name="result_msg" value="">
    <input type="hidden" name="no_iss" value="">
    <input type="hidden" name="write_date" value="<%= isueDate %>">
    </form>


    <% IF application("Svr_Info")="Dev" THEN %>
        <iframe name="ipreSave" id="ipreSave" width="100%" height="300"></iframe>
    <% else %>
        <iframe name="ipreSave" id="ipreSave" width="0" height="0"></iframe>
    <% end if %>
</div>

<script type="text/javascript">

var wehago_id_login = new wehago_id_login({
    app_key: "<%= wehagoAppKey %>",  // AppKey
    service_code: "<%= wehagoServiceCode %>",  // ServiceCode
    <% IF application("Svr_Info")="Dev" THEN %>
        //mode: "dev",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
        mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
    <% else %>
        mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
    <% end if %>
});

var wehago_common = new wehago_common({
    app_key: "<%= wehagoAppKey %>",  // AppKey
    service_code: "<%= wehagoServiceCode %>",  // ServiceCode
    <% IF application("Svr_Info")="Dev" THEN %>
        //mode: "dev",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
        mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
    <% else %>
        mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
    <% end if %>
    access_token : wehago_id_login.getAccessToken(),
});

var wehago_invoice = new wehago_invoice({
    app_key: "<%= wehagoAppKey %>",  // AppKey
    service_code: "<%= wehagoServiceCode %>",  // ServiceCode
    <% IF application("Svr_Info")="Dev" THEN %>
        //mode: "dev",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
        mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
    <% else %>
        mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
    <% end if %>
    access_token : wehago_id_login.getAccessToken(),
});

// 가져온 토크 표기
$("#access_token").text(wehago_id_login.getAccessToken());

// 필수값 체크후 회사코드 받아서 토큰 발급받아서 세금계산서 자동 발행
getAutoSendTax();

</script>

<%
function Get3PLUpcheInfoByTPLCompanyid(tplcompanyid, byRef tplcompanyname, byRef tplgroupid, byRef tplbillUserID, byRef tplbillUserPass)
	dim sqlStr

	sqlStr = " select top 1 t.tplcompanyid, t.tplcompanyname, t.groupid as tplgroupid, billUserID as tplbillUserID, billUserPass as tplbillUserPass "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner_tpl t with (nolock)"
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.tplcompanyid = '" + CStr(tplcompanyid) + "' "
	'' response.write sqlStr
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF then
		tplcompanyname = db2html(rsget("tplcompanyname"))
		tplgroupid = rsget("tplgroupid")
		tplbillUserID = rsget("tplbillUserID")
		tplbillUserPass = rsget("tplbillUserPass")
	end if
	rsget.close
end function

set oTax = Nothing

session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->