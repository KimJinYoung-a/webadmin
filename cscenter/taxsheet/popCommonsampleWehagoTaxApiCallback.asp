<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>위하고 API Callback (개발서버 cno=4로 되어있음)</title>
    <script type="text/javascript" src="https://static.wehago.com/support/jquery/jquery-3.2.1.min.js"></script>
    <script type="text/javascript" src="https://static.wehago.com/support/jquery/jquery.cookie.js"></script>
    <script type="text/javascript" src="https://static.wehago.com/support/wehago.0.1.2.min.js" charset="utf-8"></script>
    <script type="text/javascript" src="https://static.wehago.com/support/wehagoLogin-1.1.6.min.js" charset="utf-8"></script>
    <script type="text/javascript" src="https://static.wehago.com/support/service/common/wehagoCommon-0.2.8.min.js" charset="utf-8"></script>
    <script type="text/javascript" src="https://static.wehago.com/support/service/invoice/wehagoInvoice-0.0.5.min.js" charset="utf-8"></script>
    <script type="text/javascript" src="/cscenter/taxsheet/wehago_globals.js" charset="utf-8"></script>
</head>
<body>

<div style="width: 600px; margin: 100px auto 0; padding: 50px; border: 1px solid #9297a4;">
    <h1>위하고 API Callback</h1>
    <p>
        * access_token : <label id="access_token"></label>
        <br />
        * wehago_id : <label id="wehago_id"></label>
    </p>

    <button type="button" onclick="getCompanyList()">소속회사 리스트 조회</button>
    <br />
    <div id="company_list"></div>
    <br />
    회사 관리 시퀀스 : <input id="cno" name="cno" type="text">
    <button type="button" onclick="getServiceToken()">서비스토큰 발급</button>
    <br />
    <span >- 사용할 회사의 시퀀스를 입력하셔야 발급받은 서비스 토큰으로 API를 사용하실 수 있습니다.</span>
    <br />
    <br />
    <button type="button" onclick="getWehagoUserProfile()">위하고 사용자 프로필 조회</button>
    <br />
    <div style="padding: 5px;">* wehago_id : <label id="_wehago_id"></label></div>
    <div style="padding: 5px;">* user_name : <label id="_user_name"></label></div>
    <div style="padding: 5px;">* user_contact : <label id="_user_contact"></label></div>
    <div style="padding: 5px;">* user_email : <label id="_user_email"></label></div>
    <div style="padding: 5px;">* user_no : <label id="_user_no"></label></div>
    <br />
    <button type="button" onclick="getInvoiceSendTax()">세금계산서 발행</button>
    <button type="button" onclick="getInvoiceSendTaxCancel()">세금계산서 발행취소</button>
    <button type="button" onclick="getInvoiceSearchTaxAccountlist()">세금계산서 리스트조회</button>
    <button type="button" onclick="getInvoiceSearchTaxAccount()">세금계산서 단건조회</button>
    <br>
    <button type="button" onclick="getAutoSendTax()">세금계산서발행자동처리</button>
</div>

<script type="text/javascript">
    var wehago_id_login = new wehago_id_login({
        app_key: "98730f8cfdef4f77af17ce8ee08282fb",  // AppKey
        service_code: "10x10",  // ServiceCode
        mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
    });

    var wehago_common = new wehago_common({
        app_key: "98730f8cfdef4f77af17ce8ee08282fb",  // AppKey
        service_code: "10x10",  // ServiceCode
        mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
        access_token : wehago_id_login.getAccessToken(),
    });

    var wehago_invoice = new wehago_invoice({
        app_key: "98730f8cfdef4f77af17ce8ee08282fb",  // AppKey
        service_code: "10x10",  // ServiceCode
        mode: "live",  // dev-개발, live-운영 (기본값=live, 운영 반영시 생략 가능합니다.)
        access_token : wehago_id_login.getAccessToken(),
    });

    $("#access_token").text(wehago_id_login.getAccessToken());
    $("#wehago_id").text(wehago_id_login.getId());

    // 소속회사 리스트 조회
    function getCompanyList() {
        wehago_common.get_company_list_uncert(function (response) {
            for (var i=0; i<=response.resultData.length-1; i++) {
                var companyName = response.resultData[i].company_name_kr;
                var companyNo = response.resultData[i].company_no;
                var txt = document.createTextNode("company_name: " + companyName + ", cno: " + companyNo);
                var br = document.createElement("br");
                document.querySelector("#company_list").appendChild(txt);
                document.querySelector("#company_list").appendChild(br);
            }
        }, callbackError);
    }

    // 접근 토큰 값 및 위하고 아이디 출력
    function getServiceToken() {
        // 소속회사 리스트 조회한 cno값 입력 후, 토큰값 발급
        var service_hash_key = $.cookie(wehago_id_login.service_code + "_token");
        if($("#cno").val() == null || $("#cno").val() == undefined || $("#cno").val() == "") {
            alert("회사시퀀스를 입력하세요.");
            return;
        }
        if (service_hash_key == undefined || $.cookie(wehago_id_login.service_code + "_selected_company_no") != $("#cno").val()) {
            wehago.getToken({
                "cno": $("#cno").val(),
                "thirdparty_a_token": wehago_id_login.getAccessToken()
            }, function (result) {
                if (result.resultCode == 200) {
                    alert("토큰발급 성공!");
                } else {
                    alert("토큰발급 실패!");
                    // 필요시 실패 처리로직 구현
                }
            });
        } else {
            alert("토큰발급 성공!");
        }
    }

    // 위하고 사용자 프로필 조회
    function getWehagoUserProfile() {
        wehago_id_login.get_wehago_userprofile("wehagoSignInCallback()");
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
        var param = {
            "TB_TAX": {
                "DC_RMK": "",
                "DC_RMK2": "",
                "DC_RMK3": "",
                "FG_VAT": "1",
                "YN_TURN": "Y",
                "FG_IO": "1",
                "FG_PC": "1",
                "FG_BILL": "2",
                "SELL_NO_BIZ": "2222222227",
                "SELL_NM_CORP": "더존비즈온(본점)",
                "SELL_NM_CEO": "김대표",
                "SELL_ADDR1": "강원 춘천시 남산면 강촌구곡길 4",
                "SELL_ADDR2": "3434",
                "SELL_BIZ_STATUS": "농업, 임업 및 어업",
                "SELL_BIZ_TYPE": "곡물 및 기타 식량작물 재배업",
                "SELL_DAM_NM": "김우진",
                "SELL_DAM_TEL1": "0507",
                "SELL_DAM_TEL2": "6212",
                "SELL_DAM_TEL3": "0125",
                "SELL_DAM_MOBIL1": "010",
                "SELL_DAM_MOBIL2": "5143",
                "SELL_DAM_MOBIL3": "6051",
                "BUY_DAM_TEL1": "",
                "BUY_DAM_TEL2": "",
                "BUY_DAM_TEL3": "",
                "BUY_DAM_MOBIL1": "",
                "BUY_DAM_MOBIL2": "",
                "BUY_DAM_MOBIL3": "",
                "BUY_NO_BIZ": "1111111119",
                "BUY_NM_CORP": "상호명0001",
                "BUY_NM_CEO": "대표자이름",
                "BUY_ADDR1": "강원도 춘천시 남산면 강촌길17-5-33",
                "BUY_ADDR2": null,
                "BUY_BIZ_STATUS": "도소매",
                "BUY_BIZ_TYPE": "자동차",
                "BUY_DAM_DEPT": "",
                "BUY_DAM_NM": "",
                "SELL_REG_ID": "1234",
                "BUY_REG_ID": "1234",
                "SELL_DAM_EMAIL": "kwjdia@naver.com",
                "BUY_DAM_EMAIL": "test@duzon.com",
                "AMT": 1100,
                "AMT_CASH": 1100,
                "AMT_CHECK": 0,
                "AMT_NOTE": 0,
                "AMT_AR": 0,
                "AM": 1000,
                "AM_VAT": 100,
                "YMD_WRITE": "20220816",
                "NO_ISSUE": "09",
                "NO_VOL": "2019",
                "NO_SERIAL": "10",
                "trade_connect_key": "TRADE000001394229",
                "CD_SVC": "100W",
                "SEND_RMK": "",
                "FG_TURN": "00",
                "FG_FINAL": "1",
                "NM_SENDER_SYS": "WEHAGO"
            },
            "TB_TAX_LINE_LIST": [{
                "MM_WRITE": "09",
                "DD_WRITE": "26",
                "NM_ITEM": "품목1",
                "ITEM_STD": "1",
                "QTY": 1,
                "UM": 1000,
                "AM": 1000,
                "AM_VAT": 100,
                "AMT": 1100
            }],
            "TB_SERVICE": {
                "CD_SVC": "100W"
            },
            "YN_MGMT": "Y",
            "CERT_INFO": {},
            "YN_CERT": "Y",
            "ccode": "biz201610170000001"
        };
        wehago_invoice.get_invoice_sendtax(callbackSuccess, callbackError, param);
    }

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

    function callbackSuccess(response) {
        console.log("callback_success: ", response);
    }

    function callbackError(response) {
        console.log("callback_error: ", response);
    }

    // 세금계산서 발행
    function getInvoiceSendTax2() {
        var param = {
            "TB_TAX": {
                "DC_RMK": "",
                "DC_RMK2": "",
                "DC_RMK3": "",
                "FG_VAT": "1",
                "YN_TURN": "Y",
                "FG_IO": "1",
                "FG_PC": "1",
                "FG_BILL": "2",
                "SELL_NO_BIZ": "2222222227",
                "SELL_NM_CORP": "더존비즈온(본점)",
                "SELL_NM_CEO": "김대표",
                "SELL_ADDR1": "강원 춘천시 남산면 강촌구곡길 4",
                "SELL_ADDR2": "3434",
                "SELL_BIZ_STATUS": "농업, 임업 및 어업",
                "SELL_BIZ_TYPE": "곡물 및 기타 식량작물 재배업",
                "SELL_DAM_NM": "김우진",
                "SELL_DAM_TEL1": "0507",
                "SELL_DAM_TEL2": "6212",
                "SELL_DAM_TEL3": "0125",
                "SELL_DAM_MOBIL1": "010",
                "SELL_DAM_MOBIL2": "5143",
                "SELL_DAM_MOBIL3": "6051",
                "BUY_DAM_TEL1": "",
                "BUY_DAM_TEL2": "",
                "BUY_DAM_TEL3": "",
                "BUY_DAM_MOBIL1": "",
                "BUY_DAM_MOBIL2": "",
                "BUY_DAM_MOBIL3": "",
                "BUY_NO_BIZ": "1111111119",
                "BUY_NM_CORP": "상호명0001",
                "BUY_NM_CEO": "대표자이름",
                "BUY_ADDR1": "강원도 춘천시 남산면 강촌길17-5-33",
                "BUY_ADDR2": null,
                "BUY_BIZ_STATUS": "도소매",
                "BUY_BIZ_TYPE": "자동차",
                "BUY_DAM_DEPT": "",
                "BUY_DAM_NM": "",
                "SELL_REG_ID": "1234",
                "BUY_REG_ID": "1234",
                "SELL_DAM_EMAIL": "kwjdia@naver.com",
                "BUY_DAM_EMAIL": "test@duzon.com",
                "AMT": 1100,
                "AMT_CASH": 1100,
                "AMT_CHECK": 0,
                "AMT_NOTE": 0,
                "AMT_AR": 0,
                "AM": 1000,
                "AM_VAT": 100,
                "YMD_WRITE": "20220816",
                "NO_ISSUE": "09",
                "NO_VOL": "2019",
                "NO_SERIAL": "10",
                "trade_connect_key": "TRADE000001394229",
                "CD_SVC": "100W",
                "SEND_RMK": "",
                "FG_TURN": "00",
                "FG_FINAL": "1",
                "NM_SENDER_SYS": "WEHAGO"
            },
            "TB_TAX_LINE_LIST": [{
                "MM_WRITE": "09",
                "DD_WRITE": "26",
                "NM_ITEM": "품목1",
                "ITEM_STD": "1",
                "QTY": 1,
                "UM": 1000,
                "AM": 1000,
                "AM_VAT": 100,
                "AMT": 1100
            }],
            "TB_SERVICE": {
                "CD_SVC": "100W"
            },
            "YN_MGMT": "Y",
            "CERT_INFO": {},
            "YN_CERT": "Y",
            "ccode": "biz201610170000001"
        };
        wehago_invoice.get_invoice_sendtax(callbackSuccess2, callbackError, param);
    }

    function callbackSuccess2(response) {
        console.log("callback_success: ", response);
        //alert( response.resultCode );
        //alert( response.resultMsg );

        if (response.resultData.RESULT=="00000"){
            alert("세금계산서 발행 성공");
            //var resultDataArr = JSON.parse(response.resultData.TB_TAX);
            //alert( resultDataArr[0] );
            alert( response.resultData.TB_TAX.NO_TAX );
            alert( response.resultData.TB_TAX.YN_ISS );
            alert( response.resultData.TB_TAX.NO_ISS );
            alert( response.resultData.TB_TAX.FG_FINAL );
        }else if (response.resultData.RESULT=="99991"){
            alert("이미 발행된 세금계산서 입니다.");
            return;
        }else{
            alert( response.resultData.RESULT );
            alert( response.resultData.RESULT_MSG );
        }
    }

    // cno조회
    function getCompanyList2() {
        wehago_common.get_company_list_uncert(function (response) {
            for (var i=0; i<=response.resultData.length-1; i++) {
                var companyName = response.resultData[i].company_name_kr;
                var companyNo = response.resultData[i].company_no;
                $("#cno").val(companyNo);
                var txt = document.createTextNode("company_name: " + companyName + ", cno: " + companyNo);
                var br = document.createElement("br");
                document.querySelector("#company_list").appendChild(txt);
                document.querySelector("#company_list").appendChild(br);
            }
        }, callbackError);
    }

    // 토큰조회
    function getServiceToken2() {
        // 소속회사 리스트 조회한 cno값 입력 후, 토큰값 발급
        var service_hash_key = $.cookie(wehago_id_login.service_code + "_token");
        if($("#cno").val() == null || $("#cno").val() == undefined || $("#cno").val() == "") {
            alert("회사시퀀스를 입력하세요.");
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

                    // 세금계산서 발행
                    getInvoiceSendTax2();
                } else {
                    alert("토큰발급 실패!");
                    // 필요시 실패 처리로직 구현
                }
            });
        } else {
            //alert("토큰발급 성공!");

            // 위하고 사용자 프로필 조회
            getWehagoUserProfile();

            // 세금계산서 발행
            getInvoiceSendTax2();
        }
    }

    function getAutoSendTax() {
        // cno조회
        getCompanyList2();

        // 토큰조회
        setTimeout(function(){
            getServiceToken2();
        },500);
    }

//쿠키 값 가져오는 함수
function get_cookie(name) {
    var value = document.cookie.match('(^|;) ?' + name + '=([^;]*)(;|$)');
    return value? value[2] : null;
}
//쿠키 저장하는 함수
function set_cookie(name, value, unixTime) {
    var date = new Date();
    date.setTime(date.getTime() + unixTime);
    document.cookie = encodeURIComponent(name) + '=' + encodeURIComponent(value) + ';expires=' + date.toUTCString() + ';path=/';
}

//alert( get_cookie('pinfo') );
//set_cookie('access_token','4eIOQLbF0e2H45bBmKjgGZsykuGbUh','10')

</script>
<%
' 위하고측에서 쿠키유지를 8시간까지 하고 있다고 함
' 처음 위하고 접속시 토큰값을 저장하고 그 토큰으로 8시간 통신을 한다. 토큰이 없을때만 로그인해서 토큰 받아옴.
if session("WEHAGO_time") then
    session("WEHAGO_access_token") = request("access_token")
    session("WEHAGO_state") = request("state")
    session("WEHAGO_wehago_id") = request("wehago_id")
    session("WEHAGO_time") = now()
    Call fn_RDS_SSN_SET()
end if
%>
</body>
</html>