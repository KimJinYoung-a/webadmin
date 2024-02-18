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
' Description : 수수료 세금계산서 발행 위하고 api 연동
' History : 2022.11.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyheadutf8.asp"-->
<!-- #include virtual="/lib/util/htmllib_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls_utf8.asp"-->
<!-- #include virtual="/common/jungsan/wehagoApiFunction.asp" -->
<%
' 위하고측에서 쿠키유지를 8시간까지 하고 있다고 함
' 처음 위하고 접속시 토큰값을 저장하고 그 토큰으로 8시간 통신을 한다. 토큰이 없을때만 로그인해서 토큰 받아옴.
if session("WEHAGO_time")="" then
    session("WEHAGO_access_token") = request("access_token")
    session("WEHAGO_state") = request("state")
    session("WEHAGO_wehago_id") = request("wehago_id")
    session("WEHAGO_time") = now()
    Call fn_RDS_SSN_SET()
end if

dim i, repEmail, jungsan_name, isueDate, autotype
dim makerid, yyyy1,mm1, onoffGubun, jidx, isauto, nextjidx, FG_VAT
makerid 		= requestCheckvar(request("makerid"),32)
yyyy1   		= requestCheckvar(request("yyyy1"),10)
mm1     		= requestCheckvar(request("mm1"),10)
onoffGubun     	= requestCheckvar(request("onoffGubun"),10)
jidx            = requestCheckvar(request("jidx"),10)
isauto          = requestCheckvar(request("isauto"),10)
nextjidx        = requestCheckvar(request("nextjidx"),10)
autotype 		= requestCheckvar(request("autotype"),32)
dim groupid
groupid = getPartnerId2GroupID(makerid)

dim ojungsanTaxCC
set ojungsanTaxCC = new CUpcheJungsanTax
ojungsanTaxCC.FRectMakerid = makerid
ojungsanTaxCC.FRectTargetGbn = onoffGubun
ojungsanTaxCC.FRectJjungsanIdx = jidx
ojungsanTaxCC.getOneUpcheJungsanTax

dim PrdCommissionSum : PrdCommissionSum = 0

if (ojungsanTaxCC.FresultCount>0) then
	if (ojungsanTaxCC.FOneItem.IsCommissionTax) then
	    PrdCommissionSum = ojungsanTaxCC.FOneItem.Ftotalcommission
	end if

    FG_VAT = ojungsanTaxCC.FOneItem.getBill_FG_VAT
end if

IF application("Svr_Info")<>"Dev" THEN
    if PrdCommissionSum = 0 then
        if (request("autotype")="V2") then
        response.write "<script type='text/javascript'>"&vbCRLF
        response.write "opener.addResultLog('"&request("jidx")&"','수수료0');"&vbCRLF
        response.write "opener.fnNextEvalProc();"&vbCRLF
        response.write "</script>"
        else
        response.write "<script type='text/javascript'>alert('수수료 매출정보가 없습니다.');</script>"
        response.write "수수료 매출정보가 없습니다"
        end if
        session.codePage = 949
        dbget.close() : response.end
    end if
end if
if ojungsanTaxCC.FOneItem.IsEvaledTax then
    if (request("autotype")="V2") then
    response.write "<script type='text/javascript'>"&vbCRLF
    response.write "opener.addResultLog('"&request("jidx")&"','기정산확정');"&vbCRLF
    response.write "opener.fnNextEvalProc();"&vbCRLF
    response.write "</script>"
    else
    response.write "<script type='text/javascript'>alert('이미 정산 확정된 내역입니다.');</script>"
    response.write "이미 정산 확정된 내역입니다."
    end if
    session.codePage = 949
    dbget.close()	:	response.End
end if

dim opartner, ogroup
dim stypename

set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid
opartner.FPageSize = 1
opartner.GetOnePartnerNUser

set ogroup = new CPartnerGroup
ogroup.FRectGroupid = ojungsanTaxCC.FOneItem.Fgroupid
ogroup.GetOneGroupInfo

if ogroup.FResultCount<1 then
    if (request("autotype")="V2") then
        response.write "<script type='text/javascript'>"&vbCRLF
        response.write "opener.addResultLog('"&request("jidx")&"','그룹미지정/정산정보없음');"&vbCRLF
        response.write "opener.fnNextEvalProc();"&vbCRLF
        response.write "</script>"
    else
        response.write "<script type='text/javascript'>alert('그룹 코드가 지정되지 않았거나, 정산정보가 없습니다.');</script>"
        response.write "그룹 코드가 지정되지 않았거나, 정산정보가 없습니다"
    end if
    session.codePage = 949
	dbget.close()	:	response.End
end if

dim MaySocialNo : MaySocialNo=FALSE ''주민번호로 발급
if IsMaySocialNo(ogroup.FOneItem.Fcompany_no) then
    MaySocialNo = true
    ogroup.FOneItem.Fcompany_no = ogroup.FOneItem.FdecCompNo
end if

jungsan_name=ogroup.FOneItem.Fjungsan_name

if (NOT MaySocialNo) then
    if LEN(replace(ogroup.FOneItem.Fcompany_no,"-",""))<>10 then
        if (request("autotype")="V2") then
            response.write "<script type='text/javascript'>"&vbCRLF
            response.write "opener.addResultLog('"&request("jidx")&"','사업자번호');"&vbCRLF
            response.write "opener.fnNextEvalProc();"&vbCRLF
            response.write "</script>"
        else
            response.write "<script type='text/javascript'>alert('사업자 번호가 올바르지 않습니다.');</script>"
            response.write "사업자 번호가 올바르지 않습니다."& replace(ogroup.FOneItem.Fcompany_no,"-","") & "::" & LEN(replace(ogroup.FOneItem.Fcompany_no,"-",""))
        end if
        session.codePage = 949
    	dbget.close()	:	response.End
    end if
end if

stypename = "세금계산서"
dim jungsan_hpall, jungsan_hp1,jungsan_hp2,jungsan_hp3, reg_socno, busiNo, buyceoname, buycompany_address1, buycompany_address2
jungsan_hpall = Trim(ogroup.FOneItem.Fjungsan_hp)
jungsan_hpall = split(jungsan_hpall,"-")

if UBound(jungsan_hpall)>=0 then
	jungsan_hp1 = jungsan_hpall(0)
end if
if UBound(jungsan_hpall)>=1 then
	jungsan_hp2 = jungsan_hpall(1)
end if
if UBound(jungsan_hpall)>=2 then
	jungsan_hp3 = jungsan_hpall(2)
end if

if (jungsan_hp2="") and (jungsan_hp3="") and (Len(jungsan_hp1)=11) then
    jungsan_hp3 = MID(jungsan_hp1,8,4)
    jungsan_hp2 = MID(jungsan_hp1,4,4)
    jungsan_hp1 = LEFT(jungsan_hp1,3)
end if
repEmail = db2html(ogroup.FOneItem.Fjungsan_email)
reg_socno = "211-87-00620"
busiNo = ogroup.FOneItem.Fcompany_no
buyceoname=ogroup.FOneItem.Fceoname
buycompany_address1 = ogroup.FOneItem.Fcompany_address
buycompany_address2 = ogroup.FOneItem.Fcompany_address2

IF application("Svr_Info")="Dev" THEN
    reg_socno = "2222222227"
    busiNo = "1111111119"
    buyceoname = "한용민"
    buycompany_address1 = "서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층"
    buycompany_address2 = "서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 15층"
    jungsan_hp1 = "010"
    jungsan_hp2 = "9177"
    jungsan_hp3 = "8708"
    repEmail = "tozzinet@10x10.co.kr"
    jungsan_name = "한용민"

    isueDate = date()
end if

Dim EVAL_CompanyNo  : EVAL_CompanyNo = "2118700620"

IF application("Svr_Info")<>"Dev" THEN
    if (replace(ogroup.FOneItem.Fcompany_no,"-","")=EVAL_CompanyNo) then
        if (request("autotype")="V2") then
            response.write "<script type='text/javascript'>"&vbCRLF
            response.write "opener.addResultLog('"&request("jidx")&"','텐바이텐사업자발행불가');"&vbCRLF
            response.write "opener.fnNextEvalProc();"&vbCRLF
            response.write "</script>"
        else
            response.write "<script type='text/javascript'>alert('텐바이텐 사업자 발행 불가.');</script>"
            response.write "텐바이텐 사업자 발행 불가."
        end if
        session.codePage = 949
        dbget.close()	:	response.End
    end if
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

var fxStarted = false;
var NO_SENDER_PK='';
var last_NO_SENDER_PK='';
var pLogIdx = 0;

// 발행전 저장
function preSaveLog(){
    var frm = document.frm;
    <% if (jungsan_hp1="") or (jungsan_hp2="") or (jungsan_hp3="") or (Len(jungsan_hp1)>3) or (Len(jungsan_hp2)>4) or (Len(jungsan_hp3)>4) then %>
        <% if (request("autotype")="V2") then %>
            opener.addResultLog('<%=request("jidx")%>','<strong>휴대폰번호</strong>');
            opener.fnNextEvalProc();
        <% else %>
            alert('정산 담당자 핸드폰 번호가 올바르지 않습니다. \n업체정보수정에서 정산담당자 핸드폰을 000-000-0000 대시 형태로 수정후 사용하세요.');
        <% end if %>
        return;
    <% end if %>
    <% if repEmail="" or isnull(repEmail) then %>
        <% if (request("autotype")="V2") then %>
            opener.addResultLog('<%=request("jidx")%>','<strong>이메일주소</strong>');
            opener.fnNextEvalProc();
        <% else %>
            alert('정산 이메일주소가 올바르지 않습니다. \n업체정보수정에서 정산담당자 이메일주소 수정후 사용하세요.');
        <% end if %>
        return;
    <% end if %>
    <% if jungsan_name="" or isnull(jungsan_name) then %>
        <% if (request("autotype")="V2") then %>
            opener.addResultLog('<%=request("jidx")%>','<strong>정산담당자명</strong>');
            opener.fnNextEvalProc();
        <% else %>
            alert('정산 담당자명이 올바르지 않습니다. \n업체정보수정에서 정산 담당자명 수정후 사용하세요.');
        <% end if %>
        return;
    <% end if %>

    if (fxStarted){
        <% if (request("autotype")="V2") then %>
            fxStarted = false;
            opener.addResultLog('<%=request("jidx")%>','<strong>세금계산서 중복발행요청</strong>');
            opener.fnNextEvalProc();
        <% else %>
            alert('세금계산서 중복발행요청.');
        <% end if %>
        return;
    }
    fxStarted = true;

    frm.action="/admin/upchejungsan/dotaxregAdm_utf8.asp";
	frm.target = "ipreSave";
	frm.submit();
}

//발행후 저장
function saveTaxEvalResult(result,no_tax,result_msg,no_iss){
    var frm = taxSaveFrm;
    frm.action="/admin/upchejungsan/saveTaxResultAdm_utf8.asp";
    frm.idx.value = pLogIdx;
    frm.result.value = result;
    frm.no_tax.value = no_tax;
    frm.no_iss.value = no_iss;
    frm.result_msg.value = result_msg;
	frm.target = "ipreSave";
	frm.submit();

	fxStarted = false;
}

function getAutoSendTax() {
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
        <% if (request("autotype")="V2") then %>
            opener.addResultLog('<%=request("jidx")%>','<strong>위하고쪽 실패 메세지4 [회사번호를 입력하세요]</strong>');
            opener.fnNextEvalProc();
        <% else %>
            alert("위하고쪽 실패 메세지4 [회사번호를 입력하세요]");
        <% end if %>
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
                <% if (request("autotype")="V2") then %>
                    opener.addResultLog('<%=request("jidx")%>','<strong>위하고쪽 실패 메세지5 [코드 : ' + result.resultCode + ']' + result.resultMsg + '</strong>');
                    opener.fnNextEvalProc();
                <% else %>
                    alert( "위하고쪽 실패 메세지5 [코드 : " + result.resultCode + "]" + result.resultMsg );
                <% end if %>
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

    // 발행전 저장
    preSaveLog();
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
function getInvoiceSendTax(pidx) {
    pLogIdx = pidx;
    var FG_FINAL = "1"; 
    var NO_SENDER_PK="<%= ojungsanTaxCC.FOneItem.getBill_NO_SENDER_PK %>";

    // 위하고 서버와 통신 지연시 사용자가 버튼 두번 이상 클릭 방지
    if (NO_SENDER_PK==last_NO_SENDER_PK){
        <% if (request("autotype")="V2") then %>
            opener.addResultLog('<%=request("jidx")%>','<strong>세금계산서발행키중복</strong>');
            opener.fnNextEvalProc();
        <% else %>
            alert('세금계산서발행키중복.');
        <% end if %>
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
            <% if (MaySocialNo) then %>
                "FG_PC": "2",       // 기업, 개인구분 ‘1’ : 기업, ‘2’ : 개인, ‘3’ : 외국인  // 2016/09/29 핑거스 작품 개인
            <% else %>
                "FG_PC": "1",       // 기업, 개인구분 ‘1’ : 기업, ‘2’ : 개인, ‘3’ : 외국인
            <% end if %>
            "FG_BILL": "<%= ojungsanTaxCC.FOneItem.getBill_FG_BILL %>",     // 청구유형 코드 ‘1’ : 청구, ‘2’ : 영수
            "YN_FX": "N",     // 수정 세금계산서 여부  Y:수정 세금 계산서, N: 정상 발행
            //"YN_DLV_ISS": "N",     // "Y" :지연교부 발행 , "N" : 지연교부 발행 않음 (과산세 부여되는 부분이므로 옵션 확인 후 발행할것)
            "NO_SENDER_PK": NO_SENDER_PK,     // ERP내의 고유키 (PK). NO_SENDER_PK 로 세금계산서 조회 가능함.
            "YN_CSMT": "N",      // 위수탁구분코드 ‘Y’:위수탁발행, ‘N’:정상발행

            // 공급자
            "SELL_NO_BIZ": "<%= Replace(reg_socno, "-", "") %>",    // 공급자 사업자등록번호
            "SELL_NM_CORP": "(주)텐바이텐",     // 상호
            "SELL_NM_CEO": "최은희",       // 대표자명
            "SELL_ADDR1": "서울시 종로구 대학로 57",        // 주소
            "SELL_ADDR2": "홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐",       // 상세주소
            "SELL_DAM_DEPT": "재경팀",        // 담당자부서명
            "SELL_BIZ_STATUS": "도소매외",     // 업태
            "SELL_BIZ_TYPE": "전자상거래외",      // 업종
            "SELL_DAM_NM": "계산서담당자",      // 공급자 담당자명
            "SELL_DAM_TEL1": "02",      // 전화번호
            "SELL_DAM_TEL2": "554",      // 전화번호
            "SELL_DAM_TEL3": "2033",      // 전화번호
            "SELL_DAM_MOBIL1": "02",    // 휴대폰번호
            "SELL_DAM_MOBIL2": "554",    // 휴대폰번호
            "SELL_DAM_MOBIL3": "2033",    // 휴대폰번호
            "BUY_DAM_TEL1": "<%= jungsan_hp1 %>",         // 전화번호
            "BUY_DAM_TEL2": "<%= jungsan_hp2 %>",         // 전화번호
            "BUY_DAM_TEL3": "<%= jungsan_hp3 %>",         // 전화번호
            "BUY_DAM_MOBIL1": "<%= jungsan_hp1 %>",     // 휴대폰번호
            "BUY_DAM_MOBIL2": "<%= jungsan_hp2 %>",    // 휴대폰번호
            "BUY_DAM_MOBIL3": "<%= jungsan_hp3 %>",   // 휴대폰번호
            "BUY_NO_BIZ": "<%= replace(replace(busiNo,"-","")," ","") %>",     // 공급받는자 사업자등록번호
            "BUY_NM_CORP": "<%= ogroup.FOneItem.FCompany_name %>",    // 상호
            "BUY_NM_CEO": "<%= buyceoname %>",     // 공급받는자 대표자 성명
            "BUY_ADDR1": "<%= buycompany_address1 %>",      // 공급받는자 주소
            "BUY_ADDR2": "<%= buycompany_address2 %>",      // 공급받는자 상세주소
            "BUY_BIZ_STATUS": "<%= ogroup.FOneItem.Fcompany_uptae %>",
            "BUY_BIZ_TYPE": "<%= ogroup.FOneItem.Fcompany_upjong %>",
            "BUY_DAM_DEPT": "",     // 담당자부서명
            "BUY_DAM_NM": "<%= jungsan_name %>",       // 담당자명
            //"SELL_REG_ID": "",      // 공급자 종사업장번호
            //"BUY_REG_ID": "",       // 공급받는자 종사업장번호
            "SELL_DAM_EMAIL": "accounts@10x10.co.kr",     // 담당자이메일
            "BUY_DAM_EMAIL": "<%= repEmail %>",        // 담당자이메일
            "AMT": <%= ojungsanTaxCC.FOneItem.getJungsanTaxSum %>,      // 합계금액
            "AMT_CASH": 0,      // 현금
            "AMT_CHECK": 0,     // 수표
            "AMT_NOTE": 0,     // 어음
            "AMT_AR": 0,       // 외상미수금
            "AM": <%= ojungsanTaxCC.FOneItem.getJungsanTaxSuply %>,    // 공급가액
            "AM_VAT": <%= ojungsanTaxCC.FOneItem.getJungsanTaxVat %>,      // 부가세액
            "YMD_WRITE": "<%= Replace(ojungsanTaxCC.FOneItem.GetPreFixSegumil,"-","") %>",     // 작성일자 (YYYYMMDD)
            //"NO_ISSUE": "",       // 책번호 – 호
            //"NO_VOL": "",       // 책번호 – 권
            "NO_SERIAL": pidx,      // 책번호 – 일련번호
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
            "MM_WRITE": "<%= Mid(ojungsanTaxCC.FOneItem.GetPreFixSegumil,6,2) %>",       // 작성월(MM) ‘05’
            "DD_WRITE": "<%= Mid(ojungsanTaxCC.FOneItem.GetPreFixSegumil,9,2) %>",       // 작성일(DD) ‘25’
            "NM_ITEM": "<%= ojungsanTaxCC.FOneItem.getBill_NM_ITEM %>",     // 품목명
            "ITEM_STD": "<%= Right(Replace(ojungsanTaxCC.FOneItem.Fyyyymm,"-",""),4) %>",    // 규격
            "QTY": 1,       // 수량
            //"UM": <%= ojungsanTaxCC.FOneItem.getJungsanTaxSuply %>,     // 단가
            "AM": <%= ojungsanTaxCC.FOneItem.getJungsanTaxSuply %>,     // 공급가액
            "AM_VAT": <%= ojungsanTaxCC.FOneItem.getJungsanTaxVat %>,      // 부가세액
            "AMT": <%= ojungsanTaxCC.FOneItem.getJungsanTaxSum %>     // 합계금액
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
    }else if (response.resultData.RESULT=="99991"){
        <% if (request("autotype")="V2") then %>
            opener.addResultLog('<%=request("jidx")%>','<strong>이미 발행된 세금계산서 입니다.</strong>');
            opener.fnNextEvalProc();
        <% else %>
            alert("이미 발행된 세금계산서 입니다.");
            opener.location.reload(); window.close();
        <% end if %>
        return;
    }else{
        //alert( response.resultData.RESULT );
        <% if (request("autotype")="V2") then %>
            opener.addResultLog('<%=request("jidx")%>','<strong>위하고쪽 실패 메세지1 ' + response.resultData.RESULT_MSG + '</strong>');
            opener.fnNextEvalProc();
        <% else %>
            alert( "위하고쪽 실패 메세지1 : " + response.resultData.RESULT_MSG );
            opener.location.reload(); window.close();
        <% end if %>
        return;
    }
}

function callbackErrorSendTax(response) {
    console.log("callback_error: ", response);
    <% if (request("autotype")="V2") then %>
        opener.addResultLog('<%=request("jidx")%>','<strong>위하고쪽 실패 메세지2 ' + response.resultMsg + '</strong>');
        opener.fnNextEvalProc();
    <% else %>
        alert( "위하고쪽 실패 메세지2 : " + response.resultMsg );
        opener.location.reload(); window.close();
    <% end if %>
    return;
}

function callbackErrorcno(response) {
    console.log("callback_error: ", response);
    <% if (request("autotype")="V2") then %>
        opener.addResultLog('<%=request("jidx")%>','<strong>위하고쪽 실패 메세지3 ' + response.resultMsg + '</strong>');
        opener.fnNextEvalProc();
    <% else %>
        alert( "위하고쪽 실패 메세지3 : " + response.resultMsg );
        opener.location.reload(); window.close();
    <% end if %>
    return;
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

<form name="frm" method="post" action="/admin/upchejungsan/dotaxregAdm_utf8.asp" style="margin:0px;" >
<input type="hidden" name="jungsanid" value="<%= ojungsanTaxCC.FOneItem.FId %>">
<input type="hidden" name="jungsanname" value="<%= ojungsanTaxCC.FOneItem.Ftitle %>">
<input type="hidden" name="jungsangubun" value="<%= ojungsanTaxCC.FOneItem.FtargetGbn %>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="jgubun" value="<%= ojungsanTaxCC.FOneItem.Fjgubun %>">
<input type="hidden" name="biz_no" value="<%= replace(replace(socialnoReplace(ogroup.FOneItem.Fcompany_no),"-","")," ","") %>" >
<input type="hidden" name="corp_nm" value="<%= ogroup.FOneItem.FCompany_name %>">
<input type="hidden" name="ceo_nm" value="<%= ogroup.FOneItem.Fceoname %>">
<input type="hidden" name="biz_status" value="<%= ogroup.FOneItem.Fcompany_uptae %>">
<input type="hidden" name="biz_type" value="<%= ogroup.FOneItem.Fcompany_upjong %>">
<input type="hidden" name="addr" value="<%= ogroup.FOneItem.Fcompany_address %> <%= ogroup.FOneItem.Fcompany_address2 %>">
<input type="hidden" name="dam_nm" value="<%= jungsan_name %>">
<input type="hidden" name="email" value="<%= ogroup.FOneItem.Fjungsan_email %>">
<input type="hidden" name="hp_no1" value="<%= jungsan_hp1 %>">
<input type="hidden" name="hp_no2" value="<%= jungsan_hp2 %>">
<input type="hidden" name="hp_no3" value="<%= jungsan_hp3 %>">
<input type="hidden" name="sb_type" value="01"> <!-- 매출 01 매입 02 -->
<input type="hidden" name="tax_type" value="<%= ojungsanTaxCC.FOneItem.Ftaxtype %>">
<input type="hidden" name="bill_type" value="01"> <!-- 영수 01 청구 18 -->
<input type="hidden" name="pc_gbn" value="C"> <!-- 개인 P 기업 C -->
<input type="hidden" name="item_count" value="1">
<input type="hidden" name="item_nm" value="<%= ojungsanTaxCC.FOneItem.getBill_NM_ITEM %>" size=25>
<input type="hidden" name="item_qty" value="1">
<input type="hidden" name="item_price" value="<%= ojungsanTaxCC.FOneItem.getJungsanTaxSum %>">
<input type="hidden" name="item_amt" value="<%= ojungsanTaxCC.FOneItem.getJungsanTaxSuply %>">
<input type="hidden" name="item_vat" value="<%= ojungsanTaxCC.FOneItem.getJungsanTaxVat %>">
<input type="hidden" name="item_remark" value="">
<input type="hidden" name="credit_amt" value="<%= ojungsanTaxCC.FOneItem.getJungsanTaxSum %>">
<input type="hidden" name="cur_u_user_no" value="261744"> <!-- DEV 1000394, REAL 244730, ON 261744 -->
<input type="hidden" name="cur_dam_nm" value="재경팀">
<input type="hidden" name="cur_email" value="accounts@10x10.co.kr">
<input type="hidden" name="cur_hp_no1" value="02">
<input type="hidden" name="cur_hp_no2" value="554">
<input type="hidden" name="cur_hp_no3" value="2033">
<input type="hidden" name="autotype" value="<%=request("autotype")%>">
<input type="hidden" name="billSite" value="WE">
</form>
<form name="taxSaveFrm" method="post" style="margin:0px;" >
<input type="hidden" name="idx" value="">
<input type="hidden" name="result" value="">
<input type="hidden" name="no_tax" value="">
<input type="hidden" name="no_iss" value="">
<input type="hidden" name="billsiteCode" value="WE"> <!-- 더존B, 웹캐시W, 위하고WE -->
<input type="hidden" name="result_msg" value="">
<input type="hidden" name="jungsangubun" value="<%= ojungsanTaxCC.FOneItem.FtargetGbn %>">
<input type="hidden" name="write_date" value="<%= ojungsanTaxCC.FOneItem.GetPreFixSegumil %>">
<input type="hidden" name="jungsanid" value="<%= ojungsanTaxCC.FOneItem.FId %>">
<input type="hidden" name="isauto" value="<%= isauto %>">
</form>
<% IF application("Svr_Info")="Dev" THEN %>
    <iframe name="ipreSave" id="ipreSave" width="100%" height="300"></iframe>
<% else %>
    <iframe name="ipreSave" id="ipreSave" width="100%" height="50"></iframe>
<% end if %>

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
set ojungsanTaxCC = Nothing
set opartner = Nothing
set ogroup = Nothing
%>
<script type="text/javascript">

function reActEval(){
    <% if (nextjidx<>"") then %>
        <% if (jidx<>nextjidx) then %>
        opener.evalOneTax_V2(<%=nextjidx%>)
        <% end if %>
    <% elseif (request("autotype")="V2") then %>
        opener.addResultLog("<%=jidx%>","v")
        opener.fnNextEvalProc()
    <% end if %>
}

</script>
<%
function IsMaySocialNo(icompanyno)
    IsMaySocialNo = false
    if isNULL(icompanyno) then Exit function
    IsMaySocialNo = LEN(trim(replace(icompanyno,"-","")))=13
end function

session.codePage = 949
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->