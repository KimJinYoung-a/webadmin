<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 세금계산서 발행 빌36524 api 연동
' History : 2021.01.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderutf8.asp"-->
<!-- #include virtual="/lib/util/htmllib_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
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

'// ============================================================================
if (oTax.FOneItem.Fbilldiv = "52") or (oTax.FOneItem.Fbilldiv = "55") then
	response.write "텐바이텐 이외 사업자 발행불가"
    session.codePage = 949
    dbget.close() : response.end
end if


'// ============================================================================
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
IF application("Svr_Info")="Dev" THEN
    reg_socno = "2222222227"
    busiNo = "1111111119"
    reg_managerphone	= "01091778708"
    reg_managermail		= "tozzinet@10x10.co.kr"
end if

'// ============================================================================
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


'// ============================================================================
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

'// ============================================================================
dim ipkumdate : ipkumdate = ""

if IsNull(oTax.FOneItem.Fipkumdate) then
	oTax.FOneItem.Fipkumdate = ""
end if

'// 고객 주문의 경우 입금일자
ipkumdate = oTax.FOneItem.Fipkumdate


'// ============================================================================
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


'// ============================================================================
if (oTax.FOneItem.Fbilldiv = "99") then
	Call Get3PLUpcheInfoByTPLCompanyid(oTax.FOneItem.Ftplcompanyid, tplcompanyname, tplgroupid, tplbillUserID, tplbillUserPass)
end if

dim USER_IP, loginUrl
    USER_IP = "110.93.128.93"

    ' Bill36524 웹서비스용 WEBAPI 연동모듈
    IF application("Svr_Info")="Dev" THEN
        loginUrl = "https://realtest.bill36524.com:1443/action.dox"
        isueDate = date()
    else
        loginUrl = "https://www.bill36524.com:443/action.dox"
    end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" src="/js/bil/AC_OETags.js"></script>
<script language="javascript" src="/js/bil/TSToolkitConfig.js"></script>
<script language="javascript" src="/js/bil/TSToolkitObject.js"></script>
<script language="javascript" src="/js/bil/TSToolkitUtil.js"></script>

<script type="text/javascript" src="/js/jquery.blockUI.js"></script>
<script type="text/javascript" src="/js/page.js"></script>

<%	
' 제공된 샘플에 있으나 안쓰는듯?
'	<object id="TSToolkit"
'		classid="clsid:55D9860A-AB9C-44A1-BB74-75AF7F805333"
'		codebase="http://webadmin.10x10.co.kr/common/cab/TSToolkit.cab#version=2,0,7,0"
'		style="LEFT: 0px; TOP: 0px" width="0" height="0" viewastext=""></object>
%>
<script type="text/javascript">

var fxStarted = false;

// 리턴 객체 저장용 전역 변수
var g_recvArr = new Array();
var g_sendArr = new Array();
var g_searArr = new Array();

// 파일첨부 호출 후 받아지는 값 저장
var filename = "";
var filepath = "";

var allcnt=300000;

var APIObject = "dzWebApiMgr"; 

// 동시에 여러건 발행 가능 여부 설정 0:_CallCount 체크 함 (여러건 발생 불가능)1:_CallCount 체크 안함(여러건 발생 가능)
var _SendTaxMuch;

// 서버 연결 지연시 사용자가 세금계산서 발행 버튼을 두번 클릭  방지용
// 발행 가능한 상태 (동시에 여러건 발행 못함  0: 발행 가능, 1: 방행 진행중, 2: 발행 성공. 3: 발행 오류)
var _CallCount; 

// 서버 인증서 사용 확인 여부
var _EnableCertification = 0;
var isLoadCert = false;
var CERT_VALUE = "";
var CERT_KEY = "";
var CERT_RVALUE = "";
var str_YN_TURN;

//-------------------------- 로그인 -------------------------------------------------------------------
function FxLogin(iid,ipwd){
    if (fxStarted) return;
    fxStarted = true;

    var ID = iid;
    var PASSWD = ipwd;
    
    //접속  ip 입력 부탁드립니다. 
    var USER_IP = "<%= USER_IP %>";

    var jsonObj = {
            service : "dzWebApiMgr",
            method : "fxLogin",
            ID : ID,	
            USER_IP : USER_IP,
            PASSWD : PASSWD,
        };
    
    var data = {};
    data['json'] = JSON.stringify(jsonObj);

    $.ajax({
        url : "<%= loginUrl %>",
        type : "POST",
        data : data,
        dataType : "json",
        crossDomain:true,
        async : true,
        success	: function(data, textStatus) {			
            FxLoginResult(data.hashtable, textStatus);
            $.unblockUI();
        },
        error : function(xhr, textStatus, errorThrown) {
            alert("조회중 오류가 발생했습니다.");
            $.unblockUI();
        },
        complete : function(xhr, textStatus) {
            //alert('complete');
        },
        beforeSend : function(xhr) {
            fnBlockUI();
        }
    });	
        
}

var loginSess;

//로그인 결과     
function FxLoginResult(data, textStatus){
    var retObj = data;

    //성공
    if(retObj.RESULT == "00000"){
        loginSess = retObj;
        
        form1.output.value =  " [Login]\n" ;
        form1.output.value += " NO_CUST : " + retObj.NO_CUST+"\n USER_ID : "+retObj.USER_ID+"\n";
        form1.output.value += " FG_CUST : " + retObj.FG_CUST+"\n NM_CUST : "+retObj.NM_CUST+"\n";
        form1.output.value += " NM_USER : " + retObj.NM_USER+"\n NO_ID : "+retObj.NO_ID+"\n";
        form1.output.value += " NO_USER : " + retObj.NO_USER+"\n";
        form1.output.value += " TEL1 : " + retObj.TEL1+"\n";
        form1.output.value += " TEL2 : " + retObj.TEL2+"\n";
        form1.output.value += " TEL3 : " + retObj.TEL3+"\n";
        form1.output.value += " MOBILE1 : " + retObj.MOBILE1+"\n";
        form1.output.value += " MOBILE2 : " + retObj.MOBILE2+"\n";
        form1.output.value += " MOBILE3 : " + retObj.MOBILE3+"\n";
        form1.output.value += " YMD_FG_PAY : " + retObj.YMD_FG_PAY+"\n";
        form1.output.value += " NM_CEO : " + retObj.NM_CEO+"\n";
        form1.output.value += " ADDR1 : " + retObj.ADDR1+"\n";
        form1.output.value += " ADDR2 : " + retObj.ADDR2+"\n";
        form1.output.value += " NO_ID : " + retObj.NO_ID+"\n";
        form1.output.value += " NM_CUST : " + retObj.NM_CUST+"\n";
        form1.output.value += " EMAIL : " + retObj.EMAIL+"\n";
        form1.output.value += " BIZ_STATUS : " + retObj.BIZ_STATUS+"\n";
        form1.output.value += " BIZ_TYPE : " + retObj.BIZ_TYPE+"\n";
        form1.output.value += " YN_ISS_VAT3 : " + retObj.YN_ISS_VAT3+"\n";      
        form1.output.value += " POINT : " + retObj.POINT+"\n";    
                    
        form1.output.value += " NM_DEPT : " + retObj.NM_DEPT+"\n";
        form1.output.value += " YN_PW_UPDATE : " + retObj.YN_PW_UPDATE+"\n";

		//document.getElementById("loginForm").style.display = "none";
		document.getElementById("loginView").style.display = "";

        var loginfo = "<strong>" + retObj.NM_CUST + " | " + retObj.NM_USER + "</strong>";//업체명, 고객명		
        document.getElementById("logInfo").innerHTML = loginfo;
        var point = "<strong> 현재포인트 : " + retObj.POINT + "</strong>";			
        document.getElementById("curPoint").innerHTML = point;

        FxSendTaxAccount()
    //실패
    }else{
        alert(retObj.RESULT_MSG);
    }
}
//-------------------------- 로그인 -------------------------------------------------------------------

//-------------------------- 세금계산서 발행 -------------------------------------------------------------------
function FxSendTaxAccount(){
    if (loginSess != undefined) {
        if (loginSess.NO_CUST == null || loginSess.NO_CUST == "" || loginSess.NO_CUST == undefined
            || loginSess.NO_USER == null || loginSess.NO_USER == "" || loginSess.NO_USER == undefined
            || loginSess.ID == null || loginSess.ID == "" || loginSess.ID == undefined 
            || loginSess.NO_ID == null || loginSess.NO_ID == "" || loginSess.NO_ID == undefined) 
        {
            alert("로그인 정보가 없습니다.");
            return;
        }
    } else {
        alert("로그인 정보가 없습니다.");
        return;
    } 
   
    var arrObject = new Object();
    
    arrObject.AM = "<%= oTax.FOneItem.FtotalPrice-oTax.FOneItem.FtotalTax %>";      // 공급가액
    arrObject.AM_VAT = "<%= oTax.FOneItem.FtotalTax %>";        // 부가세액
    arrObject.AMT = "<%= oTax.FOneItem.FtotalPrice %>";         // 합계금액

    <% if (oTax.FOneItem.Fbilldiv = "01" or oTax.FOneItem.Fbilldiv = "11") then %>
        arrObject.AMT_CASH = "<%= oTax.FOneItem.FtotalPrice %>";			// 현금
    <% else %>
        arrObject.AMT_AR = "<%= oTax.FOneItem.FtotalPrice %>";			// 외상미수금
    <% end if %>

    arrObject.AMT_CHECK = "0";
    arrObject.AMT_NOTE = "0";

	<% if (ipkumdate <> "") then %>
        arrObject.FG_BILL = "2";    //청구유형 코드 (1 : 청구, 2 : 영수)
	<% else %>
        arrObject.FG_BILL = "1";    //청구유형 코드 (1 : 청구, 2 : 영수)
	<% end if %>

    arrObject.YN_FX = "N";      // 수정 세금계산서 여부  Y:수정 세금 계산서, N: 정상 발행 <== 필수 입력 입니다
    arrObject.YN_TURN = "Y"     //Y정발행 N역발행  :: 역발행시 발행요청 , 정발행시 승인요청
    arrObject.FG_IO = "1";      //1매출 2매입
    //arrObject.FG_TURN = getRadioValue(form1.FG_TURN);      //역발행시는 무조건 1. 개인이 역발행하지 않음
    arrObject.FG_PC = "1";      //1기업 2개인
    arrObject.FG_VAT = "<%= FG_VAT %>";     // 1과세, 3면세, 2영세(잘못된 것 아님)

    //발행 상태
    var FG_FINAL = "1"; 
    arrObject.FG_FINAL = FG_FINAL;      //0저장 1 발송 2승인 3반려 4승인취소요청
    
    arrObject.YN_CSMT = "<%= consignYN %>";
    //arrObject.YN_DLV_ISS = getRadioValue(form1.YN_DLV_ISS);   // 지연교부

    <%
    ' 1. 비고에 값이 있고 첫 두글자가 SO 로 되어 있으면 출고분계산서, 아니면 주문번호로를 PK 로 한다. (SO_주문번호, CUST_주문번호)
    ' 2. 비고에 값이 없고 orderidx 에 0 이 아닌 값이 있으면 가맹점계산서(FRAN_orderidx)
    ' 3. 비고에 값이 없고 orderidx 에 0 이면 추가발행계산서(TAX_taxIdx)
    %>
    <% if (Trim(oTax.FOneItem.Forderserial) <> "") and (Left(oTax.FOneItem.Forderserial, 2) = "SO") then %>
        // 출고코드
        arrObject.NO_SENDER_PK = "SO_" + "<%= Trim(oTax.FOneItem.Forderserial) %>";
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
        arrObject.NO_SENDER_PK = "CUST_" + "<%= Trim(osePK) %>";
    <% else %>
        // 기타
        arrObject.NO_SENDER_PK = "TAX_" + "<%= Trim(CStr(oTax.FOneItem.FtaxIdx)) %>";
    <% end if %>

    //arrObject.DC_RMK = form1.DC_RMK.value;    // 비고
    //arrObject.CD_SVC = form1.CD_SVC.value;    // 과금코드
    //arrObject.YN_LIQUOR = "N";
    //arrObject.APP_NO_USER = form1.APP_NO_USER.value;    //협력업체코드   
    arrObject.YMD_WRITE = "<%= Replace(isueDate,"-","") %>";    // 작성일자

    // 공급자
    arrObject.SELL_NO_BIZ = "<%= Replace(reg_socno, "-", "") %>";       // 등록번호
    arrObject.SELL_NM_CORP = "<%= reg_socname %>";       // 상호
    arrObject.SELL_NM_CEO = "<%= reg_ceoname %>";     // 대표자명
    arrObject.SELL_ADDR1 = "<%= reg_socaddr %>";    	 // 주소 
    arrObject.SELL_ADDR2 = "";       // 상세주소
    arrObject.SELL_DAM_DEPT = "";       // 담당자부서명
    arrObject.SELL_DAM_NM = "<%= reg_managername %>";       // 담당자명
    arrObject.SELL_DAM_EMAIL = "<%= reg_managermail %>";    // 담당자이메일
    arrObject.SELL_DAM_MOBIL1 = "<%= sell_hp1 %>";      // 휴대폰번호
    arrObject.SELL_DAM_MOBIL2 = "<%= sell_hp2 %>";
    arrObject.SELL_DAM_MOBIL3 = "<%= sell_hp3 %>";
    arrObject.SELL_DAM_TEL1 = "<%= sell_hp1 %>";        // 전화번호
    arrObject.SELL_DAM_TEL2 = "<%= sell_hp2 %>";
    arrObject.SELL_DAM_TEL3 = "<%= sell_hp3 %>";
    arrObject.SELL_BIZ_STATUS = "<%= reg_socstatus %>";     // 업태
    arrObject.SELL_BIZ_TYPE = "<%= reg_socevent %>";        // 업종
    arrObject.SELL_REG_ID = "<%= reg_subsocno %>";      // 종사업장번호
    
    // 공급받는자
    arrObject.BUY_NO_BIZ = "<%= replace(replace(busiNo,"-","")," ","") %>";       // 등록번호
    arrObject.BUY_NM_CORP = "<%= oTax.FOneItem.FbusiName %>";       // 상호
    arrObject.BUY_NM_CEO = "<%= oTax.FOneItem.FbusiCEOName %>";     // 대표자명
    arrObject.BUY_ADDR1 = "<%= oTax.FOneItem.FbusiAddr %>";    	 // 주소 
    arrObject.BUY_ADDR2 = "";       // 상세주소
    arrObject.BUY_DAM_DEPT = "";       // 담당자부서명
    arrObject.BUY_DAM_NM = "<%= db2html(oTax.FOneItem.FrepName) %>";       // 담당자명
    arrObject.BUY_DAM_EMAIL = "<%= repEmail %>";    // 담당자이메일
    arrObject.BUY_DAM_MOBIL1 = "<%= buy_hp1 %>";      // 휴대폰번호
    arrObject.BUY_DAM_MOBIL2 = "<%= buy_hp2 %>";
    arrObject.BUY_DAM_MOBIL3 = "<%= buy_hp3 %>";
    arrObject.BUY_DAM_TEL1 = "<%= buy_hp1 %>";        // 전화번호
    arrObject.BUY_DAM_TEL2 = "<%= buy_hp2 %>";
    arrObject.BUY_DAM_TEL3 = "<%= buy_hp3 %>";
    arrObject.BUY_BIZ_STATUS = "<%= oTax.FOneItem.FbusiType %>";
    arrObject.BUY_BIZ_TYPE = "<%= oTax.FOneItem.FbusiItem %>";
    
    arrObject.BUY_REG_ID = "<%= Trim(CStr(NULL2Blank(oTax.FOneItem.FbusiSubNo))) %>";
    //arrObject.NO_VOL = form1.NO_VOL.value;    // 책번호(권)
    //arrObject.NO_SERIAL = form1.NO_SERIAL.value;      // 일련번호
    //arrObject.NO_ISSUE = form1.NO_ISSUE.value;        // 책번호(권)
    //arrObject.NM_SENDER_SYS = form1.NM_SENDER_SYS.value;      // 시스템명(NM_SENDER_SYS)
    
    arrObject.YN_ISS = "0";     //FG_VAT 가 3(면세) 일경우 YN_ISS : NULL 일경우 전송제외 YN_ISS : 0 일경우 국세청 전송요청
    //arrObject.YN_PAPER = form1.NO_VOL.value;      // 종이세금계산서 여부
    //arrObject.NO_TAX = form1.NO_TAX_SEND.value;     // 관리번호(NO_TAX)
  
    //수정발행시 Y 아닐경우 N
    var arr = new Array(); 		
    arr.push(arrObject);          
    
    var msg = "";	          
    // 품목 정보 
    var arrlineArray = new Array();
    var cnt = 1;
    var arrLineObject = new Object();
    
    for(var i=0; i<cnt; i++)
    {
        //console.log("i ", i);
        //arrLineObject.ITEM_STD = form1.TL_ITEM_STD.value + "_" + i;       // 규격
        arrLineObject.NM_ITEM = "<%= oTax.FOneItem.Fitemname %>";       // 품목명
        arrLineObject.AM = "<%= oTax.FOneItem.FtotalPrice-oTax.FOneItem.FtotalTax %>";      // 공급가액
        arrLineObject.AM_VAT = "<%= oTax.FOneItem.FtotalTax %>";        // 부가세액
        arrLineObject.AMT = "<%= oTax.FOneItem.FtotalPrice %>";     // 합계금액
        //arrLineObject.UM = form1.TL_UM.value;     // 단가
        arrLineObject.DD_WRITE = "<%= Mid(isueDate,9,2) %>";
        arrLineObject.MM_WRITE = "<%= Mid(isueDate,6,2) %>";
        arrLineObject.QTY = cnt;       // 수량

    var arrline1 = new Array();    
    arrline1.push(arrLineObject); 

    arrlineArray[i]= arrline1;
    }  

    //------------------------------------------------------------------------------
    // 파일 첨부
    // 파일 첨부는  API 파일 첨부 발행.ppt 파일을 참고 해주시기 바랍니다.
    var arrFile = new Array();
    var arrFileObject = new Object();

    var FILE_NAME = filename;	// 파일 이름
    var FILE_PATH = filepath;	// 파일 경로

    if(	filename == null || filepath == null){
        arrFile = null;
    }else{
        arrFileObject.FILE_NAME = FILE_NAME;
        arrFileObject.FILE_PATH = FILE_PATH;	
        
        arrFile.push(arrFileObject); 
    }
        
    //arrFile = null;   //임시
    //------------------------------------------------------------------------------  

    //------------------------------------------------------------------------------
    // 수정 세금 계산서 arr 항목에서 YN_FX 필드가 Y일 때 사용 합니다.
          var arrTbFx = new Array();

		/* CD_FX 코드가 다음과 같을때 현재 소스 코드에서 위쪽에 arr 배열에서 DC_RMK에 해당 사항 입력
		01 : 기재 사항 착오, 정정
			필요적 기재사항 등이 착오로 잘못 기재된경우
		 	세무서장이 결정하여 통지하기 전까지 세금계산서 작성 필요
		 	당초 발급분에 대한 부(-)와 착오사항을 정정한 정(+)의 세금계산서를 각 1장씩 발행
		
		02 :공급가액 변동
			공급가액에 추가 또는 차감되는 금액이 발생한경우
			증감사유가 발생한 날을 작성일자로 기재, 비고란에 당초 세금계산서 교부일자 기재
			증감 금액에 따라 당초 발급분에 대한 부(-) 또는 정(+)의 세금계산서 1장을 발행
			
		03 : 환입
			당초 공급한 재화가 환입된 경우
			재화가 환입된 날을 작성일자로 기재하며, 비고란에는 당초 세금계산서 교부일자 기재
			참고 ) FG_VAT 가 2(영세) 일때 AM_VAT (부가세)금액이 0보다 작거나 같아야 함
			환입된 금액만큼 부(-)의 세금계산서를 1장 발행
		
		04 : 계약의 해제
			계약의 해제로 인하여 재화 또는 용역이 공급되지 아니한경우
			당초 세금계산서 작성일을 작성일자에 기재 하고 비고란에 계약 해제 일자를 기재
			금액 변경은 반영 안된다
			참고 ) FG_VAT 가 2(영세) 일때 AM_VAT (부가세)금액이 0보다 작거나 같아야 함
			당초 발급분에 대한 부(-)의 세금계산서를 1장 발행
		
		05 : 내국신용장 사후 개설
			공급시기가 속하는 과세기간 종료 후 20일 이내에 내국신용장 등이 개설된 경우
			작성일자를 당초 세금계산서 작성일자로 기재
			당초 발급분에 대한 부(-)의 세금계산서와 함께 열세율로 작성된 정(+)의 세금계산서를
			각 1장씩 발행

		06 : 착오에 의한 이중발급
		           단순 착오로 인한 계산서를 이중으로 발급한 경우,
		           당초 발급분에 대한 부(-)의 계산서를 1장 발행함
		*/

		var cdfx = "";
		var notxsrc = "";
		var notxrel = "";
		var noisssrc = "";
		
		var arrTbFxObject = new Object();

        /*
		if(getRadioValue(form1.YN_FX) == "Y"){
			
			cdfx = getValue(form1.CD_FX);
			notxsrc = getValue(form1.NO_TX_SRC);
			notxrel = getValue(form1.NO_TX_REL);
			noisssrc = getValue(form1.NO_ISS_SRC);

			if(cdfx == null || cdfx == ""){
				alert("사유코드를 선택하세요.");
				return;
			}

			if(noisssrc == null || noisssrc == ""){
				alert("원천 국세청전송번호를 입력하세요.");
				return;
			}
		}
		*/

  		var CD_FX = cdfx;					//사유코드 
  		var NO_TX_SRC = notxsrc;			//원천 세금계산서 관리번호
  		var NO_TX_REL = notxrel;			//관련세금계산서 (-) (+) 2장발행할경우 (-) 세금계산서의 관리번호 
		var NO_ISS_SRC = noisssrc;			//원천 세금계산서 국세청 전송번호 (NO_ISS) //필수
  		
		arrTbFxObject.CD_FX = CD_FX; 
		arrTbFxObject.NO_TX_SRC = NO_TX_SRC; 
		arrTbFxObject.NO_TX_REL = NO_TX_REL;
		arrTbFxObject.NO_ISS_SRC = NO_ISS_SRC;
		
		arrTbFx.push(arrTbFxObject);
		
    //------------------------------------------------------------------------------
    // 서버인증서 사용 : 0 , 로컬인증서 사용 : 1
    //thisMovie(APIObject).EnableCertification(getRadioValue(form1.CERT));
    //getRadioValue(form1.CERT);

    //console.log("arrlineArray" , arrlineArray);
    if(arrlineArray.length <= 0)
    {
        alert("품목은 최소 1개 이상 있어야 합니다.");
        return;
    }

    // 페이지 새로고침 없이 여러건 발행 가능 여부 설정 
    // 0: 체크 함 (여러건 발행 불가능)1: 체크 안함(여러건 발행 가능)
    // 무조건 1로 세팅 
    //thisMovie(APIObject).SendTaxMuch(1); 
    _SendTaxMuch = '1';

        // bill36524 서버와 통신 지연시 상요자가 버튼 두번 이상 클릭 방지
    if(!SendStateCheck()){
        return;   			
    }

    jsonObj = {
            service : "dzWebApiMgr",
            method : "FxSendTaxAccount",        // SaveTaxAccount, FxSendTaxAccount, FxSaveTaxAccount
            arrTax : arr,
            arrLine : arrlineArray,
            arrTbFx : arrTbFx,      // 수정 세금 계산서 arr 항목에서 YN_FX 필드가 Y일 때 사용 합니다.
            arrFile : arrFile,      // 파일첨부
            LOGIN_DATA: loginSess,
            CMD_USR_ID : "",
        };
    
    var data = {};
    data['json'] = JSON.stringify(jsonObj);

    $.ajax({
        url : "<%= loginUrl %>",
        type : "POST",
        data : data,
        dataType : "json",
        crossDomain:true,
        async : true,
        success	: function(data, textStatus) {			
            FxSendTaxAccountResult(data.hashtable, textStatus);
            $.unblockUI();
        },
        error : function(xhr, textStatus, errorThrown) {
            alert("조회중 오류가 발생했습니다.");
            $.unblockUI();
        },
        complete : function(xhr, textStatus) {
                
        },
        beforeSend : function(xhr) {
            fnBlockUI();
        }
    });	

}

//세금계산서 발행 결과 품목이 5건 미만일때
function FxSendTaxAccountResult(data, textStatus){
    var retObj = data;
    //console.log("+++++++ " ,retObj);
    g_recvArr.push(retObj);
    //console.log("g_recvArr" + g_recvArr)
    RecvProcess();
}

function RecvProcess(){
    //배열에 값이 있으면 처리
    if(g_recvArr.length > 0){    
        //전역 배열로 부터 한건을 꺼내와 다음 처리를 시작한다.
        var result = g_recvArr.pop();
        var tb_tax = result.OBJ_TBTAX;
        
        if(result.RESULT == "00000"){ // 발행 성공
            alert("세금계산서 발행 성공");
            form1.output.value += " [SendTaxAccount]\n";
            form1.output.value += " tb_tax.NO_TAX : " + tb_tax.NO_TAX + "\n" + " tb_tax.NO_SENDER_PK : " + tb_tax.NO_SENDER_PK + "\n";
            form1.output.value += " tb_tax.NO_ISS : " + tb_tax.NO_ISS + "\n" + " tb_tax.YN_ISS : " + tb_tax.YN_ISS + "\n";
            form1.output.value += " tb_tax.FG_FINAL : " + tb_tax.FG_FINAL +"\n"; 
            // FxSendTaxAccountAll()함수를 호출하여 세금계산서 대량 발행 일때 g_sendArr.length 는 1보다 많습니다.
        
            //scrollEnd();

            var saveresult = result.RESULT;
            var saveresult_msg  = result.RESULT_MSG;
            if (tb_tax!=null){
                var saveno_tax = tb_tax.NO_TAX;
                var saveno_iss = tb_tax.NO_ISS;     //국세청승인번호
            }else{
                var saveno_tax = "";
            }

            saveTaxEvalResult(saveresult,saveno_tax,saveresult_msg,saveno_iss);
            setTimeout("opener.location.reload(); window.close();",2000)
        } else {    // 발행 실패
            _CallCount = 0;
            _oldNOSENDPK = "";
            alert("발행실패 - result.RESULT_MSG:" + result.RESULT_MSG);        
        }
    }
}

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
	fxStarted = false;
}

function SendStateCheck() {
    if(_SendTaxMuch==0){
        // 인증서 중복 발행 체크
        // 서버 연결 지연시 사용자가 세금계산서 발행 버튼을 두번 클릭 하면
        // 중복 발생 가능성 이 있어 두번 이상의 버튼 클릭 방지용도  0: 발행 가능, 1: 방행 진행중, 2: 발행 성공. 3: 발행 오류
        switch(_CallCount){
            case 1:
            {
                alert("세금계산서 발행 진행 중입니다. 잠시만 기다려 주세요!	");
                return false;
                break;
            }
            case 2:
            {
                alert("이미 세금계산서 발행을 완료 하였습니다.");
                return false;
                break;
            }
            case 3:
            {
                alert("세금계산서 발행중 오류가 발생 하였습니다.");
                return false;
                break;
            }
        }
    }
    return true;   		// 입력 받은 데이터  length 체크
    if(!DataLengthCheck(argsTbTax,arrTbTaxLine,arrTbFx))
        return;
}

// 데이터 필드 length 체크 및 유효성 체크
function DataLengthCheck(argsTbTax,arrTbTaxLine,arrTbFx){
    var newNOSENDPK;
    
    var objDataLengh = GetDataLength();
    
    var i = 0;
    
    for(i = 0; i<argsTbTax.length; i++){
        // argsTbTax Array  데이터 length 체크   
        if(String(argsTbTax[i]["value"]).length > objDataLengh[argsTbTax[i]["key"]])
        {
            _CallCount = 0;
            alert(argsTbTax[i]["key"] + " Over Length Error");
            return false;
        }
        
        //페이지 새로 고침 없이 바로 전dml  no_send_pk 번호가 중복 되는 지  검사 하기 위해 임시 보관 
        if(argsTbTax[i]["key"] == "NO_SENDER_PK")
        {
            newNOSENDPK = argsTbTax[i]["value"];
        }
        
        
        // 정발행 역발행 구분 
        if(argsTbTax[i]["key"] == "YN_TURN")
        {
            if(argsTbTax[i]["value"] == "N")
            {
                str_YN_TURN = false;
            }
            else if(argsTbTax[i]["value"] == "Y")
            {
                str_YN_TURN = true;
            }
            else
            {
                _CallCount = 0;
                alert(arrTbTaxLine[i]["key"] + " UnValid");
                return false;	
            }
        }
    }
    
    // arrTbTaxLine Array 요소 데이터 length 체크
    for(i = 0; i<arrTbTaxLine.length; i++)
    {   
        if(String(arrTbTaxLine[i]["value"]).length > objDataLengh[arrTbTaxLine[i]["key"]])
        {
            alert(arrTbTaxLine[i]["key"] + " Over Length Error");
            return false;
        }
    }
    
    //페이지 새로 고침 없이 바로 전의  no_send_pk 번호가 중복 되는 지  검사 
    if(newNOSENDPK == _oldNOSENDPK && _oldNOSENDPK.length > 0)
    {
        _CallCount = 0;
        
        alert("NO_SENDER_PK 값이 중복 되었습니다.");
        return false;
    }
    
    // 수정 세금계산서 필수 항목 체크
    if(arrTbFx !=null)
    {
        for(i = 0; i<arrTbFx.length; i++)
        {   
            if(String(arrTbFx[i]["value"]).length > objDataLengh[arrTbFx[i]["key"]])
            {
                alert(arrTbFx[i]["key"] + " Over Length Error");
                return false;
            }
        }
    }
    
    _oldNOSENDPK = newNOSENDPK;
    
    return true;
}
//-------------------------- 세금계산서 발행 -------------------------------------------------------------------

function evalTx(){
    if (confirm('발행일 : <%= isueDate %>\n발행 하시겠습니까?')){
        
		<%
        IF application("Svr_Info")="Dev" THEN
		    response.write "FxLogin('BILLTEST02','bizon#720');"
        else
            Select Case oTax.FOneItem.Fbilldiv
                Case "01"
                    '// 고객 - 공급자 텐바이텐
                    response.write "FxLogin('customer','20011010');"
                Case "11"
                    '// 고객 - 공급자(업체별)
                    response.write "FxLogin('customer','20011010');"
                Case "02"
                    '// 가맹점 - 공급자 텐바이텐
                    response.write "FxLogin('accounts','20011010');"
                Case "03"
                    '// 프로모션 - 공급자 텐바이텐
                    response.write "FxLogin('promotion','20011010');"
                Case "51"
                    '// 기타 - 공급자 텐바이텐
                    response.write "FxLogin('accounts','20011010');"
                Case "99"
                    '// 3PL업체
                    response.write "FxLogin('" & tplbillUserID & "','" & tplbillUserPass & "');"
                Case Else
                    response.write "FxLogin('customer','20011010');"
            End Select
        end if
		%>
    }
}

function fnBlockUI() {
    var msg = "<div style='text-align: center;'>";
    msg += "<p style='margin:8px;font-size:14px;font-family: dotum;font-weight: bold;color:#999999;'>처리중 입니다.</p></div>";
    $.blockUI({ message: msg,
        overlayCSS:{ 
            backgroundColor: '#000000', 
            opacity: 0.01
        },
        css:{
            backgroundColor: "#EFEFEF",
            width: '180px',
            left: '42%', 
            border: '#629CD8 solid 2px',
            '-webkit-border-radius': '10px', 
            '-moz-border-radius':    '10px'
        },
        fadeIn:  0,
        fadeOut: 0
    });
}	

function getOnLoad(){
    setTimeout("evalTx();",1000)
}

</script>

<form id="form1" name="form1" onsubmit="return false;" style="margin:0px;" >
<div style="" id="loginView">
    <div class="tb_terms">
        <table>
            <tbody>
                <tr>
                    <th colspan="2" style="font-size: 15px; text-align: left">로그인정보</th>
                </tr>
                <tr>
                    <td><div id="logInfo"></div></td>

                </tr>
                <tr>
                    <td>
                        <div id="curPoint"></div> 
                    </td>
                </tr>
            </tbody>
        </table>
        <br>
        결과 (Bill36524.com) :<br>
        <textarea id="output" rows="40" cols="80"></textarea>
        <br>
    </div>
</div>

<input type="button" value="발행" onclick="evalTx();"> 오래 기다려도 자동발행이 안될경우 눌러주세요.
</form>
<form name="taxSaveFrm" method="post" style="margin:0px;" >
<input type="hidden" name="taxIdx" value="<%= taxIdx %>">
<input type="hidden" name="result" value="">
<input type="hidden" name="no_tax" value="">
<input type="hidden" name="result_msg" value="">
<input type="hidden" name="no_iss" value="">
<input type="hidden" name="write_date" value="<%= isueDate %>">
</form>
<% IF application("Svr_Info")="Dev" THEN %>
    <iframe name="ipreSave" id="ipreSave" width="100%" height="200"></iframe>
<% else %>
    <iframe name="ipreSave" id="ipreSave" width="0" height="0"></iframe>
<% end if %>

<script type="text/javascript">
    window.onload=getOnLoad;
</script>

<%
function Get3PLUpcheInfoByTPLCompanyid(tplcompanyid, byRef tplcompanyname, byRef tplgroupid, byRef tplbillUserID, byRef tplbillUserPass)
	dim sqlStr

	sqlStr = " select top 1 t.tplcompanyid, t.tplcompanyname, t.groupid as tplgroupid, billUserID as tplbillUserID, billUserPass as tplbillUserPass "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner_tpl t "
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

session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->