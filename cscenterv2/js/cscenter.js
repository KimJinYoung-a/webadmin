
// ============================================================================
// 1:1 상담게시판
function PopMyQnaList(userid, orderserial, finishyn){
    var window_width = 1024;
    var window_height = 768;
	var popwin = window.open("/cscenterv2/board/cscenter_qna_board_list.asp?userid=" + userid + "&orderserial=" + orderserial + "&finishyn=" + finishyn + "&qadiv=","PopMyQnaList","width=" + window_width + " height=" + window_height + " left=50 top=50 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}


// ============================================================================
// 주문내역
function Poporder(){
    var window_width = 1280;
    var window_height = 1024;
	var popwin = window.open("/cscenterv2/order/order.asp","Poporder","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function PoporderWithCallRing(){
    var window_width = 1280;
    var window_height = 1024;
	var popwin = window.open("/cscenterv2/order/orderWithCallRing.asp","PoporderWithCallRing","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function PoporderWithCallRingOrderserial(iorderserial){
    var window_width = 1280;
    var window_height = 1024;
	var popwin = window.open("/cscenterv2/order/orderWithCallRing.asp?orderserial=" + iorderserial,"PoporderWithCallRing","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function PopBuyerInfo(orderserial) {
	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    var window_width = 300;
    var window_height = 250;
	var popwin = window.open("/cscenterv2/order/order_buyer_info.asp?orderserial=" + orderserial,"PopBuyerInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopReceiverInfo(orderserial) {
	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    var window_width = 300;
    var window_height = 250;
	var popwin = window.open("/cscenterv2/order/order_receiver_info.asp?orderserial=" + orderserial,"PopReceiverInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopFlowerDeliverInfo(orderserial) {
	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    var window_width = 300;
    var window_height = 250;
	var popwin = window.open("/cscenterv2/order/order_flower_info.asp?orderserial=" + orderserial,"PopFlowerDeliverInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopSearchBankPayByName(accountname) {
	if (accountname == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    alert("작업중입니다.");
}

function PopNextIpkumDiv(orderserial){
    if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    var window_width = 300;
    var window_height = 160;
	var popwin = window.open("/cscenterv2/order/order_nextstep.asp?orderserial=" + orderserial,"PopNextIpkumDiv","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

// ============================================================================
// 보조함수

function GetCenterX(window_width) {
	return (screen.width - window_width)/2;
}

function GetCenterY(window_height) {
	return (screen.height - window_height)/2;
}


// ============================
//전체취소
function PopOpenCancelOrder(orderserial){
    var mode    = "regcsas";
    var divcd   = "A008";
    var ckAll   = "on";

	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,ckAll);
}

function PopOpenCancelOrderLecture(orderserial){
    var mode    = "regcsas";
    var divcd   = "A008";
    var ckAll   = "on";

	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
    }

    PopLectureCSActionCom('',orderserial,mode,divcd,ckAll);
}

//부분취소
function PopOpenCancelItem(orderserial){
    var mode    = "regcsas";
    var divcd   = "A008";
    var ckAll   = "";

	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,ckAll);
}

function PopOpenCancelItemLecture(orderserial){
    var mode    = "regcsas";
    var divcd   = "A008";
    var ckAll   = "";

	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
    }

    PopLectureCSActionCom('',orderserial,mode,divcd,ckAll);
}

//배송유의
function PopOpenReadMe(orderserial){
    var mode    = "regcsas";
    var divcd   = "A006";
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

//기타내역등록
function PopOpenEtcNote(orderserial){
    var mode    = "regcsas";
    var divcd   = "A009";
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}



//회수요청
function PopOpenReceiveItemByTenTen(orderserial){
    var mode    = "regcsas";
    var divcd   = "A010";
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

//반품접수
function PopOpenReceiveItemByUpche(orderserial){
    var mode    = "regcsas";
    var divcd   = "A004";

	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

// 강좌확정 후 환불
function PopOpenReceiveItemByUpcheLecture(orderserial){
    var mode    = "regcsas";
    var divcd   = "A004";

	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    PopLectureCSActionCom('',orderserial,mode,divcd,'');
}

//맞교환
function PopOpenServiceItemChange(orderserial){
    var mode    = "regcsas";
    var divcd   = "A000";

	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

//누락재발송
function PopOpenServiceItemOmit(orderserial){
    var mode    = "regcsas";
    var divcd   = "A001";

	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

//서비스발송
function PopOpenServiceItemMore(orderserial){
    var mode    = "regcsas";
    var divcd   = "A002";

	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

// ============================
// CS처리리스트 관리

function Cscenter_Action_List(orderserial, userid, divcd) {
    var window_width = 1280;
    var window_height = 960;
	var popwin = window.open("/cscenterv2/cs/frame.asp?orderserial=" + orderserial + "&userid=" + userid + "&divcd=" + divcd,"cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function Cscenter_Action_List_Lecture(orderserial, userid, divcd) {
    var window_width = 1280;
    var window_height = 960;
	var popwin = window.open("/cscenterv2/cs_lecture/frame.asp?orderserial=" + orderserial + "&userid=" + userid + "&divcd=" + divcd,"cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

// CS처리 등록/수정
function PopCSActionReg(orderserial){
    PopCSActionCom('',orderserial,'regcsas','','');
}

// CS처리 등록/수정
function PopCSActionEdit(id,mode){
    PopCSActionCom(id,'',mode,'','');

}

function PopCSActionEdit_Lecture(id,mode){
    PopLectureCSActionCom(id,'',mode,'','');

}

// CS처리 등록/수정
function PopCSActionFinish(id,mode){
    PopCSActionCom(id,'',mode,'','');

}

function PopCSActionFinish_Lecture(id,mode){
    PopLectureCSActionCom(id,'',mode,'','');

}

// CS처리 등록/수정 공통
function PopCSActionCom(id,orderserial,mode,divcd,ckAll){
    var popwin=window.open("/cscenterv2/cs/pop_cs_register.asp?orderserial=" + orderserial + "&id=" + id + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll,"pop_cs_action_reg_" + divcd,"width=1000 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

// CS처리 등록/수정 공통
function PopLectureCSActionCom(id,orderserial,mode,divcd,ckAll){
    var popwin=window.open("/cscenterv2/cs_lecture/pop_lec_cs_register.asp?orderserial=" + orderserial + "&id=" + id + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll,"pop_cs_action_reg_" + divcd,"width=1000 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

// CS 등록 사유 선택 Box 팝업
function popCsAsGubunSelect(frm,gubun01,gubun02,gubun01name,gubun02name){
    var popwin=window.open("/cscenterv2/cs/pop_cs_gubun_select.asp?frm=" + frm + "&gubun01=" + gubun01 + "&gubun02=" + gubun02 + "&gubun01name=" + gubun01name + "&gubun02name=" + gubun02name,"pop_cs_gubun_select","width=500 height=300 scrollbars=yes resizable=yes");
    popwin.focus();
}

// CS 환불정보 수정
function popCSrefundInfoEdit(id){
    var popwin=window.open("/cscenterv2/cs/pop_cs_refundinfoedit.asp?id=" + id ,"popCSrefundInfoEdit","width=700 height=400 scrollbars=yes resizable=yes");
    popwin.focus();
}

// 이전 환불 계좌
function popPreReturnAcct(userid, frmname, rebankaccount, rebankownername, rebankname){
    var popwin=window.open("/cscenterv2/cs/pop_cs_PreRefundAccount.asp?userid=" + userid + "&frmname=" + frmname + "&rebankaccount=" + rebankaccount + "&rebankownername=" + rebankownername + "&rebankname=" + rebankname ,"popPreReturnAcct","width=500 height=400 scrollbars=yes resizable=yes");
    popwin.focus();
}


//----------------------
//주문메일 발송
function PopCSMailSendOrder(orderserial){
    var window_width = 250;
    var window_height = 120;
    var popwin=window.open("/cscenterv2/cs/pop_cs_mail_sendorder.asp?orderserial=" + orderserial,"PopCSMailSendOrder","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=no");
    popwin.focus();
}

//고객메일 발송 - 커스터머 메일로 발송 - 회신불가
function PopCSMailSend(email,orderserial,userid){
    var window_width = 300;
    var window_height = 100;
    var popwin=window.open("/cscenterv2/cs/pop_cs_mail_send.asp?email=" + email + "&orderserial=" + orderserial + "&userid=" + userid ,"PopCSMailSend","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}

//고객SMS 발송
function PopCSSMSSend(reqhp,orderserial,userid,defaultMsg){
    var window_width = 250;
    var window_height = 210;
    var popwin=window.open("/cscenterv2/cs/pop_cs_sms_send.asp?reqhp=" + reqhp + "&orderserial=" + orderserial + "&userid=" + userid + "&defaultMsg=" + defaultMsg,"PopCSSMSSend","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + "  scrollbars=no resizable=no");
    popwin.focus();
}

//고객 회수/맞교환/서비스발송 주소지 변경
function popEditCsDelivery(id){
    var window_width = 600;
    var window_height = 450;
    var popwin=window.open("/cscenterv2/cs/pop_CsDeliveryEdit.asp?id=" + id ,"pop_CsDeliveryEdit","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}

//영수증 출력
function popOrderReceipt(orderserial){
    var window_width = 750;
    var window_height = 700;
    var popwin=window.open("/cscenterv2/common/pop_order_receipt.asp?orderserial=" + orderserial ,"popOrderReceipt","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}


function changeSongjang(csid){
    var popwin = window.open('/cscenterv2/cs/popChangeSongjang.asp?id=' + csid,'popChangeSongjang','width=400,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popForeignDeliverInfo(countryCode){
    if ((countryCode=="")||(countryCode=="KR")){
        alert('해외배송이 아닙니다.');
        return;
    }


    var url = "http://service.epost.go.kr/front.EmsApplyGoCondition.postal?nation="+countryCode;
	var popwin = window.open(url,'popEmsServiceArea','width=650,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();

    //중량
    //http://2009www.10x10.co.kr/inipay/popEmsCharge.asp?areaCode=1
}

function popForeignDeliverPay(areaCode){
    var url = "http://www.10x10.co.kr/inipay/popEmsCharge.asp?areaCode="+areaCode;
	var popwin = window.open(url,'popEmsServiceAreaPrice','width=650,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//주문제작문구 수정
function EditRequireDetail(iorderserial,idetailid){
	alert('[차후 작업 예정] 주문 수정 & 저장하는 팝업 뜨는 자리');
	return;

    var popwin = window.open('/cscenterv2/cs/popChangeRequireDetail.asp?orderserial=' + iorderserial + '&id=' + idetailid,'popChangeRequireDetail','width=400,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

/*
  popCallRingFingers({
      sitename:"academy",
      orderserial:"B6102178401"
  });
*/
// sitename, phoneNumber, orderserial, id, userid
function popCallRingFingers(args) {
    var popwinName = "popCallRingFingers_" + Math.floor(Date.now() / 1000);
	var params = object2queryparams(args);

    var popwin = window.open('/cscenterv2/ordermaster/ordermasterWithCallRing_FIN.asp' + params, popwinName, 'width=1680,height=1000,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//고객SMS 발송
// reqhp, orderserial, userid, defaultMsg, makerid, itemid, orderdetailidx
function PopCSSMSSendNew(args) {
	var params = object2queryparams(args);
    var window_width = 450;
    var window_height = 400;
    var popwin=window.open("/cscenterv2/action/pop_cs_sms_send.asp" + params,"PopCSSMSSendNew","width=" + window_width + " height=" + window_height + "  scrollbars=no resizable=yes");
    popwin.focus();
}

//고객메일 발송 - 커스터머 메일로 발송 - 회신불가
// email, orderserial, userid
function PopCSMailSend(args) {
	var params = object2queryparams(args);
    var window_width = 600;
    var window_height = 450;
    var popwin=window.open("/cscenterv2/action/pop_cs_mail_send.asp" + params,"PopCSMailSend","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}

function object2queryparams(obj) {
	var ret = "", name;

	for (name in obj) {
		if (ret === "") {
			ret = "?" + name + "=" + obj[name];
		} else {
			ret = ret + "&" + name + "=" + obj[name];
		}
	}

	return ret;
}
