// 1:1 상담게시판		'/차후 사용안함 밑에꺼 쓸것 용만
function PopMyQnaList(userid, orderserial, finishyn){
    var window_width = 1280;
    var window_height = 960;
	var popwin = window.open("/cscenter/board/cscenter_qna_board_list.asp?userid=" + userid + "&orderserial=" + orderserial + "&finishyn=" + finishyn + "&qadiv=","PopMyQnaList","width=" + window_width + " height=" + window_height + " left=50 top=50 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

// 1:1 상담게시판		'/2016.03.28 한용민 생성
function PopMyQna(userid, orderserial, finishyn, qadiv, chargeid, writeid, replyqadiv, userlevel, evalPoint){
    var window_width = 1400;
    var window_height = 960;
	var popwinQna = window.open("/cscenter/board/cscenter_qna_board_list.asp?userid="+userid+"&orderserial="+orderserial+"&finishyn="+finishyn+"&qadiv="+qadiv+"&chargeid="+chargeid+"&writeid="+writeid+"&replyqadiv="+replyqadiv+"&userlevel="+userlevel+"&evalPoint="+evalPoint,"PopMyQna","width=" + window_width + " height=" + window_height + " left=50 top=50 scrollbars=yes resizable=yes status=yes");
	popwinQna.focus();
}

function PopMyQnaListByChargeId(chargeid, finishyn){
    var window_width = 1280;
    var window_height = 960;

	if (chargeid == "vipmanager") {
		// VIP 담당자
		chargeid = "";
		finishyn = "V";
	}

	if (chargeid == "vipnormal") {
		// VIP 일반상담
		chargeid = "";
		finishyn = "E";
	}

	if (chargeid == "vipdeliver") {
		// VIP 배송문의
		chargeid = "";
		finishyn = "D";
	}

	var popwin = window.open("/cscenter/board/cscenter_qna_board_list.asp?chargeid=" + chargeid + "&finishyn=" + finishyn + "&qadiv=","PopMyQnaListByChargeId","width=" + window_width + " height=" + window_height + " left=50 top=50 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}


// ============================================================================
// 주문내역
function PopOrderMaster(){
    var window_width = 1280;
    var window_height = 1024;
	var popwin = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function PopOrderMasterWithCallRing(){
    var window_width = 1280;
    var window_height = 1024;
	var popwin = window.open("/cscenter/ordermaster/ordermasterWithCallRing.asp","PopOrderMasterWithCallRing","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function PopOrderMasterWithCallRingOrderserial(iorderserial){
    var window_width = 1280;
    var window_height = 1024;
	var popwin = window.open("/cscenter/ordermaster/ordermasterWithCallRing.asp?orderserial=" + iorderserial,"PopOrderMasterWithCallRing","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function PopOrderMasterWithCallRingUserid(iuserid){
    var window_width = 1280;
    var window_height = 1024;
	var popwin = window.open("/cscenter/ordermaster/ordermasterWithCallRing.asp?userid=" + iuserid,"PopOrderMasterWithCallRing","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function PopBuyerInfo(orderserial) {
	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    var window_width = 800;
    var window_height = 600;
	var popwin = window.open("/cscenter/ordermaster/order_buyer_info.asp?orderserial=" + orderserial,"PopBuyerInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopReceiverInfo(orderserial) {
	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    var window_width = 800;
    var window_height = 600;
	var popwin = window.open("/cscenter/ordermaster/order_receiver_info.asp?orderserial=" + orderserial,"PopReceiverInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopFlowerDeliverInfo(orderserial) {
	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    var window_width = 300;
    var window_height = 250;
	var popwin = window.open("/cscenter/ordermaster/order_flower_info.asp?orderserial=" + orderserial,"PopFlowerDeliverInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
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

    var window_width = 600;
    var window_height = 400;
	var popwin = window.open("/cscenter/ordermaster/order_nextstep.asp?orderserial=" + orderserial,"PopNextIpkumDiv","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
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

	//var popwin = window.open("/cscenter/action/pop_cs_write_cancel.asp?divcd=20&orderserial=" + orderserial,"PopOpenCancelOrder","width=1000 height=800 scrollbars=yes resizable=yes");
	//var popwin = window.open("/cscenter/action/pop_cs_action_reg.asp?mode=" + mode + "&divcd=" + divcd + "&orderserial=" + orderserial + "&ckAll=" + ckAll,"PopOpenCancelOrder","width=900 height=800 scrollbars=yes resizable=yes");

	//popwin.focus();
}

function PopOpenAddPayment(orderserial){
    var mode    = "regcsas";
    var divcd   = "A999";
    var ckAll   = "";

	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,ckAll);
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

	//var popwin = window.open("/cscenter/action/pop_cs_write_cancel.asp?divcd=21&orderserial=" + orderserial,"PopOpenCancelItem","width=1000 height=800 scrollbars=yes resizable=yes");
	//var popwin = window.open("/cscenter/action/pop_cs_action_reg.asp?mode=" + mode + "&divcd=" + divcd + "&orderserial=" + orderserial + "&ckAll=" + ckAll,"PopOpenCancelOrder","width=900 height=800 scrollbars=yes resizable=yes");
	//popwin.focus();
}


//업체긴급문의
function PopOpenNowReadMe(orderserial){
    var mode    = "regcsas";
    var divcd   = "A060";
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
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
	//var popwin = window.open("/cscenter/action/pop_cs_write_etc.asp?divcd=6&orderserial=" + orderserial,"PopOpenReadMe","width=1000 height=800 scrollbars=yes resizable=yes");
	//var popwin = window.open("/cscenter/action/pop_cs_action_reg.asp?mode=" + mode + "&divcd=" + divcd + "&orderserial=" + orderserial,"PopOpenCancelOrder","width=900 height=800 scrollbars=yes resizable=yes");
	//popwin.focus();
}

// 업체추가정산
function PopOpenUpcheAddJungsan(orderserial){
    var mode    = "regcsas";
    var divcd   = "A700";
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

	//var popwin = window.open("/cscenter/action/pop_cs_write_etc.asp?divcd=9&orderserial=" + orderserial,"PopOpenReadMe","width=1000 height=800 scrollbars=yes resizable=yes");
	//var popwin = window.open("/cscenter/action/pop_cs_action_reg.asp?mode=" + mode + "&divcd=" + divcd + "&orderserial=" + orderserial,"PopOpenCancelOrder","width=900 height=800 scrollbars=yes resizable=yes");
	//popwin.focus();
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

	//var popwin = window.open("/cscenter/action/pop_cs_write_receive.asp?divcd=10&orderserial=" + orderserial,"PopOpenReceiveItemByTenTen","width=1000 height=800 scrollbars=yes resizable=yes");
	//popwin.focus();
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

	//var popwin = window.open("/cscenter/action/pop_cs_write_receive.asp?divcd=4&orderserial=" + orderserial,"PopOpenReceiveItemByUpche","width=1000 height=800 scrollbars=yes resizable=yes");
	//popwin.focus();
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
	//var popwin = window.open("/cscenter/action/pop_cs_write_service.asp?divcd=0&orderserial=" + orderserial,"PopOpenReceiveItemByUpche","width=1000 height=800 scrollbars=yes resizable=yes");
	//popwin.focus();
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
	//var popwin = window.open("/cscenter/action/pop_cs_write_service.asp?divcd=1&orderserial=" + orderserial,"PopOpenServiceItemOmit","width=1000 height=800 scrollbars=yes resizable=yes");
	//popwin.focus();
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
	//var popwin = window.open("/cscenter/action/pop_cs_write_service.asp?divcd=2&orderserial=" + orderserial,"PopOpenServiceItemMore","width=1000 height=800 scrollbars=yes resizable=yes");
	//popwin.focus();
}

// 기타회수
function PopOpenServiceRecvItemMore(orderserial){
    var mode    = "regcsas";
    var divcd   = "A200";

	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
	//var popwin = window.open("/cscenter/action/pop_cs_write_service.asp?divcd=2&orderserial=" + orderserial,"PopOpenServiceRecvItemMore","width=1000 height=800 scrollbars=yes resizable=yes");
	//popwin.focus();
}


// ============================
// CS처리리스트 관리

function Cscenter_Action_List(orderserial, userid, divcd) {
    var window_width = 1280;
    var window_height = 800;
	var popwin = window.open("/cscenter/action/cs_action.asp?orderserial=" + orderserial + "&userid=" + userid + "&divcd=" + divcd,"cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function Cscenter_Action_List_3PL(orderserial, userid, divcd) {
    var window_width = 1280;
    var window_height = 960;
	var popwin = window.open("/admin/etc/3pl/cscenter/action/cs_action_3PL.asp?orderserial=" + orderserial + "&userid=" + userid + "&divcd=" + divcd,"cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

// CS처리 등록/수정
function PopCSActionReg(orderserial){
    PopCSActionCom('',orderserial,'regcsas','','');
}

function PopCSActionReg_3PL(orderserial){
    PopCSActionCom_3PL('',orderserial,'regcsas','','');
}

// CS처리 등록/수정
function PopCSActionEdit(id,mode){
    PopCSActionCom(id,'',mode,'','');

}

function PopCSActionEdit_3PL(id,mode){
    PopCSActionCom_3PL(id,'',mode,'','');

}

// CS처리 등록/수정
function PopCSActionFinish(id,mode){
    PopCSActionCom(id,'',mode,'','');

}

function PopCSActionFinish_3PL(id,mode){
    PopCSActionCom_3PL(id,'',mode,'','');

}

// CS처리 등록/수정 공통
function PopCSActionCom(id,orderserial,mode,divcd,ckAll){
    // var popwin=window.open("/cscenter/action/pop_cs_action_reg.asp?orderserial=" + orderserial + "&id=" + id + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll,"pop_cs_action_reg_" + divcd,"width=1000 height=600 scrollbars=yes resizable=yes");
    var popwin=window.open("/cscenter/action/pop_cs_action_new.asp?orderserial=" + orderserial + "&id=" + id + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll,"pop_cs_action_reg_" + divcd,"width=1400 height=800 scrollbars=yes resizable=yes");
    popwin.focus();
}

function PopCSActionCom_3PL(id,orderserial,mode,divcd,ckAll){
    // var popwin=window.open("/cscenter/action/pop_cs_action_reg.asp?orderserial=" + orderserial + "&id=" + id + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll,"pop_cs_action_reg_" + divcd,"width=1000 height=600 scrollbars=yes resizable=yes");
    var popwin=window.open("/admin/etc/3pl/cscenter/action/pop_cs_action_new_3PL.asp?orderserial=" + orderserial + "&id=" + id + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll,"pop_cs_action_reg_" + divcd,"width=1400 height=800 scrollbars=yes resizable=yes");
    popwin.focus();
}

// CS 등록 사유 선택 Box 팝업
function popCsAsGubunSelect(frm,gubun01,gubun02,gubun01name,gubun02name){
    var popwin=window.open("/cscenter/action/pop_cs_gubun_select.asp?frm=" + frm + "&gubun01=" + gubun01 + "&gubun02=" + gubun02 + "&gubun01name=" + gubun01name + "&gubun02name=" + gubun02name,"pop_cs_gubun_select","width=500 height=300 scrollbars=yes resizable=yes");
    popwin.focus();
    //var retVal=window.showModalDialog("/cscenter/action/pop_cs_gubun_select.asp?frm=" + frm + "&gubun01=" + gubun01 + "&gubun02=" + gubun02 + "&gubun01name=" + gubun01name + "&gubun02name=" + gubun02name,"pop_cs_gubun_select","dialogwidth:400px;dialogheight:300px;center:yes;scroll:yes;resizable:yes;status:no;help:no;");
    //alert(retVal);
}

// CS 환불정보 수정
function popCSrefundInfoEdit(id){
    var popwin=window.open("/cscenter/action/pop_cs_refundinfoedit.asp?id=" + id ,"popCSrefundInfoEdit","width=700 height=400 scrollbars=yes resizable=yes");
    popwin.focus();
}

// 이전 환불 계좌
function popPreReturnAcct(userid, frmname, rebankaccount, rebankownername, rebankname){
    var popwin=window.open("/cscenter/action/pop_cs_PreRefundAccount.asp?userid=" + userid + "&frmname=" + frmname + "&rebankaccount=" + rebankaccount + "&rebankownername=" + rebankownername + "&rebankname=" + rebankname ,"popPreReturnAcct","width=500 height=400 scrollbars=yes resizable=yes");
    popwin.focus();
}

function popRequestReturnAcctLMS(id, orderserial, buyhp) {
    var popwin=window.open("/cscenter/action/pop_cs_RequestRefundAccountLMS.asp?id=" + id + "&orderserial=" + orderserial + "&buyhp=" + buyhp, "popRequestReturnAcctLMS","width=700 height=350 scrollbars=yes resizable=yes");
    popwin.focus();
}


//----------------------
//주문메일 발송
function PopCSMailSendOrder(orderserial){
    var window_width = 250;
    var window_height = 120;
    var popwin=window.open("/cscenter/action/pop_cs_mail_sendorder.asp?orderserial=" + orderserial,"PopCSMailSendOrder","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=no");
    popwin.focus();
}

//고객메일 발송 - 커스터머 메일로 발송 - 회신불가
function PopCSMailSend(email,orderserial,userid){
    var window_width = 600;
    var window_height = 450;
    var popwin=window.open("/cscenter/action/pop_cs_mail_send.asp?email=" + email + "&orderserial=" + orderserial + "&userid=" + userid ,"PopCSMailSend","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}

// 고객파일전송관리     // 2019.11.22 한용민
function PopCSfileSend(userhp,orderserial,userid,asmasteridx){
    var window_width = 1400;
    var window_height = 960;
    var popwin=window.open("/cscenter/action/pop_cs_file_send.asp?userhp=" + userhp + "&orderserial=" + orderserial + "&userid=" + userid + "&asmasteridx=" + asmasteridx ,"PopCSfileSend","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}

//고객SMS 발송
function PopCSSMSSend(reqhp,orderserial,userid,defaultMsg){
    var window_width = 450;
    var window_height = 400;
    var popwin=window.open("/cscenter/action/pop_cs_sms_send.asp?reqhp=" + reqhp + "&orderserial=" + orderserial + "&userid=" + userid + "&defaultMsg=" + defaultMsg,"PopCSSMSSend","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + "  scrollbars=no resizable=yes");
    popwin.focus();
}

//고객SMS 발송 v2
function PopCSSMSSendNew(reqhp, orderserial, userid, makerid, itemid, orderdetailidx, defaultMsg) {
    var window_width = 450;
    var window_height = 400;
    var popwin=window.open("/cscenter/action/pop_cs_sms_send.asp?reqhp=" + reqhp + "&orderserial=" + orderserial + "&userid=" + userid + "&makerid=" + makerid + "&itemid=" + itemid + "&orderdetailidx=" + orderdetailidx + "&defaultMsg=" + defaultMsg,"PopCSSMSSendNew","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + "  scrollbars=no resizable=yes");
    popwin.focus();
}

//고객 회수/맞교환/서비스발송 주소지 변경
function popEditCsDelivery(id){
    var window_width = 600;
    var window_height = 450;
    var popwin=window.open("/cscenter/action/pop_CsDeliveryEdit.asp?id=" + id ,"pop_CsDeliveryEdit","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}

function popEditCsDelivery_3PL(id){
    var window_width = 600;
    var window_height = 450;
    var popwin=window.open("/admin/etc/3pl/cscenter/action/pop_CsDeliveryEdit_3PL.asp?id=" + id ,"pop_CsDeliveryEdit","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}

//영수증 출력
function popOrderReceipt(orderserial){
    var window_width = 750;
    var window_height = 700;
    var popwin=window.open("/common/pop_order_receipt.asp?orderserial=" + orderserial ,"popOrderReceipt","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}


function changeSongjang(csid){
    var popwin = window.open('/cscenter/action/popChangeSongjang.asp?id=' + csid,'popChangeSongjang','width=400,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function changeSongjang_3PL(csid){
    var popwin = window.open('/admin/etc/3pl/cscenter/action/popChangeSongjang_3PL.asp?id=' + csid,'popChangeSongjang','width=400,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popForeignDeliverInfo(countryCode, companyCode){
    if ((countryCode=="")||(countryCode=="KR")){
        alert('해외배송이 아닙니다.');
        return;
    }

	if (companyCode == undefined) {
		companyCode = 'EMS';
	}

	if (companyCode == 'EMS') {
		var url = "http://service.epost.go.kr/front.EmsApplyGoCondition.postal?nation="+countryCode;
		var popwin = window.open(url,'popEmsServiceArea','width=650,height=600,scrollbars=yes,resizable=yes');
		popwin.focus();
	} else if (companyCode == 'UPS') {
		var url = "https://wwwapps.ups.com/ctc/request?loc=ko_KR";
		var popwin = window.open(url,'popEmsServiceArea','width=1200,height=600,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

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
    var popwin = window.open('/cscenter/action/popChangeRequireDetail_UTF8.asp?orderserial=' + iorderserial + '&id=' + idetailid,'popChangeRequireDetail','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}
