
// ============================================================================
// 1:1 ���Խ���
function PopMyQnaList(userid, orderserial, finishyn){
    var window_width = 1024;
    var window_height = 768;
	var popwin = window.open("/cscenterv2/board/cscenter_qna_board_list.asp?userid=" + userid + "&orderserial=" + orderserial + "&finishyn=" + finishyn + "&qadiv=","PopMyQnaList","width=" + window_width + " height=" + window_height + " left=50 top=50 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}


// ============================================================================
// �ֹ�����
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
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 300;
    var window_height = 250;
	var popwin = window.open("/cscenterv2/order/order_buyer_info.asp?orderserial=" + orderserial,"PopBuyerInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopReceiverInfo(orderserial) {
	if (orderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 300;
    var window_height = 250;
	var popwin = window.open("/cscenterv2/order/order_receiver_info.asp?orderserial=" + orderserial,"PopReceiverInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopFlowerDeliverInfo(orderserial) {
	if (orderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 300;
    var window_height = 250;
	var popwin = window.open("/cscenterv2/order/order_flower_info.asp?orderserial=" + orderserial,"PopFlowerDeliverInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopSearchBankPayByName(accountname) {
	if (accountname == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    alert("�۾����Դϴ�.");
}

function PopNextIpkumDiv(orderserial){
    if (orderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 300;
    var window_height = 160;
	var popwin = window.open("/cscenterv2/order/order_nextstep.asp?orderserial=" + orderserial,"PopNextIpkumDiv","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

// ============================================================================
// �����Լ�

function GetCenterX(window_width) {
	return (screen.width - window_width)/2;
}

function GetCenterY(window_height) {
	return (screen.height - window_height)/2;
}


// ============================
//��ü���
function PopOpenCancelOrder(orderserial){
    var mode    = "regcsas";
    var divcd   = "A008";
    var ckAll   = "on";

	if (orderserial == "") {
	        alert("���� �ֹ��� �����ϼ���.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,ckAll);
}

function PopOpenCancelOrderLecture(orderserial){
    var mode    = "regcsas";
    var divcd   = "A008";
    var ckAll   = "on";

	if (orderserial == "") {
	        alert("���� �ֹ��� �����ϼ���.");
	        return;
    }

    PopLectureCSActionCom('',orderserial,mode,divcd,ckAll);
}

//�κ����
function PopOpenCancelItem(orderserial){
    var mode    = "regcsas";
    var divcd   = "A008";
    var ckAll   = "";

	if (orderserial == "") {
	        alert("���� �ֹ��� �����ϼ���.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,ckAll);
}

function PopOpenCancelItemLecture(orderserial){
    var mode    = "regcsas";
    var divcd   = "A008";
    var ckAll   = "";

	if (orderserial == "") {
	        alert("���� �ֹ��� �����ϼ���.");
	        return;
    }

    PopLectureCSActionCom('',orderserial,mode,divcd,ckAll);
}

//�������
function PopOpenReadMe(orderserial){
    var mode    = "regcsas";
    var divcd   = "A006";
	if (orderserial == "") {
	        alert("���� �ֹ��� �����ϼ���.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

//��Ÿ�������
function PopOpenEtcNote(orderserial){
    var mode    = "regcsas";
    var divcd   = "A009";
	if (orderserial == "") {
	        alert("���� �ֹ��� �����ϼ���.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}



//ȸ����û
function PopOpenReceiveItemByTenTen(orderserial){
    var mode    = "regcsas";
    var divcd   = "A010";
	if (orderserial == "") {
	        alert("���� �ֹ��� �����ϼ���.");
	        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

//��ǰ����
function PopOpenReceiveItemByUpche(orderserial){
    var mode    = "regcsas";
    var divcd   = "A004";

	if (orderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

// ����Ȯ�� �� ȯ��
function PopOpenReceiveItemByUpcheLecture(orderserial){
    var mode    = "regcsas";
    var divcd   = "A004";

	if (orderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    PopLectureCSActionCom('',orderserial,mode,divcd,'');
}

//�±�ȯ
function PopOpenServiceItemChange(orderserial){
    var mode    = "regcsas";
    var divcd   = "A000";

	if (orderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

//������߼�
function PopOpenServiceItemOmit(orderserial){
    var mode    = "regcsas";
    var divcd   = "A001";

	if (orderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

//���񽺹߼�
function PopOpenServiceItemMore(orderserial){
    var mode    = "regcsas";
    var divcd   = "A002";

	if (orderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    PopCSActionCom('',orderserial,mode,divcd,'');
}

// ============================
// CSó������Ʈ ����

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

// CSó�� ���/����
function PopCSActionReg(orderserial){
    PopCSActionCom('',orderserial,'regcsas','','');
}

// CSó�� ���/����
function PopCSActionEdit(id,mode){
    PopCSActionCom(id,'',mode,'','');

}

function PopCSActionEdit_Lecture(id,mode){
    PopLectureCSActionCom(id,'',mode,'','');

}

// CSó�� ���/����
function PopCSActionFinish(id,mode){
    PopCSActionCom(id,'',mode,'','');

}

function PopCSActionFinish_Lecture(id,mode){
    PopLectureCSActionCom(id,'',mode,'','');

}

// CSó�� ���/���� ����
function PopCSActionCom(id,orderserial,mode,divcd,ckAll){
    var popwin=window.open("/cscenterv2/cs/pop_cs_register.asp?orderserial=" + orderserial + "&id=" + id + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll,"pop_cs_action_reg_" + divcd,"width=1000 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

// CSó�� ���/���� ����
function PopLectureCSActionCom(id,orderserial,mode,divcd,ckAll){
    var popwin=window.open("/cscenterv2/cs_lecture/pop_lec_cs_register.asp?orderserial=" + orderserial + "&id=" + id + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll,"pop_cs_action_reg_" + divcd,"width=1000 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

// CS ��� ���� ���� Box �˾�
function popCsAsGubunSelect(frm,gubun01,gubun02,gubun01name,gubun02name){
    var popwin=window.open("/cscenterv2/cs/pop_cs_gubun_select.asp?frm=" + frm + "&gubun01=" + gubun01 + "&gubun02=" + gubun02 + "&gubun01name=" + gubun01name + "&gubun02name=" + gubun02name,"pop_cs_gubun_select","width=500 height=300 scrollbars=yes resizable=yes");
    popwin.focus();
}

// CS ȯ������ ����
function popCSrefundInfoEdit(id){
    var popwin=window.open("/cscenterv2/cs/pop_cs_refundinfoedit.asp?id=" + id ,"popCSrefundInfoEdit","width=700 height=400 scrollbars=yes resizable=yes");
    popwin.focus();
}

// ���� ȯ�� ����
function popPreReturnAcct(userid, frmname, rebankaccount, rebankownername, rebankname){
    var popwin=window.open("/cscenterv2/cs/pop_cs_PreRefundAccount.asp?userid=" + userid + "&frmname=" + frmname + "&rebankaccount=" + rebankaccount + "&rebankownername=" + rebankownername + "&rebankname=" + rebankname ,"popPreReturnAcct","width=500 height=400 scrollbars=yes resizable=yes");
    popwin.focus();
}


//----------------------
//�ֹ����� �߼�
function PopCSMailSendOrder(orderserial){
    var window_width = 250;
    var window_height = 120;
    var popwin=window.open("/cscenterv2/cs/pop_cs_mail_sendorder.asp?orderserial=" + orderserial,"PopCSMailSendOrder","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=no");
    popwin.focus();
}

//������ �߼� - Ŀ���͸� ���Ϸ� �߼� - ȸ�źҰ�
function PopCSMailSend(email,orderserial,userid){
    var window_width = 300;
    var window_height = 100;
    var popwin=window.open("/cscenterv2/cs/pop_cs_mail_send.asp?email=" + email + "&orderserial=" + orderserial + "&userid=" + userid ,"PopCSMailSend","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}

//��SMS �߼�
function PopCSSMSSend(reqhp,orderserial,userid,defaultMsg){
    var window_width = 250;
    var window_height = 210;
    var popwin=window.open("/cscenterv2/cs/pop_cs_sms_send.asp?reqhp=" + reqhp + "&orderserial=" + orderserial + "&userid=" + userid + "&defaultMsg=" + defaultMsg,"PopCSSMSSend","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + "  scrollbars=no resizable=no");
    popwin.focus();
}

//�� ȸ��/�±�ȯ/���񽺹߼� �ּ��� ����
function popEditCsDelivery(id){
    var window_width = 600;
    var window_height = 450;
    var popwin=window.open("/cscenterv2/cs/pop_CsDeliveryEdit.asp?id=" + id ,"pop_CsDeliveryEdit","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}

//������ ���
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
        alert('�ؿܹ���� �ƴմϴ�.');
        return;
    }


    var url = "http://service.epost.go.kr/front.EmsApplyGoCondition.postal?nation="+countryCode;
	var popwin = window.open(url,'popEmsServiceArea','width=650,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();

    //�߷�
    //http://2009www.10x10.co.kr/inipay/popEmsCharge.asp?areaCode=1
}

function popForeignDeliverPay(areaCode){
    var url = "http://www.10x10.co.kr/inipay/popEmsCharge.asp?areaCode="+areaCode;
	var popwin = window.open(url,'popEmsServiceAreaPrice','width=650,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//�ֹ����۹��� ����
function EditRequireDetail(iorderserial,idetailid){
	alert('[���� �۾� ����] �ֹ� ���� & �����ϴ� �˾� �ߴ� �ڸ�');
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

//��SMS �߼�
// reqhp, orderserial, userid, defaultMsg, makerid, itemid, orderdetailidx
function PopCSSMSSendNew(args) {
	var params = object2queryparams(args);
    var window_width = 450;
    var window_height = 400;
    var popwin=window.open("/cscenterv2/action/pop_cs_sms_send.asp" + params,"PopCSSMSSendNew","width=" + window_width + " height=" + window_height + "  scrollbars=no resizable=yes");
    popwin.focus();
}

//������ �߼� - Ŀ���͸� ���Ϸ� �߼� - ȸ�źҰ�
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
