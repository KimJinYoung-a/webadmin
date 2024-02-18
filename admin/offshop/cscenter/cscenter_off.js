//오프라인 cs 센터 js
//2011.03.08 한용민 생성

function GetCenterX(window_width) {
	return (screen.width - window_width)/2;
}

function GetCenterY(window_height) {
	return (screen.height - window_height)/2;
}

//구매자 정보 수정
function PopBuyerInfo_off(masteridx) {
	if (masteridx == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    var window_width = 300;
    var window_height = 250;
	var PopBuyerInfo_off = window.open("/admin/offshop/cscenter/order/order_buyer_info.asp?masteridx=" + masteridx,"PopBuyerInfo_off","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	PopBuyerInfo_off.focus();
}

//고객SMS 발송
function PopCSSMSSend_off(reqhp,masteridx,defaultMsg){
    var window_width = 250;
    var window_height = 210;
    var popwin=window.open("/admin/offshop/cscenter/action/pop_cs_sms_send.asp?reqhp=" + reqhp + "&masteridx=" + masteridx + "&defaultMsg=" + defaultMsg,"PopCSSMSSend_off","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + "  scrollbars=no resizable=no");
    popwin.focus();
}

//고객메일 발송 - 커스터머 메일로 발송 - 회신불가
function PopCSMailSend_off(email,masteridx){
    var window_width = 600;
    var window_height = 450;
    var PopCSMailSend_off=window.open("/admin/offshop/cscenter/action/pop_cs_mail_send.asp?email=" + email + "&masteridx=" + masteridx,"PopCSMailSend_off","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    PopCSMailSend_off.focus();
}

//배송지정보수정
function PopReceiverInfo_off(masteridx) {
	if (masteridx == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    var window_width = 1000;
    var window_height = 250;
	var PopReceiverInfo_off = window.open("/admin/offshop/cscenter/order/order_receiver_info.asp?masteridx=" + masteridx,"PopReceiverInfo_off","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	PopReceiverInfo_off.focus();
}

//관련 cs 리스트
function Cscenter_Action_List_off(masteridx, divcd) {
    var window_width = 1024;
    var window_height = 768;
	var Cscenter_Action_List_off = window.open("/admin/offshop/cscenter/action/cs_action.asp?masteridx=" + masteridx + "&divcd=" + divcd,"Cscenter_Action_List_off","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	Cscenter_Action_List_off.focus();
}

//주문상품 수정
function popOrderDetailEdit(detailidx){
	var popwin = window.open('/admin/offshop/cscenter/order/orderdetailedit.asp?detailidx=' + detailidx,'orderdetailedit','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// CS처리 등록/수정
function PopCSActionEdit(masteridx,mode,csmasteridx){
    PopCSActionCom(masteridx,'',mode,'','',csmasteridx);
}

// CS처리 등록/수정 공통
function PopCSActionCom(masteridx,orderno,mode,divcd,ckAll,csmasteridx){
    var popwin=window.open("/admin/offshop/cscenter/action/pop_cs_action_new.asp?orderno=" + orderno + "&masteridx=" + masteridx + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll+ "&csmasteridx=" + csmasteridx,"pop_cs_action_reg_" + divcd,"width=1000 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

// CS처리 등록/수정
function PopCSActionFinish(masteridx,mode,csmasteridx){
    PopCSActionCom(masteridx,'',mode,'','',csmasteridx);
}

//서비스발송
function PopOpenServiceItemMore(masteridx){
    var mode    = "regcsas";
    var divcd   = "A002";

	if (masteridx == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    PopCSActionCom(masteridx,'',mode,divcd,'','');
}

//취소
function PopOpenCancelItem(masteridx){
    var mode    = "regcsas";
    var divcd   = "A008";

	if (masteridx == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    PopCSActionCom(masteridx,'',mode,divcd,'','');
}

//맞교환
function PopOpenServiceItemChange(masteridx){
    var mode    = "regcsas";
    var divcd   = "A000";

	if (masteridx == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }
		
    PopCSActionCom(masteridx,'',mode,divcd,'','');
}

//미출고상품보기
function misendmaster(v){
	var popwin = window.open("/admin/offshop/cscenter/order/misendmaster_main.asp?masteridx=" + v,"misendmaster","width=1027 height=768 scrollbars=yes resizable=yes");
	popwin.focus();
}

//영수증 출력
function popOrderReceipt(masteridx){
    var window_width = 750;
    var window_height = 700;
    var popwin=window.open("/admin/offshop/cscenter/order/pop_order_receipt.asp?masteridx=" + masteridx ,"popOrderReceipt","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}

//배송유의
function PopOpenReadMe(masteridx,csmasteridx){
    var mode    = "regcsas";
    var divcd   = "A006";
	if (masteridx == "" && csmasteridx == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    PopCSActionCom(masteridx,'',mode,divcd,'',csmasteridx);
}

// 선택상품 저장
function SaveCheckedItemList(frm) {
    var e;
    var ischecked = false;
    var checkitemExists = false;

    var orderdetailidx = "";
	var regitemno = "";
    var causecontent = "";

    frm.detailitemlist.value = "";

    for (var i = 0; i < frm.length; i++) {
        e = frm.elements[i];

        if (e.name == "dummystarter") {
            ischecked = false;
            orderdetailidx = "";
            regitemno = "";
            causecontent = "";            
        }

        if (e.name == "orderdetailidx") {
        	if (e.type != "checkbox") {
        		continue;
        	}

            if (e.checked == true) {
                ischecked = true;
                orderdetailidx = e.value;
                checkitemExists = true;
            } else {
                ischecked = false;
                orderdetailidx = "";
            }
        }

        if ((ischecked == true) && (e.name.indexOf("regitemno") == 0)) {
			if (IsStatusEdit && IsCSReturnProcess) {
	            if ((e.value*1)<0){
	                alert('수량을 입력하세요.');
	                e.focus();
	                e.select();
	                return false;
	            }
			} else {
	            if ((e.value*1)==0){
	                alert('수량을 입력하세요.');
	                e.focus();
	                e.select();
	                return false;
	            }
			}

            regitemno = e.value;
        }

        if (e.name == "dummystopper") {
            if (ischecked == true) {
                frm.detailitemlist.value = frm.detailitemlist.value + "|" + orderdetailidx + "\t" + regitemno + "\t" + causecontent;
                ischecked = false;
                regitemno = "";
                causecontent = "";
            }
        }
    }

    //기타내역, 서비스발송 , 환불요청, 출고시유의사항, 업체 추가 정산 - 상세내역 체크 안함.
    if ((divcd=="A009")||(divcd=="A002")||(divcd=="A003")||(divcd=="A005")||(divcd=="A006")||(divcd=="A700")){
        // no- check

    }else{
        if (!checkitemExists){
            alert('선택된 상세내역이 없습니다.');
            return false;
        }
    }

    return true;    
}

// 상품선택시 확인할 것들
function CheckSelect(comp){
    var chkidx = comp.value;
    var frm = document.frmaction;

    // 단일 브랜드만 선택 가능하게
    // 반품 담당 브랜드 저장
    DisableUpcheDeliver(frm);

	// 업배 반품/맞교환의 경우 대상 브랜드 체크
	DispCheckedUpcheID(frm);
}

// 업체 추가정산 관련 브랜드 아이디 가져오기
function DispCheckedUpcheID(frm) {
    var checkedUpcheid = "";
    var UpcheDuplicated = false;
    var IsUpcheReturn;

	if ((divcd == "A004") || (divcd == "A000")) { // 반품접수(업체배송), 맞교환출고
		IsUpcheReturn = true;
	} else {
		IsUpcheReturn = false;
	}

    if (!frm.buf_requiremakerid) {
        return;
    }

	// 선택된 디테일중에서
	//  - 브랜드 가져오기, 서로 다른 두개의 브랜드가 있으면 중복 표시(선택은 하나의 브랜드 상품으로만 해야한다.)
	//
	//  - 반품접수(업체배송), 맞교환출고 이고 업체배송이면 가져오기
	//  - 이외의 경우 브랜드 가져오기
    for(var i = 0; i < frm.orderdetailidx.length; i++) {
		if (frm.orderdetailidx[i].type != "checkbox") {
			continue;
		}

        if (frm.orderdetailidx[i].checked){
            if (IsUpcheReturn){
                if (frm.isupchebeasong[i].value=="Y"){
                    if (checkedUpcheid!="") {
                        if (checkedUpcheid != frm.makerid[i].value){
                            UpcheDuplicated = true;
                        }
                        checkedUpcheid = frm.makerid[i].value;
                    }
                    checkedUpcheid = frm.makerid[i].value;
                }
            }else{
                if (checkedUpcheid!="") {
                    if (checkedUpcheid != frm.makerid[i].value){
                        UpcheDuplicated = true;
                    }
                    checkedUpcheid = frm.makerid[i].value;
                }
                checkedUpcheid = frm.makerid[i].value;
            }
        }
    }

    frm.buf_requiremakerid.value = "";

	if (UpcheDuplicated) {
		alert("두개의 브랜드가 동시에 선택되어 있습니다.(중복불가) 내역을 삭제하세요.");
	}

    if ((!UpcheDuplicated)&&(checkedUpcheid!="")){
        if (frm.buf_requiremakerid){
            frm.buf_requiremakerid.value = checkedUpcheid;
        }
    }
}

// ============================================================================
// 반품시 단일 브랜드만 선택할수 있도록 변경
// 반품 담당 브랜드 저장
// 텐배 체크시 - 업배 Disable
// 업배 체크시 - 텐배 및 다른 브랜드 Disable
// ============================================================================
function DisableUpcheDeliver(frm) {
    var upbeaMakerid;
    var checkfound;

    var objdeliver, objitem;
    var e;

    if (!frm.Deliverdetailidx) return;
    if (!frm.orderdetailidx) return;

	checkfound = false;
	upbeaMakerid = "";
	
    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        objitem = frm.orderdetailidx[i];

        if (objitem.type != "checkbox") {
        	continue;
        }

        if (objitem.checked) {
        	if ((frm.odlvtype[i].value == "1") || (frm.odlvtype[i].value == "4") || (frm.odlvtype[i].value == "0")) {
        		// 텐배
	        	upbeaMakerid = "";
        	} else {
        		// 업배
	        	upbeaMakerid = frm.makerid[i].value;
        	}
        	checkfound = true;
        	break;
        }
    }

	// 반품 담당 브랜드 저장
	if (checkfound != true) {
        frm.requireupche.value = "";
        frm.requiremakerid.value = "";
	} else {
		if (upbeaMakerid.length < 1) {
	        frm.requireupche.value = "N";
	        frm.requiremakerid.value = "";
		} else {
	        frm.requireupche.value = "Y";
	        frm.requiremakerid.value = upbeaMakerid;
		}
	}

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        objitem = frm.orderdetailidx[i];

        if (objitem.type != "checkbox") {
        	continue;
        }

		if (checkfound != true) {
		    // TR 의 bgColor 를 구한다.(색이 FFFFFF 인것만 활성화 할 수 있다.)
		    e = objitem;
			while (e.tagName!="TR") {
				e = e.parentElement;
			}

			if (e.bgColor == "#ffffff") {
				objitem.disabled = false;
			}
		} else {
	        if (
	        	((upbeaMakerid.length < 1) && ((frm.odlvtype[i].value != "1") && (frm.odlvtype[i].value != "4") && (frm.odlvtype[i].value != "0") ))
	        	||
	        	((upbeaMakerid.length > 0) && (upbeaMakerid != frm.makerid[i].value))
	        ) {
	            objitem.checked = false;
	            objitem.disabled = true;
	        }
		}
    }
}

//cs접수
function CsRegProc(frm){

    if (frm.divcd.value.length<1){
        alert("접수 구분을 선택하세요.");
        frm.divcd.focus();
        return;
    }

    if (frm.title.value.length<1) {
        alert("제목을 입력하세요.");
        frm.title.focus();
        return;
    }
	
	//주문취소
	if(divcd =='A008'){
		//출고완료가 아닐경우
		if (IsOrderMasterState != '8'){
//			if (frm.cancelorderno.value.length<1) {
//			    alert("취소한 주문번호를 입력해주세요(마이너스주문)");
//			    frm.cancelorderno.focus();
//			    return;
//			}
		}
	}
    
    //선택 상품 체크
    if (!SaveCheckedItemList(frm)) {
		return;
    }
         
    if(confirm("접수 하시겠습니까?")){
        frm.submit();
    }
}

//CS 수정
function CsRegEditProc(frm){
    if (frm.divcd.value.length<1){
        alert("접수 구분을 선택하세요.");
        frm.divcd.focus();
        return;
    }

    if (frm.title.value.length<1) {
        alert("제목을 입력하세요.");
        frm.title.focus();
        return;
    }

    if (confirm('수정 하시겠습니까?')){
        frm.mode.value = "editcsas";
        frm.submit();
    }
}

//cs완료처리
function CsRegFinishProc(frm){
    var divcd = frm.divcd.value;

    var confirmMsg ;
    confirmMsg = '완료처리 진행 하시겠습니까?';

	//주문취소
	if(divcd =='A008'){
//		//출고완료가 아닐경우
//		if (IsOrderMasterState != '8'){
//			if (frm.cancelorderno.value.length<1) {
//			    alert("취소한 주문번호를 입력해주세요(마이너스주문)");
//			    frm.cancelorderno.focus();
//			    return;
//			}
//		}
	}

    if (confirm(confirmMsg )){

        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "";
        frm.submit();
    }
}

// 디테일 상품 입력수량 체크
function CheckMaxItemNo(obj, maxno) {
    if (obj.value*1 > maxno*1) {
        alert("주문수량 이상으로 상품수량을 수정할수 없습니다.");
        obj.value = maxno;
    }
}

//고객 회수/맞교환/서비스발송 주소지 변경
function popEditCsDelivery(CsAsID){	
    var window_width = 600;
    var window_height = 450;
    
    var popEditCsDelivery=window.open("/admin/offshop/cscenter/action/pop_CsDeliveryEdit.asp?CsAsID=" + CsAsID ,"popEditCsDelivery","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popEditCsDelivery.focus();
}

//누락재발송
function PopOpenServiceItemOmit(masteridx){
    var mode    = "regcsas";
    var divcd   = "A001";

	if (masteridx == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    PopCSActionCom(masteridx,'',mode,divcd,'','');
}

function changeSongjang(csid){
    var popwin = window.open('/admin/offshop/cscenter/action/popChangeSongjang.asp?masteridx=' + csid,'popChangeSongjang','width=400,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function searchDetail(csmasteridx){
    location.href='/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx='+csmasteridx;
}

// 삭제
function CsRegCancelProc(frm){
    if (confirm('등록된 접수 내역을 삭제 하시겠습니까?')){
        frm.mode.value = "deletecsas";
        frm.submit();
    }
}

// 업체처리완료=>접수 변경
function CsUpcheConfirm2RegProc(frm){
    if (confirm('접수 상태로 변경 하시겠습니까?')){
        frm.mode.value = "upcheconfirm2jupsu";
        frm.submit();
    }
}