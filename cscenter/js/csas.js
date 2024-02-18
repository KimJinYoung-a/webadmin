


// ============================================================================
// CS 사유구분 표시 (AJAX)
// ============================================================================
function divCsAsGubunSelect(value_gubun01,value_gubun02,name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv){
	if (IsFinishProcState) {
	    alert('수정창에서 수정해주세요. - 완료처리시 수정불가');
	    return;
	}

    var params = "?gubun01=" + value_gubun01 + "&gubun02=" + value_gubun02 + "&name_gubun01=" + name_gubun01 + "&name_gubun02=" + name_gubun02 + "&name_gubun01name=" + name_gubun01name + "&name_gubun02name=" + name_gubun02name +"&name_frm=" + name_frm + "&targetDiv=" + targetDiv;
    initializeURL("/cscenter/action/ajax_cs_gubun_select.asp" + params);
    initializeReturnFunction("processAjaxCSGubunSelect(" + targetDiv + ")");
    initializeErrorFunction("onErrorAjaxCSGubunSelect()");
    startRequest();
}

function processAjaxCSGubunSelect(targetDiv) {
    eval(targetDiv).innerHTML = xmlHttp.responseText;
}

function onErrorAjaxCSGubunSelect() {
    alert("데이터를 읽는 도중에 에러가 발생했습니다. 잠시후 다시 시도해보시기 바랍니다.[CODE:" + xmlHttp.status + "]");
}

function colseCausepop(targetDiv){
    eval(targetDiv).innerHTML = "";
}

function delGubun(name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv){
    eval("document." + name_frm + "." + name_gubun01).value = "";
    eval("document." + name_frm + "." + name_gubun02).value = "";
    eval("document." + name_frm + "." + name_gubun01name).value = "";
    eval("document." + name_frm + "." + name_gubun02name).value = "";

    eval(targetDiv).innerHTML = "";
}


function selectGubun(value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_gubun01, name_gubun02, name_gubun01name, name_gubun02name ,name_frm, targetDiv){
	if (IsFinishProcState) {
	    alert('수정창에서 수정해주세요. - 완료처리시 수정불가');
	    return;
	}

    var frm = eval(name_frm);

    eval("document." + name_frm + "." + name_gubun01).value = value_gubun01;
    eval("document." + name_frm + "." + name_gubun02).value = value_gubun02;
    eval("document." + name_frm + "." + name_gubun01name).value = value_gubun01name;
    eval("document." + name_frm + "." + name_gubun02name).value = value_gubun02name;

    eval(targetDiv).innerHTML = "";

    // 전체 사유구분을 입력한 경우 선택된 Detail에 같은 사유 세팅
    if (targetDiv=="causepop"){
        for (var i=0;i<frm.elements.length;i++){
            var e = frm.elements[i];

            if ((e.type=="checkbox")&&(e.checked)&&(e.name=="orderdetailidx")){
                setDetailCause(e.value, value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_frm);
            }
        }
    }

    // 업체개별배송 체크
    CheckUpcheDeliverPay(frm);

    // 배송비 체크
    CheckDeliverPay(frmaction);

	// 금액 재계산
    CalculateAndApplyItemCostSum(frmaction);
}

function setDetailCause(idx, value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_frm) {
    var ogubun01 = eval(name_frm + ".gubun01_" + idx);
    ogubun01.value = value_gubun01;

    var ogubun02 = eval(name_frm + ".gubun02_" + idx);
    ogubun02.value = value_gubun02;

    var ogubun01name = eval(name_frm + ".gubun01name_" + idx);
    ogubun01name.value = value_gubun01name;

    var ogubun02name = eval(name_frm + ".gubun02name_" + idx);
    ogubun02name.value = value_gubun02name;
}



// ============================================================================
// 페이지 새로고침
// ============================================================================
function reloadMe(comp){
    var divcd = comp.value;
    var mode  = Fmode;
    var orderserial = Forderserial;
    document.location = "?mode=" + mode + "&divcd=" + divcd + "&orderserial=" + orderserial;
}



// ============================================================================
// 디테일 상품 입력수량 체크
// ============================================================================
function CheckMaxItemNo(obj, maxno) {
    if (obj.value*1 > maxno*1) {
        alert("주문수량 이상으로 상품수량을 수정할수 없습니다.");
        obj.value = maxno;
    }

	if (IsEditState && IsReturnProcess) {
	    if (obj.value < 0) {
	        alert("0개 미만 으로  수정할수 없습니다.");
	        obj.value = maxno;
	    }
	} else {
	    if (obj.value < 1) {
	        alert("0개 이하로  수정할수 없습니다. 상품선택을 해지해주세요.");
	        obj.value = maxno;
	    }
	}
}

// 상품선택시 확인할 것들
function CheckSelect(comp){
    var chkidx = comp.value;
    var frm = document.frmaction;

    if (comp.name!="Deliverdetailidx"){
        if (comp.checked){
            // CS 사유구분 복사
            eval("frm.gubun01_" + chkidx).value = frm.gubun01.value;
            eval("frm.gubun02_" + chkidx).value = frm.gubun02.value;
            eval("frm.gubun01name_" + chkidx).value = frm.gubun01name.value;
            eval("frm.gubun02name_" + chkidx).value = frm.gubun02name.value;
        }else{
            delGubun("gubun01_" + chkidx,"gubun02_" + chkidx,"gubun01name_" + chkidx,"gubun02name_" + chkidx,frm.name,causepop);
        }
    }

    // 업체개별배송 체크
    CheckUpcheDeliverPay(frm);

    // 배송비 체크
    CheckDeliverPay(frm);

	// 금액 재계산
    CalculateAndApplyItemCostSum(frm);

    // 선택 상품의 관련 브랜드 검색
    DispCheckedUpcheID(frmaction);
}



// ============================================================================
// 업체배송비 합계
function GetCheckedUpcheBeasongPay(frm){
    var retVal = 0;
    var e;

    if (!frm.Deliverdetailidx) return retVal;

    for (var i = 0; i < frm.Deliverdetailidx.length; i++) {
        e = frm.Deliverdetailidx[i];

        if ((e.type == "checkbox") && (e.checked)) {
            retVal = retVal + (frm.Deliveritemcost[i].value * 1);
        }
    }

    return retVal;
}

// 업체 배송비 체크
// 텐배   - 텐배이고, 브랜드 전체상품이 체크된경우 브랜드 없는 배송비 전부 체크한다.
// 브랜드 - 업배이고, 브랜드 전체상품이 체크된경우 브랜드 배송비도 체크한다.
// 단순변심인경우 체크하지 않는다.
function CheckUpcheDeliverPay(frm){
    var upbeaMakerid;
    var itemMakerid;

    var NotCheckExists;				// 체크안된 상품이 있는가
    var isCheckValExists;			// 한개라도 선택된 상품이 있는가

    var value_gubun02 = frm.gubun02.value;
    var objdeliver, objitem

    if (!frm.Deliverdetailidx) return;
    if (!frm.orderdetailidx) return;

    if ((!IsCancelProcess)&&(!IsReturnProcess)) return;

    for (var i = 0; i < frm.Deliverdetailidx.length; i++) {
        objdeliver = frm.Deliverdetailidx[i];

        if (objdeliver.type == "checkbox") {
            NotCheckExists=false;
            isCheckValExists=false;

            upbeaMakerid = frm.DeliverMakerid[i].value;

            if (upbeaMakerid.length < 1) {
            	// 배송비 브랜드가 없으면 : 텐텐배송
                for (var j = 0; j < frm.orderdetailidx.length; j++) {
                    objitem = frm.orderdetailidx[i];

                    if (objitem.type == "checkbox") {
	                    if ((frm.odlvtype[j].value == "1") || (frm.odlvtype[j].value == "4")) {
	                        // 텐텐배송이면
	                        isCheckValExists = true;
	                        NotCheckExists = (NotCheckExists) || (!frm.orderdetailidx[j].checked);
	                        NotCheckExists = (NotCheckExists) || ((frm.orderdetailidx[j].checked == true) && (frm.itemno[j].value != frm.regitemno[j].value));
	                    }
                    }
                }
            } else {
            	// 배송비 브랜드가 있으면 : 업체배송
                for (var j = 0; j < frm.orderdetailidx.length; j++) {
                    objitem = frm.orderdetailidx[j];

                    if (objitem.type == "checkbox") {

                        itemMakerid = frm.makerid[j].value;
                        if (upbeaMakerid == itemMakerid) {
                        	isCheckValExists = true;
                            NotCheckExists = (NotCheckExists)||(!frm.orderdetailidx[j].checked);
                            NotCheckExists = (NotCheckExists)||((frm.orderdetailidx[j].checked == true) && (frm.itemno[j].value != frm.regitemno[j].value));
                        }
                    }
                }
            }

            // 전부가 선택된 경우 배송비도 같이 체크한다.(브랜드 전상품 취소)
            frm.Deliverdetailidx[i].checked = ((!NotCheckExists) && (isCheckValExists));

            //반품 프로세스고 단순변심인 경우 체크해제하지 않는다.
            if ((IsReturnProcess) && ((value_gubun02 == "CD01")||(value_gubun02 == "CD06"))) {
                frm.Deliverdetailidx[i].checked = false;
            }
            AnCheckClick(frm.Deliverdetailidx[i]);
        }
    }
}

//배송비 체크
function CheckDeliverPay(frm){
    var allselected = IsAllSelected(frm);
    var brandallselected = IsBrandAllSelected(frm);

    var value_gubun02 = frm.gubun02.value;

    if (IsCancelProcess) {
        // 취소 Process

        if (allselected) {
            // 마일리지, 쿠폰 환원
            frm.milereturn.checked = true;
            frm.couponreturn.checked = true;
        } else {
            frm.milereturn.checked = false;
            frm.couponreturn.checked = false;
        }

    }else if (IsReturnProcess) {
    	// 반품 Process

        // 회수배송비 관련
        if ((value_gubun02=="CD01")||(value_gubun02=="CD06")){
        	// 단순변심
            if (frm.divcd.value=="A010"){
            	// 회수신청(텐바이텐배송)
            	if (brandallselected == true) {
	                frm.ckreturnpay.checked = true;
	                frm.ckreturnpayHalf.checked = false;
            	} else {
	                frm.ckreturnpay.checked = false;
	                frm.ckreturnpayHalf.checked = true;
            	}
            }else{
                frm.ckreturnpay.checked = false;
                frm.ckreturnpayHalf.checked = false;
            }
        }else{
            frm.ckreturnpay.checked = false;
            frm.ckreturnpayHalf.checked = false;
        }

    }
}

function ChangeReturnMethod(comp){
    if (comp==undefined) return;

    var returnmethod = comp.value;

    document.all.refundinfo_R007.style.display = "none";
    //document.all.refundinfo_R020.style.display = "none";
    document.all.refundinfo_R050.style.display = "none";
    //document.all.refundinfo_R080.style.display = "none";
    document.all.refundinfo_R100.style.display = "none";
    document.all.refundinfo_R900.style.display = "none";

    if (comp.value=="R007"){
        //무통장 환불
        document.all.refundinfo_R007.style.display = "inline";
    }else if((comp.value=="R020")||(comp.value=="R080")||(comp.value=="R100")||(comp.value=="R400")){
        //실시간 이체 취소//ALL@ 결제 취소 //신용카드 결제 취소//휴대폰
        document.all.refundinfo_R100.style.display = "inline";
    }else if(comp.value=="R050"){
        //입점몰 결제 취소
        document.all.refundinfo_R050.style.display = "inline";
    }else if(comp.value=="R900"){
        //마일리지 환불
        document.all.refundinfo_R900.style.display = "inline";
    }

}

//CS 접수
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

    if (frm.gubun01.value.length<1) {
        alert("사유 구분을 입력하세요.");
        return;
    }

    //선택 상품 체크
    if (!SaveCheckedItemList(frm)) {
        return;
    }

    if (IsReturnProcess){
        if (!checkReturnProcessAvail(frm)){
            return;
        }
    }

    if (IsServiceDeliverProcess){
        if (!checkReturnProcessAvail(frm)){
            return;
        }
    }

    if ((Fdivcd=="A009")||(Fdivcd=="A006")){
        if (!checkReturnProcessAvail(frm)){
            return;
        }
    }

    if ((IsCancelProcess)||(IsReturnProcess)){
        if (frm.subtotalprice!=undefined){


			if (frm.returnmethod){

				// 환불 없음은 가능하다.
				if (frm.returnmethod.value != "R000") {
		            if (frm.subtotalprice.value*1 < frm.canceltotal.value*1) {
		                alert('환불 예정액이 결제금액 보다 클 수 없습니다. - 쿠폰이나 마일리지 환급을 체크해 주세요.');

		                if (IsAdminLogin != true) {
		                	return;
		                }
		            }

		            if (frm.canceltotal.value*1<1){
		                alert('환불예정액이 없습니다. 환불없음을 선택하세요.');

		                if (IsAdminLogin != true) {
		                	return;
		                }
		            }
				}
            }

            if (frm.nextsubtotal.value*1<0){
                alert('취소 후 결제 금액이 마이너스가 될 수 없습니다. - 쿠폰이나 마일리지 환급을 체크해 주세요.');

                if (IsAdminLogin != true) {
                	return;
                }
            }

            if (frm.canceltotal.value*1<0){
                alert('실 취소 금액이 0보다 작을 수 없습니다. - 쿠폰이나 마일리지 환급을 UnCheck 주세요.');

                if (IsAdminLogin != true) {
                	return;
                }
            }

            //returnmethod R400 인경우 당월 취소만 가능함
            if (frm.returnmethod){
	            if (frm.returnmethod.value=="R400"){
			    	if (IsOrderFound && (IsThisMonthJumun != true)) {
				        alert('휴대폰 결제는 당월 취소만 가능합니다. 다른 환불방식을 선택해 주세요.');
				        frm.returnmethod.focus();
				        return;
			    	}
	            }
            }
        }
    }

    if ((frm.subtotalprice!=undefined)&&(frm.returnmethod!=undefined)){
        if (!CheckReturnMethod(frm)){
            return;
        }
    }

    //추가 정산 관련
    if (frm.add_upchejungsandeliverypay){
        if (!IsInteger(frm.add_upchejungsandeliverypay.value)){
            alert('숫자만 가능합니다.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*1!=0){
            if (frm.add_upchejungsancause.value.length<1){
                alert('추가 정산 사유를 입력하세요.');
                frm.add_upchejungsancause.focus();
                return;
            }else if ((frm.add_upchejungsancause.value=='직접입력')&&(frm.add_upchejungsancauseText.value.length<1)){
                alert('추가 정산 사유를 입력하세요.');
                frm.add_upchejungsancauseText.focus();
                return;
            }

            if (frm.buf_requiremakerid.value.length<1){
                alert('추가 정산액이 있는경우 브랜드 아이디가 지정되어야 합니다. ');
                return;
            }

            //주문 내역에 아이디가 있는 경우만.

        }else{

			if (Fdivcd == "A700") {
	            alert('추가 정산액을 입력하세요.');
	            frm.add_upchejungsandeliverypay.focus();
	            return;
			}
        }
    }

    if (IsCancelProcess){
        if(confirm("취소 접수 하시겠습니까?")){
            frm.submit();
        }
    }else if (IsReturnProcess){
        if (frm.ForceReturnByTen.checked){
            frm.requireupche.value = "N";
            frm.requiremakerid.value = "";
        }

        if (frm.requireupche.value=="Y"){
            if(confirm("업체 [" + frm.requiremakerid.value +"]로 반품/회수/교환 접수 하시겠습니까?")){
                frm.submit();
            }
        }else{
            if(confirm("[텐바이텐 물류센터]로 반품/회수/교환 접수 하시겠습니까?")){
                frm.submit();
            }
        }
    }else if (IsRefundProcess){
        if (frm.returnmethod.value.length<1){
            alert('환불 방식을 선택해 주세요.');
            frm.returnmethod.focus();
            return;
        }

        if (frm.returnmethod.value=="R000"){
            frm.refundrequire.value = "0";

        }


        if ((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R080")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R400")){
            alert('신용카드/실시간/올엣/휴대폰 환불 접수시 원 결제 금액 그대로 접수됩니다.');
            frm.refundrequire.value = frm.orgsubtotalprice.value;
        }

        if ((frm.refundrequire.value.length<1)||(!IsDigit(frm.refundrequire.value))){
            alert('환불 금액을 입력하세요.');
            frm.refundrequire.focus();
            return;
        }

        if(confirm("환불 접수 하시겠습니까?")){
            frm.submit();
        }
    }else if(confirm("접수 하시겠습니까?")){
        frm.submit();
    }
}

//업체처리완료=>접수 변경
function CsUpcheConfirm2RegProc(frm){
    if (confirm('접수 상태로 변경 하시겠습니까?')){
        frm.mode.value = "upcheconfirm2jupsu";
        frm.submit();
    }
}

//수정
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

    if (frm.gubun01.value.length<1) {
        alert("사유 구분을 입력하세요.");
        return;
    }

	//반품 관련 수정일경우 체크함
	if (IsEditState && IsReturnProcess) {
	    if (!SaveCheckedItemList(frm)) {
	        return;
	    }
	}

    if ((frm.returnmethod!=undefined)){
        if (!CheckReturnForm(frm)){
            return;
        }
    }

    //추가 정산 관련
    if (frm.add_upchejungsandeliverypay){
        if (!IsDigit(frm.add_upchejungsandeliverypay.value)){
            alert('숫자만 가능합니다.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*1!=0){
            if (frm.add_upchejungsancause.value.length<1){
                alert('추가 정산 사유를 입력하세요.');
                frm.add_upchejungsancause.focus();
                return;
            }else if ((frm.add_upchejungsancause.value=='직접입력')&&(frm.add_upchejungsancauseText.value.length<1)){
                alert('추가 정산 사유를 입력하세요.');
                frm.add_upchejungsancauseText.focus();
                return;
            }

            if (frm.buf_requiremakerid.value.length<1){
                alert('추가 정산액이 있는경우 브랜드 아이디가 지정되어야 합니다. ');
                return;
            }
        }
    }

    //returnmethod R400 인경우 당월 취소만 가능함
    if (frm.returnmethod){
	    if (frm.returnmethod.value=="R400"){

	    	if (IsOrderFound && (IsThisMonthJumun != true)) {
		        alert('휴대폰 결제는 당월 취소만 가능합니다. 다른 환불방식을 선택해 주세요.');
		        frm.returnmethod.focus();
		        return;
	    	}
	    }
    }

    if (confirm('수정 하시겠습니까?')){
        frm.mode.value = "editcsas";
        frm.submit();
    }
}

// 업체 추가정산 관련 브랜드 아이디 가져오기
function DispCheckedUpcheID(frm) {
    var checkedUpcheid = "";
    var UpcheDuplicated = false;
    var IsUpcheReturn;

	if ((Fdivcd == "A004") || (Fdivcd == "A000")) { // 반품접수(업체배송), 맞교환출고
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

    if ((!UpcheDuplicated)&&(checkedUpcheid!="")){
        if (frm.buf_requiremakerid){
            frm.buf_requiremakerid.value = checkedUpcheid;
        }
    }
}


//회수관련 체크
function checkReturnProcessAvail(frm){
    var TenBeasongExists = false;
    var UpcheBeasongExists = false;
    var UpcheDuplicated = false;
    var checkedUpcheid = "";

	// 서비스 발송시 브랜드를 지정하게..=> 상품을 선택하는것이 아니므로.
    for(var i = 0; i < frm.orderdetailidx.length; i++) {
    	if (frm.orderdetailidx[i].type != "checkbox") {
    		continue;
    	}

        if (frm.orderdetailidx[i].checked){
            if (frm.isupchebeasong[i].value=="Y"){
                UpcheBeasongExists = true;

                if (checkedUpcheid!="") {
                    if (checkedUpcheid != frm.makerid[i].value){
                        UpcheDuplicated = true;
                    }

                    checkedUpcheid = frm.makerid[i].value;
                }
                checkedUpcheid = frm.makerid[i].value;
            }else{
                TenBeasongExists   = true;
            }
        }
    }

    if ((UpcheBeasongExists)&&(TenBeasongExists)){
        alert('텐바이텐 배송과 업체배송을 동시에 선택 하실 수 없습니다.');
        return false;
    }

    if (UpcheDuplicated){
        alert('업체배송 선택시 브랜드 별로 접수해 주세요.');
        return false;
    }

    if (checkedUpcheid!=""){
        frm.requireupche.value = "Y";
        frm.requiremakerid.value = checkedUpcheid;

        if (frm.buf_requiremakerid){
            frm.buf_requiremakerid.value = frm.requiremakerid.value;
        }
    }else{
        frm.requireupche.value = "N";
        frm.requiremakerid.value = "";

        if (frm.buf_requiremakerid){
            frm.buf_requiremakerid.value = "";
        }
    }
    return true;
}

//환불 관련 체크 form
function CheckReturnForm(frm){
    if (frm.returnmethod){
	    if (frm.returnmethod.value=="R007"){
	        var mooconfirm = false;
	        if ((frm.rebankaccount!=undefined)&&(frm.rebankaccount.value.length<1)){
	            //alert('환불 계좌를 입력해 주세요.');
	            //frm.rebankaccount.focus();
	            mooconfirm=true;
	            //return false;
	        }

	        if ((frm.rebankownername!=undefined)&&(frm.rebankownername.value.length<1)){
	            //alert('예금주명을  입력해 주세요.');
	            //frm.rebankownername.focus();
	            mooconfirm=true;
	            //return false;
	        }

	        if ((frm.rebankname!=undefined)&&(frm.rebankname.value.length<1)){
	            //alert('환불 은행을 선택해 주세요.');
	            //frm.rebankname.focus();
	            mooconfirm=true;
	            //return false;
	        }

	        if ((frm.refundrequire!=undefined)&&(frm.refundrequire.value.length<1)){
	            alert('환불 금액을 재 계산하세요');
	            return false;
	        }

	        if (mooconfirm){
	            if (!confirm('환불 계좌가 없습니다. \n\n환불 계좌 없이 등록 하시겠습니까?')){
	                frm.rebankaccount.focus();
	                return false;
	            }
	        }

    }else if (frm.returnmethod.value=="R900"){
	        if ((frm.refundbymile_userid!=undefined)&&(frm.refundbymile_userid.value.length<1)){
	            alert('사용자 아이디가 없습니다. 다른 환불 방식을 선택하세요.');
	            return false;
	        }

	        if ((frm.refundbymile_sum!=undefined)&&(frm.refundbymile_sum.value.length<1)){
	            alert('환불 마일리지가 없습니다. 재계산 해주세요.');
	            return false;
	        }

	    }

	    if ((frm.returnmethod.value!="")&&(frm.returnmethod.value!="R000")&&(!IsDigit(frm.refundrequire.value))){
	        alert('환불 금액은 양수(+) 만 가능합니다.');
	        return false;
	    }
    }
    return true;
}

//환불 관련 체크
function CheckReturnMethod(frm){
    var allselected = IsAllSelected(frm);

    var PayedNCancelEqual = (frm.subtotalprice.value*1==frm.canceltotal.value*1);

    if ((allselected)&&(!PayedNCancelEqual)&&(IsCancelProcess)){
        alert('전체 취소인경우 결제금액 전체를 환불해야합니다. - 원배송비 환급, 마일리지, 할인권 등을 체크해주세요.');
        return false;
    }

    if (((!PayedNCancelEqual))&&((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R080")||(frm.returnmethod.value=="R400"))){
        alert('전체 취소인 경우만 신용카드/실시간 이체/휴대폰 취소가 가능합니다. \n\n무통장 환불 또는 마일리지 환급을 선택해 주세요');
        frm.returnmethod.focus();
        return false;
    }

    if (frm.returnmethod.value.length<1){
        alert('환불 방식을 선택해 주세요.');
        frm.returnmethod.focus();
        return false;
    }

    if (!CheckReturnForm(frm)){
        return false;
    }

	if (FSiteName != "10x10") {
	    //외부몰인 경우 외부몰 환불접수만 가능..
	    if ((frm.returnmethod.value != "R050") && (frm.returnmethod.value != "R000")) {
	        alert('외부몰인 경우 환불 없음 또는 외부몰 환불을 선택하세요. \n\n제휴 담당자를 통해 제휴몰에서 취소 환불 처리 합니다.');
	        frm.returnmethod.focus();
	        return
	    }
	}

    if (frm.refundrequire.value!=frm.canceltotal.value){
        if ((frm.returnmethod.value!="R007")&&(frm.returnmethod.value!="R900")&&(frm.returnmethod.value!="R000")){
            alert('환불 금액과 취소금액이 다를경우 무통장 또는 마일리지 환불만 가능합니다.');
            return false;
        }

        if (!confirm('환불 금액이 취소 금액과 다르게 수정될경우 보정치로 금액이 입력됩니다.\n\n진행 하시겠습니까?')){
            return false;
        }
    }

    return true;
}


//삭제
function CsRegCancelProc(frm){
    if (confirm('등록된 접수 내역을 삭제 하시겠습니까?')){
        frm.mode.value = "deletecsas";
        frm.submit();
    }
}

//접수상태로 변경
function CsRegStateChg(frm){
    if (confirm('접수 상태로 변경 하시겠습니까?')){
        frm.mode.value = "state2jupsu";
        frm.submit();
    }
}

//완료처리
function CsRegFinishProc(frm){
    var divcd = frm.divcd.value;

    //환불요청 , 신용카드 취소요청
    if ((divcd=="A003")||(divcd=="A007")){
        if (frm.contents_finish.value.length<1){
            alert('처리 내용을 입력하세요.');
            frm.contents_finish.focus();
            return;
        }
    }

    var confirmMsg ;
    confirmMsg = '완료처리 진행 하시겠습니까?';

    if ((divcd=="A004")||(divcd=="A010")){
        confirmMsg = '완료처리 진행시 마이너스 주문 및 환불이 자동 접수됩니다. 진행 하시겠습니까?';
    }

    //20090601추가
    if (divcd="A003"){
        if (!CheckReturnForm(frm)){
            return;
        }
    }

    if (confirm(confirmMsg )){

        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "";
        frm.submit();
    }
}

function CsRegFinishProcNoRefund(frm){
    var divcd = frm.divcd.value;

    if (confirm('환불 및 마이너스 등록 없이 완료처리 진행 하시겠습니까?')){
        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "norefund";
        frm.submit();
    }
}

//상품 전체 체크 되었는지
function IsAllSelected(frm){
    var allselected = false;

	if (frm.orderdetailidx) {
	    for (var i = 0; i < frm.orderdetailidx.length; i++) {
	        e = frm.orderdetailidx[i];

	        if (e.type=="checkbox") {
	            if (e.checked == true) {
	                    allselected = true;
	            } else {
	                    return false;
	            }

	            if (frm.regitemno[i].value != frm.itemno[i].value) {
	                return false;
	            }
	        }
	    }
	}

    return allselected;
}

// 브랜드 상품 전체 체크 되었는지
// 업배는 해당 브랜드 전체
// 텐배의 경우 텐배 전체
function IsBrandAllSelected(frm){
    var brandallselected = false;
    var makerid = "";
	var isupchebeasong = "";

	if (frm.orderdetailidx) {
	    for (var i = 0; i < frm.orderdetailidx.length; i++) {
	        e = frm.orderdetailidx[i];

	        if (e.type == "checkbox") {
	        	if (e.checked == true) {
	        		if (frm.isupchebeasong[i].value == "Y") {
	        			makerid = frm.makerid[i].value;
	        		} else {
	        			makerid = "";
	        		}

		            if (IsOneBrandAllSelected(frm, makerid) != true) {
		            	return false;
		            }
		            brandallselected = true;
	            }
	        }
	    }
	}

    return brandallselected;
}

// 한개 브랜드 전체 선택되었는지 체크
// 브랜드가 빈값이면 텐텐배송 전체체크
function IsOneBrandAllSelected(frm, makerid) {
    var onebrandallselected = false;

	if (frm.orderdetailidx) {
	    for (var i = 0; i < frm.orderdetailidx.length; i++) {
	        e = frm.orderdetailidx[i];

	        if (e.type == "checkbox") {
	        	if (makerid == "") {
	        		if (frm.isupchebeasong[i].value != "Y") {
	        			if (e.checked != true) {
	        				return false;
	        			}

			            if (frm.regitemno[i].value != frm.itemno[i].value) {
			                return false;
			            }
	        		} else {
	        			onebrandallselected = true;
	        		}
	        	} else {
	        		if ((frm.isupchebeasong[i].value == "Y") && (frm.makerid[i].value == makerid)) {
	        			if (e.checked != true) {
	        				return false;
	        			}

			            if (frm.regitemno[i].value != frm.itemno[i].value) {
			                return false;
			            }
	        		} else {
	        			onebrandallselected = true;
	        		}
	        	}
	        }
	    }
	}

    return onebrandallselected;
}

function SaveCheckedItemList(frm) {
    var e;
    var ischecked = false;
    var checkitemExists = false;

    var orderdetailidx = "";
    var gubun01 = "";
    var gubun02 = "";
    var regitemno = "";
    var causecontent = "";

    frm.detailitemlist.value = "";

    for (var i = 0; i < frm.length; i++) {
        e = frm.elements[i];

        if (e.name == "dummystarter") {
            ischecked = false;
            orderdetailidx = "";
            gubun01 = "";
            gubun02 = "";
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

        if ((ischecked == true) && (e.name.indexOf("gubun01_") == 0)) {
            gubun01 = e.value;
        }

        if ((ischecked == true) && (e.name.indexOf("gubun02_") == 0)) {
            gubun02 = e.value;
        }

        if ((ischecked == true) && (e.name.indexOf("regitemno") == 0)) {
			if (IsEditState && IsReturnProcess) {
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

        // 사용안함.
        //if ((ischecked == true) && (e.name.indexOf("causecontent") == 0)) {
        //        causecontent = e.value;
        //}

        if (e.name == "dummystopper") {
            if (ischecked == true) {
                frm.detailitemlist.value = frm.detailitemlist.value + "|" + orderdetailidx + "\t" + gubun01 + "\t" + gubun02 + "\t" + regitemno + "\t" + causecontent;
                ischecked = false;
                gubun01 = "";
                gubun02 = "";
                regitemno = "";
                causecontent = "";
            }
        }
    }

    //배송비 ----------------------------------------
    var upchedeliverPayStr = '';
    if (frm.Deliverdetailidx){

        for (var i = 0; i < frm.Deliverdetailidx.length; i++) {
        	e = frm.Deliverdetailidx[i];

            if ((e.type=="checkbox") && (e.checked)) {
                upchedeliverPayStr = e.value + "\t" + frm.gubun01.value + "\t" + frm.gubun02.value + "\t" + "1" + "\t" + frm.Deliveritemcost[i].value;
            }
        }
    }

    if ((upchedeliverPayStr.length>0)&&(frm.detailitemlist.value.length>0)){
        frm.detailitemlist.value = frm.detailitemlist.value + "|" + upchedeliverPayStr
    }
    //--------------------------------------------------

    //기타내역, 서비스발송 , 환불요청, 출고시유의사항, 업체 추가 정산 - 상세내역 체크 안함.
    if ((Fdivcd=="A009")||(Fdivcd=="A002")||(Fdivcd=="A003")||(Fdivcd=="A005")||(Fdivcd=="A006")||(Fdivcd=="A700")){
        // no- check

    }else{
        if (!checkitemExists){
            alert('선택된 상세내역이 없습니다.');
            return false;
        }
    }

    return true;
}

//업체 추가 정산 삭제시
function clearAddUpchejungsan(frm){
    frm.add_upchejungsandeliverypay.value = "0";
    frm.add_upchejungsancause.value = "";
    frm.add_upchejungsancauseText.value = "";
}

function setTitle(frm,titlestr){
    frm.title.value=titlestr;
}

// 이 펑션 사용 안함. 사용 하지 말것.
function getReCalcuBeasongPay(userlevel, itemsum){
    //Yellow : 5만원이상, 그린회원 : 4만원이상, 블루회원 : 3만원이상, vip회원 : 항상 무료 , 마니아 2만원이상

    //무료배송인경우 체크. 추가 해야함

    if (userlevel=="0"){
        if ((itemsum>=50000)||(itemsum<1)){
            return 0;
        }else{
            return 2000;
        }
    }else if (userlevel=="1"){
        if ((itemsum>=40000)||(itemsum<1)){
            return 0;
        }else{
            return 2000;
        }
    }else if (userlevel=="2"){
        if ((itemsum>=30000)||(itemsum<1)){
            return 0;
        }else{
            return 2000;
        }
    }else if (userlevel=="9"){
        if ((itemsum>=20000)||(itemsum<1)){
            return 0;
        }else{
            return 2000;
        }
    }else if (userlevel=="3"){
        return 0;
    }else{
        return 2000;
    }
}

function CalculateAndApplyItemCostSum(frm) {

    var e;
    var ischecked       = false;
    var regitemno       = 0;
    var itemno          = 0;
    var itemcost        = 0;
    var refunditemcostsum   = 0;
    var allatitemdiscount    = 0;
    var allatitemdiscountSum = 0;

    var percentBonusCouponDiscount =0;
    var percentBonusCouponDiscountSum =0;

    var orgitemcostsum     = 0;
    var orgbeasongpay       = 0;
    var refundadjustpay     = 0;

    var refundmileagesum    = 0;
    var refundcouponsum     = 0;
    var allatsubtractsum    = 0;

    var remainitemcostsum   = 0;
    var remainmileagesum    = 0;
    var remaincouponsum     = 0;
    var remainallatdiscount = 0;
    var refunddeliverypay   = 0;

    var recalcubeasongpay   = 0;
    var refundbeasongpay    = 0;

    //업체배송비 환급 추가
    var refundupchebeasongpay = GetCheckedUpcheBeasongPay(frm);

    for (var i = 0; i < frm.length; i++) {
        e = frm.elements[i];

        if (e.name == "dummystarter") {
            ischecked = false;
            regitemno = 0;
            itemno = 0;
            itemcost = 0;
        }

        if (e.name == "orderdetailidx") {
	        if (e.type != "checkbox") {
	        	ischecked = false;
	            continue;
	        }

            if (e.checked == true) {
                ischecked = true;
            }
        }

        if ((ischecked == true) && (e.name == "regitemno")) {
            if ((e.value * 0) == 0) {
                regitemno = e.value;
            } else {
                regitemno = 0;
            }
        }

        if ((ischecked == true) && (e.name == "itemno")) {
            if ((e.value * 0) == 0) {
                itemno = e.value;
            } else {
                itemno = 0;
            }
        }

        if ((ischecked == true) && (e.name == "itemcost")) {
            if ((e.value * 0) == 0) {
                itemcost = e.value;
            } else {
                itemcost = 0;
            }
        }

        if ((ischecked == true) && (e.name == "allatitemdiscount")) {
            if ((e.value * 0) == 0) {
                allatitemdiscount = e.value;
            } else {
                allatitemdiscount = 0;
            }
        }

        if ((ischecked == true) && (e.name == "percentBonusCouponDiscount")) {
            if ((e.value * 0) == 0) {
                percentBonusCouponDiscount = e.value;
            } else {
                percentBonusCouponDiscount = 0;
            }
        }

        if (e.name == "dummystopper") {
            if (ischecked == true) {
                refunditemcostsum = refunditemcostsum + (itemcost * regitemno * 1);
                allatitemdiscountSum = allatitemdiscountSum + (allatitemdiscount * regitemno * 1);
                percentBonusCouponDiscountSum = percentBonusCouponDiscountSum + (percentBonusCouponDiscount * regitemno * 1);
            }

            ischecked = false;
            regitemno = 0;
            itemno = 0;
            itemcost = 0;
        }
    }

    // 원래 상품 합계 금액
    if (frm.orgitemcostsum!=undefined){
        orgitemcostsum = frm.orgitemcostsum.value*1;
    }

    // 상품목록 하단 선택 합계
    if (frm.itemcanceltotal!=undefined){
         frm.itemcanceltotal.value = refunditemcostsum;
    }

	// 등록 및 수정 시에만 재계산.
	if (IsRegisterState || IsRefundInfoFound) {
	    // 취소/반품 상품총액
	    if (frm.refunditemcostsum!=undefined){
	        frm.refunditemcostsum.value = refunditemcostsum;
	    }

	    // 선택안한(나머지) 상품총액
	    if (frm.remainitemcostsum!=undefined){
	        frm.remainitemcostsum.value = orgitemcostsum - refunditemcostsum;

	        remainitemcostsum = frm.remainitemcostsum.value;
	    }

	    // 사용 마일리지 환원
	    if (frm.milereturn!=undefined){
	        if (frm.milereturn.checked){

				// 환원마일리지가 취소금액보다 큰경우
	            if ((frm.itemcanceltotal.value*1) < (frm.miletotalprice.value*1)) {
	            	frm.refundmileagesum.value = frm.itemcanceltotal.value*-1;
	            } else {
	            	frm.refundmileagesum.value = frm.miletotalprice.value*-1;
	            }

	        }else{
	            frm.refundmileagesum.value = 0;
	        }
	        frm.remainmileagesum.value = (frm.miletotalprice.value*1 + frm.refundmileagesum.value*1)*-1;

	        refundmileagesum = frm.refundmileagesum.value;
	        remainmileagesum = frm.remainmileagesum.value;
	    }

	    // 사용 할인권 환원
	    if (frm.couponreturn!=undefined){
	        if (frm.couponreturn.checked){
	            frm.refundcouponsum.value = frm.tencardspend.value*-1;
	        }else{
	            frm.refundcouponsum.value = 0;

	            // % 할인권 사용시.
	            if (percentBonusCouponDiscountSum!=0){
	                frm.refundcouponsum.value = percentBonusCouponDiscountSum*-1;

	                if (percentBonusCouponDiscountSum*-1==frm.tencardspend.value*-1){
	                    frm.couponreturn.checked = true;
	                }
	            }

	        }
	        frm.remaincouponsum.value = (frm.tencardspend.value*1 + frm.refundcouponsum.value*1)*-1;

	        refundcouponsum = frm.refundcouponsum.value;
	        remaincouponsum = frm.remaincouponsum.value;
	    }

	    // 카드 할인 차감
	    if (frm.allatsubtractsum){
            frm.allatsubtractsum.value = 0;

            // 카드 할인 사용시 200906추가
            if (allatitemdiscountSum!=0){
                frm.allatsubtractsum.value = allatitemdiscountSum*-1;

                if (allatitemdiscountSum*-1==frm.allatdiscountprice.value*-1){
                    //frm.allatsubtract.checked = true;
                }
            }

	        frm.remainallatdiscount.value = (frm.allatdiscountprice.value*1 + frm.allatsubtractsum.value*1)*-1;

	        allatsubtractsum    = frm.allatsubtractsum.value ;
	        remainallatdiscount = frm.remainallatdiscount.value ;
	    }

	    // 기타보정금액
	    if (frm.refundadjustpay!=undefined){
	        refundadjustpay = frm.refundadjustpay.value*1;
	    }

	    //원 배송비
	    if (frm.orgbeasongpay!=undefined){
	        orgbeasongpay  = frm.orgbeasongpay.value*1;
	    }

	    // 취소프로세스시 원 배송비 처리
	    if (IsCancelProcess) {
	        if (frm.ckbeasongpayAssign.checked){
	            refundbeasongpay  = orgbeasongpay;
	            recalcubeasongpay = 0;
	        }else{
	            refundbeasongpay  = refundupchebeasongpay;
	            recalcubeasongpay = orgbeasongpay-refundupchebeasongpay;
	        }
	        refundbeasongpay  = orgbeasongpay - recalcubeasongpay;
	        frm.recalcubeasongpay.value = recalcubeasongpay;
	    }

	    // 반품프로세스시 원 배송비 처리
	    if (IsReturnProcess){

	        if (frm.ckbeasongpayAssign.checked){
	            refundbeasongpay  = orgbeasongpay;
	            recalcubeasongpay = 0;
	            frm.recalcubeasongpay.value = 0;
	        }else{
	            recalcubeasongpay = 0;
	            refundbeasongpay  = refundupchebeasongpay;
	            frm.recalcubeasongpay.value = refundupchebeasongpay;
	        }
	    }

	    //취소 배송비 = refundbeasongpay - recalcubeasongpay
	    if (frm.refundbeasongpay!=undefined){
	        frm.refundbeasongpay.value = refundbeasongpay
	    }

	    // 회수 배송비 -
	    if (frm.ckreturnpay!=undefined){
	        if (frm.ckreturnpayHalf.checked){
	            refunddeliverypay = CDEFAULTBEASONGPAY*-1;

	        }else if (frm.ckreturnpay.checked){
	            refunddeliverypay = CDEFAULTBEASONGPAY*2*-1;

	        }else{
	            refunddeliverypay = 0;
	        }
	    }

	    if (frm.refunddeliverypay!=undefined){
	        frm.refunddeliverypay.value = refunddeliverypay*1;
	    }

	    if (frm.buf_refunddeliverypay!=undefined){
	        frm.buf_refunddeliverypay.value = refunddeliverypay*-1;
	        frm.buf_totupchejungsandeliverypay.value = frm.buf_refunddeliverypay.value*1 + frm.add_upchejungsandeliverypay.value*1;
	    }

	    //취소금액 합계
	    if (frm.canceltotal!=undefined){
	        frm.canceltotal.value  = refunditemcostsum + refundmileagesum*1 + refundcouponsum*1 + allatsubtractsum*1 + refundbeasongpay*1 + refundadjustpay*1 + refunddeliverypay*1;
	    }

	    //취소후 금액 합계
	    if (frm.nextsubtotal!=undefined){
	        frm.nextsubtotal.value = remainitemcostsum*1 + remainmileagesum*1 + remaincouponsum*1 + remainallatdiscount*1 + recalcubeasongpay*1 ;
	    }

	    if (parseInt(frm.ipkumdiv.value) >= 4) {
	        if ((IsCancelProcess)||(IsReturnProcess)) {
	            if (frm.refundrequire!=undefined){
	                frm.refundrequire.value = frm.canceltotal.value*1;
	            }
	        }
	    }
	}
}

function CheckDoubleCheck(frm,comp){
    if (comp.name=="ckreturnpay"){
        if (frm.ckreturnpay.checked){
            frm.ckreturnpayHalf.checked = false;
        }
    }else if (comp.name=="ckreturnpayHalf"){
        if (frm.ckreturnpayHalf.checked){
            frm.ckreturnpay.checked = false;
        }
    }
}

function ShowOLDCSList(){

}

//불량상품등록
function popBadItemReg(barcode,itemcount){
    var popwin = window.open('/common/do_bad_item_input.asp?mode=insert&itemcount=' + itemcount + '&itemid=' + barcode,'popBadItemReg','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function ChangeColor(a,b,c){
    //Nothing
}
function searchDetail(a){
    //Nothing
}

//추가정산배송비
function Change_add_upchejungsandeliverypay(comp){
    comp.form.buf_totupchejungsandeliverypay.value = comp.form.buf_refunddeliverypay.value*1 + comp.value*1;

    if (isNaN(comp.form.buf_totupchejungsandeliverypay.value)){
        comp.form.buf_totupchejungsandeliverypay.value = "0";
    }
}

//추가정산배송비 사유
function Change_add_upchejungsancause(comp){
    if (comp.value=="직접입력") {
        document.all.span_add_upchejungsancauseText.style.display = "inline";
    }else{
        document.all.span_add_upchejungsancauseText.style.display = "none";
    }
}

// 배송추적 이력
function popDeliveryTrace(traceUrl, songjangNo)
{
	var f = document.popForm;
	f.traceUrl.value	= traceUrl;
	f.songjangNo.value	= songjangNo;
	f.submit();
}