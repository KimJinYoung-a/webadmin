
// ============================================================================
// 마스터 사유구분 설정
// ============================================================================
function selectGubun(value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_gubun01, name_gubun02, name_gubun01name, name_gubun02name ,name_frm, targetDiv) {

	if ((IsPossibleModifyCSMaster != true) || (IsPossibleModifyItemList != true)) {
		alert(ERROR_MSG_TRY_MODIFY);
		return;
	}

    var frm = eval(name_frm);

    eval("document." + name_frm + "." + name_gubun01).value = value_gubun01;
    eval("document." + name_frm + "." + name_gubun02).value = value_gubun02;
    eval("document." + name_frm + "." + name_gubun01name).value = value_gubun01name;
    eval("document." + name_frm + "." + name_gubun02name).value = value_gubun02name;

    eval(targetDiv).innerHTML = "";

	if (!frm) {
		return;
	}

	var arrorderdetailidx = document.getElementsByName("orderdetailidx");
	var arrchangecsdetailidx = document.getElementsByName("changecsdetailidx");
	var obj;


    // 전체 사유구분을 입력한 경우 선택된 Detail에 같은 사유 세팅
    if (targetDiv=="causepop") {
    	// 주문 상품
    	for (var i = 0; i < arrorderdetailidx.length; i++) {
    		obj = arrorderdetailidx[i];

	        if (obj.type != "checkbox") {
	        	continue;
	        }

	        if (obj.checked != true) {
	        	continue;
	        }

	        setDetailCause(obj.value, value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_frm);
    	}

		// 상품변경 맞교환출고 상품
    	for (var i = 0; i < arrchangecsdetailidx.length; i++) {
    		obj = arrchangecsdetailidx[i];

	        if (obj.type != "checkbox") {
	        	continue;
	        }

	        if (obj.checked != true) {
	        	continue;
	        }

	        setDetailCause(obj.value, value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_frm);
    	}
    }

	// 상품변경시 체크
	CheckForItemChanged();

	// 상품변경 맞교환시 배송비
	CheckBeasongPayCutItemChange(document.frmaction);
}

// ============================================================================
// 상품 사유구분 설정
// ============================================================================
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
// 상품 사유구분 삭제
// ============================================================================
function delGubun(name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv) {
    eval("document." + name_frm + "." + name_gubun01).value = "";
    eval("document." + name_frm + "." + name_gubun02).value = "";
    eval("document." + name_frm + "." + name_gubun01name).value = "";
    eval("document." + name_frm + "." + name_gubun02name).value = "";

    eval(targetDiv).innerHTML = "";
}

// ============================================================================
// CS 사유구분 표시 (AJAX)
// ============================================================================
function divCsAsGubunSelect(value_gubun01,value_gubun02,name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv) {

	if ((IsPossibleModifyCSMaster != true) || (IsPossibleModifyItemList != true)) {
		alert(ERROR_MSG_TRY_MODIFY);
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

// ============================================================================
// 상품변경시 체크
// ============================================================================
function CheckForItemChanged() {
	var frm = document.frmaction;

	// ========================================================================
	// 상품취소,반품시 배송비 동시취소(또는 마이너스 주문등록)
	// ========================================================================
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true)) {
	    CheckUpcheDeliverPay(frm);
	}

	// ========================================================================
	// 반품시 체크
	// ========================================================================
	if (IsCSReturnProcess == true) {
		// 배송비 차감(단순변심)
		CheckBeasongPayCut(frm);

		// 원주문수량 이상의 반품이 있는지 체크
		CheckOverReturnItemno(frm);
	}

	if (IsStatusRegister == true) {
		if (IsOnlyOneBrandAvailable == true) {
		    // 단일 브랜드만 선택 가능하게, 담당 브랜드 저장

		    if (IsStatusRegister == true) {
		    	EnableOnlyOneBrand(frm);
		    }

			// 업배 반품/맞교환 : 업체 추가정산 관련 브랜드 아이디 가져오기
			InsertCheckedUpcheID(frm);
		}

		// 체크된 상품/배송비 색바꾸기
		// AnCheckClickAll(frmaction);

		// 주문전체 선택시 마일리지, 할인권 자동체크
		CheckMileageETC(frm);
	}

	// 접수상품 금액 재계산
    CalculateAndApplyItemCostSum(frm);

	/*
	TODO : 상품취소로 인한 배송비 추가
    if (IsAddBrandBeasongPayNeed(frm, "") == true) {
    	alert("텐바이텐 배송 상품중 남는상품의 금액이 30000원 미만이므로 배송비 2000원이 추가됩니다.");
    }
    */
}

// ============================================================================
// 원주문수량 이상의 반품이 있는지 체크(기존 반품, 맞교환회수 CS완료내역)
// ============================================================================
function CheckOverReturnItemno(frm) {
	var chk, regitemno, itemno, prevcsreturnfinishno

	if ((IsCSReturnProcess != true) || (IsStatusFinished == true)) {
		return;
	}

	if (!frm.orderdetailidx) {
		return;
	}

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        chk = frm.orderdetailidx[i];
        regitemno = frm.regitemno[i];
        itemno = frm.itemno[i];
        prevcsreturnfinishno = frm.prevcsreturnfinishno[i];

        if (chk.type != "checkbox") {
        	continue;
        }

        if (chk.checked != true) {
        	continue;
        }

		if ((regitemno.value*1 + prevcsreturnfinishno.value*1) > itemno.value*1) {
			alert("주문수량을 초과하여 접수되는 상품이 있습니다.");

			if ((regitemno.value*1 - prevcsreturnfinishno.value*1)  > 0) {
				regitemno.value = (regitemno.value*1 - prevcsreturnfinishno.value*1);
			} else {
				chk.checked = false;
				AnCheckClick(frm.orderdetailidx[i]);
				delGubun("gubun01_" + chk.value,"gubun02_" + chk.value,"gubun01name_" + chk.value,"gubun02name_" + chk.value, frm.name, causepop);
			}

			CheckUpcheDeliverPay(frm);
		}
    }
}

// ============================================================================
// 디테일 상품 입력수량 체크
// ============================================================================
function CheckMaxItemNo(obj, maxno) {
	if (obj.value*0 != 0) {
		return;
	}

    if (obj.value*1 > maxno*1) {
        alert("주문수량 이상으로 상품수량을 수정할수 없습니다.");
        obj.value = maxno;
    }

	// 상품변경시 체크
	CheckForItemChanged();
}

// ============================================================================
// 상품선택시 확인할 것들
// ============================================================================
function CheckSelect(comp) {
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

	// 상품변경시 체크
	CheckForItemChanged();
}

// ============================================================================
// 배송비 동시 체크
// 텐배   - 텐배        상품 전체가 선택된 경우, 배송비 체크한다.
// 업배   - 업배 브랜드 상품 전체가 선택된 경우, 배송비 체크한다.
// XXXXXXXXXXXXXXXXXX반품이고, 단순변심인경우 체크하지 않는다.(CS 마스터 사유구분 기준)
// 모든상품이 선택되면 항상 배송비도 체크한다.
// ============================================================================
function CheckUpcheDeliverPay(frm) {
    var upbeaMakerid;

    var value_gubun02 = frm.gubun02.value;
    var objdeliver, objitem;

    if (!frm.Deliverdetailidx) return;
    if (!frm.orderdetailidx) return;
    if ((IsCSCancelProcess != true) && (IsCSReturnProcess != true)) return;

    for (var i = 0; i < frm.Deliverdetailidx.length; i++) {
        objdeliver = frm.Deliverdetailidx[i];

        if (objdeliver.type != "checkbox") {
        	continue;
        }

        upbeaMakerid = frm.DeliverMakerid[i].value;

		if ((IsOneBrandAllSelected(frm, upbeaMakerid) == true) || (IsForceCancelBeasongPayChecked(frm, upbeaMakerid))) {
			frm.Deliverdetailidx[i].checked = true;
		} else {
			frm.Deliverdetailidx[i].checked = false;
		}

		/*
        // 반품 프로세스고 단순변심인 경우 체크해제하지 않는다.
        if ((IsCSReturnProcess) && (value_gubun02 == "CD01")) {
            frm.Deliverdetailidx[i].checked = false;
        }
        */

        AnCheckClick(frm.Deliverdetailidx[i]);

        // 취소정보의 배송비 체크
        CheckUpcheDeliverPayCancel(frm, upbeaMakerid, frm.Deliverdetailidx[i].checked);
    }
}

function ForceCheckUpcheDeliverPay(frm) {
	CheckUpcheDeliverPay(frm);
	CalculateAndApplyItemCostSum(frm);
}

// ============================================================================
// 배송비가 선택되면 취소정보상의 배송비도 같이 체크
// ============================================================================
function CheckUpcheDeliverPayCancel(frm, upbeaMakerid, ischecked) {
	var objdeliver;

	if (!frm.ckbeasongpayAssign) {
		return;
	}

    for (var i = 0; i < frm.ckbeasongpayAssign.length; i++) {
        objdeliver = frm.ckbeasongpayAssign[i];

        if (objdeliver.type != "checkbox") {
        	continue;
        }

		if (upbeaMakerid != frm.CancelDeliverMakerid[i].value) {
			continue;
		}

		if ((ischecked == true) || (frm.ckbeasongpayAssign[i].checked == true)) {
			frm.refundbrandbeasongpay[i].value = frm.orgbrandbeasongpay[i].value;
		} else {
			frm.refundbrandbeasongpay[i].value = 0;
		}

		frm.remainbrandbeasongpay[i].value = frm.orgbrandbeasongpay[i].value*1 - frm.refundbrandbeasongpay[i].value*1;

		return;
    }
}

function IsForceCancelBeasongPayChecked(frm, upbeaMakerid) {
	var objdeliver;

    for (var i = 0; i < frm.ckbeasongpayAssign.length; i++) {
        objdeliver = frm.ckbeasongpayAssign[i];

        if (objdeliver.type != "checkbox") {
        	continue;
        }

		if (upbeaMakerid != frm.CancelDeliverMakerid[i].value) {
			continue;
		}

		return frm.ckbeasongpayAssign[i].checked;
    }

    return false;
}

// ============================================================================
// 배송비 강제취소
// ============================================================================
function FouceCheckDeliverPay(frm, upbeaMakerid, ischecked) {
	var objdeliver;

	// ========================================================================
	// 상품취소,반품시 배송비 동시취소(또는 마이너스 주문등록)
	// ========================================================================
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true)) {
	    // 배송비 동시취소
	    CheckUpcheDeliverPay(frm);
	}

    for (var i = 0; i < frm.Deliverdetailidx.length; i++) {
        objdeliver = frm.Deliverdetailidx[i];

        if (objdeliver.type != "checkbox") {
        	continue;
        }

		if (upbeaMakerid == frm.DeliverMakerid[i].value) {
			frm.Deliverdetailidx[i].checked = ischecked;
			AnCheckClick(frm.Deliverdetailidx[i]);
			CheckUpcheDeliverPayCancel(frm, upbeaMakerid, ischecked);

			CalculateAndApplyItemCostSum(frm);

			return;
		}
    }

    CalculateAndApplyItemCostSum(frm);
}


// ============================================================================
// 단일 브랜드만 선택 가능하게, 담당 브랜드 저장
//
// 텐배 체크시 - 업배 Disable
// 업배 체크시 - 텐배 및 다른 브랜드 Disable
// 업배상품 텐배회수일 경우 Disable 안함(여러브랜드 동시접수 가능)
// ============================================================================
function EnableOnlyOneBrand(frm) {
    var upbeaMakerid;
    var checkfound;

    var objdeliver, objitem;
    var e;
    var forcereturnbyten = GetForceReturnByTen(frm);
    var forcereturnbycustomer = GetForceReturnByCustomer(frm);

    if (!frm.Deliverdetailidx) return;
    if (!frm.orderdetailidx) return;

	if (IsOnlyOneBrandAvailable != true) {
		return;
	}

	// 접수시에만 계산한다.
	if (IsStatusRegister != true) {
		return;
	}

	// ========================================================================
	// 체크된 상품 검색
	checkfound = IsCheckedItemExist(frm);
	upbeaMakerid = GetCheckedItemMaker(frm);

	// ========================================================================
	// 담당 브랜드 저장
	if (checkfound != true) {
        frm.requireupche.value = "";
        frm.requiremakerid.value = "";
	} else {
		if (forcereturnbycustomer == true) {
			// 텐텐물류 고객반품
			frm.requireupche.value = "Y";
			frm.requiremakerid.value = "10x10logistics";
		} else if ((upbeaMakerid.length < 1) || (forcereturnbyten == true)) {
			// 텐텐회수
	        frm.requireupche.value = "N";
	        frm.requiremakerid.value = "";
		} else {
			// 업체반품
			frm.requireupche.value = "Y";
			frm.requiremakerid.value = upbeaMakerid;
		}
	}

	// ========================================================================
	// 업배상품 텐배회수일 경우 Disable 안함
	// 텐바이텐 물류로 고객직접반품일 경우 Disable 안함
	// ========================================================================
    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        objitem = frm.orderdetailidx[i];

        if (objitem.type != "checkbox") {
        	continue;
        }

		if ((checkfound != true) || (forcereturnbyten == true) || (forcereturnbycustomer == true)) {
			// 선택가능한 상품 전부 활성화
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
	        	((upbeaMakerid.length < 1) && ((frm.odlvtype[i].value != "1") && (frm.odlvtype[i].value != "4")))
	        	||
	        	((upbeaMakerid.length > 0) && (upbeaMakerid != frm.makerid[i].value))
	        ) {
	            objitem.checked = false;
	            objitem.disabled = true;
	        }
		}
    }
}

// ============================================================================
// 선택된 상품이 있는지
// ============================================================================
function IsCheckedItemExist(frm) {
    var objitem;

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        objitem = frm.orderdetailidx[i];

        if (objitem.type != "checkbox") {
        	continue;
        }

        if (objitem.checked) {
        	return true;
        }
    }

    return false;
}

// ============================================================================
// 선택된 상품의 브랜드 가져오기
// ============================================================================
function GetCheckedItemMaker(frm) {
    var objitem;
    var upbeaMakerid;

    upbeaMakerid = "";
    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        objitem = frm.orderdetailidx[i];

        if (objitem.type != "checkbox") {
        	continue;
        }

        if (objitem.checked) {
        	if ((frm.odlvtype[i].value == "1") || (frm.odlvtype[i].value == "4")) {
        		// 텐배
	        	upbeaMakerid = "";
        	} else {
        		// 업배
	        	upbeaMakerid = frm.makerid[i].value;
        	}
        	break;
        }
    }

    return upbeaMakerid;
}

// ============================================================================
// 전체 취소/반품인경우 -  마일리지/할인권 환원 체크 한다.
// ============================================================================
function CheckMileageETC(frm) {

    var allselected = IsAllSelected(frm);

	if (!frm.forcecouponreturn) {
		return;
	}

	// 보너스쿠폰(정액), 마일리지, Gift카드, 예치금 환원
	if ((IsCSCancelProcess || IsCSReturnProcess) && (IsStatusRegister == true)) {
		if (allselected) {
			frm.forcecouponreturn.checked = true;
			frm.forcemileagereturn.checked = true;
			frm.forcegiftcardreturn.checked = true;
			frm.forcedepositreturn.checked = true;
		} else {
			// 이미 체크되어 있으면 해제하지 않는다.
		}
	}
}

// ============================================================================
// 배송비
//
// - 단순변심 + 브랜드전체반품 = 왕복배송비(브랜드별 배송정책에 따라, 무료배송=2000원) 차감
//
// - 단순변심 + 기타 = 반품배송비 차감
//
// - 상품불량반품 등 = 배송비 차감 없음
// ============================================================================
function CheckBeasongPayCut(frm) {
    var allselected = IsAllSelected(frm);
    var brandallselected = IsBrandAllSelected(frm);
    var doubledeductionexist;
    var forcereturnbyten = GetForceReturnByTen(frm);
    var forcereturnbycustomer = GetForceReturnByCustomer(frm);
    var isupchebeasong = IsUpcheReturnState(frm);
    var makerid;

    var value_gubun02 = frm.gubun02.value;

	if (!frm.ckreturnpay) {
		return;
	}

	// ========================================================================
	// 반품시만 계산한다.
	if (IsCSReturnProcess != true) {
		return;
	}

	// 회수배송비 관련
    frm.ckreturnpay.checked = false;
    frm.ckreturnpayHalf.checked = false;
    frm.ckreturnpayZero.checked = false;

	// ========================================================================
	// 단순변심(CD01)이 아니면 차감 없음, 사이즈않맞음(CD06)추가
	if ((value_gubun02 != "CD01")&&(value_gubun02 != "CD06")) {
		return;
	}

	// ========================================================================
	// 브랜드 전체선택이이면 왕복배송비 차감, 아니면 2000원 차감.
    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if (e.type != "checkbox") {
       		continue;
        }

    	if (e.checked == true) {
			// 텐텐회수인경우 체크되어 있는 전브랜드 확인하여 한개의 브랜드라도 4000원 차감 조건을 충족하면 4000원 차감한다.
			// 업체반품의 경우도 텐텐 고객반품이 있으므로 모두 체크한다.
			if (frm.isupchebeasong[i].value != "Y") {
				makerid = "";
			} else {
				makerid = frm.makerid[i].value;
			}

			if ((IsOneBrandAllSelected(frm, makerid) == true) && (GetNotCheckedUpcheBeasongPayByBrand(frm, makerid) == 0)) {
		        frm.ckreturnpayHalf.checked = false;
		        frm.ckreturnpay.checked = true;
		        return;
    		} else if ((frm.regitemno[i].value*0 == 0) && (frm.regitemno[i].value*1 > 0)) {
    			frm.ckreturnpayHalf.checked = true;
			}
        }
    }
}

// ============================================================================
// 배송비(상품변경 맞교환)
//
// - 단순변심 = 왕복배송비(무료배송=2000*2 원) 차감
//
// - 상품불량반품 등 = 배송비 차감 없음
// ============================================================================
function CheckBeasongPayCutItemChange(frm) {
    var makerid;

    var value_gubun02 = frm.gubun02.value;

	if (frm.add_customeraddbeasongpay == undefined) {
		return;
	}

	frm.add_customeraddbeasongpay.value = 0;
	frm.add_customeraddmethod.value = "";

	// ========================================================================
	// 단순변심(CD01)이 아니면 차감 없음, 사이즈않맞음(CD06)추가, 사이즈교환(CD04)
	if ((value_gubun02 != "CD01") && (value_gubun02 != "CD06") && (value_gubun02 != "CD04")) {
		return;
	}

	frm.add_customeraddbeasongpay.value = GetUpcheDeliverPay(frm.requiremakerid.value) * 2;
	frm.add_customeraddmethod.value = "1";		// 박스동봉
}

function ForceCheckBeasongPayCut(frm) {
	CheckBeasongPayCut(frm);
	CalculateAndApplyItemCostSum(frm);
}

function GetForceReturnByTen(frm) {
    if (!frm.ForceReturnByTen) {
    	return false;
    } else {
    	return frm.ForceReturnByTen.checked;
    }
}

function GetForceReturnByCustomer(frm) {
    if (!frm.ForceReturnByCustomer) {
    	return false;
    } else {
    	return frm.ForceReturnByCustomer.checked;
    }
}

// ============================================================================
// 업체반품인가
// 접수시 선택된 상품으로 판단.
// 접수이후에는 divcd 로 판단한다.
// ============================================================================
function IsUpcheReturnState(frm) {
	if (IsStatusRegister != true) {
		return (divcd == "A004");
	}

	var forcereturnbyten = GetForceReturnByTen(frm);

    if (forcereturnbyten == true) {
    	return false;
    }

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if (e.type != "checkbox") {
        	continue;
        }

    	if (e.checked == true) {
    		if (frm.isupchebeasong[i].value == "Y") {
    			return true;
    		}
        }
    }

    return false;
}

// ============================================================================
// 업체 맞교환인가
// ============================================================================
function IsUpcheReturnStateItemChange(frm) {
	return (((divcd == "A000") || (divcd == "A100")) && (frm.requireupche.value == "Y"));
}

// ============================================================================
// 상품 전체 체크 되었는지
// ============================================================================
function IsAllSelected(frm) {
    var allselected = false;

	if (!frm.orderdetailidx) return allselected;

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if (e.type != "checkbox") {
        	continue;
        }

		if (frm.cancelyn[i].value == "Y") {
			continue;
		}

        if (e.checked == true) {
                allselected = true;
        } else {
                return false;
        }

        if (frm.regitemno[i].value != frm.itemno[i].value) {
            return false;
        }
    }

    return allselected;
}

// ============================================================================
// 브랜드 상품 전체 체크 되었는지(전체)
// 업배는 해당 브랜드 전체
// 텐배의 경우 텐배 전체
// ============================================================================
function IsBrandAllSelected(frm) {
    var brandallselected = false;
    var makerid = "";
	var isupchebeasong = "";
	var checkedmakeridlist = ",";

	if (!frm.orderdetailidx) { return brandallselected; }

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if (e.type != "checkbox") {
        	continue;
        }

		if (frm.isupchebeasong[i].value == "Y") {
			makerid = frm.makerid[i].value;
		} else {
			makerid = "";
		}

		if (checkedmakeridlist.match("," + makerid + ",") != null) {
			continue;
		}

        if (IsOneBrandAllSelected(frm, makerid) != true) {
        	return false;
        }
        brandallselected = true;
        checkedmakeridlist = checkedmakeridlist + makerid + ",";
    }

    return brandallselected;
}

// ============================================================================
// 한개 브랜드 전체 선택되었는지 체크
// 브랜드가 빈값이면 텐텐배송 전체체크
// 반품인경우 기존 반품,맞교환회수CS 완료내역 수량 합산
// ============================================================================
function IsOneBrandAllSelected(frm, makerid) {
    var onebrandallselected = false;
    var checkeditemexist = false;

    if (!frm.orderdetailidx) { return onebrandallselected; }

	// checkeditemexist = true;

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if (e.type != "checkbox") {
        	continue;
        }

		if (frm.cancelyn[i].value == "Y") {
			if (e.checked == true) {
				checkeditemexist = true;
			}
			continue;
		}

    	if (makerid == "") {
    		if (frm.isupchebeasong[i].value != "Y") {
    			if (e.checked != true) {
	                if ((IsCSReturnProcess != true) || (frm.prevcsreturnfinishno[i].value*1 != frm.itemno[i].value*1)) {
	                	return false;
	                }
    			} else {
    				checkeditemexist = true;
    			}

	            if (frm.regitemno[i].value != frm.itemno[i].value) {
	                if ((IsCSReturnProcess != true) || ((frm.regitemno[i].value*1 + frm.prevcsreturnfinishno[i].value*1) != frm.itemno[i].value*1)) {
	                	return false;
	                }
	            }
    		}
    	} else {
    		if ((frm.isupchebeasong[i].value == "Y") && (frm.makerid[i].value == makerid)) {
    			if (e.checked != true) {
	                if ((IsCSReturnProcess != true) || (frm.prevcsreturnfinishno[i].value*1 != frm.itemno[i].value*1)) {
	                	return false;
	                }
    			} else {
    				checkeditemexist = true;
    			}

	            if (frm.regitemno[i].value != frm.itemno[i].value) {
	                if ((IsCSReturnProcess != true) || ((frm.regitemno[i].value*1 + frm.prevcsreturnfinishno[i].value*1) != frm.itemno[i].value*1)) {
	                	return false;
	                }
	            }
    		}
    	}
    }

	if (checkeditemexist == true) {
		onebrandallselected = true;
	}

    return onebrandallselected;
}

// ============================================================================
// 선택안된 브랜드 상품금액 합계(TODO : 상품쿠폰 적용가 -> 판매가(할인가))
// ============================================================================
function GetOneBrandNotSelectedItemcost(frm, makerid) {
	var result = 0;

    if (!frm.orderdetailidx) { return onebrandallselected; }

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if (e.type != "checkbox") {
        	continue;
        }

    	if (makerid == "") {
    		if (frm.isupchebeasong[i].value != "Y") {
    			if (e.checked != true) {
    				result = result + (frm.itemcost[i].value * frm.itemno[i].value);
    				continue;
    			}

	            if (frm.regitemno[i].value != frm.itemno[i].value) {
    				result = result + (frm.itemcost[i].value * (frm.itemno[i].value - frm.regitemno[i].value));
    				continue;
	            }
    		}
    	} else {
    		if ((frm.isupchebeasong[i].value == "Y") && (frm.makerid[i].value == makerid)) {
    			if (e.checked != true) {
    				result = result + (frm.itemcost[i].value * frm.itemno[i].value);
    				continue;
    			}

	            if (frm.regitemno[i].value != frm.itemno[i].value) {
    				result = result + (frm.itemcost[i].value * (frm.itemno[i].value - frm.regitemno[i].value));
    				continue;
	            }
    		}
    	}
    }

    return result;
}

function CalculateAndApplyItemCostSum(frm) {

	// 환불대상 상품가격(상품쿠폰적용가)합계, 올엣카드할인합계, 비율쿠폰할인합계
	CalculateCancelItemSUM(frm);

	// 배송비 합계 계산
	CalculateBeasongPaySum(frm);

	// 구매총액 계산
	CalculateTotalBuyPaySum(frm);

	// 반품시 회수 배송비 계산
	CalculateReturnBeasongPay(frm);

	// 정액쿠폰 환원 재계산, 환불 쿠폰합계 재계산
	CalculateFixedCoupon(frm);

	// 마일리지 환원 재계산
	CalculateMileage(frm);

	// Gift카드 재계산
	CalculateGiftCard(frm);

	// 예치금 재계산
	CalculateDeposit(frm);

	// 취소 금액 합계 계산
	CalculateTotal(frm);
}

// Gift카드 재계산
function CalculateGiftCard(frm) {
    var orggiftcardsum	    = 0;
    var refundgiftcardsum    = 0;
    var remaingiftcardsum    = 0;

    var prevrefundsubtotalprice = 0;	// 구매취소합계 - 비율쿠폰 - 기타할인 - 정액쿠폰 - 마일리지 + 반품배송비 + 보정금액

	if (!frm.orggiftcardsum) { return; }

	prevrefundsubtotalprice = frm.refundtotalbuypaysum.value*1 + frm.refundpercentcouponsum.value*1 + frm.refundallatsubtractsum.value*1 + frm.refundfixedcouponsum.value*1 + frm.refundmileagesum.value*1;
	prevrefundsubtotalprice = prevrefundsubtotalprice + frm.refunddeliverypay.value*1 + frm.refundadjustpay.value*1;

	orggiftcardsum = frm.orggiftcardsum.value;
	refundgiftcardsum = frm.refundgiftcardsum.value;

	if (IsStatusFinished != true) {
		if (frm.forcegiftcardreturn.checked) {
			// 환원 Gift카드가 취소금액보다 큰경우
	        if ((prevrefundsubtotalprice) < (frm.orggiftcardsum.value*-1)) {
	        	refundgiftcardsum = prevrefundsubtotalprice*-1;
	        } else {
	        	refundgiftcardsum = orggiftcardsum;
	        }
		} else {
			refundgiftcardsum = 0;
		}
	}

	remaingiftcardsum = orggiftcardsum - refundgiftcardsum;

	// ========================================================================
	// Gift카드 재계산
	frm.refundgiftcardsum.value = refundgiftcardsum;
	frm.remaingiftcardsum.value = remaingiftcardsum;
}

// 예치금 재계산
function CalculateDeposit(frm) {
    var orgdepositsum	    = 0;
    var refunddepositsum    = 0;
    var remaindepositsum    = 0;

	var prevrefundsubtotalprice = 0;	// 구매취소합계 - 비율쿠폰 - 기타할인 - 정액쿠폰 - 마일리지 - Gift카드 + 반품배송비 + 보정금액

	if (!frm.orgdepositsum) { return; }

	prevrefundsubtotalprice = frm.refundtotalbuypaysum.value*1 + frm.refundpercentcouponsum.value*1 + frm.refundallatsubtractsum.value*1 + frm.refundfixedcouponsum.value*1 + frm.refundmileagesum.value*1;
	prevrefundsubtotalprice = prevrefundsubtotalprice + frm.refundgiftcardsum.value*1;
	prevrefundsubtotalprice = prevrefundsubtotalprice + frm.refunddeliverypay.value*1 + frm.refundadjustpay.value*1;

	orgdepositsum = frm.orgdepositsum.value;
	refunddepositsum = frm.refunddepositsum.value;

	if (IsStatusFinished != true) {
		if (frm.forcedepositreturn.checked) {
			// 환원 예치금이 취소금액보다 큰경우
	        if ((prevrefundsubtotalprice) < (orgdepositsum*-1)) {
	        	refunddepositsum = prevrefundsubtotalprice*-1;
	        } else {
	        	refunddepositsum = orgdepositsum;
	        }
		} else {
			refunddepositsum = 0;
		}
	}

	remaindepositsum = orgdepositsum - refunddepositsum;

	// ========================================================================
	// 예치금 재계산
	frm.refunddepositsum.value = refunddepositsum;
	frm.remaindepositsum.value = remaindepositsum;
}

// 구매총액 계산
function CalculateTotalBuyPaySum(frm) {
	if (!frm.orgtotalbuypaysum) {
		return;
	}

	frm.refundtotalbuypaysum.value = frm.refundbeasongpay.value*1 + frm.refunditemcostsum.value*1;
	frm.remaintotalbuypaysum.value = frm.remainbeasongpay.value*1 + frm.remainitemcostsum.value*1;
}

// 환불대상 상품가격(상품쿠폰적용가)합계, 올엣카드할인합계, 비율쿠폰할인합계
function CalculateCancelItemSUM(frm) {

    var e;
    var ischecked       	= false;
    var regitemno       	= 0;
    var itemno          	= 0;
    var cancelyn			= "Y";

    var itemcost        				= 0;
    var allatitemdiscount				= 0;
    var percentBonusCouponDiscount 		= 0;

    var orgitemcostsum     				= 0;
    var refunditemcostsum   			= 0;

    var orgallatitemdiscountSum 		= 0;
    var refundallatitemdiscountSum 		= 0;

    var orgpercentBonusCouponDiscountSum 		= 0;
    var refundpercentBonusCouponDiscountSum 	= 0;

	// ========================================================================
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

        if (e.name == "regitemno") {
            if ((e.value * 0) == 0) {
                regitemno = e.value;
            } else {
                regitemno = 0;
            }
        }

        if (e.name == "itemno") {
            if ((e.value * 0) == 0) {
                itemno = e.value;
            } else {
                itemno = 0;
            }
        }

        if (e.name == "itemcost") {
            if ((e.value * 0) == 0) {
                itemcost = e.value;
            } else {
                itemcost = 0;
            }
        }

        if (e.name == "cancelyn") {
            cancelyn = e.value;
        }

        if (e.name == "allatitemdiscount") {
            if ((e.value * 0) == 0) {
               allatitemdiscount = e.value;
            } else {
                allatitemdiscount = 0;
            }
        }

        if (e.name == "percentBonusCouponDiscount") {
            if ((e.value * 0) == 0) {
                percentBonusCouponDiscount = e.value;
            } else {
                percentBonusCouponDiscount = 0;
            }
        }

        if (e.name == "dummystopper") {
            // 선택상품합계
            if (ischecked == true) {
                refunditemcostsum 					= refunditemcostsum + (itemcost * regitemno * 1);
                refundallatitemdiscountSum 			= refundallatitemdiscountSum + (allatitemdiscount * regitemno * 1);
                refundpercentBonusCouponDiscountSum = refundpercentBonusCouponDiscountSum + (percentBonusCouponDiscount * regitemno * 1);
            }

			// 취소안된 상품 전체합계
			if (cancelyn != "Y") {
				orgitemcostsum 						= orgitemcostsum + (itemcost * itemno * 1);
				orgallatitemdiscountSum 			= orgallatitemdiscountSum + (allatitemdiscount * itemno * 1);
				orgpercentBonusCouponDiscountSum 	= orgpercentBonusCouponDiscountSum + (percentBonusCouponDiscount * itemno * 1);
			}

            ischecked = false;
            regitemno = 0;
            itemno = 0;
            itemcost = 0;
            cancelyn = "Y";
        }
    }

	// ========================================================================
    // 상품목록 하단 선택 합계
    if (frm.itemcanceltotal!=undefined){
		frm.itemcanceltotal.value = refunditemcostsum;
    }

	// ========================================================================
    // 취소안된 상품 전체합계 금액(최초접수시), 접수이후는 접수당시 취소않된 상품금액합계
    if (frm.orgitemcostsum!=undefined){
		orgitemcostsum = frm.orgitemcostsum.value;
    }

    // 취소/반품 상품총액
    if (frm.refunditemcostsum!=undefined){
        frm.refunditemcostsum.value = refunditemcostsum;
    }

    // 선택안한(나머지) 상품총액
    if (frm.remainitemcostsum!=undefined){
        frm.remainitemcostsum.value = orgitemcostsum - refunditemcostsum;
    }

	// ========================================================================
    // 취소안된 비율쿠폰 전체합계 금액
    if (frm.orgpercentcouponsum!=undefined){
		// aaaaaaaaaaaaaaaa frm.orgpercentcouponsum.value = orgpercentBonusCouponDiscountSum * -1;
    }

    // 취소/반품 비율쿠폰 총액
    if (frm.refundpercentcouponsum!=undefined){
        frm.refundpercentcouponsum.value = refundpercentBonusCouponDiscountSum * -1;
    }

    // 선택안한(나머지) 비율쿠폰 총액
    if (frm.remainpercentcouponsum!=undefined){
        frm.remainpercentcouponsum.value = frm.orgpercentcouponsum.value*1 - frm.refundpercentcouponsum.value*1;
    }

	// ========================================================================
    // 취소안된 기타할인 전체합계 금액
    if (frm.orgallatsubtractsum!=undefined){
		// aaaaaaaaaaaaaaaaa frm.orgallatsubtractsum.value = orgallatitemdiscountSum * -1;
    }

    // 취소/반품 기타할인 총액
    if (frm.refundallatsubtractsum!=undefined){
        frm.refundallatsubtractsum.value = refundallatitemdiscountSum * -1;
    }

    // 선택안한(나머지) 기타할인 총액
    if (frm.remainallatsubtractsum!=undefined){
        frm.remainallatsubtractsum.value = (orgallatitemdiscountSum * -1) - (refundallatitemdiscountSum * -1);
    }

    // 과거 변수명도 설정해준다.(pop_cs_action_new_process.asp 에 값을 넘기기 위해 필요)
    if (frm.allatdiscountprice!=undefined){
		frm.allatdiscountprice.value = orgallatitemdiscountSum * -1;
    }

    if (frm.allatsubtractsum!=undefined){
		frm.allatsubtractsum.value = refundallatitemdiscountSum * -1;
    }

    if (frm.remainallatdiscount!=undefined){
		frm.remainallatdiscount.value = (orgallatitemdiscountSum * -1) - (refundallatitemdiscountSum * -1);
    }
}


// 배송비 합계 계산
function CalculateBeasongPaySum(frm) {
	var orgbeasongpay = 0;
	var refundbeasongpay = 0;
	var remainbeasongpay = 0;

	var objdeliver;

	if (!frm.ckbeasongpayAssign) {
		return;
	}

    for (var i = 0; i < frm.ckbeasongpayAssign.length; i++) {
        objdeliver = frm.ckbeasongpayAssign[i];

        if (objdeliver.type != "checkbox") {
        	continue;
        }

		frm.remainbrandbeasongpay[i].value = frm.orgbrandbeasongpay[i].value*1 - frm.refundbrandbeasongpay[i].value*1;

		orgbeasongpay = orgbeasongpay + frm.orgbrandbeasongpay[i].value*1;
		refundbeasongpay = refundbeasongpay + frm.refundbrandbeasongpay[i].value*1;
		remainbeasongpay = remainbeasongpay + frm.remainbrandbeasongpay[i].value*1;
    }

    frm.refundbeasongpay.value = refundbeasongpay;
    frm.remainbeasongpay.value = remainbeasongpay;
}

// 환불 정액쿠폰 계산
function CalculateFixedCoupon(frm) {

	if (!frm.orgpercentcouponsum) {
		return;
	}

	var orgfixedcouponsum 		= 0;
	var refundfixedcouponsum	= 0;
	var remainfixedcouponsum 	= 0;

	var orgcouponsum			= 0; // tencardspend(주문취소 처리완료 이후 달라짐)
    var refundcouponsum    		= 0;
    var remaincouponsum    		= 0;

	var prevrefundsubtotalprice = 0;	// 구매취소합계 - 비율쿠폰 - 기타할인

	prevrefundsubtotalprice = frm.refundtotalbuypaysum.value*1 + frm.refundpercentcouponsum.value*1 + frm.refundallatsubtractsum.value*1;

	// ========================================================================
	// 정액보너스쿠폰 계산
	orgfixedcouponsum = frm.orgcouponsum.value - frm.orgpercentcouponsum.value;
	refundfixedcouponsum = frm.refundfixedcouponsum.value;

	if (IsStatusFinished != true) {
		if (frm.forcecouponreturn.checked) {
			// 환원쿠폰이 취소금액보다 큰경우
	        if ((prevrefundsubtotalprice) < (orgfixedcouponsum*-1)) {
	        	refundfixedcouponsum = prevrefundsubtotalprice*-1;
	        } else {
	        	refundfixedcouponsum = orgfixedcouponsum;
	        }
		} else {
			refundfixedcouponsum = 0;
		}
	}

	remainfixedcouponsum = orgfixedcouponsum - refundfixedcouponsum;

	frm.remainfixedcouponsum.value = remainfixedcouponsum;
	frm.refundfixedcouponsum.value = refundfixedcouponsum;

	// ========================================================================
	// 보너스쿠폰 합산(비율 + 정액)액 계산
	frm.remaincouponsum.value = frm.remainpercentcouponsum.value*1 + frm.remainfixedcouponsum.value*1;
	frm.refundcouponsum.value = frm.orgcouponsum.value - frm.remaincouponsum.value;
}

// 환불 마일리지 계산
function CalculateMileage(frm) {

    var orgmileagesum	    = 0;	// miletotalprice
    var refundmileagesum    = 0;
    var remainmileagesum    = 0;

    var prevrefundsubtotalprice = 0;	// 구매취소합계 - 비율쿠폰 - 기타할인 - 정액쿠폰 + 반품배송비 + 보정금액

	if (!frm.forcemileagereturn) {
		return;
	}

	prevrefundsubtotalprice = frm.refundtotalbuypaysum.value*1 + frm.refundpercentcouponsum.value*1 + frm.refundallatsubtractsum.value*1 + frm.refundfixedcouponsum.value*1;
	prevrefundsubtotalprice = prevrefundsubtotalprice + frm.refunddeliverypay.value*1 + frm.refundadjustpay.value*1;

	orgmileagesum = frm.orgmileagesum.value;
	refundmileagesum = frm.refundmileagesum.value;

	if (IsStatusFinished != true) {
		if (frm.forcemileagereturn.checked) {
			// 환원마일리지가 취소금액보다 큰경우
	        if ((prevrefundsubtotalprice) < (frm.orgmileagesum.value*-1)) {
	        	refundmileagesum = prevrefundsubtotalprice*-1;
	        } else {
	        	refundmileagesum = orgmileagesum;
	        }
		} else {
			refundmileagesum = 0;
		}
	}

	remainmileagesum = orgmileagesum - refundmileagesum;

	// ========================================================================
	// 마일리지 재계산
	frm.refundmileagesum.value = refundmileagesum;
	frm.remainmileagesum.value = remainmileagesum;
}

// 반품시 회수 배송비 계산
function CalculateReturnBeasongPay(frm) {
    var orgbeasongpay       	= 0;
    var refunddeliverypay   	= 0;
    var remainbeasongpay	   	= 0;
    var refundbeasongpay    	= 0;

	var makerid					= "";

	if (IsCSCancelInfoNeeded != true) {
		return;
	}

	if (frm.buf_requiremakerid) {
		makerid = frm.buf_requiremakerid.value;
	}

    // 회수 배송비
    if (frm.ckreturnpay!=undefined){
        if (frm.ckreturnpayHalf.checked){
            refunddeliverypay = GetUpcheDeliverPay(makerid) * -1;

        }else if (frm.ckreturnpay.checked){
            refunddeliverypay = GetUpcheDeliverPay(makerid) * -2;
        }else{
            refunddeliverypay = 0;
        }
    }

    if (frm.refunddeliverypay!=undefined){
        frm.refunddeliverypay.value = refunddeliverypay*1;
    }

    if (frm.buf_refunddeliverypay!=undefined){
    	if (makerid != "10x10logistics") {
	        frm.buf_refunddeliverypay.value = refunddeliverypay*-1;
    	} else {
	        frm.buf_refunddeliverypay.value = 0;
    	}
    	frm.buf_totupchejungsandeliverypay.value = frm.buf_refunddeliverypay.value*1 + frm.add_upchejungsandeliverypay.value*1;
    }
}

// 모든 금액 합산
function CalculateTotal(frm) {

	// ========================================================================
	// 구매총액
	var orgtotalbuypaysum = 0;
	var refundtotalbuypaysum = 0;
	var remaintotalbuypaysum = 0;

	if (!frm.orgtotalbuypaysum) {
		return;
	}

	orgtotalbuypaysum = frm.orgtotalbuypaysum.value*1;
	refundtotalbuypaysum = frm.refundtotalbuypaysum.value*1;
	remaintotalbuypaysum = frm.remaintotalbuypaysum.value*1;

	// ========================================================================
	// 사용 보너스쿠폰(비율)
	var orgpercentcouponsum = 0;
	var refundpercentcouponsum = 0;
	var remainpercentcouponsum = 0;

	orgpercentcouponsum = frm.orgpercentcouponsum.value*1;
	refundpercentcouponsum = frm.refundpercentcouponsum.value*1;
	remainpercentcouponsum = frm.remainpercentcouponsum.value*1;

	// ========================================================================
	// 사용 기타할인
	var orgallatsubtractsum = 0;
	var refundallatsubtractsum = 0;
	var remainallatsubtractsum = 0;

	orgallatsubtractsum = frm.orgallatsubtractsum.value*1;
	refundallatsubtractsum = frm.refundallatsubtractsum.value*1;
	remainallatsubtractsum = frm.remainallatsubtractsum.value*1;

	// ========================================================================
	// 사용 보너스쿠폰(정액)
	var orgfixedcouponsum = 0;
	var refundfixedcouponsum = 0;
	var remainfixedcouponsum = 0;

	orgfixedcouponsum = frm.orgfixedcouponsum.value*1;
	refundfixedcouponsum = frm.refundfixedcouponsum.value*1;
	remainfixedcouponsum = frm.remainfixedcouponsum.value*1;

	// ========================================================================
	// 사용 마일리지
	var orgmileagesum = 0;
	var refundmileagesum = 0;
	var remainmileagesum = 0;

	orgmileagesum = frm.orgmileagesum.value*1;
	refundmileagesum = frm.refundmileagesum.value*1;
	remainmileagesum = frm.remainmileagesum.value*1;

	// ========================================================================
	// 사용 Gift카드
	var orggiftcardsum = 0;
	var refundgiftcardsum = 0;
	var remaingiftcardsum = 0;

	orggiftcardsum = frm.orggiftcardsum.value*1;
	refundgiftcardsum = frm.refundgiftcardsum.value*1;
	remaingiftcardsum = frm.remaingiftcardsum.value*1;

	// ========================================================================
	// 사용 예치금
	var orgdepositsum = 0;
	var refunddepositsum = 0;
	var remaindepositsum = 0;

	orgdepositsum = frm.orgdepositsum.value*1;
	refunddepositsum = frm.refunddepositsum.value*1;
	remaindepositsum = frm.remaindepositsum.value*1;

	// ========================================================================
	// 결제금액
	var orgtotalrealbuypaysum = 0;
	var refundtotalrealbuypaysum = 0;
	var remaintotalrealbuypaysum = 0;

	orgtotalrealbuypaysum = orgtotalbuypaysum + orgpercentcouponsum + orgallatsubtractsum + orgfixedcouponsum + orgmileagesum + orggiftcardsum + orgdepositsum;
	remaintotalrealbuypaysum = remaintotalbuypaysum + remainpercentcouponsum + remainallatsubtractsum + remainfixedcouponsum + remainmileagesum + remaingiftcardsum + remaindepositsum;

	refundtotalrealbuypaysum = orgtotalrealbuypaysum - remaintotalrealbuypaysum;

	frm.orgtotalrealbuypaysum.value = orgtotalrealbuypaysum;
	frm.refundtotalrealbuypaysum.value = refundtotalrealbuypaysum;
	frm.remaintotalrealbuypaysum.value = remaintotalrealbuypaysum;

	// ========================================================================
	// 회수 배송비
	var refunddeliverypay = 0;

	refunddeliverypay = frm.refunddeliverypay.value*1;

	// 기타보정금액
	var refundadjustpay = 0;
    // 티켓주문인경우..================================================
    if ((IsTicketOrder==true)&&(mayTicketCancelChargePro>0)){
        if ((refundtotalbuypaysum!=0)&&(frm.refundadjustpay.value*1==0)){
            var mayTicketCancelPro = getFieldValue(frm.tRefundPro)*1;
            if (mayTicketCancelPro>0){
                alert( ticketCancelStr + '티켓 취소 수수료 ' + mayTicketCancelPro + '% 차감 \n\n(단, 당일 주문건 취소시는 제외)' );
                frm.refundadjustpay.value = (refundtotalbuypaysum*mayTicketCancelPro/100)*-1;
            }
        }
    }
    // 티켓주문인경우..================================================


	refundadjustpay = frm.refundadjustpay.value*1;

	// ========================================================================
	// 취소금액
	var orgsubtotalprice = 0;
	var refundsubtotalprice = 0;
	var remainsubtotalprice = 0;

	orgsubtotalprice = orgtotalrealbuypaysum;
	refundsubtotalprice = refundtotalrealbuypaysum + refunddeliverypay + refundadjustpay;
	remainsubtotalprice = remaintotalrealbuypaysum;

	frm.orgsubtotalprice.value = orgsubtotalprice;
	frm.refundsubtotalprice.value = refundsubtotalprice;
	frm.remainsubtotalprice.value = remainsubtotalprice;

	// 과거 변수명도 설정해준다.(pop_cs_action_new_process.asp 에 값을 넘기기 위해 필요)
	frm.subtotalprice.value = orgsubtotalprice;
	frm.canceltotal.value = refundsubtotalprice;
	frm.nextsubtotal.value = remainsubtotalprice;

	// ========================================================================
	// 환불금액 입력
    if (parseInt(frm.ipkumdiv.value) >= 4) {
        if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true)) {
            if (frm.refundrequire!=undefined) {
                frm.refundrequire.value = frm.refundsubtotalprice.value*1;
            }
        }
    }
}

// ============================================================================
// 선택된 업체배송비 합계
// ============================================================================
function GetCheckedUpcheBeasongPay(frm) {
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

// ============================================================================
// 선택않된 상품 안보이기
// ============================================================================
function ShowOnlySelectedItem(frm) {
    var e, t;

    if (!frm.orderdetailidx) return;

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];
        t = frm.orderdetailidx[i];

        if (e.type == "checkbox") {
			while (t.tagName != "TR") {
				t = t.parentElement;
			}

			if (e.checked == true) {
				t.style.display = '';
			} else {
				t.style.display = 'none';
			}
        }
    }
}

// ============================================================================
// 전체 상품 표시
// ============================================================================
function ShowAllItem(frm) {
    var e, t;

    if (!frm.orderdetailidx) return;

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];
        t = frm.orderdetailidx[i];

        if (e.type == "checkbox") {
			while (t.tagName != "TR") {
				t = t.parentElement;
			}

			t.style.display = '';
        }
    }
}

// ============================================================================
// 선택않된 업체배송비 합계(전체)
// ============================================================================
function GetNotCheckedUpcheBeasongPay(frm) {

    var checkfound, upbeaMakerid;
    var objdeliver, objitem;

    if (!frm.Deliverdetailidx) return;
    if (!frm.orderdetailidx) return;

	if (IsCSReturnProcess != true) {
		return 0;
	}

	checkfound = IsCheckedItemExist(frm);
	upbeaMakerid = GetCheckedItemMaker(frm);

    if (checkfound == false) {
    	return 0;
    }

    for (var i = 0; i < frm.Deliverdetailidx.length; i++) {
        e = frm.Deliverdetailidx[i];

        if ((e.type == "checkbox") && (e.checked == false) && (upbeaMakerid == frm.DeliverMakerid[i].value)) {
            return frm.Deliveritemcost[i].value*1;
        }
    }

    return 0;
}

// ============================================================================
// 선택않된 업체배송비 합계(브랜드별)
// ============================================================================
function GetNotCheckedUpcheBeasongPayByBrand(frm, makerid) {

    if (!frm.Deliverdetailidx) return 0;
    if (!frm.orderdetailidx) return 0;

	if (IsCSReturnProcess != true) {
		return 0;
	}

    for (var i = 0; i < frm.Deliverdetailidx.length; i++) {
        e = frm.Deliverdetailidx[i];

        if (e.type != "checkbox") {
        	continue;
        }

        if ((e.checked == false) && (makerid == frm.DeliverMakerid[i].value)) {
            return frm.Deliveritemcost[i].value*1;
        }
    }

    return 0;
}

// ============================================================================
// 회수 배송비 중복체크
// ============================================================================
function CheckDoubleCheck(frm,comp) {
    if (comp.name=="ckreturnpay"){
        if (frm.ckreturnpay.checked){
            frm.ckreturnpayHalf.checked = false;
            frm.ckreturnpayZero.checked = false;
        }
    }else if (comp.name=="ckreturnpayHalf"){
        if (frm.ckreturnpayHalf.checked){
            frm.ckreturnpay.checked = false;
            frm.ckreturnpayZero.checked = false;
        }
    }else if (comp.name=="ckreturnpayZero"){
        if (frm.ckreturnpayZero.checked){
            frm.ckreturnpay.checked = false;
            frm.ckreturnpayHalf.checked = false;
        }
    }
}

// ============================================================================
// 텐텐 반품시 업배상품 강제 선택
// ============================================================================
function CheckForceReturnByTen(obj) {
	// 상품변경시 체크
	CheckForItemChanged();

	if (obj.checked == true) {
		frmaction.ForceReturnByCustomer.checked = false;
		frmaction.ForceReturnByCustomer.disabled = true;
	} else {
		frmaction.ForceReturnByCustomer.disabled = false;
	}
}

// ============================================================================
// 텐텐 반품시 고객직접반품 강제 선택
// ============================================================================
function CheckForceReturnByCustomer(obj) {
	// 상품변경시 체크
	CheckForItemChanged();

	if (obj.checked == true) {
		frmaction.ForceReturnByTen.checked = false;
		frmaction.ForceReturnByTen.disabled = true;
	} else {
		frmaction.ForceReturnByTen.disabled = false;
	}
}

// ============================================================================
// 환불방식 변경
// ============================================================================
function ChangeReturnMethod(comp){
    if (comp==undefined) return;

    var returnmethod = comp.value;

    document.all.refundinfo_R007.style.display = "none";
    document.all.refundinfo_R050.style.display = "none";
    document.all.refundinfo_R100.style.display = "none";
    document.all.refundinfo_R550.style.display = "none";
    document.all.refundinfo_R560.style.display = "none";
    document.all.refundinfo_R900.style.display = "none";

    if (comp.value=="R007"){
        //무통장 환불
        document.all.refundinfo_R007.style.display = "inline";
    }else if((comp.value=="R020")||(comp.value=="R080")||(comp.value=="R100")||(comp.value=="R120")||(comp.value=="R400")){
        //실시간 이체 취소//ALL@ 결제 취소 //신용카드 결제 취소 //신용카드 부분취소//휴대폰
        document.all.refundinfo_R100.style.display = "inline";
    }else if(comp.value=="R550"){
        //기프팅 결제 취소
        document.all.refundinfo_R550.style.display = "inline";
    }else if(comp.value=="R560"){
        //기프티콘 결제 취소
        document.all.refundinfo_R560.style.display = "inline";
    }else if(comp.value=="R050"){
        //입점몰 결제 취소
        document.all.refundinfo_R050.style.display = "inline";
    }else if ((comp.value=="R900") || (comp.value=="R910")) {
        //마일리지 환불, 예치금 환불
        document.all.refundinfo_R900.style.display = "inline";
    }

}

// ============================================================================
// 업체 추가정산 관련 브랜드 아이디 가져오기
// ============================================================================
function InsertCheckedUpcheID(frm) {
    var checkedUpcheid = "";
    var UpcheDuplicated = false;
    var IsUpcheReturn;

	// 접수시에만 변경한다.
	if (IsStatusRegister != true) {
		return;
	}

	if ((IsUpcheReturnState(frm) == true) || (divcd == "A000") || (divcd == "A700")) { // 반품접수(업체배송), 맞교환출고, 업체기타정산
		IsUpcheReturn = true;
	} else {
		IsUpcheReturn = false;
	}

    if (!frm.buf_requiremakerid) {
        return;
    }

	frm.buf_requiremakerid.value = "";
	if (IsUpcheReturn == true) {
		frm.buf_requiremakerid.value = GetCheckedItemMaker(frm);
	} else {
		// 텐텐배송상품 고객 직접반품
		if (frm.ForceReturnByCustomer.checked == true) {
			frm.buf_requiremakerid.value = "10x10logistics";
		}
	}
}

// ============================================================================
// CS 접수시에는 반품접수(업체배송)/회수신청(텐바이텐배송) 을 구분하지 않고
// 저장시 브랜드지정이 있는경우 반품접수(업체배송), 없는경우 회수신청(텐바이텐배송) 으로 저장한다.
// 제목을 수정한경우 제목은 변경하지 않는다.
// 고객직접반품이 설정된경우 업체배송으로 저장한다.
// ============================================================================
function ChangeCSTitleGubun(frm) {
	if (IsStatusRegister != true) {
		return;
	}

	if ((divcd != "A004") && (divcd != "A010")) {
		return;
	}

	if (!frm.ForceReturnByTen) { return; }

	if ((frm.buf_requiremakerid.value == "") || ((frm.ForceReturnByTen.checked == true) && (frm.ForceReturnByCustomer.checked == false))) {
		// 텐텐반품
		frm.divcd.value = "A010";
		if (frm.title.value == "반품접수") {
			frm.title.value = "회수신청(텐바이텐배송)";
		}
	} else {
		// 업체반품
		frm.divcd.value = "A004";
		if (frm.title.value == "반품접수") {
			frm.title.value = "반품접수(업체배송)";
		}
	}
}

// ============================================================================
//추가정산배송비 사유 변경
// ============================================================================
function Change_add_upchejungsancause(comp){
    if (comp.value=="직접입력") {
        document.all.span_add_upchejungsancauseText.style.display = "inline";
    }else{
        document.all.span_add_upchejungsancauseText.style.display = "none";
    }
}

// ============================================================================
//추가정산배송비 입력
// ============================================================================
function Change_add_upchejungsandeliverypay(comp){
    comp.form.buf_totupchejungsandeliverypay.value = comp.form.buf_refunddeliverypay.value*1 + comp.value*1;

    if (isNaN(comp.form.buf_totupchejungsandeliverypay.value)){
        comp.form.buf_totupchejungsandeliverypay.value = "0";
    }
}

// ============================================================================
// 업체 추가 정산 입력 삭제
// ============================================================================
function clearAddUpchejungsan(frm){
	if ((IsPossibleModifyCSMaster != true) || (IsPossibleModifyItemList != true)) {
		alert(ERROR_MSG_TRY_MODIFY);
		return;
	}

    frm.add_upchejungsandeliverypay.value = "0";
    frm.add_upchejungsancause.value = "";
    frm.add_upchejungsancauseText.value = "";
    frm.buf_totupchejungsandeliverypay.value = "0";

	// 상품변경시 체크
	CheckForItemChanged();
}

// ============================================================================
// 업체배송비 무료배송 조건금액
// ============================================================================
function GetUpcheFreeBeasongLimit(makerid) {
	for (var i = 0; i < arrmakerid.length; i++) {
		if (arrmakerid[i] == makerid) {
			return arrdefaultfreebeasonglimit[i];
		}
	}

	// 없으면 텐바이텐 기준금액
	return 30000;
}

// ============================================================================
// 업체배송비
// ============================================================================
function GetUpcheDeliverPay(makerid) {

	var frm = document.frmaction;

	var savedrefunddeliverypay = 0;

	if (frm.refunddeliverypay) {
		savedrefunddeliverypay = frm.refunddeliverypay.value * -1;
	}

	// 접수 이후에는 배송비 정책이 바껴도 입력된 금액으로 배송비를 산정한다.
	if (IsStatusRegister != true) {
		if (savedrefunddeliverypay >= 4000) {
			return (savedrefunddeliverypay / 2);
		}

		if ((savedrefunddeliverypay > 0) && (savedrefunddeliverypay < 4000)) {
			return savedrefunddeliverypay;
		}
	}

	for (var i = 0; i < arrmakerid.length; i++) {
		if ((arrmakerid[i] == makerid) && (arrdefaultdeliverpay[i] != 0)) {
			return arrdefaultdeliverpay[i];
		}
	}

	// 없으면 텐바이텐배송비
	return CDEFAULTBEASONGPAY;
}


// ============================================================================
// 업체배송비(상품변경 맞교환)
// ============================================================================
function GetUpcheDeliverPayItemChange(makerid) {

	var frm = document.frmaction;

	for (var i = 0; i < arrmakerid.length; i++) {
		if ((arrmakerid[i] == makerid) && (arrdefaultdeliverpay[i] != 0)) {
			return arrdefaultdeliverpay[i];
		}
	}

	// 없으면 텐바이텐배송비
	return CDEFAULTBEASONGPAY;
}


// ============================================================================
// 주문취소로 배송비추가가 필요한지
// ============================================================================
function IsAddBrandBeasongPayNeed(frm, makerid) {
	var value_gubun02 = frm.gubun02.value;

	// 취소시에만 계산, 단순변심 이외사유에는 추가 안함.
	if ((IsCSCancelProcess != true) || (value_gubun02 != "CD01")) {
		return false;
	}

	// 전체선택
	if (IsOneBrandAllSelected(frm, makerid) == true) {
		return false;
	}

	// 이미 배송비가 있는 경우
	if (GetNotCheckedUpcheBeasongPayByBrand(frm, makerid) > 0) {
		return false;
	}

	// aaaaaaaaaaaaaaaaa
	// 무료배송상품 또는 착불 상품이 있는지

	// 선택안된상품이 30000만원 미만인 경우
	if (GetOneBrandNotSelectedItemcost(frm, "") < GetUpcheFreeBeasongLimit(makerid)) {
		return true;
	}

	return false;
}

// ============================================================================
// CS 접수
// ============================================================================
function CsRegProc(frm) {
    if ((IsTicketOrder==true)&&(ticketCancelDisabled==true)){
        if (!confirm('취소 기한이 지나 취소 불가합니다. ' + ticketCancelStr + ' \n\n 계속하시겠습니까?')){
            return;
        }

        //권한 있는지 check
        if (IsCsPowerUser!=true){
            alert('권한이 없습니다. 파트장급 문의 요망');
            return;
        }
    }

    if (((IsGiftingOrder == true) || (IsGiftiConOrder == true)) && (IsOrderCancelDisabled == true)) {
        if (!confirm('취소불가 주문입니다. : ' + OrderCancelDisableStr + ' \n\n 계속하시겠습니까? [관리자권한 필요]')) {
            return;
        }

        //권한 있는지 check
        if (IsCsPowerUser != true) {
            alert('권한이 없습니다. 파트장급 문의 요망');
            return;
        }
    }

	var forcereturnbyten = GetForceReturnByTen(frm);

	// 마스터 체크
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	// 선택 상품 체크 및 저장
	if ((divcd != "A100") && (divcd != "A111")) {
	    if (!SaveCheckedItemList(frm)) {
	        return;
	    }
	} else {
	    if (!SaveChangeCheckedItemList(frm)) {
	        return;
	    }
	}

	// 취소, 반품, 환불
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true) || (divcd == "A003")) {
        if (CheckReturnForm(frm) != true) {
            return;
        }
	}

	// 강제 환불요청 허용
	if (divcd == "A003") {
		if (frm.refundrequire) {
			if ((frm.refundrequire.value*1 > RefundAllowLimit) && (RefundAllowLimit != -1)) {
		        alert('권한이 없습니다. 정직원 또는 파트장급 문의 요망');
		        frm.refundrequire.focus();
		        return;
			}
		}
	}

	// 취소, 반품
    if ((IsCSCancelProcess) || (IsCSReturnProcess)) {
		// 환불정보가 정상인지 체크
		if (IsRefundInfoOK(frm) != true) {
			return;
		}

		// 환불수단이 옳바른지
		if (CheckReturnMethod(frm) != true) {
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

			if (divcd == "A700") {
	            alert('추가 정산액을 입력하세요.');
	            frm.add_upchejungsandeliverypay.focus();
	            return;
			}
        }
    }

    if (IsCSCancelProcess){
        if(confirm("취소 접수 하시겠습니까?")){
            frm.submit();
        }
    }else if (IsCSReturnProcess){
        if (frm.requireupche.value=="Y"){
            if(confirm("업체 [" + frm.requiremakerid.value +"]로 반품/회수/교환 접수 하시겠습니까?")){
                ChangeCSTitleGubun(frm);
                frm.submit();
            }
        }else{
            if(confirm("[텐바이텐 물류센터]로 반품/회수/교환 접수 하시겠습니까?")){
                ChangeCSTitleGubun(frm);
                frm.submit();
            }
        }
    }else if (IsCSRefundNeeded) {
        if(confirm("환불 접수 하시겠습니까?")){
            frm.submit();
        }
    }else if(confirm("접수 하시겠습니까?")){
        frm.submit();
    }
}

function CheckCSMasterForSave(frm) {
    if (frm.divcd.value.length<1){
        alert("접수 구분을 선택하세요.");
        frm.divcd.focus();
        return false;
    }

    if (frm.title.value.length<1) {
        alert("제목을 입력하세요.");
        frm.title.focus();
        return false;
    }

    if (frm.gubun01.value.length<1) {
        alert("사유 구분을 입력하세요.");
        return false;
    }

    return true;
}

// ============================================================================
// 환불 금액이 적절한지
// ============================================================================
function IsRefundInfoOK(frm) {

	if ((IsCSCancelProcess != true) && (IsCSReturnProcess != true)) {
		return true;
	}

	// 취소, 반품시 : 환불금액과 비교
	if (frm.orgsubtotalprice && frm.refundsubtotalprice) {
	    if (frm.orgsubtotalprice.value*1 < frm.refundsubtotalprice.value*1) {
	        alert('결제금액 이상으로 환불할 수 없습니다.\n\n마일리지, 쿠폰 등을 환원체크하세요.');
	        if (IsAdminLogin != true) {
	        	return false;
	        }
	    }
	}

	if (frm.returnmethod) {
		if (frm.returnmethod.value == "R000") {
			// 환불없음 이면 체크 안한다.
			frm.refundrequire.value = "0";
			return true;
		}
	}

	if (frm.refundrequire && frm.returnmethod) {
	    if ((frm.refundsubtotalprice.value*1 < 1) && ((frm.returnmethod.value != "R000"))) {
	        alert('환불대상 금액이 없습니다.\n\n환불없음 또는 쿠폰, 마일리지 등을 환원체크 해제하세요');
	        if (IsAdminLogin != true) {
	        	return false;
	        }
	    }
	}

	if (frm.remainsubtotalprice) {
	    if (frm.remainsubtotalprice.value*1 < 0) {
	        alert('취소 후 결제 금액이 마이너스가 될 수 없습니다. - 쿠폰이나 마일리지 환급을 체크해 주세요.');
	        if (IsAdminLogin != true) {
	        	return false;
	        }
	    }
	}

	return true;
}

// ============================================================================
// 선택상품 저장
// ============================================================================
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
			if (e.value*0 != 0) {
                alert('수량을 입력하세요.');
                e.focus();
                e.select();
                return false;
			}

			if ((IsStatusRegister == true) && ((e.value*1) == 0)) {
                alert('수량을 입력하거나 선택을 해제하세요.');
                e.focus();
                e.select();
                return false;
			}

			if ((IsStatusRegister != true) && ((e.value*1) < 0)) {
                alert('수량은 0 보다 작을 수 없습니다.');
                e.focus();
                e.select();
                return false;
			}

            regitemno = e.value;
        }

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
                upchedeliverPayStr = upchedeliverPayStr + "|" + e.value + "\t" + frm.gubun01.value + "\t" + frm.gubun02.value + "\t" + "1" + "\t" + frm.Deliveritemcost[i].value;
            }
        }
    }

    if ((upchedeliverPayStr.length>0)&&(frm.detailitemlist.value.length>0)){
        frm.detailitemlist.value = frm.detailitemlist.value + "|" + upchedeliverPayStr
    }
    //--------------------------------------------------

    //기타내역, 서비스발송 , 환불요청, 출고시유의사항, 업체 추가 정산 - 상세내역 체크 안함.
    if ((divcd=="A009") || (divcd=="A002") || (divcd=="A003") || (divcd=="A005") || (divcd=="A006") || (divcd=="A007") || (divcd=="A700") || (divcd=="A100") || (divcd=="A111")) {
        // no- check

    }else{
        if (!checkitemExists){
            alert('선택된 상세내역이 없습니다.');
            return false;
        }
    }

    return true;
}

// ============================================================================
// 선택상품 저장(상품변경 맞교환)
// ============================================================================
function SaveChangeCheckedItemList(frm) {
    var e;
    var ischecked = false;

	var prevoneorderdetailidx = "";		// 맞교환 회수하는 상품(order detail idx or 0(여러개인경우))
    var orderdetailidx = "";
    var changecsdetailidx = "";
    var reforderdetailidx = "";

    var gubun01 = "";
    var gubun02 = "";
    var regitemno = "";
    var causecontent = "";

    var itemid = "";
    var itemoption = "";

	if (IsStatusRegister) {
		// CS 접수창으로 접수할 수 없다.(옵션 변경창 또는 주문변경창에서 접수)
		alert("잘못된 접근입니다.");
		return false;
	}

    frm.detailitemlist.value = "";
    frm.csdetailitemlist.value = "";

	// 맞교환 회수하는 상품
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

        if (e.name == "changecsdetailidx") {
        	if (e.type != "checkbox") {
        		continue;
        	}

            if (e.checked == true) {
                ischecked = true;
                orderdetailidx = e.value;
                checkitemExists = true;
            }
        }

        if ((ischecked == true) && (e.name.indexOf("gubun01_") == 0)) {
            gubun01 = e.value;
        }

        if ((ischecked == true) && (e.name.indexOf("gubun02_") == 0)) {
            gubun02 = e.value;
        }

        if ((ischecked == true) && (e.name.indexOf("regitemno") == 0)) {
            regitemno = e.value;
        }

        if (e.name == "dummystopper") {
            if (ischecked == true) {
            	if (orderdetailidx != "") {
            		if (prevoneorderdetailidx == "") {
            			prevoneorderdetailidx = orderdetailidx;
            		} else {
            			prevoneorderdetailidx = "0";
            		}

            		frm.detailitemlist.value = frm.detailitemlist.value + "|" + orderdetailidx + "\t" + gubun01 + "\t" + gubun02 + "\t" + regitemno + "\t" + causecontent;
            	}

                ischecked = false;
                gubun01 = "";
                gubun02 = "";
                regitemno = "";
                causecontent = "";
            }
        }
    }

	// 맞교환 출고하는 상품
    for (var i = 0; i < frm.length; i++) {
        e = frm.elements[i];

        if (e.name == "dummystarter") {
            ischecked = false;
            changecsdetailidx = "";
            gubun01 = "";
            gubun02 = "";
            regitemno = "";
            itemid = "";
            itemoption = "";
        }

        if (e.name == "changecsdetailidx") {
        	if (e.type != "checkbox") {
        		continue;
        	}

            if (e.checked == true) {
                ischecked = true;
                changecsdetailidx = e.value;
                checkitemExists = true;
            }
        }

        if ((ischecked == true) && (e.name.indexOf("reforderdetailidx_") == 0)) {
            reforderdetailidx = e.value;
        }

        if ((ischecked == true) && (e.name.indexOf("gubun01_") == 0)) {
            gubun01 = e.value;
        }

        if ((ischecked == true) && (e.name.indexOf("gubun02_") == 0)) {
            gubun02 = e.value;
        }

        if ((ischecked == true) && (e.name.indexOf("itemid_") == 0)) {
            itemid = e.value;
        }

        if ((ischecked == true) && (e.name.indexOf("itemoption_") == 0)) {
            itemoption = e.value;
        }

        if ((ischecked == true) && (e.name.indexOf("regitemno") == 0)) {
        	if (e.value*1 < 0) {
        		changecsdetailidx = "";
        	}

            regitemno = e.value;
        }

        if (e.name == "dummystopper") {
            if (ischecked == true) {
            	if (changecsdetailidx != "") {
           			frm.csdetailitemlist.value = frm.csdetailitemlist.value + "|" + reforderdetailidx + "\t" + gubun01 + "\t" + gubun02 + "\t" + regitemno + "\t" + itemid + "\t" + itemoption;
            	}

                ischecked = false;
                gubun01 = "";
                gubun02 = "";
                regitemno = "";
	            itemid = "";
	            itemoption = "";
            }
        }
    }

    return true;
}

// ============================================================================
// 환불 수단이 적절한지
// ============================================================================
function CheckReturnMethod(frm) {
	if (!frm.returnmethod) { return true; }
	if (!frm.refundsubtotalprice) { return true; }

	var allselected = IsAllSelected(frm);
    var PayedNCancelEqual = (frm.orgsubtotalprice.value*1==frm.refundsubtotalprice.value*1);

    if (MainPaymentOrg > 0) {
    	if (frm.refundsubtotalprice.value*1 == MainPaymentOrg) {
    		PayedNCancelEqual = true;
    	} else {
    		PayedNCancelEqual = false;
    	}
    }

	if (frm.returnmethod.value == "R000") {
		// 환불없음 이면 체크 안한다.
		frm.refundrequire.value = "0";
		return true;
	}

    //2011-05-24 수정 / 기취소건 있을경우 frm.refundsubtotalprice.value*1<>MainPaymentOrg 이곳을통과 못함;;
    //if ((allselected) && (PayedNCancelEqual != true) && (IsCSCancelProcess)) {
    if ((allselected) && ((frm.orgsubtotalprice.value*1!=frm.refundsubtotalprice.value*1)) && (IsCSCancelProcess)) {
        if (!IsTicketOrder){
            alert('전체 취소인경우 결제금액 전체를 환불해야합니다. - 원배송비 환급, 마일리지, 할인권 등을 체크해주세요.');
            //alert(MainPaymentOrg);
            //alert(frm.orgsubtotalprice.value*1);
            //alert(frm.refundsubtotalprice.value*1);
            return false;
        }
    }

    if (((PayedNCancelEqual != true)) && ((frm.returnmethod.value=="R100") || (frm.returnmethod.value=="R550") || (frm.returnmethod.value=="R560") || (frm.returnmethod.value=="R020") || (frm.returnmethod.value=="R080") || (frm.returnmethod.value=="R400"))) {
        alert('일부취소 이후에는 사용할 수 없는 환불수단입니다.(신용카드 일부취소 등)\n\n[원 주결제금액 : ' + MainPaymentOrg + ']. \n\n부분취소 또는 다른 환불수단을 선택해 주세요');
        frm.returnmethod.focus();
        return false;
    }

    //원금액과 같은경우 전체취소를 해야함.
    if ((PayedNCancelEqual == true) && (frm.returnmethod.value=="R120")){
        alert('환불 금액과 원 주결제금액이 동일한 경우, 부분취소 사용불가  \n\n[원 주결제금액 : ' + MainPaymentOrg + ']. \n\n- 전체취소 사용요망.');
        frm.returnmethod.focus();
        return false;
    }

    //부분취소 관련 ALERT
    if (frm.returnmethod.value=="R120") {
        //alert(cardPartialCancelok + "," + frm.cardcode.value + "," + installment + "," + MainPaymentOrg + "," + precardcancelsum + "," + frm.refundrequire.value)
        if (cardPartialCancelok!="Y"){
            alert('부분 취소 가능 카드가 아닙니다.');
            return false;
        }
        // --1 BC 카드의 경우 90%까지만 부분취소 가능 // BC 카드의 경우 부분취소후 잔액이 5만원 미만이어도 원거래 할부가 그대로 적용됨.
        if (frm.cardcode.value=="11"){
            if (MainPaymentOrg*1!=0){
                if (precardcancelsum*1 + frm.refundrequire.value*1!=precardcancelsum*1){  //마지막 취소 (전체)는 상관없음.
                    if (((precardcancelsum*1 + frm.refundrequire.value*1)/MainPaymentOrg*1)>90){
                        alert('BC카드 의 경우 부분취소 합계액이 원금액의 90% 이상 이 될 수 없습니다. 다른 환불수단으로 처리하세요.');
                        return false;
                    }
                }
            }
        }

        // --// BC 카드가 아닌 경우 부분취소후 잔액이 5만원 미만이면, 일시불로 변경될 수 있음을 안내.
        // (국민, 외환, 신한, 현대, 삼성 - 매입후에는 원거래 할부가 그대로 적용됨. 롯데는 무조건 일시불로 변경.)
        if ((MainPaymentOrg*1>=50000)&&(MainPaymentOrg*1-(frm.refundrequire.value*1 + precardcancelsum*1)<50000)&&(installment*1>0)){
            //롯데.
            if (frm.cardcode.value=="03"){
                if (!confirm('롯데 카드의 경우 부분취소후 잔액이 5만원 미만이면 일시불로 전환됩니다.')){
                    return false;
                }
            }
            if (isThisdateCancel=="Y"){
                //당일취소
                if (!confirm('국민,외환,신한,삼성 카드의 경우 매입전 부분취소(당일 취소)후 잔액이 5만원 미만이면 일시불로 전환됩니다.')){
                    return false;
                }
            }
        }

    }

	if (sitename != "10x10") {
	    //외부몰인 경우 외부몰 환불접수만 가능..
	    if ((frm.returnmethod.value != "R050") && (frm.returnmethod.value != "R000")) {
	        alert('외부몰인 경우 환불 없음 또는 외부몰 환불을 선택하세요. \n\n제휴 담당자를 통해 제휴몰에서 취소 환불 처리 합니다.');
	        frm.returnmethod.focus();
	        return
	    }
	}

    if (frm.refundrequire.value*1 != frm.refundsubtotalprice.value*1) {
        if ((frm.returnmethod.value!="R007") && (frm.returnmethod.value!="R900") && (frm.returnmethod.value!="R910") && (frm.returnmethod.value!="R000")) {
            alert('환불 금액과 취소금액이 다를경우 무통장/마일리지환불/예치금환불 만 가능합니다.');
            return false;
        }

        if (!confirm('환불 금액이 취소 금액과 다르게 수정될경우 보정치로 금액이 입력됩니다.\n\n진행 하시겠습니까?')) {
            return false;
        }
    }

    //returnmethod R400 인경우 당월 취소만 가능함
    if (frm.returnmethod.value == "R400") {
    	if (IsOrderFound && (IsThisMonthJumun != true)) {
	        alert('휴대폰 결제는 당월 취소만 가능합니다. 다른 환불방식을 선택해 주세요.');
	        frm.returnmethod.focus();
	        return;
    	}
    }

    return true;
}

// ============================================================================
// 환불 관련 체크 Form
// ============================================================================
function CheckReturnForm(frm) {
    if (!frm.returnmethod) { return true; }
    if (!frm.refundrequire) { return true; }

    if (frm.returnmethod.value.length < 1) {
        alert('환불 방식을 선택해 주세요.');
        frm.returnmethod.focus();
        return false;
    }

	if (frm.returnmethod.value == "R000") {
		// 환불없음 이면 체크 안한다.
		frm.refundrequire.value = "0";
		return true;
	}

	if (frm.refundrequire.value*0 != 0) {
        alert('환불 금액을 입력하세요.');
        frm.refundrequire.focus();
        return;
	}

	if ((frm.refundrequire.value*1 <= 0) && (frm.returnmethod.value != "R000")) {
		alert('환불 금액은 0 보다 커야합니다. 또는 환불없음을 선택하세요.');
        return false;
	}

    if ((frm.returnmethod.value=="R100") || (frm.returnmethod.value=="R550") || (frm.returnmethod.value=="R560") || (frm.returnmethod.value=="R020") || (frm.returnmethod.value=="R080") || (frm.returnmethod.value=="R400")) {
    	if (MainPaymentOrg > 0) {

    		if (frm.refundsubtotalprice) {
	    		if (frm.refundsubtotalprice.value*1 != MainPaymentOrg) {
	    		    //alert(MainPaymentOrg);
	    		    //alert(frm.refundsubtotalprice.value*1);
			        alert('일부취소 이후에는 사용할 수 없는 환불방식입니다.(신용카드 일부취소 등)\n\n[원 주결제금액 : ' + MainPaymentOrg + ']. \n\n다른 환불수단을 선택해 주세요..');
			        frm.returnmethod.focus();
			        return false;
				}
    		} else {
	    		if (frm.refundrequire.value*1 != MainPaymentOrg) {
	    		    //alert(MainPaymentOrg);
	    		    //alert(frm.refundsubtotalprice.value*1);
			        alert('일부취소 이후에는 사용할 수 없는 환불방식입니다.(신용카드 일부취소 등)\n\n[원 주결제금액 : ' + MainPaymentOrg + ']. \n\n다른 환불수단을 선택해 주세요..');
			        frm.returnmethod.focus();
			        return false;
				}
    		}
		}
    }


	// ====================================================================
	if (frm.returnmethod.value=="R007") {
		// 무통장
        var mooconfirm = false;
        if ((frm.rebankaccount) && (frm.rebankaccount.value.length < 1)) {
            mooconfirm = true;
        }

        if ((frm.rebankownername) && (frm.rebankownername.value.length < 1)) {
            mooconfirm = true;
        }

        if ((frm.rebankname) && (frm.rebankname.value.length < 1)) {
            mooconfirm = true;
        }

        if (mooconfirm == true) {
        	// 무통장 계좌정보는 나중에 별도로 입력할 수 있다.
            if (!confirm('환불 계좌가 없습니다. \n\n환불 계좌 없이 등록 하시겠습니까?')) {
                if ((IsStatusRegister == true) || (IsStatusEdit == true)) {
                	frm.rebankaccount.focus();
                }
                return false;
            }
        }
	}

	if (frm.returnmethod.value == "R900") {
    	if (confirm("CS서비스가 아닌경우(결제금액환불) 마일리지 대신 예치금으로 환불하세요.\n\n마일리지 환불 하시겠습니까?") != true) {
    		return false;
    	}
	}

	// ====================================================================
	if ((frm.returnmethod.value=="R900") || (frm.returnmethod.value=="R910")) {
		// 마일리지, 예치금환불
        if ((frm.refund_userid) && (frm.refund_userid.value.length<1)) {
            alert('비회원에게 가능하지 않은 환불방식입니다. 다른 환불 방식을 선택하세요.');
            return false;
        }
	}

    return true;
}

// ============================================================================
// 삭제
// ============================================================================
function CsRegCancelProc(frm) {
    if (confirm('등록된 접수 내역을 삭제 하시겠습니까?')){
        frm.mode.value = "deletecsas";
        frm.submit();
    }
}

// ============================================================================
// 완료처리
// ============================================================================
function CsRegFinishProc(frm) {
	// 마스터 체크
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	// 선택 상품 체크 및 저장
	if ((divcd != "A100") && (divcd != "A111")) {
	    if (!SaveCheckedItemList(frm)) {
	        return;
	    }
	} else {
	    if (!SaveChangeCheckedItemList(frm)) {
	        return;
	    }
	}

	// 취소, 반품, 환불
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true) || (divcd == "A003")) {
        if (CheckReturnForm(frm) != true) {
            return;
        }
	}

	// 취소, 반품
    if ((IsCSCancelProcess) || (IsCSReturnProcess)) {
		// 환불정보가 정상인지 체크
		if (IsRefundInfoOK(frm) != true) {
			return;
		}

		// 환불수단이 옳바른지
		if (CheckReturnMethod(frm) != true) {
			return;
		}
    }

	if (IsStatusFinishing && (divcd == "A007" || ((divcd == "A003") && (frm.returnmethod.value=="R007")))) {
		if (IsAdminLogin) {
			alert('이곳에서 완료처리 하여도 신용카드 승인취소/무통장 환불처리는 이루어 지지 않습니다.[어드민권한]');
		} else {
			alert('이곳에서 완료처리 하여도 신용카드 승인취소/무통장 환불처리는 이루어 지지 않습니다.\n\n완료 처리 할 수 없습니다.');
			return;
		}
	}

    //환불요청 , 신용카드 취소요청
    if ((divcd == "A003") || (divcd == "A007")) {
        if (frm.contents_finish.value.length<1){
            alert('처리 내용을 입력하세요.');
            frm.contents_finish.focus();
            return;
        }
    }

    var confirmMsg ;
    confirmMsg = '완료처리 진행 하시겠습니까?';

    if ((divcd == "A004") || (divcd == "A010")) {
        confirmMsg = '완료처리 진행시 마이너스 주문 및 환불이 자동 접수됩니다. 진행 하시겠습니까?';
    }

    if (confirm(confirmMsg )) {
        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "";
        frm.submit();
    }
}

// ============================================================================
// 업체처리완료=>접수 변경
// ============================================================================
function CsUpcheConfirm2RegProc(frm) {
    if (confirm('접수 상태로 변경 하시겠습니까?')){
        frm.mode.value = "upcheconfirm2jupsu";
        frm.submit();
    }
}

// ============================================================================
// 수정
// ============================================================================
function CsRegEditProc(frm) {

	// 마스터 체크
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	// 선택 상품 체크 및 저장
	if ((divcd != "A100") && (divcd != "A111")) {
	    if (!SaveCheckedItemList(frm)) {
	        return;
	    }
	} else {
	    if (!SaveChangeCheckedItemList(frm)) {
	        return;
	    }
	}

	if (divcd == "A111") {
		if (frm.customerrealbeasongpay) {
			if (frm.customerrealbeasongpay.value == "") {
				frm.customerrealbeasongpay.value = "0";
			}

			if (frm.customerrealbeasongpay.value*0 != 0) {
				alert("금액은 숫자로 입력해야 합니다.");
				frm.customerrealbeasongpay.focus();
				return;
			}
		}
	}

	// 취소, 반품, 환불
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true) || (divcd == "A003")) {
        if (CheckReturnForm(frm) != true) {
            return;
        }
	}

	// 취소, 반품
    if ((IsCSCancelProcess) || (IsCSReturnProcess)) {
		// 환불정보가 정상인지 체크
		if (IsRefundInfoOK(frm) != true) {
			return;
		}

		// 환불수단이 옳바른지
		if (CheckReturnMethod(frm) != true) {
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

    if (confirm('수정 하시겠습니까?')){
        frm.mode.value = "editcsas";
        frm.submit();
    }
}

// ============================================================================
// 환불요청 없는 완료 처리
// ============================================================================
function CsRegFinishProcNoRefund(frm){
    var divcd = frm.divcd.value;

    if (confirm('환불 및 마이너스 등록 없이 완료처리 진행 하시겠습니까?')){
        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "norefund";
        frm.submit();
    }
}

// ============================================================================
// 체크된 상품/배송비 색바꾸기
// ============================================================================
function AnCheckClickAll(frm) {
	if (frm.Deliverdetailidx) {
	    for(var i = 0; i < frm.Deliverdetailidx.length; i++) {
			AnCheckClick(frm.Deliverdetailidx[i]);
	    }
	}

	if (frm.orderdetailidx) {
	    for(var i = 0; i < frm.orderdetailidx.length; i++) {
			AnCheckClick(frm.orderdetailidx[i]);
	    }
	}
}

// ============================================================================
// 업배 텐텐회수 인가
// ============================================================================
function SetForceReturnByTen(frm) {
	var e;

	if ((IsStatusRegister == true) || (divcd != "A010")) {
		return;
	}

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if (e.type != "checkbox") {
       		continue;
        }

    	if (e.checked == true) {
    		if (frm.isupchebeasong[i].value == "Y") {
    			frm.ForceReturnByTen.checked = true;
    			return;
    		}
        }
    }
}

// ============================================================================
// 텐텐물류로 고객 직접반품인가
// ============================================================================
function SetForceReturnByCustomer(frm) {
	var e;

	if ((IsStatusRegister == true) || (divcd != "A004")) {
		return;
	}

	// requiremakerid 가 빈값이면 텐텐회수, requiremakerid = 10x10logistics 이면 텐텐물류 고객반품, 기타 업체반품
	if (frm.requiremakerid.value == "10x10logistics") {
		frm.ForceReturnByCustomer.checked = true;
	} else {
		frm.ForceReturnByCustomer.checked = false;
	}
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

// 티켓주문인경우..
function calcuTicketCancelCharge(comp){
    var frm = comp.form;
    var mayTicketCancelPro = comp.value*1;

    if (mayTicketCancelPro>0){
        alert('티켓 취소 수수료 ' + mayTicketCancelPro + '% 차감' );
    }else{
        alert('티켓 취소 수수료 차감 안함' );
    }
    frm.refundadjustpay.value = (frm.refundtotalbuypaysum.value*mayTicketCancelPro/100)*-1;

    CheckForItemChanged();
}

// ============================================================================
// 선택 상품 전부 텐배 상품 인지
// ============================================================================
function IsAllTenbaeItem(frm) {
    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if (e.type != "checkbox") {
        	continue;
        }

    	if (e.checked == true) {
    		if (frm.isupchebeasong[i].value == "Y") {
    			return false;
    		}
        }
    }

    return true;
}

// ============================================================================
// 고객 직접반품 전환(A010 -> A004)
// ============================================================================
function ChangeDivcdToA004(frm) {
	if (IsDeletedCS) {
		alert("삭제된 내역입니다.");
		return;
	}

	if (IsLogicsSended) {
		alert("이미 택배사에 전송된 내역입니다.");
		return;
	}

	if (IsStatusEdit != true) {
		alert("수정할 수 없습니다.");
		return;
	}

    if (confirm('고객 직접반품으로 전환 하시겠습니까?')){
        frm.mode.value = "changedivcdtoa004";
        frm.submit();
    }
}

// ============================================================================
// 회수신청 전환(A004 -> A010)
// ============================================================================
function ChangeDivcdToA010(frm) {
	if (IsDeletedCS) {
		alert("삭제된 내역입니다.");
		return;
	}

	if (IsLogicsSended) {
		alert("이미 택배사에 전송된 내역입니다.");
		return;
	}

	if (IsStatusEdit != true) {
		alert("수정할 수 없습니다.");
		return;
	}

	if (frm.buf_requiremakerid) {
		if (frm.buf_requiremakerid.value != "10x10logistics") {
			alert("업체배송 상품입니다. 수정할 수 없습니다.");
			return;
		}
	} else {
		alert("잘못된 접근입니다.");
		return;
	}

	if (IsAllTenbaeItem(frm) != true) {
		alert("업체배송 상품입니다. 수정할 수 없습니다.");
		return;
	}

    if (confirm('회수신청으로 전환 하시겠습니까?')){
        frm.mode.value = "changedivcdtoa010";
        frm.submit();
    }
}
