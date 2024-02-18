
/* global IsPossibleModifyCSMaster, IsPossibleModifyItemList, ERROR_MSG_TRY_MODIFY, initializeURL, initializeReturnFunction, initializeErrorFunction, startRequest, $ */

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

	// 접수상품 금액 재계산
    CalculateAndApplyItemCostSum(frm);

	if (IsStatusRegister == true) {
		if (IsOnlyOneBrandAvailable == true) {
		    // 단일 브랜드만 선택 가능하게, 담당 브랜드 저장

	    	EnableOnlyOneBrand(frm);

			// 업배 반품/맞교환 : 업체 추가정산 관련 브랜드 아이디 가져오기
			InsertCheckedUpcheID(frm);
		}
	}

	// ========================================================================
	// 반품시 체크
	// ========================================================================
	if (IsCSReturnProcess == true) {
		// 원주문수량 이상의 반품이 있는지 체크
		CheckOverReturnItemno(frm);

		if (IsStatusRegister == true) {
			if (IsOnlyOneBrandAvailable == true) {
			    // 단일 브랜드만 선택 가능하게, 담당 브랜드 저장

				// 업배 반품/맞교환 : 업체 추가정산 관련 브랜드 아이디 가져오기
				// InsertCheckedUpcheID(frm);
			}
		}

		// 배송비 차감(단순변심)
		CheckBeasongPayCut(frm);

		// 반품시 회수 배송비 계산
		CalculateReturnBeasongPay(frm);
	}

	if (IsStatusRegister == true) {
		if (IsOnlyOneBrandAvailable == true) {
		    // 단일 브랜드만 선택 가능하게, 담당 브랜드 저장

		    if (IsStatusRegister == true) {
		    	// EnableOnlyOneBrand(frm);
		    }

			// 업배 반품/맞교환 : 업체 추가정산 관련 브랜드 아이디 가져오기
			// InsertCheckedUpcheID(frm);
		}

		// 체크된 상품/배송비 색바꾸기
		// AnCheckClickAll(frmaction);

		// 주문전체 선택시 마일리지, 할인권 자동체크
		CheckMileageETC(frm);

		// 쿠폰 재발급 체크
		jsCheckCopyCoupon(frm);
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
	if ((IsCSReturnProcess != true) || (IsStatusFinished == true)) {
		return;
	}

	var arrorderdetailidx 		= document.getElementsByName("orderdetailidx");
	var arrregitemno 			= document.getElementsByName("regitemno");
	var arritemno 				= document.getElementsByName("itemno");
	var arrprevcsreturnfinishno = document.getElementsByName("prevcsreturnfinishno");

	if (arrorderdetailidx.length < 1) {
		return;
	}

	for (var i = 0; i < arrorderdetailidx.length; i++) {
        if (arrorderdetailidx[i].type != "checkbox") {
        	continue;
        }

        if (arrorderdetailidx[i].checked != true) {
        	continue;
        }

		if ((arrregitemno[i].value*1 + arrprevcsreturnfinishno[i].value*1) > arritemno[i].value*1) {
			alert("주문수량을 초과하여 접수되는 상품이 있습니다.");

			if ((arrregitemno[i].value*1 - arrprevcsreturnfinishno[i].value*1)  > 0) {
				arrregitemno[i].value = (arrregitemno[i].value*1 - arrprevcsreturnfinishno[i].value*1);
			} else {
				arrorderdetailidx[i].checked = false;
				AnCheckClick(arrorderdetailidx[i]);
				delGubun("gubun01_" + arrorderdetailidx[i].value,"gubun02_" + arrorderdetailidx[i].value,"gubun01name_" + arrorderdetailidx[i].value,"gubun02name_" + arrorderdetailidx[i].value, frm.name, causepop);
			}
		}
	}
}

// ============================================================================
// 디테일 상품 입력수량 체크
// ============================================================================
function CheckMaxItemNo(obj, maxno) {
	var frm = document.frmaction;

	if (obj.value*0 != 0) {
		return;
	}

    if (obj.value*1 > maxno*1) {
        alert("주문수량 이상으로 상품수량을 수정할수 없습니다.");
        obj.value = maxno;
    }

	// ========================================================================
	// 상품취소,반품시 배송비 동시취소(또는 마이너스 주문등록)
	// ========================================================================
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true)) {
	    CheckUpcheDeliverPay(frm);

		// 상품 전체 체크 된 경우 체크안된 배송비 동시체크
		if (IsCSCancelProcess == true) {
			CheckBeasongPayIfAllItemSelected(frm);
		}
	}

	// 상품변경시 체크
	CheckForItemChanged();
}

function CheckBeasongPayOnlyRemain(frm) {
    var upbeaMakerid;
	var cancelyn;

    var objdeliver, objitem;

    if (!frm.orderdetailidx) return true;
    if ((IsCSCancelProcess != true) && (IsCSReturnProcess != true)) return true;

	// 브랜드 기준 배송비 체크
    for (var i = 0; i < frm.orderdetailidx.length; i++) {
		if ((frm.itemid[i].value != "") && (frm.itemid[i].value*1 == 0)) {
			objdeliver = frm.orderdetailidx[i];

			if (objdeliver.type != "checkbox") {
        		continue;
			}

			upbeaMakerid = frm.makerid[i].value;
			cancelyn = frm.cancelyn[i].value;

			/*
			if (cancelyn != "Y") {
				if ((IsOneBrandAllSelected(frm, upbeaMakerid) == true) && (objdeliver.checked == false)) {
					alert("브랜드 전체취소(반품)이면서 배송비가 선택되지 않았습니다.");
					return false;
				}
			}
			*/
		}
	}

	// 전체취소
	/*
	if (IsCSCancelProcess == true) {
		if (IsAllSelected(frm) == true) {
			for (var i = 0; i < frm.orderdetailidx.length; i++) {
				if ((frm.itemid[i].value != "") && (frm.itemid[i].value*1 == 0)) {
					objdeliver = frm.orderdetailidx[i];

					if (objdeliver.type != "checkbox") {
        				continue;
					}

					cancelyn = frm.cancelyn[i].value;

					if ((cancelyn != "Y") && (objdeliver.checked == false)) {
						alert("주문 전체취소이면서 선택않된 배송비가 있습니다.");
						return false;
					}
				}
			}
		}

	}
	*/

	return true;
}

// ============================================================================
// 상품선택시 확인할 것들
// ============================================================================
function CheckSelect(comp, isbeasongpay) {
    var chkidx = comp.value;
    var frm = document.frmaction;
	var objdeliver;

	if (isbeasongpay == false) {
		// ========================================================================
		// 상품취소,반품시 배송비 동시취소(또는 마이너스 주문등록)
		// ========================================================================
		if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true)) {
			CheckUpcheDeliverPay(frm);
		}

		// 상품 전체 체크 된 경우 체크안된 배송비 동시체크
		if (IsCSCancelProcess == true) {
			CheckBeasongPayIfAllItemSelected(frm);
		}
	}

	// 배송비 취소의 경우 출고상품이 있는지 체크
	if (IsCSCancelProcess && (isbeasongpay == true)) {
		for (var i = 0; i < frm.orderdetailidx.length; i++) {
			if ((frm.itemid[i].value != "") && (frm.itemid[i].value*1 == 0)) {
				objdeliver = frm.orderdetailidx[i];

				if (objdeliver.type != "checkbox") {
        			continue;
				}

				if (objdeliver.checked != true) {
        			continue;
				}

				if (CheckChulgoItemExist(frm, frm.makerid[i].value) == true) {
					if (IsAdminLogin == true) {
						if (confirm("[관리자 권한]\n\n출고상품이 있습니다.\n배송비를 취소하시겠습니까??") != true) {
							objdeliver.checked = false;
							AnCheckClick(frm.orderdetailidx[i]);
							return;
						}
					} else {
						alert("출고상품이 있는 경우 배송비를 취소할 수 없습니다.");
						objdeliver.checked = false;
						AnCheckClick(frm.orderdetailidx[i]);
						return;
					}
				}
			}
		}
	}

	if (comp.checked){
		// CS 사유구분 복사
		eval("frm.gubun01_" + chkidx).value = frm.gubun01.value;
		eval("frm.gubun02_" + chkidx).value = frm.gubun02.value;
		eval("frm.gubun01name_" + chkidx).value = frm.gubun01name.value;
		eval("frm.gubun02name_" + chkidx).value = frm.gubun02name.value;
	}else{
		delGubun("gubun01_" + chkidx,"gubun02_" + chkidx,"gubun01name_" + chkidx,"gubun02name_" + chkidx,frm.name,causepop);
	}

	// 상품변경시 체크
	CheckForItemChanged(comp);
}

// ============================================================================
// 배송비 동시 체크
// ============================================================================
//
// 이미 배송비 환급이 체크되어 있으면 변경 않한다.
//
// 취소,반품시
// 텐배   - 텐배 브랜드 전체가 선택된 경우, 배송비 체크한다.
// 업배   - 업배 브랜드 전체가 선택된 경우, 배송비 체크한다.(변심반품은 배송비차감으로 처리한다.)
//
// ============================================================================
function CheckUpcheDeliverPay(frm) {
    var upbeaMakerid;
	var cancelyn;

    var objdeliver, objitem;

    if (!frm.orderdetailidx) return;
    if ((IsCSCancelProcess != true) && (IsCSReturnProcess != true)) return;

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
		if ((frm.itemid[i].value != "") && (frm.itemid[i].value*1 == 0)) {
			objdeliver = frm.orderdetailidx[i];

			if (objdeliver.type != "checkbox") {
        		continue;
			}

			upbeaMakerid = frm.makerid[i].value;
			cancelyn = frm.cancelyn[i].value;

			if (cancelyn != "Y") {
				if (IsOneBrandAllSelected(frm, upbeaMakerid) == true) {
					objdeliver.checked = true;
				} else {
					objdeliver.checked = false;
				}
			}

			if (objdeliver.checked){
				// CS 사유구분 복사
				eval("frm.gubun01_" + objdeliver.value).value = frm.gubun01.value;
				eval("frm.gubun02_" + objdeliver.value).value = frm.gubun02.value;
				eval("frm.gubun01name_" + objdeliver.value).value = frm.gubun01name.value;
				eval("frm.gubun02name_" + objdeliver.value).value = frm.gubun02name.value;
			}else{
				delGubun("gubun01_" + objdeliver.value,"gubun02_" + objdeliver.value,"gubun01name_" + objdeliver.value,"gubun02name_" + objdeliver.value,frm.name,causepop);
			}

			AnCheckClick(frm.orderdetailidx[i]);
		}
    }
}

// 배송완료 상품 있는지(텐배, 업배 브랜드별)
// 출고상품 있는 경우 배송비 취소 불가
function CheckChulgoItemExist(frm, makerid) {
    for (var i = 0; i < frm.orderdetailidx.length; i++) {
		if ((frm.itemid[i].value != "") && (frm.itemid[i].value*1 != 0)) {
			if (makerid == "") {
				if ((frm.isupchebeasong[i].value == "N") && (frm.orderdetailcurrstate[i].value == "7")) {
					return true;
				}
			} else {
				if ((frm.makerid[i].value == makerid) && (frm.orderdetailcurrstate[i].value == "7")) {
					return true;
				}
			}
		}
	}

	return false;
}

function ForceCheckUpcheDeliverPay(frm) {
	CheckUpcheDeliverPay(frm);
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
			if ((upbeaMakerid.length < 1) && (frm.odlvtype[i].value == "9") && (frm.makerid[i].value == "")) {
				alert("ERROR ERROR!!\n\n시스템팀문의!! ===============");
			}
	        if (
	        	((upbeaMakerid.length < 1) && (frm.odlvtype[i].value != "1") && (frm.odlvtype[i].value != "4") && (frm.makerid[i].value != ""))
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
// 선택된 상품에서 특정 브랜드 있는지
// ============================================================================
function GetCheckedItemMakerExist(frm, makerid) {
    var objitem;

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        objitem = frm.orderdetailidx[i];

        if (objitem.type != "checkbox") {
        	continue;
        }

        if (objitem.checked) {
			if (makerid == frm.makerid[i].value) {
				return true;
			}
        }
    }

    return false;
}

function jsCheckApplyEvent(frm) {
	if (IsTempEventAvail != true) {
		alert("ERR : 적용불가");
		return;
	}

	if (GetCheckedItemMakerExist(frm, IsTempEventAvail_Makerid) == true) {
		if (((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD01")) || ((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD06"))) {
			selectGubun('C004','CD11','공통','무료반품','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');
			return;
		} else {
			alert("변심반품일 경우에만 적용하세요.\n\n(무료반품은 주문당 1회만 적용가능!!)");
			return;
		}
		//
	} else {
		alert("이벤트 브랜드(" + IsTempEventAvail_Makerid + ") 상품이 선택되지 않았습니다.");
		return;
	}
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
// 쿠폰 재발행 체크
// ============================================================================
function jsCheckCopyCoupon(frm) {

	if (!frm.tmpcopycouponinfo) {
		return;
	}

	if (frm.tmpcopycouponinfo.disabled == true) {
		return;
	}

	if (IsCSCancelProcess == true) {
		var allselected = IsAllSelected(frm);

		if (allselected && frm.tmpcopycouponinfo.checked == true) {
			alert("재발급불가!!\n\n전체 취소의 경우 사용한 쿠폰이 환원됩니다.");
			frm.tmpcopycouponinfo.checked = false;
		}
	}

	if (frm.tmpcopycouponinfo.checked == true) {
		frm.copycouponinfo.value = "Y";
	} else {
		frm.copycouponinfo.value = "N";
	}
}

// ============================================================================
// 배송비 차감
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
	IsCkReturnPayChanged = true;

	// ========================================================================
	// 단순변심(CD01)이 아니면 차감 없음, 사이즈않맞음(CD06)추가
	if ((value_gubun02 != "CD01")&&(value_gubun02 != "CD06")) {
		return;
	}

	if (IsChangeOrder) {
		// 교환주문은 상품금액,배송비 금액으로 계산한다.
		if ((frm.orgitemcostsum.value*1 == frm.refunditemcostsum.value*1) && (frm.orgbeasongpay.value*1 == 0)) {
			// 브랜드 전체상품선택, 업체무료배송
			frm.ckreturnpay.checked = true;
		} else {
			frm.ckreturnpayHalf.checked = true;
		}

		return;
	}

	// ========================================================================
    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if (e.type != "checkbox") {
       		continue;
        }

    	if (e.checked == true) {
			if (frm.isupchebeasong[i].value != "Y") {
				makerid = "";
			} else {
				makerid = frm.makerid[i].value;
			}

			// 브랜드 전체선택이이면 왕복배송비 차감
			// 텐텐회수인경우 체크되어 있는 전브랜드 확인하여 한개의 브랜드라도 4000원 차감 조건을 충족하면 4000원 차감한다.
			// 업체반품의 경우도 텐텐 고객반품이 있으므로 모두 체크한다.
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
// (배송비는 체크안함)
// ============================================================================
function IsAllSelected(frm) {
	var orderdetailidx, cancelyn, itemid;
	var prevcsreturnfinishno, itemno, regitemno;

	if (!frm.orderdetailidx) return allselected;

    for (var i = 0; ; i++) {
		orderdetailidx = document.getElementById("orderdetailidx_" + i);
		cancelyn = document.getElementById("cancelyn_" + i);
		itemid = document.getElementById("itemid_" + i);
		prevcsreturnfinishno = document.getElementById("prevcsreturnfinishno_" + i);
		itemno = document.getElementById("itemno_" + i);
		regitemno = document.getElementById("regitemno_" + i);

		if (orderdetailidx === null) {
			break;
		}

		if (cancelyn.value === "Y") { continue; }
		if (parseInt(itemid.value,10) === 0) { continue; }
		if (orderdetailidx.checked !== true) { return false; }

		// 기존 반품 포함 수량체크
		if (parseInt(itemno.value) !== (parseInt(regitemno.value) + parseInt(prevcsreturnfinishno.value))) { return false; }
    }

    return true;
}

// ============================================================================
// 상품 전체 체크 된 경우 체크안된 배송비 동시체크
// ============================================================================
function CheckBeasongPayIfAllItemSelected(frm) {
    var allselected = false;
	var e;

	if (!frm.orderdetailidx) return;

	if (IsCSCancelProcess != true) {
		return;
	}

	if (IsAllSelected(frm) == true) {
		for (var i = 0; i < frm.orderdetailidx.length; i++) {
			e = frm.orderdetailidx[i];

			if (e.type != "checkbox") {
        		continue;
			}

			if (frm.cancelyn[i].value == "Y") {
				continue;
			}

			if (frm.itemid[i].value*1 != 0) {
				continue;
			}

			if (e.checked == true) {
                continue;
			}

			if (frm.prevcsreturnfinishno[i].value*1 != 0) {
				continue;
			}

			e.checked = true;
			AnCheckClick(frm.orderdetailidx[i]);
			// CS 사유구분 복사
			eval("frm.gubun01_" + frm.orderdetailidx[i].value).value = frm.gubun01.value;
			eval("frm.gubun02_" + frm.orderdetailidx[i].value).value = frm.gubun02.value;
			eval("frm.gubun01name_" + frm.orderdetailidx[i].value).value = frm.gubun01name.value;
			eval("frm.gubun02name_" + frm.orderdetailidx[i].value).value = frm.gubun02name.value;
		}
	}
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
function IsOneBrandAllSelected(frm, targetmakerid) {
	var onebrandallselected = false;
	var checkeditemexist = false;
	var orderdetailidx, cancelyn, itemid;
	var prevcsreturnfinishno, itemno, regitemno, isupchebeasong, makerid;

	if (!frm.orderdetailidx) return allselected;

    for (var i = 0; ; i++) {
		orderdetailidx = document.getElementById("orderdetailidx_" + i);
		cancelyn = document.getElementById("cancelyn_" + i);
		itemid = document.getElementById("itemid_" + i);
		prevcsreturnfinishno = document.getElementById("prevcsreturnfinishno_" + i);
		itemno = document.getElementById("itemno_" + i);
		regitemno = document.getElementById("regitemno_" + i);
		isupchebeasong = document.getElementById("isupchebeasong_" + i);
		makerid = document.getElementById("makerid_" + i);

		if (orderdetailidx === null) {
			break;
		}

		if (targetmakerid === "") {
			// 텐배체크
			if (isupchebeasong.value === "Y") { continue; }
		} else {
			// 업배체크
			if ((isupchebeasong.value !== "Y") || (targetmakerid !== makerid.value)) { continue; }
		}

		if (parseInt(itemid.value,10) === 0) { continue; }
		if (cancelyn.value === "Y") {
			onebrandallselected = true;
			continue;
		}

		if (orderdetailidx.checked !== true) {
			if (IsCSReturnProcess != true) { return false; }
			if (parseInt(itemno.value,10) !== parseInt(prevcsreturnfinishno.value,10)) { return false; }
		} else {
			checkeditemexist = true;
			if (parseInt(itemno.value,10) !== (parseInt(regitemno.value,10) + parseInt(prevcsreturnfinishno.value,10))) { return false; }
		}

		onebrandallselected = true;
    }

    return onebrandallselected && checkeditemexist;
}

/*
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

		if (frm.itemid[i].value*1 == 0) {
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
*/

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

// ============================================================================
// 선택된 상품수
// ============================================================================
function GetSelectedItemNO(frm) {
	var regitemnoSUM = 0;

    if (!frm.orderdetailidx) { return regitemnoSUM; }

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if (e.type != "checkbox") {
        	continue;
        }

		if (e.checked != true) {
			continue;
		}

		if (frm.itemid[i].value == "0") {
			continue;
		}

		regitemnoSUM = regitemnoSUM + frm.regitemno[i].value*1;
    }

    return regitemnoSUM;
}

// ============================================================================
// 두개 이상의 상품을 품절 등록하는지
// ============================================================================
function IsMultiStockOutItemChecked(frm) {
    var stockoutitemcount = 0;
    var ele;
    var ischecked, gubun01, gubun02;

    for (var i = 0; i < frm.length; i++) {
        ele = frm.elements[i];

        if (ele.name == "dummystarter") {
            ischecked = false;
            gubun01 = "";
            gubun02 = "";
        }

        if (ele.name == "orderdetailidx") {
        	if (ele.type != "checkbox") {
        		continue;
        	}

            if (ele.checked == true) {
                ischecked = true;
            }

            continue;
        }

		if (ischecked != true) {
			continue;
		}

        if (ele.name.indexOf("gubun01_") == 0) {
            gubun01 = ele.value;
            continue;
        }

        if (ele.name.indexOf("gubun02_") == 0) {
            gubun02 = ele.value;
            continue;
        }

        if (ele.name == "dummystopper") {
        	if ((gubun01 == "C004") && (gubun02 == "CD05")) {
        		stockoutitemcount = stockoutitemcount + 1;

        		if (stockoutitemcount >= 2) {
        			return true;
        		}
        	}
        }
    }

    return false;
}

function CalculateAndApplyItemCostSum(frm) {
	// 환불대상 상품가격(상품쿠폰적용가)합계, 올엣카드할인합계, 비율쿠폰할인합계
	CalculateCancelItemSUM(frm);

	// 배송비 합계 계산
	CalculateBeasongPaySum(frm);

	// 구매총액 계산
	CalculateTotalBuyPaySum(frm);

	// 반품시 회수 배송비 계산
	// CalculateReturnBeasongPay(frm);

	// 업체 추가정산
	CalculateUpcheReturnBeasongPay(frm);

	// 정액쿠폰 환원 재계산, 환불 쿠폰합계 재계산
	// 정액쿠폰 없음, 샹품별로 안분되어 비율쿠폰과 동일함. 2018-04-19
	// CalculateFixedCoupon(frm);

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

	if ((IsStatusFinished != true) && (IsStatusFinishing != true)) {
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

	if ((IsStatusFinished != true) && (IsStatusFinishing != true)) {
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
	var itemid          	= 0;

    var itemcost        				= 0;
    var allatitemdiscount				= 0;
    var percentBonusCouponDiscount 		= 0;
	var etcDiscountDiscount		 		= 0;

    var orgitemcostsum     				= 0;
    var refunditemcostsum   			= 0;

    var orgallatitemdiscountSum 		= 0;
    var refundallatitemdiscountSum 		= 0;

    var orgpercentBonusCouponDiscountSum 		= 0;
    var refundpercentBonusCouponDiscountSum 	= 0;
    var orgetcDiscountDiscountSum 		= 0;
    var refundetcDiscountDiscountSum 	= 0;

	// ========================================================================
    for (var i = 0; i < frm.length; i++) {
        e = frm.elements[i];

        if (e.name == "dummystarter") {
            ischecked = false;
            regitemno = 0;
            itemno = 0;
            itemcost = 0;
			itemid = 0;
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

        if (e.name == "itemid") {
            itemid = e.value;
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

        if (e.name == "etcDiscountDiscount") {
            if ((e.value * 0) == 0) {
                etcDiscountDiscount = e.value;
            } else {
                etcDiscountDiscount = 0;
            }
        }

        if (e.name == "dummystopper") {
			// 선택상품합계
			if (ischecked == true) {
				if (itemid*1 != 0) {
					refunditemcostsum 					= refunditemcostsum + (itemcost * regitemno * 1);
					refundallatitemdiscountSum 			= refundallatitemdiscountSum + (allatitemdiscount * regitemno * 1);
				}

				refundpercentBonusCouponDiscountSum = refundpercentBonusCouponDiscountSum + (percentBonusCouponDiscount * regitemno * 1);
				refundetcDiscountDiscountSum = refundetcDiscountDiscountSum + (etcDiscountDiscount * regitemno * 1);
			}

			// 취소안된 상품 전체합계
			if (cancelyn != "Y") {
				if (itemid*1 != 0) {
					orgitemcostsum 						= orgitemcostsum + (itemcost * itemno * 1);
					orgallatitemdiscountSum 			= orgallatitemdiscountSum + (allatitemdiscount * itemno * 1);
				}

				orgpercentBonusCouponDiscountSum 	= orgpercentBonusCouponDiscountSum + (percentBonusCouponDiscount * itemno * 1);
				orgetcDiscountDiscountSum 	= orgetcDiscountDiscountSum + (etcDiscountDiscount * itemno * 1);
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
		frm.refundcouponsum.value = frm.refundpercentcouponsum.value;
    }

    // 선택안한(나머지) 비율쿠폰 총액
    if (frm.remainpercentcouponsum!=undefined){
		if (IsCSReturnProcess && IsStatusEdit) {
			// 반품접수 후 제휴몰 쿠폰금액 입력된 케이스
			if ((frm.tencardspend.value*1 == 0) && (frm.orgcouponsum.value*1 == 0) && (frm.orgpercentcouponsum.value*1 == 0) && (frm.refundpercentcouponsum.value*1 < 0)) {
				frm.tencardspend.value = frm.refundpercentcouponsum.value*-1;
				frm.orgcouponsum.value = frm.refundpercentcouponsum.value*1;
				frm.orgpercentcouponsum.value = frm.refundpercentcouponsum.value*1;
			}
		}
        frm.remainpercentcouponsum.value = frm.orgpercentcouponsum.value*1 - frm.refundpercentcouponsum.value*1;
		frm.remaincouponsum.value = frm.remainpercentcouponsum.value;
    }

	// ========================================================================
    // 취소안된 기타할인 전체합계 금액
    if (frm.orgallatsubtractsum!=undefined){
		// aaaaaaaaaaaaaaaaa frm.orgallatsubtractsum.value = orgetcDiscountDiscountSum * -1;
    }

    // 취소/반품 기타할인 총액
    if (frm.refundallatsubtractsum!=undefined){
        frm.refundallatsubtractsum.value = refundetcDiscountDiscountSum * -1;
    }

    // 선택안한(나머지) 기타할인 총액
    if (frm.remainallatsubtractsum!=undefined){
        frm.remainallatsubtractsum.value = frm.orgallatsubtractsum.value*1 - frm.refundallatsubtractsum.value*1;
    }

    // 과거 변수명도 설정해준다.(pop_cs_action_new_process.asp 에 값을 넘기기 위해 필요)
    if (frm.allatdiscountprice!=undefined){
		// aaaaaaaaaaaaaaaaa frm.allatdiscountprice.value = orgetcDiscountDiscountSum;
    }

    if (frm.allatsubtractsum!=undefined){
		frm.allatsubtractsum.value = refundetcDiscountDiscountSum * -1;
    }

    if (frm.remainallatdiscount!=undefined){
		frm.remainallatdiscount.value = frm.orgallatsubtractsum.value*1 - frm.refundallatsubtractsum.value*1;
    }
}


// 배송비 합계 계산
function CalculateBeasongPaySum(frm) {
	var orgbeasongpay = 0;
	var refundbeasongpay = 0;
	var remainbeasongpay = 0;

	var objdeliver;
	var objitemid;
	var objitemcost;
	var objitemno;

	if (!frm.orderdetailidx) {
		return;
	}

	if (!frm.refundbeasongpay) {
		return;
	}

	for (var i = 0; i < frm.orderdetailidx.length; i++) {
		objdeliver = frm.orderdetailidx[i];
		objitemid = frm.itemid[i];
		objitemcost = frm.itemcost[i];
		objitemno = frm.regitemno[i];

        if (objdeliver.type != "checkbox") {
        	continue;
        }

        if (objdeliver.checked != true) {
        	continue;
        }

        if (objitemid.value*1 != 0) {
        	continue;
        }

		refundbeasongpay = refundbeasongpay + ((objitemcost.value*1) * (objitemno.value*1));
	}

    frm.refundbeasongpay.value = refundbeasongpay;
    frm.remainbeasongpay.value = frm.orgbeasongpay.value*1 - refundbeasongpay;
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

	if ((IsStatusFinished != true) && (IsStatusFinishing != true)) {
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

	if ((IsStatusFinished != true) && (IsStatusFinishing != true)) {
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
var IsCkReturnPayChanged = true;
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


	if ((IsTravelOrder == true) && (travelItemExist == true) && ((IsStatusRegister == true) || (IsStatusEdit == true))) {
		if (IsTravelItemSelected(frm) == true) {
			refunddeliverypay = GetTravelItemDeliverPaySUM(frm) * -1;
		}
	}

    if (frm.refunddeliverypay!=undefined){
        frm.refunddeliverypay.value = refunddeliverypay*1;
		if (IsCkReturnPayChanged == true) {
			frm.addbeasongpay.value = frm.refunddeliverypay.value;
			IsCkReturnPayChanged = false;
		}
    }

	// 업체 추가정산
    CalculateUpcheReturnBeasongPay(frm);
}

// 업체 추가정산
function CalculateUpcheReturnBeasongPay(frm) {
	return;

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

	refunddeliverypay = frm.refunddeliverypay.value*1;

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
				if (mayTicketCancelPro > 100) {
					// 금액(한도 10%)
					if ((refundtotalbuypaysum*10/100) > mayTicketCancelPro) {
						alert( ticketCancelStr + '티켓 취소 수수료 ' + mayTicketCancelPro + '원 차감 \n\n(단, 당일 주문건 취소시는 제외)' );
						frm.refundadjustpay.value = mayTicketCancelPro*-1;
					} else {
						alert( ticketCancelStr + '티켓 취소 수수료 10% 차감 \n\n(단, 당일 주문건 취소시는 제외)' );
						frm.refundadjustpay.value = (refundtotalbuypaysum*10/100)*-1;
					}
				} else {
					// 퍼센트
					alert( ticketCancelStr + '티켓 취소 수수료 ' + mayTicketCancelPro + '% 차감 \n\n(단, 당일 주문건 취소시는 제외)' );
					frm.refundadjustpay.value = (refundtotalbuypaysum*mayTicketCancelPro/100)*-1;
				}
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
	remainsubtotalprice = orgsubtotalprice - refundsubtotalprice;

	frm.orgsubtotalprice.value = orgsubtotalprice;
	frm.refundsubtotalprice.value = refundsubtotalprice;
	frm.remainsubtotalprice.value = remainsubtotalprice;

	// 과거 변수명도 설정해준다.(pop_cs_action_new_process.asp 에 값을 넘기기 위해 필요)
	frm.subtotalprice.value = orgsubtotalprice;
	frm.canceltotal.value = refundsubtotalprice;
	frm.nextsubtotal.value = remainsubtotalprice;

	// ========================================================================
	// 환불금액 입력
	if ((parseInt(frm.ipkumdiv.value) >= 4) || IsChangeOrder) {
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
/*
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
 */

// ============================================================================
// 여행상품 선택되었는지
// ============================================================================
function IsTravelItemSelected(frm) {
	var orderdetailidx;

	travelCancelDisabled		= false;
	travelCancelStr             = '';

    if (!frm.orderdetailidx) return false;

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
		orderdetailidx = frm.orderdetailidx[i];
		if (orderdetailidx.type == "checkbox") {
			if (orderdetailidx.checked == true) {
				for (var j = 0; j < travelItemInfoArr.length; j++) {
					if (travelItemInfoArr[j][0]*1 == orderdetailidx.value*1) {
						if (travelItemInfoArr[j][1] == "N") {
							travelCancelDisabled = true;
							travelCancelStr = travelItemInfoArr[j][3];
						}
						return true;
					}
				}
			}
		}
    }

	return false;
}

function GetTravelItemDeliverPaySUM(frm) {
	var orderdetailidx;
	var paysum = 0;

    if (!frm.orderdetailidx) return paysum;

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
		orderdetailidx = frm.orderdetailidx[i];
		if (orderdetailidx.type == "checkbox") {
			if (orderdetailidx.checked == true) {
				for (var j = 0; j < travelItemInfoArr.length; j++) {
					if (travelItemInfoArr[j][0]*1 == orderdetailidx.value*1) {
						paysum = paysum + travelItemInfoArr[j][2]*frm.regitemno[i].value;
					}
				}
			}
		}
    }

	return paysum;
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

    if (!frm.orderdetailidx) return;

	if (IsCSReturnProcess != true) {
		return 0;
	}

	checkfound = IsCheckedItemExist(frm);
	upbeaMakerid = GetCheckedItemMaker(frm);

    if (checkfound == false) {
    	return 0;
    }

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if ((e.type == "checkbox") && (e.checked == false) && (frm.itemid[i].value*1 == 0) && (upbeaMakerid == frm.makerid[i].value)) {
            return frm.itemcost[i].value*1;
        }
    }

    return 0;
}

// ============================================================================
// 선택않된 업체배송비 합계(브랜드별)
// ============================================================================
function GetNotCheckedUpcheBeasongPayByBrand(frm, makerid) {

    if (!frm.orderdetailidx) return 0;

	if (IsCSReturnProcess != true) {
		return 0;
	}

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];

        if (e.type != "checkbox") {
        	continue;
        }

		if (frm.itemid[i].value*1 != 0) {
			continue;
		}

        if ((e.checked == false) && (makerid == frm.makerid[i].value)) {
            return frm.itemcost[i].value*1;
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
        document.all.refundinfo_R007.style.display = "";
    }else if((comp.value=="R020")||(comp.value=="R022")||(comp.value=="R080")||(comp.value=="R100")||(comp.value=="R120")||(comp.value=="R400")||(comp.value=="R420")){
        //실시간 이체 취소//실시간 이체 부분 취소//ALL@ 결제 취소 //신용카드 결제 취소 //신용카드 부분취소//휴대폰 결제 취소 //휴대폰 부분취소
        document.all.refundinfo_R100.style.display = "";
    }else if(comp.value=="R550"){
        //기프팅 결제 취소
        document.all.refundinfo_R550.style.display = "";
    }else if(comp.value=="R560"){
        //기프티콘 결제 취소
        document.all.refundinfo_R560.style.display = "";
    }else if(comp.value=="R050"){
        //입점몰 결제 취소
        document.all.refundinfo_R050.style.display = "";
    }else if ((comp.value=="R900") || (comp.value=="R910")) {
        //마일리지 환불, 예치금 환불
        document.all.refundinfo_R900.style.display = "";
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

	// 반품접수(업체배송), 맞교환출고, 상품변경 맞교환출고, 업체기타정산, 누락재발송, 서비스발송, 기타회수
	if ((IsUpcheReturnState(frm) == true) || (divcd == "A000") || (divcd == "A100") || (divcd == "A700") || (divcd == "A001") || (divcd == "A002") || (divcd == "A200")) {
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
function Change_add_upchejungsancause(comp) {
	var jungsancause = document.all.add_upchejungsancause;

	if (divcd != "A700") {
		// 업체기타정산 아니면 배송비만 선택가능
		if ((jungsancause.value != "배송비") && (jungsancause.value != "도선료")) {
			alert("\n\n업체기타정산이 아니면 [배송비,도선료] 만 선택가능합니다.\n\n상품대금 등 다른 정산내역은 업체기타정산으로 등록하세요.\n[사원이상사용가능]\n\n");
			jungsancause.value = "";
		}
	} else {
		if ((jungsancause.value != "배송비") && (jungsancause.value != "도선료")) {
			//if ((IsCsPowerUser != true) && (HasAuthUpcheJungsanItemPrice != true)) {
			if ( (C_CSpermanentUser != true) && (IsC_ADMIN_AUTH != true) ) {
				alert("\n\n[배송비,도선료] 이외의 사유는 사원이상사용가능\n\n");
				jungsancause.value = "";
			} else {
				if (jungsancause.value == "상품대금") {
					alert("\n\n상품대금(상품별 매입가)은 추가정산금액이 0원으로 입력됩니다.\n(물류입고내역 작성 후 정산됩니다.)\n\n접수내용에 상품별 매입가를 정확히 기재하세요.");
					document.all.add_upchejungsandeliverypay.value = "0";
					document.all.buf_totupchejungsandeliverypay.value = "0";
				//} else if (jungsancause.value = "직접입력") {
				//	jungsancause.value = "";
				//	alert("접수불가!!!");
				}
			}
		}
	}

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
	var jungsancause = document.all.add_upchejungsancause;

	if ((jungsancause.value == "상품대금") && (comp.value*1 != 0)) {
		comp.value = "0";
		alert("\n\n상품대금은 추가정산금액이 0원으로 입력됩니다.\n(물류입고내역 작성 후 정산됩니다.)\n\n접수내용에 상품별 매입가를 정확히 기재하세요.");
	}

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
/*
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
*/

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

    if ((IsTravelOrder==true)&&(travelCancelDisabled==true)){
        if (!confirm('취소 기한이 지나 취소 불가합니다. ' + travelCancelStr + ' \n\n 계속하시겠습니까?')){
            return;
        }

        //권한 있는지 check
        if (IsCsPowerUser!=true){
            alert('권한이 없습니다. 파트장급 문의 요망');
            return;
        }
    }

    if (((IsGiftingOrder == true) || (IsGiftiConOrder == true)) && (IsOrderCancelDisabled == true) && IsCSReturnProcess) {
        if (!confirm('반품불가 주문입니다. : ' + OrderCancelDisableStr + ' \n\n 계속하시겠습니까? [관리자권한 필요]')) {
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

	if (divcd == "A000") {
		// 동일상품 교환출고의 경우, 사이즈교환 선택 못함
		if ((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD04")) {
	        alert('사이즈교환일 경우 옵션변경 교환출고를 등록하세요');
	        return;
		}
	}

	// 취소, 반품
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true)) {
		// 브랜드 전체취소 또는 주문취소이면서 배송비만 남는 경우 체크
        if (CheckBeasongPayOnlyRemain(frm) != true) {
            return;
        }
	}

	// 선택 상품 체크 및 저장
	if ((divcd != "A100") && (divcd != "A111") && (divcd != "A112")) {
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

	// 반품배송비 체크
	if (IsCSReturnProcess == true) {
		if (CheckRefundDeliverYpay(frm) != true) {
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

	// 2018-01-19, skyer9, 쿠폰 재발급 안내창 띄우기
	if (IsCSReturnProcess) {
		if ((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD01")) {
			// 변심반품 제외
		} else {
			if (frm.tmpcopycouponinfo) {
				if (frm.tmpcopycouponinfo.checked == false) {
					if (confirm("\n변심반품이 아닌 경우 쿠폰을 재발급해 드릴 수 있습니다.\n(한장만 재발급 가능합니다.)\n\n쿠폰 재발급하시겠습니까?\n\n\n") == true) {
						frm.tmpcopycouponinfo.checked = true;
						jsCheckCopyCoupon(frm);
					}
				}
			}
		}
	}

	if ($("#needChkYN_X").val()) {
		if ($("#needChkYN_X").prop("checked") === false && $("#needChkYN_F").prop("checked") === false) {
			alert('완료구분을 선택하세요.');
			$("#needChkYN_X").focus();
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

			if ((divcd == "A700") && (frm.add_upchejungsancause.value != '상품대금')) {
	            alert('추가 정산액을 입력하세요.');
	            frm.add_upchejungsandeliverypay.focus();
	            return;
			}
        }
    }

	if ((IsCSCancelProcess == true) && IsStatusRegister) {
        if (frm.modifyitemstockoutyn) {
        	if (frm.modifyitemstockoutyn.checked == true) {
		        if (IsMultiStockOutItemChecked(frm) == true) {
		            // 두개 이상의 상품이 품절설정 되어 있는 경우
		            if (confirm("\n\n\n============ 여러 상품을 [품절] 등록합니다. =================\n\n진행하시겠습니까?") != true) {
		            	return;
		            }
		        }
        	}
        }
	}

    if (IsCSCancelProcess){
        if(confirm("취소 접수 하시겠습니까?")) {
			DisableAllButton();
            frm.submit();
        }
    }else if (IsCSReturnProcess){
        if (frm.requireupche.value=="Y"){
            if(confirm("업체 [" + frm.requiremakerid.value +"]로 반품/회수/교환 접수 하시겠습니까?")){
                ChangeCSTitleGubun(frm);
				DisableAllButton();
                frm.submit();
            }
        }else{
            if(confirm("[텐바이텐 물류센터]로 반품/회수/교환 접수 하시겠습니까?")){
                ChangeCSTitleGubun(frm);
				DisableAllButton();
                frm.submit();
            }
        }
    }else if (IsCSRefundNeeded) {
        if(confirm("환불 접수 하시겠습니까?")){
			DisableAllButton();
            frm.submit();
        }
    }else if(confirm("접수 하시겠습니까?")){
		DisableAllButton();
        frm.submit();
    }
}

function DisableAllButton() {
	var inputs = document.getElementsByTagName("INPUT");
	for (var i = 0; i < inputs.length; i++) {
		if ((inputs[i].type === 'button') && (inputs[i].name === '')) {
			inputs[i].disabled = true;
		}
	}
}

// 2014-02-06 skyer9
function CheckRefundDeliverYpay(frm) {
	if (frm.refunddeliverypay) {
		if (frm.refunddeliverypay.value == "") {
			alert("반품배송비를 입력하세요.");
			frm.refunddeliverypay.focus();
			return false;
		}

		if (frm.refunddeliverypay.value*0 != 0) {
			alert("반품배송비는 숫자만 가능합니다.");
			frm.refunddeliverypay.focus();
			return false;
		}

		if (frm.refunddeliverypay.value*1 > 0) {
			alert("반품배송비는 마이너스 금액만 입력가능합니다.");
			frm.refunddeliverypay.focus();
			return false;
		}

		if (frm.addbeasongpay) {
			if ((frm.addmethod[1].checked != true) && (frm.addmethod[2].checked != true) && (frm.addmethod[3].checked != true)) {
				frm.addmethod[0].checked = true;
			}

			if (frm.addbeasongpay.value*1 != 0) {
				if ((frm.refunddeliverypay.value*1 != frm.addbeasongpay.value*1) && (frm.addmethod[0].checked == true)) {
					alert("배송비 추가결제 방식을 선택하세요.");
					frm.addmethod[0].focus();
					return false;
				} if (frm.refunddeliverypay.value*1 == frm.addbeasongpay.value*1) {
					frm.addmethod[0].checked = true;
				}
			}
		}
	}

	return true;
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
			// 어드민도 불가!!, 2017-05-04, skyer9
	        // if (IsAdminLogin != true) {
	        	return false;
	        // }
	    }
	}

	if (frm.mainpaymentorg && frm.prevrefundsum && frm.refundrequire) {
		if (frm.mainpaymentorg.value*1 > 0) {
			// 입점몰 제외
			if (frm.mainpaymentorg.value*1 < (frm.prevrefundsum.value*1 + frm.refundrequire.value*1)) {
				alert('결제금액 이상으로 환불할 수 없습니다. - 쿠폰이나 마일리지 환급을 체크해 주세요.');
				return false;
			}
		}
	}

	if (frm.remainsubtotalprice) {
	    if (frm.remainsubtotalprice.value*1 < 0) {
	        alert('취소 후 결제 금액이 마이너스가 될 수 없습니다. - 쿠폰이나 마일리지 환급을 체크해 주세요.');
			return false;

			// 관리자도 불가, skyer9, 2018-01-04
			/*
	        if (IsAdminLogin != true) {
	        	return false;
	        }
			*/
	    }
	}

	if (frm.orgpercentcouponsum) {
		if ((frm.orgpercentcouponsum.value*1 != 0) && (frm.orgfixedcouponsum.value*1 != 0)) {
			alert('\n\n쿠폰 금액 오류 : 시스템팀 문의[쿠폰이중]!!!\n\n');
			return false;
		}

		if ((frm.orgpercentcouponsum.value*1 > 0) || (frm.orgallatsubtractsum.value*1 > 0) || (frm.orgfixedcouponsum.value*1 > 0) || (frm.orgcouponsum.value*1 > 0) || (frm.orgmileagesum.value*1 > 0) || (frm.orggiftcardsum.value*1 > 0) || (frm.orgdepositsum.value*1 > 0)) {
			alert('\n\n접수 금액 오류 : 시스템팀 문의[금액플러스]!!!\n\n');
			return false;
		}

		if ((frm.refundpercentcouponsum.value*1 > 0) || (frm.refundallatsubtractsum.value*1 > 0) || (frm.refundfixedcouponsum.value*1 > 0) || (frm.refundcouponsum.value*1 > 0) || (frm.refundmileagesum.value*1 > 0) || (frm.refundgiftcardsum.value*1 > 0) || (frm.refunddepositsum.value*1 > 0)) {
			alert('\n\환불 금액 오류 : 시스템팀 문의!!!\n\n');
			return false;
		}

		if ((frm.orgpercentcouponsum.value*1 == 0) && (frm.refundpercentcouponsum.value*1 != 0)) {
			alert('\n\n접수/환불 금액 오류 : 시스템팀 문의[비율쿠폰]!!!\n\n');
			return false;
		}

		if ((frm.orgallatsubtractsum.value*1 == 0) && (frm.refundallatsubtractsum.value*1 != 0)) {
			alert('\n\n접수/환불 금액 오류 : 시스템팀 문의[기타할인]!!!\n\n');
			return false;
		}

		if ((frm.orgfixedcouponsum.value*1 == 0) && (frm.refundfixedcouponsum.value*1 != 0)) {
			alert('\n\n접수/환불 금액 오류 : 시스템팀 문의[정액쿠폰]!!!\n\n');
			return false;
		}

		if ((frm.orgcouponsum.value*1 == 0) && (frm.refundcouponsum.value*1 != 0)) {
			alert('\n\n접수/환불 금액 오류 : 시스템팀 문의[쿠폰]!!!\n\n');
			return false;
		}

		if ((frm.orgmileagesum.value*1 == 0) && (frm.refundmileagesum.value*1 != 0)) {
			alert('\n\n접수/환불 금액 오류 : 시스템팀 문의[마일리지]!!!\n\n');
			return false;
		}

		if ((frm.orggiftcardsum.value*1 == 0) && (frm.refundgiftcardsum.value*1 != 0)) {
			alert('\n\n접수/환불 금액 오류 : 시스템팀 문의[기프트]!!!\n\n');
			return false;
		}

		if ((frm.orgdepositsum.value*1 == 0) && (frm.refunddepositsum.value*1 != 0)) {
			alert('\n\n접수/환불 금액 오류 : 시스템팀 문의[예치금]!!!\n\n');
			return false;
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

	/*
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
	*/

    //기타내역, 서비스발송 , 기타회수, 환불요청, 출고시유의사항, 업체 추가 정산 - 상세내역 체크 안함.
    if ((divcd=="A009") || (divcd=="A002") || (divcd=="A200") || (divcd=="A003") || (divcd=="A005") || (divcd=="A006") || (divcd=="A007") || (divcd=="A700") || (divcd=="A100") || (divcd=="A111")) {
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
		/*
		// 환불없음 이면 체크 안한다.
		frm.refundrequire.value = "0";
		return true;
		 */

		// 환불금액 있는 경우 환불없음 선택못하게 수정(2015-01-05, skyer9)
		if (frm.refundsubtotalprice.value*1 > 0) {
			alert('환불금액이 있는 경우, 환불없음 선택할 수 없습니다.');
			frm.returnmethod.focus();
			return false;
		}
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
    if ((PayedNCancelEqual == true) && ((frm.returnmethod.value=="R120") || (frm.returnmethod.value=="R022") || (frm.returnmethod.value=="R420"))){
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

    //부분취소 관련 ALERT
    if (frm.returnmethod.value=="R420") {
        if (phonePartialCancelok!="Y") {
            alert('부분 취소 가능 결제가 아닙니다.');
            return false;
        }
    }

    //부분취소 관련 ALERT (네이버페이 실시간 이체)
    if (frm.returnmethod.value=="R022") {
        if (pggubun=="NP" || pggubun=="KK"){
            return true;
		}
		else{
			alert('네이버페이만 실시간이체부분취소가 가능합니다.');
            return false;
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
        return false;
	}

	if ((frm.refundrequire.value*1 <= 0) && (frm.returnmethod.value != "R000")) {
		if (IsCsPowerUser != true) {
			alert('환불 금액은 0 보다 커야합니다. 또는 환불없음을 선택하세요.[0]');
			return false;
		} else if (frm.refundrequire.value*1 != 0) {
			alert('환불 금액은 0 보다 커야합니다. 또는 환불없음을 선택하세요.[1]');
			return false;
		}
	}

    if (((IsGiftingOrder == true) || (IsGiftiConOrder == true)) && (IsOrderCancelDisabled == true) && IsCSCancelProcess) {
        if (frm.returnmethod.value != "R910") {
        	// 예치금 환불의 경우 관리자 권한 필요없음
	        if (!confirm('취소불가 주문입니다. : ' + OrderCancelDisableStr + ' \n\n 계속하시겠습니까? [관리자권한 필요]')) {
	            return false;
	        }

	        //권한 있는지 check
	        if (IsCsPowerUser != true) {
	            alert('권한이 없습니다. 파트장급 문의 요망');
	            return false;
	        }
        }
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

    //네이버 페이는 원결제수단으로 환불 해야.. 얼럿만. 2016/08/05
    if (pggubun=="NP"){
        if ((orgaccountdiv=="100")&&((frm.returnmethod.value!="R100")&&(frm.returnmethod.value!="R120"))){
            if(!confirm('네이버 페이인경우 원결제수단(카드)으로 환불하는것이 기본입니다.\r\n선택한 결제수단으로 환불 처리 계속 하시겠습니까?'))return;
        }

        if ((orgaccountdiv=="20")&&((frm.returnmethod.value!="R020")&&(frm.returnmethod.value!="R022"))){
            if(!confirm('네이버 페이인경우 원결제수단(실시간이체)으로 환불하는것이 기본입니다.\r\n선택한 결제수단으로 환불 처리 계속 하시겠습니까?'))return;
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

function CsRegCancelFinishedProc(frm) {
	if (confirm('\n\n완료CS내역(취소,반품)을 삭제하시겠습니까?\n\n')) {
        frm.mode.value = "delfinishedcsas";
        frm.submit();
    }
}

// ============================================================================
// 완료처리
// ============================================================================
function CsRegFinishProc(frm) {
	var btn = document.getElementById("btnFinishReturn");

	// 마스터 체크
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	// 선택 상품 체크 및 저장
	if ((divcd != "A100") && (divcd != "A111") && (divcd != "A112")) {
	    if (!SaveCheckedItemList(frm)) {
	        return;
	    }
	} else {
	    if (!SaveChangeCheckedItemList(frm)) {
	        return;
	    }
	}

	// 취소, 반품, 환불
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true) || (divcd == "A003") || ((divcd == "A100") && frm.refundrequire && frm.refundrequire.value*1 !== 0)) {
        if (CheckReturnForm(frm) != true) {
            return;
        }

		if (frm.refundrequire) {
			if (frm.refundrequire_org.value*1 != frm.refundrequire.value*1) {
				alert("[시스템팀 문의] 환불 예정액 오류!!!!");
				return;
			}
		}
	}

	// 취소, 반품
    if ((IsCSCancelProcess) || (IsCSReturnProcess) || ((divcd == "A100") && frm.refundrequire && frm.refundrequire.value*1 !== 0)) {
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

    if (IsChangeOrder && (OrderMasterState != "8")) {
        if (!confirm('교환회수완료 이전 반품입니다.\n\n 계속하시겠습니까? [관리자권한 필요]')) {
            return;
        }

        //권한 있는지 check
        if (IsCsPowerUser != true) {
            alert('권한이 없습니다. 파트장급 문의 요망');
            return;
        }
    }

    var confirmMsg ;
    confirmMsg = '완료처리 진행 하시겠습니까?';

    if ((divcd == "A004") || (divcd == "A010")) {
        confirmMsg = '완료처리 진행시 마이너스 주문 및 환불이 자동 접수됩니다. 진행 하시겠습니까?';
    }

	if (btn) {
		btn.disabled = true;
	}

    if (confirm(confirmMsg )) {
        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "";
        frm.submit();
    }

	if (btn) {
		btn.disabled = false;
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
// 교환주문 수기생성
// ============================================================================
function CsChangeOrderRegProc(frm) {
    if (confirm('교환주문을 생성 하시겠습니까?')){
        frm.mode.value = "changeorderreg";
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
	if ((divcd != "A100") && (divcd != "A111") && (divcd != "A112")) {
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

	// 반품배송비 체크
	if (IsCSReturnProcess == true) {
		if (CheckRefundDeliverYpay(frm) != true) {
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
        if (frm.add_upchejungsandeliverypay.value*0 != 0) {
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

	if ($("#needChkYN_X").val()) {
		if ($("#needChkYN_X").prop("checked") === false && $("#needChkYN_F").prop("checked") === false) {
			alert('완료구분을 선택하세요.');
			$("#needChkYN_X").focus();
			return;
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

// ============================================================================
// 접수창 사이즈 증가
// ============================================================================
function resizeTextArea(textarea, textareawidth) {
	if (!textarea) { return; }
	var lines = textarea.value.split('\n');
	var textareaminheight = textarea.rows;

	var textareaheight = 1;
	for (x = 0; x < lines.length; x++) {
		// textareawidth 넘어가면 한줄 추가
		c = lines[x].length;
		if (c >= textareawidth) {
			textareaheight += Math.ceil(c / textareawidth);
		}
	}
	textareaheight += lines.length;

	for (x = (lines.length - 1); x >= 0; x--) {
		// 맨밑의 엔터라인은 무시
		c = lines[x].length;
		if (c == 0) {
			textareaheight = textareaheight - 1;
		} else {
			break;
		}
	}

	if (textareaheight < textareaminheight) {
		textareaheight = textareaminheight;
	} else {
		textareaheight += 1;
	}

	textarea.rows = textareaheight;
}

// 배송비취소
function CsRegCancelBeasongPayProc(frm, orderdetailidx) {
	if (CheckedProductExist(frm, orderdetailidx) == true) {
		alert("배송비취소 불가\n\n체크된 상품이 있는 경우 배송비를 선택취소할 수 없습니다.");
		return;
	}

	var e;
	var ischecked = false;
	for (var i = 0; i < frm.length; i++) {
		e = frm.elements[i];
		if (e.name == "orderdetailidx") {
			if (e.value*1 == orderdetailidx) {
				e.checked = true;
				AnCheckClick(e);
				CheckSelect(e, true);

				break;
			}
		}
	}

	CsRegProc(frm);

	try {
		e.checked = false;
		AnCheckClick(e);
		CheckSelect(e, true);
	} catch (err) { }
}

function CheckedProductExist(frm, orderdetailidx) {
    var e;
    var ischecked = false;

	for (var i = 0; i < frm.length; i++) {
		e = frm.elements[i];
		if ((e.name == "orderdetailidx") && (e.value*1 != orderdetailidx)) {
			if (e.checked == true) {
				return true;
			}
		}
	}

	return false;
}
