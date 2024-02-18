
/* global IsPossibleModifyCSMaster, IsPossibleModifyItemList, ERROR_MSG_TRY_MODIFY, initializeURL, initializeReturnFunction, initializeErrorFunction, startRequest, $ */

// ============================================================================
// ������ �������� ����
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


    // ��ü ���������� �Է��� ��� ���õ� Detail�� ���� ���� ����
    if (targetDiv=="causepop") {
    	// �ֹ� ��ǰ
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

		// ��ǰ���� �±�ȯ��� ��ǰ
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

	// ��ǰ����� üũ
	CheckForItemChanged();

	// ��ǰ���� �±�ȯ�� ��ۺ�
	CheckBeasongPayCutItemChange(document.frmaction);
}

// ============================================================================
// ��ǰ �������� ����
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
// ��ǰ �������� ����
// ============================================================================
function delGubun(name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv) {
    eval("document." + name_frm + "." + name_gubun01).value = "";
    eval("document." + name_frm + "." + name_gubun02).value = "";
    eval("document." + name_frm + "." + name_gubun01name).value = "";
    eval("document." + name_frm + "." + name_gubun02name).value = "";

    eval(targetDiv).innerHTML = "";
}

// ============================================================================
// CS �������� ǥ�� (AJAX)
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
    alert("�����͸� �д� ���߿� ������ �߻��߽��ϴ�. ����� �ٽ� �õ��غ��ñ� �ٶ��ϴ�.[CODE:" + xmlHttp.status + "]");
}

function colseCausepop(targetDiv){
    eval(targetDiv).innerHTML = "";
}

// ============================================================================
// ��ǰ����� üũ
// ============================================================================
function CheckForItemChanged() {
	var frm = document.frmaction;

	// ������ǰ �ݾ� ����
    CalculateAndApplyItemCostSum(frm);

	if (IsStatusRegister == true) {
		if (IsOnlyOneBrandAvailable == true) {
		    // ���� �귣�常 ���� �����ϰ�, ��� �귣�� ����

	    	EnableOnlyOneBrand(frm);

			// ���� ��ǰ/�±�ȯ : ��ü �߰����� ���� �귣�� ���̵� ��������
			InsertCheckedUpcheID(frm);
		}
	}

	// ========================================================================
	// ��ǰ�� üũ
	// ========================================================================
	if (IsCSReturnProcess == true) {
		// ���ֹ����� �̻��� ��ǰ�� �ִ��� üũ
		CheckOverReturnItemno(frm);

		if (IsStatusRegister == true) {
			if (IsOnlyOneBrandAvailable == true) {
			    // ���� �귣�常 ���� �����ϰ�, ��� �귣�� ����

				// ���� ��ǰ/�±�ȯ : ��ü �߰����� ���� �귣�� ���̵� ��������
				// InsertCheckedUpcheID(frm);
			}
		}

		// ��ۺ� ����(�ܼ�����)
		CheckBeasongPayCut(frm);

		// ��ǰ�� ȸ�� ��ۺ� ���
		CalculateReturnBeasongPay(frm);
	}

	if (IsStatusRegister == true) {
		if (IsOnlyOneBrandAvailable == true) {
		    // ���� �귣�常 ���� �����ϰ�, ��� �귣�� ����

		    if (IsStatusRegister == true) {
		    	// EnableOnlyOneBrand(frm);
		    }

			// ���� ��ǰ/�±�ȯ : ��ü �߰����� ���� �귣�� ���̵� ��������
			// InsertCheckedUpcheID(frm);
		}

		// üũ�� ��ǰ/��ۺ� ���ٲٱ�
		// AnCheckClickAll(frmaction);

		// �ֹ���ü ���ý� ���ϸ���, ���α� �ڵ�üũ
		CheckMileageETC(frm);

		// ���� ��߱� üũ
		jsCheckCopyCoupon(frm);
	}

	// ������ǰ �ݾ� ����
    CalculateAndApplyItemCostSum(frm);

	/*
	TODO : ��ǰ��ҷ� ���� ��ۺ� �߰�
    if (IsAddBrandBeasongPayNeed(frm, "") == true) {
    	alert("�ٹ����� ��� ��ǰ�� ���»�ǰ�� �ݾ��� 30000�� �̸��̹Ƿ� ��ۺ� 2000���� �߰��˴ϴ�.");
    }
    */
}

// ============================================================================
// ���ֹ����� �̻��� ��ǰ�� �ִ��� üũ(���� ��ǰ, �±�ȯȸ�� CS�Ϸ᳻��)
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
			alert("�ֹ������� �ʰ��Ͽ� �����Ǵ� ��ǰ�� �ֽ��ϴ�.");

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
// ������ ��ǰ �Է¼��� üũ
// ============================================================================
function CheckMaxItemNo(obj, maxno) {
	var frm = document.frmaction;

	if (obj.value*0 != 0) {
		return;
	}

    if (obj.value*1 > maxno*1) {
        alert("�ֹ����� �̻����� ��ǰ������ �����Ҽ� �����ϴ�.");
        obj.value = maxno;
    }

	// ========================================================================
	// ��ǰ���,��ǰ�� ��ۺ� �������(�Ǵ� ���̳ʽ� �ֹ����)
	// ========================================================================
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true)) {
	    CheckUpcheDeliverPay(frm);

		// ��ǰ ��ü üũ �� ��� üũ�ȵ� ��ۺ� ����üũ
		if (IsCSCancelProcess == true) {
			CheckBeasongPayIfAllItemSelected(frm);
		}
	}

	// ��ǰ����� üũ
	CheckForItemChanged();
}

function CheckBeasongPayOnlyRemain(frm) {
    var upbeaMakerid;
	var cancelyn;

    var objdeliver, objitem;

    if (!frm.orderdetailidx) return true;
    if ((IsCSCancelProcess != true) && (IsCSReturnProcess != true)) return true;

	// �귣�� ���� ��ۺ� üũ
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
					alert("�귣�� ��ü���(��ǰ)�̸鼭 ��ۺ� ���õ��� �ʾҽ��ϴ�.");
					return false;
				}
			}
			*/
		}
	}

	// ��ü���
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
						alert("�ֹ� ��ü����̸鼭 ���þʵ� ��ۺ� �ֽ��ϴ�.");
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
// ��ǰ���ý� Ȯ���� �͵�
// ============================================================================
function CheckSelect(comp, isbeasongpay) {
    var chkidx = comp.value;
    var frm = document.frmaction;
	var objdeliver;

	if (isbeasongpay == false) {
		// ========================================================================
		// ��ǰ���,��ǰ�� ��ۺ� �������(�Ǵ� ���̳ʽ� �ֹ����)
		// ========================================================================
		if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true)) {
			CheckUpcheDeliverPay(frm);
		}

		// ��ǰ ��ü üũ �� ��� üũ�ȵ� ��ۺ� ����üũ
		if (IsCSCancelProcess == true) {
			CheckBeasongPayIfAllItemSelected(frm);
		}
	}

	// ��ۺ� ����� ��� ����ǰ�� �ִ��� üũ
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
						if (confirm("[������ ����]\n\n����ǰ�� �ֽ��ϴ�.\n��ۺ� ����Ͻðڽ��ϱ�??") != true) {
							objdeliver.checked = false;
							AnCheckClick(frm.orderdetailidx[i]);
							return;
						}
					} else {
						alert("����ǰ�� �ִ� ��� ��ۺ� ����� �� �����ϴ�.");
						objdeliver.checked = false;
						AnCheckClick(frm.orderdetailidx[i]);
						return;
					}
				}
			}
		}
	}

	if (comp.checked){
		// CS �������� ����
		eval("frm.gubun01_" + chkidx).value = frm.gubun01.value;
		eval("frm.gubun02_" + chkidx).value = frm.gubun02.value;
		eval("frm.gubun01name_" + chkidx).value = frm.gubun01name.value;
		eval("frm.gubun02name_" + chkidx).value = frm.gubun02name.value;
	}else{
		delGubun("gubun01_" + chkidx,"gubun02_" + chkidx,"gubun01name_" + chkidx,"gubun02name_" + chkidx,frm.name,causepop);
	}

	// ��ǰ����� üũ
	CheckForItemChanged(comp);
}

// ============================================================================
// ��ۺ� ���� üũ
// ============================================================================
//
// �̹� ��ۺ� ȯ���� üũ�Ǿ� ������ ���� ���Ѵ�.
//
// ���,��ǰ��
// �ٹ�   - �ٹ� �귣�� ��ü�� ���õ� ���, ��ۺ� üũ�Ѵ�.
// ����   - ���� �귣�� ��ü�� ���õ� ���, ��ۺ� üũ�Ѵ�.(���ɹ�ǰ�� ��ۺ��������� ó���Ѵ�.)
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
				// CS �������� ����
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

// ��ۿϷ� ��ǰ �ִ���(�ٹ�, ���� �귣�庰)
// ����ǰ �ִ� ��� ��ۺ� ��� �Ұ�
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
// ���� �귣�常 ���� �����ϰ�, ��� �귣�� ����
//
// �ٹ� üũ�� - ���� Disable
// ���� üũ�� - �ٹ� �� �ٸ� �귣�� Disable
// �����ǰ �ٹ�ȸ���� ��� Disable ����(�����귣�� �������� ����)
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

	// �����ÿ��� ����Ѵ�.
	if (IsStatusRegister != true) {
		return;
	}

	// ========================================================================
	// üũ�� ��ǰ �˻�
	checkfound = IsCheckedItemExist(frm);
	upbeaMakerid = GetCheckedItemMaker(frm);

	// ========================================================================
	// ��� �귣�� ����
	if (checkfound != true) {
        frm.requireupche.value = "";
        frm.requiremakerid.value = "";
	} else {
		if (forcereturnbycustomer == true) {
			// ���ٹ��� ����ǰ
			frm.requireupche.value = "Y";
			frm.requiremakerid.value = "10x10logistics";
		} else if ((upbeaMakerid.length < 1) || (forcereturnbyten == true)) {
			// ����ȸ��
	        frm.requireupche.value = "N";
	        frm.requiremakerid.value = "";
		} else {
			// ��ü��ǰ
			frm.requireupche.value = "Y";
			frm.requiremakerid.value = upbeaMakerid;
		}
	}

	// ========================================================================
	// �����ǰ �ٹ�ȸ���� ��� Disable ����
	// �ٹ����� ������ ��������ǰ�� ��� Disable ����
	// ========================================================================
    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        objitem = frm.orderdetailidx[i];

        if (objitem.type != "checkbox") {
        	continue;
        }

		if ((checkfound != true) || (forcereturnbyten == true) || (forcereturnbycustomer == true)) {
			// ���ð����� ��ǰ ���� Ȱ��ȭ
		    // TR �� bgColor �� ���Ѵ�.(���� FFFFFF �ΰ͸� Ȱ��ȭ �� �� �ִ�.)
		    e = objitem;
			while (e.tagName!="TR") {
				e = e.parentElement;
			}

			if (e.bgColor == "#ffffff") {
				objitem.disabled = false;
			}
		} else {
			if ((upbeaMakerid.length < 1) && (frm.odlvtype[i].value == "9") && (frm.makerid[i].value == "")) {
				alert("ERROR ERROR!!\n\n�ý���������!! ===============");
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
// ���õ� ��ǰ�� �ִ���
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
// ���õ� ��ǰ�� �귣�� ��������
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
        		// �ٹ�
	        	upbeaMakerid = "";
        	} else {
        		// ����
	        	upbeaMakerid = frm.makerid[i].value;
        	}
        	break;
        }
    }

    return upbeaMakerid;
}

// ============================================================================
// ���õ� ��ǰ���� Ư�� �귣�� �ִ���
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
		alert("ERR : ����Ұ�");
		return;
	}

	if (GetCheckedItemMakerExist(frm, IsTempEventAvail_Makerid) == true) {
		if (((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD01")) || ((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD06"))) {
			selectGubun('C004','CD11','����','�����ǰ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');
			return;
		} else {
			alert("���ɹ�ǰ�� ��쿡�� �����ϼ���.\n\n(�����ǰ�� �ֹ��� 1ȸ�� ���밡��!!)");
			return;
		}
		//
	} else {
		alert("�̺�Ʈ �귣��(" + IsTempEventAvail_Makerid + ") ��ǰ�� ���õ��� �ʾҽ��ϴ�.");
		return;
	}
}

// ============================================================================
// ��ü ���/��ǰ�ΰ�� -  ���ϸ���/���α� ȯ�� üũ �Ѵ�.
// ============================================================================
function CheckMileageETC(frm) {

    var allselected = IsAllSelected(frm);

	if (!frm.forcecouponreturn) {
		return;
	}

	// ���ʽ�����(����), ���ϸ���, Giftī��, ��ġ�� ȯ��
	if ((IsCSCancelProcess || IsCSReturnProcess) && (IsStatusRegister == true)) {
		if (allselected) {
			frm.forcecouponreturn.checked = true;
			frm.forcemileagereturn.checked = true;
			frm.forcegiftcardreturn.checked = true;
			frm.forcedepositreturn.checked = true;
		} else {
			// �̹� üũ�Ǿ� ������ �������� �ʴ´�.
		}
	}
}

// ============================================================================
// ���� ����� üũ
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
			alert("��߱޺Ұ�!!\n\n��ü ����� ��� ����� ������ ȯ���˴ϴ�.");
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
// ��ۺ� ����
//
// - �ܼ����� + �귣����ü��ǰ = �պ���ۺ�(�귣�庰 �����å�� ����, ������=2000��) ����
//
// - �ܼ����� + ��Ÿ = ��ǰ��ۺ� ����
//
// - ��ǰ�ҷ���ǰ �� = ��ۺ� ���� ����
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
	// ��ǰ�ø� ����Ѵ�.
	if (IsCSReturnProcess != true) {
		return;
	}

	// ȸ����ۺ� ����
    frm.ckreturnpay.checked = false;
    frm.ckreturnpayHalf.checked = false;
    frm.ckreturnpayZero.checked = false;
	IsCkReturnPayChanged = true;

	// ========================================================================
	// �ܼ�����(CD01)�� �ƴϸ� ���� ����, ������ʸ���(CD06)�߰�
	if ((value_gubun02 != "CD01")&&(value_gubun02 != "CD06")) {
		return;
	}

	if (IsChangeOrder) {
		// ��ȯ�ֹ��� ��ǰ�ݾ�,��ۺ� �ݾ����� ����Ѵ�.
		if ((frm.orgitemcostsum.value*1 == frm.refunditemcostsum.value*1) && (frm.orgbeasongpay.value*1 == 0)) {
			// �귣�� ��ü��ǰ����, ��ü������
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

			// �귣�� ��ü�������̸� �պ���ۺ� ����
			// ����ȸ���ΰ�� üũ�Ǿ� �ִ� ���귣�� Ȯ���Ͽ� �Ѱ��� �귣��� 4000�� ���� ������ �����ϸ� 4000�� �����Ѵ�.
			// ��ü��ǰ�� ��쵵 ���� ����ǰ�� �����Ƿ� ��� üũ�Ѵ�.
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
// ��ۺ�(��ǰ���� �±�ȯ)
//
// - �ܼ����� = �պ���ۺ�(������=2000*2 ��) ����
//
// - ��ǰ�ҷ���ǰ �� = ��ۺ� ���� ����
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
	// �ܼ�����(CD01)�� �ƴϸ� ���� ����, ������ʸ���(CD06)�߰�, �����ȯ(CD04)
	if ((value_gubun02 != "CD01") && (value_gubun02 != "CD06") && (value_gubun02 != "CD04")) {
		return;
	}

	frm.add_customeraddbeasongpay.value = GetUpcheDeliverPay(frm.requiremakerid.value) * 2;
	frm.add_customeraddmethod.value = "1";		// �ڽ�����
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
// ��ü��ǰ�ΰ�
// ������ ���õ� ��ǰ���� �Ǵ�.
// �������Ŀ��� divcd �� �Ǵ��Ѵ�.
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
// ��ü �±�ȯ�ΰ�
// ============================================================================
function IsUpcheReturnStateItemChange(frm) {
	return (((divcd == "A000") || (divcd == "A100")) && (frm.requireupche.value == "Y"));
}

// ============================================================================
// ��ǰ ��ü üũ �Ǿ�����
// (��ۺ�� üũ����)
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

		// ���� ��ǰ ���� ����üũ
		if (parseInt(itemno.value) !== (parseInt(regitemno.value) + parseInt(prevcsreturnfinishno.value))) { return false; }
    }

    return true;
}

// ============================================================================
// ��ǰ ��ü üũ �� ��� üũ�ȵ� ��ۺ� ����üũ
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
			// CS �������� ����
			eval("frm.gubun01_" + frm.orderdetailidx[i].value).value = frm.gubun01.value;
			eval("frm.gubun02_" + frm.orderdetailidx[i].value).value = frm.gubun02.value;
			eval("frm.gubun01name_" + frm.orderdetailidx[i].value).value = frm.gubun01name.value;
			eval("frm.gubun02name_" + frm.orderdetailidx[i].value).value = frm.gubun02name.value;
		}
	}
}

// ============================================================================
// �귣�� ��ǰ ��ü üũ �Ǿ�����(��ü)
// ����� �ش� �귣�� ��ü
// �ٹ��� ��� �ٹ� ��ü
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
// �Ѱ� �귣�� ��ü ���õǾ����� üũ
// �귣�尡 ���̸� ���ٹ�� ��üüũ
// ��ǰ�ΰ�� ���� ��ǰ,�±�ȯȸ��CS �Ϸ᳻�� ���� �ջ�
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
			// �ٹ�üũ
			if (isupchebeasong.value === "Y") { continue; }
		} else {
			// ����üũ
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
// ���þȵ� �귣�� ��ǰ�ݾ� �հ�(TODO : ��ǰ���� ���밡 -> �ǸŰ�(���ΰ�))
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
// ���õ� ��ǰ��
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
// �ΰ� �̻��� ��ǰ�� ǰ�� ����ϴ���
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
	// ȯ�Ҵ�� ��ǰ����(��ǰ�������밡)�հ�, �ÿ�ī�������հ�, �������������հ�
	CalculateCancelItemSUM(frm);

	// ��ۺ� �հ� ���
	CalculateBeasongPaySum(frm);

	// �����Ѿ� ���
	CalculateTotalBuyPaySum(frm);

	// ��ǰ�� ȸ�� ��ۺ� ���
	// CalculateReturnBeasongPay(frm);

	// ��ü �߰�����
	CalculateUpcheReturnBeasongPay(frm);

	// �������� ȯ�� ����, ȯ�� �����հ� ����
	// �������� ����, ��ǰ���� �ȺеǾ� ���������� ������. 2018-04-19
	// CalculateFixedCoupon(frm);

	// ���ϸ��� ȯ�� ����
	CalculateMileage(frm);

	// Giftī�� ����
	CalculateGiftCard(frm);

	// ��ġ�� ����
	CalculateDeposit(frm);

	// ��� �ݾ� �հ� ���
	CalculateTotal(frm);
}

// Giftī�� ����
function CalculateGiftCard(frm) {
    var orggiftcardsum	    = 0;
    var refundgiftcardsum    = 0;
    var remaingiftcardsum    = 0;

    var prevrefundsubtotalprice = 0;	// ��������հ� - �������� - ��Ÿ���� - �������� - ���ϸ��� + ��ǰ��ۺ� + �����ݾ�

	if (!frm.orggiftcardsum) { return; }

	prevrefundsubtotalprice = frm.refundtotalbuypaysum.value*1 + frm.refundpercentcouponsum.value*1 + frm.refundallatsubtractsum.value*1 + frm.refundfixedcouponsum.value*1 + frm.refundmileagesum.value*1;
	prevrefundsubtotalprice = prevrefundsubtotalprice + frm.refunddeliverypay.value*1 + frm.refundadjustpay.value*1;

	orggiftcardsum = frm.orggiftcardsum.value;
	refundgiftcardsum = frm.refundgiftcardsum.value;

	if ((IsStatusFinished != true) && (IsStatusFinishing != true)) {
		if (frm.forcegiftcardreturn.checked) {
			// ȯ�� Giftī�尡 ��ұݾ׺��� ū���
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
	// Giftī�� ����
	frm.refundgiftcardsum.value = refundgiftcardsum;
	frm.remaingiftcardsum.value = remaingiftcardsum;
}

// ��ġ�� ����
function CalculateDeposit(frm) {
    var orgdepositsum	    = 0;
    var refunddepositsum    = 0;
    var remaindepositsum    = 0;

	var prevrefundsubtotalprice = 0;	// ��������հ� - �������� - ��Ÿ���� - �������� - ���ϸ��� - Giftī�� + ��ǰ��ۺ� + �����ݾ�

	if (!frm.orgdepositsum) { return; }

	prevrefundsubtotalprice = frm.refundtotalbuypaysum.value*1 + frm.refundpercentcouponsum.value*1 + frm.refundallatsubtractsum.value*1 + frm.refundfixedcouponsum.value*1 + frm.refundmileagesum.value*1;
	prevrefundsubtotalprice = prevrefundsubtotalprice + frm.refundgiftcardsum.value*1;
	prevrefundsubtotalprice = prevrefundsubtotalprice + frm.refunddeliverypay.value*1 + frm.refundadjustpay.value*1;

	orgdepositsum = frm.orgdepositsum.value;
	refunddepositsum = frm.refunddepositsum.value;

	if ((IsStatusFinished != true) && (IsStatusFinishing != true)) {
		if (frm.forcedepositreturn.checked) {
			// ȯ�� ��ġ���� ��ұݾ׺��� ū���
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
	// ��ġ�� ����
	frm.refunddepositsum.value = refunddepositsum;
	frm.remaindepositsum.value = remaindepositsum;
}

// �����Ѿ� ���
function CalculateTotalBuyPaySum(frm) {
	if (!frm.orgtotalbuypaysum) {
		return;
	}

	frm.refundtotalbuypaysum.value = frm.refundbeasongpay.value*1 + frm.refunditemcostsum.value*1;
	frm.remaintotalbuypaysum.value = frm.remainbeasongpay.value*1 + frm.remainitemcostsum.value*1;
}

// ȯ�Ҵ�� ��ǰ����(��ǰ�������밡)�հ�, �ÿ�ī�������հ�, �������������հ�
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
			// ���û�ǰ�հ�
			if (ischecked == true) {
				if (itemid*1 != 0) {
					refunditemcostsum 					= refunditemcostsum + (itemcost * regitemno * 1);
					refundallatitemdiscountSum 			= refundallatitemdiscountSum + (allatitemdiscount * regitemno * 1);
				}

				refundpercentBonusCouponDiscountSum = refundpercentBonusCouponDiscountSum + (percentBonusCouponDiscount * regitemno * 1);
				refundetcDiscountDiscountSum = refundetcDiscountDiscountSum + (etcDiscountDiscount * regitemno * 1);
			}

			// ��Ҿȵ� ��ǰ ��ü�հ�
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
    // ��ǰ��� �ϴ� ���� �հ�
    if (frm.itemcanceltotal!=undefined){
		frm.itemcanceltotal.value = refunditemcostsum;
    }

	// ========================================================================
    // ��Ҿȵ� ��ǰ ��ü�հ� �ݾ�(����������), �������Ĵ� ������� ��Ҿʵ� ��ǰ�ݾ��հ�
    if (frm.orgitemcostsum!=undefined){
		orgitemcostsum = frm.orgitemcostsum.value;
    }

    // ���/��ǰ ��ǰ�Ѿ�
    if (frm.refunditemcostsum!=undefined){
        frm.refunditemcostsum.value = refunditemcostsum;
    }

    // ���þ���(������) ��ǰ�Ѿ�
    if (frm.remainitemcostsum!=undefined){
        frm.remainitemcostsum.value = orgitemcostsum - refunditemcostsum;
    }

	// ========================================================================
    // ��Ҿȵ� �������� ��ü�հ� �ݾ�
    if (frm.orgpercentcouponsum!=undefined){
		// aaaaaaaaaaaaaaaa frm.orgpercentcouponsum.value = orgpercentBonusCouponDiscountSum * -1;
    }

    // ���/��ǰ �������� �Ѿ�
    if (frm.refundpercentcouponsum!=undefined){
        frm.refundpercentcouponsum.value = refundpercentBonusCouponDiscountSum * -1;
		frm.refundcouponsum.value = frm.refundpercentcouponsum.value;
    }

    // ���þ���(������) �������� �Ѿ�
    if (frm.remainpercentcouponsum!=undefined){
		if (IsCSReturnProcess && IsStatusEdit) {
			// ��ǰ���� �� ���޸� �����ݾ� �Էµ� ���̽�
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
    // ��Ҿȵ� ��Ÿ���� ��ü�հ� �ݾ�
    if (frm.orgallatsubtractsum!=undefined){
		// aaaaaaaaaaaaaaaaa frm.orgallatsubtractsum.value = orgetcDiscountDiscountSum * -1;
    }

    // ���/��ǰ ��Ÿ���� �Ѿ�
    if (frm.refundallatsubtractsum!=undefined){
        frm.refundallatsubtractsum.value = refundetcDiscountDiscountSum * -1;
    }

    // ���þ���(������) ��Ÿ���� �Ѿ�
    if (frm.remainallatsubtractsum!=undefined){
        frm.remainallatsubtractsum.value = frm.orgallatsubtractsum.value*1 - frm.refundallatsubtractsum.value*1;
    }

    // ���� ������ �������ش�.(pop_cs_action_new_process.asp �� ���� �ѱ�� ���� �ʿ�)
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


// ��ۺ� �հ� ���
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

// ȯ�� �������� ���
function CalculateFixedCoupon(frm) {

	if (!frm.orgpercentcouponsum) {
		return;
	}

	var orgfixedcouponsum 		= 0;
	var refundfixedcouponsum	= 0;
	var remainfixedcouponsum 	= 0;

	var orgcouponsum			= 0; // tencardspend(�ֹ���� ó���Ϸ� ���� �޶���)
    var refundcouponsum    		= 0;
    var remaincouponsum    		= 0;

	var prevrefundsubtotalprice = 0;	// ��������հ� - �������� - ��Ÿ����

	prevrefundsubtotalprice = frm.refundtotalbuypaysum.value*1 + frm.refundpercentcouponsum.value*1 + frm.refundallatsubtractsum.value*1;

	// ========================================================================
	// ���׺��ʽ����� ���
	orgfixedcouponsum = frm.orgcouponsum.value - frm.orgpercentcouponsum.value;
	refundfixedcouponsum = frm.refundfixedcouponsum.value;

	if ((IsStatusFinished != true) && (IsStatusFinishing != true)) {
		if (frm.forcecouponreturn.checked) {
			// ȯ�������� ��ұݾ׺��� ū���
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
	// ���ʽ����� �ջ�(���� + ����)�� ���
	frm.remaincouponsum.value = frm.remainpercentcouponsum.value*1 + frm.remainfixedcouponsum.value*1;
	frm.refundcouponsum.value = frm.orgcouponsum.value - frm.remaincouponsum.value;
}

// ȯ�� ���ϸ��� ���
function CalculateMileage(frm) {

    var orgmileagesum	    = 0;	// miletotalprice
    var refundmileagesum    = 0;
    var remainmileagesum    = 0;

    var prevrefundsubtotalprice = 0;	// ��������հ� - �������� - ��Ÿ���� - �������� + ��ǰ��ۺ� + �����ݾ�

	if (!frm.forcemileagereturn) {
		return;
	}

	prevrefundsubtotalprice = frm.refundtotalbuypaysum.value*1 + frm.refundpercentcouponsum.value*1 + frm.refundallatsubtractsum.value*1 + frm.refundfixedcouponsum.value*1;
	prevrefundsubtotalprice = prevrefundsubtotalprice + frm.refunddeliverypay.value*1 + frm.refundadjustpay.value*1;

	orgmileagesum = frm.orgmileagesum.value;
	refundmileagesum = frm.refundmileagesum.value;

	if ((IsStatusFinished != true) && (IsStatusFinishing != true)) {
		if (frm.forcemileagereturn.checked) {
			// ȯ�����ϸ����� ��ұݾ׺��� ū���
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
	// ���ϸ��� ����
	frm.refundmileagesum.value = refundmileagesum;
	frm.remainmileagesum.value = remainmileagesum;
}

// ��ǰ�� ȸ�� ��ۺ� ���
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

    // ȸ�� ��ۺ�
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

	// ��ü �߰�����
    CalculateUpcheReturnBeasongPay(frm);
}

// ��ü �߰�����
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

// ��� �ݾ� �ջ�
function CalculateTotal(frm) {

	// ========================================================================
	// �����Ѿ�
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
	// ��� ���ʽ�����(����)
	var orgpercentcouponsum = 0;
	var refundpercentcouponsum = 0;
	var remainpercentcouponsum = 0;

	orgpercentcouponsum = frm.orgpercentcouponsum.value*1;
	refundpercentcouponsum = frm.refundpercentcouponsum.value*1;
	remainpercentcouponsum = frm.remainpercentcouponsum.value*1;

	// ========================================================================
	// ��� ��Ÿ����
	var orgallatsubtractsum = 0;
	var refundallatsubtractsum = 0;
	var remainallatsubtractsum = 0;

	orgallatsubtractsum = frm.orgallatsubtractsum.value*1;
	refundallatsubtractsum = frm.refundallatsubtractsum.value*1;
	remainallatsubtractsum = frm.remainallatsubtractsum.value*1;

	// ========================================================================
	// ��� ���ʽ�����(����)
	var orgfixedcouponsum = 0;
	var refundfixedcouponsum = 0;
	var remainfixedcouponsum = 0;

	orgfixedcouponsum = frm.orgfixedcouponsum.value*1;
	refundfixedcouponsum = frm.refundfixedcouponsum.value*1;
	remainfixedcouponsum = frm.remainfixedcouponsum.value*1;

	// ========================================================================
	// ��� ���ϸ���
	var orgmileagesum = 0;
	var refundmileagesum = 0;
	var remainmileagesum = 0;

	orgmileagesum = frm.orgmileagesum.value*1;
	refundmileagesum = frm.refundmileagesum.value*1;
	remainmileagesum = frm.remainmileagesum.value*1;

	// ========================================================================
	// ��� Giftī��
	var orggiftcardsum = 0;
	var refundgiftcardsum = 0;
	var remaingiftcardsum = 0;

	orggiftcardsum = frm.orggiftcardsum.value*1;
	refundgiftcardsum = frm.refundgiftcardsum.value*1;
	remaingiftcardsum = frm.remaingiftcardsum.value*1;

	// ========================================================================
	// ��� ��ġ��
	var orgdepositsum = 0;
	var refunddepositsum = 0;
	var remaindepositsum = 0;

	orgdepositsum = frm.orgdepositsum.value*1;
	refunddepositsum = frm.refunddepositsum.value*1;
	remaindepositsum = frm.remaindepositsum.value*1;

	// ========================================================================
	// �����ݾ�
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
	// ȸ�� ��ۺ�
	var refunddeliverypay = 0;

	refunddeliverypay = frm.refunddeliverypay.value*1;

	// ��Ÿ�����ݾ�
	var refundadjustpay = 0;
    // Ƽ���ֹ��ΰ��..================================================
    if ((IsTicketOrder==true)&&(mayTicketCancelChargePro>0)){
        if ((refundtotalbuypaysum!=0)&&(frm.refundadjustpay.value*1==0)){
            var mayTicketCancelPro = getFieldValue(frm.tRefundPro)*1;
            if (mayTicketCancelPro>0){
				if (mayTicketCancelPro > 100) {
					// �ݾ�(�ѵ� 10%)
					if ((refundtotalbuypaysum*10/100) > mayTicketCancelPro) {
						alert( ticketCancelStr + 'Ƽ�� ��� ������ ' + mayTicketCancelPro + '�� ���� \n\n(��, ���� �ֹ��� ��ҽô� ����)' );
						frm.refundadjustpay.value = mayTicketCancelPro*-1;
					} else {
						alert( ticketCancelStr + 'Ƽ�� ��� ������ 10% ���� \n\n(��, ���� �ֹ��� ��ҽô� ����)' );
						frm.refundadjustpay.value = (refundtotalbuypaysum*10/100)*-1;
					}
				} else {
					// �ۼ�Ʈ
					alert( ticketCancelStr + 'Ƽ�� ��� ������ ' + mayTicketCancelPro + '% ���� \n\n(��, ���� �ֹ��� ��ҽô� ����)' );
					frm.refundadjustpay.value = (refundtotalbuypaysum*mayTicketCancelPro/100)*-1;
				}
            }
        }
    }
    // Ƽ���ֹ��ΰ��..================================================


	refundadjustpay = frm.refundadjustpay.value*1;

	// ========================================================================
	// ��ұݾ�
	var orgsubtotalprice = 0;
	var refundsubtotalprice = 0;
	var remainsubtotalprice = 0;

	orgsubtotalprice = orgtotalrealbuypaysum;
	refundsubtotalprice = refundtotalrealbuypaysum + refunddeliverypay + refundadjustpay;
	remainsubtotalprice = orgsubtotalprice - refundsubtotalprice;

	frm.orgsubtotalprice.value = orgsubtotalprice;
	frm.refundsubtotalprice.value = refundsubtotalprice;
	frm.remainsubtotalprice.value = remainsubtotalprice;

	// ���� ������ �������ش�.(pop_cs_action_new_process.asp �� ���� �ѱ�� ���� �ʿ�)
	frm.subtotalprice.value = orgsubtotalprice;
	frm.canceltotal.value = refundsubtotalprice;
	frm.nextsubtotal.value = remainsubtotalprice;

	// ========================================================================
	// ȯ�ұݾ� �Է�
	if ((parseInt(frm.ipkumdiv.value) >= 4) || IsChangeOrder) {
        if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true)) {
            if (frm.refundrequire!=undefined) {
                frm.refundrequire.value = frm.refundsubtotalprice.value*1;
            }
        }
    }
}

// ============================================================================
// ���õ� ��ü��ۺ� �հ�
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
// �����ǰ ���õǾ�����
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
// ���þʵ� ��ǰ �Ⱥ��̱�
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
// ��ü ��ǰ ǥ��
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
// ���þʵ� ��ü��ۺ� �հ�(��ü)
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
// ���þʵ� ��ü��ۺ� �հ�(�귣�庰)
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
// ȸ�� ��ۺ� �ߺ�üũ
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
// ���� ��ǰ�� �����ǰ ���� ����
// ============================================================================
function CheckForceReturnByTen(obj) {
	// ��ǰ����� üũ
	CheckForItemChanged();

	if (obj.checked == true) {
		frmaction.ForceReturnByCustomer.checked = false;
		frmaction.ForceReturnByCustomer.disabled = true;
	} else {
		frmaction.ForceReturnByCustomer.disabled = false;
	}
}

// ============================================================================
// ���� ��ǰ�� ��������ǰ ���� ����
// ============================================================================
function CheckForceReturnByCustomer(obj) {
	// ��ǰ����� üũ
	CheckForItemChanged();

	if (obj.checked == true) {
		frmaction.ForceReturnByTen.checked = false;
		frmaction.ForceReturnByTen.disabled = true;
	} else {
		frmaction.ForceReturnByTen.disabled = false;
	}
}

// ============================================================================
// ȯ�ҹ�� ����
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
        //������ ȯ��
        document.all.refundinfo_R007.style.display = "";
    }else if((comp.value=="R020")||(comp.value=="R022")||(comp.value=="R080")||(comp.value=="R100")||(comp.value=="R120")||(comp.value=="R400")||(comp.value=="R420")){
        //�ǽð� ��ü ���//�ǽð� ��ü �κ� ���//ALL@ ���� ��� //�ſ�ī�� ���� ��� //�ſ�ī�� �κ����//�޴��� ���� ��� //�޴��� �κ����
        document.all.refundinfo_R100.style.display = "";
    }else if(comp.value=="R550"){
        //������ ���� ���
        document.all.refundinfo_R550.style.display = "";
    }else if(comp.value=="R560"){
        //����Ƽ�� ���� ���
        document.all.refundinfo_R560.style.display = "";
    }else if(comp.value=="R050"){
        //������ ���� ���
        document.all.refundinfo_R050.style.display = "";
    }else if ((comp.value=="R900") || (comp.value=="R910")) {
        //���ϸ��� ȯ��, ��ġ�� ȯ��
        document.all.refundinfo_R900.style.display = "";
    }

}

// ============================================================================
// ��ü �߰����� ���� �귣�� ���̵� ��������
// ============================================================================
function InsertCheckedUpcheID(frm) {
    var checkedUpcheid = "";
    var UpcheDuplicated = false;
    var IsUpcheReturn;

	// �����ÿ��� �����Ѵ�.
	if (IsStatusRegister != true) {
		return;
	}

	// ��ǰ����(��ü���), �±�ȯ���, ��ǰ���� �±�ȯ���, ��ü��Ÿ����, ������߼�, ���񽺹߼�, ��Ÿȸ��
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
		// ���ٹ�ۻ�ǰ �� ������ǰ
		if (frm.ForceReturnByCustomer.checked == true) {
			frm.buf_requiremakerid.value = "10x10logistics";
		}
	}
}

// ============================================================================
// CS �����ÿ��� ��ǰ����(��ü���)/ȸ����û(�ٹ����ٹ��) �� �������� �ʰ�
// ����� �귣�������� �ִ°�� ��ǰ����(��ü���), ���°�� ȸ����û(�ٹ����ٹ��) ���� �����Ѵ�.
// ������ �����Ѱ�� ������ �������� �ʴ´�.
// ��������ǰ�� �����Ȱ�� ��ü������� �����Ѵ�.
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
		// ���ٹ�ǰ
		frm.divcd.value = "A010";
		if (frm.title.value == "��ǰ����") {
			frm.title.value = "ȸ����û(�ٹ����ٹ��)";
		}
	} else {
		// ��ü��ǰ
		frm.divcd.value = "A004";
		if (frm.title.value == "��ǰ����") {
			frm.title.value = "��ǰ����(��ü���)";
		}
	}
}

// ============================================================================
//�߰������ۺ� ���� ����
// ============================================================================
function Change_add_upchejungsancause(comp) {
	var jungsancause = document.all.add_upchejungsancause;

	if (divcd != "A700") {
		// ��ü��Ÿ���� �ƴϸ� ��ۺ� ���ð���
		if ((jungsancause.value != "��ۺ�") && (jungsancause.value != "������")) {
			alert("\n\n��ü��Ÿ������ �ƴϸ� [��ۺ�,������] �� ���ð����մϴ�.\n\n��ǰ��� �� �ٸ� ���곻���� ��ü��Ÿ�������� ����ϼ���.\n[����̻��밡��]\n\n");
			jungsancause.value = "";
		}
	} else {
		if ((jungsancause.value != "��ۺ�") && (jungsancause.value != "������")) {
			//if ((IsCsPowerUser != true) && (HasAuthUpcheJungsanItemPrice != true)) {
			if ( (C_CSpermanentUser != true) && (IsC_ADMIN_AUTH != true) ) {
				alert("\n\n[��ۺ�,������] �̿��� ������ ����̻��밡��\n\n");
				jungsancause.value = "";
			} else {
				if (jungsancause.value == "��ǰ���") {
					alert("\n\n��ǰ���(��ǰ�� ���԰�)�� �߰�����ݾ��� 0������ �Էµ˴ϴ�.\n(�����԰��� �ۼ� �� ����˴ϴ�.)\n\n�������뿡 ��ǰ�� ���԰��� ��Ȯ�� �����ϼ���.");
					document.all.add_upchejungsandeliverypay.value = "0";
					document.all.buf_totupchejungsandeliverypay.value = "0";
				//} else if (jungsancause.value = "�����Է�") {
				//	jungsancause.value = "";
				//	alert("�����Ұ�!!!");
				}
			}
		}
	}

    if (comp.value=="�����Է�") {
        document.all.span_add_upchejungsancauseText.style.display = "inline";
    }else{
        document.all.span_add_upchejungsancauseText.style.display = "none";
    }
}

// ============================================================================
//�߰������ۺ� �Է�
// ============================================================================
function Change_add_upchejungsandeliverypay(comp){
	var jungsancause = document.all.add_upchejungsancause;

	if ((jungsancause.value == "��ǰ���") && (comp.value*1 != 0)) {
		comp.value = "0";
		alert("\n\n��ǰ����� �߰�����ݾ��� 0������ �Էµ˴ϴ�.\n(�����԰��� �ۼ� �� ����˴ϴ�.)\n\n�������뿡 ��ǰ�� ���԰��� ��Ȯ�� �����ϼ���.");
	}

    comp.form.buf_totupchejungsandeliverypay.value = comp.form.buf_refunddeliverypay.value*1 + comp.value*1;

    if (isNaN(comp.form.buf_totupchejungsandeliverypay.value)){
        comp.form.buf_totupchejungsandeliverypay.value = "0";
    }
}

// ============================================================================
// ��ü �߰� ���� �Է� ����
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

	// ��ǰ����� üũ
	CheckForItemChanged();
}

// ============================================================================
// ��ü��ۺ� ������ ���Ǳݾ�
// ============================================================================
function GetUpcheFreeBeasongLimit(makerid) {
	for (var i = 0; i < arrmakerid.length; i++) {
		if (arrmakerid[i] == makerid) {
			return arrdefaultfreebeasonglimit[i];
		}
	}

	// ������ �ٹ����� ���رݾ�
	return 30000;
}

// ============================================================================
// ��ü��ۺ�
// ============================================================================
function GetUpcheDeliverPay(makerid) {

	var frm = document.frmaction;

	var savedrefunddeliverypay = 0;

	if (frm.refunddeliverypay) {
		savedrefunddeliverypay = frm.refunddeliverypay.value * -1;
	}

	// ���� ���Ŀ��� ��ۺ� ��å�� �ٲ��� �Էµ� �ݾ����� ��ۺ� �����Ѵ�.
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

	// ������ �ٹ����ٹ�ۺ�
	return CDEFAULTBEASONGPAY;
}


// ============================================================================
// ��ü��ۺ�(��ǰ���� �±�ȯ)
// ============================================================================
function GetUpcheDeliverPayItemChange(makerid) {

	var frm = document.frmaction;

	for (var i = 0; i < arrmakerid.length; i++) {
		if ((arrmakerid[i] == makerid) && (arrdefaultdeliverpay[i] != 0)) {
			return arrdefaultdeliverpay[i];
		}
	}

	// ������ �ٹ����ٹ�ۺ�
	return CDEFAULTBEASONGPAY;
}


// ============================================================================
// �ֹ���ҷ� ��ۺ��߰��� �ʿ�����
// ============================================================================
/*
function IsAddBrandBeasongPayNeed(frm, makerid) {
	var value_gubun02 = frm.gubun02.value;

	// ��ҽÿ��� ���, �ܼ����� �ܻ̿������� �߰� ����.
	if ((IsCSCancelProcess != true) || (value_gubun02 != "CD01")) {
		return false;
	}

	// ��ü����
	if (IsOneBrandAllSelected(frm, makerid) == true) {
		return false;
	}

	// �̹� ��ۺ� �ִ� ���
	if (GetNotCheckedUpcheBeasongPayByBrand(frm, makerid) > 0) {
		return false;
	}

	// aaaaaaaaaaaaaaaaa
	// �����ۻ�ǰ �Ǵ� ���� ��ǰ�� �ִ���

	// ���þȵȻ�ǰ�� 30000���� �̸��� ���
	if (GetOneBrandNotSelectedItemcost(frm, "") < GetUpcheFreeBeasongLimit(makerid)) {
		return true;
	}

	return false;
}
*/

// ============================================================================
// CS ����
// ============================================================================
function CsRegProc(frm) {
    if ((IsTicketOrder==true)&&(ticketCancelDisabled==true)){
        if (!confirm('��� ������ ���� ��� �Ұ��մϴ�. ' + ticketCancelStr + ' \n\n ����Ͻðڽ��ϱ�?')){
            return;
        }

        //���� �ִ��� check
        if (IsCsPowerUser!=true){
            alert('������ �����ϴ�. ��Ʈ��� ���� ���');
            return;
        }
    }

    if ((IsTravelOrder==true)&&(travelCancelDisabled==true)){
        if (!confirm('��� ������ ���� ��� �Ұ��մϴ�. ' + travelCancelStr + ' \n\n ����Ͻðڽ��ϱ�?')){
            return;
        }

        //���� �ִ��� check
        if (IsCsPowerUser!=true){
            alert('������ �����ϴ�. ��Ʈ��� ���� ���');
            return;
        }
    }

    if (((IsGiftingOrder == true) || (IsGiftiConOrder == true)) && (IsOrderCancelDisabled == true) && IsCSReturnProcess) {
        if (!confirm('��ǰ�Ұ� �ֹ��Դϴ�. : ' + OrderCancelDisableStr + ' \n\n ����Ͻðڽ��ϱ�? [�����ڱ��� �ʿ�]')) {
            return;
        }

        //���� �ִ��� check
        if (IsCsPowerUser != true) {
            alert('������ �����ϴ�. ��Ʈ��� ���� ���');
            return;
        }
    }

	var forcereturnbyten = GetForceReturnByTen(frm);

	// ������ üũ
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	if (divcd == "A000") {
		// ���ϻ�ǰ ��ȯ����� ���, �����ȯ ���� ����
		if ((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD04")) {
	        alert('�����ȯ�� ��� �ɼǺ��� ��ȯ��� ����ϼ���');
	        return;
		}
	}

	// ���, ��ǰ
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true)) {
		// �귣�� ��ü��� �Ǵ� �ֹ�����̸鼭 ��ۺ� ���� ��� üũ
        if (CheckBeasongPayOnlyRemain(frm) != true) {
            return;
        }
	}

	// ���� ��ǰ üũ �� ����
	if ((divcd != "A100") && (divcd != "A111") && (divcd != "A112")) {
	    if (!SaveCheckedItemList(frm)) {
	        return;
	    }
	} else {
	    if (!SaveChangeCheckedItemList(frm)) {
	        return;
	    }
	}

	// ���, ��ǰ, ȯ��
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true) || (divcd == "A003")) {
        if (CheckReturnForm(frm) != true) {
            return;
        }
	}

	// ��ǰ��ۺ� üũ
	if (IsCSReturnProcess == true) {
		if (CheckRefundDeliverYpay(frm) != true) {
            return;
        }
	}

	// ���� ȯ�ҿ�û ���
	if (divcd == "A003") {
		if (frm.refundrequire) {
			if ((frm.refundrequire.value*1 > RefundAllowLimit) && (RefundAllowLimit != -1)) {
		        alert('������ �����ϴ�. ������ �Ǵ� ��Ʈ��� ���� ���');
		        frm.refundrequire.focus();
		        return;
			}
		}
	}

	// ���, ��ǰ
    if ((IsCSCancelProcess) || (IsCSReturnProcess)) {
		// ȯ�������� �������� üũ
		if (IsRefundInfoOK(frm) != true) {
			return;
		}

		// ȯ�Ҽ����� �ǹٸ���
		if (CheckReturnMethod(frm) != true) {
			return;
		}
    }

	// 2018-01-19, skyer9, ���� ��߱� �ȳ�â ����
	if (IsCSReturnProcess) {
		if ((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD01")) {
			// ���ɹ�ǰ ����
		} else {
			if (frm.tmpcopycouponinfo) {
				if (frm.tmpcopycouponinfo.checked == false) {
					if (confirm("\n���ɹ�ǰ�� �ƴ� ��� ������ ��߱��� �帱 �� �ֽ��ϴ�.\n(���常 ��߱� �����մϴ�.)\n\n���� ��߱��Ͻðڽ��ϱ�?\n\n\n") == true) {
						frm.tmpcopycouponinfo.checked = true;
						jsCheckCopyCoupon(frm);
					}
				}
			}
		}
	}

	if ($("#needChkYN_X").val()) {
		if ($("#needChkYN_X").prop("checked") === false && $("#needChkYN_F").prop("checked") === false) {
			alert('�Ϸᱸ���� �����ϼ���.');
			$("#needChkYN_X").focus();
			return;
		}
	}

    //�߰� ���� ����
    if (frm.add_upchejungsandeliverypay){
        if (!IsInteger(frm.add_upchejungsandeliverypay.value)){
            alert('���ڸ� �����մϴ�.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*1!=0){
            if (frm.add_upchejungsancause.value.length<1){
                alert('�߰� ���� ������ �Է��ϼ���.');
                frm.add_upchejungsancause.focus();
                return;
            }else if ((frm.add_upchejungsancause.value=='�����Է�')&&(frm.add_upchejungsancauseText.value.length<1)){
                alert('�߰� ���� ������ �Է��ϼ���.');
                frm.add_upchejungsancauseText.focus();
                return;
            }

            if (frm.buf_requiremakerid.value.length<1){
                alert('�߰� ������� �ִ°�� �귣�� ���̵� �����Ǿ�� �մϴ�. ');
                return;
            }

            //�ֹ� ������ ���̵� �ִ� ��츸.
        }else{

			if ((divcd == "A700") && (frm.add_upchejungsancause.value != '��ǰ���')) {
	            alert('�߰� ������� �Է��ϼ���.');
	            frm.add_upchejungsandeliverypay.focus();
	            return;
			}
        }
    }

	if ((IsCSCancelProcess == true) && IsStatusRegister) {
        if (frm.modifyitemstockoutyn) {
        	if (frm.modifyitemstockoutyn.checked == true) {
		        if (IsMultiStockOutItemChecked(frm) == true) {
		            // �ΰ� �̻��� ��ǰ�� ǰ������ �Ǿ� �ִ� ���
		            if (confirm("\n\n\n============ ���� ��ǰ�� [ǰ��] ����մϴ�. =================\n\n�����Ͻðڽ��ϱ�?") != true) {
		            	return;
		            }
		        }
        	}
        }
	}

    if (IsCSCancelProcess){
        if(confirm("��� ���� �Ͻðڽ��ϱ�?")) {
			DisableAllButton();
            frm.submit();
        }
    }else if (IsCSReturnProcess){
        if (frm.requireupche.value=="Y"){
            if(confirm("��ü [" + frm.requiremakerid.value +"]�� ��ǰ/ȸ��/��ȯ ���� �Ͻðڽ��ϱ�?")){
                ChangeCSTitleGubun(frm);
				DisableAllButton();
                frm.submit();
            }
        }else{
            if(confirm("[�ٹ����� ��������]�� ��ǰ/ȸ��/��ȯ ���� �Ͻðڽ��ϱ�?")){
                ChangeCSTitleGubun(frm);
				DisableAllButton();
                frm.submit();
            }
        }
    }else if (IsCSRefundNeeded) {
        if(confirm("ȯ�� ���� �Ͻðڽ��ϱ�?")){
			DisableAllButton();
            frm.submit();
        }
    }else if(confirm("���� �Ͻðڽ��ϱ�?")){
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
			alert("��ǰ��ۺ� �Է��ϼ���.");
			frm.refunddeliverypay.focus();
			return false;
		}

		if (frm.refunddeliverypay.value*0 != 0) {
			alert("��ǰ��ۺ�� ���ڸ� �����մϴ�.");
			frm.refunddeliverypay.focus();
			return false;
		}

		if (frm.refunddeliverypay.value*1 > 0) {
			alert("��ǰ��ۺ�� ���̳ʽ� �ݾ׸� �Է°����մϴ�.");
			frm.refunddeliverypay.focus();
			return false;
		}

		if (frm.addbeasongpay) {
			if ((frm.addmethod[1].checked != true) && (frm.addmethod[2].checked != true) && (frm.addmethod[3].checked != true)) {
				frm.addmethod[0].checked = true;
			}

			if (frm.addbeasongpay.value*1 != 0) {
				if ((frm.refunddeliverypay.value*1 != frm.addbeasongpay.value*1) && (frm.addmethod[0].checked == true)) {
					alert("��ۺ� �߰����� ����� �����ϼ���.");
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
        alert("���� ������ �����ϼ���.");
        frm.divcd.focus();
        return false;
    }

    if (frm.title.value.length<1) {
        alert("������ �Է��ϼ���.");
        frm.title.focus();
        return false;
    }

    if (frm.gubun01.value.length<1) {
        alert("���� ������ �Է��ϼ���.");
        return false;
    }

    return true;
}

// ============================================================================
// ȯ�� �ݾ��� ��������
// ============================================================================
function IsRefundInfoOK(frm) {

	if ((IsCSCancelProcess != true) && (IsCSReturnProcess != true)) {
		return true;
	}

	// ���, ��ǰ�� : ȯ�ұݾװ� ��
	if (frm.orgsubtotalprice && frm.refundsubtotalprice) {
	    if (frm.orgsubtotalprice.value*1 < frm.refundsubtotalprice.value*1) {
	        alert('�����ݾ� �̻����� ȯ���� �� �����ϴ�.\n\n���ϸ���, ���� ���� ȯ��üũ�ϼ���.');
	        if (IsAdminLogin != true) {
	        	return false;
	        }
	    }
	}

	if (frm.returnmethod) {
		if (frm.returnmethod.value == "R000") {
			// ȯ�Ҿ��� �̸� üũ ���Ѵ�.
			frm.refundrequire.value = "0";
			return true;
		}
	}

	if (frm.refundrequire && frm.returnmethod) {
	    if ((frm.refundsubtotalprice.value*1 < 1) && ((frm.returnmethod.value != "R000"))) {
	        alert('ȯ�Ҵ�� �ݾ��� �����ϴ�.\n\nȯ�Ҿ��� �Ǵ� ����, ���ϸ��� ���� ȯ��üũ �����ϼ���');
			// ���ε� �Ұ�!!, 2017-05-04, skyer9
	        // if (IsAdminLogin != true) {
	        	return false;
	        // }
	    }
	}

	if (frm.mainpaymentorg && frm.prevrefundsum && frm.refundrequire) {
		if (frm.mainpaymentorg.value*1 > 0) {
			// ������ ����
			if (frm.mainpaymentorg.value*1 < (frm.prevrefundsum.value*1 + frm.refundrequire.value*1)) {
				alert('�����ݾ� �̻����� ȯ���� �� �����ϴ�. - �����̳� ���ϸ��� ȯ���� üũ�� �ּ���.');
				return false;
			}
		}
	}

	if (frm.remainsubtotalprice) {
	    if (frm.remainsubtotalprice.value*1 < 0) {
	        alert('��� �� ���� �ݾ��� ���̳ʽ��� �� �� �����ϴ�. - �����̳� ���ϸ��� ȯ���� üũ�� �ּ���.');
			return false;

			// �����ڵ� �Ұ�, skyer9, 2018-01-04
			/*
	        if (IsAdminLogin != true) {
	        	return false;
	        }
			*/
	    }
	}

	if (frm.orgpercentcouponsum) {
		if ((frm.orgpercentcouponsum.value*1 != 0) && (frm.orgfixedcouponsum.value*1 != 0)) {
			alert('\n\n���� �ݾ� ���� : �ý����� ����[��������]!!!\n\n');
			return false;
		}

		if ((frm.orgpercentcouponsum.value*1 > 0) || (frm.orgallatsubtractsum.value*1 > 0) || (frm.orgfixedcouponsum.value*1 > 0) || (frm.orgcouponsum.value*1 > 0) || (frm.orgmileagesum.value*1 > 0) || (frm.orggiftcardsum.value*1 > 0) || (frm.orgdepositsum.value*1 > 0)) {
			alert('\n\n���� �ݾ� ���� : �ý����� ����[�ݾ��÷���]!!!\n\n');
			return false;
		}

		if ((frm.refundpercentcouponsum.value*1 > 0) || (frm.refundallatsubtractsum.value*1 > 0) || (frm.refundfixedcouponsum.value*1 > 0) || (frm.refundcouponsum.value*1 > 0) || (frm.refundmileagesum.value*1 > 0) || (frm.refundgiftcardsum.value*1 > 0) || (frm.refunddepositsum.value*1 > 0)) {
			alert('\n\ȯ�� �ݾ� ���� : �ý����� ����!!!\n\n');
			return false;
		}

		if ((frm.orgpercentcouponsum.value*1 == 0) && (frm.refundpercentcouponsum.value*1 != 0)) {
			alert('\n\n����/ȯ�� �ݾ� ���� : �ý����� ����[��������]!!!\n\n');
			return false;
		}

		if ((frm.orgallatsubtractsum.value*1 == 0) && (frm.refundallatsubtractsum.value*1 != 0)) {
			alert('\n\n����/ȯ�� �ݾ� ���� : �ý����� ����[��Ÿ����]!!!\n\n');
			return false;
		}

		if ((frm.orgfixedcouponsum.value*1 == 0) && (frm.refundfixedcouponsum.value*1 != 0)) {
			alert('\n\n����/ȯ�� �ݾ� ���� : �ý����� ����[��������]!!!\n\n');
			return false;
		}

		if ((frm.orgcouponsum.value*1 == 0) && (frm.refundcouponsum.value*1 != 0)) {
			alert('\n\n����/ȯ�� �ݾ� ���� : �ý����� ����[����]!!!\n\n');
			return false;
		}

		if ((frm.orgmileagesum.value*1 == 0) && (frm.refundmileagesum.value*1 != 0)) {
			alert('\n\n����/ȯ�� �ݾ� ���� : �ý����� ����[���ϸ���]!!!\n\n');
			return false;
		}

		if ((frm.orggiftcardsum.value*1 == 0) && (frm.refundgiftcardsum.value*1 != 0)) {
			alert('\n\n����/ȯ�� �ݾ� ���� : �ý����� ����[����Ʈ]!!!\n\n');
			return false;
		}

		if ((frm.orgdepositsum.value*1 == 0) && (frm.refunddepositsum.value*1 != 0)) {
			alert('\n\n����/ȯ�� �ݾ� ���� : �ý����� ����[��ġ��]!!!\n\n');
			return false;
		}
	}

	return true;
}

// ============================================================================
// ���û�ǰ ����
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
                alert('������ �Է��ϼ���.');
                e.focus();
                e.select();
                return false;
			}

			if ((IsStatusRegister == true) && ((e.value*1) == 0)) {
                alert('������ �Է��ϰų� ������ �����ϼ���.');
                e.focus();
                e.select();
                return false;
			}

			if ((IsStatusRegister != true) && ((e.value*1) < 0)) {
                alert('������ 0 ���� ���� �� �����ϴ�.');
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
    //��ۺ� ----------------------------------------
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

    //��Ÿ����, ���񽺹߼� , ��Ÿȸ��, ȯ�ҿ�û, �������ǻ���, ��ü �߰� ���� - �󼼳��� üũ ����.
    if ((divcd=="A009") || (divcd=="A002") || (divcd=="A200") || (divcd=="A003") || (divcd=="A005") || (divcd=="A006") || (divcd=="A007") || (divcd=="A700") || (divcd=="A100") || (divcd=="A111")) {
        // no- check

    }else{
        if (!checkitemExists){
            alert('���õ� �󼼳����� �����ϴ�.');
            return false;
        }
    }

    return true;
}

// ============================================================================
// ���û�ǰ ����(��ǰ���� �±�ȯ)
// ============================================================================
function SaveChangeCheckedItemList(frm) {
    var e;
    var ischecked = false;

	var prevoneorderdetailidx = "";		// �±�ȯ ȸ���ϴ� ��ǰ(order detail idx or 0(�������ΰ��))
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
		// CS ����â���� ������ �� ����.(�ɼ� ����â �Ǵ� �ֹ�����â���� ����)
		alert("�߸��� �����Դϴ�.");
		return false;
	}

    frm.detailitemlist.value = "";
    frm.csdetailitemlist.value = "";

	// �±�ȯ ȸ���ϴ� ��ǰ
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

	// �±�ȯ ����ϴ� ��ǰ
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
// ȯ�� ������ ��������
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
		// ȯ�Ҿ��� �̸� üũ ���Ѵ�.
		frm.refundrequire.value = "0";
		return true;
		 */

		// ȯ�ұݾ� �ִ� ��� ȯ�Ҿ��� ���ø��ϰ� ����(2015-01-05, skyer9)
		if (frm.refundsubtotalprice.value*1 > 0) {
			alert('ȯ�ұݾ��� �ִ� ���, ȯ�Ҿ��� ������ �� �����ϴ�.');
			frm.returnmethod.focus();
			return false;
		}
	}

    //2011-05-24 ���� / ����Ұ� ������� frm.refundsubtotalprice.value*1<>MainPaymentOrg �̰������ ����;;
    //if ((allselected) && (PayedNCancelEqual != true) && (IsCSCancelProcess)) {
    if ((allselected) && ((frm.orgsubtotalprice.value*1!=frm.refundsubtotalprice.value*1)) && (IsCSCancelProcess)) {
        if (!IsTicketOrder){
            alert('��ü ����ΰ�� �����ݾ� ��ü�� ȯ���ؾ��մϴ�. - ����ۺ� ȯ��, ���ϸ���, ���α� ���� üũ���ּ���.');
            //alert(MainPaymentOrg);
            //alert(frm.orgsubtotalprice.value*1);
            //alert(frm.refundsubtotalprice.value*1);
            return false;
        }
    }

    if (((PayedNCancelEqual != true)) && ((frm.returnmethod.value=="R100") || (frm.returnmethod.value=="R550") || (frm.returnmethod.value=="R560") || (frm.returnmethod.value=="R020") || (frm.returnmethod.value=="R080") || (frm.returnmethod.value=="R400"))) {
        alert('�Ϻ���� ���Ŀ��� ����� �� ���� ȯ�Ҽ����Դϴ�.(�ſ�ī�� �Ϻ���� ��)\n\n[�� �ְ����ݾ� : ' + MainPaymentOrg + ']. \n\n�κ���� �Ǵ� �ٸ� ȯ�Ҽ����� ������ �ּ���');
        frm.returnmethod.focus();
        return false;
    }

    //���ݾװ� ������� ��ü��Ҹ� �ؾ���.
    if ((PayedNCancelEqual == true) && ((frm.returnmethod.value=="R120") || (frm.returnmethod.value=="R022") || (frm.returnmethod.value=="R420"))){
        alert('ȯ�� �ݾװ� �� �ְ����ݾ��� ������ ���, �κ���� ���Ұ�  \n\n[�� �ְ����ݾ� : ' + MainPaymentOrg + ']. \n\n- ��ü��� �����.');
        frm.returnmethod.focus();
        return false;
    }

    //�κ���� ���� ALERT
    if (frm.returnmethod.value=="R120") {
        //alert(cardPartialCancelok + "," + frm.cardcode.value + "," + installment + "," + MainPaymentOrg + "," + precardcancelsum + "," + frm.refundrequire.value)
        if (cardPartialCancelok!="Y"){
            alert('�κ� ��� ���� ī�尡 �ƴմϴ�.');
            return false;
        }
        // --1 BC ī���� ��� 90%������ �κ���� ���� // BC ī���� ��� �κ������ �ܾ��� 5���� �̸��̾ ���ŷ� �Һΰ� �״�� �����.
        if (frm.cardcode.value=="11"){
            if (MainPaymentOrg*1!=0){
                if (precardcancelsum*1 + frm.refundrequire.value*1!=precardcancelsum*1){  //������ ��� (��ü)�� �������.
                    if (((precardcancelsum*1 + frm.refundrequire.value*1)/MainPaymentOrg*1)>90){
                        alert('BCī�� �� ��� �κ���� �հ���� ���ݾ��� 90% �̻� �� �� �� �����ϴ�. �ٸ� ȯ�Ҽ������� ó���ϼ���.');
                        return false;
                    }
                }
            }
        }

        // --// BC ī�尡 �ƴ� ��� �κ������ �ܾ��� 5���� �̸��̸�, �Ͻúҷ� ����� �� ������ �ȳ�.
        // (����, ��ȯ, ����, ����, �Ｚ - �����Ŀ��� ���ŷ� �Һΰ� �״�� �����. �Ե��� ������ �Ͻúҷ� ����.)
        if ((MainPaymentOrg*1>=50000)&&(MainPaymentOrg*1-(frm.refundrequire.value*1 + precardcancelsum*1)<50000)&&(installment*1>0)){
            //�Ե�.
            if (frm.cardcode.value=="03"){
                if (!confirm('�Ե� ī���� ��� �κ������ �ܾ��� 5���� �̸��̸� �Ͻúҷ� ��ȯ�˴ϴ�.')){
                    return false;
                }
            }
            if (isThisdateCancel=="Y"){
                //�������
                if (!confirm('����,��ȯ,����,�Ｚ ī���� ��� ������ �κ����(���� ���)�� �ܾ��� 5���� �̸��̸� �Ͻúҷ� ��ȯ�˴ϴ�.')){
                    return false;
                }
            }
        }

    }

    //�κ���� ���� ALERT
    if (frm.returnmethod.value=="R420") {
        if (phonePartialCancelok!="Y") {
            alert('�κ� ��� ���� ������ �ƴմϴ�.');
            return false;
        }
    }

    //�κ���� ���� ALERT (���̹����� �ǽð� ��ü)
    if (frm.returnmethod.value=="R022") {
        if (pggubun=="NP" || pggubun=="KK"){
            return true;
		}
		else{
			alert('���̹����̸� �ǽð���ü�κ���Ұ� �����մϴ�.');
            return false;
		}
    }


	if (sitename != "10x10") {
	    //�ܺθ��� ��� �ܺθ� ȯ�������� ����..
	    if ((frm.returnmethod.value != "R050") && (frm.returnmethod.value != "R000")) {
	        alert('�ܺθ��� ��� ȯ�� ���� �Ǵ� �ܺθ� ȯ���� �����ϼ���. \n\n���� ����ڸ� ���� ���޸����� ��� ȯ�� ó�� �մϴ�.');
	        frm.returnmethod.focus();
	        return
	    }
	}

    if (frm.refundrequire.value*1 != frm.refundsubtotalprice.value*1) {
        if ((frm.returnmethod.value!="R007") && (frm.returnmethod.value!="R900") && (frm.returnmethod.value!="R910") && (frm.returnmethod.value!="R000")) {
            alert('ȯ�� �ݾװ� ��ұݾ��� �ٸ���� ������/���ϸ���ȯ��/��ġ��ȯ�� �� �����մϴ�.');
            return false;
        }

        if (!confirm('ȯ�� �ݾ��� ��� �ݾװ� �ٸ��� �����ɰ�� ����ġ�� �ݾ��� �Էµ˴ϴ�.\n\n���� �Ͻðڽ��ϱ�?')) {
            return false;
        }
    }

    //returnmethod R400 �ΰ�� ��� ��Ҹ� ������
    if (frm.returnmethod.value == "R400") {
    	if (IsOrderFound && (IsThisMonthJumun != true)) {
	        alert('�޴��� ������ ��� ��Ҹ� �����մϴ�. �ٸ� ȯ�ҹ���� ������ �ּ���.');
	        frm.returnmethod.focus();
	        return;
    	}
    }

    return true;
}

// ============================================================================
// ȯ�� ���� üũ Form
// ============================================================================
function CheckReturnForm(frm) {
    if (!frm.returnmethod) { return true; }
    if (!frm.refundrequire) { return true; }

    if (frm.returnmethod.value.length < 1) {
        alert('ȯ�� ����� ������ �ּ���.');
        frm.returnmethod.focus();
        return false;
    }

	if (frm.returnmethod.value == "R000") {
		// ȯ�Ҿ��� �̸� üũ ���Ѵ�.
		frm.refundrequire.value = "0";
		return true;
	}

	if (frm.refundrequire.value*0 != 0) {
        alert('ȯ�� �ݾ��� �Է��ϼ���.');
        frm.refundrequire.focus();
        return false;
	}

	if ((frm.refundrequire.value*1 <= 0) && (frm.returnmethod.value != "R000")) {
		if (IsCsPowerUser != true) {
			alert('ȯ�� �ݾ��� 0 ���� Ŀ���մϴ�. �Ǵ� ȯ�Ҿ����� �����ϼ���.[0]');
			return false;
		} else if (frm.refundrequire.value*1 != 0) {
			alert('ȯ�� �ݾ��� 0 ���� Ŀ���մϴ�. �Ǵ� ȯ�Ҿ����� �����ϼ���.[1]');
			return false;
		}
	}

    if (((IsGiftingOrder == true) || (IsGiftiConOrder == true)) && (IsOrderCancelDisabled == true) && IsCSCancelProcess) {
        if (frm.returnmethod.value != "R910") {
        	// ��ġ�� ȯ���� ��� ������ ���� �ʿ����
	        if (!confirm('��ҺҰ� �ֹ��Դϴ�. : ' + OrderCancelDisableStr + ' \n\n ����Ͻðڽ��ϱ�? [�����ڱ��� �ʿ�]')) {
	            return false;
	        }

	        //���� �ִ��� check
	        if (IsCsPowerUser != true) {
	            alert('������ �����ϴ�. ��Ʈ��� ���� ���');
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
			        alert('�Ϻ���� ���Ŀ��� ����� �� ���� ȯ�ҹ���Դϴ�.(�ſ�ī�� �Ϻ���� ��)\n\n[�� �ְ����ݾ� : ' + MainPaymentOrg + ']. \n\n�ٸ� ȯ�Ҽ����� ������ �ּ���..');
			        frm.returnmethod.focus();
			        return false;
				}
    		} else {
	    		if (frm.refundrequire.value*1 != MainPaymentOrg) {
	    		    //alert(MainPaymentOrg);
	    		    //alert(frm.refundsubtotalprice.value*1);
			        alert('�Ϻ���� ���Ŀ��� ����� �� ���� ȯ�ҹ���Դϴ�.(�ſ�ī�� �Ϻ���� ��)\n\n[�� �ְ����ݾ� : ' + MainPaymentOrg + ']. \n\n�ٸ� ȯ�Ҽ����� ������ �ּ���..');
			        frm.returnmethod.focus();
			        return false;
				}
    		}
		}
    }

    //���̹� ���̴� �������������� ȯ�� �ؾ�.. �󷵸�. 2016/08/05
    if (pggubun=="NP"){
        if ((orgaccountdiv=="100")&&((frm.returnmethod.value!="R100")&&(frm.returnmethod.value!="R120"))){
            if(!confirm('���̹� �����ΰ�� ����������(ī��)���� ȯ���ϴ°��� �⺻�Դϴ�.\r\n������ ������������ ȯ�� ó�� ��� �Ͻðڽ��ϱ�?'))return;
        }

        if ((orgaccountdiv=="20")&&((frm.returnmethod.value!="R020")&&(frm.returnmethod.value!="R022"))){
            if(!confirm('���̹� �����ΰ�� ����������(�ǽð���ü)���� ȯ���ϴ°��� �⺻�Դϴ�.\r\n������ ������������ ȯ�� ó�� ��� �Ͻðڽ��ϱ�?'))return;
        }
    }

	// ====================================================================
	if (frm.returnmethod.value=="R007") {
		// ������
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
        	// ������ ���������� ���߿� ������ �Է��� �� �ִ�.
            if (!confirm('ȯ�� ���°� �����ϴ�. \n\nȯ�� ���� ���� ��� �Ͻðڽ��ϱ�?')) {
                if ((IsStatusRegister == true) || (IsStatusEdit == true)) {
                	frm.rebankaccount.focus();
                }
                return false;
            }
        }
	}

	if (frm.returnmethod.value == "R900") {
    	if (confirm("CS���񽺰� �ƴѰ��(�����ݾ�ȯ��) ���ϸ��� ��� ��ġ������ ȯ���ϼ���.\n\n���ϸ��� ȯ�� �Ͻðڽ��ϱ�?") != true) {
    		return false;
    	}
	}

	// ====================================================================
	if ((frm.returnmethod.value=="R900") || (frm.returnmethod.value=="R910")) {
		// ���ϸ���, ��ġ��ȯ��
        if ((frm.refund_userid) && (frm.refund_userid.value.length<1)) {
            alert('��ȸ������ �������� ���� ȯ�ҹ���Դϴ�. �ٸ� ȯ�� ����� �����ϼ���.');
            return false;
        }
	}

    return true;
}

// ============================================================================
// ����
// ============================================================================
function CsRegCancelProc(frm) {
    if (confirm('��ϵ� ���� ������ ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "deletecsas";
        frm.submit();
    }
}

function CsRegCancelFinishedProc(frm) {
	if (confirm('\n\n�Ϸ�CS����(���,��ǰ)�� �����Ͻðڽ��ϱ�?\n\n')) {
        frm.mode.value = "delfinishedcsas";
        frm.submit();
    }
}

// ============================================================================
// �Ϸ�ó��
// ============================================================================
function CsRegFinishProc(frm) {
	var btn = document.getElementById("btnFinishReturn");

	// ������ üũ
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	// ���� ��ǰ üũ �� ����
	if ((divcd != "A100") && (divcd != "A111") && (divcd != "A112")) {
	    if (!SaveCheckedItemList(frm)) {
	        return;
	    }
	} else {
	    if (!SaveChangeCheckedItemList(frm)) {
	        return;
	    }
	}

	// ���, ��ǰ, ȯ��
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true) || (divcd == "A003") || ((divcd == "A100") && frm.refundrequire && frm.refundrequire.value*1 !== 0)) {
        if (CheckReturnForm(frm) != true) {
            return;
        }

		if (frm.refundrequire) {
			if (frm.refundrequire_org.value*1 != frm.refundrequire.value*1) {
				alert("[�ý����� ����] ȯ�� ������ ����!!!!");
				return;
			}
		}
	}

	// ���, ��ǰ
    if ((IsCSCancelProcess) || (IsCSReturnProcess) || ((divcd == "A100") && frm.refundrequire && frm.refundrequire.value*1 !== 0)) {
		// ȯ�������� �������� üũ
		if (IsRefundInfoOK(frm) != true) {
			return;
		}

		// ȯ�Ҽ����� �ǹٸ���
		if (CheckReturnMethod(frm) != true) {
			return;
		}
    }

	if (IsStatusFinishing && (divcd == "A007" || ((divcd == "A003") && (frm.returnmethod.value=="R007")))) {
		if (IsAdminLogin) {
			alert('�̰����� �Ϸ�ó�� �Ͽ��� �ſ�ī�� �������/������ ȯ��ó���� �̷�� ���� �ʽ��ϴ�.[���α���]');
		} else {
			alert('�̰����� �Ϸ�ó�� �Ͽ��� �ſ�ī�� �������/������ ȯ��ó���� �̷�� ���� �ʽ��ϴ�.\n\n�Ϸ� ó�� �� �� �����ϴ�.');
			return;
		}
	}

    //ȯ�ҿ�û , �ſ�ī�� ��ҿ�û
    if ((divcd == "A003") || (divcd == "A007")) {
        if (frm.contents_finish.value.length<1){
            alert('ó�� ������ �Է��ϼ���.');
            frm.contents_finish.focus();
            return;
        }
    }

    if (IsChangeOrder && (OrderMasterState != "8")) {
        if (!confirm('��ȯȸ���Ϸ� ���� ��ǰ�Դϴ�.\n\n ����Ͻðڽ��ϱ�? [�����ڱ��� �ʿ�]')) {
            return;
        }

        //���� �ִ��� check
        if (IsCsPowerUser != true) {
            alert('������ �����ϴ�. ��Ʈ��� ���� ���');
            return;
        }
    }

    var confirmMsg ;
    confirmMsg = '�Ϸ�ó�� ���� �Ͻðڽ��ϱ�?';

    if ((divcd == "A004") || (divcd == "A010")) {
        confirmMsg = '�Ϸ�ó�� ����� ���̳ʽ� �ֹ� �� ȯ���� �ڵ� �����˴ϴ�. ���� �Ͻðڽ��ϱ�?';
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
// ��üó���Ϸ�=>���� ����
// ============================================================================
function CsUpcheConfirm2RegProc(frm) {
    if (confirm('���� ���·� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "upcheconfirm2jupsu";
        frm.submit();
    }
}

// ============================================================================
// ��ȯ�ֹ� �������
// ============================================================================
function CsChangeOrderRegProc(frm) {
    if (confirm('��ȯ�ֹ��� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "changeorderreg";
        frm.submit();
    }
}

// ============================================================================
// ����
// ============================================================================
function CsRegEditProc(frm) {

	// ������ üũ
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	// ���� ��ǰ üũ �� ����
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
				alert("�ݾ��� ���ڷ� �Է��ؾ� �մϴ�.");
				frm.customerrealbeasongpay.focus();
				return;
			}
		}
	}

	// ���, ��ǰ, ȯ��
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true) || (divcd == "A003")) {
        if (CheckReturnForm(frm) != true) {
            return;
        }
	}

	// ��ǰ��ۺ� üũ
	if (IsCSReturnProcess == true) {
		if (CheckRefundDeliverYpay(frm) != true) {
            return;
        }
	}

	// ���, ��ǰ
    if ((IsCSCancelProcess) || (IsCSReturnProcess)) {
		// ȯ�������� �������� üũ
		if (IsRefundInfoOK(frm) != true) {
			return;
		}

		// ȯ�Ҽ����� �ǹٸ���
		if (CheckReturnMethod(frm) != true) {
			return;
		}
    }

    //�߰� ���� ����
    if (frm.add_upchejungsandeliverypay){
        if (frm.add_upchejungsandeliverypay.value*0 != 0) {
            alert('���ڸ� �����մϴ�.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*1!=0){
            if (frm.add_upchejungsancause.value.length<1){
                alert('�߰� ���� ������ �Է��ϼ���.');
                frm.add_upchejungsancause.focus();
                return;
            }else if ((frm.add_upchejungsancause.value=='�����Է�')&&(frm.add_upchejungsancauseText.value.length<1)){
                alert('�߰� ���� ������ �Է��ϼ���.');
                frm.add_upchejungsancauseText.focus();
                return;
            }

            if (frm.buf_requiremakerid.value.length<1){
                alert('�߰� ������� �ִ°�� �귣�� ���̵� �����Ǿ�� �մϴ�. ');
                return;
            }
        }
    }

	if ($("#needChkYN_X").val()) {
		if ($("#needChkYN_X").prop("checked") === false && $("#needChkYN_F").prop("checked") === false) {
			alert('�Ϸᱸ���� �����ϼ���.');
			$("#needChkYN_X").focus();
			return;
		}
	}

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "editcsas";
        frm.submit();
    }
}

// ============================================================================
// ȯ�ҿ�û ���� �Ϸ� ó��
// ============================================================================
function CsRegFinishProcNoRefund(frm){
    var divcd = frm.divcd.value;

    if (confirm('ȯ�� �� ���̳ʽ� ��� ���� �Ϸ�ó�� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "norefund";
        frm.submit();
    }
}

// ============================================================================
// üũ�� ��ǰ/��ۺ� ���ٲٱ�
// ============================================================================
function AnCheckClickAll(frm) {
	if (frm.orderdetailidx) {
	    for(var i = 0; i < frm.orderdetailidx.length; i++) {
			AnCheckClick(frm.orderdetailidx[i]);
	    }
	}
}

// ============================================================================
// ���� ����ȸ�� �ΰ�
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
// ���ٹ����� �� ������ǰ�ΰ�
// ============================================================================
function SetForceReturnByCustomer(frm) {
	var e;

	if ((IsStatusRegister == true) || (divcd != "A004")) {
		return;
	}

	// requiremakerid �� ���̸� ����ȸ��, requiremakerid = 10x10logistics �̸� ���ٹ��� ����ǰ, ��Ÿ ��ü��ǰ
	if (frm.requiremakerid.value == "10x10logistics") {
		frm.ForceReturnByCustomer.checked = true;
	} else {
		frm.ForceReturnByCustomer.checked = false;
	}
}

//�ҷ���ǰ���
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

// Ƽ���ֹ��ΰ��..
function calcuTicketCancelCharge(comp){
    var frm = comp.form;
    var mayTicketCancelPro = comp.value*1;

    if (mayTicketCancelPro>0){
        alert('Ƽ�� ��� ������ ' + mayTicketCancelPro + '% ����' );
    }else{
        alert('Ƽ�� ��� ������ ���� ����' );
    }
    frm.refundadjustpay.value = (frm.refundtotalbuypaysum.value*mayTicketCancelPro/100)*-1;

    CheckForItemChanged();
}

// ============================================================================
// ���� ��ǰ ���� �ٹ� ��ǰ ����
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
// �� ������ǰ ��ȯ(A010 -> A004)
// ============================================================================
function ChangeDivcdToA004(frm) {
	if (IsDeletedCS) {
		alert("������ �����Դϴ�.");
		return;
	}

	if (IsLogicsSended) {
		alert("�̹� �ù�翡 ���۵� �����Դϴ�.");
		return;
	}

	if (IsStatusEdit != true) {
		alert("������ �� �����ϴ�.");
		return;
	}

    if (confirm('�� ������ǰ���� ��ȯ �Ͻðڽ��ϱ�?')){
        frm.mode.value = "changedivcdtoa004";
        frm.submit();
    }
}

// ============================================================================
// ȸ����û ��ȯ(A004 -> A010)
// ============================================================================
function ChangeDivcdToA010(frm) {
	if (IsDeletedCS) {
		alert("������ �����Դϴ�.");
		return;
	}

	if (IsLogicsSended) {
		alert("�̹� �ù�翡 ���۵� �����Դϴ�.");
		return;
	}

	if (IsStatusEdit != true) {
		alert("������ �� �����ϴ�.");
		return;
	}

	if (frm.buf_requiremakerid) {
		if (frm.buf_requiremakerid.value != "10x10logistics") {
			alert("��ü��� ��ǰ�Դϴ�. ������ �� �����ϴ�.");
			return;
		}
	} else {
		alert("�߸��� �����Դϴ�.");
		return;
	}

	if (IsAllTenbaeItem(frm) != true) {
		alert("��ü��� ��ǰ�Դϴ�. ������ �� �����ϴ�.");
		return;
	}

    if (confirm('ȸ����û���� ��ȯ �Ͻðڽ��ϱ�?')){
        frm.mode.value = "changedivcdtoa010";
        frm.submit();
    }
}

// ============================================================================
// ����â ������ ����
// ============================================================================
function resizeTextArea(textarea, textareawidth) {
	if (!textarea) { return; }
	var lines = textarea.value.split('\n');
	var textareaminheight = textarea.rows;

	var textareaheight = 1;
	for (x = 0; x < lines.length; x++) {
		// textareawidth �Ѿ�� ���� �߰�
		c = lines[x].length;
		if (c >= textareawidth) {
			textareaheight += Math.ceil(c / textareawidth);
		}
	}
	textareaheight += lines.length;

	for (x = (lines.length - 1); x >= 0; x--) {
		// �ǹ��� ���Ͷ����� ����
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

// ��ۺ����
function CsRegCancelBeasongPayProc(frm, orderdetailidx) {
	if (CheckedProductExist(frm, orderdetailidx) == true) {
		alert("��ۺ���� �Ұ�\n\nüũ�� ��ǰ�� �ִ� ��� ��ۺ� ��������� �� �����ϴ�.");
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
