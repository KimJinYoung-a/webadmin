


// ============================================================================
// CS �������� ǥ�� (AJAX)
// ============================================================================
function divCsAsGubunSelect(value_gubun01,value_gubun02,name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv){
	if (IsFinishProcState) {
	    alert('����â���� �������ּ���. - �Ϸ�ó���� �����Ұ�');
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

function delGubun(name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv){
    eval("document." + name_frm + "." + name_gubun01).value = "";
    eval("document." + name_frm + "." + name_gubun02).value = "";
    eval("document." + name_frm + "." + name_gubun01name).value = "";
    eval("document." + name_frm + "." + name_gubun02name).value = "";

    eval(targetDiv).innerHTML = "";
}


function selectGubun(value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_gubun01, name_gubun02, name_gubun01name, name_gubun02name ,name_frm, targetDiv){
	if (IsFinishProcState) {
	    alert('����â���� �������ּ���. - �Ϸ�ó���� �����Ұ�');
	    return;
	}

    var frm = eval(name_frm);

    eval("document." + name_frm + "." + name_gubun01).value = value_gubun01;
    eval("document." + name_frm + "." + name_gubun02).value = value_gubun02;
    eval("document." + name_frm + "." + name_gubun01name).value = value_gubun01name;
    eval("document." + name_frm + "." + name_gubun02name).value = value_gubun02name;

    eval(targetDiv).innerHTML = "";

    // ��ü ���������� �Է��� ��� ���õ� Detail�� ���� ���� ����
    if (targetDiv=="causepop"){
        for (var i=0;i<frm.elements.length;i++){
            var e = frm.elements[i];

            if ((e.type=="checkbox")&&(e.checked)&&(e.name=="orderdetailidx")){
                setDetailCause(e.value, value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_frm);
            }
        }
    }

    // ��ü������� üũ
    CheckUpcheDeliverPay(frm);

    // ��ۺ� üũ
    CheckDeliverPay(frmaction);

	// �ݾ� ����
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
// ������ ���ΰ�ħ
// ============================================================================
function reloadMe(comp){
    var divcd = comp.value;
    var mode  = Fmode;
    var orderserial = Forderserial;
    document.location = "?mode=" + mode + "&divcd=" + divcd + "&orderserial=" + orderserial;
}



// ============================================================================
// ������ ��ǰ �Է¼��� üũ
// ============================================================================
function CheckMaxItemNo(obj, maxno) {
    if (obj.value*1 > maxno*1) {
        alert("�ֹ����� �̻����� ��ǰ������ �����Ҽ� �����ϴ�.");
        obj.value = maxno;
    }

	if (IsEditState && IsReturnProcess) {
	    if (obj.value < 0) {
	        alert("0�� �̸� ����  �����Ҽ� �����ϴ�.");
	        obj.value = maxno;
	    }
	} else {
	    if (obj.value < 1) {
	        alert("0�� ���Ϸ�  �����Ҽ� �����ϴ�. ��ǰ������ �������ּ���.");
	        obj.value = maxno;
	    }
	}
}

// ��ǰ���ý� Ȯ���� �͵�
function CheckSelect(comp){
    var chkidx = comp.value;
    var frm = document.frmaction;

    if (comp.name!="Deliverdetailidx"){
        if (comp.checked){
            // CS �������� ����
            eval("frm.gubun01_" + chkidx).value = frm.gubun01.value;
            eval("frm.gubun02_" + chkidx).value = frm.gubun02.value;
            eval("frm.gubun01name_" + chkidx).value = frm.gubun01name.value;
            eval("frm.gubun02name_" + chkidx).value = frm.gubun02name.value;
        }else{
            delGubun("gubun01_" + chkidx,"gubun02_" + chkidx,"gubun01name_" + chkidx,"gubun02name_" + chkidx,frm.name,causepop);
        }
    }

    // ��ü������� üũ
    CheckUpcheDeliverPay(frm);

    // ��ۺ� üũ
    CheckDeliverPay(frm);

	// �ݾ� ����
    CalculateAndApplyItemCostSum(frm);

    // ���� ��ǰ�� ���� �귣�� �˻�
    DispCheckedUpcheID(frmaction);
}



// ============================================================================
// ��ü��ۺ� �հ�
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

// ��ü ��ۺ� üũ
// �ٹ�   - �ٹ��̰�, �귣�� ��ü��ǰ�� üũ�Ȱ�� �귣�� ���� ��ۺ� ���� üũ�Ѵ�.
// �귣�� - �����̰�, �귣�� ��ü��ǰ�� üũ�Ȱ�� �귣�� ��ۺ� üũ�Ѵ�.
// �ܼ������ΰ�� üũ���� �ʴ´�.
function CheckUpcheDeliverPay(frm){
    var upbeaMakerid;
    var itemMakerid;

    var NotCheckExists;				// üũ�ȵ� ��ǰ�� �ִ°�
    var isCheckValExists;			// �Ѱ��� ���õ� ��ǰ�� �ִ°�

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
            	// ��ۺ� �귣�尡 ������ : ���ٹ��
                for (var j = 0; j < frm.orderdetailidx.length; j++) {
                    objitem = frm.orderdetailidx[i];

                    if (objitem.type == "checkbox") {
	                    if ((frm.odlvtype[j].value == "1") || (frm.odlvtype[j].value == "4")) {
	                        // ���ٹ���̸�
	                        isCheckValExists = true;
	                        NotCheckExists = (NotCheckExists) || (!frm.orderdetailidx[j].checked);
	                        NotCheckExists = (NotCheckExists) || ((frm.orderdetailidx[j].checked == true) && (frm.itemno[j].value != frm.regitemno[j].value));
	                    }
                    }
                }
            } else {
            	// ��ۺ� �귣�尡 ������ : ��ü���
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

            // ���ΰ� ���õ� ��� ��ۺ� ���� üũ�Ѵ�.(�귣�� ����ǰ ���)
            frm.Deliverdetailidx[i].checked = ((!NotCheckExists) && (isCheckValExists));

            //��ǰ ���μ����� �ܼ������� ��� üũ�������� �ʴ´�.
            if ((IsReturnProcess) && ((value_gubun02 == "CD01")||(value_gubun02 == "CD06"))) {
                frm.Deliverdetailidx[i].checked = false;
            }
            AnCheckClick(frm.Deliverdetailidx[i]);
        }
    }
}

//��ۺ� üũ
function CheckDeliverPay(frm){
    var allselected = IsAllSelected(frm);
    var brandallselected = IsBrandAllSelected(frm);

    var value_gubun02 = frm.gubun02.value;

    if (IsCancelProcess) {
        // ��� Process

        if (allselected) {
            // ���ϸ���, ���� ȯ��
            frm.milereturn.checked = true;
            frm.couponreturn.checked = true;
        } else {
            frm.milereturn.checked = false;
            frm.couponreturn.checked = false;
        }

    }else if (IsReturnProcess) {
    	// ��ǰ Process

        // ȸ����ۺ� ����
        if ((value_gubun02=="CD01")||(value_gubun02=="CD06")){
        	// �ܼ�����
            if (frm.divcd.value=="A010"){
            	// ȸ����û(�ٹ����ٹ��)
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
        //������ ȯ��
        document.all.refundinfo_R007.style.display = "inline";
    }else if((comp.value=="R020")||(comp.value=="R080")||(comp.value=="R100")||(comp.value=="R400")){
        //�ǽð� ��ü ���//ALL@ ���� ��� //�ſ�ī�� ���� ���//�޴���
        document.all.refundinfo_R100.style.display = "inline";
    }else if(comp.value=="R050"){
        //������ ���� ���
        document.all.refundinfo_R050.style.display = "inline";
    }else if(comp.value=="R900"){
        //���ϸ��� ȯ��
        document.all.refundinfo_R900.style.display = "inline";
    }

}

//CS ����
function CsRegProc(frm){

    if (frm.divcd.value.length<1){
        alert("���� ������ �����ϼ���.");
        frm.divcd.focus();
        return;
    }

    if (frm.title.value.length<1) {
        alert("������ �Է��ϼ���.");
        frm.title.focus();
        return;
    }

    if (frm.gubun01.value.length<1) {
        alert("���� ������ �Է��ϼ���.");
        return;
    }

    //���� ��ǰ üũ
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

				// ȯ�� ������ �����ϴ�.
				if (frm.returnmethod.value != "R000") {
		            if (frm.subtotalprice.value*1 < frm.canceltotal.value*1) {
		                alert('ȯ�� �������� �����ݾ� ���� Ŭ �� �����ϴ�. - �����̳� ���ϸ��� ȯ���� üũ�� �ּ���.');

		                if (IsAdminLogin != true) {
		                	return;
		                }
		            }

		            if (frm.canceltotal.value*1<1){
		                alert('ȯ�ҿ������� �����ϴ�. ȯ�Ҿ����� �����ϼ���.');

		                if (IsAdminLogin != true) {
		                	return;
		                }
		            }
				}
            }

            if (frm.nextsubtotal.value*1<0){
                alert('��� �� ���� �ݾ��� ���̳ʽ��� �� �� �����ϴ�. - �����̳� ���ϸ��� ȯ���� üũ�� �ּ���.');

                if (IsAdminLogin != true) {
                	return;
                }
            }

            if (frm.canceltotal.value*1<0){
                alert('�� ��� �ݾ��� 0���� ���� �� �����ϴ�. - �����̳� ���ϸ��� ȯ���� UnCheck �ּ���.');

                if (IsAdminLogin != true) {
                	return;
                }
            }

            //returnmethod R400 �ΰ�� ��� ��Ҹ� ������
            if (frm.returnmethod){
	            if (frm.returnmethod.value=="R400"){
			    	if (IsOrderFound && (IsThisMonthJumun != true)) {
				        alert('�޴��� ������ ��� ��Ҹ� �����մϴ�. �ٸ� ȯ�ҹ���� ������ �ּ���.');
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

			if (Fdivcd == "A700") {
	            alert('�߰� ������� �Է��ϼ���.');
	            frm.add_upchejungsandeliverypay.focus();
	            return;
			}
        }
    }

    if (IsCancelProcess){
        if(confirm("��� ���� �Ͻðڽ��ϱ�?")){
            frm.submit();
        }
    }else if (IsReturnProcess){
        if (frm.ForceReturnByTen.checked){
            frm.requireupche.value = "N";
            frm.requiremakerid.value = "";
        }

        if (frm.requireupche.value=="Y"){
            if(confirm("��ü [" + frm.requiremakerid.value +"]�� ��ǰ/ȸ��/��ȯ ���� �Ͻðڽ��ϱ�?")){
                frm.submit();
            }
        }else{
            if(confirm("[�ٹ����� ��������]�� ��ǰ/ȸ��/��ȯ ���� �Ͻðڽ��ϱ�?")){
                frm.submit();
            }
        }
    }else if (IsRefundProcess){
        if (frm.returnmethod.value.length<1){
            alert('ȯ�� ����� ������ �ּ���.');
            frm.returnmethod.focus();
            return;
        }

        if (frm.returnmethod.value=="R000"){
            frm.refundrequire.value = "0";

        }


        if ((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R080")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R400")){
            alert('�ſ�ī��/�ǽð�/�ÿ�/�޴��� ȯ�� ������ �� ���� �ݾ� �״�� �����˴ϴ�.');
            frm.refundrequire.value = frm.orgsubtotalprice.value;
        }

        if ((frm.refundrequire.value.length<1)||(!IsDigit(frm.refundrequire.value))){
            alert('ȯ�� �ݾ��� �Է��ϼ���.');
            frm.refundrequire.focus();
            return;
        }

        if(confirm("ȯ�� ���� �Ͻðڽ��ϱ�?")){
            frm.submit();
        }
    }else if(confirm("���� �Ͻðڽ��ϱ�?")){
        frm.submit();
    }
}

//��üó���Ϸ�=>���� ����
function CsUpcheConfirm2RegProc(frm){
    if (confirm('���� ���·� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "upcheconfirm2jupsu";
        frm.submit();
    }
}

//����
function CsRegEditProc(frm){
    if (frm.divcd.value.length<1){
        alert("���� ������ �����ϼ���.");
        frm.divcd.focus();
        return;
    }

    if (frm.title.value.length<1) {
        alert("������ �Է��ϼ���.");
        frm.title.focus();
        return;
    }

    if (frm.gubun01.value.length<1) {
        alert("���� ������ �Է��ϼ���.");
        return;
    }

	//��ǰ ���� �����ϰ�� üũ��
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

    //�߰� ���� ����
    if (frm.add_upchejungsandeliverypay){
        if (!IsDigit(frm.add_upchejungsandeliverypay.value)){
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

    //returnmethod R400 �ΰ�� ��� ��Ҹ� ������
    if (frm.returnmethod){
	    if (frm.returnmethod.value=="R400"){

	    	if (IsOrderFound && (IsThisMonthJumun != true)) {
		        alert('�޴��� ������ ��� ��Ҹ� �����մϴ�. �ٸ� ȯ�ҹ���� ������ �ּ���.');
		        frm.returnmethod.focus();
		        return;
	    	}
	    }
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "editcsas";
        frm.submit();
    }
}

// ��ü �߰����� ���� �귣�� ���̵� ��������
function DispCheckedUpcheID(frm) {
    var checkedUpcheid = "";
    var UpcheDuplicated = false;
    var IsUpcheReturn;

	if ((Fdivcd == "A004") || (Fdivcd == "A000")) { // ��ǰ����(��ü���), �±�ȯ���
		IsUpcheReturn = true;
	} else {
		IsUpcheReturn = false;
	}

    if (!frm.buf_requiremakerid) {
        return;
    }

	// ���õ� �������߿���
	//  - �귣�� ��������, ���� �ٸ� �ΰ��� �귣�尡 ������ �ߺ� ǥ��(������ �ϳ��� �귣�� ��ǰ���θ� �ؾ��Ѵ�.)
	//
	//  - ��ǰ����(��ü���), �±�ȯ��� �̰� ��ü����̸� ��������
	//  - �̿��� ��� �귣�� ��������
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


//ȸ������ üũ
function checkReturnProcessAvail(frm){
    var TenBeasongExists = false;
    var UpcheBeasongExists = false;
    var UpcheDuplicated = false;
    var checkedUpcheid = "";

	// ���� �߼۽� �귣�带 �����ϰ�..=> ��ǰ�� �����ϴ°��� �ƴϹǷ�.
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
        alert('�ٹ����� ��۰� ��ü����� ���ÿ� ���� �Ͻ� �� �����ϴ�.');
        return false;
    }

    if (UpcheDuplicated){
        alert('��ü��� ���ý� �귣�� ���� ������ �ּ���.');
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

//ȯ�� ���� üũ form
function CheckReturnForm(frm){
    if (frm.returnmethod){
	    if (frm.returnmethod.value=="R007"){
	        var mooconfirm = false;
	        if ((frm.rebankaccount!=undefined)&&(frm.rebankaccount.value.length<1)){
	            //alert('ȯ�� ���¸� �Է��� �ּ���.');
	            //frm.rebankaccount.focus();
	            mooconfirm=true;
	            //return false;
	        }

	        if ((frm.rebankownername!=undefined)&&(frm.rebankownername.value.length<1)){
	            //alert('�����ָ���  �Է��� �ּ���.');
	            //frm.rebankownername.focus();
	            mooconfirm=true;
	            //return false;
	        }

	        if ((frm.rebankname!=undefined)&&(frm.rebankname.value.length<1)){
	            //alert('ȯ�� ������ ������ �ּ���.');
	            //frm.rebankname.focus();
	            mooconfirm=true;
	            //return false;
	        }

	        if ((frm.refundrequire!=undefined)&&(frm.refundrequire.value.length<1)){
	            alert('ȯ�� �ݾ��� �� ����ϼ���');
	            return false;
	        }

	        if (mooconfirm){
	            if (!confirm('ȯ�� ���°� �����ϴ�. \n\nȯ�� ���� ���� ��� �Ͻðڽ��ϱ�?')){
	                frm.rebankaccount.focus();
	                return false;
	            }
	        }

    }else if (frm.returnmethod.value=="R900"){
	        if ((frm.refundbymile_userid!=undefined)&&(frm.refundbymile_userid.value.length<1)){
	            alert('����� ���̵� �����ϴ�. �ٸ� ȯ�� ����� �����ϼ���.');
	            return false;
	        }

	        if ((frm.refundbymile_sum!=undefined)&&(frm.refundbymile_sum.value.length<1)){
	            alert('ȯ�� ���ϸ����� �����ϴ�. ���� ���ּ���.');
	            return false;
	        }

	    }

	    if ((frm.returnmethod.value!="")&&(frm.returnmethod.value!="R000")&&(!IsDigit(frm.refundrequire.value))){
	        alert('ȯ�� �ݾ��� ���(+) �� �����մϴ�.');
	        return false;
	    }
    }
    return true;
}

//ȯ�� ���� üũ
function CheckReturnMethod(frm){
    var allselected = IsAllSelected(frm);

    var PayedNCancelEqual = (frm.subtotalprice.value*1==frm.canceltotal.value*1);

    if ((allselected)&&(!PayedNCancelEqual)&&(IsCancelProcess)){
        alert('��ü ����ΰ�� �����ݾ� ��ü�� ȯ���ؾ��մϴ�. - ����ۺ� ȯ��, ���ϸ���, ���α� ���� üũ���ּ���.');
        return false;
    }

    if (((!PayedNCancelEqual))&&((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R080")||(frm.returnmethod.value=="R400"))){
        alert('��ü ����� ��츸 �ſ�ī��/�ǽð� ��ü/�޴��� ��Ұ� �����մϴ�. \n\n������ ȯ�� �Ǵ� ���ϸ��� ȯ���� ������ �ּ���');
        frm.returnmethod.focus();
        return false;
    }

    if (frm.returnmethod.value.length<1){
        alert('ȯ�� ����� ������ �ּ���.');
        frm.returnmethod.focus();
        return false;
    }

    if (!CheckReturnForm(frm)){
        return false;
    }

	if (FSiteName != "10x10") {
	    //�ܺθ��� ��� �ܺθ� ȯ�������� ����..
	    if ((frm.returnmethod.value != "R050") && (frm.returnmethod.value != "R000")) {
	        alert('�ܺθ��� ��� ȯ�� ���� �Ǵ� �ܺθ� ȯ���� �����ϼ���. \n\n���� ����ڸ� ���� ���޸����� ��� ȯ�� ó�� �մϴ�.');
	        frm.returnmethod.focus();
	        return
	    }
	}

    if (frm.refundrequire.value!=frm.canceltotal.value){
        if ((frm.returnmethod.value!="R007")&&(frm.returnmethod.value!="R900")&&(frm.returnmethod.value!="R000")){
            alert('ȯ�� �ݾװ� ��ұݾ��� �ٸ���� ������ �Ǵ� ���ϸ��� ȯ�Ҹ� �����մϴ�.');
            return false;
        }

        if (!confirm('ȯ�� �ݾ��� ��� �ݾװ� �ٸ��� �����ɰ�� ����ġ�� �ݾ��� �Էµ˴ϴ�.\n\n���� �Ͻðڽ��ϱ�?')){
            return false;
        }
    }

    return true;
}


//����
function CsRegCancelProc(frm){
    if (confirm('��ϵ� ���� ������ ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "deletecsas";
        frm.submit();
    }
}

//�������·� ����
function CsRegStateChg(frm){
    if (confirm('���� ���·� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "state2jupsu";
        frm.submit();
    }
}

//�Ϸ�ó��
function CsRegFinishProc(frm){
    var divcd = frm.divcd.value;

    //ȯ�ҿ�û , �ſ�ī�� ��ҿ�û
    if ((divcd=="A003")||(divcd=="A007")){
        if (frm.contents_finish.value.length<1){
            alert('ó�� ������ �Է��ϼ���.');
            frm.contents_finish.focus();
            return;
        }
    }

    var confirmMsg ;
    confirmMsg = '�Ϸ�ó�� ���� �Ͻðڽ��ϱ�?';

    if ((divcd=="A004")||(divcd=="A010")){
        confirmMsg = '�Ϸ�ó�� ����� ���̳ʽ� �ֹ� �� ȯ���� �ڵ� �����˴ϴ�. ���� �Ͻðڽ��ϱ�?';
    }

    //20090601�߰�
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

    if (confirm('ȯ�� �� ���̳ʽ� ��� ���� �Ϸ�ó�� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "norefund";
        frm.submit();
    }
}

//��ǰ ��ü üũ �Ǿ�����
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

// �귣�� ��ǰ ��ü üũ �Ǿ�����
// ����� �ش� �귣�� ��ü
// �ٹ��� ��� �ٹ� ��ü
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

// �Ѱ� �귣�� ��ü ���õǾ����� üũ
// �귣�尡 ���̸� ���ٹ�� ��üüũ
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
	                alert('������ �Է��ϼ���.');
	                e.focus();
	                e.select();
	                return false;
	            }
			} else {
	            if ((e.value*1)==0){
	                alert('������ �Է��ϼ���.');
	                e.focus();
	                e.select();
	                return false;
	            }
			}

            regitemno = e.value;
        }

        // ������.
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

    //��ۺ� ----------------------------------------
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

    //��Ÿ����, ���񽺹߼� , ȯ�ҿ�û, �������ǻ���, ��ü �߰� ���� - �󼼳��� üũ ����.
    if ((Fdivcd=="A009")||(Fdivcd=="A002")||(Fdivcd=="A003")||(Fdivcd=="A005")||(Fdivcd=="A006")||(Fdivcd=="A700")){
        // no- check

    }else{
        if (!checkitemExists){
            alert('���õ� �󼼳����� �����ϴ�.');
            return false;
        }
    }

    return true;
}

//��ü �߰� ���� ������
function clearAddUpchejungsan(frm){
    frm.add_upchejungsandeliverypay.value = "0";
    frm.add_upchejungsancause.value = "";
    frm.add_upchejungsancauseText.value = "";
}

function setTitle(frm,titlestr){
    frm.title.value=titlestr;
}

// �� ��� ��� ����. ��� ���� ����.
function getReCalcuBeasongPay(userlevel, itemsum){
    //Yellow : 5�����̻�, �׸�ȸ�� : 4�����̻�, ���ȸ�� : 3�����̻�, vipȸ�� : �׻� ���� , ���Ͼ� 2�����̻�

    //�������ΰ�� üũ. �߰� �ؾ���

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

    //��ü��ۺ� ȯ�� �߰�
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

    // ���� ��ǰ �հ� �ݾ�
    if (frm.orgitemcostsum!=undefined){
        orgitemcostsum = frm.orgitemcostsum.value*1;
    }

    // ��ǰ��� �ϴ� ���� �հ�
    if (frm.itemcanceltotal!=undefined){
         frm.itemcanceltotal.value = refunditemcostsum;
    }

	// ��� �� ���� �ÿ��� ����.
	if (IsRegisterState || IsRefundInfoFound) {
	    // ���/��ǰ ��ǰ�Ѿ�
	    if (frm.refunditemcostsum!=undefined){
	        frm.refunditemcostsum.value = refunditemcostsum;
	    }

	    // ���þ���(������) ��ǰ�Ѿ�
	    if (frm.remainitemcostsum!=undefined){
	        frm.remainitemcostsum.value = orgitemcostsum - refunditemcostsum;

	        remainitemcostsum = frm.remainitemcostsum.value;
	    }

	    // ��� ���ϸ��� ȯ��
	    if (frm.milereturn!=undefined){
	        if (frm.milereturn.checked){

				// ȯ�����ϸ����� ��ұݾ׺��� ū���
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

	    // ��� ���α� ȯ��
	    if (frm.couponreturn!=undefined){
	        if (frm.couponreturn.checked){
	            frm.refundcouponsum.value = frm.tencardspend.value*-1;
	        }else{
	            frm.refundcouponsum.value = 0;

	            // % ���α� ����.
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

	    // ī�� ���� ����
	    if (frm.allatsubtractsum){
            frm.allatsubtractsum.value = 0;

            // ī�� ���� ���� 200906�߰�
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

	    // ��Ÿ�����ݾ�
	    if (frm.refundadjustpay!=undefined){
	        refundadjustpay = frm.refundadjustpay.value*1;
	    }

	    //�� ��ۺ�
	    if (frm.orgbeasongpay!=undefined){
	        orgbeasongpay  = frm.orgbeasongpay.value*1;
	    }

	    // ������μ����� �� ��ۺ� ó��
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

	    // ��ǰ���μ����� �� ��ۺ� ó��
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

	    //��� ��ۺ� = refundbeasongpay - recalcubeasongpay
	    if (frm.refundbeasongpay!=undefined){
	        frm.refundbeasongpay.value = refundbeasongpay
	    }

	    // ȸ�� ��ۺ� -
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

	    //��ұݾ� �հ�
	    if (frm.canceltotal!=undefined){
	        frm.canceltotal.value  = refunditemcostsum + refundmileagesum*1 + refundcouponsum*1 + allatsubtractsum*1 + refundbeasongpay*1 + refundadjustpay*1 + refunddeliverypay*1;
	    }

	    //����� �ݾ� �հ�
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

//�߰������ۺ�
function Change_add_upchejungsandeliverypay(comp){
    comp.form.buf_totupchejungsandeliverypay.value = comp.form.buf_refunddeliverypay.value*1 + comp.value*1;

    if (isNaN(comp.form.buf_totupchejungsandeliverypay.value)){
        comp.form.buf_totupchejungsandeliverypay.value = "0";
    }
}

//�߰������ۺ� ����
function Change_add_upchejungsancause(comp){
    if (comp.value=="�����Է�") {
        document.all.span_add_upchejungsancauseText.style.display = "inline";
    }else{
        document.all.span_add_upchejungsancauseText.style.display = "none";
    }
}

// ������� �̷�
function popDeliveryTrace(traceUrl, songjangNo)
{
	var f = document.popForm;
	f.traceUrl.value	= traceUrl;
	f.songjangNo.value	= songjangNo;
	f.submit();
}