//���� cs ���� js
//2012.03.20 �ѿ�� ����

function searchDetail(csmasteridx){
    location.href='/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?csmasteridx='+csmasteridx;
}

//��SMS �߼�
function PopCSSMSSend_off(reqhp,masteridx,orderno,defaultMsg){
    var window_width = 250;
    var window_height = 210;
    var popwin=window.open("/admin/offshop/shopcscenter/action/pop_cs_sms_send.asp?orderno=" + orderno + "&reqhp=" + reqhp + "&masteridx=" + masteridx + "&defaultMsg=" + defaultMsg,"PopCSSMSSend_off","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + "  scrollbars=no resizable=no");
    popwin.focus();
}

function GetCenterX(window_width) {
	return (screen.width - window_width)/2;
}

function GetCenterY(window_height) {
	return (screen.height - window_height)/2;
}

// a/s
function PopOpenServiceItemas(masteridx){
    var mode    = "regcsas";
    var divcd   = "A030";

	if (masteridx == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }
		
    PopCSActionCom(masteridx,'',mode,divcd,'','');
}

// CSó�� ���/����
function PopCSActionFinish(masteridx,mode,csmasteridx){
    PopCSActionCom(masteridx,'',mode,'','',csmasteridx);
}

// CSó�� ���/����
function PopCSActionEdit(masteridx,mode,csmasteridx){
    PopCSActionCom(masteridx,'',mode,'','',csmasteridx);
}

// CSó�� ���/���� ����
function PopCSActionCom(masteridx,orderno,mode,divcd,ckAll,csmasteridx){
    var popwin=window.open("/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?orderno=" + orderno + "&masteridx=" + masteridx + "&mode=" + mode + "&divcd=" + divcd + "&ckAll=" + ckAll+ "&csmasteridx=" + csmasteridx,"pop_cs_action_reg_" + divcd,"width=1024 height=768 scrollbars=yes resizable=yes");
    popwin.focus();
}

function changeSongjang(csid){
    var popwin = window.open('/admin/offshop/shopcscenter/action/popChangeSongjang.asp?masteridx=' + csid,'popChangeSongjang','width=400,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//CS ����
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

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "editcsas";
        frm.submit();
    }
}

// ��ǰ���ý� Ȯ���� �͵�
function CheckSelect(comp){
    var chkidx = comp.value;
    var frm = document.frmaction;

    // ���� �귣�常 ���� �����ϰ�
    // ��ǰ ��� �귣�� ����
    DisableUpcheDeliver(frm);

	// ���� ��ǰ/�±�ȯ�� ��� ��� �귣�� üũ
	//DispCheckedUpcheID(frm);
}

// ������ ��ǰ �Է¼��� üũ
function CheckMaxItemNo(obj, maxno) {
    if (obj.value*1 > maxno*1) {
        alert("�ֹ����� �̻����� ��ǰ������ �����Ҽ� �����ϴ�.");
        obj.value = maxno;
    }
}

// ��ü �߰����� ���� �귣�� ���̵� ��������
function DispCheckedUpcheID(frm) {
    var checkedUpcheid = "";
    var UpcheDuplicated = false;
    var IsUpcheReturn;

	if ((divcd == "A004") || (divcd == "A000")) { // ��ǰ����(��ü���), �±�ȯ���
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

	if (UpcheDuplicated) {
		alert("�ΰ��� �귣�尡 ���ÿ� ���õǾ� �ֽ��ϴ�.(�ߺ��Ұ�) ������ �����ϼ���.");
	}

    if ((!UpcheDuplicated)&&(checkedUpcheid!="")){
        if (frm.buf_requiremakerid){
            frm.buf_requiremakerid.value = checkedUpcheid;
        }
    }
}

// ============================================================================
// ��ǰ�� ���� �귣�常 �����Ҽ� �ֵ��� ����
// ��ǰ ��� �귣�� ����
// �ٹ� üũ�� - ���� Disable
// ���� üũ�� - �ٹ� �� �ٸ� �귣�� Disable
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
        		// �ٹ�
	        	upbeaMakerid = "";
        	} else {
        		// ����
	        	upbeaMakerid = frm.makerid[i].value;
        	}
        	checkfound = true;
        	break;
        }
    }

	// ��ǰ ��� �귣�� ����
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
		    // TR �� bgColor �� ���Ѵ�.(���� FFFFFF �ΰ͸� Ȱ��ȭ �� �� �ִ�.)
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

//cs����
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

    if (frm.reqname.value.length<1) {
        alert("������ �̸��� �Է��ϼ���.");
        frm.reqname.focus();
        return;
    }

    if (frm.reqhp.value.length<1) {
        alert("�޴��� ��ȣ�� �Է��ϼ���.");
        frm.reqhp.focus();
        return;
    }

    if (frm.reqzipcode.value.length<1) {
        alert("�����ȣ�� �Է��ϼ���.");
        frm.reqzipcode.focus();
        return;
    }    

    if (frm.reqzipaddr.value.length<1) {
        alert("�ּҸ� �Է��ϼ���");
        frm.reqzipaddr.focus();
        return;
    }    

    if (frm.reqzipcode.value.length<1) {
        alert("�� �ּҸ� �Է��ϼ���.");
        frm.reqzipcode.focus();
        return;
    }    
    
	//�ֹ����
	if(divcd =='A008'){
		//���Ϸᰡ �ƴҰ��
		if (IsOrderMasterState != '8'){
			if (frm.cancelorderno.value.length<1) {
			    alert("����� �ֹ���ȣ�� �Է����ּ���(���̳ʽ��ֹ�)");
			    frm.cancelorderno.focus();
			    return;
			}
		}
	}
    
    //���� ��ǰ üũ
    if (!SaveCheckedItemList(frm)) {
		return;
    }
         
    if(confirm("���� �Ͻðڽ��ϱ�?")){
        frm.submit();
    }
}

// ���û�ǰ ����
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

        if (e.name == "dummystopper") {
            if (ischecked == true) {
                frm.detailitemlist.value = frm.detailitemlist.value + "|" + orderdetailidx + "\t" + regitemno + "\t" + causecontent + "\t" + frm.detailitemlist.value;
                ischecked = false;
                regitemno = "";
                causecontent = "";
            }
        }
    }

    //��Ÿ����, ���񽺹߼� , ȯ�ҿ�û, �������ǻ���, ��ü �߰� ���� - �󼼳��� üũ ����.
    if ((divcd=="A009")||(divcd=="A002")||(divcd=="A003")||(divcd=="A005")||(divcd=="A006")||(divcd=="A700")){
        // no- check

    }else{
        if (!checkitemExists){
            alert('���õ� �󼼳����� �����ϴ�.');
            return false;
        }
    }

    return true;    
}

// ��üó���Ϸ�=>���� ����
function CsUpcheConfirm2RegProc(frm){
    if (confirm('���� ���·� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "upcheconfirm2jupsu";
        frm.submit();
    }
}

//cs�Ϸ�ó��
function CsRegFinishProc(frm){
    var divcd = frm.divcd.value;

    var confirmMsg ;
    confirmMsg = '�Ϸ�ó�� ���� �Ͻðڽ��ϱ�?';

    if (confirm(confirmMsg )){

        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "";
        frm.submit();
    }
}

//���� cs ����Ʈ
function Cscenter_Action_List_off(masteridx,orderno, divcd,currstate,shopid) {
    var window_width = 1024;
    var window_height = 768;
	var Cscenter_Action_List_off = window.open("/admin/offshop/shopcscenter/action/cs_action.asp?masteridx=" + masteridx + "&orderno=" + orderno + "&divcd=" + divcd+"&currstate="+currstate+"&shopid="+shopid,"Cscenter_Action_List_off","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	Cscenter_Action_List_off.focus();
}

// ����
function CsRegCancelProc(frm){
    if (confirm('��ϵ� ���� ������ ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "deletecsas";
        frm.submit();
    }
}

//��üa/s , ��üa/s(����ȸ��) �ּ��� ����
function popEditCsDelivery(CsAsID){	
    var window_width = 600;
    var window_height = 450;
    
    var popEditCsDelivery=window.open("/admin/offshop/shopcscenter/action/pop_CsDeliveryEdit.asp?CsAsID=" + CsAsID ,"popEditCsDelivery","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popEditCsDelivery.focus();
}

//���� a/s ó�� ������
function PopmaejangAction(orderno,shopid,divcd,currstate){
    var PopmaejangAction=window.open("/common/offshop/shopcscenter/shop_cslist.asp?searchfield=01&searchstring="+orderno+"&shopid="+shopid+"&divcd="+divcd+"&currstate="+currstate,"PopmaejangAction","width=1024 height=768 scrollbars=yes resizable=yes");
    PopmaejangAction.focus();
}

//������ ���
function popOrderReceipt(orderno){
    var window_width = 750;
    var window_height = 700;
    var popwin=window.open("/admin/offshop/shopcscenter/order/pop_order_receipt.asp?orderno=" + orderno ,"popOrderReceipt","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}