<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB

Dim npartner
set npartner = new CPartnerUser
npartner.FRectDesignerID = session("ssBctID")
npartner.GetOnePartnerNUser


dim i,j,k
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script>
function UseTemplate() {
	window.open("/common/pop_basic_item_info_list.asp", "UseTemplate", "width=700, height=450, scrollbars=yes, resizable=yes");
}

// ============================================================================
// ��ü�����ڵ��Է�
function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value = varArray[0];
	document.itemreg.margin.value = varArray[1];

}

function CalcuAuto(frm){
	var imargin, isellcash, ibuycash;
	var isellvat, ibuyvat, imileage;
	imargin = frm.margin.value;
	isellcash = frm.sellcash.value;

	isvatinclude = frm.vatinclude[0].checked;

	if (imargin.length<1){
		alert('������ �Է��ϼ���.');
		frm.margin.focus();
		return;
	}

	if (isellcash.length<1){
		alert('�ǸŰ��� �Է��ϼ���.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDouble(imargin)){
		alert('������ ���ڷ� �Է��ϼ���.');
		frm.margin.focus();
		return;
	}

	if (!IsDigit(isellcash)){
		alert('�ǸŰ��� ���ڷ� �Է��ϼ���.');
		frm.sellcash.focus();
		return;
	}

	if (isvatinclude==true){
		isellvat = parseInt(parseInt(1/11 * parseInt(isellcash)));
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		ibuyvat = parseInt(parseInt(1/11 * parseInt(ibuycash)));
		imileage = parseInt(isellcash*0.01) ;
	}else{
		isellvat = 0;
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		ibuyvat = 0;
		imileage = parseInt(isellcash*0.01) ;
	}

	frm.sellvat.value = isellvat;
	frm.buycash.value = ibuycash;
	frm.buyvat.value = ibuyvat;
	frm.mileage.value = imileage;
}

// ============================================================================
// ī�װ����
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.itemreg;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;

	//ī�װ��� ���� Enable ���� -�ö��
	EnDisableFlowerShop();
}

// ============================================================================
// �ɼǼ���
function editItemOption(itemid, waityn) {
	var param = "itemid=" + itemid + "&waityn=" + waityn;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=800,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function doEditItemOption(itemid, waityn, arrmode, arritemoption, arritemoptionname, arroptuseyn, arroptsellyn, arroptlimityn, arroptlimitno, arroptlimitsold) {
	alert("a");
	// var param = "itemid=" + itemid + "&waityn=" + waityn;

	// popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=700,height=400,scrollbars=yes,resizable=yes');
	// popwin.focus();
}

function popNormalOptionAdd() {
	popwin = window.open('/common/module/normalitemoptionadd.asp' ,'popNormalOptionAdd','width=540,height=260,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popEtcOptionAdd(){
	popwin = window.open('/common/module/etcitemoptionadd.asp' ,'normalitemoptionadd','width=540,height=260,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// �ɼ��� �߰��Ѵ�
function InsertOptionWithGubun(ioptTypeName, ft, fv) {
	var frm = document.itemreg;

	//�ɼǰ��� �������� ������ skip ,����ɼ��ΰ�� ����
	if (fv!="0000"){
		for (i=0;i<frm.realopt.length;i++){
			if (frm.realopt[i].value==fv){
				return;
			}
		}
	}

    frm.optTypeNm.value = ioptTypeName;
	frm.elements['realopt'].options[frm.realopt.options.length] = new Option(ft, fv);
}

// ���õ� �ɼ� ����
function delItemOptionAdd()
{
	var frm = document.itemreg;
	var sidx = frm.realopt.options.selectedIndex;

	if(sidx<0){
		alert("������ �ɼ��� �������ֽʿ�.");
	}else{
	    for(i=0; i<frm.realopt.options.length; i++){
    		if(frm.realopt.options[i].selected){
    			frm.realopt.options[i] = null;
    			i=i-1;
    		}
    	}

		if (frm.realopt.options.length<1){
		    frm.optTypeNm.value = '';
		}

		//frm.realopt.options[sidx] = null;
	}
}


// ============================================================================
// �̹���ǥ��
function ClearImage(img,fsize,wd,ht) {
    img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";
}

function CheckExtension(imgname, allowext) {
    var ext = imgname.lastIndexOf(".");
    if (ext < 0) {
        return false;
    }

    ext = imgname.toLowerCase().substring(ext + 1);
    allowext = "," + allowext + ",";
    if (allowext.indexOf(ext) < 0) {
        return false;
    }

    return true;
}

function CheckImage(img, filesize, imagewidth, imageheight, extname, fsize)
{
    var ext;
    var filename;

	filename = img.value;
	if (img.value == "") { return false; }

    if (CheckExtension(filename, extname) != true) {
        alert("�̹���ȭ���� ������ ȭ�ϸ� ����ϼ���.[" + extname + "]");
        ClearImage(img,fsize, imagewidth, imageheight);
        return false;
    }

    return true;
}

function getByteLength(inputValue) {
     var byteLength = 0;
     for (var inx = 0; inx < inputValue.length; inx++) {
         var oneChar = escape(inputValue.charAt(inx));
         if ( oneChar.length == 1 ) {
             byteLength ++;
         } else if (oneChar.indexOf("%u") != -1) {
             byteLength += 2;
         } else if (oneChar.indexOf("%") != -1) {
             byteLength += oneChar.length/3;
         }
     }
     return byteLength;
 }


// ============================================================================
// �����ϱ�
function SubmitSave() {

	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
		return;
	}

	if (!$("input[name='isDefault'][value='y']").length){
		alert("���� ī�װ��� �����ϼ���.\n�� ���� �⺻ ī�װ��� �ʼ� �ֽ��ϴ�.");
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }

	//��ǰ�� ����üũ �߰� 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("��ǰ���� �ִ� 64byte ���Ϸ� �Է����ּ���.(�ѱ�32�� �Ǵ� ����64��)");
		itemreg.itemname.focus();
		return;
	}

    //��۱��� üũ =======================================
    //��ü ���ǹ��
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[3].checked){
            alert('��� ������ Ȯ�����ּ���. [��ü ���ǹ��] ��ü�� �ƴմϴ�.');
            itemreg.deliverytype[3].focus();
            return;
        }
    }

    //��ü���ҹ�� : ���ǹ�۵� ���Ҽ�������
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype[4].checked)){
        alert('��� ������ Ȯ�����ּ���. [��ü ���ҹ��,��ü ���ǹ��] ��ü�� �ƴմϴ�.');
        itemreg.deliverytype[4].focus();
        return;
    }

    //��۱��� ��ü�̳� ���Ա����� ��ü�� �ƴѰ�.
    if ((itemreg.deliverytype[1].checked)||(itemreg.deliverytype[3].checked)||(itemreg.deliverytype[4].checked)){
        if ((itemreg.mwdiv[0].checked)||(itemreg.mwdiv[1].checked)){
            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
            itemreg.deliverytype[1].focus();
            return;
        }
    }

    //���Ա����� ��ü�̳� ��۱����� ��ü�� �ƴѰ�.
    if (itemreg.mwdiv[2].checked){
        if ((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
            itemreg.deliverytype[0].focus();
            return;
        }
    }

    //��ü��۸� �ֹ����� ����.
    if ((!itemreg.mwdiv[2].checked)&&(itemreg.itemdiv[1].checked)){
        alert('�ֹ����� ��ǰ�� ��ü����ΰ�츸 �����մϴ�.');
        itemreg.itemdiv[0].focus();
        return;
    }

	if(document.itemreg.deliverfixday[1].checked) {
		if(document.itemreg.freight_min.value<=0||document.itemreg.freight_max.value<=0) {
            alert('ȭ����� ����� �Է����ּ���.');
            document.itemreg.freight_min.focus();
            return;
		}
	}

    //==================================================================================

	if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 20000000){
		alert("�Ǹ� ������ 400�� �̻� 20,000,000�� �̸����� ��� �����մϴ�.");
		itemreg.sellcash.focus();
		return;
	}

	if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
        alert("���ϸ����� �ǸŰ����� Ŭ �� �����ϴ�.");
        itemreg.mileage.focus();
        return;
    }

	if((itemreg.sellcash.value*0.05) <= itemreg.mileage.value*1){
	  	alert("���ϸ����� 1% �̻� 5% ���Ϸθ� ��� �����մϴ�.");
	  	itemreg.mileage.focus();
	  	return;
	}

	//��ǰ ���� �Ұ��׸� �˻�
	var cntRe = /.js["'>\s]/gi;
	if(cntRe.test(itemreg.itemcontent.value)) {
        alert('��ǰ������ js������ ���� �� �����ϴ�.');
        itemreg.itemcontent.focus();
        return;
	}

	//��ǰ ǰ������
    if (!itemreg.infoDiv.value){
        alert('��ǰ�� �ش��ϴ� ǰ���� �������ֽʽÿ�.');
        itemreg.infoDiv.focus();
        return;
    } else if(itemreg.infoDiv.value=="35") {
    	if(!itemreg.itemsource.value) {
	        alert('��ǰ�� ������ �Է����ּ���.');
	        itemreg.itemsource.focus();
	        return;
    	}
    	if(!itemreg.itemsize.value) {
	        alert('��ǰ�� ũ�⸦ �Է����ּ���.');
	        itemreg.itemsize.focus();
	        return;
    	}
    }

	//������������
    if (itemreg.safetyYn[0].checked){
	    if (!itemreg.safetyDiv.value){
	        alert('�������������� �������ּ���.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
	    if (!itemreg.safetyNum.value){
	        alert('����������ȣ�� �Է����ּ���.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
    }

    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("���������� �Է����ּ���!");
        itemreg.limitno.focus();
        return;
    }

	//=== �ɼ� ================================================
    if (itemreg.useoptionyn[0].checked == true) {
	    if (itemreg.optlevel[0].checked == true) {
	    //���Ͽɼ�
    	    if (itemreg.realopt.length < 1) {
                alert("�߰��� �ɼ��� �����ϴ�.");
                // itemreg.useoptionyn.focus();
                return;
            }

    	    if (itemreg.realopt.length < 2) {
                alert("�ɼ��� �ΰ� �̻��̾�� �մϴ�.(�ɼǺ��� ����/���ü����� �����մϴ�.)");
                // itemreg.useoptionyn.focus();
                return;
            }
        }else if (itemreg.optlevel[1].checked == true) {
        //���߿ɼ�
            if ((itemreg.optionTypename1.value.length<1)||(itemreg.optionTypename2.value.length<1)){
                alert("���߿ɼ��� ����� ��� �ɼǱ��и� �� �ּ� 2�� �̻� ����ϼž� �մϴ�.");
                itemreg.optionTypename2.focus();
                return;
            }

            if ((fnTrim(itemreg.optionTypename1.value)==fnTrim(itemreg.optionTypename2.value))||(fnTrim(itemreg.optionTypename2.value)==fnTrim(itemreg.optionTypename3.value))||(fnTrim(itemreg.optionTypename1.value)==fnTrim(itemreg.optionTypename3.value))){
                alert('���߿ɼ��� �ɼ� ���и��� ���� �ٸ��� �����ؾ� �մϴ�.');
                itemreg.optionTypename2.focus();
                return;
            }

            var chkCnt=0;
            for (var i=0;i<itemreg.optionName1.length;i++){
                if (itemreg.optionName1[i].value.length>0) chkCnt++;
            }

            if (chkCnt<2){
                alert("�ɼ��� �� ���д� 2�� �̻��̾�� �մϴ�.");
                itemreg.optionName1[1].focus();
                return;
            }

            chkCnt=0;

            for (var i=0;i<itemreg.optionName2.length;i++){
                if (itemreg.optionName2[i].value.length>0) chkCnt++;
            }

            if (chkCnt<2){
                alert("�ɼ��� �� ���д� 2�� �̻��̾�� �մϴ�.");
                itemreg.optionName2[1].focus();
                return;
            }

            if (itemreg.optionTypename3.value.length>0){
                chkCnt=0;

                for (var i=0;i<itemreg.optionName3.length;i++){
                    if (itemreg.optionName3[i].value.length>0) chkCnt++;
                }

                if (chkCnt<2){
                    alert("�ɼ��� �� ���д� 2�� �̻��̾�� �մϴ�.");
                    itemreg.optionName3[1].focus();
                    return;
                }

            }
        }
	}

    var optiont = "";
    var optionv = "";
    var optvalue = 11; // ����ɼ�(11 - 99)
    for(var i = 0; i < itemreg.realopt.options.length; i++) {
        optiont += (itemreg.realopt.options[i].text + "|");

        // ����ɼ��߰�
        if (itemreg.realopt.options[i].value == "0000") {
            if (optvalue > 99) {
                alert("�ʹ����� �ɼ��� �߰��ϼ̽��ϴ�.");
                return;
            }
            itemreg.realopt.options[i].value = "00" + optvalue;
            optvalue = optvalue + 1;
        }

        optionv += (itemreg.realopt.options[i].value + "|");
    }

	// �⺻ ������
	if(!itemreg.DFcolorCD.value) {
        alert("��ǰ�� �⺻ ������ �������ּ���.");
        return;
	}
    if (itemreg.imgDFColor.value != "") {
        if (CheckImage(itemreg.imgDFColor, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40) != true) {
            return;
        }
    }

	//=== ��ǰ �̹��� ================================
    //if(itemreg.regimg.checked==false) {
	    if (itemreg.imgbasic.value=="") {
	        alert("�⺻�̹����� �ʼ��Դϴ�.");
	        return;
	    } else {
	        if (CheckImage(itemreg.imgbasic, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40) != true) {
	            return;
	        }
	    }
	//}

    if (itemreg.imgmask.value != "") {
        if (CheckImage(itemreg.imgmask, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40) != true) {
            return;
        }
    }

    if (itemreg.imgadd1.value != "") {
        if (CheckImage(itemreg.imgadd1, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40) != true) {
            return;
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage(itemreg.imgadd2, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40) != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage(itemreg.imgadd3, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40) != true) {
            return;
        }
    }

    if (itemreg.imgadd4.value != "") {
        if (CheckImage(itemreg.imgadd4, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40) != true) {
            return;
        }
    }

    if (itemreg.imgadd5.value != "") {
        if (CheckImage(itemreg.imgadd5, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40) != true) {
            return;
        }
    }

    if (itemreg.imgmain.value != "") {
        if (CheckImage(itemreg.imgmain, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40) != true) {
            return;
        }
    }

    if (itemreg.imgmain2.value != "") {
        if (CheckImage(itemreg.imgmain2, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40) != true) {
            return;
        }
    }

    if (itemreg.imgmain3.value != "") {
        if (CheckImage(itemreg.imgmain3, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40) != true) {
            return;
        }
    }

    // ���󰡰�
	if (confirm("�Һ��ڰ�(" + itemreg.sellcash.value + ")/���ް�(" + itemreg.buycash.value + ")�� ��Ȯ�� �ԷµǾ����ϱ�?") == false) {
		itemreg.sellcash.focus();
		return;
    }

    if(confirm("��ǰ�� �ø��ðڽ��ϱ�? \n���MD ������ �ݿ� �˴ϴ�.") == true){
        itemreg.itemoptioncode2.value = optionv;
        itemreg.itemoptioncode3.value = optiont;

		itemreg.deliverytype[0].disabled=false;
		itemreg.deliverytype[1].disabled=false;
		itemreg.deliverytype[2].disabled=false;
        itemreg.deliverytype[3].disabled=false;
        itemreg.deliverytype[4].disabled=false;

        itemreg.target = "FrameCKP";
        itemreg.submit();
    }

}

function TnCheckUpcheYN(frm){
	if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
		frm.deliverytype[0].checked=true;	// �⺻üũ
		// ��۱��� ����(�ٹ�����)
		frm.deliverytype[0].disabled=false;
		frm.deliverytype[1].disabled=true;
		frm.deliverytype[2].disabled=true;  //�ٹ����� �������� üũ �� �� ����.
		frm.deliverytype[3].disabled=true;  //��ü�������(9)
		frm.deliverytype[4].disabled=true;  //��ü���ҹ��(7) : ��ü���� �����Ұ�
        frm.optlevel[0].checked=true;
        frm.optlevel[1].disabled=true;
	}
	else if(frm.mwdiv[2].checked){
		// ��۱��� ����(��ü���)
	    if ((frm.defaultFreeBeasongLimit.value*1>0)&&(frm.defaultDeliverPay.value*1>0)){
	        frm.deliverytype[3].checked=true;	// ��ü���ǹ�� �⺻ üũ
	    }else if(frm.defaultDeliveryType.value=="7"){
	        frm.deliverytype[4].checked=true;	// ��ü���ҹ�� �⺻ üũ
	    }else{
	        frm.deliverytype[1].checked=true;	// �⺻ üũ
	    }

		frm.deliverytype[0].disabled=true;
		frm.deliverytype[1].disabled=false;
		frm.deliverytype[2].disabled=true;
        frm.deliverytype[3].disabled=false; //��ü�������(9)
        frm.deliverytype[4].disabled=false;  //��ü���ҹ��(7) : ��ü���� �����Ұ�
        frm.optlevel[1].disabled=false;
	}
}

function TnCheckFixday(frm) {
	if(frm.deliverfixday[0].checked) {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=true;
		frm.deliverarea[2].disabled=true;
		document.getElementById("lyrFreightRng").style.display="none";
	} else if(frm.deliverfixday[1].checked) {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=true;
		frm.deliverarea[2].disabled=true;
		document.getElementById("lyrFreightRng").style.display="";
	} else {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=false;
		frm.deliverarea[2].disabled=false;
		document.getElementById("lyrFreightRng").style.display="none";
	}
}

function EnDisableFlowerShop(){

    var frm = document.itemreg;
    if ((frm.cd1.value=="110")&&(frm.cd2.value=="060")){
        frm.deliverfixday[2].disabled = false;
    }else{
        frm.deliverfixday[2].disabled = true;
        frm.deliverfixday[0].checked = true;
        frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=true;
		frm.deliverarea[2].disabled=true;
    }
}

function TnAutoChkDeliver() {
	var frm = document.itemreg
	switch(frm.defaultmaeipdiv.value) {
		case "M" :
			frm.mwdiv[0].checked=true;
			break;
		case "W" :
			frm.mwdiv[1].checked=true;
			break;
		case "U" :
			frm.mwdiv[2].checked=true;
			break;
	}
	TnCheckUpcheYN(frm);
}

function TnGoClear(frm){
	frm.sellvat.value = "";
	frm.buycash.value = "";
	frm.buyvat.value = "";
	frm.mileage.value = "";
}

function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if (frm.mwdiv[2].checked){
			alert("����Ư�� ������ ��ü�� ���\n��۱����� �ٹ����� ������� ���� �Ͻ� �� �����ϴ�!!\n����Ư�������� Ȯ�����ּ���!!");
			frm.mwdiv[0].checked=true;
			frm.optlevel[1].checked=false;
			frm.optlevel[1].disabled=true;
		}
	}
	//else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
	else if(frm.deliverytype[1].checked){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("����Ư�� ������ �����̳� Ư���� ���\n��۱�����  ��ü������� ���� �Ͻ� �� �����ϴ�!!!\n����Ư�������� Ȯ�����ּ���!!");
			frm.mwdiv[2].checked=true;
			frm.optlevel[1].disabled=false;
		}
	}
}

function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // �ɼǻ��
	    opttype.style.display="";

	    if (frm.optlevel[1].checked==true){
	    	document.getElementById("optlist").style.display="none";
	    	document.getElementById("optlist2").style.display="";
	    } else {
	    	document.getElementById("optlist").style.display="";
	    	document.getElementById("optlist2").style.display="none";
	    }

        document.itemreg.optlevel.value= "1";

	} else {
	    // �ɼǾ���
	    opttype.style.display="none";
	    while (frm.realopt.length > 0) {
	        frm.realopt.options[0] = null;
        }
    	document.getElementById("optlist").style.display="none";
    	document.getElementById("optlist2").style.display="none";

        frm.itemoptioncode2.value = "";
        frm.itemoptioncode3.value = "";
    }
}

//�����ڵ� ����
function selColorChip(cd) {
	var i;
	itemreg.DFcolorCD.value= cd;
	for(i=0;i<=31;i++) {
		document.all("cline"+i).bgColor='#DDDDDD';
	}
	if(!cd) document.all("cline0").bgColor='#DD3300';
	else document.all("cline"+cd).bgColor='#DD3300';
}

// ============================================================================
// �̹��� �˸�â
function PopImageInformation(){
	window.open("itemreg_info_win.asp","PopImageInformation","width=920,height=600,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no");
}

function ClearVal(comp){
    comp.value = "";
}

function checkItemDiv(comp){
    var frm = comp.form;

    if (comp.name=="itemdiv"){
        if (frm.itemdiv[1].checked){
            frm.reqMsg.disabled=false;
        }else{
            //frm.reqMsg.checked=false;
            frm.reqMsg.disabled=true;
        }
    }

    //�ֹ����� ��ǰ�ΰ��.
    if (frm.itemdiv[1].checked){
        if (frm.reqMsg.checked){
            frm.itemdiv[1].value="06";
        }else{
            frm.itemdiv[1].value="16";
        }
    }
}

// ������������ ����
function chgSafetyYn(frm) {
	if(frm.safetyYn[0].checked) {
		frm.safetyDiv.disabled=false;
		frm.safetyNum.disabled=false;
	} else {
		frm.safetyDiv.disabled=true;
		frm.safetyNum.disabled=true;
	}
}

//ǰ�� ���� / ǰ�񳻿� ǥ��
function chgInfoDiv(v) {
	$("#itemInfoList").empty();

	if(v=="") {
		$("#itemInfoCont").hide();
	} else {
		$("#itemInfoCont").show();

		var str = $.ajax({
			type: "POST",
			url: "/admin/itemmaster/act_waitItemInfoDivForm.asp",
			data: "ifdv="+v,
			dataType: "html",
			async: false
		}).responseText;

		if(str!="") {
			$("#itemInfoList").html(str);
		}
	}

	if(v=="35") {
		$("#lyItemSrc").show();
		$("#lyItemSize").show();
	} else {
		$("#lyItemSrc").hide();
		$("#lyItemSize").hide();
	}
}

//�ܼ� ���� ������
function chgInfoChk(fm) {
	$(fm).parent().parent().find('[name="infoChk"]').val($(fm).val());
}

//���� ���� ������
function chgInfoSel(fm) {
	$(fm).parent().parent().find('[name="infoChk"]').val($(fm).val());
	$(fm).parent().parent().find('[name="infoCont"]').val($(fm).attr("msg"));

	if($(fm).val()=="Y") {
		$(fm).parent().parent().find('[name="infoCont"]').removeAttr("readonly");
		$(fm).parent().parent().find('[name="infoCont"]').removeClass("text_ro");
		$(fm).parent().parent().find('[name="infoCont"]').addClass("text");
	} else {
		$(fm).parent().parent().find('[name="infoCont"]').attr("readonly", true);
		$(fm).parent().parent().find('[name="infoCont"]').addClass("text_ro");
	}
}

<!--2013 ������ �߰� ------->
$(function(){
	// �ε��� ��ǰ�Ӽ� ���� ���
	printItemAttribute();
	// �ε��� �⺻ ������� ����
	TnAutoChkDeliver();
});

// ��ǰ�Ӽ� ���
function printItemAttribute() {
	var arrDispCd="";
	$("input[name='catecode']").each(function(i){
		if(i>0) arrDispCd += ",";
		arrDispCd += $(this).val();
	});
	$.ajax({
		url: "/common/module/act_waitItemAttribSelect.asp?itemid=0&arrDispCate="+arrDispCd,
		cache: false,
		success: function(message) {
			$("#lyrItemAttribAdd").empty().append(message);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

// ����ī�װ� ���� �˾�
function popDispCateSelect(){
	var dCnt = $("input[name='isDefault'][value='y']").length;
	var cCnt = $("input[name='isDefault']").length;
	$.ajax({
		url: "/common/module/act_DispCategorySelectUpche.asp?isDft="+dCnt+"&chk="+cCnt,
		cache: false,
		success: function(message) {
			$("#lyrDispCateAdd").empty().append(message).fadeIn();
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

// ���̾�� ����ī�װ� �߰�
function addDispCateItem(dcd,cnm,div,dpt) {
	// ������ ���� �ߺ� ī�װ� ���� �˻�
	if(tbl_DispCate.rows.length>0)	{
		if(tbl_DispCate.rows.length>1)	{
			for(l=0;l<document.all.isDefault.length;l++)	{
			    if((document.all.catecode[l].value==dcd)) {
					alert("�̹� ������ ���� ī�װ��� �ֽ��ϴ�..");
					return;
				}
			}
		}
		else {
		    if((document.all.catecode.value==dcd)) {
				alert("�̹� ������ ���� ī�װ��� �ֽ��ϴ�..");
				return;
			}
		}
	}

	// ���߰�
	var oRow = tbl_DispCate.insertRow();
	oRow.onmouseover=function(){tbl_DispCate.clickedRowIndex=this.rowIndex};

	// ���߰� (����,ī�װ�,������ư)
	var oCell1 = oRow.insertCell();
	var oCell2 = oRow.insertCell();
	var oCell3 = oRow.insertCell();

	if(div=="y") {
		oCell1.innerHTML = "<font color='darkred'><b>[�⺻]<b></font><input type='hidden' name='isDefault' value='y'>";
	} else {
		oCell1.innerHTML = "<font color='darkblue'>[�߰�]</font><input type='hidden' name='isDefault' value='n'>";
	}
	$(cnm).each(function(i){
		if(dpt>i) {
			if(i>0) oCell2.innerHTML += " >> ";
			oCell2.innerHTML += $(this).text();
		}
	});
	oCell2.innerHTML += "<input type='hidden' name='catecode' value='" + dcd + "'>";
	oCell2.innerHTML += "<input type='hidden' name='catedepth' value='" + dpt + "'>";
	oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle>";
	$("#lyrDispCateAdd").fadeOut();

	if($("input[name='isDefault']").length>1) {
		$("#btnAddCate").hide();
	}

	//��ǰ�Ӽ� ���
	printItemAttribute();
}

// ���� ����ī�װ� ����
function delDispCateItem() {
	if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?")) {
		tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

		if($("input[name='isDefault']").length<2) {
			$("#btnAddCate").show();
		}

		//��ǰ�Ӽ� ���
		printItemAttribute();
	}
}

</script>

<form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/WaitUpcheItemRegister_Process.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="itemoptioncode2">
<input type="hidden" name="itemoptioncode3">
<input type="hidden" name="defultmargine" value="<%= npartner.FOneItem.Fdefaultmargine %>">
<input type="hidden" name="defaultmaeipdiv" value="<%= npartner.FOneItem.Fmaeipdiv %>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%= npartner.FOneItem.FdefaultFreeBeasongLimit %>">
<input type="hidden" name="defaultDeliverPay" value="<%= npartner.FOneItem.FdefaultDeliverPay %>">
<input type="hidden" name="defaultDeliveryType" value="<%= npartner.FOneItem.FdefaultDeliveryType %>">
<input type="hidden" name="DFcolorCD" value="">

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- 1.�Ϲ����� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>1.�Ϲ�����</strong>
        </td>
        <td align="right">
          <input type="button" value="�⺻Ʋ����" class="button" onClick="UseTemplate();">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�귣��ID :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="designerid"  value="<%= session("ssBctID") %>" class="text_ro" readonly size="30" id="[on,off,off,off][�귣��ID]">&nbsp;
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemname" maxlength="64" size="50" class="text" id="[on,off,off,off][��ǰ��]">&nbsp;
  	</td>
  </tr>
</table>

<!-- 2.���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>2.����</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF" title="���/���� ���� ���� ī�װ�" style="cursor:help;">���� ī�װ� :</td>
    <input type="hidden" name="cd1" value="">
    <input type="hidden" name="cd2" value="">
    <input type="hidden" name="cd3" value="">
  	<td bgcolor="#FFFFFF" colspan="2">
      <input type="text" name="cd1_name" value="" id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro">
      <input type="text" name="cd2_name" value="" id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro">
      <input type="text" name="cd3_name" value="" id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro">

      <input type="button" value="ī�װ� ����" class="button" onclick="editCategory(itemreg.cd1.value,itemreg.cd2.value,itemreg.cd3.value);">
  	</td>
  </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="����Ʈ�� ������ ī�װ�" style="cursor:help;">���� ī�װ� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table class=a>
		<tr>
			<td id="lyrDispList">
				<table id='tbl_DispCate' class=a></table>
			</td>
			<td valign="bottom"><input id="btnAddCate" type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF" >
      <label><input type="radio" name="itemdiv" value="01" checked onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�Ϲݻ�ǰ</label>
      <br>
	  <label><input type="radio" name="itemdiv" value="06" onClick="document.getElementById('lyRequre').style.display='block';checkItemDiv(this);">�ֹ� ���ۻ�ǰ</label>
	  <input type="checkbox" name="reqMsg" value="10" onClick="checkItemDiv(this);">�ֹ����� ���� �ʿ�<font color=red>(�ֹ��� �̴ϼȵ� ���۹����� �ʿ��Ѱ�� üũ)</font>
  	</td>
  	<td bgcolor="#FFFFFF" >
  	    <div id="lyRequre" style="display:none;padding-left:22px;">
      	�������ۼҿ��� <input type="text" name="requireMakeDay" value="0" size="2" class="text" id="[off,on,off,off][�������ۼҿ���]">��
      	<font color="red">(��ǰ�߼��� ��ǰ���� �Ⱓ)</font>
      </div>
  	</td>
  </tr>
</table>

<!-- 3.�������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>3.��������</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
    <td width="15%" bgcolor="#DDDDFF">����, �鼼 ���� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="radio" name="vatinclude" value="Y" checked onclick="TnGoClear(this.form);">����
      <input type="radio" name="vatinclude" value="N" onclick="TnGoClear(this.form);">�鼼
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">�⺻ ���� ���� :</td>
  	<td bgcolor="#FFFFFF">
      <input type="text" name="margin" maxlength="32" size="5" id="[off,off,off,off][����]" onKeyUp="CalcuAuto(itemreg);" value="<% =npartner.FOneItem.Fdefaultmargine %>">%
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ǸŰ�(�Һ��ڰ�) :</td>
  	<input type="hidden" name="sellvat">
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="sellcash" size="12" id="[on,on,off,off][�Һ��ڰ�]" onKeyUp="CalcuAuto(itemreg);" maxlength="8" class="text">��
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">���ް� :</td>
  	<input type="hidden" name="buyvat">
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="buycash" maxlength="16" size="12" id="[on,on,off,off][���ް�]" readonly class="text_ro">��
      (<b>�ΰ��� ���԰�</b>)
  	</td>
  </tr>
  <tr>
  	<td bgcolor="#DDDDFF"></td>
  	<td bgcolor="#F8F8F8" colspan="3">
      - ���ް��� <b>�ΰ��� ���԰�</b>�Դϴ�.<br>
      - �Һ��ڰ�(���ΰ�)�� ����(���θ���)�� �Է��ϸ� ���ް��� ���ϸ����� �ڵ����˴ϴ�.
  	</td>
  </tr>
  <input type="hidden" name="mileage" value="0">
</table>

<!-- 4.�������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>4.��������</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="��ǰ �� �Ӽ�" style="cursor:help;">��ǰ�Ӽ� :</td>
	<td id="lyrItemAttribAdd" bgcolor="#FFFFFF" colspan="3"></td>
</tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ü��ǰ�ڵ� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	    <input type="text" name="upchemanagecode" value="" size="20" maxlength="32" class="text" id="[off,off,off,off][��ü��ǰ�ڵ�]">
  	    (��ü���� �����ϴ� �ڵ� �ִ� 32�� - ����/���ڸ� ����)
  	</td>
  </tr>
</table>

<!-- 5.�⺻���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>5.�⺻����</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="makername" maxlength="32" size="25" class="text" id="[on,off,off,off][������]">&nbsp;(������ü��)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="sourcearea" maxlength="64" size="25" class="text" id="[on,off,off,off][������]">&nbsp;(ex:�ѱ�,�߱�,�߱�OEM,�Ϻ�...)
      <br>( ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�˻�Ű���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="keywords" maxlength="128" size="60" class="text" id="[on,off,off,off][�˻�Ű����]">&nbsp;(�޸��α��� ex: Ŀ��,Ƽ����,����)
  	</td>
  </tr>
</table>

<!-- 5-1.ǰ������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- ǰ������� </strong> &nbsp;<font color=gray>��ǰ����������� ���� ���� ������ ���� �Ʒ� ������ ��Ȯ�� �Է����ֽñ� �ٶ��ϴ�.</font></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ���� :</td>
	<td bgcolor="#FFFFFF">
		<select name="infoDiv" class="select" onchange="chgInfoDiv(this.value)">
		<option value="">::��ǰǰ��::</option>
		<option value="01">�Ƿ�</option>
		<option value="02">����/�Ź�</option>
		<option value="03">����</option>
		<option value="04">�м���ȭ(����/��Ʈ/�׼�����)</option>
		<option value="05">ħ����/Ŀư</option>
		<option value="06">����(ħ��/����/��ũ��/DIY��ǰ)</option>
		<option value="07">������(TV��)</option>
		<option value="08">������ ������ǰ(�����/��Ź��/�ı⼼ô��/���ڷ�����)</option>
		<option value="09">��������(������/��ǳ��)</option>
		<option value="10">�繫����(��ǻ��/��Ʈ��/������)</option>
		<option value="11">���б��(������ī�޶�/ķ�ڴ�)</option>
		<option value="12">��������(MP3/���ڻ��� ��)</option>
		<option value="14">������̼�</option>
		<option value="15">�ڵ�����ǰ(�ڵ�����ǰ/��Ÿ �ڵ�����ǰ)</option>
		<option value="16">�Ƿ���</option>
		<option value="17">�ֹ��ǰ</option>
		<option value="18">ȭ��ǰ</option>
		<option value="19">�ͱݼ�/����/�ð��</option>
		<option value="20">��ǰ(����깰)</option>
		<option value="21">������ǰ</option>
		<option value="22">�ǰ���ɽ�ǰ/ü��������ǰ</option>
		<option value="23">�����ƿ�ǰ</option>
		<option value="24">�Ǳ�</option>
		<option value="25">��������ǰ</option>
		<option value="26">����</option>
		<option value="35">��Ÿ</option>
		</select>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:none">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ�񳻿� :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList"></td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">��ǰ���������� ������ ���� �Ǿ��ִ��� ��Ȯ�� �Է¹ٶ��ϴ�. ����Ȯ�ϰų� �߸��� ���� �Է½�, �׿� ���� å���� ���� ���� �ֽ��ϴ�.</font></td>
</tr>
<tr align="left" id="lyItemSrc" style="display:none;">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsource" maxlength="64" size="50" class="text">&nbsp;(ex:�ö�ƽ,����,��,...)
	</td>
</tr>
<tr align="left" id="lyItemSize" style="display:none;">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ������ :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsize" maxlength="64" size="50" class="text">
		<select name="unit" class="select">
			<option value="">�����Է�</option>
			<option value="mm">mm</option>
			<option value="cm" selected>cm</option>
			<option value="m��">m��</option>
			<option value="km">km</option>
			<option value="m��">m��</option>
			<option value="km��">km��</option>
			<option value="ha">ha</option>
			<option value="m��">m��</option>
			<option value="cm��">cm��</option>
			<option value="L">L</option>
			<option value="g">g</option>
			<option value="Kg">Kg</option>
			<option value="t">t</option>
		</select>
		&nbsp;(ex:7.5x15(cm))
	</td>
</tr>
</table>
<!-- 5-2.������������ -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- ������������</strong></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">����������� :</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="safetyYn" value="Y" onclick="chgSafetyYn(document.itemreg)">���</label>
		<label><input type="radio" name="safetyYn" value="N" checked onclick="chgSafetyYn(document.itemreg)">���ƴ�</label><br />
		<select name="safetyDiv" disabled class="select">
		<option value="">::������������::</option>
		<option value="10">������������(KC��ũ)</option>
		<option value="20">�����ǰ ��������</option>
		<option value="30">KPS �������� ǥ��</option>
		<option value="40">KPS �������� Ȯ�� ǥ��</option>
		<option value="50">KPS ��� ��ȣ���� ǥ��</option>
		</select>
		������ȣ <input type="text" name="safetyNum" disabled size="35" maxlength="25" class="text" value="" />
		<font color="darkred">���ƿ�ǰ�̳� �����ǰ�� ��� �ʼ� �Է�</font>
	</td>
</tr>
</table>

<!-- 6.������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>6.�������</strong>
        </td>
        <td align="right">
        	<input type="button" class="button" value="����������� ����" onclick="TnAutoChkDeliver()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">����Ư������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="mwdiv" value="M" checked onclick="TnCheckUpcheYN(this.form);">����</label>
		<label><input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);">Ư��</label>
		<label><input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);">��ü���</label>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��۱��� :</td>
  	<td width="85%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverytype" value="1" checked  onclick="TnCheckUpcheDeliverYN(this.form);">�ٹ����ٹ��</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);">��ü(����)���</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);">�ٹ����ٹ�����</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);">��ü���ǹ��(���� ��ۺ�ΰ�)</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);">��ü���ҹ��</label>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��۹�� :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverfixday" value="" checked onclick="TnCheckFixday(this.form)">�ù�(�Ϲ�)</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="X" onclick="TnCheckFixday(this.form)">ȭ��</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="C" disabled onclick="TnCheckFixday(this.form)">�ö��������</label>
		<span id="lyrFreightRng" style="display:none;">
			<br />&nbsp;
			��ǰ/��ȯ �� ȭ����� ���(��) :
			�ּ� <input type="text" name="freight_min" class="text" size="6" value="0" style="text-align:right;">�� ~
			�ִ� <input type="text" name="freight_max" class="text" size="6" value="0" style="text-align:right;">��
		</span>
		<br>&nbsp;<font color="red">(�ö�� ��ǰ�� ��츸 �����ǹ��, ������, �ö�������� �ɼ��� ��밡���մϴ�.)</font>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">������� :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverarea" value="" checked>�������</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="C" disabled >�����ǹ��</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="S" disabled >������</label>
  	</td>
  </tr>
  <input type="hidden" name="pojangok" value="N">
  <input type="hidden" name="sellyn" value="N">
  <input type="hidden" name="dispyn" value="N">
  <input type="hidden" name="isusing" value="Y">
</table>

<!-- 7.�ɼ����� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>7.�ɼ�����</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ɼǱ��� :</td>
	<td width="85%" bgcolor="#FFFFFF">
		<label><input type="radio" name="useoptionyn" value="Y" onClick="TnCheckOptionYN(this.form);">�ɼǻ����</label>&nbsp;&nbsp;
		<label><input type="radio" name="useoptionyn" value="N" onClick="TnCheckOptionYN(this.form);" checked>�ɼǻ�����</label>
	</td>
</tr>
<!----- �ɼǱ��� DIV ----->
<tr id="opttype" style="display:none" height="40">
    <td width="15%" bgcolor="#DDDDFF">�ɼ� ����  :</td>
    <td width="85%" bgcolor="#FFFFFF">
        <input type="radio" name="optlevel" value="1" onClick="TnCheckOptionYN(this.form);" checked >���� �ɼ� (�ɼ� ���� 1��)
        <input type="radio" name="optlevel" value="2" onClick="TnCheckOptionYN(this.form);" disabled >���� �ɼ� (�ɼ� ���� �ִ� 3��) <font color="blue">�� ����Ư�������� ��ü����� ��츸 ���ð����մϴ�.</font>
    </td>
</tr>
<!----- ���� �ɼ� DIV ----->
<tr id="optlist" style="display:none" height="30">
    <td width="15%" bgcolor="#DDDDFF">�ɼ� ���� :</td>
  	<td width="85%" bgcolor="#FFFFFF">
      	<table width="500" border="0" cellspacing="0" cellpadding="0" class="a" >
      	<tr>
      	    <td width="100">�ɼ� ���и� :</td>
      	    <td width="400"><input type="text" name="optTypeNm" value="" size="20" maxlength="16" class="text" id="[off,off,off,off][�ɼ� ���и�]"></td>
      	</tr>
      	<tr>
      	    <td colspan="2">
              <select multiple name="realopt" class="select" style="width:400px;height:120px;"></select>
            </td>
        </tr>
        <tr>
            <td colspan="2">
              <input type="button" value="�⺻�ɼ��߰�" name="btnoptadd" class="button" onclick="popNormalOptionAdd();" >
              <input type="button" value="����ɼ��߰�" name="btnetcoptadd" class="button" onclick="popEtcOptionAdd();">
              <input type="button" value="���ÿɼǻ���" name="btnoptdel" class="button" onclick="delItemOptionAdd()" >
              <br><br>
              - �⺻�ɼ��߰� : ����, ������� �⺻������ ���ǵ� �ɼ��� �߰� �Ͻ� �� �ֽ��ϴ�.<br>
              - ����ɼ��߰� : �⺻�ɼǿ� ���ǵ��� ���� ��ǰ����ɼ��� �����Ͻ� �� �ֽ��ϴ�.<br>
              - ���ÿɼǻ��� : ���õ� �ɼ��� �����մϴ�.<br>
              - ���ǻ��� : �ѹ� ����� �ɼ��� <font color=red>������ �Ұ���</font>�մϴ�.<br>
              <br>
            </td>
        </tr>
        </table>
  	</td>
</tr>
<%
dim iMaxCols : iMaxCols = 3
dim iMaxRows : iMaxRows = 20
%>
<!----- ��Ƽ �ɼ� DIV ----->
<tr id="optlist2" style="display:none" height="30">
    <td width="15%" bgcolor="#DDDDFF">�ɼǼ��� :</td>
    <td width="85%" bgcolor="#FFFFFF">
        <table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
        <tr align="center"  bgcolor="#DDDDFF">
            <td width="100">�ɼǱ��и�</td>
            <% for j=0 to iMaxCols-1 %>
            <td>
                <input type="text" name="optionTypename<%= j+1 %>" value="" size="18" maxlength="16" class="text" id="[off,off,off,off][�ɼ� ���и�<%= j %>]">
            </td>
            <% Next %>
            <td width="80">(��Ͽ���)<br>����</td>
            <td width="80">(��Ͽ���)<br>������</td>
        </tr>
        <tr height="2" bgcolor="#FFFFFF">
            <td colspan="6"></td>
        </tr>
        <% for i=0 to iMaxRows-1 %>
        <tr align="center"  bgcolor="#FFFFFF">
            <td>�ɼǸ� <%= i+1 %></td>
            <% for j=0 to iMaxCols-1 %>
            <td>
                <input type="hidden" name="itemoption<%= j+1 %>" value="">
                <input type="text" name="optionName<%= j+1 %>" size="18" maxlength="18" class="text" id="[off,off,off,off][�ɼǸ�<%= i %><%= j %>]">
            </td>
            <% next %>
            <td>
                <% if i=0 then %>
                ����
                <% elseif i=1 then %>
                �Ķ�
                <% elseif i=2 then %>
                ���
                <% elseif i=3 then %>
                ������
                <% end if %>
            </td>
            <td>
                <% if i=0 then %>
                XL
                <% elseif i=1 then %>
                L
                <% elseif i=2 then %>
                S
                <% end if %>
            </td>
        </tr>
        <% next %>
        </table>
     </td>
</tr>

<!----- �⺻ ���� DIV ----->
<tr id="lyDFColor" height="30" style="display:;">
	<td colspan="2" bgcolor="#FFFFFF" style="padding:0px;">
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="0">
		<tr>
			<td width="15%" bgcolor="#DDDDFF">�⺻ ������ :</td>
			<td width="85%" bgcolor="#FFFFFF" style="border-left:1px solid <%= adminColor("tablebg") %>;"><%=FnSelectColorBar("",25)%></td>
		</tr>
		<tr>
			<td width="15%" rowspan="2" bgcolor="#DDDDFF" style="border-top:1px solid <%= adminColor("tablebg") %>;">���� ��ǰ�̹��� :</td>
			<td width="85%" bgcolor="#FFFFFF" style="border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;">
				<input type="file" size="40" name="imgDFColor" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text">
				<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgDFColor, 40, 1000, 1000)"> (����,1000X1000,<b><font color="red">jpg</font></b>)
			</td>
		</tr>
		<tr>
			<td width="85%" bgcolor="#FFFFFF" style="border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;">
		      - ���� �̹����� ������ ����� ���������� ��ǰ �⺻�̹����� ���˴ϴ�.
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<!-- 8.�������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>8.��������</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td width="15%" bgcolor="#DDDDFF" rowspan="2">�����Ǹű��� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	  <input type="radio" name="limityn" value="N" onClick="this.form.limitno.readOnly=true; this.form.limitno.value=''; this.form.limitno.className='text_ro';" checked>�������Ǹ�&nbsp;&nbsp;
  	  <input type="radio" name="limityn" value="Y" onClick="this.form.limitno.readOnly=false; this.form.limitno.className='text';">�����Ǹ�
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">�������� :</td>
  	<td width="35%" bgcolor="#FFFFFF" >
      <input type="text" name="limitno" maxlength="32" size="8" readonly class="text_ro" id="[off,on,off,off][��������]">(��)
  	</td>
  </tr>
  <tr>
  	<td colspan="3" bgcolor="#FFFFFF"><font color="red">** �ɼ��� �ִ°�� �ɼǺ��� ���������� �ϰ� �����˴ϴ�.(���������� ����� ��������)</font></td>
  </tr>
</table>

<!-- 9.��ǰ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>9.��ǰ����</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="usinghtml" value="N" checked >�Ϲ�TEXT
      <input type="radio" name="usinghtml" value="H">TEXT+HTML
      <input type="radio" name="usinghtml" value="Y">HTML���
      <br>
      <textarea name="itemcontent" rows="18" class="textarea" style="width:100%" id="[on,off,off,off][��ǰ����]"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ��� ���ǻ��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="ordercomment" rows="5" cols="90" class="textarea" id="[off,off,off,off][���ǻ���]"></textarea><br>
      <font color="red">Ư���� ��۱Ⱓ�̳� �ֹ��� Ȯ���ؾ߸� �ϴ� ����</font>�� �Է��Ͻø� ���Ҹ��̳� ȯ���� ���ϼ� �ֽ��ϴ�.
  	</td>
  </tr>
</table>

<!-- 10.�̹������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left" style="padding-bottom:5px;">
          <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>10.�̹�������</strong>
			<br>- �ٹ����ٿ��� �̹����� ����� ��쿡�� �ʼ��׸��� �⺻�̹����� �Է��Ͻñ� �ٶ��ϴ�.
			<br>- �̹����� <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> ���� �ø��� �� �ֽ��ϴ�.
			<br>&nbsp;&nbsp;(�̹�������� <font color=red>���μ������� ������</font>�� �԰ݿ� ���� �ʰ� ������ּ���. �԰��ʰ��� ����� ���� �ʽ��ϴ�.)
			<br>- <font color=red>����޿��� Save For Web����, Optimizeüũ, ������ 80%����</font>�� ����� �� �÷��ֽñ� �ٶ��ϴ�.
        </td>
        <td align="right" valign="bottom" style="padding-bottom:5px;">
        	<a href="javascript:PopImageInformation()"><b><font color=red>[�ʵ�]�̹��� ��Ͽ��</font></b> <img src="/images/icon_help.gif" border="0" align="absbottom"></a>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �⺻�̹��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgbasic" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgbasic,40, 1000, 1000)"> (<font color=red>�ʼ�</font>,1000X1000,<b><font color="red">jpg</font></b>)
      <!-- //������// <br><input type="checkbox" name="regimg"> ������̹������ - �̹����� <font color=red>���߿� ���</font>�Ұ�쿡�� ������̹�������� üũ�ϼ���.-->
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ ����(����)�̹��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgmask" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmask,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd1" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd1,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd2" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd2, 40, 1000, 1000)"> (����,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd3" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd3, 40, 1000, 1000)"> (����,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���4 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd4" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd4, 40, 1000, 1000)"> (����,1000X1000,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���5 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd5" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd5,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
   	</td>
  </tr>
  <tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgmain" onchange="CheckImage(this, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmain, 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #2:</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgmain2" onchange="CheckImage(this, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmain2, 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #3:</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgmain3" onchange="CheckImage(this, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40);" class="text" size="40">
      <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmain3, 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="�����ϱ�" class="button_s" onClick="SubmitSave()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
</form>
<script language='javascript'>
function getOnload(){
    EnDisableFlowerShop();
}
window.onload = getOnload;
</script>
<div style="padding:5px;text-align:right;font-size:8pt">Ver1.0  lastupdate: 2013.12.16 </div>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->