<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->

<%
CONST CBASIC_IMG_MAXSIZE = 180   'KB
CONST CMAIN_IMG_MAXSIZE = 500   'KB

Dim npartner
set npartner = new CPartnerUser
npartner.FRectDesignerID = session("ssBctID")
npartner.GetOnePartnerNUser


dim i,j,k 
%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language="javascript" >
function UseTemplate() {
	window.open("/common/pop_basic_item_info_list.asp", "UseTemplate", "width=700, height=450, scrollbars=yes, resizable=yes");
}

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
function InsertOption(ft, fv) {
	var frm = document.itemreg;

	//�ɼǰ��� �������� ������ skip ,����ɼ��ΰ�� ����
	if (fv!="0000"){
		for (i=0;i<frm.realopt.length;i++){
			if (frm.realopt[i].value==fv){
				return;
			}
		}
	}
	frm.elements['realopt'].options[frm.realopt.options.length] = new Option(ft, fv);
}

//2008�� ��
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

// �̹���ǥ��
function ClearImage(img) {
    var e = eval("itemreg." + img);

	if (img == "imgbasic") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg');\" size='40'>";
    } else if (img == "imgmain") {
       e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');\" size='40'>";
    } else if (img == "imgadd1") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');\" size='40'>";
    } else {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');\" size='40'>";
    }

    e = eval("document.all.div" + img);
    e.style.display = "none";

    if (img == "imgbasic") {
        e = eval("itemreg.basic");
        e.value = "del";
    } else if (img == "imgadd1") {
        e = eval("itemreg.add1");
        e.value = "del";
    } else if (img == "imgadd2") {
        e = eval("itemreg.add2");
        e.value = "del";
    } else if (img == "imgadd3") {
        e = eval("itemreg.add3");
        e.value = "del";
    } else if (img == "imgadd4") {
        e = eval("itemreg.add4");
        e.value = "del";
    } else if (img == "imgadd5") {
        e = eval("itemreg.add5");
        e.value = "del";
    } else if (img == "imgmain") {
        e = eval("itemreg.main");
        e.value = "del";
    }
}

function ShowImage(img) {
	var e = eval("document.all.div" + img);
    e.style.display = "";

    var filename;
    e = eval("itemreg." + img );
    filename = e.value;

	eval("document.all." + img + "_img").src=filename;
    //document.getElementById(img).src=filename;


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

function pause(numberMillis) {
     var now = new Date();
     var exitTime = now.getTime() + numberMillis;


     while (true) {
          now = new Date();
          if (now.getTime() > exitTime)
              return;
     }
}

function CheckImage(img, filesize, imagewidth, imageheight, extname)
{
    var preview;
    var e;
    var ext;
    var filename;

    e = eval("itemreg." + img);
    filename = e.value;

    e = eval("itemreg." + img);
    if (e.value == "") { return false; }

	ShowImage(img);

    if (CheckExtension(filename, extname) != true) {
        alert("�̹���ȭ���� ������ ȭ�ϸ� ����ϼ���.[" + extname + "]");
        ClearImage(img);
        return false;
    }
    
    try{
        // iframe �ӿ� �̹����� �ְ�, ������/ũ�⸦ üũ�Ѵ�.
        document.imgpreview.document.getElementById("imgpreview").src = filename;
        // �ð����̷� �̹��� �ε����� �Ѿ�� ����
        preview = document.imgpreview.document.getElementById("imgpreview");
    
        if(preview.fileSize > (filesize * 1024)){
            alert("���ϻ������ " + filesize + "Kbyte�� �ѱ�� �� �����ϴ�.");
            ClearImage(img);
            return false;
        }
    
        if(preview.width > (imagewidth)){
            alert("�������� " + imagewidth + "�ȼ��� �ѱ�� �� �����ϴ�.");
            ClearImage(img);
            return false;
        }
    
        if(preview.height > (imageheight)){
            alert("�������� " + imageheight + "�ȼ��� �ѱ�� �� �����ϴ�.");
            ClearImage(img);
            return false;
        }
    }catch(ex){
        // nothing;
    }

    if (img == "imgbasic") {
        e = eval("itemreg.basic");
        e.value = "";
    } else if (img == "imgadd1") {
        e = eval("itemreg.add1");
        e.value = "";
    } else if (img == "imgadd2") {
        e = eval("itemreg.add2");
        e.value = "";
    } else if (img == "imgadd3") {
        e = eval("itemreg.add3");
        e.value = "";
    } else if (img == "imgadd4") {
        e = eval("itemreg.add4");
        e.value = "";
    } else if (img == "imgadd5") {
        e = eval("itemreg.add5");
        e.value = "";
    } else if (img == "imgmain") {
        e = eval("itemreg.main");
        e.value = "";
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


// �����ϱ�
function SubmitSave() {
//alert('���� ���� �۾� ������ ��ǰ ���/ ������ �Ұ��մϴ�.');
//return;
	
	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
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
	
	if (itemreg.itemsize.value!=''){
		if (itemreg.unit.value!=''){
			itemreg.itemsize.value=itemreg.itemsize.value + '(' + itemreg.unit.value + ')';
		}
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

    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("���������� �Է����ּ���!");
        itemreg.limitno.focus();
        return;
    }

	
	
    if (itemreg.imgbasic.value == "") {
        // alert("�⺻�̹����� �ʼ��Դϴ�.");
        // return;
    } else {
        if (CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg') != true) {
            return;
        }
    }

    if (itemreg.imgadd1.value != "") {
        if (CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd4.value != "") {
        if (CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd5.value != "") {
        if (CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgmain.value != "") {
        if (CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif') != true) {
            return;
        }
    }

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
	}
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
		}
	}
	//else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
	else if(frm.deliverytype[1].checked){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("����Ư�� ������ �����̳� Ư���� ���\n��۱�����  ��ü������� ���� �Ͻ� �� �����ϴ�!!!\n����Ư�������� Ȯ�����ּ���!!");
			frm.mwdiv[2].checked=true;
		}
	}
}

function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // �ɼǻ��
        
        opttype.style.display="inline";
        
        if (frm.optlevel[1].checked==true){
            optlist.style.display ="none";
            optlist2.style.display ="inline";
        }else{
            optlist.style.display="inline";
            optlist2.style.display="none";
        }
        
	} else {
	    // �ɼǾ���
	    while (frm.realopt.length > 0) {
	        frm.realopt.options[0] = null;
        }
        opttype.style.display="none";
        document.all.optlist2.style.display="none";
		document.all.optlist.style.display="none";

        frm.itemoptioncode2.value = "";
        frm.itemoptioncode3.value = "";
    }
}

// �̹��� �˸�â
function PopImageInformation(){
	window.open("itemreg_info_win.asp","PopImageInformation","width=920,height=600,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no");
}

function EnDisableFlowerShop(){
    
    var frm = document.itemreg;
    if ((frm.cd1.value=="110")&&(frm.cd2.value=="060")){
        frm.deliverarea[1].disabled = false;
        frm.deliverarea[2].disabled = false;
        
        frm.deliverfixday.disabled = false;
    }else{
        frm.deliverarea[1].disabled = true;
        frm.deliverarea[2].disabled = true;
        
        frm.deliverfixday.disabled = true;
        frm.deliverfixday.checked = false;
    }
}

function ClearVal(comp){
    comp.value = "";
}
</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>�⺻����</strong>
        </td>
        <td align="right">
          <input type="button" value="�⺻Ʋ����" onClick="UseTemplate();">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/WaitUpcheItemRegister_Process.asp" onsubmit="return false;" enctype="multipart/form-data">
  <input type="hidden" name="itemoptioncode2">
  <input type="hidden" name="itemoptioncode3">
  <input type="hidden" name="designerid" value="<%= session("ssBctID") %>">
  <input type="hidden" name="defultmargine" value="<%= npartner.FOneItem.Fdefaultmargine %>">
  <input type="hidden" name="defaultmaeipdiv" value="<%= npartner.FOneItem.Fmaeipdiv %>">
  <input type="hidden" name="defaultFreeBeasongLimit" value="<%= npartner.FOneItem.FdefaultFreeBeasongLimit %>">
  <input type="hidden" name="defaultDeliverPay" value="<%= npartner.FOneItem.FdefaultDeliverPay %>">
  <input type="hidden" name="defaultDeliveryType" value="<%= npartner.FOneItem.FdefaultDeliveryType %>">
  
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">ī�װ� ���� :</td>
    <input type="hidden" name="cd1" value="">
    <input type="hidden" name="cd2" value="">
    <input type="hidden" name="cd3" value="">
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="cd1_name" value="" id="[on,off,off,off][ī�װ�]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" value="" id="[on,off,off,off][ī�װ�]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" value="" id="[on,off,off,off][ī�װ�]" size="20" readonly style="background-color:#E6E6E6">

      <input type="button" value="ī�װ� ����" onclick="editCategory(itemreg.cd1.value,itemreg.cd2.value,itemreg.cd3.value);">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="itemdiv" value="01" checked>�Ϲݻ�ǰ
      <input type="radio" name="itemdiv" value="06">�ֹ����ۻ�ǰ
      <font color="red">(�ֹ����� �޼����� �ʿ��� ���, ������� ����û �̴ϼ��� �־��ٰ��)</font>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemname" maxlength="64" size="50" id="[on,off,off,off][��ǰ��]">&nbsp;
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsource" maxlength="64" size="50" id="[on,off,off,off][��ǰ����]">&nbsp;(ex:�ö�ƽ,����,��,...)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsize" maxlength="64" size="50" id="[on,off,off,off][��ǰ������]">
      	<select name="unit">
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
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="sourcearea" maxlength="64" size="25" id="[on,off,off,off][������]">&nbsp;(ex:�ѱ�,�߱�,�߱�OEM,�Ϻ�...)
      <br>( ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="makername" maxlength="32" size="25" id="[on,off,off,off][������]">&nbsp;(������ü��)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�˻�Ű���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="keywords" maxlength="50" size="50" id="[on,off,off,off][�˻�Ű����]">&nbsp;(�޸��α��� ex: Ŀ��,Ƽ����,����)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ü��ǰ�ڵ� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	    <input type="text" name="upchemanagecode" value="" size="20" maxlength="32" id="[off,off,off,off][��ü��ǰ�ڵ�]">
  	    (��ü���� �����ϴ� �ڵ� �ִ� 32�� - ����/���ڸ� ����)
  	</td>
  </tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>��������</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">

  <tr align="left">
    <td width="15%" bgcolor="#DDDDFF">����, �鼼 ���� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="radio" name="vatinclude" value="Y" checked onclick="TnGoClear(this.form);">����
      <input type="radio" name="vatinclude" value="N" onclick="TnGoClear(this.form);">�鼼
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">�⺻ ���� ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="margin" maxlength="32" size="5" id="[off,off,off,off][����]" value="<% =npartner.FOneItem.Fdefaultmargine %>" readonly style="background-color:#E6E6E6;">%
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ǸŰ�(�Һ��ڰ�) :</td>
  	<input type="hidden" name="sellvat">
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="sellcash" size="12" id="[on,on,off,off][�Һ��ڰ�]" onKeyUp="CalcuAuto(itemreg);" maxlength="7">��
      <!--<input type="button" value="���ް� �ڵ����" onclick="CalcuAuto(itemreg);">-->
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">���ް� :</td>
  	<input type="hidden" name="buyvat">
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="buycash" maxlength="16" size="12" id="[on,on,off,off][���ް�]" >��
      (<b>�ΰ��� ���԰�</b>)
  	</td>
  </tr>
  <input type="hidden" name="mileage" value="0">
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>�Ǹ�����</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">����Ư������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="mwdiv" value="M" checked onclick="TnCheckUpcheYN(this.form);">����
      <input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);">Ư��
      <input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);">��ü���
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��۱��� :</td>
  	<td width="85%" bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="deliverytype" value="1" checked  onclick="TnCheckUpcheDeliverYN(this.form);">�ٹ����ٹ��&nbsp;
      <input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);">��ü(����)���&nbsp;
      <input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);">�ٹ����ٹ�����&nbsp;
      <input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);">��ü���ǹ��(���� ��ۺ�ΰ�)&nbsp;
      <input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);">��ü���ҹ��
     
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">������� :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="deliverarea" value="" checked>�����ù�(�Ϲ�)&nbsp;
      <input type="radio" name="deliverarea" value="C" disabled >�����ǹ��&nbsp;
      <input type="radio" name="deliverarea" value="S" disabled >������&nbsp;
      <input type="checkbox" name="deliverfixday" value="C" disabled >�ö��������
      <br>&nbsp;<font color="red">(�ö�� ��ǰ�� ��츸 �����ǹ��, ������, �ö�������� �ɼ��� ��밡���մϴ�.)</font>
  	</td>
  </tr>
  <!-- ������
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�Ǹ�����(����)�� :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="sellEndDate" id="[off,off,off,off][�Ǹ�����(����)��]"  size="10" value="" > 
  	    <a href="javascript:calendarOpen(itemreg.sellEndDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
  	    <a href="javascript:ClearVal(itemreg.sellEndDate);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
  	</td>
  </tr>
  -->
  <input type="hidden" name="pojangok" value="N">
  <input type="hidden" name="sellyn" value="N">
  <input type="hidden" name="dispyn" value="N">
  <input type="hidden" name="isusing" value="Y">
</table>
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>��ǰ����</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">������ ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="usinghtml" value="N" checked >�Ϲ�TEXT
      <input type="radio" name="usinghtml" value="H">TEXT+HTML
      <input type="radio" name="usinghtml" value="Y">HTML���
      <br>
      <textarea name="itemcontent" rows="10" cols="80" id="[on,off,off,off][�����ۼ���]"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ��� ���ǻ��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="ordercomment" rows="5" cols="80" id="[off,off,off,off][���ǻ���]"></textarea><br>
      <font color="red">Ư���� ��۱Ⱓ�̳� �ֹ��� Ȯ���ؾ߸� �ϴ� ����</font>�� �Է��Ͻø� ���Ҹ��̳� ȯ���� ���ϼ� �ֽ��ϴ�.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ü�ڸ�Ʈ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="designercomment" size="50" maxlength="40" id="[off,off,off,off][��ü�ڸ�Ʈ]"><br>
      ��ǰ������ ���丮�� ��̳� �̾߱⸦ �����ּ���...
  	</td>
  </tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>�ɼ�����/��������</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ɼǱ��� :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="useoptionyn" value="Y" onClick="TnCheckOptionYN(this.form);">�ɼǻ����&nbsp;&nbsp;
      <input type="radio" name="useoptionyn" value="N" onClick="TnCheckOptionYN(this.form);" checked>�ɼǻ�����
  	</td>
  </tr>

  <tr align="left">
  	<td width="15%" bgcolor="#DDDDFF" rowspan="2">�����Ǹű��� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	  <input type="radio" name="limityn" value="N" onClick="this.form.limitno.readonly=true; this.form.limitno.value=''; this.form.limitno.style.background='#E6E6E6'; this.form.limitno.readOnly=true" checked>�������Ǹ�&nbsp;&nbsp;
  	  <input type="radio" name="limityn" value="Y" onClick="this.form.limitno.readonly=false;this.form.limitno.style.background='#FFFFFF'; this.form.limitno.readOnly=false">�����Ǹ�
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">�������� :</td>
  	<td width="35%" bgcolor="#FFFFFF" >
      <input type="text" name="limitno" maxlength="32" size="8" readonly style="background-color:#E6E6E6;" id="[off,on,off,off][��������]">(��)
  	</td>
  </tr>
  <tr>
  	<td colspan="3" bgcolor="#FFFFFF"><font color="red">** �ɼ��� �ִ°�� �ɼǺ��� ���������� �ϰ� �����˴ϴ�.(���������� ����� ��������)</font></td>
  </tr>
</table>

<div id="opttype" style="display:none">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr height="40">
    <td width="15%" bgcolor="#DDDDFF">�ɼ� ����  :</td>
    <td width="85%" bgcolor="#FFFFFF">
        <input type="radio" name="optlevel" value="1" onClick="TnCheckOptionYN(this.form);" checked >���� �ɼ� (�ɼ� ���� 1��)
        <input type="radio" name="optlevel" value="2" onClick="TnCheckOptionYN(this.form);">���� �ɼ� (�ɼ� ���� �ִ� 3��)
    </td>
  </tr>
</table>
</div>

<div id="optlist" style="display:none" >
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr>
    <td width="15%" bgcolor="#DDDDFF">�ɼ� ���� :</td>
  	<td width="85%" bgcolor="#FFFFFF" colspan="3">
      	<table width="500" border="0" cellspacing="0" cellpadding="0" class="a" >
      	<tr>
      	    <td width="100">�ɼ� ���и� :</td>
      	    <td width="400"><input type="text" name="optTypeNm" value="" size="20" maxlength="20" id="[off,off,off,off][�ɼ� ���и�]"></td>
      	</tr>
      	<tr>
      	    <td colspan="2">
              <select multiple name="realopt" size="10" style="width:400">
              </select>
              <br>
            </td>
        </tr>
        <tr>
            <td colspan="2">
              <input type="button" value="�⺻�ɼ��߰�" name="btnoptadd" onclick="popNormalOptionAdd();" >
              <input type="button" value="����ɼ��߰�" name="btnetcoptadd" onclick="popEtcOptionAdd();">
              <input type="button" value="���ÿɼǻ���" name="btnoptdel" onclick="delItemOptionAdd()" >
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
</table>
</div>


<%
dim iMaxCols : iMaxCols = 3
dim iMaxRows : iMaxRows = 9
%>
<div id="optlist2" style="display:none">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr>
    <td width="15%" bgcolor="#DDDDFF">�ɼǼ��� :</td>
    <td width="85%" bgcolor="#FFFFFF" colspan="3">
        <table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
        <tr align="center"  bgcolor="#DDDDFF">
            <td width="100">�ɼǱ��и�</td>
            <% for j=0 to iMaxCols-1 %>
            <td>
                <input type="text" name="optionTypename<%= j+1 %>" value="" size="18" maxlength="20" id="[off,off,off,off][�ɼ� ���и�<%= j %>]">
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
<td><input type="hidden" name="itemoption<%= j+1 %>" value="">
    <input type="text" name="optionName<%= j+1 %>" size="18" maxlength="20" id="[off,off,off,off][�ɼǸ�<%= i %><%= j %>]"></td>
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
 </table>
</div>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="100">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
			<img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>�̹������</strong>
			<br>- �ٹ����ٿ��� �̹����� ����� ��� ���� �Է����� ���ñ� �ٶ��ϴ�.
			<br>- �̹����� <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> ���� �ø��� �� �ֽ��ϴ�.
			<br>&nbsp;&nbsp;(�̹�������� <font color=red>���μ������� ������</font>�� �԰ݿ� ���� �ʰ� ������ּ���. �԰��ʰ��� ����� ���� �ʽ��ϴ�.)
			<br>- <font color=red>����޿��� Save For Web����, Optimizeüũ, ������ 80%����</font>�� ����� �� �÷��ֽñ� �ٶ��ϴ�.
        </td>
        <td align="right" valign="bottom">
        	<a href="javascript:PopImageInformation()"><b><font color=red>[�ʵ�]�̹��� ��Ͽ��</font></b> <img src="/images/icon_help.gif" border="0" align="absbottom"></a>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<iframe name="imgpreview" src="iframe_imagepreview.asp" frameborder="0" width="0" height="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�⺻�̹��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgbasic" onchange="CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg');" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgbasic')"><input type="hidden" name="basic"> (<font color=red>�ʼ�</font>,400X400,<b><font color="red">jpg</font></b>)
	  <div id="divimgbasic" style="display:none;">
      <table width="400" height="400" >
        <tr>
          <td>
          	<img id="imgbasic_img" src=""> 
          </td>
        </tr>
      </table>
      </div>
  	</td>
  </tr>
  <tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd1" onchange="CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd1')"><input type="hidden" name="add1"> (����,400X400,jpg,gif)
	  <div id="divimgadd1" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd1_img" src=""></td>
        </tr>
        
      </table>
      </div>
  	</td>
  </tr>
  <tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <div id="divimgadd2" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd2_img" src=""></td>
        </tr>
       
      </table>
      </div>

      <input type="file" name="imgadd2" onchange="CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd2')"><input type="hidden" name="add2"> (����,400X400,jpg,gif)
  	</td>
  </tr>
  <tr height="2" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
	  <div id="divimgadd3" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd3_img" src=""></td>
        </tr>
        
      </table>
      </div>

      <input type="file" name="imgadd3" onchange="CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd3')"><input type="hidden" name="add3"> (����,400X400,jpg,gif)

  	</td>
  </tr>
  <tr height="2" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���4 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">

      <input type="file" name="imgadd4" onchange="CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd4')"><input type="hidden" name="add4"> (����,400X400,jpg,gif)

	  <div id="divimgadd4" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd4_img" src=""></td>
        </tr>
        
      </table>
      </div>
  	</td>
  </tr>
  <tr height="2" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���5 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd5" onchange="CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd5')"><input type="hidden" name="add5"> (����,400X400,jpg,gif)

      <div id="divimgadd5" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd5_img" src=""></td>
        </tr>
       
      </table>
      </div>
   	</td>
  </tr>
  <tr height="2" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgmain" onchange="CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgmain')"><input type="hidden" name="main"> (����,600X2000, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
	  <div id="divimgmain" style="display:none;">
      <table width="400" height="400">
        <tr>
          <td>
          <img id="imgmain_img" src="">
          </td>
        </tr>
      </table>
      </div>
  	</td>
  </tr>
  </form>
</table>


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="�����ϱ�" onClick="SubmitSave()">
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

<script language='javascript'>
function getOnload(){
    EnDisableFlowerShop();
}
window.onload = getOnload;
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->