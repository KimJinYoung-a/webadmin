$(function(){
	// �ε��� ��ǰ�Ӽ� ���� ���
	printItemAttribute();
});

function printItemAttribute() {
	var arrDispCd="";
	$("input[name='catecode']").each(function(i){
		if(i>0) arrDispCd += ",";
		arrDispCd += $(this).val();
	});
	$.ajax({
		url: "/common/module/act_ItemAttribSelect.asp?itemid=<%=itemid%>&arrDispCate="+arrDispCd,
		cache: false,
		success: function(message) {
			$("#lyrItemAttribAdd").empty().append(message);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

function UseTemplate() {
	window.open("/common/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
}

function popMultiLangEdit(iid) {
	window.open("/common/item/pop_MultiLangItemCont.asp?itemid="+iid+"&lang=EN", "multiLang_win", "width=600, height=500, scrollbars=yes, resizable=yes");
}

// ============================================================================
// ��ü �⺻���� �ڵ��Է� - �귣�庯��ø�.
function TnDesignerNMargineAppl(){
	var varArray;
	varArray = document.itemreg.marginData.value.split(',');

	document.itemreg.designerid.value = document.itemreg.designer.value;
	document.itemreg.margin.value = varArray[0];
    
    document.itemreg.defaultmargin.value            = varArray[0];  //��ü�⺻����.
	document.itemreg.defaultmaeipdiv.value          = varArray[1];
    document.itemreg.defaultFreeBeasongLimit.value  = varArray[2];
    document.itemreg.defaultDeliverPay.value        = varArray[3];
    document.itemreg.defaultDeliveryType.value      = varArray[4];
    
    if(document.itemreg.mwdiv.length>0){
        if (document.itemreg.defaultmaeipdiv.value=="M"){
            document.itemreg.mwdiv[0].checked = true; //����
        }else if (document.itemreg.defaultmaeipdiv.value=="W"){
            document.itemreg.mwdiv[1].checked = true; //��Ź
        }else if (document.itemreg.defaultmaeipdiv.value=="U"){
            document.itemreg.mwdiv[2].checked = true; //��ü
        }
    }else{
        document.itemreg.mwdiv.value=document.itemreg.defaultmaeipdiv.value;
    }
    
    TnCheckUpcheYN(document.itemreg);
}

function CalcuAuto(frm){
	var isvatinclude, imileage;
	var isellcash, ibuycash, imargin;
	var isailprice, isailsuplycash, isailpricevat, isailsuplycashvat, isailmargin;

    isvatinclude = frm.vatinclude[0].checked;

	if (frm.sailyn[0].checked == true) {
	    // ���󰡰�
	    isellcash = frm.sellcash.value;
	    imargin = frm.margin.value;

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
    		ibuycash = isellcash - Math.round(isellcash*imargin/100);   //parseInt-> round�� ���� 
			imileage = parseInt(isellcash*0.005) ;
    	}else{
    		ibuycash = isellcash - Math.round(isellcash*imargin/100);   //parseInt-> round�� ���� 
			imileage = parseInt(isellcash*0.005) ;
    	}

    	frm.buycash.value = ibuycash;
    	frm.mileage.value = imileage;
	} else {
	    // ���ϰ���
	    isailprice = frm.sailprice.value;
	    isailmargin = frm.sailmargin.value;

    	if (isailmargin.length<1){
    		alert('���ϸ����� �Է��ϼ���.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (isailprice.length<1){
    		alert('�����ǸŰ��� �Է��ϼ���.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (!IsDouble(isailmargin)){
    		alert('���ϸ����� ���ڷ� �Է��ϼ���.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (!IsDigit(isailprice)){
    		alert('�����ǸŰ��� ���ڷ� �Է��ϼ���.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (isvatinclude==true){
    		isailpricevat = parseInt(parseInt(1/11 * parseInt(isailprice)));
    		isailsuplycash = isailprice - Math.round(isailprice*isailmargin/100);         //parseInt-> round�� ���� 
    		isailsuplycashvat = parseInt(parseInt(1/11 * parseInt(isailsuplycash)));
			imileage = parseInt(isailprice*0.005) ;
    	}else{
    		isailpricevat = 0;
    		isailsuplycash = isailprice - Math.round(isailprice*isailmargin/100);         //parseInt-> round�� ���� 
    		isailsuplycashvat = 0;
			imileage = parseInt(isailprice*0.005) ;
    	}

    	frm.sailpricevat.value = isailpricevat;
    	frm.sailsuplycash.value = isailsuplycash;
    	frm.sailsuplycashvat.value = isailsuplycashvat;
    	frm.mileage.value = imileage;
    }

	//������ ���
	if (frm.sailyn[0].checked == true) {
		document.getElementById("lyrPct").innerHTML = "";
	} else {
		isellcash = frm.sellcash.value;
		isailprice = frm.sailprice.value;
		var isalePercent = parseInt(Math.round((isellcash-isailprice)/isellcash*1000))/10;
		document.getElementById("lyrPct").innerHTML = "������: <font color='#EE0000'><strong>" + isalePercent + "%</strong></font>";
	}
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
}

// ============================================================================
// �ɼǼ���
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=800,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function doEditItemOption(optioncnt, optlimitno, optlimitsold, optlimitstock) {
    // �ɼ�â���� ����â����
    itemreg.optioncnt.value = optioncnt;

    itemreg.limitno.value = optlimitno;
    itemreg.limitsold.value = optlimitsold;
    itemreg.limitstock.value = optlimitstock;
}

function popNormalOptionAdd() {
	popwin = window.open('/common/module/normalitemoptionadd.asp' ,'popNormalOptionAdd','width=800,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// �ɼ��� �߰��Ѵ�
function InsertOption(ft, fv) {
	var frm = document.itemreg;
	frm.elements['realopt'].options[frm.realopt.options.length] = new Option(ft, fv);
}

// ���õ� �ɼ� ����
function delItemOptionAdd()
{
	var frm = document.itemreg;
	var sidx = frm.realopt.options.selectedIndex;

	if(sidx<0)
		alert("������ �ɼ��� �������ֽʿ�.");
	else
	{
		frm.realopt.options[sidx] = null;
	}
}


// ============================================================================
// �̹���ǥ��
function ClearImage(img,fsize,wd,ht) {
	img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";

    document.getElementById("div"+ img.name).style.display = "none";

	var e = eval("itemreg."+img.name.substr(3,img.name.length));
	e.value = "del";
}

function ClearImage2(img,fsize,wd,ht,num) {
	var imgcnt = $('input[name="addimgname"]').length;
    img.outerHTML = "<input type='file' name='" + img.name + "' onchange=\"CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +", "+num+");\" class='text' size='"+ fsize +"'>";
	$("#divaddimgname"+(num+1)+"").hide();
	
	if(imgcnt > 1){
    	document.itemreg.addimgdel[num].value = "del";
    }else{
    	document.itemreg.addimgdel.value = "del";
    }
}

function oldClearImage(img,fsize,wd,ht) {
	$("#divimg"+img+"").hide();
	$("input[name='"+img+"']").val("del");
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
        alert("�̹��������� ������ ���ϸ� ����ϼ���.[" + extname + "]");
        ClearImage(img,fsize, imagewidth, imageheight, num);
        return false;
    }

	var e = eval("itemreg."+img.name.substr(3,img.name.length));
	e.value = "";

    return true;
}

function CheckImage2(img, filesize, imagewidth, imageheight, extname, fsize, num)
{
    var ext;
    var filename;
    var imgcnt = $('input[name="addimgname"]').length;

	filename = img.value;
	if (img.value == "") { return false; }

    if (CheckExtension(filename, extname) != true) {
        alert("�̹��������� ������ ���ϸ� ����ϼ���.[" + extname + "]");
        ClearImage2(img,fsize, imagewidth, imageheight, num);
        return false;
    }

	if(imgcnt > 1){
    	document.itemreg.addimgdel[num].value = "";
    }else{
    	document.itemreg.addimgdel.value = "";
    }

    return true;
}


// ============================================================================
// �����ϱ�
function SubmitSave() {
	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
		return;
	}

	if (!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("[�⺻] ���� ī�װ��� �����ϼ���.\n�� [�߰�] ���� ī�װ��� ���� �� �����ϴ�.");
		return;
	}

    // ī�װ� �������� �˻�
	if(tbl_Category.rows.length>0)	{
		if(tbl_Category.rows.length>1)	{
			var chk=0;
			for(l=0;l<document.all.cate_div.length;l++)	{
				if(document.all.cate_div[l].value=="D") chk++;
			}
			if(chk==0) {
				alert("ī�װ��� �⺻ ī�װ��� �������ּ���.\n�ر⺻ ī�װ��� �ʼ��׸��Դϴ�.");
				return;
			} else if(chk>1) {
				alert("ī�װ��� �⺻ ī�װ��� �Ѱ��� �������ּ���.");
				return;
			}
		}
		else {
			if(document.all.cate_div.length){
				if(document.all.cate_div[0].value!="D") {
					alert("ī�װ��� �⺻ ī�װ��� �������ּ���.\n�ر⺻ ī�װ��� �ʼ��׸��Դϴ�.");
					return;
				}
			} else {
				if(document.all.cate_div.value!="D") {
					alert("ī�װ��� �⺻ ī�װ��� �������ּ���.\n�ر⺻ ī�װ��� �ʼ��׸��Դϴ�.");
					return;
				}
			}
		}
	} else {
		alert("ī�װ��� �������ּ���.");
		return;
	}
	
	//��۱��� üũ =========================================================================
    //��ü ���ǹ��
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[3].checked){
            alert('��� ������ Ȯ�����ּ���. ������� ��ü�� �ƴմϴ�.');
            return;
        }
    }
    
    //��ü���ҹ�� : ���ǹ�۵� ���Ҽ�������
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype[4].checked)){
        alert('��� ������ Ȯ�����ּ���. [��ü ���ҹ��,��ü ���ǹ��] ��ü�� �ƴմϴ�.');
        itemreg.deliverytype[4].focus();
        return;
    }
    
    if ((itemreg.deliverytype[1].checked)||(itemreg.deliverytype[3].checked)||(itemreg.deliverytype[4].checked)){
        if(document.itemreg.mwdiv.length>0){
            if ((document.itemreg.mwdiv[0].checked)||(document.itemreg.mwdiv[1].checked)){
                alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
                return;
            }
        }else{
            if ((document.itemreg.mwdiv.value=="M")||(document.itemreg.mwdiv.value=="W")){
	            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
	            return;
	        }
        }
        
       // if (itemreg.deliverOverseas.checked){
       //     alert('�ٹ����� ����� ��쿡�� �ؿܹ���� �Ͻ� �� �ֽ��ϴ�.');
       //     return;
       // }
    }
    
    if(document.itemreg.mwdiv.length>0){
        if (document.itemreg.mwdiv[2].checked){
            if ((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
                alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
                return;
            }
        }
    }else{
        if (document.itemreg.mwdiv.value=="U"){
	        if ((document.itemreg.deliverytype[0].checked)||(document.itemreg.deliverytype[2].checked)){
	            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
	            return;
	        }
	    }
    }
    
    //��ü��۸� �ֹ����� ����.
    if(document.itemreg.mwdiv.length>0){
        if ((!document.itemreg.mwdiv[2].checked)&&(document.itemreg.itemdiv[1].checked)){
            alert('�ֹ� ���ۻ�ǰ�� ��ü����ΰ�츸 �����մϴ�.');
            return;
        }
    }else{
        if ((document.itemreg.mwdiv.value!="U")&&(document.itemreg.itemdiv[1].checked)){
            alert('�ֹ� ���ۻ�ǰ�� ��ü����ΰ�츸 �����մϴ�.');
            return;
        }
    }
    
	if(document.itemreg.deliverfixday[1].checked) {
		if(document.itemreg.freight_min.value<=0||document.itemreg.freight_max.value<=0) {
            alert('ȭ����� ����� �Է����ּ���.');
            document.itemreg.freight_min.focus();
            return;
		}
	}

    //==================================================================================
    
   //-------------------------------------------------------------------------------- 2014.02.14 ������ �߰�
	//1.����ڰ� [���̰�����] �� ���, ���Ի�ǰ ��� �Ұ� / ��ü,��Ź ��ǰ�� ��ϰ���
	if(document.itemreg.mwdiv.length>0){
    	if((document.itemreg.jungsangubun.value =="���̰���")&&(document.itemreg.mwdiv[0].checked)){
    		alert("����ڰ� [���̰�����]�� ���, [����]��ǰ�� ��ϺҰ����մϴ�. \n[��Ź],[��ü���]��ǰ�� ��ϰ����մϴ�. ");
    		document.itemreg.mwdiv[0].focus();
    		return;
    	}
    }else{
        if((document.itemreg.jungsangubun.value =="���̰���")&&(document.itemreg.mwdiv.value=="M")){
    		alert("����ڰ� [���̰�����]�� ���, [����]��ǰ�� ��ϺҰ����մϴ�. \n[��Ź],[��ü���]��ǰ�� ��ϰ����մϴ�. ");
    		return;
    	}
    }
	
	//2.����ڰ� [�鼼�����] �� ���, �鼼��ǰ���θ� ��ϰ��� 
	if((document.itemreg.jungsangubun.value =="�鼼")&&(document.itemreg.vatinclude[0].checked)){
		alert("����ڰ� [�鼼�����]�� ���, [����]��ǰ�� ��ϺҰ����մϴ�. \n[�鼼]��ǰ�� ��ϰ����մϴ�. ");
		document.itemreg.vatinclude[1].focus();
		return; 
	}
	
	//3.����ڰ� [�ٹ�����]�� ���, ���Ի�ǰ�� ��� ����
	if(document.itemreg.mwdiv.length>0){
    	if((document.itemreg.companyno.value =="211-87-00620")&&(!document.itemreg.mwdiv[0].checked)){
    		alert("����ڰ� [�ٹ�����]�� ���, [����]��ǰ�� ��ϰ����մϴ�. ");
    		document.itemreg.mwdiv[0].focus();
    		return;
    	}
    }else{
        if((document.itemreg.companyno.value =="211-87-00620")&&(!document.itemreg.mwdiv.value=="M")){
    		alert("����ڰ� [�ٹ�����]�� ���, [����]��ǰ�� ��ϰ����մϴ�. ");
    		return;
    	}
    }
	 //--------------------------------------------------------------------------------  
    if (validate(itemreg)==false) {
        return;
    }
    
    //��ǰ�� ����üũ �߰� 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("��ǰ���� �ִ� 64byte ���Ϸ� �Է����ּ���.(�ѱ�32�� �Ǵ� ����64��)");
		itemreg.itemname.focus();
		return;
	}
	
    if (itemreg.sailyn[0].checked == true) {
        // ���󰡰�
        if (Math.round((itemreg.sellcash.value*1) * (itemreg.margin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {
    		alert("���ް��� �߸��ԷµǾ����ϴ�.[�Һ��ڰ�*���� = ���ް�]");
    		itemreg.sellcash.focus();

    		if (!confirm('�������� ��� �� �� ������ ���ް��� �Է��ϸ� �������� ���ް��� ���� ���˴ϴ�. \n��� ���� �Ͻðڽ��ϱ�?')){
				return;
			}
        }

        if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
            alert("���ϸ����� �ǸŰ����� Ŭ �� �����ϴ�.");
            itemreg.mileage.focus();
            return;
        }

        if(!itemreg.itemdiv[3].checked) { //Present��ǰ�� �ǸŰ� 0�� ����
	        if (itemreg.sellcash.value*1 < 300 || itemreg.sellcash.value*1 >= 20000000){
				alert("�Ǹ� ������ 300�� �̻� 20,000,000�� �̸����� ��� �����մϴ�.");
				itemreg.sellcash.focus();
				return;
			}
		}

    } else {
        // ���ΰ���
        if (Math.round((itemreg.sailprice.value*1) * (itemreg.sailmargin.value*1) / 100) != ((itemreg.sailprice.value*1) - (itemreg.sailsuplycash.value*1))) {
    		alert("���ް��� �߸��ԷµǾ����ϴ�.[���μҺ��ڰ�*���θ��� = ���ΰ��ް�]");
    		itemreg.sailprice.focus();

    		if (!confirm('��� ���� �Ͻðڽ��ϱ�?')){
				return;
			}
        }

        if (itemreg.mileage.value*1 > itemreg.sailprice.value*1){
            alert("���ϸ����� �ǸŰ����� Ŭ �� �����ϴ�.");
            itemreg.mileage.focus();
            return;
        }

        if(!itemreg.itemdiv[3].checked) { //Present��ǰ�� �ǸŰ� 0�� ����
	        if (itemreg.sailprice.value*1 < 300 || itemreg.sailprice.value*1 >= 20000000){
				alert("�Ǹ� ������ 300�� �̻� 20,000,000���� �̸����� ��� �����մϴ�.");
				itemreg.sailprice.focus();
				return;
			}
		}
    }

    //���ϰ����� ���󰡰� ���� Ŭ �� ����.
    if (itemreg.sailprice.value*1>itemreg.sellcash.value*1){
        alert('���ϰ����� ���󰡺��� Ŭ �� �����ϴ�.');
        return;
    }
    
    if (itemreg.sailsuplycash.value*1>itemreg.buycash.value*1){
        alert('���ϸ��԰��� ���� ���԰����� Ŭ �� �����ϴ�.');
        return;
    }

	// �����Էµ� ���ݺ��� ������ ������ 20%�̻� ���̰� ���� Ȯ�� �޽���
	if(document.itemreg.sellcash.value<<%=fix(oitem.FOneItem.Fsellcash*0.8)%>) {
		if(!confirm("\n\n\n\n�Է��Ͻ� �Һ��ڰ� �����ϱ� ���� ���ݺ��� �ſ� ���� ���̳��ϴ�.\n\n������ ���� [ <%=formatNumber(oitem.FOneItem.Fsellcash,0)%> ]�� �� �Է��Ͻ� ���� [ "+plusComma(document.itemreg.sellcash.value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
			return;
		}
	}
	<% if oitem.FOneItem.Fsailyn="Y" then %>
	if(document.itemreg.sailprice.value<<%=fix(oitem.FOneItem.Fsailprice*0.8)%>) {
		if(!confirm("\n\n\n\n�Է��Ͻ� ���ΰ��� �����ϱ� ���� ���ݺ��� �ſ� ���� ���̳��ϴ�.\n\n������ ���� [ <%=formatNumber(oitem.FOneItem.Fsailprice,0)%> ]�� �� �Է��Ͻ� ���� [ "+plusComma(document.itemreg.sailprice.value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
			return;
		}
	}
	<% end if %>

	// ������ �˻�(50%�̻� ���)
	if (document.itemreg.sailyn[1].checked == true) {
		if(((document.itemreg.sellcash.value-document.itemreg.sailprice.value)/document.itemreg.sellcash.value*100)>50) {
			if(!confirm("\n\n�������� �ſ� ���� �����Ǿ��ֽ��ϴ�.\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n")) {
				return;
			}
		}
	}

	if((itemreg.sellyn[0].checked||itemreg.sellyn[1].checked)&&(itemreg.isusing[1].checked)) {
        alert('�Ǹſ��ο� ��뿩�θ� Ȯ�����ּ���.\n\n�ػ������ �ʴ� ��ǰ�� �Ǹ����� ������ �� �����ϴ�.');
        return;
	}

		itemreg.chkModSR.value = "N"; //�⺻�� ����, ���� ����� �� ����
 //���¿�������϶� �Ǹſ��� ����ó��
 if(itemreg.sellreservedate.value != ""){
 	 if(itemreg.sellyn[0].checked){
 	 	if(confirm(itemreg.sellreservedate.value+"�� ���¿���� ��ǰ�Դϴ�. �Ǹ������� ���� �����Ͻø�, ��ǰ���¿��༳���� ��ҵ˴ϴ�. ����Ͻðڽ��ϱ�? ")){
 	 		itemreg.sellreservedate.value = "";
 	 		itemreg.chkModSR.value = "Y";
 	 	}else{
 	 		itemreg.sellyn[0].focus();
 	 		return;
 	 	}
 	}
 	
 	if(itemreg.sellyn[1].checked){
 	 	if(confirm(itemreg.sellreservedate.value+"�� ���¿���� ��ǰ�Դϴ�. �Ͻ�ǰ���� ���� �����Ͻø�, ��ǰ���¿��༳���� ��ҵ˴ϴ�. ����Ͻðڽ��ϱ�? ")){
 	 		itemreg.sellreservedate.value = "";
 	 		itemreg.chkModSR.value = "Y";
 	 	}else{
 	 		itemreg.sellyn[1].focus();
 	 		return;
 	 	}
 	}
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

	//������������.
    if (itemreg.safetyYn[0].checked){
  		if($("#real_safetynum").val() == ""){
  			alert("�������������� �����ϰ� ������ȣ�� �Է��� �߰���ư�� Ŭ�����ּ���.");
  			return;
  		}
    }

    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("���������� �Է����ּ���!");
        itemreg.limitno.focus();
        return;
    }

	if(itemreg.orderMinNum.value<1||document.itemreg.orderMinNum.value>32000) {
        alert('�ּ��Ǹż��� 1~32,000 ������ ���ڷ� �Է����ּ���.');
        itemreg.orderMinNum.focus();
        return;
	}
	if(itemreg.orderMaxNum.value<1||document.itemreg.orderMaxNum.value>32000) {
        alert('�ִ��Ǹż��� 1~32,000 ������ ���ڷ� �Է����ּ���.');
        itemreg.orderMaxNum.focus();
        return;
	}
	if(parseInt(itemreg.orderMinNum.value)>parseInt(itemreg.orderMaxNum.value)) {
        alert('�ִ��Ǹż����� �ּ��Ǹż��� Ŭ �� �����ϴ�.');
        itemreg.orderMinNum.focus();
        return;
	}

    if (itemreg.basic.value == "del") {
        alert("�⺻�̹����� �ʼ��Դϴ�.");
        return;
    } else {
        if (itemreg.imgbasic.value != "") {
            if (CheckImage(itemreg.imgbasic, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40) != true) {
                return;
            }
        }
    }

    if (itemreg.imgmask.value != "") {
        if (CheckImage(itemreg.imgmask, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd1.value != "") {
        if (CheckImage(itemreg.imgadd1, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage(itemreg.imgadd2, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage(itemreg.imgadd3, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd4.value != "") {
        if (CheckImage(itemreg.imgadd4, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd5.value != "") {
        if (CheckImage(itemreg.imgadd5, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }


    if(confirm("��ǰ�� �ø��ðڽ��ϱ�?") == true){
		// �������� api�� ��ȸ �� ���� ������ db���� �� ����idx�� �޾� ����
		if(itemreg.safetyYn[0].checked) {
			$("#real_safetyidx").val(jsCallAPIsafety($("#real_safetynum").val(),"u",$("#real_safetydiv").val()));
		}

        itemreg.deliverytype[0].disabled=false;
		itemreg.deliverytype[1].disabled=false;
		itemreg.deliverytype[2].disabled=false;
        itemreg.deliverytype[3].disabled=false;
        itemreg.deliverytype[4].disabled=false;
        itemreg.target = "FrameCKP";
        itemreg.submit();
    }

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
 
function TnCheckUpcheYN(frm){
    if(frm.mwdiv.length>0){
    	if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
    		frm.deliverytype[0].checked=true;	// �⺻üũ
    		// ��۱��� ����(�ٹ�����)
    		frm.deliverytype[0].disabled=false;
    		frm.deliverytype[1].disabled=true;
    		frm.deliverytype[2].disabled=false;
    		frm.deliverytype[3].disabled=true;  //��ü���ǹ��(9)
    		frm.deliverytype[4].disabled=true;  //��ü���ҹ��(7) 
    	}
    	else if(frm.mwdiv[2].checked){
    	    // ��۱��� ����(��ü���ǹ��)
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
            frm.deliverytype[3].disabled=false;
            frm.deliverytype[4].disabled=false;  //��ü���ҹ��(7) 
            
    	}
    }else{
        if (frm.mwdiv.value=="M" || frm.mwdiv.value=="W"){
    		frm.deliverytype[0].checked=true;	// �⺻üũ
    		// ��۱��� ����(�ٹ�����)
    		frm.deliverytype[0].disabled=false;
    		frm.deliverytype[1].disabled=true;
    		frm.deliverytype[2].disabled=false;
    		frm.deliverytype[3].disabled=true;  //��ü���ǹ��(9)
    		frm.deliverytype[4].disabled=true;  //��ü���ҹ��(7) 
    	}
    	else if(frm.mwdiv.value=="U"){
    	    // ��۱��� ����(��ü���ǹ��)
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
            frm.deliverytype[3].disabled=false;
            frm.deliverytype[4].disabled=false;  //��ü���ҹ��(7) 
            
    	}
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

function TnCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// ������
		frm.limitno.readonly=true;
		frm.limitno.style.background='#E6E6E6';

		frm.limitsold.readonly=true;
		frm.limitsold.style.background='#E6E6E6';
		
		document.all.dvDisp.style.display = 'none';
		this.form.limitdispyn[0].checked = false;
		this.form.limitdispyn[1].checked = true;
	}
	else {
		// ����
		if ((frm.optioncnt.value*1) > 0) {
		    // �ɼǻ����
		    alert("�ɼ��� ����Ұ�� ���������� �ɼ�â���� ���������մϴ�.");
		    frm.limityn[0].checked = true;
		    return;
        }

		frm.limitno.readonly = false;
		frm.limitno.style.background = '#FFFFFF';

		frm.limitsold.readonly = false;
		frm.limitsold.style.background = '#FFFFFF';
		
		document.all.dvDisp.style.display = '';
	}
}

function TnSilentCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// ������
		frm.limitno.readonly=true;
		frm.limitno.style.background='#E6E6E6';

		frm.limitsold.readonly=true;
		frm.limitsold.style.background='#E6E6E6';
	}
	else {
		// ����
		if ((frm.optioncnt.value*1) > 0) {
		    // �ɼǻ����
		    // alert("���������� �ɼ�â���� ���������մϴ�.");
		    return;
        }

		frm.limitno.readonly=false;
		frm.limitno.style.background='#FFFFFF';

		frm.limitsold.readonly=false;
		frm.limitsold.style.background='#FFFFFF';
	}
}

function TnGoClear(frm){
	frm.buycash.value = "";
	frm.mileage.value = "";
}

function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
	    if(frm.mwdiv.length>0){
    		if (frm.mwdiv[2].checked){
    			alert("������Ź ������ ��ü�� ���\n��۱����� �ٹ����� ������� ���� �Ͻ� �� �����ϴ�!!\n������Ź������ Ȯ�����ּ���!!");
    			frm.mwdiv[0].checked=true;
    		}
    	}else{
    	    if (frm.mwdiv.value=="U"){
    			alert("������Ź ������ ��ü�� ���\n��۱����� �ٹ����� ������� ���� �Ͻ� �� �����ϴ�!!\n������Ź������ Ȯ�����ּ���!!");
    		}
    	}
	}
	else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
	//else if(frm.deliverytype[1].checked ){
	    if(frm.mwdiv.length>0){
    		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
    			alert("������Ź ������ �����̳� ��Ź�� ���\n��۱�����  ��ü������� ���� �Ͻ� �� �����ϴ�!!!\n������Ź������ Ȯ�����ּ���!!");
    			frm.mwdiv[2].checked=true;
    		}
    	}else{
    	    if (frm.mwdiv.value=="M" || frm.mwdiv.value=="W"){
    			alert("������Ź ������ �����̳� ��Ź�� ���\n��۱�����  ��ü������� ���� �Ͻ� �� �����ϴ�!!!\n������Ź������ Ȯ�����ּ���!!");
    		}
    	}
	}
}

function TnChkIsUsing(frm) {
	if(frm.isusing[0].checked) {
		frm.sellyn[0].disabled=false;
		frm.sellyn[1].disabled=false;
	} else {
		if(frm.sellyn[0].checked||frm.sellyn[1].checked) {
			alert("��뿩�θ� ���������� �����ϼ̽��ϴ�.\n�Ǹſ��ΰ� [�Ǹž���]���� �ڵ������˴ϴ�.");
		}
		frm.sellyn[2].checked=true;
		frm.sellyn[0].disabled=true;
		frm.sellyn[1].disabled=true;
	}
}

function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // �ɼǻ��
        frm.btnoptadd.disabled = false;
        frm.btnoptdel.disabled = false;
	} else {
	    // �ɼǾ���
	    while (frm.realopt.length > 0) {
	        frm.realopt.options[0] = null;
        }
        frm.btnoptadd.disabled = true;
        frm.btnoptdel.disabled = true;
    }
}

function TnCheckSailYN(frm){
	CheckSailEnDisabled(frm);
    CalcuAuto(frm);
}

function CheckSailEnDisabled(frm){
	if (frm.sailyn[0].checked == true) {
	    // ���󰡰�
        frm.sellcash.readonly = false;
        frm.margin.readonly = false;

        frm.sellcash.style.background = '#FFFFFF';
        frm.margin.style.background = '#FFFFFF';

        frm.sailprice.readonly = true;
        frm.sailmargin.readonly = true;

        frm.sailprice.style.background = '#E6E6E6';
        frm.sailmargin.style.background = '#E6E6E6';
	} else {
	    // ���ϰ���
        frm.sellcash.readonly = true;
        frm.margin.readonly = true;

        frm.sellcash.style.background = '#E6E6E6';
        frm.margin.style.background = '#E6E6E6';

        frm.sailprice.readonly = false;
        frm.sailmargin.readonly = false;

        frm.sailprice.style.background = '#FFFFFF';
        frm.sailmargin.style.background = '#FFFFFF';
    }
}

function ClearVal(comp){
    comp.value = "";
}

// ============================================================================
	// ī�°� ���� �˾�
	function popCateSelect(iid){
		var popwin = window.open("/common/module/NewCategorySelect.asp?iid=" + iid, "popCateSel","width=700,height=400,scrollbars=yes,resizable=yes");
        popwin.focus();
	}

	// ����ī�װ� ���� �˾�
	function popDispCateSelect(){
		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/common/module/act_DispCategorySelect.asp?isDft="+dCnt,
			cache: false,
			success: function(message) {
				$("#lyrDispCateAdd").empty().append(message).fadeIn();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// �˾����� ���� ī�װ� �߰�
	function addCateItem(lcd,lnm,mcd,mnm,scd,snm,div)
	{
		// ������ ���� �ߺ� ī�װ� ���� �˻� - �ö�� ��������� ������;
		if(tbl_Category.rows.length>0)	{
			if(tbl_Category.rows.length>1)	{
				for(l=0;l<document.all.cate_div.length;l++)	{
				    if (!((document.all.cate_large[l].value=="110")&&(document.all.cate_mid[l].value=="060"))){
    					if((document.all.cate_large[l].value==lcd)&&(document.all.cate_mid[l].value==mcd)) {
    						alert("���� �ߺз��� �̹� ������ ī�װ��� �ֽ��ϴ�.\n���� ī�װ��� �����ϰ� �ٽ� �������ּ���.");
    						return;
    					}
    				}
				}
			}
			else {
			    if (!((document.all.cate_large.value=="110")&&(document.all.cate_mid.value=="060"))){
    				if((document.all.cate_large.value==lcd)&&(document.all.cate_mid.value==mcd)) {
    					alert("���� �ߺз��� �̹� ������ ī�װ��� �ֽ��ϴ�.\n�ر��� ī�װ��� �����ϰ� �ٽ� �������ּ���.");
    					return;
    				}
    			}
			}
		}
		
		// ���߰�
		var oRow = tbl_Category.insertRow();
		oRow.onmouseover=function(){tbl_Category.clickedRowIndex=this.rowIndex};

		// ���߰� (����,ī�װ�,������ư)
		var oCell1 = oRow.insertCell();		
		var oCell2 = oRow.insertCell();
		var oCell3 = oRow.insertCell();

		if(div=="D") {
			oCell1.innerHTML = "<font color='darkred'><b>[�⺻]<b></font><input type='hidden' name='cate_div' value='D'>";
		} else {
			oCell1.innerHTML = "<font color='darkblue'>[�߰�]</font><input type='hidden' name='cate_div' value='A'>";
		}
		oCell2.innerHTML = lnm + " >> " + mnm + " >> " + snm
					+ "<input type='hidden' name='cate_large' value='" + lcd + "'>"
					+ "<input type='hidden' name='cate_mid' value='" + mcd + "'>"
					+ "<input type='hidden' name='cate_small' value='" + scd + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delCateItem()' align=absmiddle>";
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

		//��ǰ�Ӽ� ���
		printItemAttribute();
	}

	// ���� ī�װ� ����
	function delCateItem()
	{
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?"))
			tbl_Category.deleteRow(tbl_Category.clickedRowIndex);
	}

	// ���� ����ī�װ� ����
	function delDispCateItem() {
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			//��ǰ�Ӽ� ���
			printItemAttribute();
		}
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

    //��Ż ��ǰ�ΰ��.
    if (frm.itemdiv[7].checked){
		frm.reserveItemTp[1].checked = true;
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
			url: "act_itemInfoDivForm.asp",
			data: "itemid=<%=itemid%>&ifdv="+v,
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

	// ��������üũ. ���ȹ�
	jsSafetyCheck('','');
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

//��ǰ�����̹����߰�
function InsertImageUp() {
	var f = document.all;
	var rowLen = f.imgIn.rows.length;

	if(rowLen > 6){
		alert("�̹����� �ִ� 7������ �����մϴ�.");
		return;
	}
	
	var i = rowLen;
	var r  = f.imgIn.insertRow(rowLen++);
	var c0 = r.insertCell(0);
	var c1 = r.insertCell(1);

	r.style.textAlign = 'left';
	c0.style.height = '30';
	c0.style.width = '15%';
	c0.style.background = '#DDDDFF';
	c0.innerHTML = 'PC��ǰ�����̹��� #' + rowLen + ' :';
	c1.style.background = '#FFFFFF';
	c1.innerHTML = '<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40, '+parseInt(rowLen-1)+');" class="text" size="40"> ';
	c1.innerHTML += '<input type="button" value="#'+parseInt(rowLen)+' �̹��������" class="button" onClick="ClearImage2(this.form.addimgname['+parseInt(rowLen-1)+'],40, 800, 1600, '+parseInt(rowLen-1)+')"> (����,800X1600, Max 800KB,jpg,gif)';
	c1.innerHTML += '<input type="hidden" name="addimggubun" value="'+parseInt(rowLen)+'">';
	c1.innerHTML += '<input type="hidden" name="addimgdel" value="">';
}


//����ϻ�ǰ���̹����߰�
function InsertMobileImageUp() {
	var f = document.all;
	var rowLen = f.MobileimgIn.rows.length;

	if(rowLen > 11){
		alert("�̹����� �ִ� 12������ �����մϴ�.");
		return;
	}
	
	var i = rowLen;
	var r  = f.MobileimgIn.insertRow(rowLen++);
	var c0 = r.insertCell(0);
	var c1 = r.insertCell(1);

	r.style.textAlign = 'left';
	c0.style.height = '30';
	c0.style.width = '15%';
	c0.style.background = '#DDDDFF';
	c0.innerHTML = '����ϻ�ǰ���̹��� #' + rowLen + ' :';
	c1.style.background = '#FFFFFF';
	c1.innerHTML = '<input type="file" name="addmoblieimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40);" class="text" size="40"> ';
	c1.innerHTML += '<input type="button" value="#'+parseInt(rowLen)+' �̹��������" class="button" onClick="ClearImage2(this.form.addimgname['+parseInt(rowLen-1)+'],40, 640, 1200)"> (����,640X1200, Max 400KB,jpg,gif)';
}


	
//��ǰ���� ���� ������ ���� ǥ��
function jsSetArea(iValue){ 
	var i;
	for(i=0;i<=4;i++) { 
 		eval("document.all.dvArea"+i).style.display = "none";
	}
 	eval("document.all.dvArea"+iValue).style.display = "";
}

function jsCallAPIsafety(certnum,isSave,safetydiv){
	var returnmsg = "";
	$.ajax({
		url: "/admin/itemmaster/safety_api_auth_proc.asp?itemid=<%=itemid%>&issave="+isSave+"&certnum="+certnum+"&safetydiv="+safetydiv+"&statusmode=real",
		cache: false,
		async: false,
		success: function(message)
		{
			returnmsg = message;
		}
	});
	return returnmsg;
}

//����ī�װ�(����������)�� ���� alert �޼���.
function jsAlertCatecodeSafety(){
	var auth_go_catecode = "";
	if(typeof itemreg.catecode != "undefined"){
		if(itemreg.catecode.length == undefined){
			auth_go_catecode = itemreg.catecode.value;
		}else{
			for(si=0; si<itemreg.catecode.length; si++){
				auth_go_catecode = auth_go_catecode + itemreg.catecode[si].value + ",";
			}
		}
		
		if(auth_go_catecode != ""){
			$("#auth_go_catecode").val(auth_go_catecode);
			
			var ccode = $("#auth_go_catecode").val();
			$.ajax({
					url: "/common/item/catecode_safety_info_ajax.asp?catecode="+ccode,
					cache: false,
					async: false,
					success: function(msgc)
					{
						if(msgc != ""){
							msgc = msgc.replace(/br/gi,"\n");
							alert(msgc);
						}
					}
			});
		}
	}else{
		alert("����ī�װ��� �������ּ���.");
	}
}

//�߰��� �������� ����Ʈ ���� ����
function jsSafetyDivListDel(listnum){
	var realvalue = $("#real_safetydiv").val();
	var jbSplit = $("#real_safetydiv").val().split(",");
	var jbSplitnum = $("#real_safetynum").val().split(",");
	var resultDiv = "";
	var resultNum = "";
	var del_safetynum = "";
	var del_safetydiv = "";
	
	for(var i in jbSplit){
		if(jbSplit[i] != listnum){
			resultDiv = resultDiv + jbSplit[i] + ",";
			resultNum = resultNum + jbSplitnum[i] + ",";
		}else{
			del_safetynum = jbSplitnum[i];
			del_safetydiv = jbSplit[i];
		}
	}
	
	if(resultDiv.substr(resultDiv.length-1, 1) == ","){
		resultDiv = resultDiv.substr(0, resultDiv.length-1);
		resultNum = resultNum.substr(0, resultNum.length-1);
	}
	$("#real_safetydiv").val(resultDiv);
	$("#real_safetynum").val(resultNum);
	
	$("#l"+listnum+"").remove();
	
	var tmp_num = $("#real_safetynum_delete").val();
	var tmp_div = $("#real_safetydiv_delete").val();
	if(tmp_num == ""){
		$("#real_safetynum_delete").val(del_safetynum);
		$("#real_safetydiv_delete").val(del_safetydiv);
	}else{
		$("#real_safetynum_delete").val(tmp_num + "," + del_safetynum);
		$("#real_safetydiv_delete").val(tmp_div + "," + del_safetydiv);
	}
}
