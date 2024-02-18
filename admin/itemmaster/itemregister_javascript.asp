function UseTemplate() {
	var popwin = window.open("/common/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
    popwin.focus();
}

function printItemAttribute() {
	var arrDispCd="";
	$("input[name='catecode']").each(function(i){
		if(i>0) arrDispCd += ",";
		arrDispCd += $(this).val();
	});
	$.ajax({
		url: "/common/module/act_ItemAttribSelect.asp?itemid=0&arrDispCate="+arrDispCd,
		cache: false,
		success: function(message) {
			$("#lyrItemAttribAdd").empty().append(message);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value               = varArray[0];
	document.itemreg.margin.value                   = varArray[1];
	document.itemreg.defaultmargin.value            = varArray[1];  //��ü�⺻����.
	document.itemreg.defaultmaeipdiv.value          = varArray[2];
    document.itemreg.defaultFreeBeasongLimit.value  = varArray[3];
    document.itemreg.defaultDeliverPay.value        = varArray[4];
    document.itemreg.defaultDeliveryType.value      = varArray[5];

    if (document.itemreg.defaultmaeipdiv.value=="M"){
        document.itemreg.mwdiv[0].checked = true; //����
    }else if (document.itemreg.defaultmaeipdiv.value=="W"){
        document.itemreg.mwdiv[1].checked = true; //��Ź
    }else if (document.itemreg.defaultmaeipdiv.value=="U"){
        document.itemreg.mwdiv[2].checked = true; //��ü
    }

    TnCheckUpcheYN(document.itemreg);
}

// ============================================================================
// ��ü�����ڵ��Է�
function TnDesignerNMargineAppl2(){
	var varArray;
	varArray = document.itemreg.marginData.value.split(',');

	document.itemreg.designerid.value = document.itemreg.makerid.value;
	document.itemreg.margin.value = varArray[0];

    document.itemreg.defaultmargin.value            = varArray[0];  //��ü�⺻����.
	document.itemreg.defaultmaeipdiv.value          = varArray[1];
    document.itemreg.defaultFreeBeasongLimit.value  = varArray[2];
    document.itemreg.defaultDeliverPay.value        = varArray[3];
    document.itemreg.defaultDeliveryType.value      = varArray[4];

    if (document.itemreg.defaultmaeipdiv.value=="M"){
        document.itemreg.mwdiv[0].checked = true; //����
    }else if (document.itemreg.defaultmaeipdiv.value=="W"){
        document.itemreg.mwdiv[1].checked = true; //��Ź
    }else if (document.itemreg.defaultmaeipdiv.value=="U"){
        document.itemreg.mwdiv[2].checked = true; //��ü
    }

    TnCheckUpcheYN(document.itemreg);
}

function CalcuAuto(frm){
	var imargin, isellcash, ibuycash;
	var imileage;
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
		ibuycash = isellcash - Math.round(isellcash*imargin/100);       //parseInt-> round�� ����
		imileage = parseInt(isellcash*0.005) ;
	}else{
		ibuycash = isellcash - Math.round(isellcash*imargin/100);       //parseInt-> round�� ����
		imileage = parseInt(isellcash*0.005) ;
	}

	frm.buycash.value = ibuycash;
	frm.mileage.value = imileage;

	//�ִ뱸�ż��� ����(���ݺ��)
	if(isellcash<100) {
		frm.orderMaxNum.value="1";
	} else if(isellcash<10000) {
		frm.orderMaxNum.value="500";
	} else if(isellcash<100000) {
		frm.orderMaxNum.value="200";
	} else {
		frm.orderMaxNum.value="100";
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

	EnDisableFlowerShop();
}

// ============================================================================
	// ī�°� ���� �˾�
	function popCateSelect(iid){
	    var dftDiv = "";
	    var chk = 0;

	    //�⺻ ī�װ����� �߰����� üũ

	    if (!document.all.cate_div){
	        dftDiv = "D";
	    }else{
	        if (document.all.cate_div.length==undefined){
	            if (document.all.cate_div.value=="D") chk++;
	        }else{
        	    for(l=0;l<document.all.cate_div.length;l++)	{
        			if (document.all.cate_div[l].value=="D") chk++;
        		}
        	}
		}

		if (chk<1) dftDiv="D";

		var popwin = window.open("/common/module/NewCategorySelect.asp?iid=" + iid + "&dftDiv=" + dftDiv, "popCateSel","width=700,height=400,scrollbars=yes,resizable=yes");
        popwin.focus();
	}

	// ����ī�װ� ���� �˾�
	function popDispCateSelect(){
		var designerid = document.all.itemreg.designerid.value;
		if(designerid == ""){
			alert("��ü�� �����ϼ���.");
			return;
		}

		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/common/module/act_DispCategorySelect.asp?designerid="+designerid+"&isDft="+dCnt,
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

		//��ǰ�Ӽ� ���
		printItemAttribute();
	}

	// ���� ����ī�װ� ����
	function delDispCateItem() {
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			//��ǰ�Ӽ� ���
			printItemAttribute();
		}
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

function ClearImage2(img,fsize,wd,ht) {
    img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +");\" class='text' size='"+ fsize +"'>";
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
        ClearImage(img,fsize,imagewidth,imageheight);
        return false;
    }

    return true;
}


// ============================================================================
// �����ϱ�
function SubmitSave() {
	var itemreg = document.all.itemreg;

	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.makerid.focus();
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }

	if(!$("input[name='isDefault'][value='y']").length) {
		alert("���� ī�װ��� �����ϼ���.\n�� ���� �⺻ ī�װ��� �ʼ� �ֽ��ϴ�.");
		return;
	}

	// �Է��� ������ �ٸ���� üũ
    if (itemreg.margin.value.length>0){
        if (Math.round((itemreg.sellcash.value*1) * (itemreg.margin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {
    		alert("���ް��� �߸��ԷµǾ����ϴ�.[�Һ��ڰ�*���� = ���ް�]");


    		if (!confirm('�Է��� ������ �Էµ� �ǸŰ� ��� ���԰� �ݾ��� ���� �մϴ�. ��� ���� �Ͻðڽ��ϱ�?')){
    		    itemreg.sellcash.focus();
    			return;
    		}
        }
	}

	// ��ü �⺻������ �ٸ���� üũ
	if (itemreg.defaultmargin.value.length>0){
        if (Math.round((itemreg.sellcash.value*1) * (itemreg.defaultmargin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {

    		if (!confirm('��ü �⺻ ������ �Էµ� �ǸŰ� ��� ���԰� �ݾ��� ���� �մϴ�. ��� ���� �Ͻðڽ��ϱ�?')){
    			return;
    		}
        }
	}

    //��۱��� üũ =======================================
    //��ü ���ǹ��
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[3].checked){
            alert('��� ������ Ȯ�����ּ���. [��ü ���ǹ��] ��ü�� �ƴմϴ�.');
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
        if ((itemreg.mwdiv[0].checked)||(itemreg.mwdiv[1].checked)){
            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
            return;
        }
      //  if (itemreg.deliverOverseas.checked){
       //     alert('�ٹ����� ����� ��쿡�� �ؿܹ���� �Ͻ� �� �ֽ��ϴ�.');
      //      return;
      //  }
    }

    if (itemreg.mwdiv[2].checked){
        if ((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
            return;
        }
    }

    //��ü��۸� �ֹ����� ����.
    if ((!itemreg.mwdiv[2].checked)&&(itemreg.itemdiv[1].checked)){
        alert('�ֹ� ���ۻ�ǰ�� ��ü����ΰ�츸 �����մϴ�.');
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

	// ��۹�� �ؿ����� üũ
	if (itemreg.deliverfixday[3].checked == true){
		if (itemreg.mwdiv[2].checked == false){
			alert('�ؿ������� ��ü��۸� ���� ���� �մϴ�.');
			return;
		}
		if ( !(itemreg.deliverytype[1].checked == true || itemreg.deliverytype[3].checked == true) ){
			alert('�ؿ������� ��ü�����۰� ��ü���ǹ�۸� ���� ���� �մϴ�.');
			return;
		}
		if (itemreg.deliverarea[0].checked == false){
			alert('�ؿ������� ������۸� ���� ���� �մϴ�.');
			return;
		}
	}

    //==================================================================================


	if(!itemreg.itemdiv[3].checked) { //Present��ǰ�� �ǸŰ� 0�� ����
	    if (itemreg.buycash.value*1>itemreg.sellcash.value*1){
	        alert("���԰����� �ǸŰ� ���� Ů�ϴ�.");
			itemreg.sellcash.focus();
			return;
	    }

		if (itemreg.sellcash.value*1 < 0 || itemreg.sellcash.value*1 >= 20000000){
			alert("�Ǹ� ������ 20,000,000�� �̸����� ��� �����մϴ�.");
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
	}

	if((itemreg.sellyn[0].checked)&&(itemreg.isusing[1].checked)) {
        alert('�Ǹſ��ο� ��뿩�θ� Ȯ�����ּ���.\n\n�ػ������ �ʴ� ��ǰ�� �Ǹ����� ������ �� �����ϴ�.');
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

                for (var i=0;i<itemreg.optionTypename3.length;i++){
                    if (itemreg.optionTypename3[i].value.length>0) chkCnt++;
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

	var imgupcnt = document.all.imgIn.rows.length;
	var tmp = "";
	var tmpvalue = "";
	for(var a=0;a<imgupcnt;a++){
		tmp = itemreg.addimgname[a];
	    if (tmp.value != "") {
	        if (CheckImage(tmp, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40) != true) {
	            return;
	        }
	    }
	}

    if(confirm("��ǰ�� �ø��ðڽ��ϱ�?") == true){
		//�������� api�� ��ȸ �� ���� ������ db���� �� ����idx�� �޾� ����
		if(itemreg.safetyYn[0].checked) {
			$("#real_safetyidx").val(jsCallAPIsafety($("#real_safetynum").val(),"o",$("#real_safetydiv").val()));
		}

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

//���Ա��� üũ�� ���� ��۱��� üũ
function TnCheckUpcheYN(frm){
	if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
		frm.deliverytype[0].checked=true;	// �⺻üũ
		// ��۱��� ����(�ٹ�����)
		frm.deliverytype[0].disabled=false;
		frm.deliverytype[1].disabled=true;
		frm.deliverytype[2].disabled=false;
        frm.deliverytype[3].disabled=true;  //��ü�������(9)
        frm.deliverytype[4].disabled=true;  //��ü���ҹ��(7)
        frm.deliverOverseas.checked=true;	// �ؿܹ��üũ
       // frm.optlevel[0].checked=true;
       // frm.optlevel[1].disabled=true;
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
        frm.deliverytype[3].disabled=false;
        frm.deliverytype[4].disabled=false;  //��ü���ҹ��(7)
        frm.deliverOverseas.checked=false;	// �ؿܹ��üũ����
      //  frm.optlevel[1].disabled=false;
	}

	if (frm.deliverytype[1].checked==true || frm.deliverytype[3].checked==true){
		frm.deliverfixday[3].disabled=false;	// �ؿ�����
	}
}

function TnGoClear(frm){
	frm.buycash.value = "";
	frm.mileage.value = "";
}

// ��۱���
function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if (frm.mwdiv[2].checked){
			alert("������Ź ������ ��ü�� ���\n��۱����� �ٹ����� ������� ���� �Ͻ� �� �����ϴ�!!\n������Ź������ Ȯ�����ּ���!!");
			frm.mwdiv[0].checked=true;
			//frm.optlevel[1].checked=false;
			//frm.optlevel[1].disabled=true;
		}
	}
	else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked || frm.deliverytype[4].checked){
	//else if(frm.deliverytype[1].checked){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("������Ź ������ �����̳� ��Ź�� ���\n��۱�����  ��ü������� ���� �Ͻ� �� �����ϴ�!!!\n������Ź������ Ȯ�����ּ���!!");
			frm.mwdiv[2].checked=true;
			//frm.optlevel[1].disabled=false;
		}
	}
}

function TnChkIsUsing(frm) {
	if(frm.isusing[0].checked) {
		frm.sellyn[0].disabled=false;
	} else {
		if(frm.sellyn[0].checked) {
			alert("��뿩�θ� ���������� �����ϼ̽��ϴ�.\n�Ǹſ��ΰ� [�Ǹž���]���� �ڵ������˴ϴ�.");
		}
		frm.sellyn[1].checked=true;
		frm.sellyn[0].disabled=true;
	}
}

function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // �ɼǻ��

        opttype.style.display="";

        if (frm.optlevel[1].checked==true){
            optlist.style.display ="none";
            optlist2.style.display ="";
        }else{
            optlist.style.display="";
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

// ��۹��
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

	// �ؿ�����
	} else if(frm.deliverfixday[3].checked) {
		frm.mwdiv[2].checked=true;
		frm.deliverarea[0].checked=true;

		document.getElementById("lyrFreightRng").style.display="none";
	} else {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=false;
		frm.deliverarea[2].disabled=false;
		document.getElementById("lyrFreightRng").style.display="none";
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

	//Ƽ�� ��ǰ�� ��� ��۹�� Ŭ���� ���� Ȱ��ȭ(2018-05-10 ������)
    if (comp.name=="itemdiv"){
        if (frm.itemdiv[2].checked){
            frm.deliverfixday[4].disabled=false;
        }else{
            frm.deliverfixday[4].disabled=true;
			frm.deliverfixday[0].checked=true;
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

	//��Ż��ǰ�� ���
	if (frm.itemdiv[4].checked){
		frm.orderMaxNum.value=1;
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
	c1.innerHTML = '<input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40);" class="text" size="40"> ';
	c1.innerHTML += '<input type="button" value="#'+parseInt(rowLen)+' �̹��������" class="button" onClick="ClearImage2(this.form.addimgname['+parseInt(rowLen-1)+'],40, 800, 1600)"> (����,800X1600, Max 800KB,jpg,gif)';
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
		url: "/admin/itemmaster/safety_api_auth_proc.asp?issave="+isSave+"&certnum="+certnum+"&safetydiv="+safetydiv+"&statusmode=real",
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

	for(var i in jbSplit){
		if(jbSplit[i] != listnum){
			resultDiv = resultDiv + jbSplit[i] + ",";
			resultNum = resultNum + jbSplitnum[i] + ",";
		}
	}

	if(resultDiv.substr(resultDiv.length-1, 1) == ","){
		resultDiv = resultDiv.substr(0, resultDiv.length-1);
		resultNum = resultNum.substr(0, resultNum.length-1);
	}
	$("#real_safetydiv").val(resultDiv);
	$("#real_safetynum").val(resultNum);

	$("#l"+listnum+"").remove();
}
