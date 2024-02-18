<%@ language=vbscript %>
<%
	option explicit
	session.codepage = 949
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lectureadmin/lib/Inc_AgreeReq.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<%
'CONST CBASIC_IMG_MAXSIZE = 180   'KB
'CONST CMAIN_IMG_MAXSIZE = 500   'KB

'2016 ������ �����̹��� ���� �뷮 ����?
CONST CBASIC_IMG_MAXSIZE = 600   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB

Dim npartner
set npartner = new CPartnerUser
npartner.FRectDesignerID = session("ssBctID")
npartner.GetAcademyPartnerList

if (npartner.FTotalCount < 1) then
	
	response.write "<script>alert('������ ����Ǿ����ϴ�. �ٽ� �α����ϼ���.');</script>"
	response.write "<script>history.back();</script>"
	response.end

end if


dim i,j,k
%>
<script language="JavaScript" src="/js/jquery-1.7.2.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript">
<!-- #include file="./itemregister_javascript.asp"-->
</script>
<script>
function UseTemplate() {
	window.open("/academy/comm/pop_basic_item_info_list.asp", "UseTemplate", "width=700, height=450, scrollbars=yes, resizable=yes");
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
	var imargin, isellcash, ibuycash, itemWeight;
	var isellvat, ibuyvat, imileage;
	imargin = frm.margin.value;
	isellcash = frm.sellcash.value;
	itemWeight = frm.itemWeight.value;

	isvatYn = frm.vatYn[0].checked;

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

	if (!IsDigit(itemWeight)){
		alert('���Դ� ���ڷ� �Է��ϼ���.');
		frm.itemWeight.focus();
		return;
	}

	if (isvatYn==true){
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

// ����ī�װ� ���� �˾�
function popDispCateSelect(){
	if($("input[name='catecode']").length>1){
		alert("ī�װ��� 2������ ���� �����մϴ�.");
		return;
	}

	var designerid = document.all.itemreg.designerid.value;
	if(designerid == ""){
		alert("��ü�� �����ϼ���.");
		return;
	}
	
	var dCnt = $("input[name='isDefault'][value='y']").length;
	$.ajax({
		url: "/academy/comm/act_DispCategorySelect.asp?designerid="+designerid+"&isDft="+dCnt+"&isUpche=upche",
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

		if($("input[name='catecode']").length>1){
			$("#btnAddDispCate").hide();
		}

		//��ǰ�Ӽ� ���
		//printItemAttribute();
	}
	
	// ���� ����ī�װ� ����
	function delDispCateItem() {
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			if($("input[name='catecode']").length<2){
				$("#btnAddDispCate").show();
			}

			//��ǰ�Ӽ� ���
			//printItemAttribute();
		}
	}
// ============================================================================
// �ɼǼ���
function editItemOption(itemid, waityn) {
	var param = "itemid=" + itemid + "&waityn=" + waityn;

	popwin = window.open('/academy/comm/pop_itemoption.asp?' + param ,'editItemOption','width=800,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function doEditItemOption(itemid, waityn, arrmode, arritemoption, arritemoptionname, arroptuseyn, arroptsellyn, arroptlimityn, arroptlimitno, arroptlimitsold) {
	alert("a");
	// var param = "itemid=" + itemid + "&waityn=" + waityn;

	// popwin = window.open('/academy/comm/pop_itemoption.asp?' + param ,'editItemOption','width=700,height=400,scrollbars=yes,resizable=yes');
	// popwin.focus();
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
function ClearImage(img) {
    var e = eval("itemreg." + img);

	if (img == "imgbasic") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg');CheckImageSize(this);\" size='40'>";
    } else if (img == "imgmain") {
       e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    } else if (img == "imgadd1") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    } else {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    }

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

    if (typeof img !== 'string' ){ return false; }
    e = eval("itemreg." + img);
    filename = e.value;
    
    if (filename == "") { return false; }

    if (CheckExtension(filename, extname) != true) {
        alert("�̹���ȭ���� ������ ȭ�ϸ� ����ϼ���.[" + extname + "]");
        ClearImage(img);
        return false;
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


// ============================================================================
// �����ϱ�
function SubmitSave() {
//alert('���� ���� �۾� ������ ��ǰ ���/ ������ �Ұ��մϴ�.');
//return;

	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
		return;
	}

	if ($("input[name='isDefault'][value='y']").length == 0){
		alert("[�⺻] ���� ī�װ��� �����ϼ���.");
		return;
	}
	
	if (!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("[�⺻] ���� ī�װ��� �����ϼ���.");
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

//	if (itemreg.itemsize.value!=''){
//		if (itemreg.unit.value!=''){
//			itemreg.itemsize.value=itemreg.itemsize.value + '(' + itemreg.unit.value + ')';
//		}
//	}

	if (itemreg.itemWeight.value==''){
		itemreg.itemWeight.value='0';
	}

    //��۱��� üũ =======================================
    //��ü ���ǹ��
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[1].checked){
            alert('��� ������ Ȯ�����ּ���. [��ü ���ǹ��] ��ü�� �ƴմϴ�.');
            itemreg.deliverytype[1].focus();
            return;
        }
    }

    //��ü���ҹ�� : ���ǹ�۵� ���Ҽ�������
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype[2].checked)){
        alert('��� ������ Ȯ�����ּ���. [��ü ���ҹ��,��ü ���ǹ��] ��ü�� �ƴմϴ�.');
        itemreg.deliverytype[2].focus();
        return;
    }

    //==================================================================================




//	���ް� �����Է� ����.
//    if (parseInt((itemreg.sellcash.value*1) * (itemreg.margin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {
//		alert("���ް��� �߸��ԷµǾ����ϴ�.[�Һ��ڰ�*���� = ���ް�]");
//		itemreg.sellcash.focus();
//		return;
//    }


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

    if (!IsDigit(itemreg.itemWeight.value)){
		alert('���Դ� ���ڷ� �Է��ϼ���.');
		itemreg.itemWeight.focus();
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

//    if (itemreg.imgadd4.value != "") {
//        if (CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
//            return;
//        }
//    }
//
//    if (itemreg.imgadd5.value != "") {
//        if (CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
//            return;
//        }
//    }

//    if (itemreg.imgmain.value != "") {
//        if (CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif') != true) {
//            return;
//        }
//    }

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
        
        if (itemreg.itemvideo.value.length>0){
            itemreg.itemvideo.value = itemreg.itemvideo.value.replace(/iframe/gi, "BUFiframe");
        }
		itemreg.target = "FrameCKP";
        itemreg.submit();
    }

}

function TnGoClear(frm){
	frm.buyvat.value = "";
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

// ============================================================================
// �̹��� �˸�â
function PopImageInformation(){
	//window.open("/designer/itemmaster/itemreg_info_win.asp","PopImageInformation","width=920,height=600,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no");
	window.open("https://drive.google.com/drive/folders/0B3jVc8T-HBnpR18tWTA5U3FGcHM","PopImageInformation","width=1024,height=600,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no");
}

function ClearVal(comp){
    comp.value = "";
}

function chgodr(v){
	if (v == 1){
		$("#customorder").css("display","none");
	}else{
		$("#customorder").css("display","");
	}
}

function chgodr2(v){
	if (v == 1){
		$("#subodr").css("display","none");
	}else{
		$("#subodr").css("display","");
	}
}

function requireimg(){
	var frm = document.itemreg;
	if (frm.requireimgchk.checked){
		$("#rmemail").css("display","");
	}else{
		$("#rmemail").css("display","none");
	}
}

function CheckImageSize(obj) {
	var MaxSize=600;
	if((obj.files[0].size/1024) > MaxSize){
		alert("�̹����� 600kb ���� �ø��� �� �ֽ��ϴ�. (" + ((obj.files[0].size/1024)-MaxSize).toFixed(2) + "kb �ʰ�)" );
		obj.value="";
		return;
	}
}
</script>

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


<!-- ǥ �߰��� ����-->
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
<!-- ǥ �߰��� ��-->
<% if (TRUE) or (session("ssBCTid")="fingertest01") then %>
<form name="itemreg" method="post" action="<%= uploadImgUrl %>/linkweb/academy/items/WaitDIYItemRegister_Process.asp" onsubmit="return false;" enctype="multipart/form-data">
<% else %>
<form name="itemreg" method="post" action="<%= UploadImgFingers %>/linkweb/items/WaitDIYItemRegister_Process.asp" onsubmit="return false;" enctype="multipart/form-data">
<% end if %>
<input type="hidden" name="itemoptioncode2">
<input type="hidden" name="itemoptioncode3">
<input type="hidden" name="designerid" value="<%= session("ssBctID") %>">
<input type="hidden" name="defultmargine" value="<%= npartner.FPartnerList(0).Fdiy_margin %>">
<input type="hidden" name="defaultmaeipdiv" value="<%= npartner.FPartnerList(0).Fmaeipdiv %>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%= npartner.FPartnerList(0).FdefaultFreeBeasongLimit %>">
<input type="hidden" name="defaultDeliverPay" value="<%= npartner.FPartnerList(0).FdefaultDeliverPay %>">
<input type="hidden" name="defaultDeliveryType" value="<%= npartner.FPartnerList(0).FdefaultDeliveryType %>">
<input type="hidden" name="cd1" value="999">
<input type="hidden" name="cd2" value="999">
<input type="hidden" name="cd3" value="999">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">ī�װ� ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td id="lyrDispList"><table class="a" id="tbl_DispCate"></table></td>
			<td valign="bottom"><input id="btnAddDispCate" type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="itemdiv" value="01" onclick="checkItemDiv(this);chgodr(1);" checked>�Ϲݻ�ǰ
      <input type="radio" name="itemdiv" value="06" onclick="checkItemDiv(this);chgodr(2);">�ֹ����ۻ�ǰ
	  <input type="checkbox" name="reqMsg" value="10" onClick="checkItemDiv(this);">�ֹ����� ���� �ʿ�<font color=red>(<b>�ֹ����� �޼���</b>�� �ʿ��� ���)</font>
<!-- 	  <br> -->
<!--       <input type="radio" name="itemdiv" value="20" onclick="checkItemDiv(this);chgodr(1);">�߰������ǰ -->
<!--       <font color="red">(��ǰ��Ͽ����� ����, �߰��ɼǿ����� ������)</font> -->
	  <input type="checkbox" name="requireimgchk" value="Y" onClick="requireimg();">�ֹ����� �̹��� �ʿ�
  	</td>
  </tr>
  <!-- �ֹ� ���� �̸��� -->
  <tr id="rmemail" style="display:none;">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ����� �̸��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="requireMakeEmail" value="" size="50" maxlength="100"> (ex)�۰����� ���� �ּ�)
  	</td>
  </tr>
  <!-- �ֹ� ���� �̸��� -->
  <tr align="left" id="customorder" style="display:none;">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ����� �߰��ɼ�</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="cstodr" value="1" onclick="chgodr2(1)" checked>��ù߼�
      <input type="radio" name="cstodr" value="2" onclick="chgodr2(2)" >������ �߼�<br>
	  <div id="subodr" style="display:none;">
		������ �߼� �Ⱓ <input type="text" name="requireMakeDay" value="" size="3" maxlength="2">��<br>
		&lt--Ư�̻����� �Է� ���ּ���--&gt;<br><textarea name="requirecontents" rows="5" cols="80"></textarea>
	  </div>
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
  	<td bgcolor="#FFFFFF">
      <input type="text" name="itemsize" maxlength="64" size="20" id="[on,off,off,off][��ǰ������]">
	  <input type="hidden" name="unit"/>
<!--       	<select name="unit"> -->
<!-- 			<option value="">�����Է�</option> -->
<!-- 			<option value="mm">mm</option> -->
<!-- 			<option value="cm" selected>cm</option> -->
<!-- 			<option value="m��">m��</option> -->
<!-- 			<option value="km">km</option> -->
<!-- 			<option value="m��">m��</option> -->
<!-- 			<option value="km��">km��</option> -->
<!-- 			<option value="ha">ha</option> -->
<!-- 			<option value="m��">m��</option> -->
<!-- 			<option value="cm��">cm��</option> -->
<!-- 			<option value="L">L</option> -->
<!-- 			<option value="g">g</option> -->
<!-- 			<option value="Kg">Kg</option> -->
<!-- 			<option value="t">t</option> -->
<!-- 		</select> -->
      &nbsp;(ex:7.5x15(cm))
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF">
      <input type="text" name="itemWeight" maxlength="12" size="4" id="[on,off,off,off][��ǰ����]" _onKeyUp="CalcuAuto(itemreg);" value="0">g
      &nbsp;(���Դ� g������ �Է�, �Ҽ����ԷºҰ�)
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
		<input type="radio" name="vatYn" value="Y" checked onclick="TnGoClear(this.form);">����
		<input type="radio" name="vatYn" value="N" onclick="TnGoClear(this.form);">�鼼
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">�⺻ ���� ���� :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="margin" maxlength="32" size="5" id="[off,off,off,off][����]" value="<% =npartner.FPartnerList(0).Fdiy_margin %>" readonly style="background-color:#E6E6E6;">%
	</td>
</tr>
<tr align="left">
  	<td height="30" bgcolor="#DDDDFF">�ǸŰ�(�Һ��ڰ�) :</td>
  	<input type="hidden" name="sellvat">
  	<td bgcolor="#FFFFFF">
      <input type="text" name="sellcash" size="12" id="[on,on,off,off][�Һ��ڰ�]" onKeyUp="CalcuAuto(itemreg);" maxlength="7">��
      <!--<input type="button" value="���ް� �ڵ����" onclick="CalcuAuto(itemreg);">-->
  	</td>
  	<td bgcolor="#DDDDFF">���ް� :</td>
  	<input type="hidden" name="buyvat">
  	<td bgcolor="#FFFFFF">
      <input type="text" name="buycash" maxlength="16" size="12" id="[on,on,off,off][���ް�]" >��
      (<b>�ΰ��� ���԰�</b>)
  	</td>
</tr>
<tr align="left">
  	<td height="30" bgcolor="#DDDDFF">��۱��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="deliverytype" value="2">������&nbsp;
      <input type="radio" name="deliverytype" value="9" checked>���ǹ��(���� ��ۺ�ΰ�)&nbsp;
      <input type="radio" name="deliverytype" value="7">���ҹ��
  	</td>
</tr>
  <input type="hidden" name="mwdiv" value="U"> <!-- ����Ư������ :��ü��� -->
  <input type="hidden" name="sellyn" value="N">
  <input type="hidden" name="isusing" value="Y">
  <input type="hidden" name="mileage" value="0">
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
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">������ ���� :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="radio" name="usinghtml" value="N" checked >�Ϲ�TEXT -->
<!--       <input type="radio" name="usinghtml" value="H">TEXT+HTML -->
<!--       <input type="radio" name="usinghtml" value="Y">HTML��� -->
<!--       <br> -->
<!--       <textarea name="itemcontent" rows="10" cols="80" id="[on,off,off,off][�����ۼ���]"></textarea> -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ��� ���ǻ��� :<br/>[��ۺ� �ȳ�]</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="ordercomment" rows="5" cols="80" id="[off,off,off,off][���ǻ���]"></textarea><br>
      <font color="red">Ư���� ��۱Ⱓ�̳� �ֹ��� Ȯ���ؾ߸� �ϴ� ����</font>�� �Է��Ͻø� ���Ҹ��̳� ȯ���� ���ϼ� �ֽ��ϴ�.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ȯ / ȯ�� ��å</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="refundpolicy" rows="5" cols="80" id="[off,off,off,off][ȯ����å]">- ��ǰ/ȯ���� ��ǰ�����Ϸκ��� 7�� �̳��� �����մϴ�. 
- ��� ���� ȯ�ҿ�û �� ��ǰ ȸ�� �� ó���˴ϴ�. 
- ���� ��ǰ�� ��� �պ���ۺ� ������ �ݾ��� ȯ�ҵǸ�, ��ǰ �� ���� ���°� ���Ǹ� �����Ͽ��� �մϴ�. 
- ��ǰ �ҷ��� ���� ��ۺ� ������ ������ ȯ�ҵ˴ϴ�.
- ����ǰ���� ���Ե� ��ǰ�� ��� A/S�� �Ұ��մϴ�. 
- ��ȯ/ȯ��/��ۺ�ȳ�/AS�� ���� ���������� ��ǰ�������� �ִ� ��� �۰����� ���������� �켱 ���� �˴ϴ�.</textarea><br>
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">��ü�ڸ�Ʈ :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="text" name="designercomment" size="50" maxlength="40" id="[off,off,off,off][��ü�ڸ�Ʈ]"><br> -->
<!--       ��ǰ������ ���丮�� ��̳� �̾߱⸦ �����ּ���... -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ ������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="90" id="[off,off,off,off][�����۵�����]"></textarea>
		<br>�� Youtube, Vimeo ������ ����(Youtube : �ҽ��ڵ尪 �Է�, Vimeo : �Ӻ����� �Է�)
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


<!-----------------------------�ɼ� ���� DIV -------------------------------->
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
      	    <td width="400"><input type="text" name="optTypeNm" value="" size="20" maxlength="16" id="[off,off,off,off][�ɼ� ���и�]"></td>
      	</tr>
      	<tr>
      	    <td colspan="2">
              <select multiple name="realopt" class="select" style="width:400px;height:120px;"></select>
              </select>
              <br>
            </td>
        </tr>
        <tr>
            <td colspan="2">
              <input type="button" value="�ɼ��߰�" name="btnetcoptadd" onclick="popEtcOptionAdd();">
              <input type="button" value="���ÿɼǻ���" name="btnoptdel" onclick="delItemOptionAdd()" >
              <br><br>
              - �ɼ��߰� : ��ǰ�ɼ��� �����Ͻ� �� �ֽ��ϴ�.<br>
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
                <input type="text" name="optionTypename<%= j+1 %>" value="" size="18" maxlength="16" id="[off,off,off,off][�ɼ� ���и�<%= j %>]">
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
                <input type="text" name="optionName<%= j+1 %>" size="18" maxlength="18" id="[off,off,off,off][�ɼǸ�<%= i %><%= j %>]">
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
 </table>
</div>

<!-----------------------------�ɼ� ���� DIV -------------------------------->




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

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�⺻�̹��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgbasic" onchange="CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg');CheckImageSize(this);" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgbasic')"><input type="hidden" name="basic"> (<font color=red>�ʼ�</font>,1000x667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,<b><font color="red">jpg</font></b>)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd1" onchange="CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd1')"><input type="hidden" name="add1"> (����,1000x667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd2" onchange="CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd2')"><input type="hidden" name="add2"> (����,1000x667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd3" onchange="CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd3')"><input type="hidden" name="add3"> (����,1000x667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>

</table>

<!-- 2016 ������ �߰� ���� -->
<!-- ǰ�� �� ���� ��ǰ����߰� -->
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
<!-- 		<option value="07">������(TV��)</option> -->
<!-- 		<option value="08">������ ������ǰ(�����/��Ź��/�ı⼼ô��/���ڷ�����)</option> -->
<!-- 		<option value="09">��������(������/��ǳ��)</option> -->
<!-- 		<option value="10">�繫����(��ǻ��/��Ʈ��/������)</option> -->
<!-- 		<option value="11">���б��(������ī�޶�/ķ�ڴ�)</option> -->
<!-- 		<option value="12">��������(MP3/���ڻ��� ��)</option> -->
<!-- 		<option value="14">������̼�</option> -->
		<option value="15">�ڵ�����ǰ(�ڵ�����ǰ/��Ÿ �ڵ�����ǰ)</option>
<!-- 		<option value="16">�Ƿ���</option> -->
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
<!-- 		<option value="27">ȣ��/��ǿ���</option> -->
<!-- 		<option value="28">�����ǰ</option> -->
<!-- 		<option value="29">�װ���</option> -->
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
</table>
<!-- ������������ -->
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
		<label><input type="radio" name="safetyYn" value="Y" onclick="chgSafetyYn(document.itemreg)"> ���</label>
		<label><input type="radio" name="safetyYn" value="N" checked onclick="chgSafetyYn(document.itemreg)"> ���ƴ�</label> /
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

<!-- �̹������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>�̹�������</strong>
		<br>- �̹����� <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> ���� �ø��� �� �ֽ��ϴ�.
		<br>&nbsp;&nbsp;(�̹�������� <font color=red>���μ������� ������</font>�� �԰ݿ� ���� �ʰ� ������ּ���. �԰��ʰ��� ����� ���� �ʽ��ϴ�.)
		<br>- <font color=red>����޿��� Save For Web����, Optimizeüũ, ������ 80%����</font>�� ����� �� �÷��ֽñ� �ٶ��ϴ�.
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>�� ������ ��ǰ�����̹����� ������� �ʰ� ��ǰ�����̹����� ����մϴ�. ������ ��ϵ� ��ǰ�����̹����� ����� �ϵ� �߰� ������ �����ʰ� ������ �˴ϴ�.</strong></font>
 	</td>
 </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="MobileimgIn">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #1 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#1 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[0],40, 1000, 667)"> (����,1000x667 ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	  <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #2 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#2 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[1],40, 1000, 667)"> (����,1000x667 ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	  <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #3 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#3 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[2],40, 1000, 667)"> (����,1000x667 ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	  <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #4 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#4 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[3],40, 1000, 667)"> (����,1000x667 ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	  <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #5 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#5 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[4],40, 1000, 667)"> (����,1000x667 ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	  <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #6 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#6 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[5],40, 1000, 667)"> (����,1000x667 ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	  <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #7 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#7 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[6],40, 1000, 667)"> (����,1000x667 ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	  <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>�� ��ǰ�󼼿��� �̹����� �߶� �÷��ֽñ� �ٶ��ϴ�.</strong></font>
 	</td>
 </tr>
  <tr align="left">
  	<td bgcolor="#FFFFFF">
      <input type="button" value="��ǰ���̹����߰�" class="button" onClick="InsertMobileImageUp()">
  	</td>
  </tr>
</table>
<!-- 2016 ������ �߰� ���� -->
</form>

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
<%
	Set npartner = Nothing
%>
<% if (application("Svr_Info")	= "Dev") or (session("ssBCTid")="fingertest01") then %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% else %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->