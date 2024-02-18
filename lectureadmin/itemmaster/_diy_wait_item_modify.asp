<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemRegCls.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<%

'CONST CBASIC_IMG_MAXSIZE = 180   'KB
'CONST CMAIN_IMG_MAXSIZE = 500   'KB
CONST CBASIC_IMG_MAXSIZE = 600   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB


'==============================================================================
Dim oitemdetail,oitemreg,optiontotal,ix
Dim oitemvideo
Dim fingerson : fingerson = "on" '//��ǰ��ÿ� fingersflag

set oitemdetail = new CWaitItemDetail

oitemdetail.FRectDesignerID = session("ssBctID")
oitemdetail.WaitProductDetail requestCheckvar(request("itemid"),10) '�ӽõ�� ������ �ҷ�����
oitemdetail.WaitProductDetailOption requestCheckvar(request("itemid"),10) '�ɼ� 2�� �ѹ�,�̸� �ҷ�����

'��ǰ�̹���
Dim itemaddimage

if (IsNull(oitemdetail.Fimgadd) or (oitemdetail.Fimgadd="")) then oitemdetail.Fimgadd = ",,,,"

itemaddimage = split(oitemdetail.Fimgadd,",")

'==============================================================================
set oitemreg = new CItemReg

'if oitemdetail.FResultCount <> 0 then
'	oitemreg.SearchOptionNameBig left(oitemdetail.FItemList(ix).Fitemoption,2) '�ɼ� 1�� �ҷ�����
'end if

oitemreg.SearchCategoryNameLarge oitemdetail.Flarge 'ī�װ� 1�� �ҷ�����
oitemreg.SearchCategoryNameMid oitemdetail.Flarge,oitemdetail.FMid 'ī�װ� 2�� �ҷ�����
oitemreg.SearchCategoryNameSmall oitemdetail.Flarge,oitemdetail.FMid,oitemdetail.Fsmall 'ī�װ� 3�� �ҷ�����
'==============================================================================
dim imgsubdir

imgsubdir = GetImageSubFolderByItemid(requestCheckvar(request("itemid"),10))
'==============================================================================
Dim npartner
set npartner = new CPartnerUser
npartner.FRectDesignerID = session("ssBctID")
npartner.GetAcademyPartnerList

'//������
Set oitemvideo = New CItem
oitemvideo.FRectItemId = requestCheckvar(request("itemid"),10)
oitemvideo.FRectItemVideoGubun = "video1"
oitemvideo.GetWaitItemContentsVideo
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script>
$(function(){
	if($("input[name='catecode']").length>1){
		$("#btnAddDispCate").hide();
	}
});

function UseTemplate() {
	window.open("/academy/comm/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
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
	var isvatYn, imileage;
	var isellcash, ibuycash, isellvat, ibuyvat, imargin;
	var isailprice, isailsuplycash, isailpricevat, isailsuplycashvat, isailmargin;

    isvatYn = frm.vatYn[0].checked;

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
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/academy/comm/pop_itemoption.asp?' + param ,'editItemOption','width=800,height=400,scrollbars=yes,resizable=yes');
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
	popwin = window.open('/academy/comm/normalitemoptionadd.asp' ,'popNormalOptionAdd','width=800,height=500,scrollbars=yes,resizable=yes');
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
function ClearImage(img) {
    var e = eval("itemreg." + img);

	if (img == "imgbasic") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg');CheckImageSize(this);\" size='40'>";
    } else if (img == "imgmain") {
       e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    } else if (img == "imgadd1") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    } else {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    }

	document.getElementById("div"+img).style.display='none';

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

function CheckImage2(img, filesize, imagewidth, imageheight, extname){
    
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

//�ӽ�����

function IMSISave() {
    if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
		return;
	}
	

    //if (validate(itemreg)==false) {
    //    return;
    //}
    
    //��ǰ�� ����üũ �߰� 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("��ǰ���� �ִ� 64byte ���Ϸ� �Է����ּ���.(�ѱ�32�� �Ǵ� ����64��)");
		itemreg.itemname.focus();
		return;
	}
    
    //��۱��� üũ ================================================================ 
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
    
    
    

    if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
        alert("���ϸ����� �ǸŰ����� Ŭ �� �����ϴ�.");
        itemreg.mileage.focus();
        return;
    }

    if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 20000000){
		alert("�Ǹ� ������ 400�� �̻� 20,000,000�� �̸����� ��� �����մϴ�.");
		itemreg.sellcash.focus();
		return;
	}
    
    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("���������� �Է����ּ���!");
        itemreg.limitno.focus();
        return;
    }
    
    // ���󰡰�
	if (confirm("�Һ��ڰ�(" + itemreg.sellcash.value + ")/���ް�(" + itemreg.buycash.value + ")�� ��Ȯ�� �ԷµǾ����ϱ�?") == false) {
		itemreg.sellcash.focus();
		return;
    }

    
    if(confirm("��ǰ�� �ӽ������Ͻðڽ��ϱ�??") == true){
        if (itemreg.itemvideo.value.length>0){
            itemreg.itemvideo.value = itemreg.itemvideo.value.replace(/iframe/gi, "BUFiframe");
        }

		itemreg.target = "FrameCKP";
        itemreg.submit();
    }
}

// ============================================================================
// �����ϱ�
function SubmitSave(istate) {
//alert('���� ���� �۾� ������ ��ǰ ���/ ������ �Ұ��մϴ�.');
//return;

	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
		return;
	}
	
	if (!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("[�⺻] ���� ī�װ��� �����ϼ���.\n");
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
    
    //��۱��� üũ ================================================================ 
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
    
    
    

    if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
        alert("���ϸ����� �ǸŰ����� Ŭ �� �����ϴ�.");
        itemreg.mileage.focus();
        return;
    }

    if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 20000000){
		alert("�Ǹ� ������ 400�� �̻� 20,000,000�� �̸����� ��� �����մϴ�.");
		itemreg.sellcash.focus();
		return;
	}
    
    if (!IsDigit(itemreg.itemWeight.value)){
		alert('���Դ� ���ڷ� �Է��ϼ���.');
		itemreg.itemWeight.focus();
		return;
	}
	
    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("���������� �Է����ּ���!");
        itemreg.limitno.focus();
        return;
    }
    
    // ���󰡰�
	if (confirm("�Һ��ڰ�(" + itemreg.sellcash.value + ")/���ް�(" + itemreg.buycash.value + ")�� ��Ȯ�� �ԷµǾ����ϱ�?") == false) {
		itemreg.sellcash.focus();
		return;
    }
    if (itemreg.basic.value == "del") {
        alert("�⺻�̹����� �ʼ��Դϴ�.");
        return;
    } else {
        if (itemreg.imgbasic.value != "") {
            if (CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg') != true) {
                return;
            }
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
//
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

    
    if(confirm("��ǰ�� �ø��ðڽ��ϱ�?") == true){
        if (itemreg.itemvideo.value.length>0){
            itemreg.itemvideo.value = itemreg.itemvideo.value.replace(/iframe/gi, "BUFiframe");
        }
        if (istate){
            itemreg.CurrState.value = istate;	// 2016/12/08
        }
		itemreg.target = "FrameCKP";
        itemreg.submit();
    }

}

// ���û�ϱ�
function SubmitReSave()
{
	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
		return;
	}

    if (!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("[�⺻] ���� ī�װ��� �����ϼ���.\n");
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
	
	//��۱��� üũ ================================================================ 
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
    
    // ���󰡰�
	if (confirm("�Һ��ڰ�(" + itemreg.sellcash.value + ")/���ް�(" + itemreg.buycash.value + ")�� ��Ȯ�� �ԷµǾ����ϱ�?") == false) {
		itemreg.sellcash.focus();
		return;
    }

    if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
        alert("���ϸ����� �ǸŰ����� Ŭ �� �����ϴ�.");
        itemreg.mileage.focus();
        return;
    }

    if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 5000000){
		alert("�Ǹ� ������ 400�� �̻� 5,000,000���� �̸����� ��� �����մϴ�.");
		itemreg.sellcash.focus();
		return;
	}

    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("���������� �Է����ּ���!");
        itemreg.limitno.focus();
        return;
    }

    if (itemreg.basic.value == "del") {
        alert("�⺻�̹����� �ʼ��Դϴ�.");
        return;
    } else {
        if (itemreg.imgbasic.value != "") {
            if (CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg') != true) {
                return;
            }
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
//
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
    
    // ���û ���� ���� �˾�
    reMsg = prompt("���� ��û�� ������ ������ ������ ���ּ���.","");
    if(reMsg){
        if (itemreg.itemvideo.value.length>0){
            itemreg.itemvideo.value = itemreg.itemvideo.value.replace(/iframe/gi, "BUFiframe");
        }
        itemreg.reRegMsg.value = reMsg;
        itemreg.CurrState.value = "5";	// ���������� '������û'���� ����
        itemreg.target = "FrameCKP"; //�߰�
        itemreg.submit();
    }
    else {
    	return;
    }

}

function TnCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// ������
		frm.limitno.readOnly=true;
		frm.limitno.style.background='#E6E6E6';

		frm.limitsold.readOnly=true;
		frm.limitsold.style.background='#E6E6E6';
	}
	else {
		// ����
		if ((frm.optioncnt.value*1) > 0) {
		    // �ɼǻ����
		    alert("�ɼ��� ����Ұ�� ���������� �ɼ�â���� ���������մϴ�.");
		    frm.limityn[0].checked = true;
		    return;
        }

		frm.limitno.readOnly=false;
		frm.limitno.style.background='#FFFFFF';

		frm.limitsold.readOnly=false;
		frm.limitsold.style.background='#FFFFFF';
	}
}

function TnSilentCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// ������
		frm.limitno.readOnly=true;
		frm.limitno.style.background='#E6E6E6';

		frm.limitsold.readOnly=true;
		frm.limitsold.style.background='#E6E6E6';
	}
	else {
		// ����
		if ((frm.optioncnt.value*1) > 0) {
		    // �ɼǻ����
		    // alert("���������� �ɼ�â���� ���������մϴ�.");
		    return;
        }

		frm.limitno.readOnly=false;
		frm.limitno.style.background='#FFFFFF';

		frm.limitsold.readOnly=false;
		frm.limitsold.style.background='#FFFFFF';
	}
}

function TnGoClear(frm){
	frm.buyvat.value = "";
}

function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if (frm.mwdiv[2].checked){
			alert("����Ư�� ������ ��ü�� ���\n��۱����� �ٹ����� ������� ���� �Ͻ� �� �����ϴ�!!\n����Ư�������� Ȯ�����ּ���!!");
			frm.mwdiv[0].checked=true;
		}
	}
	//else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
	else if(frm.deliverytype[1].checked ){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("����Ư�� ������ �����̳� Ư���� ���\n��۱�����  ��ü������� ���� �Ͻ� �� �����ϴ�!!!\n����Ư�������� Ȯ�����ּ���!!");
			frm.mwdiv[2].checked=true;
		}
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
        frm.sellcash.readOnly = false;
        frm.margin.readOnly = false;

        frm.sellcash.style.background = '#FFFFFF';
        frm.margin.style.background = '#FFFFFF';

        frm.sailprice.readOnly = true;
        frm.sailmargin.readOnly = true;

        frm.sailprice.style.background = '#E6E6E6';
        frm.sailmargin.style.background = '#E6E6E6';
	} else {
	    // ���ϰ���
        frm.sellcash.readOnly = true;
        frm.margin.readOnly = true;

        frm.sellcash.style.background = '#E6E6E6';
        frm.margin.style.background = '#E6E6E6';

        frm.sailprice.readOnly = false;
        frm.sailmargin.readOnly = false;

        frm.sailprice.style.background = '#FFFFFF';
        frm.sailmargin.style.background = '#FFFFFF';
    }
}
// ============================================================================
// �̸�����
function ViewItemDetail(itemno){
	//window.open('viewDIYitem.asp?itemid='+itemno ,'ViewItemDetail','width=790,height=600,scrollbars=yes,status=no');
	var popwin = window.open('/academy/itemmaster/viewDIYitem/viewDIYitem.asp?itemid='+itemno ,'window1','width=1024,height=960,scrollbars=yes,status=no');
	popwin.focus();
}

function ClearVal(comp){
    comp.value = "";
}

function popWaitItemOptionEdit(iitemid){
    var popwin = window.open('/academy/comm/pop_DIYwaititemoptionedit.asp?itemid=' + iitemid,'popWaitItemOptionEdit','width=790,height=600,scrollbars=yes,status=no');
    popwin.focus();
}

//=====������
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
			url: "/admin/itemmaster/act_waititemInfoDivForm.asp",
			data: "itemid=<%=request("itemid")%>&ifdv="+v+"&fingerson=on",
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


//��ǰ�����̹����߰�
function InsertImageUp() {
	var f = document.all;
	var rowLen = f.imgIn.rows.length;

	if(rowLen > 14){
		alert("�̹����� �ִ� 15������ �����մϴ�.");
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
	c0.innerHTML = '��ǰ���̹��� #' + rowLen + ' :';
	c1.style.background = '#FFFFFF';
	c1.innerHTML = '<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40, '+parseInt(rowLen-1)+');CheckImageSize(this);" class="text" size="40"> ';
	c1.innerHTML += '<input type="button" value="#'+parseInt(rowLen)+' �̹��������" class="button" onClick="ClearImage2(this.form.addimgname['+parseInt(rowLen-1)+'],40, 1000, 667, '+parseInt(rowLen-1)+')"> (����,1000X667, Max 600KB,jpg,gif)';
	c1.innerHTML += '<br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>';
	c1.innerHTML += '<input type="hidden" name="addimggubun" value="'+parseInt(rowLen)+'">';
	c1.innerHTML += '<input type="hidden" name="addimgdel" value="">';
}

//��ǰ���̹��������
function ClearImage2(img,fsize,wd,ht,num) {
	var imgcnt = $('input[name="addimgname"]').length;
    img.outerHTML = "<input type='file' name='" + img.name + "' onchange=\"CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +", "+num+");CheckImageSize(this);\" class='text' size='"+ fsize +"'>";
	$("#divaddimgname"+(num+1)+"").hide();
	
	if(imgcnt > 1){
    	document.itemreg.addimgdel[num].value = "del";
    }else{
    	document.itemreg.addimgdel.value = "del";
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>��ǰ����</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>��ϴ������ ��ǰ�� �����մϴ�.</b>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>

</table>
<!-- ǥ ��ܹ� ��-->


<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>�⺻����
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->
<% if (TRUE) or (session("ssBCTid")="fingertest01") then %>
<form name="itemreg" method="post" action="<%= uploadImgUrl %>/linkweb/academy/items/WaitDIYItemModify_Process.asp" onsubmit="return false;" enctype="multipart/form-data" style="margin:0px;">
<% else %>
<form name="itemreg" method="post" action="<%= UploadImgFingers %>/linkweb/items/WaitDIYItemModify_Process.asp" onsubmit="return false;" enctype="multipart/form-data" style="margin:0px;">
<% end if %>
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="itemid" value="<%= request("itemid") %>">
<input type="hidden" name="designerid" value="<%= oitemdetail.FRectDesignerID %>">
<input type="hidden" name="defultmargine" value="<% =npartner.FPartnerList(0).Fdiy_margin %>">
<input type="hidden" name="defaultmaeipdiv" value="<%= npartner.FPartnerList(0).Fmaeipdiv %>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%= npartner.FPartnerList(0).FdefaultFreeBeasongLimit %>">
<input type="hidden" name="defaultDeliverPay" value="<%= npartner.FPartnerList(0).FdefaultDeliverPay %>">
<input type="hidden" name="defaultDeliveryType" value="<%= npartner.FPartnerList(0).FdefaultDeliveryType %>">

<input type="hidden" name="isusing" value="N">
<input type="hidden" name="dispyn" value="N">
<input type="hidden" name="sellyn" value="N">
<input type="hidden" name="reRegMsg" value="">
<input type="hidden" name="CurrState" value="<%=oitemdetail.FCurrState%>">

<input type="hidden" name="cd1" value="<%= oitemdetail.Flarge %>">
<input type="hidden" name="cd2" value="<%= oitemdetail.Fmid %>">
<input type="hidden" name="cd3" value="<%= oitemdetail.Fsmall %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�ڵ� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <%= request("itemid") %>
  	  &nbsp;&nbsp;&nbsp;&nbsp;
  	  <input type="button" value="�̸�����" onclick="ViewItemDetail('<%= request("itemid") %>');">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">ī�װ� ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getDispCategoryWait(trim(request("itemid")))%></td>
			<td valign="bottom"><input id="btnAddDispCate" type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="itemdiv" value="01" <% if oitemdetail.Fitemdiv ="01" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr(1);">�Ϲݻ�ǰ
	  <input type="radio" name="itemdiv" value="06" <%=chkIIF(oitemdetail.Fitemdiv="06" or oitemdetail.Fitemdiv="16","checked","")%> onclick="checkItemDiv(this);chgodr(2);">�ֹ����ۻ�ǰ
      <input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitemdetail.Fitemdiv="06","checked","")%> <%=chkIIF(oitemdetail.Fitemdiv="06" or oitemdetail.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">�ֹ����� ���� �ʿ�<font color=red>(�ֹ����� �޼����� �ʿ��� ���)</font>
	  <input type="checkbox" name="requireimgchk" value="Y" <%=chkIIF(oitemdetail.Frequirechk="Y","checked","")%> onClick="requireimg();">�ֹ����� �̹��� �ʿ�
<!-- 	  <br> -->
<!--       <input type="radio" name="itemdiv" value="20" <% if oitemdetail.Fitemdiv ="20" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr(1);">�߰������ǰ -->
<!--       <font color="red">(��ǰ��Ͽ����� ����, �߰��ɼǿ����� ������)</font><br> -->
  	</td>
  </tr>
   <!-- �ֹ� ���� �̸��� -->
  <tr id="rmemail" style="display:<%=chkiif(oitemdetail.Frequirechk="Y","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ����� �̸��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="requireMakeEmail" value="<%=oitemdetail.FrequireEmail%>" size="50" maxlength="100"> (ex)�۰����� ���� �ּ�)
  	</td>
  </tr>
  <!-- �ֹ� ���� �̸��� -->
  <tr align="left" id="customorder" style="display:<%=chkiif(oitemdetail.Fitemdiv="06" Or oitemdetail.Fitemdiv="16","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ����� �߰��ɼ�</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="cstodr" value="1" onclick="chgodr2(1)" <%=chkiif(oitemdetail.Fcstodr="1","checked","")%>>��ù߼�
      <input type="radio" name="cstodr" value="2" onclick="chgodr2(2)" <%=chkiif(oitemdetail.Fcstodr="2","checked","")%>>������ �߼�<br>
	  <div id="subodr" style="display:<%=chkiif(oitemdetail.Fcstodr="2","block","none")%>;">
		������ �߼� �Ⱓ <input type="text" name="requireMakeDay" value="<%=oitemdetail.FrequireMakeDay%>" size="3" maxlength="2">��<br>
		&lt--Ư�̻����� �Է� ���ּ���--&gt;<br><textarea name="requirecontents" rows="5" cols="80"><%=oitemdetail.Frequirecontents%></textarea>
	  </div>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemname" maxlength="64" size="50" id="[on,off,off,off][��ǰ��]" value="<%= oitemdetail.Fitemname %>">&nbsp;
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsource" maxlength="64" size="50" id="[on,off,off,off][��ǰ����]" value="<%= oitemdetail.Fitemsource %>">&nbsp;(ex:�ö�ƽ,����,��,...)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsize" maxlength="64" size="50" id="[on,off,off,off][��ǰ������]" value="<%= oitemdetail.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemWeight" maxlength="12" size="4" id="[on,off,off,off][��ǰ����]" value="<%= oitemdetail.FitemWeight %>">g&nbsp;(���Դ� g������ �Է�)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="sourcearea" maxlength="64" size="25" id="[on,off,off,off][������]" value="<%= oitemdetail.Fsourcearea %>">&nbsp;(ex:�ѱ�,�߱�,�߱�OEM,�Ϻ�...)
      <br>( ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="makername" maxlength="32" size="25" id="[on,off,off,off][������]" value="<%= oitemdetail.Fmakername %>">&nbsp;(������ü��)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�˻�Ű���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="keywords" maxlength="50" size="50" id="[on,off,off,off][�˻�Ű����]" value="<%= oitemdetail.Fkeywords %>">&nbsp;(�޸��α��� ex: Ŀ��,Ƽ����,����)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ü��ǰ�ڵ� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	    <input type="text" name="upchemanagecode" value="<%= oitemdetail.Fupchemanagecode %>" size="20" maxlength="32" id="[off,off,off,off][��ü��ǰ�ڵ�]">
  	    (��ü���� �����ϴ� �ڵ� �ִ� 32�� - ����/���ڸ� ����)
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ ���� :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="radio" name="usinghtml" value="N" <% if oitemdetail.Fusinghtml = "N" then response.write "checked" %>>�Ϲ�TEXT -->
<!--       <input type="radio" name="usinghtml" value="H" <% if oitemdetail.Fusinghtml = "H" then response.write "checked" %>>TEXT+HTML -->
<!--       <input type="radio" name="usinghtml" value="Y" <% if oitemdetail.Fusinghtml = "Y" then response.write "checked" %>>HTML��� -->
<!--       <br> -->
<!--       <textarea name="itemcontent" rows="10" cols="80" id="[on,off,off,off][�����ۼ���]"><%= oitemdetail.Fitemcontent %></textarea> -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ��� ���ǻ��� :<br/>[��ۺ� �ȳ�]</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="ordercomment" rows="5" cols="80" id="[off,off,off,off][���ǻ���]"><%= oitemdetail.Fordercomment %></textarea><br>
      <font color="red">Ư���� ��۱Ⱓ�̳� �ֹ��� Ȯ���ؾ߸� �ϴ� ����</font>�� �Է��Ͻø� ���Ҹ��̳� ȯ���� ���ϼ� �ֽ��ϴ�.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ȯ / ȯ�� ��å</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="refundpolicy" rows="5" cols="80" id="[off,off,off,off][ȯ����å]"><%=oitemdetail.Frefundpolicy%></textarea><br>
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">��ü�ڸ�Ʈ :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="text" name="designercomment" size="50" maxlength="40" id="[off,off,off,off][��ü�ڸ�Ʈ]" value="<%= oitemdetail.Fdesignercomment %>"><br> -->
<!--       ��ǰ������ ���丮�� ��̳� �̾߱⸦ �����ּ���... -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ ������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="90" id="[off,off,off,off][�����۵�����]"><%= db2html(oitemvideo.FOneItem.FvideoFullUrl) %></textarea>
		<br>�� Youtube, Vimeo ������ ����(Youtube : �ҽ��ڵ尪 �Է�, Vimeo : �Ӻ����� �Է�)
	</td>
  </tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>��������
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <input type="hidden" name="sellvat" value="<%= oitemdetail.Fsellvat %>">
  <input type="hidden" name="buyvat" value="<%= oitemdetail.Fbuyvat %>">
  <tr align="left">
    <td width="15%" bgcolor="#DDDDFF">����, �鼼 ���� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="radio" name="vatYn" value="Y" onclick="TnGoClear(this.form);" <% if oitemdetail.FvatYn = "Y" then response.write "checked" %>>����
      <input type="radio" name="vatYn" value="N" onclick="TnGoClear(this.form);" <% if oitemdetail.FvatYn = "N" then response.write "checked" %>>�鼼
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">���� ���� :</td>
  	<td bgcolor="#FFFFFF">
  	    <% if (oitemdetail.Fsellcash=0) then ''2016/12/08 �߰� �ӽ������ǰ �ǸŰ� 0 ����. %>
  	    <input type="text" name="margin" maxlength="32" size="5" id="[off,off,off,off][����]" value="<%= npartner.FPartnerList(0).Fdiy_margin %>" readonly style="background-color:#E6E6E6;">%
  	    <% else %>
        <input type="text" name="margin" maxlength="32" size="5" id="[off,off,off,off][����]" value="<%= oitemdetail.FMargin %>" readonly style="background-color:#E6E6E6;">%
        <% end if %>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ǸŰ�(�Һ��ڰ�) :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="sellcash" size="12" id="[on,on,off,off][�Һ��ڰ�]" onKeyUp="CalcuAuto(itemreg);" maxlength="7" value="<%= oitemdetail.Fsellcash %>" >��
      <!--<input type="button" value="���ް� �ڵ����" onclick="CalcuAuto(itemreg);">-->
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">���ް� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="buycash" maxlength="16" size="12" id="[on,on,off,off][���ް�]" value="<%= oitemdetail.Fbuycash %>" >��
      (<b>�ΰ��� ���԰�</b>)
  	</td>
  </tr>
  <tr>
  	<td bgcolor="#DDDDFF"></td>
  	<td bgcolor="#FFFFFF" colspan="3">
      - ���ް��� <b>�ΰ��� ���԰�</b>�Դϴ�.<br>
      - �Һ��ڰ�(���ΰ�)�� ����(���θ���)�� �Է��ϰ� [���ް��ڵ����] ��ư�� ������ ���ް��� ���ϸ����� �ڵ����˴ϴ�.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��۱��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitemdetail.Fdeilverytype = "2" then response.write "checked" %>>��ü(����)���&nbsp;
      <input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitemdetail.Fdeilverytype = "9" then response.write "checked" %>>��ü���ǹ��(���� ��ۺ�ΰ�)&nbsp;
  	  <input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitemdetail.Fdeilverytype = "7" then response.write "checked" %>>��ü���ҹ��
  	</td>
  </tr>
  <input type="hidden" name="mileage" id="[on,off,off,off][���ϸ���]" value="<%= oitemdetail.Fmileage %>">
  <input type="hidden" name="mwdiv" value="U">
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>�ɼ�����/��������
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ɼǱ��� :</td>
  	<input type="hidden" name="optioncnt" value="<%= oitemdetail.Foptioncnt %>">
  	<td width="35%" bgcolor="#FFFFFF">
      <% if oitemdetail.Foptioncnt < 1 then %>
      �ɼǻ�����
      <% else %>
      �ɼǻ����(<%= oitemdetail.Foptioncnt %>��)
      <% end if %>
      &nbsp;&nbsp;<input type="button" class="button" value="�ɼǼ���" onClick="popWaitItemOptionEdit('<%= oitemdetail.FWaitItemID %>');">
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">�����Ǹű��� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	  <input type="radio" name="limityn" value="N" onClick="TnCheckLimitYN(itemreg)" <% if oitemdetail.Flimityn = "N" then response.write "checked" %>>�������Ǹ�&nbsp;&nbsp;
  	  <input type="radio" name="limityn" value="Y" onClick="TnCheckLimitYN(itemreg)" <% if oitemdetail.Flimityn = "Y" then response.write "checked" %>>�����Ǹ�
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�������� :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="limitno" maxlength="32" size="8" readOnly style="background-color:#E6E6E6;" id="[off,on,off,off][��������]" value="<%= oitemdetail.Flimitno %>">
      <input type="hidden" name="limitsold" value="0">
      <input type="hidden" name="limitstock" value="<%= oitemdetail.Flimitno %>">
  	</td>
  </tr>
  <tr align="left">
  	<td width="15%" bgcolor="#DDDDFF">�ɼǼ��� :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <br>
      - �ɼ������� �ɼ�â���� ���������մϴ�.<br>
      - �ɼ��� �߰��� ���������� ������ �Ұ����մϴ�. ��Ȯ�� �Է��ϼ���.<br>
      - ���������� �ɼ��� ���� ���, �ɼ�â���� ������ �����մϴ�.(���� ������ ����Ȯ�Ҽ� �ֽ��ϴ�.)<br>
      <br>
  	</td>
  </tr>
</table>


<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>�̹�������
          <br>- �ٹ����ٿ��� �̹����� ����� ��� ���� �Է����� ���ñ� �ٶ��ϴ�.
          <br>- �̹����� <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> ���� �ø��� �� �ֽ��ϴ�.
          <br>&nbsp;&nbsp;(�̹�������� <font color=red>���μ������� ������</font>�� �԰ݿ� ���� �ʰ� ������ּ���. �԰��ʰ��� ����� ���� �ʽ��ϴ�.)
          <br>- <font color=red>����޿��� Save For Web����, Optimizeüũ, ������ 80%����</font>�� ����� �� �÷��ֽñ� �ٶ��ϴ�.
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�⺻�̹��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
<% if (oitemdetail.Fimgbasic <> "") then %>
      <div id="divimgbasic" style="display:block;">
      <table id="imgbasic"  style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="<%=UploadImgFingers%>/diyItem/waitimage/basic/<%= imgsubdir  %>/<%= oitemdetail.Fimgbasic %>">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgbasic" style="display:none;">
      <table id="imgbasic" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% end if %>
      <input type="file" name="imgbasic" onchange="CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg');CheckImageSize(this);" size="40"> (<font color=red>�ʼ�</font>,1000X667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,jpg)
      <input type="button" value="�̹��������" onClick="ClearImage('imgbasic')"><input type="hidden" name="basic">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�������̹���<br>(�ڵ�����) :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	�Ǽ��� ��Ͻ� �ڵ�����
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
<% if (itemaddimage(0) <> "") then %>
      <div id="divimgadd1" style="display:block;">
      <table id="imgadd1" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="<%=UploadImgFingers%>/diyItem/waitimage/add1/<%= imgsubdir  %>/<%= itemaddimage(0) %>">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgadd1" style="display:none;">
      <table id="imgadd1" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%"" background="">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% end if %>
      <input type="file" name="imgadd1" onchange="CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40"> (����,1000X667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd1')"><input type="hidden" name="add1">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
<% if (itemaddimage(1) <> "") then %>
      <div id="divimgadd2" style="display:block;">
      <table id="imgadd2" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="<%=UploadImgFingers%>/diyItem/waitimage/add2/<%= imgsubdir  %>/<%= itemaddimage(1) %>">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgadd2" style="display:none;">
      <table id="imgadd2" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% end if %>
      <input type="file" name="imgadd2" onchange="CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40"> (����,1000X667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd2')"><input type="hidden" name="add2">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
<% if (itemaddimage(2) <> "") then %>
      <div id="divimgadd3" style="display:block;">
      <table id="imgadd3" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="<%=UploadImgFingers%>/diyItem/waitimage/add3/<%= imgsubdir  %>/<%= itemaddimage(2) %>">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgadd3" style="display:none;">
      <table id="imgadd3" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% end if %>
      <input type="file" name="imgadd3" onchange="CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40"> (����,1000X667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd3')"><input type="hidden" name="add3">
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���4 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- <% if (itemaddimage(3) <> "") then %> -->
<!--       <div id="divimgadd4" style="display:block;"> -->
<!--       <table id="imgadd4" width="400" height="400" style="background-repeat: no-repeat" background="<%=UploadImgFingers%>/diyItem/waitimage/add4/<%= imgsubdir  %>/<%= itemaddimage(3) %>"> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% else %> -->
<!--       <div id="divimgadd4" style="display:none;"> -->
<!--       <table id="imgadd4" width="400" height="400" style="background-repeat: no-repeat" background=""> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% end if %> -->
<!--       <input type="file" name="imgadd4" onchange="CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');" size="40"> (����,1000X667,jpg,gif) -->
<!--       <input type="button" value="�̹��������" onClick="ClearImage('imgadd4')"><input type="hidden" name="add4"> -->
<!--   	</td> -->
<!--   </tr> -->
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���5 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- <% if (itemaddimage(4) <> "") then %> -->
<!--       <div id="divimgadd5" style="display:block;"> -->
<!--       <table id="imgadd5" width="400" height="400" style="background-repeat: no-repeat" background="<%=UploadImgFingers%>/diyItem/waitimage/add5/<%= imgsubdir  %>/<%= itemaddimage(4) %>"> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% else %> -->
<!--       <div id="divimgadd5" style="display:none;"> -->
<!--       <table id="imgadd5" width="400" height="400" style="background-repeat: no-repeat" background=""> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% end if %> -->
<!--       <input type="file" name="imgadd5" onchange="CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');" size="40"> (����,1000X667,jpg,gif) -->
<!--       <input type="button" value="�̹��������" onClick="ClearImage('imgadd5')"><input type="hidden" name="add5"> -->
<!--   	</td> -->
<!--   </tr> -->
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- <% if (oitemdetail.Fimgmain <> "") then %> -->
<!--       <div id="divimgmain" style="display:block;"> -->
<!--       <table id="imgmain" width="400" height="400" style="background-repeat: no-repeat" background="<%=UploadImgFingers%>/diyItem/waitimage/main/<%= imgsubdir  %>/<%= oitemdetail.Fimgmain %>"> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% else %> -->
<!--       <div id="divimgmain" style="display:none;"> -->
<!--       <table id="imgmain" width="400" height="400" style="background-repeat: no-repeat" background=""> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% end if %> -->
<!--       <input type="file" name="imgmain" onchange="CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');" size="40"> (����,600X2000 Max <%= CMAIN_IMG_MAXSIZE %>Kb ����,jpg) -->
<!--       <input type="button" value="�̹��������" onClick="ClearImage('imgmain')"><input type="hidden" name="main"> -->
<!--   	</td> -->
<!--   </tr> -->

</table>

<!-- ǰ������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">ǰ������� &nbsp;<font color=gray>��ǰ����������� ���� ���� ������ ���� �Ʒ� ������ ��Ȯ�� �Է����ֽñ� �ٶ��ϴ�.</font></td>
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
		<!--
		<option value="07">������(TV��)</option>
		<option value="08">������ ������ǰ(�����/��Ź��/�ı⼼ô��/���ڷ�����)</option>
		<option value="09">��������(������/��ǳ��)</option>
		<option value="10">�繫����(��ǻ��/��Ʈ��/������)</option>
		<option value="11">���б��(������ī�޶�/ķ�ڴ�)</option>
		<option value="12">��������(MP3/���ڻ��� ��)</option>
		<option value="14">������̼�</option>
		-->
		<option value="15">�ڵ�����ǰ(�ڵ�����ǰ/��Ÿ �ڵ�����ǰ)</option>
		<!-- <option value="16">�Ƿ���</option>-->
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
		<!--<option value="27">ȣ��/��ǿ���</option>
		<option value="28">�����ǰ</option>
		<option value="29">�װ���</option>
		-->
		<option value="35">��Ÿ</option>
		</select>
		<script type="text/javascript">
		document.itemreg.infoDiv.value="<%=oitemdetail.FinfoDiv%>";
		//setTimeout(function(){
			chgInfoDiv(<%=oitemdetail.FinfoDiv%>);
		//},0);
		</script>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitemdetail.FinfoDiv) or oitemdetail.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ�񳻿� :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitemdetail.FinfoDiv) or oitemdetail.FinfoDiv="") Then
			Server.Execute("/admin/itemmaster/act_waititemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">��ǰ���������� ������ ���� �Ǿ��ִ��� ��Ȯ�� �Է¹ٶ��ϴ�. ����Ȯ�ϰų� �߸��� ���� �Է½�, �׿� ���� å���� ���� ���� �ֽ��ϴ�.</font></td>
</tr>
<!-- <tr align="left" id="lyItemSrc" _style="display:<%=chkIIF(oitemdetail.FinfoDiv="35","","none")%>;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitemdetail.Fitemsource %>">&nbsp;(ex:�ö�ƽ,����,��,...) -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr align="left" id="lyItemSize" _style="display:<%=chkIIF(oitemdetail.FinfoDiv="35","","none")%>;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ������ :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitemdetail.Fitemsize %>">&nbsp;(ex:7.5x15(cm)) -->
<!-- 	</td> -->
<!-- </tr> -->
</table>
<!-- ������������ -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">������������</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">����������� :</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitemdetail.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)"> ���</label>
		<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","checked","")%> onclick="chgSafetyYn(document.itemreg)"> ���ƴ�</label> /
		<select name="safetyDiv" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","disabled","")%> class="select">
		<option value="">::������������::</option>
		<option value="10" <%=chkIIF(oitemdetail.FsafetyDiv="10","selected","")%>>������������(KC��ũ)</option>
		<option value="20" <%=chkIIF(oitemdetail.FsafetyDiv="20","selected","")%>>�����ǰ ��������</option>
		<option value="30" <%=chkIIF(oitemdetail.FsafetyDiv="30","selected","")%>>KPS �������� ǥ��</option>
		<option value="40" <%=chkIIF(oitemdetail.FsafetyDiv="40","selected","")%>>KPS �������� Ȯ�� ǥ��</option>
		<option value="50" <%=chkIIF(oitemdetail.FsafetyDiv="50","selected","")%>>KPS ��� ��ȣ���� ǥ��</option>
		</select>
		������ȣ <input type="text" name="safetyNum" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","disabled","")%> size="35" maxlength="25" class="text" value="<%=oitemdetail.FsafetyNum%>" />
		
		<font color="darkred">���ƿ�ǰ�̳� �����ǰ�� ��� �ʼ� �Է�</font>
	</td>
</tr>
</table>

<%
	Dim cImg, k, vArr, j, txtBuf
	set cImg = new CItemAddImage
	cImg.FRectItemID = request("itemid")
	vArr = cImg.GetWaitAddImageList
%>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
	<% If isArray(vArr) Then
			If vArr(3,UBound(vArr,2)) > 0 Then
			For k = 1 To vArr(3,UBound(vArr,2))
	%>
			  <tr align="left">
			  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #<%= (k) %> :</td>
			  	<td bgcolor="#FFFFFF">
		  		<%
		  		If cImg.IsImgExist(vArr,k) Then
		    		For j = 0 To UBound(vArr,2)
		    			If CStr(vArr(3,j)) = CStr(k) AND (vArr(4,j) <> "" and isNull(vArr(4,j)) = False) Then
							Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:block;""><img src=""" & UploadImgFingers & "/diyItem/waitcontentsimage/" & GetImageSubFolderByItemid(vArr(1,j)) & "/" & vArr(4,j) & """ width=""400""></div>"
							Exit For
		    			End If
		    		Next
				Else
					Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:none;""></div>"
				End If
				%>
			      <input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40, <%= (k-1) %>);CheckImageSize(this);" class="text" size="40">
			      <input type="button" value="#<%= (k) %> �̹��������" class="button" onClick="ClearImage2(this.form.addimgname<%=CHKIIF(vArr(3,UBound(vArr,2))=1,"","["&(k-1)&"]")%>,40, 1000, 667, <%= (k-1) %>)"> (����,1000X667, Max 800KB,jpg,gif)
				  <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/>
				  <%
				  txtBuf=""
				  For j = 0 To UBound(vArr,2)
	    			If CStr(vArr(3,j)) = CStr(k) Then
	    			    txtBuf = db2html(vArr(5,j))
						Exit For
	    			End If
	    		  Next
	    		  %>
				  <textarea name="addimgtext" cols="70" rows="5"><%=txtBuf%></textarea>
			      <input type="hidden" name="addimggubun" value="<%= (k) %>">
			      <input type="hidden" name="addimgdel" value="">
			  	</td>
			  </tr>
	<%
			Next
			End IF
		Else
	%>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #1 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname1" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,0);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#1 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[0],40, 1000, 667, 0)"> (����,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="1">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #2 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname2" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,1);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#2 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[1],40, 1000, 667, 1)"> (����,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="2">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #3 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname3" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,2);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#3 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[2],40, 1000, 667, 2)"> (����,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="3">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #4 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname4" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,3);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#4 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[3],40, 1000, 667, 3)"> (����,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="4">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #5 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname5" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,4);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#5 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[4],40, 1000, 667, 4)"> (����,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="5">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #6 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname6" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,5);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#6 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[5],40, 1000, 667, 5)"> (����,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="6">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #7 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname7" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,6);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#7 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[6],40, 1000, 667, 6)"> (����,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="7">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
	<%
	   End IF %>
</table>
<%	set cImg = nothing %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td bgcolor="#FFFFFF" height="30">
      <input type="button" value="��ǰ���̹����߰�" class="button" onClick="InsertImageUp()">
      <font color="red">* ���ε尡 �� �̹����� ����� �ȳ����� ���ΰ�ħ(CTRL + F5(��Ʈ�� F5 ��ư))�� ���ּ���.</font>
  	</td>
  </tr>
</table>

</form>


    <% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"8" then %>
    <table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
      <tr align="left">
      	<td height="30" width="15%" bgcolor="#F8DDFF">��Ϻ��� ���� :</td>
      	<td bgcolor="#FFFFFF" colspan="3">
      		<%=oitemdetail.Frejectmsg & " [" & oitemdetail.FrejectDate & "]"%>
      	</td>
      </tr>
    </table>
    <% end if %>
    <% if oitemdetail.FCurrState="5" and Not(isNull(oitemdetail.FreRegMsg)) then %>
    <table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
      <tr align="left">
      	<td height="30" width="15%" bgcolor="#F8DDFF">���û �޽��� :</td>
      	<td bgcolor="#FFFFFF" colspan="3">
      		<%=oitemdetail.FreRegMsg & " [" & oitemdetail.FreRegDate & "]"%>
      	</td>
      </tr>
    </table>
    <% end if %>
    
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <% if (oitemdetail.FCurrState="8") then %>
          <input type="button" value="�ӽ�����" onClick="IMSISave()">
          <input type="button" value="��Ͽ�û" onClick="SubmitSave('1')">
          <% else %>
              <% if oitemdetail.FCurrState="1" or oitemdetail.FCurrState="5" then %>
              <input type="button" value="�����ϱ�" onClick="SubmitSave('<%=oitemdetail.FCurrState%>')">
              <% else %>
              <input type="button" value="���û�ϱ�" onClick="SubmitReSave()">
              <% end if %>
          <% end if %>
          <input type="button" value="�������" onClick="history.back()">
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
<%
	set oitemreg = Nothing
	Set oitemvideo = Nothing
	set oitemdetail = Nothing
	set npartner = Nothing
%>
<p>
<script>
// ����
TnSilentCheckLimitYN(itemreg);
// ����
// TnCheckSailYN(itemreg);
</script>
<% if (application("Svr_Info")	= "Dev") or (session("ssBctId")="fingertest01") then %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% else %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->