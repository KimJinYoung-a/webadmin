<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemregcls.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<!-- #include virtual="/lib/classes/items/waititemcls_2014.asp"--> 
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB
dim arrold
Dim clsWait, itemid ,makerid,arrlist, intLoop
makerid	= requestCheckvar(Request("designer"),32) 
itemid =  requestCheckvar(Request("itemid"),16)
'==============================================================================
Dim oitemdetail,oitemreg,optiontotal,ix,ooimage

set oitemdetail = new CWaitItemDetail

oitemdetail.FRectDesignerID = makerid
oitemdetail.WaitProductDetail itemid '�ӽõ�� ������ �ҷ�����
oitemdetail.WaitProductDetailOption itemid '�ɼ� 2�� �ѹ�,�̸� �ҷ�����


'==============================================================================
set oitemreg = new CItemReg

'if oitemdetail.FResultCount <> 0 then
'	oitemreg.SearchOptionNameBig left(oitemdetail.FItemList(ix).Fitemoption,2) '�ɼ� 1�� �ҷ�����
'end if

oitemreg.SearchCategoryNameLarge oitemdetail.Flarge 'ī�װ� 1�� �ҷ�����
oitemreg.SearchCategoryNameMid oitemdetail.Flarge,oitemdetail.FMid 'ī�װ� 2�� �ҷ�����
oitemreg.SearchCategoryNameSmall oitemdetail.Flarge,oitemdetail.FMid,oitemdetail.Fsmall 'ī�װ� 3�� �ҷ�����



'==============================================================================
set ooimage = new CWaitItemImagelist
ooimage.WaitProductImageList itemid  '�̹��� ������ �ҷ�����

Dim itemaddimage,itemaddcontent, itemstoryimage

if (IsNull(ooimage.Fimgadd) or (ooimage.Fimgadd="")) then ooimage.Fimgadd = ",,,,"
if (IsNull(ooimage.Fitemaddcontent) or (ooimage.Fitemaddcontent="")) then ooimage.Fitemaddcontent = "||||"
if (IsNull(ooimage.Fimgstory) or (ooimage.Fimgstory="")) then ooimage.Fimgstory = ",,,,"


itemaddimage = split(ooimage.Fimgadd,",")
itemaddcontent = split(ooimage.Fitemaddcontent,"|")
itemstoryimage = split(ooimage.Fimgstory,",")


'==============================================================================
dim imgsubdir

imgsubdir = GetImageSubFolderByItemid(itemid)


'==============================================================================
Dim npartner
Dim npt_defaultmargine, npt_defaultFreeBeasongLimit, npt_defaultDeliverPay, npt_defaultDeliveryType
set npartner = new CPartnerUser
npartner.FRectDesignerID = oitemdetail.FMakerid

if Not(oitemdetail.FMakerid="" or isNull(oitemdetail.FMakerid)) then
	npartner.GetOnePartnerNUser
	
	npt_defaultmargine	 = npartner.FOneItem.Fdefaultmargine
	npt_defaultFreeBeasongLimit	= npartner.FOneItem.FdefaultFreeBeasongLimit
	npt_defaultDeliverPay	= npartner.FOneItem.FdefaultDeliverPay
	npt_defaultDeliveryType	= npartner.FOneItem.FdefaultDeliveryType
end if
set npartner = Nothing

'--- ����������� 
 set clsWait = new CWaitItemlist2014
 	clsWait.Fitemid = itemid
 	arrList = clsWait.fnGetWaitItemLog
 	IF not isArray(arrList) THEN
 		arrOld = clsWait.fnGetOldWaitItemLog
	END IF
 set clsWait = nothing
%> 
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style>
	FORM {display:inline;}  
	#dialog {display:none; position:absolute;z-index:9100;background:#efefef; padding:10px;width:650;}
	#mask {display:none; position:absolute; left:0; top:0; z-index:9000; background:url(http://fiximage.10x10.co.kr/web2013/common/mask_bg.png) left top repeat;}
	#boxes .window {position:fixed; left:0; top:0; display:none; z-index:99999;}
</style> 
<script type="text/javascript">
function printItemAttribute() {
	var arrDispCd="";
	$("input[name='catecode']").each(function(i){
		if(i>0) arrDispCd += ",";
		arrDispCd += $(this).val();
	});
	$.ajax({
		url: "/common/module/act_waitItemAttribSelect.asp?itemid=<%=request("itemid")%>&arrDispCate="+arrDispCd,
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

// ============================================================================
// ��ü�����ڵ��Է�
function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value = varArray[0];
	document.itemreg.margin.value = varArray[1];

}

function CalcuAuto(frm){
	var isvatinclude, imileage;
	var isellcash, ibuycash, isellvat, ibuyvat, imargin;
	var isailprice, isailsuplycash, isailpricevat, isailsuplycashvat, isailmargin;

    isvatinclude = frm.vatinclude[0].checked;

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
		isellvat = parseInt(parseInt(1/11 * parseInt(isellcash)));
		ibuycash = isellcash - Math.round(isellcash*imargin/100);       //parseInt-> round�� ���� 
		ibuyvat = parseInt(parseInt(1/11 * parseInt(ibuycash)));
		imileage = parseInt(isellcash*0.005) ;
	}else{
		isellvat = 0;
		ibuycash = isellcash - Math.round(isellcash*imargin/100);       //parseInt-> round�� ���� 
		ibuyvat = 0;
		imileage = parseInt(isellcash*0.005) ;
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
}

$(function(){
	// �ε��� ��ǰ�Ӽ� ���� ���
	printItemAttribute();
});

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
function ClearImage(img,fsize,wd,ht) {
	img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";

    document.getElementById("div"+ img.name).style.display = "none";

	var e = eval("itemreg."+img.name.substr(3,img.name.length));
	e.value = "del";
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
        ClearImage(img,fsize, imagewidth, imageheight);
        return false;
    }

	var e = eval("itemreg."+img.name.substr(3,img.name.length));
	e.value = "";

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
function SubmitSave(processstatus) {
	var optionv="";
	var optiont = "";

	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }

	if (processstatus==true&&!$("input[name='isDefault'][value='y']").length){
		//alert("���� ī�װ��� �����ϼ���.\n�� ���� �⺻ ī�װ��� �ʼ� �ֽ��ϴ�.");
		//return;
	}

    //��ǰ�� ����üũ �߰� 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("��ǰ���� �ִ� 64byte ���Ϸ� �Է����ּ���.(�ѱ�32�� �Ǵ� ����64��)");
		itemreg.itemname.focus();
		return;
	}
	

    //��۱��� üũ =========================================================================
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
    
    //��۱��� ��ü�̳� ���Ա����� ��ü�� �ƴѰ�.
    if ((itemreg.deliverytype[1].checked)||(itemreg.deliverytype[3].checked)||(itemreg.deliverytype[4].checked)){
        if ((itemreg.mwdiv[0].checked)||(itemreg.mwdiv[1].checked)){
            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
            return;
        }
        if (itemreg.deliverOverseas.checked){
            alert('�ٹ����� ����� ��쿡�� �ؿܹ���� �Ͻ� �� �ֽ��ϴ�.');
            return;
        }
    }
    
    //���Ա����� ��ü�̳� ��۱����� ��ü�� �ƴѰ�.
    if (itemreg.mwdiv[2].checked){
        if ((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
            alert('��� ������ Ȯ�����ּ���. ���� ���а� ��ġ���� �ʽ��ϴ�..');
            return;
        }
    }

    //��ü��۸� �ֹ����� ����.
    if ((!itemreg.mwdiv[2].checked)&&(itemreg.itemdiv[1].checked)){
        alert('(�ֹ����� ��ǰ�� ��ü����ΰ�츸 �����մϴ�.');
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
    
	//��ü����ΰ�� �Ǹž��� ���þ���

	if ((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
		if (itemreg.sellyn[0].checked){
			alert('��ü����ΰ�� �Ǹſ��δ� N�� �����ϼ���.');
			itemreg.sellyn[1].focus();
			return;
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
        if (CheckImage(itemreg.imgmask, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
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

    if (itemreg.imgmain.value != "") {
        if (CheckImage(itemreg.imgmain, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgmain2.value != "") {
        if (CheckImage(itemreg.imgmain2, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgmain3.value != "") {
        if (CheckImage(itemreg.imgmain3, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (processstatus==true){
		if(confirm("["+itemreg.itemname.value+"]�� ���� �Ͻðڽ��ϱ�?\n��ü��ۻ�ǰ�� ��� ����Ʈ�� �ٷ� ����Ǹ�,\n�ٹ����ٹ�ۻ�ǰ�� �԰� �Ϸ� �� ��ǰ�� ���µ˴ϴ�.") == true){
			itemreg.action = "<%= ItemUploadUrl %>/linkweb/items/doWaitItemToReg_byadmin.asp";
			itemreg.mode.value = "realupload";
			itemreg.itemoptioncode2.value=optionv;
			itemreg.itemoptioncode3.value=optiont;
			itemreg.target = "FrameCKP";
			itemreg.submit();
		}
	}else{
		if(confirm("��ǰ�� �ӽ� ���� �Ͻðڽ��ϱ�?") == true){
			itemreg.action = "<%= ItemUploadUrl %>/linkweb/items/doWaitItemToReg_byadmin.asp";
			itemreg.mode.value = "waititemmodi";
			itemreg.itemoptioncode2.value=optionv;
			itemreg.itemoptioncode3.value=optiont;
			itemreg.target = "FrameCKP";
			itemreg.submit();
		}
	}
}

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
	}
	else if(frm.mwdiv[2].checked){
	    // ��۱��� ����(��ü���)
	    if ((frm.defaultFreeBeasongLimit.value*1>0)&&(frm.defaultDeliverPay.value*1>0)){
	        frm.deliverytype[3].checked=true;	// �⺻ üũ
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
		frm.limitno.readOnly=true;
		frm.limitno.className='text_ro';

		frm.limitsold.readOnly=true;
		frm.limitsold.className='text_ro';
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
		frm.limitno.className='text';

		frm.limitsold.readOnly=false;
		frm.limitsold.className='text';
	}
}

function TnSilentCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// ������
		frm.limitno.readOnly=true;
		frm.limitno.className='text_ro';

		frm.limitsold.readOnly=true;
		frm.limitsold.className='text_ro';
	}
	else {
		// ����
		if ((frm.optioncnt.value*1) > 0) {
		    // �ɼǻ����
		    // alert("���������� �ɼ�â���� ���������մϴ�.");
		    return;
        }

		frm.limitno.readOnly=false;
		frm.limitno.className='text';

		frm.limitsold.readOnly=false;
		frm.limitsold.className='text';
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
	else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("����Ư�� ������ �����̳� Ư���� ���\n��۱�����  ��ü������� ���� �Ͻ� �� �����ϴ�!!!\n����Ư�������� Ȯ�����ּ���!!");
			frm.mwdiv[2].checked=true;
		}
	}
	
	if ((frm.defaultFreeBeasongLimit.value*1>0)&&(frm.defaultDeliverPay.value*1>0)&&(!frm.deliverytype[3].checked)){
	    alert('��ü ���� ��� ��ü�Դϴ�. ��۱����� Ȯ���ϼ���.')
	    frm.deliverytype[3].focus();
	}
	
	if (((frm.defaultFreeBeasongLimit.value*1<1)||(frm.defaultDeliverPay.value*1<1))&&(frm.deliverytype[3].checked)){
	    alert('��ü ���� ��� ��ü�� �ƴմϴ�. ��۱����� Ȯ���ϼ���.')
	    frm.deliverytype[3].focus();
	}
	
	if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(frm.deliverytype[4].checked)){
	    alert('��ü ���� ��� ��ü�� �ƴմϴ�. ��۱����� Ȯ���ϼ���.')
	    frm.deliverytype[4].focus();
	}
}
function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // �ɼǻ��
	    if (confirm("�ɼ��� ��������/�����ο� ���� �Ǵ� ����� 2�� �̻��ΰ�� ��� �����մϴ�. �����Ͻðڽ��ϱ�?") == true) {
            frm.btnoptadd.disabled = false;
            frm.btnetcoptadd.disabled = false;
            frm.btnoptdel.disabled = false;

            optlist.style.display="";
        } else {
            frm.useoptionyn[1].checked = true;
            TnCheckOptionYN(frm);
        }
	} else {
	    // �ɼǾ���
	    // while (frm.realopt.length > 0) {
	    //     frm.realopt.options[0] = null;
        // }
        frm.btnoptadd.disabled = true;
        frm.btnetcoptadd.disabled = true;
        frm.btnoptdel.disabled = true;

		optlist.style.display="none";

        frm.itemoptioncode2.value = "";
        frm.itemoptioncode3.value = "";
    }
}
function TnCheckSailYN(frm){
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
    CalcuAuto(frm);
}


//�����ڵ� ����
function selColorChip(cd) {
	var i;
	itemreg.DFcolorCD.value= cd;
	for(i=0;i<=30;i++) {
		document.all("cline"+i).bgColor='#DDDDDD';
	}
	if(!cd) document.all("cline0").bgColor='#DD3300';
	else document.all("cline"+cd).bgColor='#DD3300';
}


// ============================================================================
// �̸�����
function ViewItemDetail(itemno){
	window.open('/designer/itemmaster/viewitem.asp?itemid='+itemno ,'window1','width=790,height=600,scrollbars=yes,status=no');
}


function jsUniWaitState(currstate, count2){
 $("#dv2").hide();
 $("#dv0").hide();    
	if(count2>=2){ 
		 document.all.chkV0[4].checked = true;
		 document.all.sM0.value = "3ȸ �̻� ������, �ݷ�ó��(���� �Ұ�)�˴ϴ�.";
	}
	var maskHeight = $(document).height();
	var maskWidth = $(document).width();

	$('#mask').css({'width':maskWidth,'height':maskHeight}); 
	$('#boxes').show();
	$('#mask').show();
		var winH = $(document).height()-500;
		var winW = $(document).width(); 
		$("#dialog").css('top', winH-$("#dialog").height());
		$("#dialog").css('left', winW/2-$("#dialog").width()/2);
		$("#dialog").show(); 
		$("#dv"+currstate).show();  
}		 

//����ó��
 function jsConfirm(currstate){ 
 	var chkCount = 0;
 	var iMsgcd = "";
 	var sMsg = "";  
 	for(i=0;i<eval("document.all.chkV"+currstate).length;i++){
 		if(eval("document.all.chkV"+currstate)[i].checked){
 		chkCount = chkCount + 1;
 		if (iMsgcd==""){
 			iMsgcd = eval("document.all.chkV"+currstate)[i].value;
 			if (eval("document.all.chkV"+currstate)[i].value==999){
 					sMsg = eval("document.all.sM"+currstate).value;
 			}else{
 				sMsg = $("#sp"+currstate+i).text(); 
 			}
 		}else{
 		iMsgcd = iMsgcd +"^"+ eval("document.all.chkV"+currstate)[i].value;
 			if (eval("document.all.chkV"+currstate)[i].value==999){
 					sMsg = sMsg +"^"+eval("document.all.sM"+currstate).value;
 			}else{
 				sMsg = sMsg +"^"+ $("#sp"+currstate+i).text();
 			}
 		} 
 	} 
 	}
 	if(chkCount == 0){
 		alert("���� �ź� ������ �Ѱ� �̻� �������ּ���");
 		return;
 	} 
 	document.borufrm.sMsgcd.value= iMsgcd;
 	document.borufrm.sMsg.value = sMsg;
 	document.borufrm.hidM.value = "U";
 	document.borufrm.sCS.value = currstate;
  document.borufrm.submit(); 
}

  function jsCancel(){
  	document.borufrm.sMsgcd.value= "";
 		document.borufrm.sMsg.value = ""; 
  	 $( "#dialog" ).hide();
  	 $('#mask').hide();
  	 $('#boxes').hide();
  }

function GetRejectMsg(falg){
    var tmp = window.showModalDialog('pop_rejectMsg.asp?falg=' + falg,null, "dialogwidth:250px;dialogheight:230px;center:yes;scroll:no;resizable:no;status:no;help:no;");
    return tmp;
}

function ClearVal(comp){
    comp.value = "";
}

function popWaitItemOptionEdit(iitemid){
    var popwin = window.open('/common/pop_upchewaititemoptionedit.asp?itemid=' + iitemid,'popWaitItemOptionEdit','width=790,height=600,scrollbars=yes,status=no');
    popwin.focus();
}

function EnDisableFlowerShop(){
    var frm = document.itemreg;
    if ((frm.cd1.value=="110")&&(frm.cd2.value=="060")){
        frm.deliverarea[1].disabled = false;
        frm.deliverarea[2].disabled = false;
        
        deliverfixday.disabled = false;
    }else{
        frm.deliverarea[1].disabled = true;
        frm.deliverarea[2].disabled = true;
        
        deliverfixday.disabled = true;
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
			url: "act_waitItemInfoDivForm.asp",
			data: "itemid=<%=request("itemid")%>&ifdv="+v,
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
		<font color="red"><strong>���δ���ǰ����</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>���δ������ ��ǰ�� ���ĵ���մϴ�.</b>
			<br><br>- �߸��� �κ��� �ӽ����� ����� �̿��Ͽ� �����ϽǼ� �ֽ��ϴ�.
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
<form name="itemreg" method="post" action="" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="itemid" value="<%= request("itemid") %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="defaultmaeipdiv" value="<%= npt_defaultmargine %>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%= npt_defaultFreeBeasongLimit %>">
<input type="hidden" name="defaultDeliverPay" value="<%= npt_defaultDeliverPay %>">
<input type="hidden" name="defaultDeliveryType" value="<%= npt_defaultDeliveryType %>">
<input type="hidden" name="DFcolorCD" value="<%=oitemdetail.FDFcolorCd%>">

<input type="hidden" name="pojangok" value="N">
<input type="hidden" name="itemoptioncode2" value="">
<input type="hidden" name="itemoptioncode3" value="">
<input type="hidden" name="isusing" value="Y">
<input type="hidden" name="adminid" value="<%=session("ssBctId") %>">
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


<!-- 1.�Ϲ����� --> 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>1.�Ϲ�����</strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�귣��ID :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="designerid"  value="<%= oitemdetail.FMakerid %>" class="text_ro" readonly size="30" id="[on,off,off,off][�귣��ID]">&nbsp;
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�ڵ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<%= request("itemid") %>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="�̸�����" onclick="ViewItemDetail('<%= request("itemid") %>');">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemname" maxlength="64" size="50" class="text" value="<%= oitemdetail.Fitemname %>" id="[on,off,off,off][��ǰ��]">&nbsp;
	</td>
</tr>
</table>

<!-- 2.���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>2.����</strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="���/���� ���� ���� ī�װ�" style="cursor:help;">���� ī�װ� :</td>
		<input type="hidden" name="cd1" value="<%= oitemdetail.Flarge %>">
		<input type="hidden" name="cd2" value="<%= oitemdetail.Fmid %>">
		<input type="hidden" name="cd3" value="<%= oitemdetail.Fsmall %>">
	<td bgcolor="#FFFFFF" colspan="2">
		<input type="text" name="cd1_name" value="<%= oitemreg.largename %>" id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro">
		<input type="text" name="cd2_name" value="<%= oitemreg.midname %>"  id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro">
		<input type="text" name="cd3_name" value="<%= oitemreg.smallname %>" id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro">
		
		<input type="button" value="ī�װ� ����" class="button" onclick="editCategory(itemreg.cd1.value,itemreg.cd2.value,itemreg.cd3.value);">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="����Ʈ�� ������ ī�װ�" style="cursor:help;">���� ī�װ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td id="lyrDispList"><%=getDispCategoryWait(request("itemid"))%></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF" >
		<label><input type="radio" name="itemdiv" value="01" <%=chkIIF(oitemdetail.Fitemdiv ="01","checked","")%> onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�Ϲݻ�ǰ</label>
		<br>
		<label><input type="radio" name="itemdiv" value="<%= oitemdetail.Fitemdiv %>" <%=chkIIF(oitemdetail.Fitemdiv="06" or oitemdetail.Fitemdiv="16","checked","")%> onClick="document.getElementById('lyRequre').style.display='block';checkItemDiv(this);">�ֹ� ���ۻ�ǰ</label>
		<input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitemdetail.Fitemdiv="06","checked","")%> <%=chkIIF(oitemdetail.Fitemdiv="06" or oitemdetail.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">�ֹ����� ���� �ʿ�<font color=red>(�ֹ��� �̴ϼȵ� ���۹����� �ʿ��Ѱ�� üũ)</font>
		<br>
	</td>
	<td bgcolor="#FFFFFF" >
	    <div id="lyRequre" style="<%=chkIIF(oitemdetail.Fitemdiv ="06" or oitemdetail.Fitemdiv ="16","","display:none;")%>padding-left:22px;">
			�������ۼҿ��� <input type="text" name="requireMakeDay" value="<%=oitemdetail.FrequireMakeDay%>" size="2" class="text" id="[off,on,off,off][�������ۼҿ���]">��
			<font color="red">(��ǰ�߼��� ��ǰ���� �Ⱓ)</font>
		</div>
	</td>
</tr>
</table>

<!-- 3.�������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>3.��������</strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td width="15%" bgcolor="#DDDDFF">����, �鼼 ���� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="radio" name="vatinclude" value="Y" onclick="TnGoClear(this.form);" <%=chkIIF(oitemdetail.Fvatinclude="Y","checked","")%>>����
		<input type="radio" name="vatinclude" value="N" onclick="TnGoClear(this.form);" <%=chkIIF(oitemdetail.Fvatinclude="N","checked","")%>>�鼼
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">�⺻ ���� ���� :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="margin" maxlength="32" size="5" id="[off,off,off,off][����]" value="<%= oitemdetail.FMargin %>" class="text">%
		<% if (CStr(npt_defaultmargine)<>CStr(oitemdetail.FMargin)) then %>
		<font Color="red">(��ü �⺻ ���� : <%= npt_defaultmargine %>)</font>
		<% end if %>
	</td>
</tr>
<tr align="left">
<input type="hidden" name="sellvat" value="<%= oitemdetail.Fsellvat %>">
<input type="hidden" name="buyvat" value="<%= oitemdetail.Fbuyvat %>">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ǸŰ�(�Һ��ڰ�) :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" name="sellcash" size="12" id="[on,on,off,off][�Һ��ڰ�]" onKeyUp="CalcuAuto(itemreg);" maxlength="8" class="text" value="<%= oitemdetail.Fsellcash %>">��
		<input type="button" value="���ް� �ڵ����" onclick="CalcuAuto(itemreg);" class="button" style="width:100px;">
	</td>
	<td width="15%" bgcolor="#DDDDFF">���ް� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" name="buycash" maxlength="16" size="12" id="[on,on,off,off][���ް�]" class="text" value="<%= oitemdetail.Fbuycash %>">��
		(<b>�ΰ��� ���԰�</b>)
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF"></td>
	<td bgcolor="#FFFFFF" colspan="3">
		- ���ް��� <b>�ΰ��� ���԰�</b>�Դϴ�.
		<br>- �ǸŰ�(�Һ��ڰ�)�� �Է��ϸ� ������ �������� ���ް��� �ڵ����˴ϴ�.
		<br>- ������ ������ ������� ���ް��� �Է��ϽǼ� �ֽ��ϴ�.
	</td>
</tr>
<input type="hidden" name="mileage" id="[on,off,off,off][���ϸ���]" value="<%= oitemdetail.Fmileage %>">
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
	<td id="lyrItemAttribAdd" bgcolor="#FFFFFF" colspan="3">������ ����ī�װ��� �����ϴ�.</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ü��ǰ�ڵ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="upchemanagecode" value="<%= oitemdetail.Fupchemanagecode %>" size="20" maxlength="32" class="text" id="[off,off,off,off][��ü��ǰ�ڵ�]">
	    (��ü���� �����ϴ� �ڵ� �ִ� 32�� - ����/���ڸ� ����)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�Ǹſ��� :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<label><input type="radio" name="sellyn" value="Y" <% if oitemdetail.Fsellyn = "Y" then response.write "checked" %>>�Ǹ���</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="N" <% if oitemdetail.Fsellyn = "N" then response.write "checked" %>>�Ǹž���</label>
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
	  <input type="text" name="makername" maxlength="32" size="25" class="text" id="[on,off,off,off][������]" value="<%= oitemdetail.Fmakername %>">&nbsp;(������ü��)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="text" name="sourcearea" maxlength="64" size="25" class="text" id="[on,off,off,off][������]" value="<%= oitemdetail.Fsourcearea %>">&nbsp;(ex:�ѱ�,�߱�,�߱�OEM,�Ϻ�...)
	  <br>( ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�˻�Ű���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="text" name="keywords" maxlength="128" size="60" class="text" id="[on,off,off,off][�˻�Ű����]" value="<%= oitemdetail.Fkeywords %>">&nbsp;(�޸��α��� ex: Ŀ��,Ƽ����,����)
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
		<script type="text/javascript">
		document.itemreg.infoDiv.value="<%=oitemdetail.FinfoDiv%>";
		</script>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitemdetail.FinfoDiv) or oitemdetail.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ�񳻿� :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitemdetail.FinfoDiv) or oitemdetail.FinfoDiv="") then
			Server.Execute("act_waitItemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">��ǰ���������� ������ ���� �Ǿ��ִ��� ��Ȯ�� �Է¹ٶ��ϴ�. ����Ȯ�ϰų� �߸��� ���� �Է½�, �׿� ���� å���� ���� ���� �ֽ��ϴ�.</font></td>
</tr>
<tr align="left" id="lyItemSrc" style="display:<%=chkIIF(oitemdetail.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF">
	  <input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitemdetail.Fitemsource %>">&nbsp;(ex:�ö�ƽ,����,��,...)
	</td>
</tr>
<tr align="left" id="lyItemSize" style="display:<%=chkIIF(oitemdetail.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ������ :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitemdetail.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
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
		<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitemdetail.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)">���</label>
		<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","checked","")%> onclick="chgSafetyYn(document.itemreg)">���ƴ�</label><br />
		<select name="safetyDiv" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","disabled","")%> class="select">
		<option value="">::������������::</option>
		<option value="10" <%=chkIIF(oitemdetail.FsafetyDiv="10","selected","")%>>������������(KC��ũ)</option>
		<option value="20" <%=chkIIF(oitemdetail.FsafetyDiv="20","selected","")%>>�����ǰ ��������</option>
		<option value="30" <%=chkIIF(oitemdetail.FsafetyDiv="30","selected","")%>>KPS �������� ǥ��</option>
		<option value="40" <%=chkIIF(oitemdetail.FsafetyDiv="40","selected","")%>>KPS �������� Ȯ�� ǥ��</option>
		<option value="50" <%=chkIIF(oitemdetail.FsafetyDiv="50","selected","")%>>KPS ��� ��ȣ���� ǥ��</option>
		</select>
		������ȣ <input type="text" name="safetyNum" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","disabled","")%> size="18" maxlength="18" class="text" value="<%=oitemdetail.FsafetyNum%>" />
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
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����Ư������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="mwdiv" value="M" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(oitemdetail.Fmwdiv="M","checked","")%>>����
	  <input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(oitemdetail.Fmwdiv="W","checked","")%>>Ư��
	  <input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(oitemdetail.Fmwdiv="U","checked","")%>>��ü���
	  &nbsp;&nbsp; - ����Ư�����п� ���� ��۱����� �޶����ϴ�. ��۱����� Ȯ�����ּ���.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��۱��� :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverytype" value="1" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="1","checked","")%>>�ٹ����ٹ��&nbsp;
	  <input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="2","checked","")%>>��ü(����)���&nbsp;
	  <input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="4","checked","")%>>�ٹ����ٹ�����&nbsp;
	  <input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="9","checked","")%>>��ü���ǹ��(���� ��ۺ�ΰ�)&nbsp;
	  <input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="7","checked","")%>>��ü���ҹ��
	 </td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��۹�� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverfixday" value="" onclick="TnCheckFixday(this.form)" <%=chkIIF(Trim(oitemdetail.Fdeliverfixday)="" or IsNull(oitemdetail.Fdeliverfixday),"checked","")%>>�ù�(�Ϲ�)&nbsp;
	  <input type="radio" name="deliverfixday" value="X" onclick="TnCheckFixday(this.form)" <%=chkIIF(oitemdetail.Fdeliverfixday="X","checked","")%>>ȭ��&nbsp;
	  <input type="radio" name="deliverfixday" value="C" onclick="TnCheckFixday(this.form)" <%=chkIIF(oitemdetail.Fdeliverfixday="C","checked","")%>>�ö��������
		<span id="lyrFreightRng" style="display:<%=chkIIF(oitemdetail.Fdeliverfixday="X","","none")%>;">
			<br />&nbsp;
			��ǰ/��ȯ �� ȭ����� ���(��) :
			�ּ� <input type="text" name="freight_min" class="text" size="6" value="<%=oitemdetail.Ffreight_min%>" style="text-align:right;">�� ~
			�ִ� <input type="text" name="freight_max" class="text" size="6" value="<%=oitemdetail.Ffreight_max%>" style="text-align:right;">��
		</span>
	  <br>&nbsp;<font color="red">(�ö�� ��ǰ�� ��츸 �����ǹ��, ������, �ö�������� �ɼ��� ��밡���մϴ�.)</font>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverarea" value="" <%=chkIIF(Trim(oitemdetail.Fdeliverarea)="" or IsNull(oitemdetail.Fdeliverarea),"checked","")%>>�������&nbsp;
	  <input type="radio" name="deliverarea" value="C" <%=chkIIF(oitemdetail.Fdeliverarea="C","checked","")%>>�����ǹ��&nbsp;
	  <input type="radio" name="deliverarea" value="S" <%=chkIIF(oitemdetail.Fdeliverarea="S","checked","")%>>������&nbsp;
	  <input type="checkbox" name="deliverOverseas" value="Y" <% if oitemdetail.Fmwdiv<>"U" then response.write "checked" %> title="�ؿܹ���� ��ǰ���԰� �Է��� �ž� �Ϸ�˴ϴ�.">�ؿܹ��
	</td>
</tr>
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
	<td height="30" width="15%" bgcolor="#DDDDFF" rowspan="2">�ɼǱ��� :</td>
	<input type="hidden" name="optioncnt" value="<%= oitemdetail.Foptioncnt %>">
	<td width="85%" bgcolor="#FFFFFF">
	  <% if oitemdetail.Foptioncnt < 1 then %>
	  �ɼǻ�����
	  <% else %>
	  �ɼǻ����(<%= oitemdetail.Foptioncnt %>��)
	  <% end if %>
	  &nbsp;&nbsp;<input type="button" class="button" value="�ɼǼ���" onClick="popWaitItemOptionEdit('<%= oitemdetail.FWaitItemID %>');">
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF">
      - �ɼ������� �ɼ�â���� ���������մϴ�.<br>
      - �ɼ��� ���ĵ���� ������ �Ұ����մϴ�. ��Ȯ�� �Է��ϼ���.
	</td>
</tr>
<tr id="lyDFColor" height="30">
	<td colspan="2" bgcolor="#FFFFFF" style="padding:0px;">
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="0">
		<tr>
			<td width="15%" bgcolor="#DDDDFF">�⺻ ������ :</td>
			<td width="85%" bgcolor="#FFFFFF" style="border-left:1px solid <%= adminColor("tablebg") %>;"><%=FnSelectColorBar(oitemdetail.FDFColorCD,25)%></td>
		</tr>
		<tr>
			<td width="15%" rowspan="2" bgcolor="#DDDDFF" style="border-top:1px solid <%= adminColor("tablebg") %>;">���� ��ǰ�̹��� :</td>
			<td width="85%" bgcolor="#FFFFFF" style="border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;">
			  <% if (oitemdetail.FDFcolorImg <> "") then %>
				<div id="divimgDFColor" style="display:block;">
				<img src="<%=partnerUrl%>/waitimage/color/<%=imgsubdir%>/<%=oitemdetail.FDFcolorImg %>" width="200">
				</div>
			  <% end if %>
				<input type="file" size="40" name="imgDFColor" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text">
				<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgDFColor, 40, 1000, 1000)"> (����,1000X1000,<b><font color="red">jpg</font></b>)
				<input type="hidden" name="DFColor">
			</td>
		</tr>
		<tr>
			<td width="85%" bgcolor="#FFFFFF" style="border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;">
		      - ���� �̹����� ������ ���ĵ���� �����ʽ��ϴ�.(Err:013) ���ĵ�Ͻÿ� �ݵ�� ������ּ���.
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
	  <input type="radio" name="limityn" value="N" onClick="TnCheckLimitYN(itemreg)" <%=chkIIF(oitemdetail.Flimityn="N","checked","")%>>�������Ǹ�&nbsp;&nbsp;
	  <input type="radio" name="limityn" value="Y" onClick="TnCheckLimitYN(itemreg)" <%=chkIIF(oitemdetail.Flimityn="Y","checked","")%>>�����Ǹ�
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">�������� :</td>
	<td width="35%" bgcolor="#FFFFFF" >
	  <input type="text" name="limitno" maxlength="32" size="8" readonly class="text_ro" id="[off,on,off,off][��������]" value="<%= oitemdetail.Flimitno %>">(��)
      <input type="hidden" name="limitsold" value="0">
      <input type="hidden" name="limitstock" value="<%= oitemdetail.Flimitno %>">
	</td>
</tr>
<tr>
	<td colspan="3" bgcolor="#FFFFFF"><font color="red">** ���������� �ɼ��� ���� ���, �ɼ�â���� ������ �����մϴ�.(���� ������ ����Ȯ�Ҽ� �ֽ��ϴ�.)</font></td>
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
	  <input type="radio" name="usinghtml" value="N" <%=chkIIF(oitemdetail.Fusinghtml="N","checked","")%>>�Ϲ�TEXT
	  <input type="radio" name="usinghtml" value="H" <%=chkIIF(oitemdetail.Fusinghtml="H","checked","")%>>TEXT+HTML
	  <input type="radio" name="usinghtml" value="Y" <%=chkIIF(oitemdetail.Fusinghtml="Y","checked","")%>>HTML���
	  <br>
	  <textarea name="itemcontent" rows="18" class="textarea" style="width:100%" id="[on,off,off,off][��ǰ����]"><%= oitemdetail.Fitemcontent %></textarea>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ��� ���ǻ��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <textarea name="ordercomment" rows="5" cols="90" class="textarea" id="[off,off,off,off][���ǻ���]"><%= oitemdetail.Fordercomment %></textarea><br>
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
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �⺻�̹��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fimgbasic <> "") then %>
		<div id="divimgbasic" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/basic/<%=imgsubdir%>/<%= ooimage.Fimgbasic %>" width="300" height="300">
		</div>
	  <% end if %>
	  <input type="file" name="imgbasic" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgbasic,40, 1000, 1000)"> (<font color=red>�ʼ�</font>,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="basic">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ<br />����(����)�̹��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fimgmask <> "") then %>
		<div id="divimgmask" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mask/<%=imgsubdir%>/<%= ooimage.Fimgmask %>" width="300" height="300">
		</div>
	  <% end if %>
	  <input type="file" name="imgmask" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmask,40, 1000, 1000)"> (����,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="mask">
	</td>
</tr>
<tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (itemaddimage(0) <> "") then %>
		<div id="divimgadd1" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add1/<%=imgsubdir%>/<%=itemaddimage(0) %>" width="300" height="300">
		</div>
	  <% end if %>
		<input type="file" name="imgadd1" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd1,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add1">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���2 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (itemaddimage(1) <> "") then %>
		<div id="divimgadd2" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add2/<%=imgsubdir%>/<%=itemaddimage(1) %>" width="300" height="300">
		</div>
	  <% end if %>
		<input type="file" name="imgadd2" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd2,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���3 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (itemaddimage(2) <> "") then %>
		<div id="divimgadd3" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add3/<%=imgsubdir%>/<%=itemaddimage(2) %>" width="300" height="300">
		</div>
	  <% end if %>
		<input type="file" name="imgadd3" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd3,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add3">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���4 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (itemaddimage(3) <> "") then %>
		<div id="divimgadd4" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add4/<%=imgsubdir%>/<%=itemaddimage(3) %>" width="300" height="300">
		</div>
	  <% end if %>
		<input type="file" name="imgadd4" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd4,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add4">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���5 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (itemaddimage(4) <> "") then %>
		<div id="divimgadd5" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add5/<%=imgsubdir%>/<%=itemaddimage(4) %>" width="300" height="300">
		</div>
	  <% end if %>
		<input type="file" name="imgadd5" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd5,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add5">
	</td>
</tr>
<tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fimgmain <> "") then %>
		<div id="divimgmain" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/main/<%=imgsubdir%>/<%=ooimage.Fimgmain %>" width="400">
		</div>
	  <% end if %>
		<input type="file" name="imgmain" onchange="CheckImage(this, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmain, 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #2 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fimgmain2 <> "") then %>
		<div id="divimgmain2" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/main2/<%=imgsubdir%>/<%=ooimage.Fimgmain2 %>" width="400">
		</div>
	  <% end if %>
		<input type="file" name="imgmain2" onchange="CheckImage(this, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmain2, 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #3 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fimgmain3 <> "") then %>
		<div id="divimgmain3" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/main3/<%=imgsubdir%>/<%=ooimage.Fimgmain3 %>" width="400">
		</div>
	  <% end if %>
		<input type="file" name="imgmain3" onchange="CheckImage(this, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmain3, 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main3">
	</td>
</tr>
</table>

 
<!-- 11.������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>11.�������</strong> </td>
    <td align="right">3ȸ �̻� ������, �ݷ� ó��(���ϺҰ�)�ǹǷ� ���� ��Ź �帳�ϴ�.</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>  
<table width="100%" border="0" align="center" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	 <tr>
	 	<td width="15%" align="center" bgcolor="#DDDDFF">��������</td>
	 	<td bgcolor="#FFFFFF">
	 		<%IF isARray(arrList) THEN%>
	 <% dim count2, strMsg, sMsgCd2, sMsgcd0
	 count2 = 0 
	 sMsgCd2 = ""
	 sMsgcd0 = ""
	 For intLoop = 0 To UBound(arrList,2)
	 strMsg = ""
	 		IF arrList(2,intLoop) = 2 THEN
	 			count2 = count2 + 1
	 			strMsg = count2&"��"
	 			sMsgCd2 = sMsgCd2 + "^" + arrList(6,intLoop)
	 		ELSEIF arrList(2,intLoop) = 0 THEN
	 				sMsgCd0 = sMsgCd0 + "^" + arrList(6,intLoop)	
	 		END IF	
	 %> 
	 <div style="padding:3"><font color="<%=GetCurrStateColor(arrList(2,intLoop))%>"><%=strMsg%><%=fnGetCurrStateShortName(arrList(2,intLoop))%></font>: <%=arrList(4,intLoop)%> &nbsp;<%IF arrList(3,intLoop) <> "" THEN%>[<%=replace(arrList(3,intLoop),"^","/")%>]<%END IF%></div>
	 <%Next%>
	  <%ELSEIF isArray(arrold) THEN
	  		IF arrold(4,0) = 5 THEN
	  	%>
	  	 <div style="padding:3">����:<%=arrold(0,0)%> &nbsp;[<%=arrold(1,0)%>] </div>
	  	 <div style="padding:3"><font color="<%=GetCurrStateColor(arrold(4,0))%>"><%=fnGetCurrStateShortName(arrold(4,0))%></font>: <%=arrold(2,0)%> &nbsp;[<%=arrold(3,0)%>]</div>
	  	<%ELSE%>
			 <font color="<%=GetCurrStateColor(arrold(4,0))%>"><%=fnGetCurrStateShortName(arrold(4,0))%></font>: <%=arrold(0,0)%> &nbsp;[<%=arrold(1,0)%>] 
			 <%END IF%>
	 <%END IF%>
 	</td>
</tr>
</table> 


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
    	<input type="button" value="�ӽ�����" class="button" onclick="SubmitSave(false);" <% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"5" then response.write "disabled" %>> 
       <%IF count2<2 THEN%>&nbsp;&nbsp;&nbsp;  
      <input type="button" value="���κ��� (���Ͽ�û)" class="button" onclick="jsUniWaitState(2,'<%=count2%>');" <% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"5" then response.write "disabled" %>>
      <%END IF%>
      &nbsp;&nbsp;&nbsp;
      <input type="button" value="���ιݷ� (���ϺҰ�)" class="button" onclick="jsUniWaitState(0,'<%=count2%>');" <% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"5" then response.write "disabled" %>>
			&nbsp;&nbsp;&nbsp;
			<input type="button" value="����" class="button" onclick="SubmitSave(true);" <% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"5" then response.write "disabled" %> >
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
</form>
<!-- ǥ �ϴܹ� ��-->
<form name="borufrm" method="post" action="doitemregboru.asp"> 
<input type="hidden" name="hidM" value="U"> 
<input type="hidden" name="itemid" value="<%= request("itemid") %>">  
<input type="hidden" name="sCS" value="">
<input type="hidden" name="sMsgcd" value="">
<input type="hidden" name="sMsg" value="">
<input type="hidden" name="sRU" value="wait_item_modify.asp?itemid=<%=itemid%>&designer=<%=makerid%>">	
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
 <div id="boxes">  
<div id="mask"></div>
<div id="dialog">  
<!-- #include virtual="/admin/itemmaster/item_confirm_inc.asp"-->
</div>  
<script>
// ����Ư������ �� ��۱��м���
TnCheckUpcheYN(itemreg);
for (var i = 0; i < itemreg.elements.length; i++) {
    if (itemreg.elements(i).name == "deliverytype") {
        if (itemreg.elements(i).value == "<%= oitemdetail.Fdeilverytype %>") {
            itemreg.elements(i).checked = true;
        }
    }
}

// ����
TnSilentCheckLimitYN(itemreg);
// ����
// TnCheckSailYN(itemreg);

<% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"5" then %>
alert('���� ��� ���°� �ƴմϴ�.');
<% end if %> 
</script>
<%
set oitemdetail = Nothing
set oitemreg = Nothing
set ooimage = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->