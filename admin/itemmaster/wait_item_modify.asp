<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �¶��� ���δ���ǰ
' History : ������ ����
'			2023.08.11 �ѿ�� ����(isbn �߰�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemregcls.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<!-- #include virtual="/lib/classes/items/waititemcls_2014.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 640   'KB

CONST CMAIN_IMG_MAXSIZE = 1230   'KB
CONST CMAIN_IMG_MAXWIDTH = 1000   'Px
CONST CMAIN_IMG_MAXHEIGHT = 3000   'Px

CONST CMOBILE_IMG_MAXSIZE = 500   'KB

dim arrold, deliverfixday, mwdiv, deliverytype, purchaseType, deliverarea
Dim clsWait, itemid ,makerid,arrlist, intLoop, i
makerid	= requestCheckvar(Request("designer"),32)
itemid =  requestCheckvar(Request("itemid"),16)

Dim oitemdetail,oitemreg,optiontotal,ix,ooimage, mainImg(10)

set oitemdetail = new CWaitItemDetail

oitemdetail.FRectDesignerID = request("designer")
oitemdetail.WaitProductDetail request("itemid") '�ӽõ�� ������ �ҷ�����
oitemdetail.WaitProductDetailOption request("itemid") '�ɼ� 2�� �ѹ�,�̸� �ҷ�����


if oitemdetail.FTotalCount>0 then
	purchaseType = oitemdetail.fpurchaseType		' ��������

	' ���������� �ؿ����� �ϰ�� ���� ����
	if purchaseType="9" then
		deliverfixday = "G"	' �ؿ�����
		mwdiv = "U"
		deliverarea = ""

		' ��ü(����)��� �ϰ��
		if oitemdetail.Fdeliverytype="2" then
			deliverytype = oitemdetail.Fdeliverytype
		else
			deliverytype = "9"
		end if
	else
		deliverfixday = oitemdetail.Fdeliverfixday	' �ؿ�����
		mwdiv = oitemdetail.Fmwdiv
		deliverarea = oitemdetail.Fdeliverarea
		deliverytype = oitemdetail.Fdeliverytype
	end if
end if

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
ooimage.WaitProductImageList request("itemid")  '�̹��� ������ �ҷ�����

Dim itemaddimage,itemaddcontent, itemstoryimage

if (IsNull(ooimage.Fimgadd) or (ooimage.Fimgadd="")) then ooimage.Fimgadd = ",,,,"
if (IsNull(ooimage.Fitemaddcontent) or (ooimage.Fitemaddcontent="")) then ooimage.Fitemaddcontent = "||||"
if (IsNull(ooimage.Fimgstory) or (ooimage.Fimgstory="")) then ooimage.Fimgstory = ",,,,"


itemaddimage = split(ooimage.Fimgadd,",")
itemaddcontent = split(ooimage.Fitemaddcontent,"|")
itemstoryimage = split(ooimage.Fimgstory,",")


'==============================================================================
dim imgsubdir

imgsubdir = GetImageSubFolderByItemid(request("itemid"))


'==============================================================================
Dim npartner
Dim npt_defaultmargine, npt_defaultFreeBeasongLimit, npt_defaultDeliverPay, npt_defaultDeliveryType
Dim npt_jungsan_gubun, npt_company_no
set npartner = new CPartnerUser
npartner.FRectDesignerID = oitemdetail.FMakerid

if Not(oitemdetail.FMakerid="" or isNull(oitemdetail.FMakerid)) then
	npartner.GetOnePartnerNUser
	if npartner.FResultCount > 0 THEN
	npt_defaultmargine	 = npartner.FOneItem.Fdefaultmargine
	npt_defaultFreeBeasongLimit	= npartner.FOneItem.FdefaultFreeBeasongLimit
	npt_defaultDeliverPay	= npartner.FOneItem.FdefaultDeliverPay
	npt_defaultDeliveryType	= npartner.FOneItem.FdefaultDeliveryType
	npt_jungsan_gubun = npartner.FOneItem.Fjungsan_gubun '2014.02.14 ������ �߰�
	npt_company_no = npartner.FOneItem.Fcompany_no
	end if
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

function getFileSize(fsz)
	if fsz>1024 then
		getFileSize = formatNumber(fsz/1024,2) & "Mb"
	else
		getFileSize = fsz & "Kb"
	end if
end function
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

function ClearImage2(img,fsize,wd,ht) {
	//img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", <!--%= CBASIC_IMG_MAXSIZE %-->, "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";

	if (img=="mobile")
	{
	    document.getElementById("divmobileimgmain").style.display = "none";
	}
	if (img=="mobile2")
	{
	    document.getElementById("divmobileimgmain2").style.display = "none";
	}
	if (img=="mobile3")
	{
	    document.getElementById("divmobileimgmain3").style.display = "none";
	}
	if (img=="mobile4")
	{
	    document.getElementById("divmobileimgmain4").style.display = "none";
	}
	if (img=="mobile5")
	{
	    document.getElementById("divmobileimgmain5").style.display = "none";
	}
	if (img=="mobile6")
	{
	    document.getElementById("divmobileimgmain6").style.display = "none";
	}
	if (img=="mobile7")
	{
	    document.getElementById("divmobileimgmain7").style.display = "none";
	}
	// 20160601�߰��Ѻκ�
	if (img=="mobile8")
	{
	    document.getElementById("divmobileimgmain8").style.display = "none";
	}
	if (img=="mobile9")
	{
	    document.getElementById("divmobileimgmain9").style.display = "none";
	}
	if (img=="mobile10")
	{
	    document.getElementById("divmobileimgmain10").style.display = "none";
	}
	if (img=="mobile11")
	{
	    document.getElementById("divmobileimgmain11").style.display = "none";
	}
	if (img=="mobile12")
	{
	    document.getElementById("divmobileimgmain12").style.display = "none";
	}
	// 20160601�߰��Ѻκ�
	var e = eval("itemreg."+img);
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
	var ctrstate = "<%=oitemdetail.Fctrstate%>";
//	if (ctrstate != "7"){
//		alert("���̿Ϸ�� �귣��� ������ �Ұ����մϴ�.\n���Ȯ�� �� ó�����ּ���");
// 	  	return;
//	}

	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }

	if (processstatus==true&&!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("���� ī�װ��� �����ϼ���.\n�� ���� �⺻ ī�װ��� �ʼ� ���õǾ�� �մϴ�.");
		return;
	}

    //��ǰ�� ����üũ �߰� 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("��ǰ���� �ִ� 64byte ���Ϸ� �Է����ּ���.(�ѱ�32�� �Ǵ� ����64��)");
		itemreg.itemname.focus();
		return;
	}

	//��ǰ���� ����üũ
 if (!IsDigit(document.itemreg.itemWeight.value)){
		alert('��ǰ���Դ�  ���ڷ� �Է��ϼ���.');
		itemreg.itemWeight.focus();
		return;
	}
	//-------------------------------------------------------------------------------- 2014.02.14 ������ �߰�
	//1.����ڰ� [���̰�����] �� ���, ���Ի�ǰ ��� �Ұ� / ��ü,��Ź ��ǰ�� ��ϰ���
	if((itemreg.jungsangubun.value =="���̰���")&&(itemreg.mwdiv[0].checked)){
		alert("����ڰ� [���̰�����]�� ���, [����]��ǰ�� ��ϺҰ����մϴ�. \n[��Ź],[��ü���]��ǰ�� ��ϰ����մϴ�. ");
		itemreg.mwdiv[0].focus();
		return;
	}

	//2.����ڰ� [�鼼�����] �� ���, �鼼��ǰ���θ� ��ϰ���
	if((itemreg.jungsangubun.value =="�鼼")&&(itemreg.vatinclude[0].checked)){
		alert("����ڰ� [�鼼�����]�� ���, [����]��ǰ�� ��ϺҰ����մϴ�. \n[�鼼]��ǰ�� ��ϰ����մϴ�. ");
		itemreg.vatinclude[1].focus();
		return;
	}

	//3.����ڰ� [�ٹ�����]�� ���, ���Ի�ǰ�� ��� ����
	if((itemreg.companyno.value =="211-87-00620")&&(!itemreg.mwdiv[0].checked)){
		alert("����ڰ� [�ٹ�����]�� ���, [����] ��ǰ�� ��ϰ����մϴ�. ");
		itemreg.mwdiv[0].focus();
		return;
	}
	 //--------------------------------------------------------------------------------
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
//        if (itemreg.deliverOverseas.checked){
//            alert('�ٹ����� ����� ��쿡�� �ؿܹ���� �Ͻ� �� �ֽ��ϴ�.');
//            return;
//        }
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
        alert('�ֹ����� ��ǰ�� ��ü����ΰ�츸 �����մϴ�.');
        itemreg.itemdiv[0].focus();
        return;
    }

		if(document.itemreg.optionaddprice.value >0 && document.itemreg.deliverOverseas.checked){
			alert("�ɼǿ� �߰������� ���� ��� �ؿܹ���� �Ұ����մϴ�. �ؿܹ��üũ�� �������ּ���" );
			document.itemreg.deliverOverseas.focus();
			 return;
		}

 if(document.itemreg.deliverOverseas.checked){
	    if(document.itemreg.itemWeight.value<=0){
	        alert("�ؿܹ�۽� ��ۺ� ������ ���� ��ǰ���Ը� �� �Է����ּ���")
	        document.itemreg.itemWeight.focus();
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

	//������������. ���� �Ⱓ�� �ΰ� Ǯ����. ���Ŀ� �ٽ� ���ƾ���. �����Ⱓ : 2018�� 1��1��??
    if (itemreg.safetyYn[0].checked){
  		if($("#real_safetynum").val() == ""){
  			alert("�������������� �����ϰ� ������ȣ�� �Է��� �߰���ư�� Ŭ�����ּ���.");
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

    if (itemreg.sellcash.value*1 < 200 || itemreg.sellcash.value*1 >= 20000000){
		alert("�Ǹ� ������ 200�� �̻� 20,000,000�� �̸����� ��� �����մϴ�.");
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

    if (itemreg.imgmain.value == "" && (itemreg.imgmain2.value != "" || itemreg.imgmain3.value != "")) {
        alert("�����̹����� #1���� ���ʷ� �־��ּ���.");
        return;
    }

    if (itemreg.imgmain2.value == "" && itemreg.imgmain3.value != "") {
        alert("�����̹����� #2���� ���ʷ� �־��ּ���.");
        return;
    }

    if (itemreg.imgmain.value != "") {
        if (CheckImage(itemreg.imgmain, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain2.value != "") {
        if (CheckImage(itemreg.imgmain2, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain3.value != "") {
        if (CheckImage(itemreg.imgmain3, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain4.value != "") {
        if (CheckImage(itemreg.imgmain4, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain5.value != "") {
        if (CheckImage(itemreg.imgmain5, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain6.value != "") {
        if (CheckImage(itemreg.imgmain6, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain7.value != "") {
        if (CheckImage(itemreg.imgmain7, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }

   if(typeof(itemreg.chkSR)=="object"){
   	if(itemreg.chkSR.checked){
	    if(itemreg.dSR.value==""){
	    	alert("���¿����� �����Ǿ��ֽ��ϴ�. ��¥�� �Է����ּ���");
	    	itemreg.dSR.focus();
	    	return;
	    }

     if((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
    	alert("[�ٹ�����(����)���]��ǰ�� ��� ���¿����� �Ұ����մϴ�.");
    	itemreg.chkSR.focus();
    	return;
    }
 	 }
   }

    if (processstatus==true){
	    if(typeof(itemreg.chkSR)=="object"){
	    	if (itemreg.chkSR.checked) {
	    		strMsg = itemreg.dSR.value+" ���¿���� ��ǰ�Դϴ�.";
	    	}else{
	    		strMsg = "��ü��ۻ�ǰ�� ��� ����Ʈ�� �ٷ� ����Ǹ�,\n�ٹ����ٹ�ۻ�ǰ�� �԰� �Ϸ� �� ��ǰ�� ���µ˴ϴ�.";
	    	}
	    }else{
	    		strMsg = "��ü��ۻ�ǰ�� ��� ����Ʈ�� �ٷ� ����Ǹ�,\n�ٹ����ٹ�ۻ�ǰ�� �԰� �Ϸ� �� ��ǰ�� ���µ˴ϴ�.";
	    }
		if(confirm("["+itemreg.itemname.value+"]�� ���� �Ͻðڽ��ϱ�?\n"+strMsg) == true){
			<% ''�������� api�� ��ȸ �� ���� ������ db���� �� ����idx�� �޾� ���� %>
			if(itemreg.safetyYn[0].checked) {
				$("#real_safetyidx").val(jsCallAPIsafety($("#real_safetynum").val(),"u"));
			}

			itemreg.action = "<%= ItemUploadUrl %>/linkweb/items/doWaitItemToReg_byadmin.asp";
			itemreg.mode.value = "realupload";
			itemreg.itemoptioncode2.value=optionv;
			itemreg.itemoptioncode3.value=optiont;
			itemreg.target = "FrameCKP";
			itemreg.submit();
		}
	}else{
		if(confirm("��ǰ�� �ӽ� ���� �Ͻðڽ��ϱ�?") == true){
			<% ''�������� api�� ��ȸ �� ���� ������ db���� �� ����idx�� �޾� ���� %>
			if(itemreg.safetyYn[0].checked) {
				$("#real_safetyidx").val(jsCallAPIsafety($("#real_safetynum").val(),"u"));
			}

			itemreg.action = "<%= ItemUploadUrl %>/linkweb/items/doWaitItemToReg_byadmin.asp";
			itemreg.mode.value = "waititemmodi";
			itemreg.itemoptioncode2.value=optionv;
			itemreg.itemoptioncode3.value=optiont;
			itemreg.target = "FrameCKP";
			itemreg.submit();
		}
	}
}

// ������Ź����
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

        <%
        ' �ؿ����� �ϰ��
        if deliverfixday="G" then
        %>
        	frm.deliverytype[4].disabled=true;  //��ü���ҹ��(7)
        <% else %>
			frm.deliverytype[4].disabled=false;  //��ü���ҹ��(7)
		<% end if %>

       // frm.deliverOverseas.checked=false;	// �ؿܹ��üũ����
	}

	if (frm.deliverytype[1].checked==true || frm.deliverytype[3].checked==true){
		frm.deliverfixday[3].disabled=false;	// �ؿ�����
	}
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

// ��۱���
function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if (frm.mwdiv[2].checked){
			alert("������Ź ������ ��ü�� ���\n��۱����� �ٹ����� ������� ���� �Ͻ� �� �����ϴ�!!\n������Ź������ Ȯ�����ּ���!!");
			frm.mwdiv[0].checked=true;
		}
	}
	else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("������Ź ������ �����̳� ��Ź�� ���\n��۱�����  ��ü������� ���� �Ͻ� �� �����ϴ�!!!\n������Ź������ Ȯ�����ּ���!!");
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
	if(typeof(itemreg.chkSR) == "object"){
		if(itemreg.chkSR.checked){
			alert("��ǰ�� �������� ������, ����� ��¥�� ��ǰ�� ���µ� �� �����ϴ�.\n�ݷ� �Ǵ� ������ �Ͻ÷��� ��ǰ���¿��� ������ ������ּ���.");
			itemreg.chkSR.focus();
			return;
		}
	}

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
		var winH = $(document).height()-800;
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
 	var iNo = "";
 	for(i=0;i<eval("document.all.chkV"+currstate).length;i++){
 		if(eval("document.all.chkV"+currstate)[i].checked){
 		chkCount = chkCount + 1;
 		iNo = eval("document.all.chkV"+currstate)[i].value;
 		if (iMsgcd==""){
 			iMsgcd = eval("document.all.chkV"+currstate)[i].value;
 			if (eval("document.all.chkV"+currstate)[i].value==999){
 					sMsg = eval("document.all.sM"+currstate).value;
 			}else{
 				sMsg = $("#sp"+currstate+iNo).text();
 			}
 		}else{
 		iMsgcd = iMsgcd +"^"+ eval("document.all.chkV"+currstate)[i].value;
 			if (eval("document.all.chkV"+currstate)[i].value==999){
 					sMsg = sMsg +"^"+eval("document.all.sM"+currstate).value;
 			}else{
 				sMsg = sMsg +"^"+ $("#sp"+currstate+iNo).text();
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

	//�޷�
	function jsPopCal(sName){
	 if(!document.all.chkSR.checked){
	 	 document.all.chkSR.checked= true;
	 	}
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();

	}

	//���¿���
	function jsChkSellReserve(){
		if(!document.all.chkSR.checked){
			document.all.dSR.value = "";
		}
	}


//��ǰ���� ���� ������ ���� ǥ��
function jsSetArea(iValue){
	var i;
	for(i=0;i<=4;i++) {
 		eval("document.all.dvArea"+i).style.display = "none";
	}
 	eval("document.all.dvArea"+iValue).style.display = "";
}

//�������� �߰� ��ư �׼�
function jsSafetyAuth(){
	//var cnum = $("#safetyNum").val();
	var cnum = itemreg.safetyNum.value.ltrim().rtrim();
	var listbody = "";
	var safetyvalue = "";
	var safetynum = "";

	if(typeof itemreg.catecode == "undefined"){
		alert("ī�װ��� ������ �ּ���.");
		return;
	}

	if($("#safetyDiv").val() == ""){
		alert("�������������� ������ �ּ���.");
		return;
	}

	var isExist = $("#real_safetydiv").attr("value").indexOf($("#safetyDiv").val()) > -1;
	if(isExist){
		alert("�̹� ���õ� ������������ �Դϴ�.");
		return;
	}
//	var isExistsafetynum = $("#real_safetynum").attr("value").indexOf(cnum) > -1;
//	if(isExistsafetynum){
//		alert("�̹� ���õ� ����������ȣ �Դϴ�.");
//		return;
//	}

	if($("#safetyDiv").val() == "30" || $("#safetyDiv").val() == "60" || $("#safetyDiv").val() == "90"){
		$("#issafetyauth").val("ok");

		safetyvalue = $("#real_safetydiv").val();
		if(safetyvalue == ""){
			$("#real_safetydiv").val($("#safetyDiv").val());
		}else{
			$("#real_safetydiv").val(safetyvalue + "," + $("#safetyDiv").val())
		}

		safetynum = $("#real_safetynum").val();
		if(safetynum == ""){
			$("#real_safetynum").val("x");
		}else{
			$("#real_safetynum").val(safetynum + "," + "x");
		}


		listbody = $("#safetyDivList").html();
		$("#safetyDivList").html(listbody + "<p id='l"+$("#safetyDiv").val()+"'>- " + $("#safetyDiv option:selected").text() + "(������ȣ ����) <input type='button' value='����' onClick='jsSafetyDivListDel("+$("#safetyDiv").val()+");' class='button'><p>");
	}else{

		var msgg = jsCallAPIsafety(cnum,"x");

		if(msgg == "����" || msgg == "����" || msgg == "�������" || msgg == "û���ǽ�"){
			$("#issafetyauth").val("ok");

			safetyvalue = $("#real_safetydiv").val();
			if(safetyvalue == ""){
				$("#real_safetydiv").val($("#safetyDiv").val());
			}else{
				$("#real_safetydiv").val(safetyvalue + "," + $("#safetyDiv").val())
			}

			safetynum = $("#real_safetynum").val();
			if(safetynum == ""){
				$("#real_safetynum").val(cnum);
			}else{
				$("#real_safetynum").val(safetynum + "," + cnum);
			}


			listbody = $("#safetyDivList").html();
			$("#safetyDivList").html(listbody + "<p id='l"+$("#safetyDiv").val()+"'>- " + $("#safetyDiv option:selected").text() + "("+cnum+") <input type='button' value='����' onClick='jsSafetyDivListDel("+$("#safetyDiv").val()+");' class='button'><p>");
		}else{
			alert("������ȣ�� ���� ���� : " + msgg);
			return;
		}
	}
	jsSafetyDefault();
}

function jsCallAPIsafety(certnum,isSave){
	var returnmsg = "";
	$.ajax({
		url: "/admin/itemmaster/safety_api_auth_proc.asp?itemid=<%=itemid%>&issave="+isSave+"&certnum="+certnum+"&statusmode=wait",
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
<input type="hidden" name="jungsangubun" value="<%=npt_jungsan_gubun%>">
<input type="hidden" name="companyno" value="<%=npt_company_no%>">

<input type="hidden" name="pojangok" value="Y">
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
<tr>
	<td  height="30" width="15%" bgcolor="#DDDDFF">�귣�� ������</td>
	<td bgcolor="#FFFFFF" colspan="3"><% IF oitemdetail.Fctrstate = "7" then%>���Ϸ�<%else%>�̰��<%end if%></td>
</tr>
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
		<input type="text" name="itemname" maxlength="64" size="50" class="text" value="<%= Replace(oitemdetail.Fitemname,"""","&quot;") %>" id="[on,off,off,off][��ǰ��]">&nbsp;
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
		<label><input type="radio" name="itemdiv" value="23" <%=chkIIF(oitemdetail.Fitemdiv ="23","checked","")%> onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">B2B��ǰ</label>
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
	<td height="30" width="15%" bgcolor="#DDDDFF">ISBN :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		ISBN 13 <input type="text" name="isbn13" class="text" value="<%= oitemdetail.Fisbn13 %>" size="13" maxlength="13">
		/ �ΰ���ȣ <input type="text" name="isbn_sub" class="text" value="<%= oitemdetail.FisbnSub %>" size="5" maxlength="5"><br />
		ISBN 10 <input type="text" name="isbn10" class="text" value="<%= oitemdetail.Fisbn10 %>" size="10" maxlength="10"> (Optional)
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
	  <p>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="0" <%if isNull(oitemdetail.Fsourcekind) or oitemdetail.Fsourcekind="0" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ��ǰ ��</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="1" <%if oitemdetail.Fsourcekind="1" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ����깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="2" <%if oitemdetail.Fsourcekind="2" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ���깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="3" <%if oitemdetail.Fsourcekind="3" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ��깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="4" <%if oitemdetail.Fsourcekind="4" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ����갡��ǰ</span>
	  </p>
	  <p><input type="text" name="sourcearea" maxlength="64" size="64" class="text" id="[on,off,off,off][������]"  value="<%= oitemdetail.Fsourcearea %>"/></p>
	  <div id="dvArea0" style="display:<%if isNull(oitemdetail.Fsourcekind) or oitemdetail.Fsourcekind="0" then%>block<%else%>none<%end if%>;">
	  <p><strong>ex: �ѱ�, �߱�, �߱�OEM, �Ϻ� �� </strong></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea1" style="display:<%if oitemdetail.Fsourcekind ="1" then%>block<%else%>none<%end if%>;">
	  <p><strong>������ :</strong> ����, ������ �Ǵ� �á�����, �á�����(���ѹα�, �ѱ�X)  <span style="margin-right:10px;">ex. ��(����)</span></BR>
	   <strong>���Ի� :</strong> ������� ���Ա����� <span style="margin-right:10px;">ex. ����(�߱���)</span></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea2" style="display:<%if oitemdetail.Fsourcekind ="2" then%>block<%else%>none<%end if%>;">
	  <p><strong>������ :</strong> ����,������ �Ǵ� �����ػ�(��� ���깰�� �á����� ����)   <span style="margin-right:10px;">ex. ��ġ(����), ��¡��(�����ػ�)</span> </BR>
	  	<strong>����� :</strong> ����� �Ǵ� �����(�ؿ���)   <span style="margin-right:10px;">ex. ��ġ[�����(�뼭��)]</span> </BR>
	    <strong>���Ի� :</strong> ������� ���Ա����� <span style="margin-right:10px;">ex. ���(�߱���)</span></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea3" style="display:<%if oitemdetail.Fsourcekind ="3" then%>block<%else%>none<%end if%>;">
	  <p>�Ұ���� ��� ������ ����(�ѿ�/����/���ұ���) �� ������   <span style="margin-right:10px;">ex. ����(Ⱦ���� �ѿ�), ����(ȣ�ֻ�)</span></BR>
	  - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea4" style="display:<%if oitemdetail.Fsourcekind ="4" then%>block<%else%>none<%end if%>;">
	  <p><strong>98%�̻� ���ᰡ �ִ� ���:</strong>  �Ѱ��� ���Ḹ ǥ�� ����    <span style="margin-right:10px;">ex. ����(�̱���)</span> </BR>
	  	<strong>���� ���Ḧ ����� ���:</strong> ȥ�պ����� ���� ������ 2�� ����   <span style="margin-right:10px;">ex. ������[�а���(�̱���),���尡��(������)]</span></BR>
	  - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemWeight" maxlength="12" size="8" id="[on,off,off,off][��ǰ����]" style="text-align:right" value="<%= oitemdetail.Fitemweight %>">g &nbsp;(�׷������� �Է�, ex:1.5kg�� 1500) / �ؿܹ�۽� ��ۺ� ������ ���� ���̹Ƿ� ��Ȯ�� �Է�.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�˻�Ű���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="text" name="keywords" maxlength="256" size="120" class="text" id="[on,off,off,off][�˻�Ű����]" value="<%= oitemdetail.Fkeywords %>">&nbsp;(�޸��α��� ex: Ŀ��,Ƽ����,����)
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
		<% DrawInfoDiv "infoDiv", oitemdetail.FinfoDiv, " onchange='chgInfoDiv(this.value);'" %>
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
<%
dim arrAuth, r, real_safetydiv, real_safetynum, safetyDivList
arrAuth = oitemdetail.FAuthInfo
if isArray(arrAuth) THEN
	For r =0 To UBound(arrAuth,2)
		real_safetydiv = real_safetydiv & arrAuth(0,r)
		if r <> UBound(arrAuth,2) then real_safetydiv = real_safetydiv & "," end if

		real_safetynum = real_safetynum & arrAuth(1,r)
		if r <> UBound(arrAuth,2) then real_safetynum = real_safetynum & "," end if

		safetyDivList = safetyDivList & "<p class='tPad05' id='l"&arrAuth(0,r)&"'>"
		safetyDivList = safetyDivList & "- "&fnSafetyDivCodeName(arrAuth(0,r))&"("&CHKIIF(arrAuth(1,r)="x","������ȣ ����",arrAuth(1,r))&")"
		safetyDivList = safetyDivList & " <input type='button' value='����' class='btn3 btnIntb' onClick='jsSafetyDivListDel("&arrAuth(0,r)&");'>"
		safetyDivList = safetyDivList & "</p>"
	Next
end if
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- ������������</strong></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">
		����������� :
		<input type="button" value="�������� �ʼ� ǰ�� Ȯ��" onclick="jsSafetyPopup();" class="button" />
	</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
		<tr align="left" height="30">
			<td bgcolor="#FFFFFF">
				<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitemdetail.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ���</label>
				<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitemdetail.FsafetyYn="N","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ���ƴ�</label>
				<label><input type="radio" name="safetyYn" value="I" <%=chkIIF(oitemdetail.FsafetyYn="I","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ��ǰ���� ǥ��</label>
				<label><input type="radio" name="safetyYn" value="S" <%=chkIIF(oitemdetail.FsafetyYn="S","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ���������ؼ�</label>
				<input type="hidden" name="auth_go_catecode" id="auth_go_catecode" value="">
				<input type="hidden" name="real_safetydiv" id="real_safetydiv" value="<%=real_safetydiv%>">
				<input type="hidden" name="real_safetynum" id="real_safetynum" value="<%=real_safetynum%>">
				<input type="hidden" name="real_safetyidx" id="real_safetyidx" value="">
				<input type="hidden" name="real_safetynum_delete" id="real_safetynum_delete" value="">
				<input type="hidden" name="real_safetydiv_delete" id="real_safetydiv_delete" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<% drawSelectBoxSafetyDivCode "safetyDiv", "", oitemdetail.FsafetyYn, "" %>
				������ȣ <input type="text" name="safetyNum" id="[off,off,off,off][�������� ������ȣ]" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","disabled","")%> size="35" maxlength="25" value="" /><%'=oitemdetail.FsafetyNum%>
				<input type="button" id="safetybtn" value="��   ��" onclick="jsSafetyAuth();" class="button">
				<input type="hidden" name="issafetyauth" id="issafetyauth" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<div id="safetyDivList">
					<%=safetyDivList%>
				</div>
				<div id="safetyYnI" style="display:none;">
					<font color="blue">��ǰ ���� ǥ��(ǥ���� ��ǰ�ΰ�� ��ǰ �� �������� ������ȣ�� �𵨸�, KC ��ũ�� �� ǥ�����ּ���.)</font>
				</div>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td bgcolor="#FFFFFF" colspan=2>
		* ���������� �Է� �� �ϰų�, �߸��� ���������� �Է��� ��� �߰� <strong><font color='red'>��� �Ǹ����� �Ǵ� ����</font></strong> �˴ϴ�.<br>
		* <strong><font color='red'>���������ؼ�</font></strong> ����ϰ�� ������ȣ�� ������, KC��ũ�� ǥ������ �ʾƾ� �˴ϴ�.<br>
		* �Է��� ���������� ��ǰ�����������Ϳ��� ������ ������ �������� ��ȸ�Ǹ�, <strong><font color='red'>�������� ���� ������ ����� �Ұ�</font></strong>���մϴ�.<br>
		* �������� ���������� �Է��������� �ұ��ϰ� ����� �ȵɰ�쿡 "��ǰ���� ǥ��"�� ������ �����ϸ�, ��ǰ �� �������� �𵨸�� ǥ���� ��ǰ�ΰ�� ������ȣ,KC��ũ�� ǥ���ؾ� �մϴ�.<br>
		* ������������ ���� ���Ǵ� Ȩ������(<u><a href="http://safetykorea.kr" target="_blank">http://safetykorea.kr</a></u>)�� Ȯ���� �ֽñ� �ٶ��ϴ�.
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
	<td height="30" width="15%" bgcolor="#DDDDFF">������Ź���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="mwdiv" value="M" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(mwdiv="M","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >����
	  <input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(mwdiv="W","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >��Ź
	  <input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(mwdiv="U","checked","")%>>��ü���
	  &nbsp;&nbsp; - ������Ź���п� ���� ��۱����� �޶����ϴ�. ��۱����� Ȯ�����ּ���.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��۱��� :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverytype" value="1" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="1","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >�ٹ����ٹ��&nbsp;
	  <input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="2","checked","")%>>��ü(����)���&nbsp;
	  <input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="4","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >�ٹ����ٹ�����&nbsp;
	  <input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="9","checked","")%>>��ü���ǹ��(���� ��ۺ�ΰ�)&nbsp;
	  <input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="7","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >��ü���ҹ��
	 </td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��۹�� :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverfixday" value="" onclick="TnCheckFixday(this.form)" <%=chkIIF(Trim(deliverfixday)="" or IsNull(deliverfixday),"checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >�ù�(�Ϲ�)&nbsp;
	  <input type="radio" name="deliverfixday" value="X" onclick="TnCheckFixday(this.form)" <%=chkIIF(deliverfixday="X","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >ȭ��&nbsp;
	  <input type="radio" name="deliverfixday" value="C" onclick="TnCheckFixday(this.form)" <%=chkIIF(deliverfixday="C","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >�ö��������
	  <input type="radio" name="deliverfixday" value="G" onclick="TnCheckFixday(this.form)" <%=chkIIF(deliverfixday="G","checked","")%> <%=chkIIF(mwdiv<>"U" or (deliverytype <> "2" and deliverytype <> "9")," disabled","")%> >�ؿ�����
		<span id="lyrFreightRng" style="display:<%=chkIIF(deliverfixday="X","","none")%>;">
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
	  <input type="radio" name="deliverarea" value="" <%=chkIIF(Trim(deliverarea)="" or IsNull(deliverarea),"checked","")%>>�������&nbsp;
	  <input type="radio" name="deliverarea" value="C" <%=chkIIF(deliverarea="C","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >�����ǹ��&nbsp;
	  <input type="radio" name="deliverarea" value="S" <%=chkIIF(deliverarea="S","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >������&nbsp;
 	  <input type="checkbox" name="deliverOverseas" value="Y" <%=chkIIF(oitemdetail.FdeliverOverseas="Y","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> title="�ؿܹ���� ��ǰ���԰� �Է��� �ž� �Ϸ�˴ϴ�.">�ؿܹ��
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
	<input type="hidden" name="optionaddprice" value="<%= oitemdetail.fnGetWaitOptAddPrice(itemid) %>">
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
		<% if ubound(itemaddimage) >= 0 then %>
	  <% if (itemaddimage(0) <> "") then %>
		<div id="divimgadd1" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add1/<%=imgsubdir%>/<%=itemaddimage(0) %>" width="300" height="300">
		</div>
	  <% end if %>
		<% end if %>
		<input type="file" name="imgadd1" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd1,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add1">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���2 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if ubound(itemaddimage) >= 1 then %>
	  <% if (itemaddimage(1) <> "") then %>
		<div id="divimgadd2" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add2/<%=imgsubdir%>/<%=itemaddimage(1) %>" width="300" height="300">
		</div>
	  <% end if %>
		<% end if %>
		<input type="file" name="imgadd2" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd2,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���3 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if ubound(itemaddimage) >= 2 then %>
	  <% if (itemaddimage(2) <> "") then %>
		<div id="divimgadd3" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add3/<%=imgsubdir%>/<%=itemaddimage(2) %>" width="300" height="300">
		</div>
	  <% end if %>
		<% end if %>
		<input type="file" name="imgadd3" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd3,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add3">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���4 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if ubound(itemaddimage) >= 3 then %>
	  <% if (itemaddimage(3) <> "") then %>
		<div id="divimgadd4" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add4/<%=imgsubdir%>/<%=itemaddimage(3) %>" width="300" height="300">
		</div>
	  <% end if %>
		<% end if %>
		<input type="file" name="imgadd4" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd4,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add4">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ �߰��̹���5 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if ubound(itemaddimage) >= 4 then %>
	  <% if (itemaddimage(4) <> "") then %>
		<div id="divimgadd5" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add5/<%=imgsubdir%>/<%=itemaddimage(4) %>" width="300" height="300">
		</div>
	  <% end if %>
		<% end if %>
		<input type="file" name="imgadd5" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd5,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add5">
	</td>
</tr>
<tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>�� ������ ��ǰ�����̹����� ������� �ʰ� ��ǰ�����̹����� ����մϴ�. ������ ��ϵ� ��ǰ�����̹����� ����� �ϵ� �߰� ������ �����ʰ� ������ �˴ϴ�.</strong></font>
 	</td>
 </tr>
<%
			'��ǰ���� �̹���
			mainImg(1) = ooimage.Fimgmain
			mainImg(2) = ooimage.Fimgmain2
			mainImg(3) = ooimage.Fimgmain3
			mainImg(4) = ooimage.Fimgmain4
			mainImg(5) = ooimage.Fimgmain5
			mainImg(6) = ooimage.Fimgmain6
			mainImg(7) = ooimage.Fimgmain7
			mainImg(8) = ooimage.Fimgmain8
			mainImg(9) = ooimage.Fimgmain9
			mainImg(10) = ooimage.Fimgmain10

			for i=1 to 7
%>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #<%=i%> :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (mainImg(i) <> "") then %>
		<div id="divimgmain<%=chkIIF(i>1,i,"")%>" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/main<%=chkIIF(i>1,i,"")%>/<%=imgsubdir%>/<%=mainImg(i) %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain<%=chkIIF(i>1,i,"")%>" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgmain<%=chkIIF(i>1,i,"")%>" onchange="CheckImage(this, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmain<%=chkIIF(i>1,i,"")%>, 40, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>)"> (����, �ʺ� <%=CMAIN_IMG_MAXWIDTH%>px, Max <%= getFileSize(CMAIN_IMG_MAXSIZE) %>, jpg,gif,png)
		<input type="hidden" name="main<%=chkIIF(i>1,i,"")%>">
	</td>
</tr>
<% next %>
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>�� ����� ��ǰ�� �̹����� ������ �� �������� ��ü �˴ϴ�. html�� ������� ���� �����̿��� �������� ���ε� ���ֽñ� �ٶ��ϴ�.<br>�� ����� ��ǰ�󼼿��� �̹����� �߶� �÷��ֽñ� �ٶ��ϴ�.</strong></font>
 	</td>
 </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain <> "") then %>
		<div id="divmobileimgmain" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #2:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain2 <> "") then %>
		<div id="divmobileimgmain2" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile2/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain2 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain2" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain2" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile2', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #3:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain3 <> "") then %>
		<div id="divmobileimgmain3" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile3/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain3 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain3" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain3" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile3', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile3">
	</td>
</tr>

<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #4:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain4 <> "") then %>
		<div id="divmobileimgmain4" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile4/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain4 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain4" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain4" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile4', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile4">
	</td>
</tr>

<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #5:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain5 <> "") then %>
		<div id="divmobileimgmain5" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile5/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain5 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain5" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain5" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile5', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile5">
	</td>
</tr>

<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #6:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain6 <> "") then %>
		<div id="divmobileimgmain6" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile6/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain6 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain6" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain6" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile6', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile6">
	</td>
</tr>

<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #7:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain7 <> "") then %>
		<div id="divmobileimgmain7" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile7/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain7 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain7" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain7" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile7', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile7">
	</td>
</tr>
<!-- 20160601�߰��Ѻκ� -->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #8:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain8 <> "") then %>
		<div id="divmobileimgmain8" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile8/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain8 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain8" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain8" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile8', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile8">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #9:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain9 <> "") then %>
		<div id="divmobileimgmain9" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile9/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain9 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain9" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain9" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile9', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile9">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #10:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain10 <> "") then %>
		<div id="divmobileimgmain10" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile10/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain10 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain10" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain10" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile10', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile10">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #11:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain11 <> "") then %>
		<div id="divmobileimgmain11" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile11/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain11 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain11" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain11" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile11', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile11">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����� ��ǰ���̹��� #12:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain12 <> "") then %>
		<div id="divmobileimgmain12" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile12/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain12 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain12" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain12" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage2('mobile12', 40, 640, 1200)"> (����,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile12">
	</td>
</tr>

<!--// 20160601�߰��Ѻκ� -->

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
	 		IF arrList(2,intLoop) = "2" THEN
	 			count2 = count2 + 1
	 			strMsg = count2&"��"
	 			sMsgCd2 = sMsgCd2 + "^" + arrList(6,intLoop)
	 		ELSEIF arrList(2,intLoop) = "0" THEN
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

 <% if oitemdetail.FCurrState="1" or oitemdetail.FCurrState="5" then %>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
	 <tr>
	 	  <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
			<td style="padding:5px">
				<input type="checkbox" name="chkSR" value="Y" onClick="jsChkSellReserve();"> ��ǰ���¿���:
				<input type="text" name="dSR" value="" size="10" class="input"   onClick="jsPopCal('dSR');">
				<input type="image" name="imgSR" src="/images/admin_calendar.png" onClick="jsPopCal('dSR');"  >
				  ��ǰ ������ ���� �ʾҰų�, ������ ������ ��� ����� �ð��� ������ ���� �ʽ��ϴ�.
			   �ٹ����� ����� ���, �԰� Ȯ�� �� ���¿����� �����մϴ�.
				</td>
			<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tR>
</table>
<% end if %>
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

<% if application("Svr_Info")	= "Dev" then %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="600" height="600"></iframe>
<% else %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<div id="boxes"></div>
<div id="mask"></div>
<div id="dialog">
<!-- #include virtual="/admin/itemmaster/item_confirm_inc.asp"-->
</div>
<script type="text/javascript">
// ������Ź���� �� ��۱��м���
TnCheckUpcheYN(itemreg);
for (var i = 0; i < itemreg.elements.length; i++) {
    if (itemreg.elements[i].name == "deliverytype") {
        if (itemreg.elements[i].value == "<%= oitemdetail.Fdeilverytype %>") {
            itemreg.elements[i].checked = true;
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

	// ��������üũ. ���ȹ�
	jsSafetyCheck('<%= oitemdetail.FsafetyYn %>','');
</script>
<%
set oitemdetail = Nothing
set oitemreg = Nothing
set ooimage = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->