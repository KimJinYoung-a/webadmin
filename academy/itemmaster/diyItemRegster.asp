<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
'CONST CBASIC_IMG_MAXSIZE = 180   'KB
'CONST CMAIN_IMG_MAXSIZE = 500   'KB

'2016 ������ �����̹��� ���� �뷮 ����?
CONST CBASIC_IMG_MAXSIZE = 600   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB

dim i,j
'==============================================================================
Sub SelectBoxDesignerItem()
   dim query1
   %><select name="designer" onchange="TnDesignerNMargineAppl(this.value);">
     <option value=''>-- ��ü���� --</option><%
	query1 = " select U.userid, U.socname_kor, L.diy_margin as defaultmargine, U.maeipdiv, IsNULL(L.defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit "
	query1 = query1 + " 	, IsNULL(L.DefaultDeliveryPay,0) as defaultDeliverPay, IsNULL(L.diy_dlv_gubun,'') as defaultDeliveryType "
	query1 = query1 + " from [TENDB].[db_user].dbo.tbl_user_c as U "
	query1 = query1 + "		Left Join [db_academy].[dbo].tbl_lec_user as L "
	query1 = query1 + " 		on U.userid=L.lecturer_id "
	query1 = query1 + " where U.isusing='Y' and U.userid<>'' and U.userdiv ='14' "
	rsACADEMYget.CursorLocation = adUseClient
    rsACADEMYget.Open query1,dbACADEMYget,adOpenForwardOnly, adLockReadOnly

   if  not rsACADEMYget.EOF  then
       do until rsACADEMYget.EOF
           response.write("<option value='"&rsACADEMYget("userid")& "," & rsACADEMYget("defaultmargine") & "," & rsACADEMYget("maeipdiv") & "," & rsACADEMYget("defaultFreeBeasongLimit") & "," & rsACADEMYget("defaultDeliverPay") & "," & rsACADEMYget("defaultDeliveryType") & "'>" & rsACADEMYget("userid") & "  [" & replace(db2html(rsACADEMYget("socname_kor")),"'","") & "]" & "</option>")
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   response.write("</select>")
End Sub

%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript">
<!-- #include file="./itemregister_javascript.asp"-->
</script>
<script>
function UseTemplate() {
	var popwin = window.open("/academy/comm/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
    popwin.focus();
}

function UseTemplateTen(){
    var popwin = window.open("/common/pop_basic_item_info_list.asp?tp=academydiy", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
    popwin.focus();
}

// ============================================================================
// ��ü�����ڵ��Է�
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
        document.itemreg.mwdiv[1].checked = true; //Ư��
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

	isvatinclude = frm.vatyn[0].checked;

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
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		imileage = parseInt(isellcash*0.01) ;
	}else{
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		imileage = parseInt(isellcash*0.01) ;
	}

	frm.buycash.value = ibuycash;
	frm.mileage.value = imileage;
}

// ============================================================================
// ī�װ����
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

	popwin = window.open('/academy/comm/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
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
			url: "/academy/comm/act_DispCategorySelect.asp?designerid="+designerid+"&isDft="+dCnt,
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
		// ������ ���� �ߺ� ī�װ� ���� �˻�
		var tbl_Category = document.all.tbl_Category;
		if(tbl_Category.rows.length>0)	{
			if(tbl_Category.rows.length>1)	{
				for(l=0;l<document.all.cate_div.length;l++)	{
					if((document.all.cate_large[l].value==lcd)&&(document.all.cate_mid[l].value==mcd)) {
						alert("���� �ߺз��� �̹� ������ ī�װ��� �ֽ��ϴ�.\n���� ī�װ��� �����ϰ� �ٽ� �������ּ���.");
						return;
					}
				}
			}
			else {
				if((document.all.cate_large.value==lcd)&&(document.all.cate_mid.value==mcd)) {
					alert("���� �ߺз��� �̹� ������ ī�װ��� �ֽ��ϴ�.\n�ر��� ī�װ��� �����ϰ� �ٽ� �������ּ���.");
					return;
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

	// ���� ī�װ� ����
	function delCateItem()
	{
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?"))
			tbl_Category.deleteRow(tbl_Category.clickedRowIndex);
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
		//printItemAttribute();
	}

	// ���� ����ī�װ� ����
	function delDispCateItem() {
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			//��ǰ�Ӽ� ���
			//printItemAttribute();
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


// ============================================================================
// �����ϱ�
function SubmitSave() {
	var itemreg = document.all.itemreg;

	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
		return;
	}

	if (!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("[�⺻] ���� ī�װ��� �����ϼ���.\n�� [�߰�] ���� ī�װ��� ���� �� �����ϴ�.");
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }

	if (itemreg.itemsize.value!=''){
		if (itemreg.unit.value!=''){
			itemreg.itemsize.value=itemreg.itemsize.value + '(' + itemreg.unit.value + ')';
		}
	}

	if (itemreg.itemWeight.value==''){
		itemreg.itemWeight.value='0';
	}

	// ��� ���� üũ

	// �Է��� ������ �ٸ���� üũ
    if (itemreg.margin.value.length>0){
        if (parseInt((itemreg.sellcash.value*1) * (itemreg.margin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {
    		alert("���ް��� �߸��ԷµǾ����ϴ�.[�Һ��ڰ�*���� = ���ް�]");


    		if (!confirm('�Է��� ������ �Էµ� �ǸŰ� ��� ���԰� �ݾ��� ���� �մϴ�. ��� ���� �Ͻðڽ��ϱ�?')){
    		    itemreg.sellcash.focus();
    			return;
    		}
        }
	}

	// ��ü �⺻������ �ٸ���� üũ
	if (itemreg.defaultmargin.value.length>0){
        if (parseInt((itemreg.sellcash.value*1) * (itemreg.defaultmargin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {

    		if (!confirm('��ü �⺻ ������ �Էµ� �ǸŰ� ��� ���԰� �ݾ��� ���� �մϴ�. ��� ���� �Ͻðڽ��ϱ�?')){
    			return;
    		}
        }
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


    if (itemreg.buycash.value*1>itemreg.sellcash.value*1){
        alert("���԰����� �ǸŰ� ���� Ů�ϴ�.");
		itemreg.sellcash.focus();
		return;
    }

	if (itemreg.sellcash.value*1 < 500 || itemreg.sellcash.value*1 >= 20000000){
		alert("�Ǹ� ������ 500�� �̻� 20,000,000�� �̸����� ��� �����մϴ�.");
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

    if (itemreg.imgbasic.value == "") {
        //alert("�⺻�̹����� �ʼ��Դϴ�.");
        //return;
    } else {
        if (CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg') != true) {
            return;
        }
    }

    if (itemreg.imgadd1.value != "") {
        if (CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg,gif') != true) {
            return;
        }
    }

//    if (itemreg.imgadd4.value != "") {
//        if (CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg,gif') != true) {
//            return;
//        }
//    }
//
//    if (itemreg.imgadd5.value != "") {
//        if (CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 1000 , 667, 'jpg,gif') != true) {
//            return;
//        }
//    }

//    if (itemreg.imgmain.value != "") {
//        if (CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg') != true) {
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

    if(confirm("��ǰ�� �ø��ðڽ��ϱ�?") == true){
        itemreg.itemoptioncode2.value = optionv;
        itemreg.itemoptioncode3.value = optiont;

		itemreg.deliverytype[0].disabled=false;
		itemreg.deliverytype[1].disabled=false;
		itemreg.deliverytype[2].disabled=false;
		//itemreg.target = "_blank";
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
        // frm.deliverOverseas.checked=true;	// �ؿܹ��üũ
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
        // frm.deliverOverseas.checked=false;	// �ؿܹ��üũ����
		//  frm.optlevel[1].disabled=false;
	}
}

function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if (frm.mwdiv[2].checked){
			alert("����Ư�� ������ ��ü�� ���\n��۱����� �ٹ����� ������� ���� �Ͻ� �� �����ϴ�!!\n����Ư�������� Ȯ�����ּ���!!");
			frm.mwdiv[0].checked=true;
			//frm.optlevel[1].checked=false;
			//frm.optlevel[1].disabled=true;
		}
	}
	else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked || frm.deliverytype[4].checked){
	//else if(frm.deliverytype[1].checked){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("����Ư�� ������ �����̳� Ư���� ���\n��۱�����  ��ü������� ���� �Ͻ� �� �����ϴ�!!!\n����Ư�������� Ȯ�����ּ���!!");
			frm.mwdiv[2].checked=true;
			//frm.optlevel[1].disabled=false;
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
	var MaxSize=<%= CBASIC_IMG_MAXSIZE %>;
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
		<font color="red"><strong>��ǰ���</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>�Ż�ǰ�� ����մϴ�.</b>
			<!--
            <br>- ���� ȭ���� ���� ����ϼž� �����Ͽ� ���� �� ������Ʈ �˴ϴ�.
            <br>- �����̳� ������ ������ ��� ���� �źε� �� �ֽ��ϴ�.
            -->
			<br>- �⺻Ʋ������ �̿��Ͽ� ������ ��ǰ�� ����Ҽ� �ֽ��ϴ�.
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
<input type="button" class="button" value="�⺻Ʋ����" onClick="UseTemplate();">

&nbsp;&nbsp;
<input type="button" class="button" value="�⺻Ʋ����(�ٹ����ٻ�ǰ)" onClick="UseTemplateTen();">
<br><br>
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
<form name="itemreg" method="post" action="<%= uploadImgUrl %>/linkweb/academy/items/itemregisterWithImage_process.asp" onsubmit="return false;" enctype="multipart/form-data">
<!--<form name="itemreg" method="post" action="<%'= UploadImgFingers %>/linkweb/items/itemregisterWithImage_process.asp" onsubmit="return false;" enctype="multipart/form-data">-->
<input type="hidden" name="designerid">
<input type="hidden" name="defaultmargin">
<input type="hidden" name="defaultmaeipdiv">
<input type="hidden" name="defaultFreeBeasongLimit">
<input type="hidden" name="defaultDeliverPay">
<input type="hidden" name="defaultDeliveryType">
<input type="hidden" name="itemoptioncode2">
<input type="hidden" name="itemoptioncode3">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ü�� :</td>
  	<td bgcolor="#FFFFFF" colspan="3"><% SelectBoxDesignerItem %> (����ü�� ǥ�õ˴ϴ�)</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">���� ī�װ� :</td>
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
		<td height="30" width="15%" bgcolor="#DDDDFF" title="����Ʈ�� ������ ī�װ�" style="cursor:help;">���� ī�װ� :</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<table class=a>
			<tr>
				<td id="lyrDispList"><table class="a" id="tbl_DispCate"></table></td>
				<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
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
	  <input type="checkbox" name="requireimgchk" value="Y" onClick="requireimg();">�ֹ����� �̹��� �ʿ�
<!-- 	  <br> -->
<!--       <input type="radio" name="itemdiv" value="20" onclick="checkItemDiv(this);chgodr(1);">�߰������ǰ -->
<!--       <font color="red">(��ǰ��Ͽ����� ����, �߰��ɼǿ����� ������)</font> -->
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
  <tr id="customorder" style="display:none;">
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
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF">
      <input type="text" name="itemWeight" maxlength="12" size="4" id="[on,off,off,off][��ǰ����]" value="0">g
      &nbsp;(���Դ� g������ �Է�)
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

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="margin" maxlength="32" size="5" id="[on,off,off,off][����]">%
      <input type="button" value="���ް� �ڵ����" onclick="CalcuAuto(itemreg);">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ǸŰ�(�Һ��ڰ�) :</td>
  	<input type="hidden" name="sellvat">
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="sellcash" maxlength="16" size="12" id="[on,on,off,off][�Һ��ڰ�]" onKeyup="CalcuAuto(itemreg);">��
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">���ް� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="buycash" maxlength="16" size="12" id="[on,on,off,off][���ް�]" >��
      (<b>�ΰ��� ���԰�</b>)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">���ϸ��� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" class="text_ro" name="mileage" maxlength="32" size="10" id="[on,on,off,off][���ϸ���]" value="0" ReadOnly > (�ǸŰ��� 1%)
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">����, �鼼 ���� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="radio" name="vatyn" value="Y" checked>����
      <input type="radio" name="vatyn" value="N">�鼼
  	</td>
  </tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>�Ǹ�����
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
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
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverytype" value="1" checked onclick="TnCheckUpcheDeliverYN(this.form);">�ٹ����ٹ��&nbsp;
      <input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);">��ü(����)���&nbsp;
	  <label><input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);">�ٹ����ٹ�����</label>&nbsp;
      <input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);">��ü���ǹ��(���� ��ۺ�ΰ�)
      <input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);">��ü���ҹ��
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�Ǹſ��� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	  <input type="radio" name="sellyn" value="Y">�Ǹ���&nbsp;&nbsp;
  	  <input type="radio" name="sellyn" value="N" checked>�Ǹž���
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">��뿩�� :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="radio" name="isusing" value="Y" checked>�����&nbsp;&nbsp;
  	  <input type="radio" name="isusing" value="N" disabled>������
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ ���� :</td> -->
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
      <textarea name="refundpolicy" rows="5" cols="80" id="[off,off,off,off][ȯ����å]" >
	  - ��ǰ/ȯ���� ��ǰ�����Ϸκ��� 7�� �̳��� �����մϴ�.
	  - ��� ���� ȯ�ҿ�û �� ��ǰ ȸ�� �� ó���˴ϴ�.
	  - ���� ��ǰ�� ��� �պ���ۺ� ������ �ݾ��� ȯ�ҵǸ�, ��ǰ �� ���� ���°� ���Ǹ� �����Ͽ��� �մϴ�.
	  - ��ǰ �ҷ��� ���� ��ۺ� ������ ������ ȯ�ҵ˴ϴ�.
	  - ����ǰ���� ���Ե� ��ǰ�� ��� A/S�� �Ұ��մϴ�.
	  - ��ȯ/ȯ��/��ۺ�ȳ�/AS�� ���� ���������� ��ǰ�������� �ִ� ��� �۰����� ���������� �켱 ���� �˴ϴ�.
	  </textarea><br>
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
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ɼ� ��� ���� :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="useoptionyn" value="Y" onClick="TnCheckOptionYN(this.form);" disabled>�ɼǻ����&nbsp;&nbsp;
      <input type="radio" name="useoptionyn" value="N" onClick="TnCheckOptionYN(this.form);" checked>�ɼǻ�����&nbsp;&nbsp;
	  <font color="red">** �ɼ��� ��ǰ��� �� �߰��ϼ���.</font>
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
              <input type="button" value="�ɼ��߰�" name="btnetcoptadd" onclick="popEtcOptionAdd();">
              <input type="button" value="���ÿɼǻ���" name="btnoptdel" onclick="delItemOptionAdd()" >
              <br><br>
              - �ɼ��߰� : ��ǰ�ɼ��� �߰��Ͻ� �� �ֽ��ϴ�.<br>
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
            <td>
                <input type="hidden" name="itemoption<%= j+1 %>" value="">
                <input type="text" name="optionName<%= j+1 %>" size="18" maxlength="20" id="[off,off,off,off][�ɼǸ�<%= i %><%= j %>]">
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
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          �̹�������
          <br>- �̹����� <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> ���� �ø��� �� �ֽ��ϴ�.
          <br>&nbsp;&nbsp;(�̹�������� <font color=red>���μ������� ������</font>�� �԰ݿ� ���� �ʰ� ������ּ���. �԰��ʰ��� ����� ���� �ʽ��ϴ�.)
          <br>- <font color=red>����޿��� Save For Web</font>���� ����� �� �÷��ֽñ� �ٶ��ϴ�.
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
      <input type="button" value="�̹��������" onClick="ClearImage('imgbasic')"><input type="hidden" name="basic"> (<font color=red>�ʼ�</font>,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,<b><font color="red">jpg</font></b>)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd1" onchange="CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd1')"><input type="hidden" name="add1"> (����,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd2" onchange="CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd2')"><input type="hidden" name="add2"> (����,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd3" onchange="CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40">
      <input type="button" value="�̹��������" onClick="ClearImage('imgadd3')"><input type="hidden" name="add3"> (����,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���4 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--  -->
<!--       <input type="file" name="imgadd4" onchange="CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');" size="40"> -->
<!--       <input type="button" value="�̹��������" onClick="ClearImage('imgadd4')"><input type="hidden" name="add4"> (����,1000x667,jpg,gif) -->
<!--   	</td> -->
<!--   </tr> -->
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���5 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--   	  <input type="file" name="imgadd5" onchange="CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');" size="40"> -->
<!--       <input type="button" value="�̹��������" onClick="ClearImage('imgadd5')"><input type="hidden" name="add5"> (����,1000x667,jpg,gif) -->
<!--    	</td> -->
<!--   </tr> -->
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� :<br/>������ ���� ������</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="file" name="imgmain" onchange="CheckImage('imgmain', 1024, 610, 2000, 'jpg,gif');" size="40"> -->
<!--       <input type="button" value="�̹��������" onClick="ClearImage('imgmain')"><input type="hidden" name="main"> (����,600X2000,1024KB,jpg,gif) -->
<!--   	</td> -->
<!--   </tr> -->
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
<!-- <tr align="left" id="lyItemSrc" style="display:none;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsource" maxlength="64" size="50" class="text">&nbsp;(ex:�ö�ƽ,����,��,...) -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr align="left" id="lyItemSize" style="display:none;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ������ :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsize" maxlength="64" size="50" class="text"> -->
<!-- 		<select name="unit" class="select"> -->
<!-- 		<option value="">�����Է�</option> -->
<!-- 		<option value="mm">mm</option> -->
<!-- 		<option value="cm" selected>cm</option> -->
<!-- 		<option value="m��">m��</option> -->
<!-- 		<option value="km">km</option> -->
<!-- 		<option value="m��">m��</option> -->
<!-- 		<option value="km��">km��</option> -->
<!-- 		<option value="ha">ha</option> -->
<!-- 		<option value="m��">m��</option> -->
<!-- 		<option value="cm��">cm��</option> -->
<!-- 		<option value="L">L</option> -->
<!-- 		<option value="g">g</option> -->
<!-- 		<option value="Kg">Kg</option> -->
<!-- 		<option value="t">t</option> -->
<!-- 		</select> -->
<!-- 		&nbsp;(ex:7.5x15(cm)) -->
<!-- 		</td> -->
<!-- </tr> -->
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
      <input type="button" value="#1 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[0],40, 1000, 667)"> (����,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #2 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#2 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[1],40, 1000, 667)"> (����,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #3 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#3 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[2],40, 1000, 667)"> (����,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #4 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#4 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[3],40, 1000, 667)"> (����,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #5 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#5 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[4],40, 1000, 667)"> (����,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #6 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#6 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[5],40, 1000, 667)"> (����,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
	   <br/><span style="color:red;font-size:15px"><strong>���̹��� ��� ���� ���� �ø� �� �����ϴ�.��</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���̹��� #7 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40);CheckImageSize(this);" class="text" size="40">
      <input type="button" value="#7 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[6],40, 1000, 667)"> (����,1000x667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
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
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
