//###########################################################
// Description : �¶��� & �������� ���� ��ٱ��� JS
// Hieditor : 2011.08.02 �ѿ�� ����
//###########################################################

//��ٱ��� ��ǰ�߰�	//onoffgubun ON:�¶��� OFF;��������
function adminshoppingbagreg(upfrm,onoffgubun,shopid){
    var frm;
    var pass = false;
    var searchfrm = document.frm;

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    var ret;

    if (!pass) {
        alert('���� �������� �����ϴ�.');
        return;
    }

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){

                if (!IsInteger(frm.itemno.value)){
                    alert('������ ������ �����մϴ�.');
                    frm.itemno.focus();
                    return;
                }

                if (frm.itemno.value==""){
                    alert('������ �Է��ϼ���.');
                    frm.itemno.focus();
                    return;
                }

                //if (frm.itemno.value=="0"){
                //    alert('0�� �ƴ� ������ �Է��ϼ���.');
                //    frm.itemno.focus();
                //    return;
                //}

                upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + ",";
                upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + ",";
                upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + ",";
                upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + ",";
            }
        }
    }

    //��ǰ������ ��ٱ��� �������� �ѱ�..  ��ٱ��� ��� �׼��� ��ٱ��� ������ ������ ó����..
    var popbag = window.open('','popbag','width=1024,height=768,scrollbars=yes,resizable=yes');
    upfrm.action='/common/item/adminshoppingbag.asp';
    upfrm.target='popbag';
    upfrm.onoffgubun.value=onoffgubun;

	//���������� ��쿡�� �������� ����
    if (onoffgubun=='OFF'){
    	upfrm.shopid.value=shopid;
    }

    upfrm.submit();
    popbag.focus();

	//���� ������ ��ǰ ������ ���� ��� ����
	upfrm.itemgubunarr.value = '';
	upfrm.itemidarr.value = '';
	upfrm.itemoptionarr.value = '';
	upfrm.itemnoarr.value = '';
	upfrm.onoffgubun.value = '';

    if (onoffgubun=='OFF'){
		upfrm.shopid.value = '';
    }

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
                frm.cksel.checked = false;
                frm.itemno.value="0"
                dL(frm.cksel);
            }
        }
    }
}

//��ٱ��� ��ǰ��ǰ�߰�	//onoffgubun ON:�¶��� OFF;��������
function adminshoppingbagregoneitem(onoffgubun,shopid,upfrm){
    if (onoffgubun=="" || upfrm.itemgubun.value=="" || upfrm.itemgubun.value=="" || upfrm.itemid.value=="" || upfrm.itemoption.value==""){
        alert('���� �����ϴ�');
        upfrm.itemno.focus();
        return;
    }

	//���������� ���
    if (onoffgubun=='OFF'){
    	//�������� ����
	    if (shopid==""){
	        alert('������ �����ϴ�');
	        return;
	    }
    }

    if (!IsInteger(upfrm.itemno.value)){
        alert('������ ������ �����մϴ�.');
        upfrm.itemno.focus();
        return;
    }

    if (upfrm.itemno.value==""){
        alert('������ �Է��ϼ���.');
        upfrm.itemno.focus();
        return;
    }

    var tmp = '&itemgubunarr='+upfrm.itemgubun.value+',&itemidarr='+upfrm.itemid.value+',&itemoptionarr='+upfrm.itemoption.value+',&itemnoarr='+upfrm.itemno.value+',';
    var popbag = window.open('/common/item/adminshoppingbag_process.asp?mode=directbagaddarr&onoffgubun='+onoffgubun+'&shopid='+shopid+tmp,'popbag','width=1024,height=768,scrollbars=yes,resizable=yes');
	popbag.focus();
}

//��ٱ��� ����	//onoffgubun ON:�¶��� OFF;��������
function adminshoppingbagview(upfrm,onoffgubun,shopid){

    var popbag = window.open('','popbag','width=1024,height=768,scrollbars=yes,resizable=yes');
    upfrm.onoffgubun.value=onoffgubun;

    if (onoffgubun=='OFF'){
    	//upfrm.shopid.value=shopid;
    }

    upfrm.action='/common/item/adminshoppingbag.asp';
    upfrm.target='popbag';
    upfrm.submit();
    popbag.focus();
}

//�ʿ����Ŭ����	������ �ʿ������ �������� �ִ´�
function inputiteno(shortitemno,formi){
    formi.itemno.value=shortitemno;

    formi.cksel.checked=true;
    AnCheckClick(formi.cksel);
}

//�˻���ư
function reg(upfrm){

    if(upfrm.itemid.value!=''){
        if (!IsDouble(upfrm.itemid.value)){
            alert('��ǰ�ڵ�� ���ڸ� �����մϴ�.');
            upfrm.itemid.focus();
            return;
        }
    }

    upfrm.submit();
}

//��ٱ��� ����
function bageditarr(upfrm){
    var frm;
    var pass = false;

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    var ret;

    if (!pass) {
        alert('���� �������� �����ϴ�.');
        return;
    }

    upfrm.bagidxarr.value = "";

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){

                if (!IsInteger(frm.itemno.value)){
                    alert('������ ������ �����մϴ�.');
                    frm.itemno.focus();
                    return;
                }

                if (frm.itemno.value=="0"){
                    alert('������ �Է��ϼ���.');
                    frm.itemno.focus();
                    return;
                }

				upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + ",";
                upfrm.bagidxarr.value = upfrm.bagidxarr.value + frm.bagidx.value + ",";
            }
        }
    }

    upfrm.action='/common/item/adminshoppingbag_process.asp';
    upfrm.mode.value='bageditarr';
    upfrm.target='view';
    upfrm.submit();
}

//��ٱ��� ��ǰ ����
function bagdelarr(upfrm){
    var frm;
    var pass = false;

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    var ret;

    if (!pass) {
        alert('���� �������� �����ϴ�.');
        return;
    }

    upfrm.bagidxarr.value = "";

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
                upfrm.bagidxarr.value = upfrm.bagidxarr.value + frm.bagidx.value + ",";
            }
        }
    }

    upfrm.action='/common/item/adminshoppingbag_process.asp';
    upfrm.mode.value='bagdelarr';
    upfrm.target='view';
    upfrm.submit();
}

//�ٹ����ٹ��� �ֹ��� �ۼ�
function AddArr(upfrm ,shopgubun){
    var frm; var tmpshopid = ''; var tmpcomm_cd013 = ''; var tmpcomm_cd011 = ''; var tmpcomm_cd031 = '';
    var pass = false;

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    var ret;

    if (!pass) {
        alert('���� �������� �����ϴ�.');
        return;
    }

    upfrm.itemgubunarr2.value = "";
    upfrm.itemidadd2.value = "";
    upfrm.itemoptionarr2.value = "";
    upfrm.sellcasharr2.value = "";
    upfrm.suplycasharr2.value = "";
    upfrm.buycasharr2.value = "";
    upfrm.itemnoarr2.value = "";
    upfrm.itemnamearr2.value = "";
    upfrm.itemoptionnamearr2.value = "";
    upfrm.designerarr2.value = "";
    upfrm.bagidxarr.value = "";

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){

                //�ٹ����ٹ��� �ֹ��� ��� �ֹ��ڰ� �����ؾ���
                if (tmpshopid==''){
                	tmpshopid = frm.shopid.value;
                } else {
                	if (tmpshopid != frm.shopid.value){
	                    alert('���� Ʋ�� ������ ���õǾ� �ֽ��ϴ� \n����(�ֹ���)�� ���� �ؾ� �մϴ�.');
	                    return;
                	}
                }

				//�ٹ����ٹ��� �ֹ��� ��� ��������� �ٹ�������Ź�� �ֹ�����
                if (frm.comm_cd.value=="B012" || frm.comm_cd.value=="B022"){
                    alert('��ü��Ź�̳� ��ü������ �ֹ� �ϽǼ� �����ϴ�.');
                    frm.itemno.focus();
                    return;
                }

				//�ٹ����� ��Ź �ֹ�
                if (frm.comm_cd.value=="B011" && tmpcomm_cd011==''){
                	tmpcomm_cd011 = frm.comm_cd.value;
                }
				//������
                if (frm.comm_cd.value=="B031" && tmpcomm_cd031==''){
                	tmpcomm_cd031 = frm.comm_cd.value;
                }
				//�����Ź
                if (frm.comm_cd.value=="B013" && tmpcomm_cd013==''){
                	tmpcomm_cd013 = frm.comm_cd.value;
                }

                //�ٹ����ٹ��� �ֹ� �����Ź�� ���, �����Ź������ �ֹ��� ������
                if (tmpcomm_cd013 != ''){

                	if (tmpcomm_cd011 != '' || tmpcomm_cd031 != ''){
                		alert('�����Ź�� ���, �ٹ�������Ź�� ������ �ֹ��� ���� �ֹ��ϽǼ� �����ϴ�.');
                		return;
                	}
                	upfrm.cwflag.value='1';
                }else{
                	upfrm.cwflag.value='0';
                }

                if (!IsInteger(frm.itemno.value)){
                    alert('������ ������ �����մϴ�.');
                    frm.itemno.focus();
                    return;
                }

                if (frm.itemno.value=="0"){
                    alert('������ �Է��ϼ���.');
                    frm.itemno.focus();
                    return;
                }

                upfrm.itemgubunarr2.value = upfrm.itemgubunarr2.value + frm.itemgubun.value + "|";
                upfrm.itemidadd2.value = upfrm.itemidadd2.value + frm.itemid.value + "|";
                upfrm.itemoptionarr2.value = upfrm.itemoptionarr2.value + frm.itemoption.value + "|";
                upfrm.sellcasharr2.value = upfrm.sellcasharr2.value + frm.shopitemprice.value + "|";

                //[db_storage].[dbo].tbl_ordersheet_master�� ���� ������ ��� ���͸��԰��� , ������԰��� ���ٷ���
                upfrm.suplycasharr2.value = upfrm.suplycasharr2.value + frm.shopbuyprice.value + "|";
                upfrm.buycasharr2.value = upfrm.buycasharr2.value + frm.shopsuplycash.value + "|";

                upfrm.itemnoarr2.value = upfrm.itemnoarr2.value + frm.itemno.value + "|";
                upfrm.itemnamearr2.value = upfrm.itemnamearr2.value + frm.itemname.value + "|";
                upfrm.itemoptionnamearr2.value = upfrm.itemoptionnamearr2.value + frm.itemoptionname.value + "|";
                upfrm.designerarr2.value = upfrm.designerarr2.value + frm.makerid.value + "|";
				frmbag.bagidxarr.value = frmbag.bagidxarr.value + frm.bagidx.value + ",";

            }
        }
    }

	//������ ���� ��ٱ��Ͽ��� ����
    frmbag.action='/common/item/adminshoppingbag_process.asp';
    frmbag.mode.value='baginsertdelarr';
    frmbag.target='view';
    frmbag.submit();

    //�ٹ����ٹ��� �ֹ��� �ۼ�������
    //����
    if (shopgubun == 'True'){
    	upfrm.action='/common/offshop/shop_jumuninput.asp';
	//����
	}else{
		upfrm.action='/admin/fran/jumuninput.asp';
	}
    upfrm.shopid.value=tmpshopid;
    upfrm.submit();
}

// ����ǰ �߰� �˾�
function jsAddNewItemOFF(upfrm,shopid ,acURL){
	var addnewItem;

		if (shopid == '') {
			alert('��ǰ�� �߰� �Ͻ� ������ �˻� �Ͻð�, ��ǰ�� �߰� �ϼ���');
			upfrm.shopid.focus();
			return;
		}

		addnewItem = window.open("/common/offshop/pop_itemAddInfoOFF.asp?shopid=" + shopid + "&acURL="+acURL, "addnewItemOFF", "width=1024,height=768,scrollbars=yes,resizable=yes");
		addnewItem.focus();
}

//��ü �ֹ��� �ۼ�
function AddArr_upche(upfrm,shopgubun){
    var frm; var tmpshopid = ''; var tmpmakerid = ''; var ret;
    var pass = false;
    var searchfrm = document.frm;

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    //if (searchfrm.makerid.value == ''){
    //    alert('�귣��(����ó)�� ������ �ּ���.');
    //    return;
    //}

    if (!pass) {
        alert('���� �������� �����ϴ�.');
        return;
    }

    upfrm.itemgubunarr2.value = "";
    upfrm.itemidadd2.value = "";
    upfrm.itemoptionarr2.value = "";
    upfrm.sellcasharr2.value = "";
    upfrm.suplycasharr2.value = "";
    upfrm.shopbuypricearr2.value = "";
    upfrm.itemnoarr2.value = "";
    upfrm.itemnamearr2.value = "";
    upfrm.itemoptionnamearr2.value = "";
    upfrm.designerarr2.value = "";
    upfrm.bagidxarr.value = "";

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){

				//��ü�ֹ��� ��� �ֹ��ڰ� �����ؾ���
                if (tmpshopid==''){
                	tmpshopid = frm.shopid.value;
                } else {
                	if (tmpshopid != frm.shopid.value){
	                    alert('���� Ʋ�� ������ ���õǾ� �ֽ��ϴ� \n����(�ֹ���)�� ���� �ؾ� �մϴ�.');
	                    return;
                	}
                }

				//��ü�ֹ��� ��� �ϳ��� �ֹ��� ����ó�� �Ѱ����� �Ѵ�
                if (tmpmakerid==''){
                	tmpmakerid = frm.makerid.value;
                } else {
                	if (tmpmakerid != frm.makerid.value){
	                    alert('���� Ʋ�� �귣��(����ó)�� ���õǾ� �ֽ��ϴ� \n��ü�ֹ��� ��� �귣��(����ó)�� �����ؾ� �մϴ�');
	                    return;
                	}
                }

				//��ü�ֹ��ǰ�� ��ü��Ź�� ��ü���Ը� �ֹ�����
                if (frm.comm_cd.value=="B011" || frm.comm_cd.value=="B031" || frm.comm_cd.value=="B013"){
                    alert('�ٹ�������Ź, ������, �����Ź�� �ֹ� �ϽǼ� �����ϴ�.');
                    frm.itemno.focus();
                    return;
                }

                if (!IsInteger(frm.itemno.value)){
                    alert('������ ������ �����մϴ�.');
                    frm.itemno.focus();
                    return;
                }

                if (frm.itemno.value=="0"){
                    alert('������ �Է��ϼ���.');
                    frm.itemno.focus();
                    return;
                }

                upfrm.itemgubunarr2.value = upfrm.itemgubunarr2.value + frm.itemgubun.value + "|";
                upfrm.itemidadd2.value = upfrm.itemidadd2.value + frm.itemid.value + "|";
                upfrm.itemoptionarr2.value = upfrm.itemoptionarr2.value + frm.itemoption.value + "|";
                upfrm.sellcasharr2.value = upfrm.sellcasharr2.value + frm.shopitemprice.value + "|";
                upfrm.suplycasharr2.value = upfrm.suplycasharr2.value + frm.shopsuplycash.value + "|";
                upfrm.shopbuypricearr2.value = upfrm.shopbuypricearr2.value + frm.shopbuyprice.value + "|";
                upfrm.itemnoarr2.value = upfrm.itemnoarr2.value + frm.itemno.value + "|";
                upfrm.itemnamearr2.value = upfrm.itemnamearr2.value + frm.itemname.value + "|";
                upfrm.itemoptionnamearr2.value = upfrm.itemoptionnamearr2.value + frm.itemoptionname.value + "|";
                upfrm.designerarr2.value = upfrm.designerarr2.value + frm.makerid.value + "|";
				frmbag.bagidxarr.value = frmbag.bagidxarr.value + frm.bagidx.value + ",";

            }
        }
    }

	//������ ���� ��ٱ��Ͽ��� ����
    frmbag.action='/common/item/adminshoppingbag_process.asp';
    frmbag.mode.value='baginsertdelarr';
    frmbag.target='view';
    frmbag.submit();

    //��ü �ֹ��� �ۼ� ������
    //����
    if (shopgubun == 'True'){
    	upfrm.action='/common/offshop/shop_ipchulinput.asp';
	//����
	}else{
		upfrm.action='/common/offshop/shop_ipchulinput.asp';
	}
    upfrm.shopid.value=tmpshopid;
    upfrm.chargeid.value=tmpmakerid;
    upfrm.submit();
}

//�귣��Ŭ����
function searchmakerid(makerid,upfrm){
    upfrm.makerid.value=makerid;
    upfrm.submit();
}

function CheckThis(frm){
    frm.cksel.checked=true;
    AnCheckClick(frm.cksel);
}

function addnewItem(onoffgubun,upfrm,shopid ,acURL){
	var addnewItem; var tmpshopid;

	tmpshopid = shopid;
	//tmpshopid = upfrm.shopid.value;

	if (onoffgubun=='ON'){

	}else if (onoffgubun=='OFF'){
		if (tmpshopid==''){
			alert('��ǰ�� �߰� �Ͻ� ������ �˻� �Ͻð�, ��ǰ�� �߰� �ϼ���');
			upfrm.shopid.focus();
			return;
		}

		addnewItem = window.open("/common/offshop/pop_itemAddInfo_off.asp?shopid="+tmpshopid+"&acURL="+acURL, "addnewItem", "width=1024,height=768,scrollbars=yes,resizable=yes");
		addnewItem.focus();
	}
}
