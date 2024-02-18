<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �귣��
' History : 2015.05.27 ������ ����
'			2022.02.09 �ѿ�� ����(����ī�װ� ���MD �߰�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%

dim i
dim designer
dim groupid

dim pcuserdiv : pcuserdiv = requestCheckVar(request("pcuserdiv"),16)

'/2013.12.02 �ѿ�� �߰�
if not(C_ADMIN_AUTH or C_AUTH) then
	if pcuserdiv="999_50" or pcuserdiv="501_21" or pcuserdiv="502_21" or pcuserdiv="503_21" or pcuserdiv="903_21" then	' 900_21 ���ó(��Ÿ)
		response.write "<script language='javascript'>"
		response.write "	alert('[���Ѿ���] ����ó�� ��� ���� �մϴ�.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
end if
%>
<script type='text/javascript'>

function CopyZip(flag,post1,post2,add,dong){
	if (flag=="s"){
		frmbrand.company_zipcode.value= post1 + "-" + post2;
		frmbrand.company_address.value= add;
		frmbrand.company_address2.value= dong;
	}else if(flag=="m"){
		frmbrand.return_zipcode.value= post1 + "-" + post2;
		frmbrand.return_address.value= add;
		frmbrand.return_address2.value= dong;
	}else if(flag=="p"){
		frmbrand.p_return_zipcode.value= post1 + "-" + post2;
		frmbrand.p_return_address.value= add;
		frmbrand.p_return_address2.value= dong;
	}
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function SameReturnAddr(bool){
    var frm = document.frmbrand;
	if (bool){
		frm.return_zipcode.value = frm.company_zipcode.value;
		frm.return_address.value = frm.company_address.value;
		frm.return_address2.value = frm.company_address2.value;
	}else{
		frm.return_zipcode.value = "";
		frm.return_address.value = "";
		frm.return_address2.value = "";
	}
}

function SameReturnAddr2(bool){
    var frm = document.frmbrand;
	if (bool){
		frm.p_return_zipcode.value = frm.return_zipcode.value;
		frm.p_return_address.value = frm.return_address.value;
		frm.p_return_address2.value = frm.return_address2.value;
	}else{
		frm.p_return_zipcode.value = "";
		frm.p_return_address.value = "";
		frm.p_return_address2.value = "";
	}
}

// �н����� ���⵵ �˻�
function fnChkComplexPassword(pwd) {
    var aAlpha = /[a-z]|[A-Z]/;
    var aNumber = /[0-9]/;
    var aSpecial = /[!|@|#|$|%|^|&|*|(|)|-|_]/;
    var sRst = true;

    if(pwd.length < 8){
        sRst=false;
        return sRst;
    }

    var numAlpha = 0;
    var numNums = 0;
    var numSpecials = 0;
    for(var i=0; i<pwd.length; i++){
        if(aAlpha.test(pwd.substr(i,1)))
            numAlpha++;
        else if(aNumber.test(pwd.substr(i,1)))
            numNums++;
        else if(aSpecial.test(pwd.substr(i,1)))
            numSpecials++;
    }

    if((numAlpha>0&&numNums>0)||(numAlpha>0&&numSpecials>0)||(numNums>0&&numSpecials>0)) {
    	sRst=true;
    } else {
    	sRst=false;
    }
    return sRst;
}

function precheck(frm){
	if (frm.company_name.value.length<1){
		alert('����� ��ϻ��� ȸ����� �Է��ϼ���.');
		frm.company_name.focus();
		return;
	}

	if (frm.ceoname.value.length<1){
		alert('����� ��ϻ��� ��ǥ�ڸ��� �Է��ϼ���.');
		frm.ceoname.focus();
		return;
	}

	if (frm.company_no.value.length<1){
		alert('����� ��� ��ȣ�� �Է��ϼ���.');
		frm.company_no.focus();
		return;
	}

	if (frm.jungsan_gubun.value.length<1){
		alert('���������� �����ϼ���.');
		frm.jungsan_gubun.focus();
		return;
	}

	var errMsg = chkIsValidJungsanGubun(frm.company_no.value, frm.jungsan_gubun.value);
	if (errMsg != "OK") {
		alert(errMsg);
		retutn;
	}

	if (frm.company_zipcode.value.length<1){
		alert('�����ȣ�� �����ϼ���.');
		frm.company_zipcode.focus();
		return;
	}

	if (frm.company_address.value.length<1){
		alert('����� ��ϻ��� �ּ�1�� �Է��ϼ���.');
		frm.company_address.focus();
		return;
	}

	if (frm.company_address2.value.length<1){
		alert('����� ��ϻ��� �ּ�2�� �Է��ϼ���.');
		frm.company_address2.focus();
		return;
	}

	if (frm.company_uptae.value.length<1){
		alert('����� ��ϻ��� ���¸� �Է��ϼ���.');
		frm.company_uptae.focus();
		return;
	}

	if (frm.company_upjong.value.length<1){
		alert('����� ��ϻ��� ������ �Է��ϼ���.');
		frm.company_upjong.focus();
		return;
	}

	if (frm.company_tel.value.length<1){
		alert('��ü ��ȭ��ȣ�� �Է��ϼ���.');
		frm.company_tel.focus();
		return;
	}

	if (frm.manager_name.value.length<1){
		alert('����� ������ �Է��ϼ���.');
		frm.manager_name.focus();
		return;
	}

	if (frm.manager_phone.value.length<1){
		alert('����� ��ȭ��ȣ�� �Է��ϼ���.');
		frm.manager_phone.focus();
		return;
	}

	if (frm.manager_email.value.length<1){
		alert('����� �̸����� �Է��ϼ���.');
		frm.manager_email.focus();
		return;
	}

	if (frm.manager_hp.value.length<1){
		alert('����� �ڵ����� �Է��ϼ���.');
		frm.manager_hp.focus();
		return;
	}

    if (frm.jungsan_date.value.length<1){
		alert('�������� �����ϼ���.');
		frm.jungsan_date.focus();
		return;
	}

    if (frm.jungsan_date_off.value.length<1){
		alert('���� �������� �����ϼ���. - �⺻�� �¶��ΰ� �����մϴ�.');
		frm.jungsan_date_off.focus();
		return;
	}

    var partnerCnt = frm.partnerCnt.value;
    if (partnerCnt=='') partnerCnt=0;

    // ���� �귣�尡 �ƴҰ��
    if (partnerCnt>0){
        if (frm.jungsan_date.value!=''){
            if (frm.jungsan_date.value!='����'){
                if (!confirm('�¶��� �������� �⺻���� ���� �Դϴ�. ��� ���� �Ͻðڽ��ϱ�?')){
                    return;
                }
            }
        }
        if (frm.jungsan_date_off.value!=''){
            if (frm.jungsan_date_off.value!='����'){
                if (!confirm('�������� �������� �⺻���� ���� �Դϴ�. ��� ���� �Ͻðڽ��ϱ�?')){
                    return;
                }
            }
        }
    }

// -----------------	//


/*
	if (frm.userdiv.value.length<1){
		alert('��ü ������ �����ϼ���.');
		frm.userdiv.focus();
		return;
	}
*/
    var pcuserdiv = getFieldValue(frm.pcuserdiv);
    // 9999_02:����ó, 9999_14:��ī����, 999_50:���޻�(�¶���) , 501_21:��������, 502_21:������, 503_21:���� ,9999_21:���ó(��Ÿ)

    if (pcuserdiv.length<1){
		alert('�귣�� ������ �����ϼ���.');
		frm.pcuserdiv[0].focus();
		return;
	}

	if (frm.uid.value.length<2){
		alert('�귣�� ���̵� �Է��ϼ���.');
		frm.uid.focus();
		return;
	}

    var regex = "^[a-zA-Z0-9_]+$";
	if(frm.uid.value.match(regex) == null){
		alert("�귣�� ���̵𿡴� ����, ����, ����(_) �� �Է��� �� �ֽ��ϴ�.");
		va.focus();
	}

	if (frm.password.value.length<1){
		alert('�귣�� �н����带 �Է��ϼ���.');
		frm.password.focus();
		return;
	}


	if (frm.password.value.length < 8 || frm.password.value.length > 16){
			alert("�н������ ������� 8~16���Դϴ�.");
			frm.password.focus();
			return ;
		 }

	if (!fnChkComplexPassword(frm.password.value)) {
			alert('�н������ ����/����/Ư������ �� �� ���� �̻��� �������� �Է����ּ���.');
			frm.password.focus();
			return;
		}

	if(frm.password.value==frm.uid.value) {
			alert("���̵�� �ٸ� ��й�ȣ�� ������ּ���.");
			frm.password.focus();
			return  ;
		}

	//if (frm.passwordS.value.length<1){
	//	alert('�귣�� 2�� �н����带 �Է��ϼ���.');
	//	frm.passwordS.focus();
	//	return;
	//}

	//if (frm.passwordS.value.length < 8 || frm.passwordS.value.length > 16){
	//	alert("2�� �н������ ������� 8~16���Դϴ�.");
	//	frm.passwordS.focus();
	//	return ;
	//}

	//if (!fnChkComplexPassword(frm.passwordS.value)) {
	//	alert('2�� �н������ ����/����/Ư������ �� �� ���� �̻��� �������� �Է����ּ���.');
	//	frm.passwordS.focus();
	//	return;
	//}

	//if(frm.passwordS.value==frm.uid.value) {
	//	alert("���̵�� �ٸ� ��й�ȣ�� ������ּ���.");
	//	frm.passwordS.focus();
	//	return  ;
	//}

	//if(frm.passwordS.value==frm.password.value) {
	//	alert("��й�ȣ��  �ٸ� ��й�ȣ�� ������ּ���.");
	//	frm.passwordS.focus();
	//	return  ;
	//}

	if (frm.socname_kor.value.length<1){
		alert('�귣���-�ѱ��� �Է��ϼ���.');
		frm.socname_kor.focus();
		return;
	}

	if (frm.socname.value.length<1){
		alert('�귣���-������ �Է��ϼ���.');
		frm.socname.focus();
		return;
	}

    if ((frm.p_return_zipcode.value.length<1)||(frm.p_return_address.value.length<1)){
        alert('���� ��ǰ�ּҸ� �Է��ϼ���.');
		frm.p_return_address.focus();
		return;
    }



    //�Ϲ� ����ó.
    if (pcuserdiv=="9999_02"){
        /*
        var selltype=getFieldValue(frm.selltype);
        if (selltype.length<1){
    		alert('�귣�� �Ǹ�ä���� �����ϼ���.');
    		frm.selltype[0].focus();
    		return;
    	}
	    */

	    if ((!frm.isusing[0].checked)&&(!frm.isusing[1].checked)){
    		alert('��뿩�θ� �����ϼ���.');
    		frm.isusing[0].focus();
    		return;
    	}

		/*
		// �׻� Y �� �����Ѵ�. ��� ���Ŀ� ��������(skyer9)
    	if ((!frm.isextusing[0].checked)&&(!frm.isextusing[1].checked)){
    		alert('���޸� ��뿩�θ� �����ϼ���.');
    		frm.isextusing[0].focus();
    		return;
    	}

        //���޻� �귣�� ���� confirm �ٹ�����Y,����N�ΰ�� (�������� N�� �����ϹǷ� ����)
        if ((frm.isusing[0].checked)&&(frm.isextusing[1].checked)){
            if (!confirm('���޸� �귣�� ��뿩�� N�ΰ�� InterPark,Lotte �� ���޸��� �Ǹ����� �ʽ��ϴ�. ����Ͻðڽ��ϱ�?')) {
                frm.isextusing[0].focus();
                return;
            }
        }
		*/

    	if ((!frm.streetusing[0].checked)&&(!frm.streetusing[1].checked)){
    		alert('��Ʈ��Ʈ ��뿩�θ� �����ϼ���.');
    		frm.streetusing[0].focus();
    		return;
    	}

		/*
    	if ((!frm.extstreetusing[0].checked)&&(!frm.extstreetusing[1].checked)){
    		alert('���޸� ��Ʈ��Ʈ ��뿩�θ� �����ϼ���.');
    		frm.extstreetusing[0].focus();
    		return;
    	}
    	*/

    	if ((!frm.specialbrand[0].checked)&&(!frm.specialbrand[1].checked)){
    		alert('Ŀ�´�Ƽ ��뿩�θ� �����ϼ���.');
    		frm.specialbrand[0].focus();
    		return;
    	}

    	if ((frm.catecode.value.length<1)&&(frm.offcatecode.value.length<1)){
    		alert('�¶��� �Ǵ� �������� ī�װ� ������ �����ϼ���. \n- �� �� �ϳ��� �ʼ� �����Դϴ�.');
    		//frm.catecode.focus();
    		return;
    	}

    	if (frm.standardmdcatecode.value.length<1){
    		alert('����ī�װ� ���MD�� �����ϼ���.');
    		frm.standardmdcatecode.focus();
    		return;
    	}

    	if ((frm.mduserid.value.length<1)&&(frm.offmduserid.value.length<1)){
    		alert('�¶��� �Ǵ� �������� ���MD ������ �����ϼ���. \n- �� �� �ϳ��� �ʼ� �����Դϴ�.');
    		frm.mduserid.focus();
    		return;
    	}

    	if (frm.maeipdiv.value.length<1){
    		alert('���� ������ �����ϼ���.');
    		frm.maeipdiv.focus();
    		return;
    	}

    	if (!IsDouble(frm.defaultmargine.value)){
    		alert('�⺻������ �Է��ϼ���. - �Ǽ��� �����մϴ�.');
    		frm.defaultmargine.focus();
    		return;
    	}

    	if(frm.defaultdeliverytype.options[1].selected == true){
    		if(frm.defaultFreeBeasongLimit.value == ""){
    			alert('���� ����� ��� �����۱��رݾ��� �Է����ּ���.');
    			frm.defaultFreeBeasongLimit.focus();
    			return;
    		}
    		if(frm.defaultDeliverPay.value == ""){
    			alert('���� ����� ��� ��ۺ� �Է����ּ���.');
    			frm.defaultDeliverPay.focus();
    			return;
    		}
    		if(isNaN(frm.defaultFreeBeasongLimit.value)){
    			alert('�ݾ��� ���ڷ� �Է����ּ���.');
    			frm.defaultFreeBeasongLimit.value = "";
    			frm.defaultFreeBeasongLimit.focus();
    			return;
    		}
    		if(isNaN(frm.defaultDeliverPay.value)){
    			alert('��ۺ�� ���ڷ� �Է����ּ���.');
    			frm.defaultDeliverPay.value = "";
    			frm.defaultDeliverPay.focus();
    			return;
    		}
            if (frm.defaultFreeBeasongLimit.value*1<=0){
                alert('���� ����� ��� �����۱��رݾ��� 0�� �̻��̾�� �մϴ�.');
                frm.defaultFreeBeasongLimit.focus();
                return;
            }
            if (frm.defaultDeliverPay.value*1<=2000){
                alert('���� ����� ��� ��ۺ�� 2000�� �̻� �Է°����Դϴ�.');
                frm.defaultDeliverPay.focus();
                return;
            }

    	}
	}

    //����ó(��ī����)
    if (pcuserdiv=="9999_14"){
        var selltype=frm.selltype.value;

        if (frm.mduserid.value.length<1){
    		alert('���MD ������ �����ϼ���.');
    		frm.mduserid.focus();
    		return;
    	}

        var lec_yn = getFieldValue(frm.lec_yn);
        var diy_yn = getFieldValue(frm.diy_yn);

        if ((lec_yn=="N")&&(diy_yn=="N")){
            alert('����/DIY ���� �ϳ��� ������� �����ϼž� �մϴ�.');
            frm.lec_yn[0].focus();
            return;
        }

        if ((lec_yn=="Y")&&(frm.lec_margin.value.length<1)){
            alert('���� �⺻ ������ �Է��ϼ���.');
            frm.lec_margin.focus();
            return;
        }

        if ((lec_yn=="Y")&&(frm.mat_margin.value.length<1)){
            alert('���� �⺻ ������ �Է��ϼ���.');
            frm.mat_margin.focus();
            return;
        }

        if ((diy_yn=="Y")&&(frm.diy_margin.value.length<1)){
            alert('DIY��ǰ �⺻ ������ �Է��ϼ���.');
            frm.diy_margin.focus();
            return;
        }

        if ((frm.diy_yn[0].checked)&&(frm.diy_dlv_gubun.value.length<1)){
            alert('DIY ��۱����� �����ϼ���.');
            frm.diy_dlv_gubun.focus();
            return;
        }

        if (frm.diy_dlv_gubun.value=="9"){
            if (!IsDigit(frm.DefaultFreebeasongLimit.value)){
                alert('��ۺ� ���� ���ڸ� �����մϴ�.');
                frm.DefaultFreebeasongLimit.focus();
                return;
            }

            if (!IsDigit(frm.DefaultDeliverPay.value)){
                alert('��ۺ�  ���ڸ� �����մϴ�.');
                frm.DefaultDeliverPay.focus();
                return;
            }

            if (frm.DefaultFreebeasongLimit.value*1<=0){
                alert('�ݾ��� 0�� �̻� �Է��ϼ���.');
                frm.DefaultFreebeasongLimit.focus();
                return;
            }

            if (frm.DefaultDeliverPay.value*1<=0){
                alert('�ݾ��� 0�� �̻� �Է��ϼ���.');
                frm.DefaultDeliverPay.focus();
                return;
            }

        }

        if ((lec_yn=="Y")&&(diy_yn=="N")){
            frm.selltype.value="10";
            frm.maeipdiv.value="M";
            frm.defaultmargine.value=frm.lec_margin.value;
        }

        if ((lec_yn=="N")&&(diy_yn=="Y")){
            frm.selltype.value="20";
            frm.maeipdiv.value="U";
            frm.defaultmargine.value=frm.diy_margin.value;

        }

        if ((lec_yn=="Y")&&(diy_yn=="Y")){
            frm.selltype.value="30";

        }

        if (diy_yn=="Y"){
            frm.maeipdiv.value="U";
            frm.defaultmargine.value=frm.diy_margin.value;
            frm.defaultdeliverytype.value = frm.diy_dlv_gubun.value;
        }
	}

	// �¶������޻�, etc���ó
	if ((pcuserdiv=="999_50") || (pcuserdiv=="900_21") || (pcuserdiv=="902_21") || (pcuserdiv=="903_21")){
	    if (frm.purchasetype.value.length<1){
	        alert('���� ����� ���� �ϼ���. �ʼ� ���Դϴ�.');
	        frm.purchasetype.focus()
	        return;
	    }
	}

	//�¶������޻�
	if (pcuserdiv=="999_50"){
	    if (frm.commission.value.length<1){
	        alert('�����Ḧ �Է� �ϼ���.');
	        frm.commission.focus()
	        return;
	    }
	}

	if ((pcuserdiv!="999_50") && (pcuserdiv!="900_21") && (pcuserdiv!="902_21") && (pcuserdiv!="903_21") && (pcuserdiv!="501_21") && (pcuserdiv!="502_21") && (pcuserdiv!="503_21")){
		//��Ʈ��Ʈ ǥ�ÿ��� ���޸��� �ٹ����ٰ� ����.
		if(frm.streetusing[0].checked){
			frm.extstreetusing.value = "Y";
		}else if(frm.streetusing[1].checked){
			frm.extstreetusing.value = "N";
		}
	}

	var ret = confirm('�귣�� ������ ���� �Ͻðڽ��ϱ�?');

	if (ret){
		if (frm.groupid.value.length<1) {
			icheckframe.location.href="icheckframe.asp?mode=CheckSocnoOnSave&socno=" + frm.company_no.value+"&pcuserdiv="+pcuserdiv;
		}else{
		    icheckframe.location.href="icheckframe.asp?uid=" + frm.uid.value + "&password=" + frm.password.value+"&pcuserdiv="+pcuserdiv;
		}
	}
}

function AddProc(mode){
    var frm = document.frmbrand;

	try {
		if (mode == "checkidpassword") {
			frm.submit();
			return;
		}

		if (mode == "CheckSocnoOnSave") {
			icheckframe.location.href="icheckframe.asp?uid=" + frm.uid.value + "&password=" + frm.password.value;
			return;
		}

		if (mode == "CheckSocno") {
			alert('��ϰ����� ����ڹ�ȣ�Դϴ�.');
			return;
		}
	} catch (err) {
		alert(err.message);
		return;
	}
}

function chkIsValidJungsanGubun(company_no, jungsan_gubun) {
	// 000-00-00000
	//
	// ��� �α��� : �����ڵ�
	// =========================================================================
	// 01-79 : ���λ����+���������
	// 90-99 : ���λ����+�鼼�����
	// ��Ÿ : ���� �鼼 ��� ����
	//
	// ���ڸ� ������ : ����(1-6) + �������Ϸù�ȣ
	// =========================================================================
	// 108 = 1(����) + 08(����)
	//
	// ���ڸ� 888 = ����(�ؿ�), ���̰���
	// =========================================================================

	if (company_no.length != 12) {
		// return "�߸��� ����ڹ�ȣ�Դϴ�.";
		return "OK";
	}

	var soc_gubun = company_no.substring(4, 6)*1;
	var IsForeign = (company_no.substring(0, 3) == "888");

	if (IsForeign) {
		if ((jungsan_gubun != "����(�ؿ�)") && (jungsan_gubun != "���̰���")) {
			return "����(�ؿ�), ���̰��� ����ڸ� ������ ����ڹ�ȣ�Դϴ�.";
		}

		return "OK";
	} else {
		if (jungsan_gubun == "����(�ؿ�)") {
			return "����(�ؿ�) ����ڷ� ���� �Ұ����� ����ڹ�ȣ�Դϴ�.";
		}

		/*
		if ((soc_gubun >= 1) && (soc_gubun <= 79)) {
			if (jungsan_gubun == "�鼼") {
				return "�鼼�� ����� �� ���� ����ڹ�ȣ�Դϴ�.";
			}

			return "OK";
		}
		*/

		if ((soc_gubun >= 90) && (soc_gubun <= 99)) {
			if (jungsan_gubun != "�鼼") {
				return "�鼼�θ� ��ϰ����� ����ڹ�ȣ�Դϴ�.";
			}

			return "OK";
		}

		return "OK";
	}
}

function SearchSocno(frm){

	if (frm.company_no.value.length<1){
		alert('����� ��� ��ȣ�� �Է��ϼ���.');
		frm.company_no.focus();
		return;
	}

	if (frm.company_no.value.length != 12){
		alert('����� ��� ��ȣ�� 000-00-00000 �������� �Է��ؾ� �մϴ�.');
		frm.company_no.focus();
		return;
	}

	if (frm.groupid.value.length<1){
		icheckframe.location.href="icheckframe.asp?mode=CheckSocno&socno=" + frm.company_no.value;
	}else{
		alert('����ڹ�ȣ�� ������ ��� ���� ������ ����˴ϴ�.');
	}

}

function ModiInfo(frm){
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		//frm.submit();
	}

}

function DisableSocInfo(frm){
<% if (C_ADMIN_AUTH <> true) then %>
	frm.company_name.readOnly = true;
	frm.company_no.readOnly = true;
	frm.ceoname.readOnly = true;
	frm.jungsan_gubun.readOnly = true;

	frm.company_name.style.background = "#EEEEEE";
	frm.company_no.style.background = "#EEEEEE";
	frm.ceoname.style.background = "#EEEEEE";
	frm.jungsan_gubun.style.background = "#EEEEEE";
<% end if %>
}

function CopyFromBrandInfo(){
	frmupche.company_name.value = frmbuf.company_name.value;
	frmupche.ceoname.value = frmbuf.ceoname.value;
	frmupche.company_no.value = frmbuf.company_no.value;
	frmupche.jungsan_gubun.value = frmbuf.jungsan_gubun.value;
	frmupche.company_zipcode.value = frmbuf.company_zipcode.value;
	frmupche.company_address.value = frmbuf.company_address.value;
	frmupche.company_address2.value = frmbuf.company_address2.value;
	frmupche.company_uptae.value = frmbuf.company_uptae.value;
	frmupche.company_upjong.value = frmbuf.company_upjong.value;
	frmupche.company_tel.value = frmbuf.company_tel.value;
	frmupche.company_fax.value = frmbuf.company_fax.value;
	frmupche.jungsan_bank.value = frmbuf.jungsan_bank.value;
	frmupche.jungsan_acctno.value = frmbuf.jungsan_acctno.value;
	frmupche.jungsan_acctname.value = frmbuf.jungsan_acctname.value;


	frmupche.manager_name.value = frmbuf.manager_name.value;
	frmupche.manager_phone.value = frmbuf.manager_phone.value;
	frmupche.manager_email.value = frmbuf.manager_email.value;
	frmupche.manager_hp.value = frmbuf.manager_hp.value;

	frmupche.deliver_name.value = frmbuf.deliver_name.value;
	frmupche.deliver_phone.value = frmbuf.deliver_phone.value;
	frmupche.deliver_email.value = frmbuf.deliver_email.value;
	frmupche.deliver_hp.value = frmbuf.deliver_hp.value;

	frmupche.jungsan_name.value = frmbuf.jungsan_name.value;
	frmupche.jungsan_phone.value = frmbuf.jungsan_phone.value;
	frmupche.jungsan_email.value = frmbuf.jungsan_email.value;
	frmupche.jungsan_hp.value = frmbuf.jungsan_hp.value;

}

function inputDeliveryType(ddt)
{
	if(ddt == "U")
	{
		document.getElementById("ddtdiv").style.display = "block";
	}
	else
	{
		document.frmbrand.defaultdeliverytype.options[0].selected = true;
		document.frmbrand.defaultFreeBeasongLimit.value = "";
		document.frmbrand.defaultDeliverPay.value = "";
		document.getElementById("ddtdiv").style.display = "none";
		document.getElementById("paydiv").style.display = "none";
	}
}

function inputDeliveryPay(pay)
{
	if(pay == "9")
	{
		document.getElementById("paydiv").style.display = "block";
	}
	else
	{
		document.frmbrand.defaultFreeBeasongLimit.value = "";
		document.frmbrand.defaultDeliverPay.value = "";
		document.getElementById("paydiv").style.display = "none";
	}
}



function clickLec(comp){

}

function clickDiy(comp){
    if (comp.value=="Y"){
        iDiyDlv.style.display="inline";
    }else{
        iDiyDlv.style.display="none";
    }
}

function stepNext(){
    var frm = document.frmNext;

    var pcuserdiv=getFieldValue(frm.pcuserdiv);

    if (pcuserdiv.length<1){
        alert('�귣�� ������ ���� �����ϼ���.');
        frm.pcuserdiv[0].focus();
        return;
    }

    frm.submit();
}

function chkCompdiygbn(comp){
    var frm = comp.form;
    if (comp.value=="9"){
        frm.DefaultFreebeasongLimit.style.background = '#FFFFFF';
        frm.DefaultDeliverPay.style.background  = '#FFFFFF';

        frm.DefaultFreebeasongLimit.readOnly = false;
        frm.DefaultDeliverPay.readOnly = false;

        frm.DefaultFreebeasongLimit.value=frm.pDFL.value;
        frm.DefaultDeliverPay.value=frm.pDDP.value;


    }else{
        frm.DefaultFreebeasongLimit.style.background = '#BBBBBB';
        frm.DefaultDeliverPay.style.background  = '#BBBBBB';

        frm.DefaultFreebeasongLimit.readOnly = true;
        frm.DefaultDeliverPay.readOnly = true;

        frm.DefaultFreebeasongLimit.value=0;
        frm.DefaultDeliverPay.value=0;
    }
}

function delcomRow(){
    //������ �̰����� ��� �Ұ�
    var f = document.frmNext;

    for (i=0;i<f.pcuserdiv.length;i++){
    	if (f.pcuserdiv[i].value=="501_21" || f.pcuserdiv[i].value=="502_21" || f.pcuserdiv[i].value=="503_21" ){
    		f.pcuserdiv.remove(i);
    		i--;
    	}
    }
}

var orgjungsan_gubun = "�Ϲݰ���";
function fnJungsanGubunChanged() {
	var frm = document.frmbrand;
	var company_no = document.getElementById("company_no");

	if ((orgjungsan_gubun != "����(�ؿ�)") && (frm.jungsan_gubun.value != "����(�ؿ�)")) {
		orgjungsan_gubun = frm.jungsan_gubun.value;
		return;
	}
	orgjungsan_gubun = frm.jungsan_gubun.value;

	if (frm.jungsan_gubun.value == "����(�ؿ�)") {
		// �ؿܴ� ����ڹ�ȣ �ڵ������ȴ�(888-00-00000)

		company_no.className = "text_ro";
		frm.company_no.readOnly = true;
		frm.company_no.value = "888-00-00000";

		frm.btnSearchSocno.disabled = true;

		frm.checksocnoyn.value = "Y";
	} else {
		company_no.className = "text";
		frm.company_no.readOnly = false;
		frm.company_no.value = "";

		frm.btnSearchSocno.disabled = false;

		frm.checksocnoyn.value = "N";
	}
}

</script>

<% if (pcuserdiv="") then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmNext">
    <tr height="30" bgcolor="#FFFFFF">
    <td width="150" bgcolor="<%= adminColor("pink") %>">1. �귣�� ���� ����</td>
    <td >
        <% drawPartnerCommCodeBox false,"pcuserdiv","pcuserdiv","9999_02","" %>

        <%'<script>delcomRow();</script>%>
    </td>
</tr>
<tr>
    <td colspan="2" height="30" bgcolor="#FFFFFF" align="center"><input type="button" value="����" onClick="stepNext();"></td>
</tr>
</form>
</table>
<% else %>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmbrand" method="post" action="/admin/member/doupcheedit.asp" target="FrameCKP">
<input type="hidden" name="mode" value="addnewupchebrand">
<input type="hidden" name="partnerCnt" value="">
	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4"><b>1.��ü��������</b></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ü�ڵ�</td>
		<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" class="text" name="groupid" value="" size="7" maxlength="5" style="background-color:#EEEEEE;" readonly>
		<input type="button" class="button" value="��ü����" onClick="PopUpcheSelect('frmbrand'); DisableSocInfo(frmbrand);">
		</td>
	</tr>
	<tr >
		<td bgcolor="<%= adminColor("tabletop") %>">�����귣��ID</td>
		<td height="25" colspan="3" bgcolor="#FFFFFF"></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**����ڵ������**&nbsp;&nbsp;&nbsp;(�ߺ��� ����ڹ�ȣ�� ����� �� �����ϴ�.)</td>
	</tr>

	<tr>
		<td width="120" bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ȸ���(��ȣ)</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_name" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
		<td width="90" bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ��ǥ��</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="ceoname" value="" size="16" maxlength="16" style="background-color:#EEEEEE;" readonly></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ����ڹ�ȣ</td>
		<input type="hidden" name="checksocnoyn" value="N">
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" id="company_no" name="company_no" value="" size="16" maxlength="20" style="background-color:#EEEEEE;" readonly>
			<!--<input type="button" class="button" name="btnSearchSocno" value="�˻�" onClick="SearchSocno(frmbrand)">//-->
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ��������</td>
		<td bgcolor="#FFFFFF">
			<select name="jungsan_gubun" class="select" onchange="fnJungsanGubunChanged()" readonly>
			<option value="�Ϲݰ���" >�Ϲݰ���</option>
			<option value="���̰���" >���̰���</option>
			<option value="��õ¡��" >��õ¡��</option>
			<option value="�鼼" >�鼼</option>
			<option value="����(�ؿ�)" >����(�ؿ�)</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ����������</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="company_zipcode" value="" size="7" maxlength="7" style="background-color:#EEEEEE;" readonly>
			<% '<input type="button" class="button_s" value="�˻�" onClick="FnFindZipNew('frmbrand','C')"> %>
			<% '<input type="button" class="button_s" value="�˻�(��)" onClick="TnFindZipNew('frmbrand','C')"> %>
			<% '<input type="button" class="button" value="�˻�(��)" onClick="popZip('s');"> %>
		    <br>
			<input type="text" class="text" name="company_address" value="" size="30" maxlength="64" style="background-color:#EEEEEE;" readonly>&nbsp;
			<input type="text" class="text" name="company_address2" value="" size="46" maxlength="64" style="background-color:#EEEEEE;" readonly>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ����</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_uptae" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ����</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_upjong" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**��ü�⺻����**</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ��ǥ��ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.company_tel)"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ѽ�</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.company_fax)"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�繫���ּ�</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="return_zipcode" value="" size="7" maxlength="7">
			<input type="button" class="button_s" value="�˻�" onClick="FnFindZipNew('frmbrand','D')">
			<input type="button" class="button_s" value="�˻�(��)" onClick="TnFindZipNew('frmbrand','D')">
			<% '<input type="button" class="button" value="�˻�(��)" onClick="popZip('m');"> %>

			<input type="checkbox" name="samezip2" onclick="SameReturnAddr(this.checked)">��
		<br>
		<input type="text" class="text" name="return_address" value="" size="30" maxlength="64">&nbsp;
		<input type="text" class="text" name="return_address2" value="" size="46" maxlength="64">
		</td>
	</tr>
	<!-- �귣�� ������ ����
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ �ּ�</td>
		<td colspan="3" height="25" bgcolor="#FFFFFF">�ʱ� ��ǰ�ּҴ� �繫�� �ּҿ� �����ϰ� �����˴ϴ�.</td>
	</tr>
	-->
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**������������**</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�ŷ�����</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		<% DrawBankCombo "jungsan_bank", "" %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">���¹�ȣ</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" class="text" name="jungsan_acctno" value="" size="24" maxlength="32" style="background-color:#EEEEEE;" readonly>
		&nbsp;&nbsp; '-'�� ���� ��ȣ�� �Է����ֽñ� �ٶ��ϴ�.
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�����ָ�</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" class="text" name="jungsan_acctname" value="" size="24" maxlength="16" style="background-color:#EEEEEE;" readonly>
		&nbsp;&nbsp; ���� ���� ���ñ� �ٶ��ϴ�.
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ������</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		�¶��� : <% DrawJungsanDateCombo "jungsan_date", "" %>
		&nbsp;
		�������� : <% DrawJungsanDateCombo "jungsan_date_off", "" %>
		</td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**��ü ���������**</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ����ڸ�</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> �Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.manager_phone)"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> �ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.manager_hp)"></td>
	</tr>


	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�������ڸ�</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.jungsan_phone)"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.jungsan_hp)"></td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="30" >
    	<td bgcolor="<%= adminColor("pink") %>" colspan="6"><b>2.�귣���������</b></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td height="25" colspan="4">**�귣�� �⺻����**</td>
    </tr>
    <tr height="30" bgcolor="#FFFFFF">
        <td width="100" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> �귣�� ����</td>
        <td colspan="3">
            <%= getPartnerCommCodeName("pcuserdiv",pcuserdiv) %>
            <input type="hidden" name="pcuserdiv" value="<%= pcuserdiv %>">
        </td>
    </tr>
	<tr height="50">
		<td width="100" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> �귣��ID</td>
		<td bgcolor="#FFFFFF" >
    		<input type="text" class="text" name="uid" value="" size="24" maxlength="24">
    		<p>(����, ����, ����(_) �� ���� Ư������ ����)</p>
			<% if (pcuserdiv="501_21") or (pcuserdiv="502_21") or (pcuserdiv="503_21") then %>
			<ul>
				<li>�ٹ����� ���� - streetshopxxx</li>
				<li>���̶�� ���� - ithinksoxxxxx, 3pl_its_xxxxx</li>
				<li>���� &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;- wholesale1xxx</li>
				<li>���� ���� &nbsp; &nbsp; &nbsp; - ygentshop1xxx, 3pl_xxx_xxxxx</li>
				<li>���̶�� �ؿ����ó &nbsp; &nbsp; &nbsp; - its_exp_xxxxx</li>
			</ul>
			<% end if %>
		</td>
		<td width="100"  bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> �н�����</td>
		<td bgcolor="#FFFFFF" >
		    <input type="password" class="text" name="password" value="" size="16" maxlength="24">
		    <%'<input type="password" class="text" name="passwordS" value="" size="16" maxlength="24">%>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> �귣���(KR)</td>
		<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="socname_kor" value="" size="30" maxlength="32">
		</td>
		<td bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> �귣���(EN)</td>
		<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="socname" value="" size="30" maxlength="32">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">�귣��� ǥ��</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<select name="socname_use" class="select">
				<option value="K">�귣���(KR)</option>
				<option value="E" selected>�귣���(EN)</option>
			</select>
		</td>
	</tr>
	<!--
	<tr >
		<td bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> ��ü����</td>
		<td bgcolor="#FFFFFF" colspan=2>
		<% DrawBrandGubunCombo "userdiv", "" %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> ī�װ�</td>
		<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "catecode", "" %></td>
	</tr>
	-->
	<tr bgcolor="#FFFFFF">
    	<td height="25" colspan="4">**�귣�� ��������**</td>
    </tr>
	<tr>
		<td height="25"  bgcolor="<%= adminColor("pink") %>">�⺻�ù��</td>
		<td bgcolor="#FFFFFF" ><% drawSelectBoxDeliverCompany "defaultsongjangdiv","" %></td>
		<td width="90" bgcolor="<%= adminColor("pink") %>" >����ȣ(����)</td>
		<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="prtidx" value="9999" size="4" maxlength="4">
		(�⺻�� : 9999)</td>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>" >����(��ǰ)�ּ�</td>
		<td bgcolor="#FFFFFF" colspan=5 >
			<input type="text" class="text" name="p_return_zipcode" value="" size="7" maxlength="7">
			<input type="button" class="button_s" value="�˻�" onClick="FnFindZipNew('frmbrand','I')">
			<input type="button" class="button_s" value="�˻�(��)" onClick="TnFindZipNew('frmbrand','I')">
			<% '<input type="button" class="button" value="�˻�(��)" onClick="popZip('p');"> %>

			<input type="checkbox" name="p_samezip" onclick="SameReturnAddr2(this.checked)">��(�繫���ּҿ�)
			<br>
			<input type="text" class="text" name="p_return_address" value="" size="30" maxlength="64">&nbsp;
			<input type="text" class="text" name="p_return_address2" value="" size="46" maxlength="64">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">��۴���ڸ�</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_name" value="" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("pink") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_phone" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.deliver_phone)"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_email" value="" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("pink") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_hp" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.deliver_hp)"></td>
	</tr>
</table>

<p>
<% ''' 9999_15 �߰� 2016/05/16 %>
<% if (pcuserdiv="9999_02") or (pcuserdiv="9999_15") then %>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
	    <% if (pcuserdiv="9999_15") then %>
	    <td height="25" colspan="6">**����ó(�ΰŽ���ǰ) �߰�����**</td>
	    <% else %>
		<td height="25" colspan="6">**����ó(�Ϲ�) �߰�����**</td>
	    <% end if %>

	</tr>
	<tr>
		<td width="100" bgcolor="<%= adminColor("pink") %>" ><!--�Ǹ�ä��--></td>
		<td bgcolor="#FFFFFF" colspan=2>
		<input type="hidden" name="selltype" value="0">
		<!--
		<input type="radio" name="selltype" value="0"> ��/OFF ��ü <input type="radio" name="selltype" value="9"> ������������
		-->
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>">��������</td>
		<td bgcolor="#FFFFFF" colspan=2>
			<% drawPartnerCommCodeBox false,"purchasetype","purchasetype","1","" %>
		</td>
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>">�¶��δ�ǥ<br>����ī�װ�</td>
		<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "catecode","" %></td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >�¶��� ���MD</td>
		<td bgcolor="#FFFFFF" colspan=2><% drawSelectBoxCoWorker_OnOff "mduserid", session("ssBctId") , "on" %></td>
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>">����ī�װ�<br>���MD</td>
		<td bgcolor="#FFFFFF" colspan=5><%= fnStandardDispCateSelectBox(1,"", "standardmdcatecode", "", "")%></td>
	</tr>
	<tr >
		<td bgcolor="<%= adminColor("pink") %>">�������� ī�װ�</td>
		<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "offcatecode", ""  %></td>
		<td bgcolor="<%= adminColor("pink") %>" >�������� ���MD</td>
		<td bgcolor="#FFFFFF" colspan=2><% drawSelectBoxCoWorker_OnOff "offmduserid", "" , "off" %></td>
	</tr>
	<tr>
		<td rowspan="3" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> �귣�� ��뿩��<br>&nbsp;&nbsp;(ī�װ�����)</td>
		<td bgcolor="#FFFFFF">�ٹ�����</td>
		<td bgcolor="#FFFFFF"><input type=radio name="isusing" value="Y" checked >��� <input type=radio name="isusing" value="N" >������</td>
		<td rowspan="3" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> ��Ʈ��Ʈ ǥ�ÿ���<br>&nbsp;&nbsp;(�귣������)</td>
		<td bgcolor="#FFFFFF">�ٹ�����</td>
		<td bgcolor="#FFFFFF"><input type=radio name="streetusing" value="Y" checked >��� <input type=radio name="streetusing" value="N" >������</td>
	</tr>
	<tr >
		<td bgcolor="#FFFFFF">���޸�</td>
		<td bgcolor="#FFFFFF">
			Y (����� ��������)
			<input type="hidden" name="isextusing" value="Y">

			<!-- ��Ʈ��Ʈ ǥ�ÿ��� ���޸� ����ó��. ��Ʈ��Ʈ ǥ�ÿ��� �ٹ����� �� �� ����. //-->
			<input type="hidden" name="extstreetusing" value="">
			<!--
			<input type=radio name="isextusing" value="Y" >��� <input type=radio name="isextusing" value="N" checked >������.

			<td bgcolor="#FFFFFF">���޸�</td>
			<td bgcolor="#FFFFFF"><input type=radio name="extstreetusing" value="Y" >��� <input type=radio name="extstreetusing" value="N" checked >������	</td>
			-->
		</td>
		<td bgcolor="#FFFFFF">Ŀ�´�Ƽ(��ǰQ/A)</td>
		<td bgcolor="#FFFFFF"><input type=radio name="specialbrand" value="Y" checked >��� <input type=radio name="specialbrand" value="N" >������</td>
	</tr>

	<tr >
		<td bgcolor="#FFFFFF" height="24">�ٹ����� OFF</td>
		<td bgcolor="#FFFFFF">
			N (����� ��������)
			<input type="hidden" name="isoffusing" value="N">
		</td>
		<td bgcolor="#FFFFFF"></td>
		<td bgcolor="#FFFFFF"></td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("pink") %>">Only ����</td>
		<td bgcolor="#FFFFFF" colspan="2"><input type=radio name="onlyflg" value="Y"  >Y <input type=radio name="onlyflg" value="N" checked >N</td>
		<td bgcolor="<%= adminColor("pink") %>">Artist ����</td>
		<td bgcolor="#FFFFFF" colspan="2"><input type=radio name="artistflg" value="Y"  >Y <input type=radio name="artistflg" value="N" checked >N</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">K-Design ����</td>
		<td bgcolor="#FFFFFF" colspan="2"><input type=radio name="kdesignflg" value="Y"  >Y <input type=radio name="kdesignflg" value="N" checked >N</td>
		<td bgcolor="#FFFFFF" colspan="3"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**�����û���**</td>
	</tr>
	<!--
	<tr bgcolor="#FFFFFF">
		<td colspan=6>
		* �¶��� �����ϰ�� -> �������ε� �������� ����.<br>
		* �¶��� ��Ź�ϰ�� -> �������ε� ��Ź���� ����.
		</td>
	</td>
	-->
	<tr>
		<td bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> �귣�� �⺻����</td>
		<td bgcolor="#FFFFFF" colspan=5>
			<table cellpadding="1" cellspacing="1" border="0" class="a">
			<tr>
				<td>
					<% DrawBrandMWUCombo_2011 "maeipdiv","" %>
					<input type="text" class="text" name="defaultmargine" value="" size="4" style="text-align:right"> %
				</td>
			</tr>
			<tr id="ddtdiv" style="display:none;">
				<td>
					��ü������Ǽ���:
					<select class='select' name="defaultdeliverytype" onchange="inputDeliveryPay(this.value)">
						<option value="null" selected>��ü������</option>
						<option value="9">��ü���ǹ��</option>
						<option value="7">��ü���ҹ��</option>
					</select>
				</td>
			</tr>
			<tr id="paydiv" style="display:none;">
				<td>
					<input type="text" name="defaultFreeBeasongLimit" value="" size="7" maxlength="7">�� �̸� ���Ž� ��۷� <input type="text" name="defaultDeliverPay" value="" size="7" maxlength="7">��
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<!--
	<tr>
		<td bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> �귣�� �⺻����</td>
		<td bgcolor="#FFFFFF" colspan="5">
			���� <input type="text" class="text" name="" value="" size="4"> %  /
			��Ź <input type="text" class="text" name="" value="" size="4"> %  /
			��ü��� <input type="text" class="text" name="" value="" size="4"> %
			(����, �� ���ذ����� ���濹��)
		</td>
	</tr>
    -->
	<tr height="40">
		<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" ���� ���� " onclick="precheck(frmbrand);"></td>
	</tr>

</table>
<% end if %>

<% if (pcuserdiv="9999_14") then %>
<!-- ��ī���� -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**����ó(��ī����) �߰�����**
		<input type="hidden" name="selltype" value="10">        <!-- �ڵ����� -->
		<input type="hidden" name="purchasetype" value="0">
		<input type="hidden" name="catecode" value="999">       <!-- ���þ��� -->
		<input type="hidden" name="offcatecode" value="999">    <!-- ���þ��� -->
		<input type="hidden" name="offmduserid" value="">       <!-- OFFMD ���� -->

		<input type="hidden" name="isextusing" value="N">       <!-- ���޸� ������ -->
		<input type="hidden" name="extstreetusing" value="N">   <!-- ���޸� Street ������ -->
		<input type="hidden" name="isoffusing" value="N">       <!-- OFF ������ -->
		<input type="hidden" name="specialbrand" value="N">     <!-- specialbrand Ŀ�´�Ƽ ������ -->
		<input type="hidden" name="onlyflg" value="N">          <!-- onlyflg ������ -->
		<input type="hidden" name="artistflg" value="N">        <!-- artistflg ������ -->
		<input type="hidden" name="kdesignflg" value="N">       <!-- kdesignflg ������ -->

		<input type="hidden" name="maeipdiv" value="M">         <!-- ���Ա��� -->
		<input type="hidden" name="defaultmargine" value="50">  <!-- �⺻����(����) -->
		<input type="hidden" name="defaultdeliverytype" value="">

		</td>
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>"></td>
		<td bgcolor="#FFFFFF" ></td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >���MD</td>
		<td bgcolor="#FFFFFF"><% drawSelectBoxCoWorker_OnOff "mduserid", "" , "fingers" %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">�귣��<br>��뿩��</td>
		<td bgcolor="#FFFFFF">
		    <input type=radio name="isusing" value="Y" checked >Y
		    <input type=radio name="isusing" value="N" >N</td>
		<td bgcolor="<%= adminColor("pink") %>">��Ʈ��Ʈ<br>ǥ�ÿ���<br>(�귣������)</td>
		<td bgcolor="#FFFFFF">
		    <input type=radio name="streetusing" value="Y" checked >Y
		    <input type=radio name="streetusing" value="N" >N</td>
	</tr>

	<tr >
		<td width="120" bgcolor="#DDDDFF" rowspan="2">���� ���� ����</td>
		<td bgcolor="#FFFFFF" rowspan="2">
		<input type="radio" name="lec_yn" value="Y" checked onClick="clickLec(this)"> Y
		<input type="radio" name="lec_yn" value="N" onClick="clickLec(this)"> N
		</td>
		<td width="120" bgcolor="#DDDDFF">���±⺻����</td>
		<td bgcolor="#FFFFFF">
		<input type="text" name="lec_margin" value="" size="4" maxlength="3"> (%)
		</td>
	</tr>
	<tr>
	    <td width="120" bgcolor="#DDDDFF">���⺻����</td>
		<td bgcolor="#FFFFFF" width="200">
		<input type="text" name="mat_margin" value="" size="4" maxlength="3"> (%)
		</td>
	</tr>
	<tr >
		<td width="120" bgcolor="#DDDDFF" >DIY ���� ����</td>
		<td  bgcolor="#FFFFFF" width="200" >
		<input type="radio" name="diy_yn" value="Y"  onClick="clickDiy(this);"> Y
		<input type="radio" name="diy_yn" value="N"  checked  onClick="clickDiy(this);"> N
		</td>
		<td width="120" bgcolor="#DDDDFF">�⺻����</td>
		<td bgcolor="#FFFFFF" width="200">
		<input type="text" name="diy_margin" value="" size="4" maxlength="3"> (%)
		</td>
	</tr>

	<tr id="iDiyDlv" style="display:none">
		<td width="120" bgcolor="#DDDDFF">DIY��۱���</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<select name="diy_dlv_gubun" onChange="chkCompdiygbn(this);">
		<option value="0" >�⺻(��ü������)
		<option value="9" selected >��ü ���ǹ��
		</select>
		<br>
		<input type="hidden" name="pDFL" value="">
		<input type="hidden" name="pDDP" value="">
		<input type="text" name="DefaultFreebeasongLimit" value="" size="9" maxlength="9">�� �̻� ������
		/�̸� ��ۺ� <input type="text" name="DefaultDeliverPay" value="" size="9" maxlength="9">��
		</td>
	</tr>
	<tr height="40">
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" ���� ���� " onclick="precheck(frmbrand);"></td>
	</tr>
</table>
<% end if %>

<%
	'///// �߰����� - 999(���޻�) /////
	if (pcuserdiv="999_50") then
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**<%=getPartnerCommCodeName("pcuserdiv",pcuserdiv)%> �߰�����**</td>
		<input type="hidden" name="catecode" value="999"> <!-- ���þ��� -->
		<input type="hidden" name="offcatecode" value="999"> <!-- ���þ��� -->

		<input type="hidden" name="isextusing" value="N"> <!-- ���޸� ������ -->
		<input type="hidden" name="streetusing" value="N"> <!--  Street ������ -->
		<input type="hidden" name="extstreetusing" value="N"> <!-- ���޸� Street ������ -->
		<input type="hidden" name="isoffusing" value="N">
		<input type="hidden" name="specialbrand" value="N"> <!-- specialbrand Ŀ�´�Ƽ ������ -->
		<input type="hidden" name="onlyflg" value="N"> <!-- onlyflg ������ -->
		<input type="hidden" name="artistflg" value="N"> <!-- artistflg ������ -->
		<input type="hidden" name="kdesignflg" value="N"> <!-- kdesignflg ������ -->

		<input type="hidden" name="maeipdiv" value="M">   <!-- ���Ա��� -->
		<input type="hidden" name="defaultmargine" value="20">  <!-- �⺻����(����) -->

		<input type="hidden" name="M_margin" value="0">
		<input type="hidden" name="W_margin" value="0">
		<input type="hidden" name="U_margin" value="0">
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>">�귣�� ��뿩��</td>
		<td bgcolor="#FFFFFF" width="200">
		    <input type=radio name="isusing" value="Y" checked >Y
		    <input type=radio name="isusing" value="N" >N
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >�����(����)</td>
		<td bgcolor="#FFFFFF"><% drawSelectBoxCoWorker_OnOff "mduserid", "", "sell" %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">�⺻ �������</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"sellacccd","selltype","","" %>

		</td>
		<td bgcolor="<%= adminColor("pink") %>">�⺻ ������</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"selljungsantype","purchasetype","","" %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">�⺻ ����μ�</td>
		<td bgcolor="#FFFFFF">
		   <%= fndrawSaleBizSecCombo(true,"sellBizCd","","") %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">������</td>
		<td bgcolor="#FFFFFF">
		    <input type="text" name="commission" value="" size="4">%
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">��꼭������</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"taxevaltype","taxevaltype","","" %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">(��Ÿ����)������</td>
		<td bgcolor="#FFFFFF">
        <% drawPartnerCommCodeBox true,"etcjungsantype","etcjungsantype","","" %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**<%=getPartnerCommCodeName("pcuserdiv",pcuserdiv)%> ���޻� ���� ����**</td>

	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">��������</td>
		<td bgcolor="#FFFFFF">
		   <% drawPartnerCommCodeBox true,"mallSellType","pmallSellType","","" %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">�������</td>
		<td bgcolor="#FFFFFF">
		    <% drawPartnerCommCodeBox true,"pcomType","pcomType","","" %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">���޾���URL</td>
		<td bgcolor="#FFFFFF" colspan="3">
		   <input type="text" name="padminUrl" value="" size="60" maxlength="120">
		</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("pink") %>">���޾��ΰ���</td>
		<td bgcolor="#FFFFFF" colspan="">
		   ID <input type="text" name="padminId" value="" size="10" maxlength="32">
		   PW <input type="password" name="padminPwd" value="" size="10" maxlength="32">
		</td>
		<td bgcolor="<%= adminColor("pink") %>">�ֹ�ó�����</td>
		<td bgcolor="#FFFFFF">
            <% drawSelectBoxCoWorker_OnOff "offmduserid", "", "sell" %>
		</td>
	</tr>
    <tr>
		<td bgcolor="<%= adminColor("pink") %>">�⺻ ��ۺ� ����</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="hidden" name="defaultdeliverytype" value="9">
			<input type="text" name="defaultFreeBeasongLimit" value="" size="8" maxlength="7">�� �̸� ���Ž�
			��۷� <input type="text" name="defaultDeliverPay" value="" size="7" maxlength="7">��
		</td>
	</tr>
	<tr height="40">
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" ���� ���� " onclick="precheck(frmbrand);"></td>
	</tr>
</table>
<% end if %>


<%
	'///// �߰����� - 900(���ó), 902(���¾�ü), 903(3PL��ǥ) /////
	if (pcuserdiv="900_21") or (pcuserdiv="902_21") or (pcuserdiv="903_21") then
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**<%=getPartnerCommCodeName("pcuserdiv",pcuserdiv)%> �߰�����**</td>
		<input type="hidden" name="catecode" value="999"> <!-- ���þ��� -->
		<input type="hidden" name="offcatecode" value="999"> <!-- ���þ��� -->

		<input type="hidden" name="isextusing" value="N"> <!-- ���޸� ������ -->
		<input type="hidden" name="streetusing" value="N"> <!--  Street ������ -->
		<input type="hidden" name="extstreetusing" value="N"> <!-- ���޸� Street ������ -->
		<input type="hidden" name="isoffusing" value="N">
		<input type="hidden" name="specialbrand" value="N"> <!-- specialbrand Ŀ�´�Ƽ ������ -->
		<input type="hidden" name="onlyflg" value="N"> <!-- onlyflg ������ -->
		<input type="hidden" name="artistflg" value="N"> <!-- artistflg ������ -->
		<input type="hidden" name="kdesignflg" value="N"> <!-- kdesignflg ������ -->

		<input type="hidden" name="maeipdiv" value="M">   <!-- ���Ա��� -->
		<input type="hidden" name="defaultmargine" value="20">  <!-- �⺻����(����) -->

		<input type="hidden" name="M_margin" value="0">
		<input type="hidden" name="W_margin" value="0">
		<input type="hidden" name="U_margin" value="0">
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>">�귣�� ��뿩��</td>
		<td bgcolor="#FFFFFF" width="200">
		    <input type=radio name="isusing" value="Y" checked >Y
		    <input type=radio name="isusing" value="N" >N
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >�����(����)</td>
		<td bgcolor="#FFFFFF"><% drawSelectBoxCoWorker_OnOff "mduserid", "", "sell" %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">�⺻ ������</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<% drawPartnerCommCodeBox true,"selljungsantype","purchasetype","","" %>
		</td>
	</tr>
	<tr height="40">
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" ���� ���� " onclick="precheck(frmbrand);"></td>
	</tr>
</table>
<% end if %>


<%
	'///// �߰����� - 501(������), 502(������), 503(����ó) /////
	if (pcuserdiv="501_21") or (pcuserdiv="502_21") or (pcuserdiv="503_21") then
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**<%=getPartnerCommCodeName("pcuserdiv",pcuserdiv)%> �߰�����**</td>
		<input type="hidden" name="catecode" value="999"> <!-- ���þ��� -->
		<input type="hidden" name="offcatecode" value="999"> <!-- ���þ��� -->

		<input type="hidden" name="isextusing" value="N"> <!-- ���޸� ������ -->
		<input type="hidden" name="streetusing" value="N"> <!--  Street ������ -->
		<input type="hidden" name="extstreetusing" value="N"> <!-- ���޸� Street ������ -->
		<input type="hidden" name="isoffusing" value="N">
		<input type="hidden" name="specialbrand" value="N"> <!-- specialbrand Ŀ�´�Ƽ ������ -->
		<input type="hidden" name="onlyflg" value="N"> <!-- onlyflg ������ -->
		<input type="hidden" name="artistflg" value="N"> <!-- artistflg ������ -->
		<input type="hidden" name="kdesignflg" value="N"> <!-- kdesignflg ������ -->

		<input type="hidden" name="maeipdiv" value="M">   <!-- ���Ա��� -->
		<input type="hidden" name="defaultmargine" value="20">  <!-- �⺻����(����) -->

		<input type="hidden" name="M_margin" value="0">
		<input type="hidden" name="W_margin" value="0">
		<input type="hidden" name="U_margin" value="0">
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>">�귣�� ��뿩��</td>
		<td bgcolor="#FFFFFF" width="200">
		    <input type=radio name="isusing" value="Y" checked >Y
		    <input type=radio name="isusing" value="N" >N
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >�����(����)</td>
		<td bgcolor="#FFFFFF"><% drawSelectBoxCoWorker_OnOff "mduserid", "", "sell" %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">�⺻ �������</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"sellacccd","selltype","","" %>

		</td>
		<td bgcolor="<%= adminColor("pink") %>">�⺻ ������</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"selljungsantype","purchasetype","0","" %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">�⺻ ����μ�</td>
		<td bgcolor="#FFFFFF">
		   <%= fndrawSaleBizSecCombo(true,"sellBizCd","","") %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">������</td>
		<td bgcolor="#FFFFFF">
		    <input type="text" name="commission" value="" size="4">%
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">��꼭������</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"taxevaltype","taxevaltype","","" %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">(��Ÿ����)������</td>
		<td bgcolor="#FFFFFF">
        <% drawPartnerCommCodeBox true,"etcjungsantype","etcjungsantype","","" %>
		</td>
	</tr>
	<tr height="40">
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" ���� ���� " onclick="precheck(frmbrand);"></td>
	</tr>
</table>
<% end if %>

<% if (pcuserdiv="") then %>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="30"><td align="center">[���� �귣�� ������ ���� �ϼ���.]</td></tr>
</table>
<% end if %>

</form>

<!--
!!!!!! icheckframe.asp �ι� ������. (AddProc() ����)
-->
<iframe src="" name="icheckframe" width="200" height="0" frameborder="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>
<iframe name="FrameCKP" src="" frameborder="0" width="600"  height="400"></iframe>
<% end if %>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
