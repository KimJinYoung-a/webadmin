<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �귣��
' History : 2018.08.03 ������ ����
'			2022.02.24 �ѿ�� ����(�Ϲ�(����)�����, ��õ¡��, �ؿܻ���� üũ������ �����ϴ� ���� ����)
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
dim groupid, categorylarge

dim pcuserdiv : pcuserdiv = requestCheckVar(request("pcuserdiv"),16)
dim hp : hp = requestCheckVar(request("hp"),16)
dim email : email = requestCheckVar(request("email"),128)
dim cd1 : cd1 = requestCheckVar(request("cd1"),3)
dim cate1 : cate1 = requestCheckVar(request("cate1"),3)
dim companyno : companyno = requestCheckVar(request("companyno"),16)

if cate1 <> "" then
	categorylarge=cate1
else
	categorylarge=cd1
end if

'/2013.12.02 �ѿ�� �߰�
if not(C_ADMIN_AUTH or C_MngPart or C_partnership_part) then
	if pcuserdiv="999_50" or pcuserdiv="501_21" or pcuserdiv="502_21" or pcuserdiv="900_21" then
		response.write "<script language='javascript'>"
		response.write "	alert('[���Ѿ���] ����ó�� ��� ���� �մϴ�.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
end if
%>
<script language="JavaScript" src="/js/jquery-1.7.2.min.js"></script>
<script type='text/javascript'>

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

function precheck(frm){

    var pcuserdiv = getFieldValue(frm.pcuserdiv);
    // 9999_02:����ó, 9999_14:��ī����, 999_50:���޻�(�¶���) , 501_21:�������� 503_21:��Ÿ���� ,9999_21:���ó(��Ÿ)

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

	if(frm.uid.value.search(/\W|\s/g) > -1){
		alert("�귣�� ���̵𿡴� Ư������ �Ǵ� ������ �Է��� �� �����ϴ�.");
		va.focus();
	}

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

	if (pcuserdiv=="9999_02"){
		if (frm.jungsan_date.value=="" && frm.jungsan_date_off.value==""){
			alert('�������� ������ �ּ���.');
			return;
		}

		if (frm.catecode.value.length<1){
			alert('�¶��� ī�װ� ������ �����ϼ���.');
			//frm.catecode.focus();
			return;
		}

		if (frm.standardmdcatecode.value.length<1){
			alert('����ī�װ� ���MD�� �����ϼ���.');
			frm.standardmdcatecode.focus();
			return;
		}
		
		if (frm.mduserid.value.length<1){
			alert('���MD ������ �����ϼ���.');
			frm.mduserid.focus();
			return;
		}
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

    	if (frm.catecode.value.length<1){
    		alert('�¶��� ī�װ� ������ �����ϼ���.');
    		//frm.catecode.focus();
    		return;
    	}
		if (frm.standardmdcatecode.value.length<1){
			alert('����ī�װ� ���MD�� �����ϼ���.');
			frm.standardmdcatecode.focus();
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
            if (frm.defaultDeliverPay.value*1<2000){
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
	if ((pcuserdiv=="999_50")||(pcuserdiv=="900_21")||(pcuserdiv=="902_21")||(pcuserdiv=="503_21")){
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

	if ((pcuserdiv!="999_50")&&(pcuserdiv!="900_21")&&(pcuserdiv!="902_21")&&(pcuserdiv!="503_21")){
		//��Ʈ��Ʈ ǥ�ÿ��� ���޸��� �ٹ����ٰ� ����.
		if(frm.streetusing[0].checked){
			frm.extstreetusing.value = "Y";
		}else if(frm.streetusing[1].checked){
			frm.extstreetusing.value = "N";
		}
	}

	if (pcuserdiv=="9999_02" || pcuserdiv=="902_21" || pcuserdiv=="503_21"){
		if (frm.email.value==""){
			alert('��ü ����� �̸����� �Է��� �ּ���.');
			frm.email.focus();
			return;
		}

		if (frm.hp.value==""){
			alert('��ü ����� �ڵ��� ��ȣ�� �Է��� �ּ���.');
			frm.hp.focus();
			return;
		}
	}

	if (frm.signtype.value==""){
		alert('�ű� ��� ��� ������ �������ּ���.');
		frm.signtype.focus();
		return;
	}

	// �ؿܻ����
	if ( $("input[name=businessgubun]:radio:checked").val()=="5" ){
		if (frm.company_no.value==""){
			alert('�ؿ� ����ڹ�ȣ�� Ȯ�����ּ���.');
			frm.company_no.focus();
			return;
		}
		if (frm.partcheck.value==""){
			alert('�ؿ� ����ڹ�ȣ�� Ȯ�����ּ���.');
			frm.partcheck.focus();
			return;
		}
	// ��õ¡��
	}else if ( $("input[name=businessgubun]:radio:checked").val()=="3" ){
		if (frm.company_no.value==""){
			alert('�ֹε�Ϲ�ȣ�� �Է����ּ���.');
			frm.company_no.focus();
			return;
		}
		if (frm.partcheck.value==""){
			alert('�ֹε�Ϲ�ȣ�� �Է��ϰ� Ȯ�����ּ���.');
			frm.partcheck.focus();
			return;
		}
	// �Ϲ�(����)�����
	}else{
		if (frm.company_no.value==""){
			alert('����ڹ�ȣ�� �Է����ּ���.');
			frm.company_no.focus();
			return;
		}
		if (frm.partcheck.value==""){
			alert('����ڹ�ȣ�� �Է��ϰ� Ȯ�����ּ���.');
			frm.partcheck.focus();
			return;
		}
	}

	if (frm.partcheck.value==""){
		alert('����ڹ�ȣ�� �Է��ϰ� Ȯ�����ּ���.');
		frm.partcheck.focus();
		return;
    }

	var ret = confirm('�귣�� ������ ���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.target="FrameCKP";
		frm.action="doupchebrand.asp";
		frm.submit();
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

function PopUpcheSelectCustom(frmname){
	document.frmbrand.mode.value = "addnewupchebrand2";
	var popwin = window.open("/admin/member/popupcheselect.asp?mode=newbrand&frmname=" + frmname,"popupcheselect","width=800 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function viewtable(){
	document.getElementById("upchediv1").style.display = "";
	document.getElementById("upchediv2").style.display = "";
	document.getElementById("upchediv3").style.display = "";
	document.getElementById("upchediv4").style.display = "";
	document.getElementById("upchediv5").style.display = "";
	document.getElementById("upchediv6").style.display = "";
	document.getElementById("upchediv7").style.display = "";
	document.getElementById("upchediv8").style.display = "";
	document.getElementById("upchediv9").style.display = "";
	document.getElementById("upchediv10").style.display = "";
	document.getElementById("upchediv11").style.display = "";
	document.getElementById("upchediv12").style.display = "";
	document.getElementById("upchediv13").style.display = "";
	//document.getElementById("upchediv14").style.display = "";
	document.getElementById("upchediv15").style.display = "";
	document.getElementById("upchediv16").style.display = "";
	document.getElementById("upchediv17").style.display = "";
	document.getElementById("upchediv18").style.display = "";
	document.getElementById("upchediv19").style.display = "";
}

function DisableSocInfo(){
	var frm = document.frmbrand;
<% if (C_ADMIN_AUTH <> true) then %>
	frm.company_name.readOnly = true;
	frm.company_no.readOnly = true;
	frm.ceoname.readOnly = true;
	frm.jungsan_gubun.readOnly = true;

	$("input[name='company_name']").css("background-color","#EEEEEE");
	$("input[name='company_no']").css("background-color","#EEEEEE");
	$("input[name='ceoname']").css("background-color","#EEEEEE");
	$("input[name='jungsan_gubun']").css("background-color","#EEEEEE");
<% end if %>
}

function businessgubun_change(){
	var businessgubun = $("input[name=businessgubun]:radio:checked").val();
	var businessgubun3 = document.getElementById("businessgubun3");

	// �ؿܻ����
	if (businessgubun=="5"){
		document.getElementById("businessgubun1").style.display = "none";
		document.getElementById("businessgubun2").style.display = "none";
		document.getElementById("businessgubun3").style.display = "";

		// �ؿܴ� ����ڹ�ȣ �ڵ������ȴ�(888-00-00000)
		$("#company_no").val("888-00-00000");
		$("#company_no3").val("888-00-00000");
		//$("#company_no3").attr("readonly",true);

		//if (frm.coSearchBtn) {
		//	frm.coSearchBtn.disabled = true;
		//}

	// ��õ¡��
	} else if (businessgubun=="3"){
		document.getElementById("businessgubun1").style.display = "none";
		document.getElementById("businessgubun2").style.display = "";
		document.getElementById("businessgubun3").style.display = "none";

	// �Ϲ�(����)�����
	} else {
		document.getElementById("businessgubun1").style.display = "";
		document.getElementById("businessgubun2").style.display = "none";
		document.getElementById("businessgubun3").style.display = "none";
	}
}

function fnCheckUpcheNo(frm){
	var businessgubun = $("input[name=businessgubun]:radio:checked").val();

	// �ؿܻ����
	if (businessgubun=="5"){
		$("#company_no").val($("#company_no3").val())
		frm.target="FrameCKP";
		frm.action="checkUpcheSelect.asp";
		frm.submit();

	// ��õ¡��
	} else if (businessgubun=="3"){
		$("#company_no").val($("#company_no2").val())
		var company_no=$("#company_no").val().replace("-", "");

		if (!jsChkSocialNum1(company_no)){
			alert('�ֹε�Ϲ�ȣ�� �ٽ� �Է��� �ּ���.');
			return;
		}
		frm.target="FrameCKP";
		frm.action="checkUpcheSelect.asp";
		frm.submit();
		
	// �Ϲ�(����)�����
	} else {
		var bizNo = frm.company_no1.value;
		bizNo = bizNo.replace(/-/gi,"");
		if(frm.company_no1.value==""){
			alert("����ڹ�ȣ�� �Է����ּ���.");
		}
		else{
			var sumMod=0;
			sumMod += parseInt(bizNo.substring(0,1));
			sumMod += parseInt(bizNo.substring(1,2)) * 3 % 10;
			sumMod += parseInt(bizNo.substring(2,3)) * 7 % 10;
			sumMod += parseInt(bizNo.substring(3,4)) * 1 % 10;
			sumMod += parseInt(bizNo.substring(4,5)) * 3 % 10;
			sumMod += parseInt(bizNo.substring(5,6)) * 7 % 10;
			sumMod += parseInt(bizNo.substring(6,7)) * 1 % 10;
			sumMod += parseInt(bizNo.substring(7,8)) * 3 % 10;
			sumMod += Math.floor(parseInt(bizNo.substring(8,9)) * 5 / 10);
			sumMod += parseInt(bizNo.substring(8,9)) * 5 % 10;
			sumMod += parseInt(bizNo.substring(9,10));

			if(sumMod % 10 != 0){
				alert("����� ��Ϲ�ȣ�� �� �� �Ǿ����ϴ�.");
				return false;
			}else if ($("#company_no1").val().length != 12){
				alert('����� ��� ��ȣ�� 000-00-00000 �������� �Է��ؾ� �մϴ�..');
				$("#company_no1").focus();
				return;
			}else{
				$("#company_no").val($("#company_no1").val())
				frm.target="FrameCKP";
				frm.action="checkUpcheSelect.asp";
				frm.submit();
			}
		}
	}
}

function fnbizNoHyphen(num) {
     num = num.replace(/-/g, "");
     var num_str = num.toString();
     var result = '';
 
      for(var i=0; i<num_str.length; i++) {
            var tmp = num_str.length-(i+1);
            if(i==5){
				result = '-' + result;
			}
			else if(i==7){
				result = '-' + result;
			}
            result = num_str.charAt(tmp) + result;
       }
       return result;
}

function fnjuminNoHyphen(num) {
     num = num.replace(/-/g, "");
     var num_str = num.toString();
     var result = '';
 
      for(var i=0; i<num_str.length; i++) {
            var tmp = num_str.length-(i+1);
            if(i==7){
				result = '-' + result;
			}
            result = num_str.charAt(tmp) + result;
       }
       return result;
}

</script>
<% if (pcuserdiv="") then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmNext">
    <tr height="30" bgcolor="#FFFFFF">
    <td width="150" bgcolor="<%= adminColor("pink") %>">1. �귣�� ���� ����</td>
    <td >
        <% drawPartnerCommCodeBox false,"pcuserdiv","pcuserdiv","9999_02","" %>

        <script>delcomRow();</script>
    </td>
</tr>
<tr>
    <td colspan="2" height="30" bgcolor="#FFFFFF" align="center"><input type="button" value="����" onClick="stepNext();"></td>
</tr>
</form>
</table>
<% else %>
<form name="frmbrand" method="post" action="/admin/member/doupchebrand.asp" target="FrameCKP">
<input type="hidden" name="mode" value="addnewupchebrand">
<input type="hidden" name="partnerCnt" value="">
<input type="hidden" name="partcheck" value="">
<input type="hidden" name="defaultsongjangdiv" value="">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="30" >
    	<td bgcolor="<%= adminColor("pink") %>" colspan="6"><b>�귣���������</b></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td height="25" colspan="4">**�귣�� �⺻����**</td>
    </tr>
	<tr height="50">
		<td width="100" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> �귣��ID</td>
		<td bgcolor="#FFFFFF" >
    		<input type="text" class="text" name="uid" value="" size="24" maxlength="24">
    		<div>(����, ���ڸ� ���� Ư������ ����)</div>
		</td>
		<td width="100"  bgcolor="<%= adminColor("pink") %>"><font color=red>*</font>�귣�� ����</td>
		<td bgcolor="#FFFFFF" >
			<%= getPartnerCommCodeName("pcuserdiv",pcuserdiv) %>
            <input type="hidden" name="pcuserdiv" value="<%= pcuserdiv %>">
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
		<td width="100" bgcolor="<%= adminColor("pink") %>" >��������</td>
		<td bgcolor="#FFFFFF" colspan=2>
		<input type="hidden" name="selltype" value="0">
			<% drawPartnerCommCodeBox false,"purchasetype","purchasetype","1","" %>
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>">������</td>
		<td bgcolor="#FFFFFF" colspan=2>
			�¶��� : <% DrawJungsanDateCombo "jungsan_date", "" %>
			&nbsp;
			�������� : <% DrawJungsanDateCombo "jungsan_date_off", "" %>
		</td>
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font>�¶��δ�ǥ<br>����ī�װ�</td>
		<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "catecode", categorylarge %></td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font>�¶��� ���MD</td>
		<td bgcolor="#FFFFFF" colspan=2><% drawSelectBoxCoWorker_OnOff "mduserid", session("ssBctId") , "on" %></td>
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font>����ī�װ�<br>���MD</td>
		<td bgcolor="#FFFFFF" colspan=5><%= fnStandardDispCateSelectBox(1,"", "standardmdcatecode", "", "")%></td>
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
					<select class="select" name="defaultdeliverytype" onchange="inputDeliveryPay(this.value)">
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
	<tr>
		<td bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> ��ü ����� E-Mail</td>
		<td bgcolor="#FFFFFF" colspan="2">
			<input type="text" class="text" name="email" size="30" value="<%=email%>">
		</td>
		<td bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> ��ü ����� �ڵ���</td>
		<td bgcolor="#FFFFFF" colspan="2"><input type="text" class="text" name="hp"  value="<%=hp%>" size="15"></td>
	</tr>


</table>
<% end if %>

<% if (pcuserdiv="902_21") or (pcuserdiv="503_21") then %>
<input type="hidden" name="jungsan_date" value="">
<input type="hidden" name="jungsan_date_off" value="">
<input type="hidden" name="defaultmargine" value="20">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td width="100" bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> ��ü ����� E-Mail</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="email" size="30" value="<%=email%>">
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> ��ü ����� �ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="hp"  value="<%=hp%>" size="15"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">�⺻ ������</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<% drawPartnerCommCodeBox true,"selljungsantype","purchasetype","","" %>
		</td>
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

<% if (pcuserdiv="999_50") or (pcuserdiv="900_21") then %>
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
	<% if (pcuserdiv="999_50") then %>
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
	<% end if %>
	<tr height="40">
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" ���� ���� " onclick="precheck(frmbrand);"></td>
	</tr>
</table>

<% end if %>
<p>
	<input type="radio" name="signtype" value="1">�ű԰����(������)
	<input type="radio" name="signtype" value="2">�ű԰����(U+���ڼ���)
	<input type="radio" name="signtype" value="3">�ű԰����(DocuSign)
<p>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4"><b>��ü��������</b></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" colspan="4">
			<input type="hidden" id="company_no" name="company_no" value="">
			<input type="hidden" name="groupid" value="">
			<input type="radio" name="businessgubun" value="1" checked onclick="businessgubun_change();" >�Ϲ�(����)�����
			<input type="radio" name="businessgubun" value="3" onclick="businessgubun_change();" >��õ¡��
			<input type="radio" name="businessgubun" value="5" onclick="businessgubun_change();" >�ؿܻ����
			&nbsp;&nbsp;
			<span id="bizCheck"></span>
		</td>
	</tr>
	<tr id="businessgubun1" style="display:">
		<td bgcolor="<%= adminColor("tabletop") %>">����� ��ȣ</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="text" class="text" id="company_no1" name="company_no1" value="<%=companyno%>" size="15" onkeyup="this.value=fnbizNoHyphen(this.value)">
			<input type="button" class="button" value="Ȯ��" onClick="fnCheckUpcheNo(this.form);">&nbsp;
		</td>
	</tr>
	<tr id="businessgubun2" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">�ֹι�ȣ</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="text" class="text" id="company_no2" name="company_no2" value="<%=companyno%>" size="15" onkeyup="this.value=fnjuminNoHyphen(this.value)">
			<input type="button" class="button" value="Ȯ��" onClick="fnCheckUpcheNo(this.form);">&nbsp;
		</td>
	</tr>
	<tr id="businessgubun3" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">�ؿܻ���� ��ȣ</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="text" class="text" id="company_no3" name="company_no3" value="<%=companyno%>" size="15" >
			<input type="button" class="button" value="Ȯ��" onClick="fnCheckUpcheNo(this.form);">&nbsp;
			�űԾ�ü�� ��� �ؿܻ���� ��ȣ�� 888-00-00000 ���� �Է��� �ּ���. �ڵ����� �˴ϴ�.
		</td>
	</tr>
	<tr id="upchediv1" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">�����귣��ID</td>
		<td height="25" colspan="3" bgcolor="#FFFFFF"></td>
	</tr>

	<tr id="upchediv2" style="display:none">
		<td colspan="4" bgcolor="#FFFFFF" height="25">**����ڵ������**&nbsp;&nbsp;&nbsp;(�ߺ��� ����ڹ�ȣ�� ����� �� �����ϴ�.)</td>
	</tr>

	<tr id="upchediv3" style="display:none">
		<td width="120" bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ȸ���(��ȣ)</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_name" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
		<td width="90" bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ��ǥ��</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="ceoname" value="" size="16" maxlength="16" style="background-color:#EEEEEE;" readonly></td>
	</tr>
	<tr id="upchediv4" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ��������</td>
		<input type="hidden" name="checksocnoyn" value="N">
		<td bgcolor="#FFFFFF">
			<select name="jungsan_gubun" class="select" onchange="fnJungsanGubunChanged()" readonly>
			<option value="�Ϲݰ���" >�Ϲݰ���</option>
			<option value="���̰���" >���̰���</option>
			<option value="��õ¡��" >��õ¡��</option>
			<option value="�鼼" >�鼼</option>
			<option value="����(�ؿ�)" >����(�ؿ�)</option>
			</select>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font></td>
		<td bgcolor="#FFFFFF">			
		</td>
	</tr>
	<tr id="upchediv5" style="display:none">
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
	<tr id="upchediv6" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ����</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_uptae" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ����</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_upjong" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
	</tr>

	<tr id="upchediv7" style="display:none">
		<td colspan="4" bgcolor="#FFFFFF" height="25">**��ü�⺻����**</td>
	</tr>

	<tr id="upchediv8" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ��ǥ��ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.company_tel)"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ѽ�</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.company_fax)"></td>
	</tr>
	<tr id="upchediv9" style="display:none">
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
	<tr id="upchediv10" style="display:none">
		<td colspan="4" bgcolor="#FFFFFF" height="25">**������������**</td>
	</tr>

	<tr id="upchediv11" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">�ŷ�����</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		<% DrawBankCombo "jungsan_bank", "" %>
		</td>
	</tr>
	<tr id="upchediv12" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">���¹�ȣ</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" class="text" name="jungsan_acctno" value="" size="24" maxlength="32" style="background-color:#EEEEEE;" readonly>
		&nbsp;&nbsp; '-'�� ���� ��ȣ�� �Է����ֽñ� �ٶ��ϴ�.
		</td>
	</tr>
	<tr id="upchediv13" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">�����ָ�</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" class="text" name="jungsan_acctname" value="" size="24" maxlength="16" style="background-color:#EEEEEE;" readonly>
		&nbsp;&nbsp; ���� ���� ���ñ� �ٶ��ϴ�.
		</td>
	</tr>
	<tr id="upchediv15" style="display:none">
		<td colspan="4" bgcolor="#FFFFFF" height="25">**��ü ���������**</td>
	</tr>

	<tr id="upchediv16" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> ����ڸ�</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> �Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.manager_phone)"></td>
	</tr>
	<tr id="upchediv17" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> �ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.manager_hp)"></td>
	</tr>


	<tr id="upchediv18" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">�������ڸ�</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.jungsan_phone)"></td>
	</tr>
	<tr id="upchediv19" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.jungsan_hp)"></td>
	</tr>
	<tr height="40">
		<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" ���� ���� " onclick="precheck(frmbrand);"></td>
	</tr>
</table>

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