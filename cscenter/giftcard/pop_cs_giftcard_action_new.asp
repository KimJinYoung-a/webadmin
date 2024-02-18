<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenter/lib/CSFunction.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%

'[�ڵ�����]
'------------------------------------------------------------------------------
'A008			�ֹ����
'
'[��������]
'------------------------------------------------------------------------------
'CSFunction.asp
'
'dim IsStatusRegister			'����
'dim IsStatusEdit				'����
'dim IsStatusFinishing			'ó���Ϸ� �õ�
'dim IsStatusFinished			'ó���Ϸ�

'dim IsDisplayPreviousCSList	'���� CS ����
'dim IsDisplayCSMaster			'CS ����������
'dim IsDisplayItemList			'��ǰ���
'dim IsDisplayRefundInfo		'ȯ������
'dim IsDisplayButton			'��ư
'
'dim IsPossibleModifyCSMaster
'dim IsPossibleModifyItemList
'dim IsPossibleModifyRefundInfo



dim i, id, mode, divcd, giftorderserial, ckAll, iPgGubun

id			= request("id")
divcd		= request("divcd")
giftorderserial	= request("giftorderserial")
mode		= request("mode")
ckAll		= request("ckAll")

dim IsOrderCanceled
dim OrderMasterState
dim IsTicketOrder



'==============================================================================
'CS���������� ��������
dim ocsaslist

set ocsaslist = New CCSASList

ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if



'==============================================================================
'CS���������� ������ ������� �ű� ����
if (ocsaslist.FResultCount<1) then
	set ocsaslist.FOneItem = new CCSASMasterItem

	ocsaslist.FOneItem.FId = 0
	ocsaslist.FOneItem.Fdivcd = divcd

	mode = "regcsas"
else
    divcd       = ocsaslist.FOneItem.Fdivcd
    giftorderserial = ocsaslist.FOneItem.Forderserial

    if (ocsaslist.FOneItem.FCurrState = "B007") then
		mode = "finished"
    else
    	if (mode = "finishreginfo") then
    		'
    	else
    		mode = "editreginfo"
    	end if
    end if
end if

Call SetCSVariable(mode, divcd)



'==============================================================================
''ȯ������
dim orefund

set orefund = New CCSASList

orefund.FRectCsAsID = ocsaslist.FOneItem.FId

orefund.GetOneRefundInfo

if (orefund.FOneItem.Fencmethod = "TBT") then
	orefund.FOneItem.Frebankaccount = Decrypt(orefund.FOneItem.FencAccount)
elseif (orefund.FOneItem.Fencmethod = "PH1") then
	orefund.FOneItem.Frebankaccount = orefund.FOneItem.Fdecaccount
end if

function Decrypt(encstr)
	if (Not IsNull(encstr)) and (encstr <> "") then
		Decrypt = TBTDecrypt(encstr)
		exit function
	end if
	Decrypt = ""
end function



'==============================================================================
''�ֹ� ����Ÿ
dim ogiftcardordermaster

set ogiftcardordermaster = new cGiftCardOrder

ogiftcardordermaster.FRectgiftorderserial = giftorderserial

if Left(giftorderserial,1)<>"G" then
	response.write "�߸��� �����Դϴ�."
	response.end
else
    ogiftcardordermaster.getCSGiftcardOrderDetail
end if

IsOrderCanceled = (ogiftcardordermaster.FOneItem.Fcancelyn = "Y")
OrderMasterState = ogiftcardordermaster.FOneItem.FIpkumDiv
'iPgGubun = (ogiftcardordermaster.FOneItem.Fpggubun)    ' ����Ʈī��� pg���� ����? ���������� ��� ������ ���� �켱 ������ �߰��ص�.

'=============================================================================='==============================================================================
'���ֹ� ��ǰ�ݾ�
dim orgitemcostsum, orgpercentcouponpricesum

'������ǰ �հ�ݾ�
dim regitemcostsum, regpercentcouponpricesum



'==============================================================================
''���� �Ұ��� �޼���
dim JupsuInValidMsg

if (Left(giftorderserial,1)<>"A") and (ogiftcardordermaster.FResultCount<1) then
    response.write "<br><br>!!! ���� �ֹ������̰ų� �ֹ� ������ �����ϴ�. - ������ ���� ���"
    dbget.close()	:	response.End
end if

''���� ���� ����
dim IsJupsuProcessAvail

if (ogiftcardordermaster.FResultCount>0) then
	if (ogiftcardordermaster.FOneItem.FCancelyn <> "N") then
		JupsuInValidMsg = "���� �ֹ��Ǹ� ��� �����մϴ�."
		IsJupsuProcessAvail = false
	else
		JupsuInValidMsg = ""
		IsJupsuProcessAvail = true
	end if

	if (ogiftcardordermaster.FOneItem.Fjumundiv = "7") then
		JupsuInValidMsg = "��ϵ� Giftī���ֹ��� ����� �� �����ϴ�. ������� ���·� ��ȯ�ϼ���." & ogiftcardordermaster.FOneItem.Fjumundiv
		IsJupsuProcessAvail = false
	end if
else
    IsJupsuProcessAvail = false
end if

if (ogiftcardordermaster.FOneItem.Fsubtotalprice < 0) and (IsJupsuProcessAvail = true) then
	IsJupsuProcessAvail = false
	JupsuInValidMsg = "���̳ʽ��ֹ��� ���� CS������ �� �����ϴ�."
end if

%>

<script language='javascript' SRC="/js/ajax.js"></script>
<script language='javascript'>
var IsCsPowerUser               = <%= LCase(C_CSPowerUser) %>;

var IsStatusRegister 			= <%= LCase(IsStatusRegister) %>;
var IsStatusEdit 				= <%= LCase(IsStatusEdit) %>;
var IsStatusFinishing 			= <%= LCase(IsStatusFinishing) %>;
var IsStatusFinished 			= <%= LCase(IsStatusFinished) %>;

var IsDisplayPreviousCSList 	= <%= LCase(IsDisplayPreviousCSList) %>;
var IsDisplayCSMaster 			= <%= LCase(IsDisplayCSMaster) %>;
var IsDisplayItemList 			= <%= LCase(IsDisplayItemList) %>;
var IsDisplayRefundInfo 		= <%= LCase(IsDisplayRefundInfo) %>;
var IsDisplayButton 			= <%= LCase(IsDisplayButton) %>;

var IsCSCancelInfoNeeded		= <%= LCase(IsCSCancelInfoNeeded(divcd)) %>;
var IsCSRefundNeeded			= <%= LCase(IsCSRefundNeeded(divcd, OrderMasterState)) %>;

var IsPossibleModifyCSMaster	= <%= LCase(IsPossibleModifyCSMaster) %>;
var IsPossibleModifyItemList	= <%= LCase(IsPossibleModifyItemList) %>;
var IsPossibleModifyRefundInfo	= <%= LCase(IsPossibleModifyRefundInfo) %>;

var IsCSCancelProcess			= <%= LCase(IsCSCancelProcess(divcd)) %>;
var IsCSReturnProcess			= <%= LCase(IsCSReturnProcess(divcd)) %>;
var IsCSServiceProcess			= <%= LCase(IsCSServiceProcess(divcd)) %>;

var IsDeletedCS 				= <%= LCase(ocsaslist.FOneITem.FDeleteyn = "Y") %>;

var ERROR_MSG_TRY_MODIFY		= "<%= ERROR_MSG_TRY_MODIFY %>";

var CDEFAULTBEASONGPAY 		= <%=Cint(getDefaultBeasongPayByDate(now())) %>; // �ٹ����� �⺻ ��ۺ�
var divcd 					= "<%= divcd %>";
var mode 					= "<%= mode %>";
var giftorderserial 			= "<%= giftorderserial %>";

var IsAdminLogin 			= IsCsPowerUser; ///<%= LCase((session("ssBctId") = "icommang") or (session("ssBctId") = "iroo4") or (session("ssBctId") = "bseo")) %>;
var IsOrderFound 			= <%= LCase(ogiftcardordermaster.FResultCount > 0) %>;
var IsRefundInfoFound 		= <%= LCase(orefund.FResultCount > 0) %>;

<% if (ogiftcardordermaster.FResultCount > 0) then %>
var IsThisMonthJumun 		= <%= LCase(datediff("m", ogiftcardordermaster.FOneItem.FRegdate, now()) <= 0) %>;
<% else %>
var IsThisMonthJumun 		= false;
<% end if %>

// ============================================================================
// ������ �������� ����
// ============================================================================
function selectGubun(value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_gubun01, name_gubun02, name_gubun01name, name_gubun02name ,name_frm, targetDiv){

	if ((IsPossibleModifyCSMaster != true) || (IsPossibleModifyItemList != true)) {
		alert(ERROR_MSG_TRY_MODIFY);
		return;
	}

    var frm = eval(name_frm);

    eval("document." + name_frm + "." + name_gubun01).value = value_gubun01;
    eval("document." + name_frm + "." + name_gubun02).value = value_gubun02;
    eval("document." + name_frm + "." + name_gubun01name).value = value_gubun01name;
    eval("document." + name_frm + "." + name_gubun02name).value = value_gubun02name;

    eval(targetDiv).innerHTML = "";
}

// ============================================================================
// CS �������� ǥ�� (AJAX)
// ============================================================================
function divCsAsGubunSelect(value_gubun01,value_gubun02,name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv) {

	if ((IsPossibleModifyCSMaster != true) || (IsPossibleModifyItemList != true)) {
		alert(ERROR_MSG_TRY_MODIFY);
		return;
	}

    var params = "?gubun01=" + value_gubun01 + "&gubun02=" + value_gubun02 + "&name_gubun01=" + name_gubun01 + "&name_gubun02=" + name_gubun02 + "&name_gubun01name=" + name_gubun01name + "&name_gubun02name=" + name_gubun02name +"&name_frm=" + name_frm + "&targetDiv=" + targetDiv;
    initializeURL("/cscenter/action/ajax_cs_gubun_select.asp" + params);
    initializeReturnFunction("processAjaxCSGubunSelect(" + targetDiv + ")");
    initializeErrorFunction("onErrorAjaxCSGubunSelect()");
    startRequest();
}

function processAjaxCSGubunSelect(targetDiv) {
    eval(targetDiv).innerHTML = xmlHttp.responseText;
}

function onErrorAjaxCSGubunSelect() {
    alert("�����͸� �д� ���߿� ������ �߻��߽��ϴ�. ����� �ٽ� �õ��غ��ñ� �ٶ��ϴ�.[CODE:" + xmlHttp.status + "]");
}

function colseCausepop(targetDiv){
    eval(targetDiv).innerHTML = "";
}

// ============================================================================
// CS ����
// ============================================================================
function CsRegProc(frm) {

	// ������ üũ
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	// ���, ��ǰ, ȯ��
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true) || (divcd == "A003")) {
        if (CheckReturnForm(frm) != true) {
            return;
        }
	}

	// ���, ��ǰ
    if ((IsCSCancelProcess) || (IsCSReturnProcess)) {
		// ȯ�������� �������� üũ
		if (IsRefundInfoOK(frm) != true) {
			return;
		}

		// ȯ�Ҽ����� �ǹٸ���
		if (CheckReturnMethod(frm) != true) {
			return;
		}
    }

    if (IsCSCancelProcess){
        if(confirm("��� ���� �Ͻðڽ��ϱ�?")){
			if (frm.returnmethod) {
				if (frm.returnmethod.value == "R000") {
					frm.refundrequire.value = "0";
				}
			}

            frm.submit();
        }
    }else if (IsCSRefundNeeded) {
        if(confirm("ȯ�� ���� �Ͻðڽ��ϱ�?")){
            frm.submit();
        }
    }else if(confirm("���� �Ͻðڽ��ϱ�?")){
        frm.submit();
    }
}

function CheckCSMasterForSave(frm) {
    if (frm.divcd.value.length<1){
        alert("���� ������ �����ϼ���.");
        frm.divcd.focus();
        return false;
    }

    if (frm.title.value.length<1) {
        alert("������ �Է��ϼ���.");
        frm.title.focus();
        return false;
    }

    if (frm.gubun01.value.length<1) {
        alert("���� ������ �Է��ϼ���.");
        return false;
    }

    return true;
}

// ============================================================================
// ȯ�� ���� üũ Form
// ============================================================================
function CheckReturnForm(frm) {
    if (!frm.returnmethod) { return true; }
    if (!frm.refundrequire) { return true; }

    if (frm.returnmethod.value.length < 1) {
        alert('ȯ�� ����� ������ �ּ���.');
        frm.returnmethod.focus();
        return false;
    }

	if (frm.returnmethod.value == "R000") {
		// ȯ�Ҿ��� �̸� üũ ���Ѵ�.
		return true;
	}

	if (frm.refundrequire.value*0 != 0) {
        alert('ȯ�� �ݾ��� �Է��ϼ���.');
        frm.refundrequire.focus();
        return;
	}

	if ((frm.refundrequire.value*1 <= 0) && (frm.returnmethod.value != "R000")) {
		alert('ȯ�� �ݾ��� 0 ���� Ŀ���մϴ�. �Ǵ� ȯ�Ҿ����� �����ϼ���.');
        return false;
	}


	// ====================================================================
	if (frm.returnmethod.value=="R007") {
		// ������
        var mooconfirm = false;
        if ((frm.rebankaccount) && (frm.rebankaccount.value.length < 1)) {
            mooconfirm = true;
        }

        if ((frm.rebankownername) && (frm.rebankownername.value.length < 1)) {
            mooconfirm = true;
        }

        if ((frm.rebankname) && (frm.rebankname.value.length < 1)) {
            mooconfirm = true;
        }

        if (mooconfirm == true) {
        	// ������ ���������� ���߿� ������ �Է��� �� �ִ�.
            if (!confirm('ȯ�� ���°� �����ϴ�. \n\nȯ�� ���� ���� ��� �Ͻðڽ��ϱ�?')) {
                if ((IsStatusRegister == true) || (IsStatusEdit == true)) {
                	frm.rebankaccount.focus();
                }
                return false;
            }
        }
	}

	if (frm.returnmethod.value == "R900") {
    	if (confirm("CS���񽺰� �ƴѰ��(�����ݾ�ȯ��) ���ϸ��� ��� ��ġ������ ȯ���ϼ���.\n\n���ϸ��� ȯ�� �Ͻðڽ��ϱ�?") != true) {
    		return false;
    	}
	}

	// ====================================================================
	if ((frm.returnmethod.value=="R900") || (frm.returnmethod.value=="R910")) {
		// ���ϸ���, ��ġ��ȯ��
        if ((frm.refund_userid) && (frm.refund_userid.value.length<1)) {
            alert('��ȸ������ �������� ���� ȯ�ҹ���Դϴ�. �ٸ� ȯ�� ����� �����ϼ���.');
            return false;
        }
	}

    return true;
}

function ChangeReturnMethod(comp){
    if (comp==undefined) return;

    var returnmethod = comp.value;

    document.all.refundinfo_R007.style.display = "none";
    document.all.refundinfo_R050.style.display = "none";
    document.all.refundinfo_R100.style.display = "none";
    document.all.refundinfo_R900.style.display = "none";

    if (comp.value=="R007"){
        //������ ȯ��
        document.all.refundinfo_R007.style.display = "";
    }else if((comp.value=="R020")||(comp.value=="R080")||(comp.value=="R100")||(comp.value=="R120")||(comp.value=="R400")){
        //�ǽð� ��ü ���//ALL@ ���� ��� //�ſ�ī�� ���� ��� //�ſ�ī�� �κ����//�޴���
        document.all.refundinfo_R100.style.display = "";
    }else if(comp.value=="R050"){
        //������ ���� ���
        document.all.refundinfo_R050.style.display = "";
    }else if ((comp.value=="R900") || (comp.value=="R910")) {
        //���ϸ��� ȯ��, ��ġ�� ȯ��
        document.all.refundinfo_R900.style.display = "";
    }

}

// ============================================================================
// ȯ�� �ݾ��� ��������
// ============================================================================
function IsRefundInfoOK(frm) {

	if ((IsCSCancelProcess != true) && (IsCSReturnProcess != true)) {
		return true;
	}

	// ���, ��ǰ�� : ȯ�ұݾװ� ��
	if (frm.orgsubtotalprice && frm.refundsubtotalprice) {
	    if (frm.orgsubtotalprice.value*1 < frm.refundsubtotalprice.value*1) {
	        alert('�����ݾ� �̻����� ȯ���� �� �����ϴ�.\n\n���ϸ���, ���� ���� ȯ��üũ�ϼ���.');
	        if (IsAdminLogin != true) {
	        	return false;
	        }
	    }
	}

	if (frm.returnmethod) {
		if (frm.returnmethod.value == "R000") {
			// ȯ�Ҿ��� �̸� üũ ���Ѵ�.
			return true;
		}
	}

	if (frm.refundrequire && frm.returnmethod) {
	    if ((frm.refundsubtotalprice.value*1 < 1) && ((frm.returnmethod.value != "R000"))) {
	        alert('ȯ�Ҵ�� �ݾ��� �����ϴ�.\n\nȯ�Ҿ��� �Ǵ� ����, ���ϸ��� ���� ȯ��üũ �����ϼ���');
	        if (IsAdminLogin != true) {
	        	return false;
	        }
	    }
	}

	if (frm.remainsubtotalprice) {
	    if (frm.remainsubtotalprice.value*1 < 0) {
	        alert('��� �� ���� �ݾ��� ���̳ʽ��� �� �� �����ϴ�. - �����̳� ���ϸ��� ȯ���� üũ�� �ּ���.');
	        if (IsAdminLogin != true) {
	        	return false;
	        }
	    }
	}

	return true;
}

function CheckReturnMethod(frm) {
	if (!frm.returnmethod) { return true; }
	if (!frm.refundsubtotalprice) { return true; }

	if ((frm.returnmethod.value != "R100") && (frm.returnmethod.value != "R007") && (frm.returnmethod.value != "R020") && (frm.returnmethod.value != "R000")) {
        alert('����Ұ� ȯ�Ҽ����Դϴ�.\n\n���밡�� ȯ�Ҽ��� : ȯ�Ҿ���, �ſ�ī��/�ǽð���ü ���, ������ȯ��');
        return false;
	}

	if ((frm.accountdiv.value == "7") && ((frm.returnmethod.value != "R007") && (frm.returnmethod.value != "R000"))) {
        alert('������ȯ�� �Ǵ� ȯ�Ҿ����� �����ϼ���');
        return false;
	}

	if ((frm.accountdiv.value == "100") && ((frm.returnmethod.value != "R100") && (frm.returnmethod.value != "R000"))) {
        alert('�ſ�ī����� �Ǵ� ȯ�Ҿ����� �����ϼ���');
        return false;
	}

	if ((frm.accountdiv.value == "20") && ((frm.returnmethod.value != "R020") && (frm.returnmethod.value != "R000"))) {
        alert('�ǽð���ü��� �Ǵ� ȯ�Ҿ����� �����ϼ���');
        return false;
	}

    return true;
}

// ============================================================================
// ����
// ============================================================================
function CsRegCancelProc(frm) {
    if (confirm('��ϵ� ���� ������ ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "deletecsas";
        frm.submit();
    }
}

// ============================================================================
// ����
// ============================================================================
function CsRegEditProc(frm) {

	// ������ üũ
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	// ���, ��ǰ, ȯ��
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true) || (divcd == "A003")) {
        if (CheckReturnForm(frm) != true) {
            return;
        }
	}

	// ���, ��ǰ
    if ((IsCSCancelProcess) || (IsCSReturnProcess)) {
		// ȯ�������� �������� üũ
		if (IsRefundInfoOK(frm) != true) {
			return;
		}

		// ȯ�Ҽ����� �ǹٸ���
		if (CheckReturnMethod(frm) != true) {
			return;
		}
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "editcsas";
        frm.submit();
    }
}

// ============================================================================
// �Ϸ�ó��
// ============================================================================
function CsRegFinishProc(frm) {
	var btn = document.getElementById("btnFinishReturn");

	// ������ üũ
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	if (IsStatusFinishing && (divcd == "A007" || ((divcd == "A003") && (frm.returnmethod.value=="R007")))) {
		if (IsAdminLogin) {
			alert('�̰����� �Ϸ�ó�� �Ͽ��� �ſ�ī�� �������/������ ȯ��ó���� �̷�� ���� �ʽ��ϴ�.[���α���]');
		} else {
			alert('�̰����� �Ϸ�ó�� �Ͽ��� �ſ�ī�� �������/������ ȯ��ó���� �̷�� ���� �ʽ��ϴ�.\n\n�Ϸ� ó�� �� �� �����ϴ�.');
			return;
		}
	}

    //ȯ�ҿ�û , �ſ�ī�� ��ҿ�û
    if ((divcd == "A003") || (divcd == "A007")) {
        if (frm.contents_finish.value.length<1){
            alert('ó�� ������ �Է��ϼ���.');
            frm.contents_finish.focus();
            return;
        }
    }

    var confirmMsg ;
    confirmMsg = '�Ϸ�ó�� ���� �Ͻðڽ��ϱ�?';

	if (btn) {
		btn.disabled = true;
	}

    if (confirm(confirmMsg )) {
        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "";
        frm.submit();
    }

	if (btn) {
		btn.disabled = false;
	}
}
</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" class="a">
<form name="frmaction" method="post" action="pop_cs_giftcard_action_new_process.asp">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="modeflag2" value="<%= mode %>">
<input type="hidden" name="giftorderserial" value="<%= giftorderserial %>" >
<input type="hidden" name="id" value="<%= ocsaslist.FOneItem.Fid %>">
<input type="hidden" name="divcd" value="<%= ocsaslist.FOneItem.FDivCd %>">
<input type="hidden" name="ipkumdiv" value="<%= ogiftcardordermaster.FOneItem.Fipkumdiv %>">
<input type="hidden" name="accountdiv" value="<%= ogiftcardordermaster.FOneItem.Faccountdiv %>">
<input type="hidden" name="orgitemcostsum" value="<%= ogiftcardordermaster.FOneItem.Fsubtotalprice %>">





<!-- ====================================================================== -->
<!-- 1. ���� CS ����                                                        -->
<!-- ====================================================================== -->
<!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_prev_cslist.asp" -->



<!-- ====================================================================== -->
<!-- 2. CS ������ ����                                                      -->
<!-- ====================================================================== -->
<!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_master_info.asp" -->



<!-- ====================================================================== -->
<!-- 3. ��ǰ����                                                            -->
<!-- ====================================================================== -->
<!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_item_list.asp" -->


</table>



<!-- ====================================================================== -->
<!-- 4. ���/ȯ��/��ü���� ����                                             -->
<!-- ====================================================================== -->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0"  class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
    <td bgcolor="#FFFFFF" width="500" valign="top">
        <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="BABABA">
        <tr height="25">
            <td colspan="5" bgcolor="<%= adminColor("topbar") %>">
            	<img src="/images/icon_star.gif" align="absbottom">
            	&nbsp;<b>��Ұ��� ����</b>
            </td>
        </tr>
		<!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_cancel_info.asp" -->
      </table>
    </td>
    <td bgcolor="#FFFFFF" valign="top" align="left">
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
        <tr height="25">
            <td colspan="2" bgcolor="<%= adminColor("topbar") %>">
            	<img src="/images/icon_star.gif" align="absbottom">
            	&nbsp;<b>ȯ�Ұ��� ����</b>
            </td>
        </tr>
        <!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_refund_info.asp" -->
        </table>

        <p>

    </td>
</tr>
</table>
<!-- ====================================================================== -->
<!-- 4. ���/ȯ��/��ü���� ����                                             -->
<!-- ====================================================================== -->



<!-- ====================================================================== -->
<!-- 5. ��ư                                                                -->
<!-- ====================================================================== -->
<!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_button.asp" -->
<!-- ====================================================================== -->
<!-- 5. ��ư                                                                -->
<!-- ====================================================================== -->

</form>

<script>

// ������ ���۽� �۵��ϴ� ��ũ��Ʈ
function getOnload(){

	if (IsStatusFinishing && (divcd == "A007" || divcd == "A003")) {
		if ((divcd == "A003") && (!frmaction.returnmethod)) {
			alert("�����Ϸ� ���� �ֹ��� ���� ȯ���� �� �����ϴ�.");
			frmaction.finishbutton.disabled = true;
		} else {
			if (divcd == "A007" || ((divcd == "A003") && (frmaction.returnmethod.value=="R007"))) {
				alert('�̰����� �Ϸ�ó�� �Ͽ��� \n\n\n�ſ�ī�� �������/������ ȯ��ó���� �̷�� ���� ������ �����Ͻñ� �ٶ��ϴ�.!\n\n\n\n\n\n');
			}
		}
	}

}

window.onload = getOnload;

</script>

<%

set ogiftcardordermaster = Nothing

%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
