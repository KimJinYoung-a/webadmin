<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : cs����
' History	:  2007.06.01 �̻� ����
'              2023.11.15 �ѿ�� ����(6�������� �����͵� ó�������ϰ� ���� ����)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_mifinishcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
dim research, page, divcd, makerid, vSiteName, itemid, Dtype, fromdate, todate, nexttodate
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, dplusOver, dplusLower, MifinishReason, MifinishState, sortby
dim exinmaychulgoday, exoldcs, exchangemindreturn, exregbycs, order6MonthBefore, csmifinish
dim OCSBrandMemo, OCSItemMemo, ix,iy
	research 	= requestCheckVar(request("research"),32)
	page 		= requestCheckVar(request("page"),32)
	divcd 		= requestCheckVar(request("divcd"),32)
	makerid 	= requestCheckVar(request("makerid"),32)
	vSiteName 	= requestCheckVar(request("vSiteName"),32)
	itemid 		= requestCheckVar(request("itemid"),32)
	Dtype 		= requestCheckVar(request("Dtype"),32)
	yyyy1   	= requestCheckVar(request("yyyy1"),4)
	mm1     	= requestCheckVar(request("mm1"),2)
	dd1     	= requestCheckVar(request("dd1"),2)
	yyyy2   	= requestCheckVar(request("yyyy2"),4)
	mm2     	= requestCheckVar(request("mm2"),2)
	dd2     	= requestCheckVar(request("dd2"),2)
	dplusOver   	= requestCheckVar(request("dplusOver"),10)
	dplusLower   	= requestCheckVar(request("dplusLower"),10)
	MifinishReason 	= requestCheckVar(request("MifinishReason"),2)
	MifinishState  	= requestCheckVar(request("MifinishState"),2)
	sortby			= requestCheckVar(request("sortby"),32)
	exinmaychulgoday	= requestCheckVar(request("exinmaychulgoday"),32)
	exoldcs				= requestCheckVar(request("exoldcs"),32)
	exchangemindreturn	= requestCheckVar(request("exchangemindreturn"),32)
	exregbycs	= requestCheckVar(request("exregbycs"),32)
	order6MonthBefore	= requestCheckVar(request("order6MonthBefore"),1)

if (page="") then page=1
if (Dtype="") then Dtype = "dday"

if (research = "") then
	if (dplusOver = "") then
		dplusOver = "7"
	end if

	exoldcs = "Y"
	''exchangemindreturn = "Y"
end if

if (yyyy1="") then
	todate = Left(CStr(now()),10)

	yyyy2 = Left(todate,4)
	mm2   = Mid(todate,6,2)
	dd2   = Mid(todate,9,2)

	fromdate = DateSerial(yyyy2,mm2-2, dd2+1)

	yyyy1 = Left(fromdate,4)
	mm1   = Mid(fromdate,6,2)
	dd1   = Mid(fromdate,9,2)
end if

nexttodate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

set csmifinish = new CCSMifinishMaster
	csmifinish.FRectDivCD = divcd
	csmifinish.FRectDesignerID = makerid
	csmifinish.FPageSize = 50
	csmifinish.FCurrPage = page
	csmifinish.FRectMifinishReason = MifinishReason
	csmifinish.FRectMifinishState  = MifinishState
	csmifinish.FRectItemID = itemid
	csmifinish.FRectSiteName = vSiteName
	csmifinish.FRectSortBy = sortby
	csmifinish.FRectExInMayChulgoDay = exinmaychulgoday
	csmifinish.FRectExOldCS = exoldcs
	csmifinish.FRectExChangeMindReturn = exchangemindreturn
	csmifinish.FRectExRegbyCS = exregbycs
	csmifinish.FRectorder6MonthBefore = order6MonthBefore

	if (Dtype = "topN") then
		todate = Left(CStr(now()),10)
		fromdate = Left(CStr(DateAdd("m", -2, now())),10)
		nexttodate = Left(CStr(DateAdd("d", 1, CDate(todate))),10)

		csmifinish.FRectRegStart = fromdate
		csmifinish.FRectRegEnd = nexttodate
		csmifinish.FPageSize = 300

		csmifinish.getUpcheMifinishList
	elseif (Dtype = "date") then
		csmifinish.FRectRegStart = LEft(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
		csmifinish.FRectRegEnd = nexttodate

		csmifinish.getUpcheMifinishList
	elseif (Dtype = "dday") then
		csmifinish.FRectdplusOver = dplusOver
		csmifinish.FRectdplusLower = dplusLower

		csmifinish.getUpcheMifinishList
	end if

set OCSBrandMemo = new CCSBrandMemo
	OCSBrandMemo.FRectMakerid = makerid

	if (makerid <> "") then
		OCSBrandMemo.GetBrandMemo
	end if

set OCSItemMemo = new CCSItemMemo
	OCSItemMemo.FRectItemId = itemid

	if (itemid <> "") then
		OCSItemMemo.GetItemidMemo
	end if

%>
<script type="text/javascript">

function chkSubmit(){
    var frm = document.frm;

    if ((frm.itemid.value.length>0)&&(!IsDigit(frm.itemid.value))){
        alert('��ǰ��ȣ�� ���ڷ� �Է��ϼ���.');
        frm.itemid.focus();
        return;
    }

    frm.yyyy1.disabled=false;
    frm.yyyy2.disabled=false;
    frm.mm1.disabled=false;
    frm.mm2.disabled=false;
    frm.dd1.disabled=false;
    frm.dd2.disabled=false;

    frm.submit();
}

function MifinishCSMaster(v){
	var popwin = window.open("/cscenter/mifinish/cs_mifinishmaster_main.asp?asid=" + v,"MifinishMaster","width=1400 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function ViewItem(itemid){
window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"sample");
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function chkComp(comp){
	var selectedval = comp.value;

	if (selectedval == "topN") {
	    comp.form.yyyy1.disabled = true;
	    comp.form.yyyy2.disabled = true;
	    comp.form.mm1.disabled = true;
	    comp.form.mm2.disabled = true;
	    comp.form.dd1.disabled = true;
	    comp.form.dd2.disabled = true;

	    comp.form.dplusOver.disabled = true;
	    comp.form.dplusLower.disabled = true;
	} else if (selectedval == "date") {
	    comp.form.yyyy1.disabled = false;
	    comp.form.yyyy2.disabled = false;
	    comp.form.mm1.disabled = false;
	    comp.form.mm2.disabled = false;
	    comp.form.dd1.disabled = false;
	    comp.form.dd2.disabled = false;

	    comp.form.dplusOver.disabled = true;
	    comp.form.dplusLower.disabled = true;
	} else if (selectedval == "dday") {
	    comp.form.yyyy1.disabled = true;
	    comp.form.yyyy2.disabled = true;
	    comp.form.mm1.disabled = true;
	    comp.form.mm2.disabled = true;
	    comp.form.dd1.disabled = true;
	    comp.form.dd2.disabled = true;

	    comp.form.dplusOver.disabled = false;
	    comp.form.dplusLower.disabled = false;
	}
}

function searchByMakerId(frm, makerid) {
	frm.makerid.value = makerid;
	frm.itemid.value = "";

	chkSubmit();
}

function searchByItemId(frm, itemid) {
	frm.itemid.value = itemid;

	chkSubmit();
}

function jsShowHideObject(id) {
	if (document.getElementById) {
		obj = document.getElementById(id);

		if (obj.style.display == "none") {
			obj.style.display = "";
		} else {
			obj.style.display = "none";
		}
	}
}

function jsUpcheBrandReturnMemo(makerid, tableobj)
{

	if (makerid == "") {
		alert("���� �귣��� �˻��ϼ���.");
		return;
	}

	jsShowHideObject(tableobj);
}

function submitSaveBrandMemo(frm)
{

	if (frm.makerid.value == "") {
		alert("���� �귣��� �˻��ϼ���.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsUpcheItemReturnMemo(itemid, tableobj)
{

	if (itemid == "") {
		alert("���� ��ǰ�ڵ�� �˻��ϼ���.");
		return;
	}

	jsShowHideObject(tableobj);
}

function submitSaveItemMemo(frm)
{

	if (frm.itemid.value == "") {
		alert("���� ��ǰ�ڵ�� �˻��ϼ���.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsMultiReturnReason(makerid, itemid, tableobj)
{

	if ((makerid == "") && (itemid == "")) {
		alert("�귣�� �Ǵ� ��ǰ�ڵ�� �˻��ϼ���.");
		return;
	}

	jsShowHideObject(tableobj);
}

function jsSetReturnReason(frm) {
	if (frm.regReturnReason.value == "26") {
		frm.nextactday.value = "";
	}

	jsSetSMSMailText(frm);
}

function jsSetSMSMailText(frm) {
	jsSetSMSText(frm);
	jsSetMailText(frm);
}

function jsSetSMSText(frm) {
	var smsText;

	smsText = "";

	if (frm.regReturnReason.value == "25") {
		smsText = "[�ٹ����� ��ü��ǰ�ȳ�] ����, �����Ͻ� ��ǰ [��ǰ��]([��ǰ�ڵ�])�� ��ü�� �ݼ� ��";
		smsText = smsText + " �ݼ��Ͻ� ������ȣ �˷��ֽø� Ȯ�� �� ȯ��ó�� �����ص帮�ڽ��ϴ�.";
		smsText = smsText + " ���� �̹ݼ��Ͻ� ��� ��ü�� �ݼ� ��Ź �帳�ϴ�.";
	} else if (frm.regReturnReason.value == "26") {
		smsText = "[�ٹ����� ��ǰöȸ�ȳ�] ����, �����Ͻ� ��ǰ [��ǰ��]([��ǰ�ڵ�])�� ���ۻ�ǰ���� ��ǰ�� �����ʴϴ�."
		smsText = smsText + " ����帮������ �˼��ϸ� ����öȸ�ʿ� ���� ���عٶ��ϴ�.."
		smsText = smsText + " Ȥ, �̹� �ݼ��Ͻ� ��� �����ͷ� ���� ��Ź�帳�ϴ�.�����մϴ�."
	}

	frm.sendsmsmsg.value = smsText;
}

function jsSetMailText(frm) {
	var mailText;

	mailText = "�ȳ��ϼ���. ����\n";
	mailText = mailText + "�ٹ����� ���ູ�����Դϴ�.\n\n";

	if (frm.regReturnReason.value == "25") {
		mailText = mailText + "���Բ��� ��ǰ�����Ͻ� ��ǰ�� ��ü�� �ݼ��ϼ̴�����?\n"
		mailText = mailText + "���� ��ǰ �����̽ø� �����Ͻ� �ù�� �̿��Ͽ� ��ü�� ��ǰ ��Ź �帮��,\n"
		mailText = mailText + "��ǰ �� �ݼ���(��ǰ �����)��ȣ��\n\n"
		mailText = mailText + "�ٹ����� Ȩ������(PCȭ��) > �����ٹ����� > ���� ��û�� ����\n\n"
		mailText = mailText + "���� ��ǰ �����Ͻ� ������ �Է��� �ֽø� ���� ���� ȯ��ó�� �������� �˷��帳�ϴ�.\n\n"
		mailText = mailText + "�����մϴ�."
		mailText = mailText + ""
	} else if (frm.regReturnReason.value == "26") {
		mailText = mailText + "����.\n"
		mailText = mailText + "�ٸ��� �ƴϿ��� ���Բ��� ��ǰ�����Ͻ� ��ǰ [��ǰ��]([��ǰ�ڵ�])�� ���ۻ�ǰ����\n"
		mailText = mailText + "��ǰ�� �����ʴϴ�\n"
		mailText = mailText + "�˼�������, �������ֽ� ��ǰ������ öȸ�Ǿ�����, ���� ���� ���غ�Ź�帮��\n"
		mailText = mailText + "����帮�� ���� ���� �˼��մϴ�\n"
		mailText = mailText + "���� ��ǰ������ ���� ��ǰ�� ���� �� ������ �ֵ��� ����ϰڽ��ϴ�\n"
		mailText = mailText + "�ƿ﷯ �̹� ��ü�� �ݼ��ϽŰ��ø�, ���ŷο�ô��� �ݼ����ȣ�� �ù�縦\n"
		mailText = mailText + "Ȯ���Ͻþ� �����ͷ� ������Ź�帳�ϴ�.\n"
		mailText = mailText + "1:1 �Խ��� �Ǵ� ����ȸ�����ּŵ� �ǽʴϴ�\n\n\n"
		mailText = mailText + "����� ����� ����ϰ�, �� ���Ծ��� �������� ������ ��ðڽ��ϴ�.\n"
		mailText = mailText + "�ٸ� �� �ñ��Ͻ� ������ �������� �����ͷ� �����ֽø� ģ���� �ȳ��ص帮������\n"
		mailText = mailText + "������ �ູ�� �� �ǽñ� �ٶ��ϴ�.~\n"
	}

	mailText = mailText + "\n\n������ �����ð�\n"
	mailText = mailText + "���� AM 09:00 ~ PM 06:00\n"
	mailText = mailText + "���ɽð� PM 12:00~01:00 ����Ϥ������� �޹�\n"
	mailText = mailText + "�� 1644-6030\n"
	mailText = mailText + "customer@10x10.co.kr\n"

	frm.sendmailmsg.value = mailText;
}

function CheckNcalendarOpen(returneason, nextactday) {
	if (returneason.value == "26") {
		// ��ǰ�Ұ�
		alert("��ǰ�Ұ� �ȳ��� ��� ����ó���������� �Է��� �� �����ϴ�.");
		return;
	}

	calendarOpen(nextactday);
}

function multiReturnInput(frm) {
	if (CheckSelected() != true) {
		alert("���õ� �ֹ��� �����ϴ�.");
		return;
	}

	if (frm.regReturnReason.value == "") {
		alert("����� ������ �����ϼ���.");
		frm.regReturnReason.focus();
		return;
	}

	if ((frm.ckSendSMS.checked != true) && (frm.ckSendEmail.checked != true)) {
		alert("SMS �� ���Ϲ߼� ���� �ϳ��� üũ�ؾ� �մϴ�.");
		return;
	}

	/*
	if ((frm.nextactday.value.length != 10) && (frm.regReturnReason.value != "26")) {
		alert("����ó���������� �Է��ϼ���.");
		frm.nextactday.focus();
		return;
	}
	*/

	if (frm.sendsmsmsg.value == "") {
		alert("SMS�߼۹����� �Է��ϼ���.");
		frm.sendsmsmsg.focus();
		return;
	}

	if (frm.sendmailmsg.value == "") {
		alert("MAIL�߼۹����� �Է��ϼ���.");
		frm.sendmailmsg.focus();
		return;
	}

	if (confirm("�ϰ������Ͻðڽ��ϱ�?") == true) {

		for (var i=0;i<document.forms.length;i++){
			f = document.forms[i];
			if (f.name.substr(0,9)=="frmBuyPrc") {
				if (f.csdetailidx.checked) {
					frm.arrcsdetailidx.value = frm.arrcsdetailidx.value + "," + f.csdetailidx.value;
				}
			}
		}

		frm.submit();
	}
}

function CheckSelected(){
	var pass = false;
	var frm;

	for (var i = 0; i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.csdetailidx.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function getOnLoad(){
    var idx = 0;

    for (var i = 0; i < document.frm.Dtype.length; i++) {
    	if (document.frm.Dtype[i].value == "<%= Dtype %>") {
    		idx = i;
    		break;
    	}
    }

    chkComp(document.frm.Dtype[idx]);
}

window.onload=getOnLoad;

</script>


<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>" height="120">�˻�<br>����</td>
		<td align="left">
			���� :
			<select name="divcd" class="select">
				<option value="">-��ü-</option>
				<option value="chulgocs" <%=CHKIIF(divcd="chulgocs","selected","")%>>��ü CS���</option>
				<option value="returncs" <%=CHKIIF(divcd="returncs","selected","")%>>��ü ��ǰ</option>
			</select>
			&nbsp;
			�귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
			&nbsp;
			����Ʈ :
            <% call drawSelectBoxXSiteOrderInputPartnerCS("vSiteName", vSiteName) %>
			&nbsp;
			��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="6" maxlength="9">
		</td>

		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="chkSubmit();">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td>
			<input type="radio" name="Dtype" value="topN" <%= cHKIIF(Dtype="topN","checked","") %> onClick="chkComp(this);" >TOP <%= CHKIIF(Dtype = "topN",csmifinish.FPageSize,100) %>��(�ֱ�2��)
			&nbsp;
			<input type="radio" name="Dtype" value="date" <%= cHKIIF(Dtype="date","checked","") %>  onClick="chkComp(this);" >�˻��Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			<input type="radio" name="Dtype" value="dday" <%= cHKIIF(Dtype="dday","checked","") %>  onClick="chkComp(this);" >�ҿ��ϼ� :
			<select class="select" name="dplusOver">
				<option value="" >��ü</option>
				<option value="below3day" <%= CHKIIF(dplusOver="below3day","selected","") %> >D+3 �̸���ü</option>
				<option value="3" <%= CHKIIF(dplusOver="3","selected","") %> >D+3 �̻�</option>
				<option value="4" <%= CHKIIF(dplusOver="4","selected","") %> >D+4�̻�</option>
				<option value="7" <%= CHKIIF(dplusOver="7","selected","") %> >D+7�̻�</option>
				<option value="14" <%= CHKIIF(dplusOver="14","selected","") %> >D+14�̻�</option>
			</select>
			~
			<select class="select" name="dplusLower">
				<option value="" >��ü</option>
				<option value="7" <%= CHKIIF(dplusLower="7","selected","") %> >D+7�̸�</option>
				<option value="30" <%= CHKIIF(dplusLower="30","selected","") %> >D+30����</option>
				<option value="60" <%= CHKIIF(dplusLower="60","selected","") %> >D+60����</option>
				<option value="90" <%= CHKIIF(dplusLower="90","selected","") %> >D+90����</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td>
			��ó������ :
			<select class="select" name="MifinishReason">
				<option value="">��ü</option>
				<option value=""> [���] --------</option>
				<option value="00" <%= CHKIIF((MifinishReason="00" and (divcd = "" or divcd = "chulgocs")),"selected","") %> >&nbsp;&nbsp;�Է�����</option>
				<option value="03" <%= CHKIIF(MifinishReason="03","selected","") %> >&nbsp;&nbsp;�������</option>
				<option value="05" <%= CHKIIF(MifinishReason="05","selected","") %> >&nbsp;&nbsp;ǰ�����Ұ�</option>
				<option value="02" <%= CHKIIF(MifinishReason="02","selected","") %> >&nbsp;&nbsp;�ֹ�����</option>
				<option value="04" <%= CHKIIF(MifinishReason="04","selected","") %> >&nbsp;&nbsp;�����ǰ</option>
				<option value=""> [��ǰ] --------</option>
				<option value="00" <%= CHKIIF((MifinishReason="00" and divcd = "returncs"),"selected","") %> >&nbsp;&nbsp;�Է�����</option>
				<option value="25" <%= CHKIIF(MifinishReason="25","selected","") %> >&nbsp;&nbsp;�����Է� �ȳ�</option>
				<option value="26" <%= CHKIIF(MifinishReason="26","selected","") %> >&nbsp;&nbsp;��ǰ�Ұ� �ȳ�</option>
				<option value="21" <%= CHKIIF(MifinishReason="21","selected","") %> >&nbsp;&nbsp;�� ����</option>
				<option value="22" <%= CHKIIF(MifinishReason="22","selected","") %> >&nbsp;&nbsp;�� ��ǰ����</option>
				<option value="23" <%= CHKIIF(MifinishReason="23","selected","") %> >&nbsp;&nbsp;CS�ù�����</option>
				<option value="12" <%= CHKIIF(MifinishReason="12","selected","") %> >&nbsp;&nbsp;��ü����</option>
				<option value="41" <%= CHKIIF(MifinishReason="41","selected","") %> >&nbsp;&nbsp;�ù�� ��������</option>
			</select>
			&nbsp;
			ó������ :
			<select class="select" name="MifinishState">
				<option value="">��ü</option>
				<option value="0" <%= CHKIIF(MifinishState="0","selected","") %> >CS(CALL)��ó��</option>
				<option value="4" <%= CHKIIF(MifinishState="4","selected","") %> >���ȳ�</option>
				<option value="6" <%= CHKIIF(MifinishState="6","selected","") %> >CSó���Ϸ�</option>
			</select>
			&nbsp;
			���ļ��� :
			<select class="select" name="sortby">
				<option value="">�ҿ��ϼ�</option>
				<option value="makerid" <%= CHKIIF(sortby="makerid","selected","") %> >�귣��</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td>
			<input type="checkbox" class="checkbox" name="exinmaychulgoday" value="Y" <%= CHKIIF(exinmaychulgoday="Y","checked","") %>> ó�������� ���� ��ó��CS ����
			<input type="checkbox" class="checkbox" name="exoldcs" value="Y" <%= CHKIIF(exoldcs="Y","checked","") %>> ��Ⱓ(3����) ��ó��CS ����
			<input type="checkbox" class="checkbox" name="exchangemindreturn" value="Y" <%= CHKIIF(exchangemindreturn="Y","checked","") %>> ���ɹ�ǰ ����
			<input type="checkbox" class="checkbox" name="exregbycs" value="Y" <%= CHKIIF(exregbycs="Y","checked","") %> > ���������� ������
			<input type="checkbox" class="checkbox" name="exchangemindreturn11" value="Y" disabled> ������Է� ������
			<input type="checkbox" class="checkbox" name="order6MonthBefore" value="Y" <% if order6MonthBefore="Y" then response.write "checked" %>>6���������ֹ�
		</td>
	</tr>
</table>
</form>
<!-- �˻� �� -->

<br>
* D+4, D+7, D+14 ���� <font color=red>�ٹ��ϼ� ����</font>�Դϴ�.<br>
* ��ȯ����� ��� ��ȯ��ǰ ��ȸ���� CS �� �����մϴ�.
<br>
<input type="button" class="button" value="�귣�� ��ǰ���� �޸�" onClick="jsUpcheBrandReturnMemo('<%= makerid %>', 'brandmemo');" <% if (divcd <> "returncs") then %>disabled<% end if %> >
<input type="button" class="button" value="��ǰ ��ǰ���� �޸�" onClick="jsUpcheItemReturnMemo('<%= itemid %>', 'itemmemo');" <% if (divcd <> "returncs") then %>disabled<% end if %> >
<input type="button" class="button" value="��ǰ���� �ȳ� �ϰ��Է�" onClick="jsMultiReturnReason('<%= makerid %>', '<%= itemid %>', 'regallreturnreason');" <% if (divcd <> "returncs") then %>disabled<% end if %> >
<br>

<form name="frmBrandMemo" method="post" action="cs_mifinishlist_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="modifybrandmemo">
<input type="hidden" name="makerid" value="<%= makerid %>">
<div id="brandmemo" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30"><b>�귣��ID</b></td>
		<td width="20%" bgcolor="FFFFFF"><%= makerid %></td>
		<td width="10%"></td>
		<td width="25%" bgcolor="FFFFFF"></td>
		<td width="10%">����������</td>
		<td bgcolor="FFFFFF"><%= OCSBrandMemo.Freturn_modifyday %></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30"></td>
		<td width="20%" bgcolor="FFFFFF"></td>
		<td width="10%"></td>
		<td width="25%" bgcolor="FFFFFF"></td>
		<td width="10%">�ۼ���</td>
		<td bgcolor="FFFFFF"><%= OCSBrandMemo.Freturn_reguserid %></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">�귣�� ��ǰ���� �޸�</td>
		<td colspan="5" bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="return_comment" cols="100" rows="7"><%= OCSBrandMemo.Freturn_comment %></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="FFFFFF" colspan = "6" height="35">
			<input type="button" class="button_s" value=" �����ϱ� " onClick="submitSaveBrandMemo(frmBrandMemo)">
		</td>
	</tr>
</table>
<br>
</div>
</form>

<form name="frmItemMemo" method="post" action="cs_mifinishlist_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="modifyitemmemo">
<input type="hidden" name="itemid" value="<%= itemid %>">
<div id="itemmemo" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30"><b>��ǰ�ڵ�</b></td>
		<td width="20%" bgcolor="FFFFFF" align="left"><%= itemid %></td>
		<td width="10%">��ǰ����</td>
		<td width="25%" bgcolor="FFFFFF" align="left">
			<input type="radio" name="return_changemindyn" value="Y" <%= CHKIIF((OCSItemMemo.Freturn_changemindyn = "Y" or OCSItemMemo.Freturn_changemindyn = ""),"checked","") %> > �Ϲ�
			<input type="radio" name="return_changemindyn" value="N" <%= CHKIIF(OCSItemMemo.Freturn_changemindyn = "N","checked","") %> > ���ɹ�ǰ �Ұ�
		</td>
		<td width="10%"></td>
		<td bgcolor="FFFFFF" align="left">

		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="15%" height="30">����������</td>
		<td width="20%" bgcolor="FFFFFF"  align="left">
			<%= OCSItemMemo.Freturn_modifyday %>
		</td>
		<td width="10%">�ۼ���</td>
		<td width="25%" bgcolor="FFFFFF" align="left">
			<%= OCSItemMemo.Freturn_reguserid %>
		</td>
		<td width="10%"></td>
		<td bgcolor="FFFFFF" align="left">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">��ǰ ��ǰ���� �޸�</td>
		<td colspan="5" bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="return_comment" cols="100" rows="7"><%= OCSItemMemo.Freturn_comment %></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="FFFFFF" colspan = "6" height="35">
			<input type="button" class="button_s" value=" �����ϱ� " onClick="submitSaveItemMemo(frmItemMemo)">
		</td>
	</tr>
</table>
</form>
<br>
</div>

<form name="frmReturnInput" method="post" action="cs_mifinishlist_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="regallreturnreason">
<input type="hidden" name="arrcsdetailidx" value="">
<div id="regallreturnreason" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="2" height="30"><b>��ǰ �ȳ� �ϰ�����</b></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="10%" height="30">��ǰ�ȳ� ����</td>
		<td bgcolor="FFFFFF" align="left">
			<select class="select" name="regReturnReason" onChange="jsSetReturnReason(frmReturnInput)">
				<option value=""></option>
				<option value="25">�����Է� �ȳ�</option>
				<option value="26">��ǰ �Ұ�</option>
			</select>
		</td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="10%">����ó��������</td>
		<td bgcolor="FFFFFF" align="left">
		    <input class="text" type="text_ro" name="nextactday" value="" size="10" maxlength="10" readonly>
		    <a href="javascript:CheckNcalendarOpen(frmReturnInput.regReturnReason, frmReturnInput.nextactday);"><img src="/images/calicon.gif" border="0" align="top" height=20></a>
		</td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="7%">���ȳ�</td>
		<td bgcolor="FFFFFF" align="left">
			<input name="ckSendSMS" type="checkbox" checked  >SMS�߼�&nbsp;
			<input name="ckSendEmail" type="checkbox" checked  >MAIL�߼�&nbsp;
		</td>
	</tr>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">SMS<br>�߼۳���</td>
		<td bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="sendsmsmsg" cols="52" rows="5"></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="30">MAIL<br>�߼۳���</td>
		<td bgcolor="FFFFFF" align="left">
			<textarea class="textarea" name="sendmailmsg" cols="90" rows="7"></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="FFFFFF" colspan = "2" height="35">
			<input type="button" class="button" value="��ǰ �ȳ� �ϰ�����" onclick="multiReturnInput(frmReturnInput);">
		</td>
	</tr>
</table>
</form>
<br>
</div>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="19">
		<% if Dtype="topN" then %>
		�˻���� : <b><% = csmifinish.FTotalCount %></b> (�ִ� <%= csmifinish.FPageSize %>�� ���� �˻��˴ϴ�.)
		<% else %>
			�˻���� : <b><% = csmifinish.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= csmifinish.FTotalpage %></b>
		<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"></td>
		<td width="30">����</td>
		<td>�귣��ID</td>
		<td width="70">�ֹ���ȣ</td>
		<td width="50">ASID</td>
		<td width="55">�ֹ���</td>
		<td width="55">������</td>
		<td width="50">��ǰ�ڵ�</td>
		<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
		<td width="30">����</td>
		<td width="60">CS�����<br>(������)</td>
		<td width="35">�ҿ�<br>�ϼ�</td>
		<td width="105">��ó������</td>
		<td width="25">����<br>�Է�</td>
		<td width="60">ó��������</td>
		<td width="65">ó������</td>
		<td width="65">��������</td>
		<td width="35">��<br>����</td>
	</tr>
	<% if csmifinish.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="19" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% else %>
	<% for ix=0 to csmifinish.FresultCount-1 %>
	<form name="frmBuyPrc_<%= ix %>" method="post" style="margin:0px;">
	<input type="hidden" name="orderserial" value="<%= csmifinish.FItemList(ix).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<% if csmifinish.FItemList(ix).IsAvailJumun then %>
	<tr class="a" align="center" bgcolor="FFFFFF">
	<% else %>
	<tr class="gray" align="center" bgcolor="DDDDDD">
	<% end if %>
		<td>
			<input type="checkbox" name="csdetailidx" value="<%= csmifinish.FItemList(ix).Fcsdetailidx %>" <% if csmifinish.FItemList(ix).FMifinishReason<>"00" or csmifinish.FItemList(ix).Fsongjangyn = "Y" then %>disabled<%end if %>>
		</td>
		<td>
			<font color="<%= csmifinish.FItemList(ix).getDivcdColor %>"><%= csmifinish.FItemList(ix).getDivcdStr %></font>
		</td>
		<td>
			<a href="javascript:searchByMakerId(frm, '<%= csmifinish.FItemList(ix).FMakerid %>')">
				<%= csmifinish.FItemList(ix).FMakerid %>
			</a>
		</td>
		<td>
			<a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= csmifinish.FItemList(ix).FOrderSerial %>')" class="zzz">
			<%= csmifinish.FItemList(ix).FOrderSerial %></a>
		</td>
		<td><a href="javascript:PopCSActionEdit(<%= csmifinish.FItemList(ix).Fasid %>,'editreginfo')"><%= csmifinish.FItemList(ix).Fasid %></a></td>
		<td>
			<%= csmifinish.FItemList(ix).FBuyname %><%'= printUserId(csmifinish.FItemList(ix).FBuyname, 1, "*") %>
		</td>
		<td>
			<%= csmifinish.FItemList(ix).FReqname %><%'= printUserId(csmifinish.FItemList(ix).FReqname, 1, "*") %>
		</td>
		<td>
			<a href="javascript:searchByItemId(frm, <%= csmifinish.FItemList(ix).FItemid %>)">
				<%= csmifinish.FItemList(ix).FItemid %>
			</a>
		</td>
		<td align="left">
			<a href="javascript:ViewItem(<% =csmifinish.FItemList(ix).FItemid  %>)"><%= csmifinish.FItemList(ix).FItemname %></a>
			<% if (csmifinish.FItemList(ix).FItemoption<>"") then %>
				<font color="blue">[<%= csmifinish.FItemList(ix).FItemoption %>]</font>
			<% end if %>
		</td>
		<td><%= csmifinish.FItemList(ix).FItemcnt %></td>
		<td><%= Left(csmifinish.FItemList(ix).Fregdate,10) %></td>
		<td><%= csmifinish.FItemList(ix).getDPlusDateStr %></td>
		<td>
		    <%= csmifinish.FItemList(ix).getMifinishText %>

		    <% if not IsNULL(csmifinish.FItemList(ix).FMifinishregdate) then %>
			    <br>(<%= Left(csmifinish.FItemList(ix).FMifinishregdate,10) %>)
		    <% end if %>
		</td>
		<td>
			<% if (csmifinish.FItemList(ix).Fsongjangyn = "Y") then %>Y<% end if %>
		</td>
		<td><%= csmifinish.FItemList(ix).FMifinishipgodate %></td>
		<td><%= csmifinish.FItemList(ix).getMifinishStateText %></td>
		<td>
			<% if Not IsNull(csmifinish.FItemList(ix).Flastupdate) then %>
				<acronym title="<%= csmifinish.FItemList(ix).Flastupdate %>"><%= Left(csmifinish.FItemList(ix).Flastupdate,10) %></acronym><br>
			<% end if %>

			<% if Not IsNull(csmifinish.FItemList(ix).Freguserid) then %>
				<%= csmifinish.FItemList(ix).Freguserid %>
			<% end if %>
		</td>
		<td>
			<a href="javascript:MifinishCSMaster('<%= csmifinish.FItemList(ix).Fasid %>');"><img src="/images/icon_search.jpg" border="0"></a>
		</td>
	</tr>
	</form>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="19" align="center">
		<% if Dtype="topN" then %>
		�ִ� <%= csmifinish.FPageSize %>�� ���� �˻��˴ϴ�.
		<% else %>
    		<% if csmifinish.HasPreScroll then %>
    			<a href="javascript:NextPage('<%= csmifinish.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>
    		<% for ix=0 + csmifinish.StartScrollPage to csmifinish.FScrollCount + csmifinish.StartScrollPage - 1 %>
    			<% if ix>csmifinish.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(ix) then %>
    			<font color="red">[<%= ix %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
    			<% end if %>
    		<% next %>

    		<% if csmifinish.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
    	<% end if %>
		</td>
	</tr>
<% end if %>
</table>

<%
set csmifinish = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
