<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �귣�� ����
' History : ������ ����
'           2021.06.18 �ѿ�� ����(����� �޴���,�̸��� �������� �������ʿ��� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim ogroup,i, groupid
dim oJungsanDiff, j

	groupid = requestCheckvar(request("groupid"),32)

set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = groupid
	ogroup.GetOneGroupInfo

set oJungsanDiff = new CPartnerGroup
    oJungsanDiff.FRectGroupid = groupid
	oJungsanDiff.GetGroupPartnerJungsanDiffList
	
Dim IsNewReg : IsNewReg = (ogroup.FResultCount<1)

dim ogroupuser
set ogroupuser = new CPartnerGroup
	ogroupuser.FPageSize = 100
	ogroupuser.FCurrPage = 1
    ogroupuser.frectgroupid = groupid
    ogroupuser.Get_partner_user_list

dim existsetccount
	existsetccount=0

if ogroupuser.FResultCount > 0 then
for i = 0 to ogroupuser.FResultCount-1
	' 10 �߰������
	if ogroupuser.FItemList(i).fgubun="10" then
		existsetccount = existsetccount + 1
	end if
next
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function CopyZip(flag,post1,post2,add,dong){
	if (flag=="s"){
		frmupche.company_zipcode.value= post1 + "-" + post2;
		frmupche.company_address.value= add;
		frmupche.company_address2.value= dong;
	}else if(flag=="m"){
		frmupche.return_zipcode.value= post1 + "-" + post2;
		frmupche.return_address.value= add;
		frmupche.return_address2.value= dong;
	}
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function SameReturnAddr(bool){
	if (bool){
		frmupche.return_zipcode.value = frmupche.company_zipcode.value;
		frmupche.return_address.value = frmupche.company_address.value;
		frmupche.return_address2.value = frmupche.company_address2.value;
	}else{
		frmupche.return_zipcode.value = "";
		frmupche.return_address.value = "";
		frmupche.return_address2.value = "";
	}
}

function SaveUpcheInfo(frm){
    var psocno =frm.psocno.value;
    <% if (IsNewReg) then %>
    if (frm.validchk.value!="Y"){
        alert('����� ��ȣ �ߺ�Ȯ�� �� ��ϰ����մϴ�.');
        return;
    }
    <% else %>
    if ((psocno!='')&&(psocno!=frm.company_no.value)){
        <% if (C_ADMIN_AUTH=true or C_OFF_AUTH or C_MD_AUTH) then %>
        if (!confirm('����ڹ�ȣ ����� ��ü�� ���� ��� �ؾ� �մϴ�. ����Ͻðڽ��ϱ�?')) return;
        <% else %>
        alert('����ڹ�ȣ ����� ��ü�� ���� ��� �ؾ� �մϴ�. ');
        return;
        <% end if %>
    }
    <% end if %>

	var errMsg = chkIsValidJungsanGubun(frm.company_no.value, frm.jungsan_gubun.value);
	if (errMsg != "OK") {
		alert(errMsg);
		retutn;
	}

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

	<% ' �������� ������ �ʼ�������.. ����� ��� %>
	if (frm.jungsan_name.value.length<1){
		alert('�������� ������ �Է��ϼ���.');
		frm.jungsan_name.focus();
		return;
	}
	if (frm.jungsan_phone.value.length<1){
		alert('�������� ��ȭ��ȣ�� �Է��ϼ���.');
		frm.jungsan_phone.focus();
		return;
	}
	if (frm.jungsan_email.value.length<1){
		alert('�������� �̸����� �Է��ϼ���.');
		frm.jungsan_email.focus();
		return;
	}
	if (frm.jungsan_hp.value.length<1){
		alert('�������� �ڵ����� �Է��ϼ���.');
		frm.jungsan_hp.focus();
		return;
	}

    <% if C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH then %>
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
	<% end if %>

	<% if existsetccount>1 then %>
		for(var i=0; i<frm.etc_idx.length; i++){
			if(frm.etc_name[i].value==''){
				alert('�߰�����ڸ��� �Է��� �ּ���.');
				frm.etc_name[i].focus()
				return;
			}
			if(frm.etc_hp[i].value==''){
				alert('�߰�������� �ڵ��� ��ȣ�� �Է��� �ּ���.');
				frm.etc_hp[i].focus()
				return;
			}
			if(frm.etc_email[i].value==''){
				alert('�߰�������� �̸����� �Է��� �ּ���.');
				frm.etc_email[i].focus()
				return;
			}
		}
	<% elseif existsetccount=1 then %>
		if(frm.etc_name.value==''){
			alert('�߰�����ڸ��� �Է��� �ּ���.');
			frm.etc_name.focus()
			return;
		}
		if(frm.etc_hp.value==''){
			alert('�߰�������� �ڵ��� ��ȣ�� �Է��� �ּ���.');
			frm.etc_hp.focus()
			return;
		}
		if(frm.etc_email.value==''){
			alert('�߰�������� �̸����� �Է��� �ּ���.');
			frm.etc_email.focus()
			return;
		}
	<% end if %>
	
	if(frm.addetc_name.value!=''){
		if(frm.addetc_hp.value==''){
			alert('�߰�������� �ڵ��� ��ȣ�� �Է��� �ּ���.');
			frm.addetc_hp.focus()
			return;
		}
		if(frm.addetc_email.value==''){
			alert('�߰�������� �̸����� �Է��� �ּ���.');
			frm.addetc_email.focus()
			return;
		}
		frm.etcaddyn.value="Y"
	}

	if (frm.groupid.value.length<1){
		var ret = confirm('��ü ������ ���� �Ͻðڽ��ϱ�?');
	}else{
		var ret = confirm('���� �׷��ڵ忡 ���� �귣�������� �ϰ� �����˴ϴ�.(��ǰ�ּ�,��۴���������� ����) \n\n���� �Ͻðڽ��ϱ�?');
	}

	if (ret){
		frm.submit();
	}
}

var bValidSoc = false;
function SearchSocno(frm){
    <% if (IsNewReg) then %>
    if (bValidSoc){
        clearSocField(false);
        return;
    }
    <% end if %>

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
	    //���� �����ϴ� ����ڷ� ���� �Ұ�
	    var psocno =frm.psocno.value;
        if ((psocno!='')&&(psocno!=frm.company_no.value)){
	        icheckframe.location.href="icheckframe.asp?mode=CheckSocno&socno=" + frm.company_no.value;
	    }
		// alert('����ڹ�ȣ�� ������ ��� ���� ������ ����˴ϴ�.');
	}

}

function AddProc(mode){
	alert('��ϰ����� ����ڹ�ȣ�Դϴ�.');
	clearSocField(true);
}

function clearSocField(bool){
    bValidSoc = bool;
    <% if (IsNewReg) then %>
	var frm = document.frmupche;

	if (!bValidSoc){
	    frm.company_no.value="";
	    frm.validchk.value="";
	    frm.company_no.style.backgroundColor ="#FFFFFF";
	}else{
	    frm.validchk.value="Y";
	    frm.company_no.style.backgroundColor ="#EEEEEE";
	}
	frm.company_no.readOnly=(bValidSoc);
	if ( document.getElementById("coSearchBtn") ) {
		document.getElementById("coSearchBtn").value=(bValidSoc)?"���Է�":"�ߺ�Ȯ��";
	}


	if (!bool){frm.company_no.focus();}
	<% end if %>
}

function PopUpcheList(frmname){
	var popwin = window.open("/admin/lib/popupchelist.asp?frmname=" + frmname,"popupchelist","width=700 height=540 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpcheReturnAddrOnly(groupid){
	if (groupid == "") {
		alert("�׷��ڵ尡 �����ϴ�.");
		return;
	}


	var popwin = window.open("/admin/member/popupchereturnaddronly.asp?groupid=" + groupid,"popupchereturnaddronly","width=900 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function EditErpCustCD(){
    var frm = document.frmupche;

    var param = "?groupid="+frm.groupid.value
    var popwin = window.open('popErpGroupLinkEdit.asp'+param,'popErpGroupLinkEdit','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

var orgjungsan_gubun = "<%= ogroup.FOneItem.Fjungsan_gubun %>";
if (orgjungsan_gubun == "") {
	orgjungsan_gubun = "�Ϲݰ���";
}
function fnJungsanGubunChanged() {
	var frm = document.frmupche;
	var company_no = document.getElementById("company_no");

	<% 
	'�̹��� �̻�� ��û(���������� ������ �ٲپ�� �մϴ�.����Ҷ� �Է��� �ȵǰ�, ������ ������ �ֳ׿�)���� �ּ�ó��.	' 2022.02.23 �ѿ��
	'if (Not IsNewReg) then 
	%>
	//	if ((orgjungsan_gubun != frm.jungsan_gubun.value) && ((orgjungsan_gubun == "����(�ؿ�)") || (frm.jungsan_gubun.value == "����(�ؿ�)"))) {
	//		alert("����(�ؿ�) - �Ϲݻ���� �� ���������� ������ �� �����ϴ�.\n\n��ü�� ���� ����ϼ���.");
	//		frm.jungsan_gubun.value = orgjungsan_gubun;
	//		return;
	//	}
	<% 'end if %>

	<% if (Not IsNewReg) then %>
		return;
	<% end if %>

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

		if (frm.coSearchBtn) {
			frm.coSearchBtn.disabled = true;
		}

		frm.validchk.value = "Y";
	} else {
		company_no.className = "text";
		frm.company_no.readOnly = false;
		frm.company_no.value = "";

		if (frm.coSearchBtn) {
			frm.coSearchBtn.disabled = false;
		}

		frm.validchk.value = "N";
	}
}

function chkIsValidJungsanGubun(company_no, jungsan_gubun) {
	// 000-00-00000
	// ��� �α��� : �����ڵ�
	// =========================================================================
	// 01-79 : ���λ����+���������
	// 90-99 : ���λ����+�鼼�����
	// ��Ÿ : ���� �鼼 ��� ����
	// ���ڸ� 888 = ����(�ؿ�)
	// =========================================================================

	if (company_no.length != 12) {
		// �ֹ���Ϲ�ȣ(��õ¡��) �� üũ ����.
		return "OK";
	}

	var soc_gubun = company_no.substring(4, 6)*1;
	var IsForeign = (company_no.substring(0, 3) == "888");

	if (IsForeign) {
		if (jungsan_gubun != "����(�ؿ�)") {
			return "����(�ؿ�) ����ڸ� ������ ����ڹ�ȣ�Դϴ�.";
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

function divadd(){
	document.all.divadd1.style.display ='';
	document.all.divadd2.style.display ='';
	document.all.divadd3.style.display ='';
}

function etc_del(etc_idx){
	if (etc_idx==''){
		alert('�������� ��ΰ� �ƴմϴ�.\n������ ��ȣ�� �����ϴ�.');
		return;
	}
	frmetcedit.etc_idx.value=etc_idx;
	frmetcedit.submit()
}

$(function(){
	<% if C_ADMIN_AUTH or C_MngPart or C_MD_AUTH or C_PSMngPart then %>
		// �귣�庰 �߰����� ���� �߰�����		// 2020.02.11 �ѿ�� ����
		$("#buttonregdiffjungsan").click(function(){
			$("#regdiffacct1").show();
			$("#regdiffacct2").show();
			$("#regdiffacct3").show();
			$("#regdiffacct4").show();
			$("#regdiffacct5").show();
		})
	<% end if %>
})

</script>

<form name="frmupche" method="post" action="/admin/member/doupcheedit.asp">
<input type="hidden" name="mode" value="groupedit">
<input type="hidden" name="uid" value="">
<input type="hidden" name="psocno" value="<%= ogroup.FOneItem.Fcompany_no %>">
<input type="hidden" name="validchk" value="">
<input type="hidden" name="etcaddyn" value="" />
<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		��ü�ڵ� : <%= ogroup.FOneItem.FGroupId %>&nbsp;&nbsp;
        ��ü�� : <%= ogroup.FOneItem.FCompany_name %>
	</td>
</tr>
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=4><b>1.��ü��������</b></td>
</tr>
<tr height="25">
	<td width="120" bgcolor="<%= adminColor("tabletop") %>">��ü�ڵ�</td>
	<td bgcolor="#FFFFFF" width="200">
		<input type="text" class="text" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" size="7" maxlength="5" style="background-color:#EEEEEE;" readonly>
	</td>
	<td width="90" bgcolor="<%= adminColor("tabletop") %>">ERP�����ڵ�</td>
	<td bgcolor="#FFFFFF" width="200">
	<% if (ogroup.FOneItem.FGroupId<>"") then %>
	    <% if (ogroup.FOneItem.FerpUsing<>1) then %>
	        ��������
	    <% else %>
		    <% if IsNULL(ogroup.FOneItem.FerpCust_CD) then %>
		    <%= ogroup.FOneItem.FGroupId %> <!--�⺻-->
		    <% else %>
    		    <% if ogroup.FOneItem.FerpCUST_USE_CD<>ogroup.FOneItem.FerpCust_CD then %>
    			<strong><%= ogroup.FOneItem.FerpCUST_USE_CD%></strong>(<%= ogroup.FOneItem.FerpCust_CD %>)
    			<% else %>
    			<%= ogroup.FOneItem.FerpCust_CD %>
    		    <% end if %>
			<% end if %>
		<% end if %>

		<% if (C_ADMIN_AUTH) or (session("ssAdminPsn") = "7") or (C_MngPart) or C_OFF_AUTH or C_MD_AUTH or C_PSMngPart then %>
			<input type="button" class="button" value="����" onClick="EditErpCustCD()">
		<% end if %>
	<% end if %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">�����귣��ID</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% if ogroup.FOneItem.getBrandList = "" then %>
			<font color="red">���� �������� �귣�尡 �����ϴ�.</font>
		<% else %>
			<%= ogroup.FOneItem.getBrandListHTML %>
		<% end if %>
	</td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**����ڵ������**</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">ȸ���(��ȣ)</td>
	<td bgcolor="#FFFFFF">
		<% if (C_ADMIN_AUTH=true or C_OFF_AUTH or C_MD_AUTH) then %>
		<input type="text" class="text" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="30" maxlength="50">
		<% else %>
		<input type="text" class="text" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="30" maxlength="50" style="background-color:#EEEEEE;" readonly>
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">��ǥ��</td>
	<td bgcolor="#FFFFFF">
		<% if (C_ADMIN_AUTH=true or C_OFF_AUTH or C_MD_AUTH) then %>
		<input type="text" class="text" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="23" maxlength="17">
		<% else %>
		<input type="text" class="text" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="23" maxlength="17" style="background-color:#EEEEEE;" readonly>
		<% end if %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">����ڹ�ȣ</td>
	<td bgcolor="#FFFFFF">
		<% if ((C_ADMIN_AUTH=true) or (C_OFF_AUTH) or C_MD_AUTH or (session("ssAdminPsn") = "7") or (IsNewReg)) and (LEN(TRIM(replace(ogroup.FOneItem.Fcompany_no,"-","")))=10) then %>
		<input type="text" class="text" id="company_no" name="company_no" value="<%= ogroup.FOneItem.Fcompany_no %>" size="16" maxlength="20" >
		<input id="coSearchBtn" name="coSearchBtn" type="button" class="button" value="�ߺ�Ȯ��" onClick="SearchSocno(frmupche)">
		<% else %>
		<input type="text" class="text" id="company_no" name="company_no" value="<%= ogroup.FOneItem.Fcompany_no %>" size="16" maxlength="20" style="background-color:#EEEEEE;" readonly>
		<% end if %>

	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
			<% if  (C_ADMIN_AUTH=true or C_OFF_AUTH or C_MD_AUTH) then %>
			<% drawSelectBoxjungsan_gubun "jungsan_gubun",ogroup.FOneItem.Fjungsan_gubun,"fnJungsanGubunChanged()" %>
		<%else%>
		<%=ogroup.FOneItem.Fjungsan_gubun%><input type="hidden" name="jungsan_gubun" value="<%=ogroup.FOneItem.Fjungsan_gubun%>"> 
		<%end if%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">����������</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" class="text" name="company_zipcode" value="<%= ogroup.FOneItem.Fcompany_zipcode %>" size="7" maxlength="7">
        <input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frmupche','C')">
		<input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('frmupche','C')">
        <% '<input type="button" class="button" value="�˻�(��)" onClick="popZip('s');"> %>
		<br>
		<input type="text" class="text" name="company_address" value="<%= ogroup.FOneItem.Fcompany_address %>" size="30" maxlength="64">&nbsp;
		<input type="text" class="text" name="company_address2" value="<%= ogroup.FOneItem.Fcompany_address2 %>" size="46" maxlength="64">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_uptae" value="<%= ogroup.FOneItem.Fcompany_uptae %>" size="30" maxlength="32"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_upjong" value="<%= ogroup.FOneItem.Fcompany_upjong %>" size="30" maxlength="32"></td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**��ü�⺻����**</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">��ǥ��ȭ</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="<%= ogroup.FOneItem.Fcompany_tel %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.company_tel)"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�ѽ�</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="<%= ogroup.FOneItem.Fcompany_fax %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.company_fax)"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�繫�� �ּ�</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<input type="text" class="text" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>" size="7" maxlength="7">
    <input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frmupche','D')">
	<input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('frmupche','D')">
    <% '<input type="button" class="button" value="�˻�(��)" onClick="popZip('m');"> %>
	<input type=checkbox name=samezip onclick="SameReturnAddr(this.checked)">��
	<br>
		<input type="text" class="text" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>" size="30" maxlength="64">&nbsp;
		<input type="text" class="text" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>" size="46" maxlength="64">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ּ�</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<input type="button" class="button" value="�귣�庰 ��۴���� �� ��ǰ�ּ� ����" onClick="PopUpcheReturnAddrOnly('<%= ogroup.FOneItem.FGroupId %>')">
	</td>
</tr>
<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">
		**������������**
		<% if C_ADMIN_AUTH or C_MngPart or C_MD_AUTH or C_PSMngPart then %>
			<input type="button" id="buttonregdiffjungsan" value="�귣�庰 �������� �߰�(��ǥ ������¿� �ٸ����)" class="button" >
		<% end if %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�ŷ�����</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<% if C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH then %>
			<% DrawBankCombo "jungsan_bank", ogroup.FOneItem.Fjungsan_bank %>
		<% else %>
			<%= ogroup.FOneItem.Fjungsan_bank %><input type="hidden" name="jungsan_bank" value="<%= ogroup.FOneItem.Fjungsan_bank %>"> 
		<% end if %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">���¹�ȣ</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<% if (C_ADMIN_AUTH) or (C_MngPart) then %>
	<input type="text" class="text" name="jungsan_acctno" value="<%= ogroup.FOneItem.Fjungsan_acctno %>" size="24" maxlength="32">
	<% else %>
	<input type="text" class="text_RO" name="jungsan_acctno" value="<%= ogroup.FOneItem.Fjungsan_acctno %>" size="20" maxlength="32" readOnly >
	<font color="red">(�濵������ ���氡��)</font>
	<% end if %>
	&nbsp;&nbsp; '-'�� ���� ��ȣ�� �Է����ֽñ� �ٶ��ϴ�.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�����ָ�</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<% if (C_ADMIN_AUTH) or (C_MngPart) then %>
	<input type="text" class="text" name="jungsan_acctname" value="<%= ogroup.FOneItem.Fjungsan_acctname %>" size="24" maxlength="16">
	<% else %>
	<input type="text" class="text_ro" name="jungsan_acctname" value="<%= ogroup.FOneItem.Fjungsan_acctname %>" size="20" maxlength="16" readOnly >
	<font color="red">(�濵������ ���氡��)</font>
	<% end if %>

	&nbsp;&nbsp; ���� ���� ���ñ� �ٶ��ϴ�.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<% if C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH then %>
			�¶��� : <% DrawJungsanDateCombo "jungsan_date", ogroup.FOneItem.Fjungsan_date %>
			&nbsp;
			�������� : <% DrawJungsanDateCombo "jungsan_date_off", ogroup.FOneItem.Fjungsan_date_off %>
		<% else %>
			�¶��� : <%= ogroup.FOneItem.Fjungsan_date %><input type="hidden" name="jungsan_date" value="<%= ogroup.FOneItem.Fjungsan_date %>">
			&nbsp;
			�������� : <%= ogroup.FOneItem.Fjungsan_date_off %><input type="hidden" name="jungsan_date_off" value="<%= ogroup.FOneItem.Fjungsan_date_off %>">
		<% end if %>
	</td>
</tr>
<% if (oJungsanDiff.FresultCount>0) then %>
<% for j=0 to oJungsanDiff.FresultCount-1 %>
<tr>
	<td colspan="4" bgcolor="CCCC22" height="25">**�귣�庰 ������������ : ��ǥ ������¿� �ٸ���� ��� ** - <strong><%=oJungsanDiff.FItemList(j).Fpartnerid %></strong></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�ŷ�����</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<input type="hidden" name="jungsan_add_brand" value="<%=oJungsanDiff.FItemList(j).Fpartnerid%>">
	<% DrawBankCombo "jungsan_bank_add", oJungsanDiff.FItemList(j).Fjungsan_bank %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">���¹�ȣ</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<% if (C_ADMIN_AUTH) or (C_MngPart) then %>
	<input type="text" class="text" name="jungsan_acctno_add" value="<%= oJungsanDiff.FItemList(j).Fjungsan_acctno %>" size="24" maxlength="32">
	<% else %>
	<input type="text" class="text_RO" name="jungsan_acctno_add" value="<%= oJungsanDiff.FItemList(j).Fjungsan_acctno %>" size="20" maxlength="32" readOnly >
	<font color="red">(�濵������ ���氡��)</font>
	<% end if %>
	&nbsp;&nbsp; '-'�� ���� ��ȣ�� �Է����ֽñ� �ٶ��ϴ�.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�����ָ�</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<% if (C_ADMIN_AUTH) or (C_MngPart) then %>
	<input type="text" class="text" name="jungsan_acctname_add" value="<%= oJungsanDiff.FItemList(j).Fjungsan_acctname %>" size="24" maxlength="16">
	<% else %>
	<input type="text" class="text_ro" name="jungsan_acctname_add" value="<%= oJungsanDiff.FItemList(j).Fjungsan_acctname %>" size="20" maxlength="16" readOnly >
	<font color="red">(�濵������ ���氡��)</font>
	<% end if %>

	&nbsp;&nbsp; ���� ���� ���ñ� �ٶ��ϴ�.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	�¶��� : <% DrawJungsanDateCombo "jungsan_date_add", oJungsanDiff.FItemList(j).Fjungsan_date %>
	&nbsp;
	�������� : <% DrawJungsanDateCombo "jungsan_date_off_add", oJungsanDiff.FItemList(j).Fjungsan_date_off %>
	</td>
</tr>
<% next %>
<% end if %>

<% if C_ADMIN_AUTH or C_MngPart or C_MD_AUTH or C_PSMngPart then %>
	<tr id="regdiffacct1" style="display:none;">
		<td colspan="4" bgcolor="CCCC99" height="25">
			**�귣�庰 ������������ �Է� : ��ǥ ������¿� �ٸ���� �Է� ** - <% DrawAcctDiffBrand "jungsan_add_brand","",groupid,"" %>
		</td>
	</tr>
	<tr id="regdiffacct2" style="display:none;">
		<td bgcolor="<%= adminColor("tabletop") %>">�ŷ�����</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<% DrawBankCombo "jungsan_bank_add", "" %>
		</td>
	</tr>
	<tr id="regdiffacct3" style="display:none;">
		<td bgcolor="<%= adminColor("tabletop") %>">���¹�ȣ</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="jungsan_acctno_add" value="" size="24" maxlength="32">
			&nbsp;&nbsp; '-'�� ���� ��ȣ�� �Է����ֽñ� �ٶ��ϴ�.
		</td>
	</tr>
	<tr id="regdiffacct4" style="display:none;">
		<td bgcolor="<%= adminColor("tabletop") %>">�����ָ�</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="jungsan_acctname_add" value="" size="24" maxlength="16">
			&nbsp;&nbsp; ���� ���� ���ñ� �ٶ��ϴ�.
		</td>
	</tr>
	<tr id="regdiffacct5" style="display:none;">
		<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			�¶��� : <% DrawJungsanDateCombo "jungsan_date_add", "" %>
			&nbsp;
			�������� : <% DrawJungsanDateCombo "jungsan_date_off_add", "" %>
		</td>
	</tr>
<% end if %>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**���������**</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">����ڸ�</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>" size="30" maxlength="32"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.manager_phone)"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>" size="30" maxlength="64"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.manager_hp)"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">��۴���ڸ�</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="button" class="button" value="�귣�庰 ��ǰ�ּ� ����" onClick="PopUpcheReturnAddrOnly('<%= ogroup.FOneItem.FGroupId %>')">
		(�귣�庰�� ���� �����մϴ�.)
	</td>
</tr>
<tr>
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">�������ڸ�</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>" size="30" maxlength="32"></td>
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.jungsan_phone)"></td>
</tr>
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>" size="30" maxlength="64"></td>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.jungsan_hp)"></td>
</tr>
<% if ogroupuser.FResultCount > 0 then %>
<% for i = 0 to ogroupuser.FResultCount-1 %>
<%
' 10 �߰������
if ogroupuser.FItemList(i).fgubun="10" then
%>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�߰�����ڸ�<%= i + 1 %><input type="hidden" name="etc_idx" value="<%= ogroupuser.FItemList(i).fidx %>"></td>
	<td bgcolor="#FFFFFF" colspan=3>
		<input type="text" class="text" name="etc_name" value="<%= ogroupuser.FItemList(i).fname %>" maxlength="32" />
		<button type="button" class="button" onClick="etc_del('<%= ogroupuser.FItemList(i).fidx %>');">����</button>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
	<td bgcolor="#FFFFFF" colspan=3><input type="text" class="text" name="etc_hp" value="<%= ogroupuser.FItemList(i).fhp %>" maxlength="16" /></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF" colspan=3><input type="text" class="text" name="etc_email" value="<%= ogroupuser.FItemList(i).femail %>" maxlength="64" /></td>
</tr>
<% end if %>
<% next %>
<% end if %>
<tr id="divadd1" style="display:none;">
	<td bgcolor="<%= adminColor("tabletop") %>">�߰�����ڸ�</td>
	<td bgcolor="#FFFFFF" colspan=3><input type="text" class="text" name="addetc_name" value="" maxlength="32" /></td>
</tr>
<tr id="divadd2" style="display:none;">
	<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
	<td bgcolor="#FFFFFF" colspan=3><input type="text" class="text" name="addetc_hp" value="" maxlength="16" /></td>
</tr>
<tr id="divadd3" style="display:none;">
	<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF" colspan=3><input type="text" class="text" name="addetc_email" value="" maxlength="64" /></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>"><button type="button" class="button" onClick="divadd();">�ű��߰������</button></td>
	<td bgcolor="#FFFFFF" colspan=3></td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<input type="button" class="button" value="��ü���� ����" onclick="SaveUpcheInfo(frmupche);">
	</td>
</tr>
</table>
</form>
<form name="frmetcedit" method="post" action="/admin/member/doupcheedit.asp">
<input type="hidden" name="mode" value="etc_del" />
<input type="hidden" name="etc_idx" value="" />
<input type="hidden" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" />
</form>
<iframe src="" name="icheckframe" width="0" height="0" frameborder="1" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>

<%
set oJungsanDiff = Nothing
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
