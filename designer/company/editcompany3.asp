<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��ü����
' History : 2009.04.17 ���ʻ����� ��
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%
dim i,page

dim opartner
set opartner = new CPartnerUser
	opartner.FCurrpage = 1
	opartner.FRectDesignerID = session("ssBctId")
	opartner.FPageSize = 1
	opartner.GetOnePartnerNUser

dim ogroup
set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = opartner.FOneItem.FGroupid
	ogroup.GetOneGroupInfo

dim ooffontract
set ooffontract = new COffContractInfo
	ooffontract.FRectDesignerID = session("ssBctId")
	ooffontract.GetPartnerOffContractInfo

dim OReturnAddr
set OReturnAddr = new CCSReturnAddress
	OReturnAddr.FRectMakerid = session("ssBctId")
	OReturnAddr.GetBrandReturnAddress

%>
<script type="text/javascript">

function SaveUpcheInfo(frm){
	if (frm.groupid.value.length<1){
		alert('�׷��ڵ尡 ������ ���� �ʽ��ϴ�.- �����ڿ��� �����ϼ���.');
		frm.groupid.focus();
		return;
	}

//	if (frm.company_name.value.length<1){
//		alert('����� ��ϻ��� ȸ����� �Է��ϼ���.');
//		frm.company_name.focus();
//		return;
//	}

//	if (frm.ceoname.value.length<1){
//		alert('����� ��ϻ��� ��ǥ�ڸ��� �Է��ϼ���.');
//		frm.ceoname.focus();
//		return;
//	}

//	if (frm.company_no.value.length<1){
//		alert('����� ��� ��ȣ�� �Է��ϼ���.');
//		frm.company_no.focus();
//		return;
//	}

	//if (frm.jungsan_gubun.value.length<1){
	//	alert('���������� �����ϼ���.');
	//	frm.jungsan_gubun.focus();
	//	return;
	//}

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

	if (frm.return_zipcode.value.length<1){
		alert('�繫���ּ� �����ȣ�� �����ϼ���.');
		frm.return_zipcode.focus();
		return;
	}

	if (frm.return_address.value.length<1){
		alert('�繫���ּ�1 �� �Է��ϼ���.');
		frm.return_address.focus();
		return;
	}

	if (frm.return_address2.value.length<1){
		alert('�繫���ּ�2�� �Է��ϼ���.');
		frm.return_address2.focus();
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

	var ret = confirm('��ü ������ ���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}

function SaveBrandReturnInfo(frm){
	var ret = confirm('�귣�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}

function CopyZip(flag,post1,post2,add,dong){
	if (flag=="s"){
		frmupche.company_zipcode.value= post1 + "-" + post2;
		frmupche.company_address.value= add;
		frmupche.company_address2.value= dong;
	}else if(flag=="m"){
		frmupche.return_zipcode.value= post1 + "-" + post2;
		frmupche.return_address.value= add;
		frmupche.return_address2.value= dong;
	}else if(flag=="b"){
		frmbrand.return_zipcode.value= post1 + "-" + post2;
		frmbrand.return_address.value= add;
		frmbrand.return_address2.value= dong;
	}
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
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

function EditPass(frm){
	if (!frm.txoldpassword.value){
		alert('���� ��й�ȣ�� �Է��ϼ���.');
		frm.txoldpassword.focus();
		return;
	}
	
	if (!frm.txnewpassword1.value){
		alert('�����Ͻ� 1�� ��й�ȣ�� �Է��ϼ���.');
		frm.txnewpassword1.focus();
		return;
	}
	
	if (frm.txnewpassword1.value.length < 8 || frm.txnewpassword1.value.length > 16){
			alert("��й�ȣ�� ������� 8~16���Դϴ�.");
			frm.txnewpassword1.focus();
			return ;
	}
	
	var uid = "<%=session("ssBctId")%>";
	
		if(frm.txnewpassword1.value==uid) {
			alert("���̵�� �ٸ� ��й�ȣ�� ������ּ���.");
			frm.txnewpassword1.focus();
			return  ;
		}
		
		 if (!fnChkComplexPassword(frm.txnewpassword1.value)) {
				alert('�н������ ����/����/Ư������ �� �� ���� �̻��� �������� �Է����ּ���.');
				frm.txnewpassword1.focus();
				return;
			}

 	if(!frm.txnewpassword2.value) {
			alert("��й�ȣ Ȯ���� �Է����ּ���.");
			frm.txnewpassword2.focus();
			return  ;
		}
		
		
	if (frm.txnewpassword1.value!=frm.txnewpassword2.value){
		alert('1�� ��й�ȣ Ȯ���� ��ġ���� �ʽ��ϴ�.');
		frm.txnewpassword2.focus();
		return;
	}
	
	if (!frm.txnewpasswordS1.value){
		alert('�����Ͻ� 2�� ��й�ȣ�� �Է��ϼ���.');
		frm.txnewpasswordS1.focus();
		return;
	}

  if (frm.txnewpasswordS1.value.length < 8 || frm.txnewpasswordS1.value.length > 16){
			alert("��й�ȣ�� ������� 8~16���Դϴ�.");
			frm.txnewpasswordS1.focus();
			return ;
	}
	
	if (!fnChkComplexPassword(frm.txnewpasswordS1.value)) {
				alert('�н������ ����/����/Ư������ �� �� ���� �̻��� �������� �Է����ּ���.');
				frm.txnewpasswordS1.focus();
				return;
			}
			
	if(!frm.txnewpasswordS2.value) {
			alert("2�� ��й�ȣ Ȯ���� �Է����ּ���.");
			frm.txnewpasswordS2.focus();
			return  ;
		}
		 
	
	if (frm.txnewpasswordS1.value!=frm.txnewpasswordS2.value){
		alert('2�� ��й�ȣ Ȯ���� ��ġ���� �ʽ��ϴ�.');
		frm.txnewpasswordS2.focus();
		return;
	}

if (frm.txnewpassword1.value==frm.txnewpasswordS1.value){
		alert('1�� ��й�ȣ��  �ٸ� ��й�ȣ�� ������ּ���.');
		frm.txnewpasswordS1.focus();
		return;
	}
	
	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret){
		frm.submit();
	}

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

function SaveBrandEtcInfo(frm){
	if (frm.socname_kor.value.length<1){
		alert('�귣��� �ѱ��� �Է��ϼ���.');
		frm.socname_kor.focus();
		return;
	}

	if (frm.socname.value.length<1){
		alert('�귣��� ������ �Է��ϼ���.');
		frm.socname.focus();
		return;
	}

//	if (!FileCheck(frm.logoimg,150000,160,110)){
//		frm.file1.focus();
//		return;
//	}

//	if (!FileCheck(frm.titleimg,150000,610,300)){
//		frm.file2.focus();
//		return;
//	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}

}

function ChangeTitle(comp,imgcomp){
	imgcomp.src = comp.value;
}

function ChangeLogo(comp,imgcomp){
	imgcomp.src = comp.value;
}

function FileCheck(comp,maxfilesize,maxwidth,maxheight){
	if(comp.fileSize > maxfilesize){
		alert("���ϻ������ "+ maxfilesize + "byte�� �ѱ�� �� �����ϴ�...");
		return false;
	}

	if ((comp.src!="")&&(comp.width <1)){
		alert('�̹����� �����մϴ�.');
		return false;
	}

	//if(comp.width > maxwidth){
	//	alert("�������� " + maxwidth + " �ȼ��� �ѱ�� �� �����ϴ�...");
	//	return false;
	//}
	//if(comp.height > maxheight){
	//	alert("�������� " + maxheight + " �ȼ��� �ѱ�� �� �����ϴ�...");
	//	return false;
	//}

	return true;
}

function PopUpcheReturnAddrOnly(){
	var popwin = window.open("popupchereturnaddronly.asp","popupchereturnaddronly","width=1100 height=450 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<table width="600" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmupche" method="post" action="doupcheedit.asp">
<input type="hidden" name="mode" value="groupedit">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="icon" value="#">
		<font color="red"><b>��ü ����� ����</b></font>
	</td>
</tr>

<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ü�ڵ�</td>
	<td bgcolor="#FFFFFF" width="200">
		<input type="text" class="text_ro" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" size="7" maxlength="5" readonly>
	</td>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ü��</td>
	<td bgcolor="#FFFFFF" >
		<%= ogroup.FOneItem.FCompany_name %>
	</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�����귣��ID</td>
	<td colspan="3" bgcolor="#FFFFFF"><%= DdotFormat(stripHTML(ogroup.FOneItem.getBrandList),100) %></td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**��ü ����ڵ������**</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">ȸ���(��ȣ)</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="20" readonly>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">��ǥ��</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="20" readonly>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">����ڹ�ȣ</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="company_no" value="<%= socialnoReplace(ogroup.FOneItem.Fcompany_no) %>" size="20" readonly>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="" value="<%= ogroup.FOneItem.Fjungsan_gubun %>" size="20" readonly>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">����������</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" class="text" name="company_zipcode" value="<%= ogroup.FOneItem.Fcompany_zipcode %>" size="7" maxlength="7">
		<input type="button" class="button" value="�˻�" onClick="TnFindZipNewdesigner('frmupche','C')">
		<input type="button" class="button" value="�˻�(��)" onclick="javascript:popZip('s');"><br>
		<input type="text" class="text" name="company_address" value="<%= ogroup.FOneItem.Fcompany_address %>" size="26" maxlength="64">&nbsp;
		<input type="text" class="text" name="company_address2" value="<%= ogroup.FOneItem.Fcompany_address2 %>" size="38" maxlength="64">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_uptae" value="<%= ogroup.FOneItem.Fcompany_uptae %>" size="20" maxlength="32"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_upjong" value="<%= ogroup.FOneItem.Fcompany_upjong %>" size="20" maxlength="32"></td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**��ü �⺻����** &nbsp;&nbsp;(��ǰ������ �귣�庰�� �Է��� �� �ֽ��ϴ�.)</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">��ǥ��ȭ</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="<%= ogroup.FOneItem.Fcompany_tel %>" size="16" maxlength="16"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�ѽ�</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="<%= ogroup.FOneItem.Fcompany_fax %>" size="16" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�繫���ּ�</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" class="text" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>" size="7" maxlength="7">
		<input type="button" class="button" value="�˻�" onClick="TnFindZipNewdesigner('frmupche','D')">
		<input type="button" class="button" value="�˻�(��)" onclick="javascript:popZip('m');">
		<input type="checkbox" class="checkbox" name=samezip onclick="SameReturnAddr(this.checked)">��<br>
		<input type="text" class="text" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>" size="26" maxlength="64">&nbsp;
		<input type="text" class="text" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>" size="38" maxlength="64">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�귣�� ��ǰ�ּ�</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<input type="button" class="button" value="�귣�庰 ���&CS ����� �� ��ǰ�ּ� ����" onclick="PopUpcheReturnAddrOnly()">
	</td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**��ü ������������** &nbsp;&nbsp;(�������� ���� ������ ���MD���� �����Ͻñ� �ٶ��ϴ�.)</td>
</tr>

<tr height="26">
	<td bgcolor="<%= adminColor("tabletop") %>">�ŷ�����</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<input type="text" class="text_ro" value="<%= ogroup.FOneItem.Fjungsan_bank %>" size="30" readonly>
	<% if (IsNULL(ogroup.FOneItem.Fjungsan_acctno)) or (ogroup.FOneItem.Fjungsan_acctno="") then %>
	<br>(���� ���¸� ����Ͻ÷���  ��� MD���� Fax�� ���� �纻�� �����ֽñ� �ٶ��ϴ�.)
	<% end if %>
	</td>
</tr>
<tr height="26">
	<td bgcolor="<%= adminColor("tabletop") %>">���¹�ȣ</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" class="text_ro" value="<%= ogroup.FOneItem.Fjungsan_acctno %>" size="30" readonly>
	</td>
</tr>
<tr height="26">
	<td bgcolor="<%= adminColor("tabletop") %>">�����ָ�</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" class="text_ro" value="<%= ogroup.FOneItem.Fjungsan_acctname %>" size="30" readonly>
	</td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**��ü ���������**</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">����ڸ�</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>" size="20" maxlength="32"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>" size="20" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>" size="25" maxlength="64"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>" size="20" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" height="30">���&CS ����ڸ�</td>
	<td bgcolor="#FFFFFF" colspan="3">��۴���� �Ǵ� CS����� ������ �Ʒ� �귣�� ��ǰ�������� ���� �����մϴ�.</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�������ڸ�</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>" size="20" maxlength="32"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>" size="20" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>" size="25" maxlength="64"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>" size="20" maxlength="16"></td>
</tr>
</form>

<tr align="center" height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="button" value="��ü���� ����" onclick="SaveUpcheInfo(frmupche);">
	</td>
</tr>

</table>

<br>
<% if (opartner.FOneItem.Fuserdiv="14") then %>
<table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="icon" value="#">
		<font color="red"><b>�귣�� ������ ����</b></font>
		(��ǰ���� �������� �޶��� �� �ֽ��ϴ�.)
	</td>
</tr>
<tr height="25">
	<td width="100" bgcolor="<%= adminColor("pink") %>">�귣��ID</td>
	<td bgcolor="#FFFFFF" width="200"><%= opartner.FOneItem.FID %></td>
	<td width="100" bgcolor="<%= adminColor("pink") %>"></td>
	<td bgcolor="#FFFFFF"></td>
</tr>
<% if (opartner.FOneItem.Fdiy_yn="Y") then %>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >��ǰ�⺻����</td>
	<td bgcolor="#FFFFFF" >
		<%= opartner.FOneItem.GetMWUName %>&nbsp;
		<%= opartner.FOneItem.Fdiy_margin %> %
		&nbsp;&nbsp;(�ΰ�������)
	</td>
	<td bgcolor="<%= adminColor("pink") %>" >������ </td>
	<td bgcolor="#FFFFFF" >
	<%= opartner.FOneItem.Fjungsan_date %>
	</td>
</tr>
<% end if %>
<% if (opartner.FOneItem.Flec_yn="Y") then %>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >���±⺻����</td>
	<td bgcolor="#FFFFFF" >
		���� : <%= opartner.FOneItem.Flec_margin %> %
		���� : <%= opartner.FOneItem.Fmat_margin %> %
	</td>
	<td bgcolor="<%= adminColor("pink") %>" >������ </td>
	<td bgcolor="#FFFFFF" >
	<%= opartner.FOneItem.Fjungsan_date %>
	</td>
</tr>
<% end if %>
</table>
<% else %>
<table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="icon" value="#">
		<font color="red"><b>�귣�� ������ ����</b></font>
		(��ǰ���� �������� �޶��� �� �ֽ��ϴ�.)
	</td>
</tr>
<tr height="25">
	<td width="100" bgcolor="<%= adminColor("pink") %>">�귣��ID</td>
	<td bgcolor="#FFFFFF" width="200"><%= opartner.FOneItem.FID %></td>
	<td width="100" bgcolor="<%= adminColor("pink") %>">��ǰ��</td>
	<td bgcolor="#FFFFFF"><%= opartner.FOneItem.FTotalitemcount %></td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >�¶��α⺻����</td>
	<td bgcolor="#FFFFFF" >
		<%= opartner.FOneItem.GetMWUName %>&nbsp;
		<%= opartner.FOneItem.Fdefaultmargine %> %
	</td>
	<td bgcolor="<%= adminColor("pink") %>" >������ </td>
	<td bgcolor="#FFFFFF" >
	<%= opartner.FOneItem.Fjungsan_date %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >��������(������)</td>
	<td bgcolor="#FFFFFF" >
		<table border=0 cellspacing=0 cellpadding=0 class=a>
			<tr>
				<td width="90"><b>��������ǥ</b></td>
				<td width="80"><%= ooffontract.GetSpecialChargeDivName("streetshop000") %></td>
				<td><%= ooffontract.GetSpecialDefaultMargin("streetshop000") %> %</td>
			</tr>
			<% for i=0 to ooffontract.FResultCount-1 %>
			<% if (ooffontract.FItemList(i).Fshopdiv="1")  then %>
			<tr>
				<td><%= ooffontract.FItemList(i).Fshopname %></td>
				<td><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
				<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
			</tr>
			<% end if %>
			<% next %>
		</table>
	</td>
	<td bgcolor="<%= adminColor("pink") %>" >������ </td>
	<td bgcolor="#FFFFFF"><%= opartner.FOneItem.Fjungsan_date_off %></td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >��������(������)</td>
	<td bgcolor="#FFFFFF" >
		<table border=0 cellspacing=0 cellpadding=0 class=a>
			<tr>
				<td width="90"><b>����������ǥ</b></td>
				<td width="80"><%= ooffontract.GetSpecialChargeDivName("streetshop800") %></td>
				<td><%= ooffontract.GetSpecialDefaultMargin("streetshop800") %> %</td>
			</tr>
			<% for i=0 to ooffontract.FResultCount-1 %>
			<% if (ooffontract.FItemList(i).Fshopdiv="3")  then %>
			<tr>
				<td><%= ooffontract.FItemList(i).Fshopname %></td>
				<td><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
				<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
			</tr>
			<% end if %>
			<% next %>

			<% for i=0 to ooffontract.FResultCount-1 %>
			<% if (ooffontract.FItemList(i).Fshopdiv="5")  then %>
			<tr>
				<td><%= ooffontract.FItemList(i).Fshopname %></td>
				<td><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
				<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
			</tr>
			<% end if %>
			<% next %>
		</table>
	</td>
	<td bgcolor="<%= adminColor("pink") %>">������ </td>
	<td bgcolor="#FFFFFF"><%= opartner.FOneItem.Fjungsan_date_frn %></td>
</tr>
</table>
<% end if %>
<br>

<table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmbrand" method="post" action="doupcheedit.asp">
<input type="hidden" name="mode" value="brandedit">
<input type="hidden" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="icon" value="#">
		<font color="red"><b>�귣�� ��۴�� �� ��ǰ�ּ�</b></font>
		(��ü��ۻ�ǰ�� ��� �Ʒ� ��ǰ�ּҸ� ���Բ� �ȳ��� �帳�ϴ�.)
	</td>
</tr>
<tr height="25">
	<td width="100" bgcolor="<%= adminColor("pink") %>">��۴����</td>
	<td bgcolor="#FFFFFF" width="200"><input type="text" class="text" name="deliver_name" value="<%= OReturnAddr.FreturnName %>" size="16" maxlength="16"></td>
	<td width="100" bgcolor="<%= adminColor("pink") %>">��ȭ��ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_phone" value="<%= OReturnAddr.FreturnPhone %>" size="16" maxlength="16"></td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>">�ڵ�����ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_hp" value="<%= OReturnAddr.Freturnhp %>" size="16" maxlength="16"></td>
	<td bgcolor="<%= adminColor("pink") %>">�̸����ּ�</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_email" value="<%= OReturnAddr.FreturnEmail %>" size="16" maxlength="128"></td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >��ǰ�ּ�</td>
	<td bgcolor="#FFFFFF" colspan=3>
		<input type="text" class="text" name="return_zipcode" value="<%= OReturnAddr.FreturnZipcode %>" size="7" maxlength="7">
		<input type="button" class="button" value="�˻�" onClick="TnFindZipNewdesigner('frmbrand','D')">
		<input type="button" class="button" value="�˻�(��)" onclick="javascript:popZip('b');"><br>
		<input type="text" class="text" name="return_address" value="<%= OReturnAddr.FreturnZipaddr %>" size="26" maxlength="64">&nbsp;
		<input type="text" class="text" name="return_address2" value="<%= OReturnAddr.FreturnEtcaddr %>" size="38" maxlength="64">
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >�ù��</td>
	<td bgcolor="#FFFFFF" colspan=3>
		<% drawSelectBoxDeliverCompany "defaultsongjangdiv" , OReturnAddr.Fsongjangdiv %>
	</td>
</tr>
<!--
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" ></td>
	<td bgcolor="#FFFFFF" colspan=3>
		<input type="checkbox" class="checkbox" name=applyallbrand value="Y"> ��ü�� ��� �귣�� �ϰ� ����
	</td>
</tr>
-->
</form>

<tr align="center" height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="button" value="�귣�� ��ǰ���� ����" onclick="SaveBrandReturnInfo(frmbrand);">
	</td>
</tr>
</table>

<br>
<% if (opartner.FOneItem.Fuserdiv="14") then %>
<!-- ǥ�� ���� 2016/08/22-->
<% else %>
<!-- <table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"> -->
<!-- <form name="frmetc" method="post" action="<%= uploadImgUrl %>/linkweb/partner_info/doprofileimageadmin.asp" enctype="multipart/form-data"> -->
<!-- <input type="hidden" name="designerid" value="<%= opartner.FOneItem.FID %>"> -->
<!-- <tr height="25" bgcolor="FFFFFF"> -->
<!-- 	<td colspan="4"> -->
<!-- 		<input type="button" class="icon" value="#"> -->
<!-- 		<font color="red"><b>�귣�� ��������(��ǥ������)</b></font> -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr> -->
<!-- 	<td width="120" bgcolor="<%= adminColor("sky") %>">�귣���(�ѱ�)</td> -->
<!-- 	<td width="180" bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" class="text" name="socname_kor" value="<%= opartner.FOneItem.Fsocname_kor %>" size=20 maxlength="20"> -->
<!-- 	</td> -->
<!-- 	<td width="120" bgcolor="<%= adminColor("sky") %>">�귣���(����)</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" class="text" name="socname" value="<%= opartner.FOneItem.Fsocname %>" size=20 maxlength="20"> -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr> -->
<!-- 	<td bgcolor="<%= adminColor("sky") %>">�ΰ�</td> -->
<!-- 	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- 		<img name="logoimg" src="<%= opartner.FOneItem.getSocLogoUrl %>" width=150 height=100><br> -->
<!-- 		(�귣�� �ΰ�� 150x100 �ȼ��� ���ε� ���ֽñ� �ٶ��ϴ�.)<br> -->
<!-- 		<input type="file" class="file" name="file1" size="40"><!--  onchange="ChangeLogo(this,frmetc.logoimg);" -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr> -->
<!-- 	<td bgcolor="<%= adminColor("sky") %>" >���</td> -->
<!-- 	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- 	<img name=titleimg src="<%= opartner.FOneItem.getTitleImgUrl %>" width=300 height=75><br> -->
<!-- 	(�̹����� 720x220 �ȼ��� ���ε� ���ֽñ� �ٶ��ϴ�.)<br> -->
<!-- 	<input type=file name=file2 size=40><!--  onchange="ChangeTitle(this,frmetc.titleimg);" --> 
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr> -->
<!-- 	<td bgcolor="<%= adminColor("sky") %>">�귣��<br>�ڸ�Ʈ</td> -->
<!-- 	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- 	<textarea class="textarea" name="dgncomment" cols="80" rows="6"><%= opartner.FOneItem.Fdgncomment %></textarea> -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- </form> -->
<!--  -->
<!-- <tr align="center" height="25" bgcolor="FFFFFF"> -->
<!-- 	<td colspan="15"> -->
<!-- 		<input type="button" class="button" value="�귣�� ���� ����" onclick="SaveBrandEtcInfo(frmetc);"> -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- </table> -->
<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmetc" method="post" action="<%= uploadImgUrl %>/linkweb/partner_info/doprofileimageadmin.asp" enctype="multipart/form-data">
<input type=hidden name=designerid value="<%= opartner.FOneItem.FID %>">
<tr>
	<td width="120" bgcolor="<%= adminColor("sky") %>">�귣���(�ѱ�)</td>
	<td width="180" bgcolor="#FFFFFF">
		<input type="text" class="text" name="socname_kor" value="<%= opartner.FOneItem.Fsocname_kor %>" size=20 maxlength="20">
	</td>
	<td width="120" bgcolor="<%= adminColor("sky") %>">�귣���(����)</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="socname" value="<%= opartner.FOneItem.Fsocname %>" size=20 maxlength="20">
	</td>
</tr>
<% if (FALSE) then %>
<tr>
	<td bgcolor="<%= adminColor("sky") %>" >������ �̹���</td>
	<td bgcolor="#FFFFFF" colspan="3">
	<img name=brandimg src="<%= opartner.FOneItem.getBrandImgUrl("") %>" width=<%=600/2%> height=<%=600/2%>><br>
	(������ �̹����� 600X600 �ȼ��̻����� �����Ǿ� �ֽ��ϴ�.)<br>
	<img src="<%= opartner.FOneItem.getBrandImgUrl("1") %>" width=<%=400/2%> height=<%=400/2%>>&nbsp;
	<img src="<%= opartner.FOneItem.getBrandImgUrl("2") %>" width=<%=200/2%> height=<%=200/2%>>&nbsp;
	<img src="<%= opartner.FOneItem.getBrandImgUrl("3") %>" width=<%=100/2%> height=<%=100/2%>>&nbsp;<br/>
	<input type="file" class="button" name="file4" size="60" onclick="ChangeTitle(this,frmetc.brandimg);">
	<% If opartner.FOneItem.getBrandImgUrl("") <> "http://webimage.10x10.co.kr/image/brandlogo/" Then %>
		<input type="checkbox" name="deltitleimg" size="60" value="Y">����
	<% End If %>
	</td>
</tr>
<% end if %>
<tr>
	<td bgcolor="<%= adminColor("sky") %>">�귣��<br>�ڸ�Ʈ</td>
	<td bgcolor="#FFFFFF" colspan="3">
	<textarea class="textarea" name="dgncomment" cols="80" rows="6"><%= opartner.FOneItem.Fdgncomment %></textarea>
	</td>
</tr>
<tr>
	<td colspan="4" align=center bgcolor="#FFFFFF"><input type="button" class="button" value="�귣�� ��Ÿ���� ����" onclick="SaveBrandEtcInfo(frmetc);"></td>
</tr>
</form>
</table>
<br>
<% end if %>

<table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmpass" method=post action="doupcheedit.asp">
<input type="hidden" name="mode" value="editpass">
<input type="hidden" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="2">
		<input type="button" class="icon" value="#">
		<font color="red"><b>��й�ȣ ����</b></font>
		&nbsp;
		(�귣�� ��й�ȣ�� �����Ͻ÷��� �Ʒ� ���� ä�� �ֽñ�ٶ��ϴ�.)
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">������й�ȣ</td>
	<td width="480" bgcolor="#FFFFFF">
		1�� :<input type="password" class="text" name="txoldpassword" size="12" value="" maxlength="32">
		2�� :<input type="password" class="text" name="txoldpasswordS" size="12" value="" maxlength="32">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�����й�ȣ 1�� </td>
	<td width="480" bgcolor="#FFFFFF">
		�Է�: <input type="password" class="text" name="txnewpassword1" size="12" value="" maxlength="32"><br>
		Ȯ��: <input type="password" class="text" name="txnewpassword2" size="12" value="" maxlength="32">
	</td>
</tr> 
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�����й�ȣ 2�� </td>
	<td width="480" bgcolor="#FFFFFF">
		�Է�: <input type="password" class="text" name="txnewpasswordS1" size="12" value="" maxlength="32"><br>
		Ȯ��: <input type="password" class="text" name="txnewpasswordS2" size="12" value="" maxlength="32">
	</td>
</tr>
</form>

<tr align="center" height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="button" value="�귣�� ��й�ȣ����" onclick="EditPass(frmpass);">
	</td>
</tr>
</table>

<%
set ogroup = Nothing
set opartner = Nothing
set ooffontract = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
