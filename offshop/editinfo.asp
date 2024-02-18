<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<%
dim opartner,i

set opartner = new CPartnerUser
opartner.FCurrpage = 1
opartner.FRectDesignerID = session("ssBctID")
opartner.FPageSize = 1
opartner.GetOnePartnerNUser 

Dim GroupIdExists 
if IsNULL(opartner.FOneItem.FGroupid) or (opartner.FOneItem.FGroupid="") then
    ''response.write "�׷��ڵ尡 �������� �ʾҽ��ϴ�. �����Ұ�."
    ''response.end
    GroupIdExists = FALSE
ELSE
    GroupIdExists = TRUE
end if

dim ogroup
set ogroup = new CPartnerGroup
ogroup.FRectGroupid = opartner.FOneItem.FGroupid
ogroup.GetOneGroupInfo

    
dim ochargeuser
set ochargeuser = new COffShopChargeUser
ochargeuser.FRectShopID = session("ssBctID")
ochargeuser.GetOffShopList


%>
<script language='javascript'>
function ModiInfo(frm){
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
    
    if (frm.shopname.value.length<1){
		alert('������� �Է��ϼ���.');
		return;
	}
	
	if (frm.return_zipcode.value.length<1){
		alert('�繫���ּ�(��ǰ�ּ�) �����ȣ�� �����ϼ���.');
		frm.return_zipcode.focus();
		return;
	}

	if (frm.return_address.value.length<1){
		alert('�繫���ּ�1(��ǰ �ּ�1)�� �Է��ϼ���.');
		frm.return_address.focus();
		return;
	}

	if (frm.return_address2.value.length<1){
		alert('�繫���ּ�2(��ǰ �ּ�2)�� �Է��ϼ���.');
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
	
	var ret = confirm('��� ��ȣ�� ���� �Ͻðڽ��ϱ�?\r\nPOS �α��� ��й�ȣ�� SCM�α��� ��й�ȣ�� ���ÿ� ����˴ϴ�.');
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
</script>

<table width="600" cellspacing="1" class="a" bgcolor=#3d3d3d>
<% if opartner.FresultCount >0 then %>


<table width="600" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmupche" method="post" action="shopinfoedit_process.asp">
	<input type="hidden" name="mode" value="groupedit">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="icon" value="#">
			<font color="red"><b>����� ����</b></font>
		</td>
	</tr>
	<% IF (Not GroupIdExists) THEN %>
	<tr height="25" bgcolor="FFFFFF">
	    <td> �׷� �ڵ尡 �������� �ʾҽ��ϴ�. �������� ���� �Ұ�</td>
	</tr>
	<% ELSE %>
	<tr>
		<td width="120" bgcolor="<%= adminColor("tabletop") %>">��ü�ڵ�</td>
		<td bgcolor="#FFFFFF" width="180">
			<input type="text" class="text_ro" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" size="7" maxlength="5" readonly>
		</td>
		<td width="120" bgcolor="<%= adminColor("tabletop") %>">��ü��</td>
		<td bgcolor="#FFFFFF" width="180">
			<%= ogroup.FOneItem.FCompany_name %>
		</td>
	</tr>
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**����ڵ������**</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">ȸ���(��ȣ)</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="30" readonly>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ��</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="30" readonly>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����ڹ�ȣ</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="company_no" value="<%= ogroup.FOneItem.Fcompany_no %>" size="30" readonly>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="" value="<%= ogroup.FOneItem.Fjungsan_gubun %>" size="30" readonly>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����������</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="company_zipcode" value="<%= ogroup.FOneItem.Fcompany_zipcode %>" size="7" maxlength="7">
			<input type="button" class="button" value="�����ȣ�˻�" onclick="javascript:popZip('s');"><br>
			<input type="text" class="text" name="company_address" value="<%= ogroup.FOneItem.Fcompany_address %>" size="26" maxlength="64">&nbsp;
			<input type="text" class="text" name="company_address2" value="<%= ogroup.FOneItem.Fcompany_address2 %>" size="38" maxlength="64">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_uptae" value="<%= ogroup.FOneItem.Fcompany_uptae %>" size="28" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_upjong" value="<%= ogroup.FOneItem.Fcompany_upjong %>" size="28" maxlength="32"></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**���� �⺻����** &nbsp;&nbsp;</td>
	</tr>

    <tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td bgcolor="#FFFFFF" colspan="3"><input type="text" class="text" name="shopname" value="<%= ochargeuser.FItemList(0).Fshopname %>" size="20" maxlength="64"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ��ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="<%= ogroup.FOneItem.Fcompany_tel %>" size="16" maxlength="16"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ѽ�</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="<%= ogroup.FOneItem.Fcompany_fax %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�繫���ּ�<br>(������ּ�)</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>" size="7" maxlength="7">
			<input type="button" class="button" value="�����ȣ�˻�" onclick="javascript:popZip('m');">
			<input type="checkbox" class="checkbox" name=samezip onclick="SameReturnAddr(this.checked)">��<br>
			<input type="text" class="text" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>" size="26" maxlength="64">&nbsp;
			<input type="text" class="text" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>" size="38" maxlength="64">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�ŷ��ù��</td>
		<td colspan="3" bgcolor="#FFFFFF"><% drawSelectBoxDeliverCompany "defaultsongjangdiv" , opartner.FOneItem.Fdefaultsongjangdiv %>
		</td>
	</tr>
	<% if (FALSE) then %>
    <!--
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**������������** &nbsp;&nbsp;(�������� ���� ������ ���MD���� �����Ͻñ� �ٶ��ϴ�.)</td>
	</tr>

	<tr height="26">
		<td bgcolor="<%= adminColor("tabletop") %>">�ŷ�����</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" class="text_ro" value="<%= ogroup.FOneItem.Fjungsan_bank %>" size="30" readonly>
		<% if (IsNULL(ogroup.FOneItem.Fjungsan_acctno)) or (ogroup.FOneItem.Fjungsan_acctno="") then %>
		(���� ���¸� ����Ͻ÷���  ��� MD���� Fax�� ���� �纻�� �����ֽñ� �ٶ��ϴ�.)
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
    -->
    <% end if %>
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**���������**</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����ڸ�</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>" size="30" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>" size="30" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��۴���ڸ�</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_name" value="<%= ogroup.FOneItem.Fdeliver_name %>" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_phone" value="<%= ogroup.FOneItem.Fdeliver_phone %>" size="30" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_email" value="<%= ogroup.FOneItem.Fdeliver_email %>" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_hp" value="<%= ogroup.FOneItem.Fdeliver_hp %>" size="30" maxlength="16"></td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�������ڸ�</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>" size="30" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>" size="30" maxlength="16"></td>
	</tr>
	
	<tr align="center" height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="button" value="��ü���� ����" onclick="ModiInfo(frmupche);">
		</td>
	</tr>
	<% END IF %>
	</form>
</table>

<br>

<p>

<table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmpass" method=post action="shopinfoedit_process.asp">
	<input type="hidden" name="mode" value="editpass">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2">
			<input type="button" class="icon" value="#">
			<font color="red"><b>��й�ȣ ����</b></font>
			&nbsp;
			(��й�ȣ�� �����Ͻ÷��� �Ʒ� ���� ä�� �ֽñ�ٶ��ϴ�.)
		</td>
	</tr>
	<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">������й�ȣ</td>
    	<td width="480" bgcolor="#FFFFFF">
    		1�� :<input type="password" class="text" name="txoldpassword" size="12" value="" maxlength="16">
    		2�� :<input type="password" class="text" name="txoldpasswordS" size="12" value="" maxlength="16">
    	</td>
    </tr>
	<tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">�����й�ȣ 1�� </td>
    	<td width="480" bgcolor="#FFFFFF">
    		�Է�: <input type="password" class="text" name="txnewpassword1" size="12" value="" maxlength="16"><br>
    		Ȯ��: <input type="password" class="text" name="txnewpassword2" size="12" value="" maxlength="16">
    	</td>
    </tr> 
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">�����й�ȣ 2�� </td>
    	<td width="480" bgcolor="#FFFFFF">
    		�Է�: <input type="password" class="text" name="txnewpasswordS1" size="12" value="" maxlength="16"><br>
    		Ȯ��: <input type="password" class="text" name="txnewpasswordS2" size="12" value="" maxlength="16">
    	</td>
    </tr>
	</form>
	
	<tr align="center" height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="button" value="��й�ȣ����" onclick="EditPass(frmpass);">
		</td>
	</tr>
</table>


<% end if %>
<%
set ochargeuser = Nothing
set ogroup = Nothing
set opartner = Nothing
%>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->