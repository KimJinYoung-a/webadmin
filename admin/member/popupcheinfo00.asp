<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
response.write "������. ���������� ���ϰ�� �������� �Ű��ϼ���."
response.end
dim ogroup,opartner,i
dim designer
dim groupid
dim oJungsanDiff, j

designer = requestCheckvar(request("designer"),32)

set opartner = new CPartnerUser
opartner.FRectDesignerID = designer
opartner.GetOnePartnerNUser

set ogroup = new CPartnerGroup
if designer<>"" then
ogroup.FRectGroupid = opartner.FOneItem.FGroupid
end if
ogroup.GetOneGroupInfo

set oJungsanDiff = new CPartnerGroup
if (designer<>"") then
if (opartner.FOneItem.FGroupid<>"") then
    oJungsanDiff.FRectGroupid = opartner.FOneItem.FGroupid
	oJungsanDiff.GetGroupPartnerJungsanDiffList
end if
end if
%>
<script language='javascript'>
function SaveBrandInfo(frm){
//alert('���� ������� �ʴ� �޴��Դϴ�. - ���� ���� ��');
//return;
	if (frm.prtidx.value.length<1){
		alert('�� ��ȣ�� �Է��ϼ���. - [�⺻�� 9999]');
		frm.prtidx.focus();
		return;
	}

	if (frm.password.value.length<1){
		alert('�귣�� �н����带 �Է��ϼ���.');
		frm.password.focus();
		return;
	}

	if (frm.socname_kor.value.length<1){
		alert('��Ʈ��Ʈ��(�ѱ�)�� �Է��ϼ���.');
		frm.socname_kor.focus();
		return;
	}

	if (frm.socname.value.length<1){
		alert('��Ʈ��Ʈ��(����)�� �Է��ϼ���.');
		frm.socname.focus();
		return;
	}

	if ((!frm.isusing[0].checked)&&(!frm.isusing[1].checked)){
		alert('��뿩�θ� �����ϼ���.');
		frm.isusing[0].focus();
		return;
	}

	if ((!frm.isextusing[0].checked)&&(!frm.isextusing[1].checked)){
		alert('���޸� ��뿩�θ� �����ϼ���.');
		frm.isextusing[0].focus();
		return;
	}

	if ((!frm.streetusing[0].checked)&&(!frm.streetusing[1].checked)){
		alert('��Ʈ��Ʈ ��뿩�θ� �����ϼ���.');
		frm.streetusing[0].focus();
		return;
	}

	if ((!frm.extstreetusing[0].checked)&&(!frm.extstreetusing[1].checked)){
		alert('���޸� ��Ʈ��Ʈ ��뿩�θ� �����ϼ���.');
		frm.extstreetusing[0].focus();
		return;
	}

	if ((!frm.specialbrand[0].checked)&&(!frm.specialbrand[1].checked)){
		alert('Ŀ�´�Ƽ ��뿩�θ� �����ϼ���.');
		frm.specialbrand[0].focus();
		return;
	}

	if (frm.userdiv.value.length<1){
		alert('�귣�� ������ �����ϼ���.');
		frm.userdiv.focus();
		return;
	}

	if (frm.maeipdiv.value.length<1){
		alert('���� ������ �����ϼ���.');
		frm.maeipdiv.focus();��
		return;
	}

	if (!IsDouble(frm.defaultmargine.value)){
		alert('�⺻������ �Է��ϼ���. - �Ǽ��� �����մϴ�.');
		frm.defaultmargine.focus();
		return;
	}

	if (frm.jungsan_date.value.length<1){
		alert('�������� �����ϼ���.');
		frm.jungsan_date.focus();
		return;
	}

	if (frm.mduserid.value.length<1){
		alert('����ڸ� �����ϼ���. - �ʼ� �����Դϴ�.');
		frm.mduserid.focus();
		return;
	}

	var ret = confirm('�귣�� ������ ���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}

function SaveUpcheInfo(frm){
    //2016/12/14 �������.
alert('���� ������� �ʴ� �޴��Դϴ�. - ���� ���� ��');
return;

	if (frm.groupid.value.length<1){
		var ret = confirm('��ü ������ ���� �Ͻðڽ��ϱ�?');
	}else{
		var ret = confirm('���� �׷��ڵ忡 �ִ� ���� ��ü ������ �����˴ϴ�. ���� �Ͻðڽ��ϱ�?');
	}

	if (ret){
		frm.submit();
	}
}

function ModiInfo(frm){
//alert('���� ������� �ʴ� �޴��Դϴ�. - ���� ���� ��');
//return;
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		//frm.submit();
	}

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

function PopUpcheList(frmname){
	var popwin = window.open("/admin/lib/popupchelist.asp?frmname=" + frmname,"popupchelist","width=700 height=540 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>


<table width="600" align="center" border="0" cellpadding="3" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr>
		<td class="a">
		�귣�� :<%	drawSelectBoxDesignerWithName "designer", designer %>
		</td>
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<% if designer<>"" then %>
<form name=frmbuf>
<input type=hidden name=company_name value="<%= opartner.FOneItem.FCompany_name %>">
<input type=hidden name=ceoname value="<%= opartner.FOneItem.Fceoname %>">
<input type=hidden name=company_no value="<%= opartner.FOneItem.Fcompany_no %>">
<input type=hidden name=jungsan_gubun value="<%= opartner.FOneItem.Fjungsan_gubun %>">
<input type=hidden name=company_zipcode value="<%= opartner.FOneItem.Fzipcode %>">
<input type=hidden name=company_address value="<%= opartner.FOneItem.Faddress %>">
<input type=hidden name=company_address2 value="<%= opartner.FOneItem.Fmanager_address %>">
<input type=hidden name=company_uptae value="<%= opartner.FOneItem.Fcompany_uptae %>">
<input type=hidden name=company_upjong value="<%= opartner.FOneItem.Fcompany_upjong %>">
<input type=hidden name=company_tel value="<%= opartner.FOneItem.Ftel %>">
<input type=hidden name=company_fax value="<%= opartner.FOneItem.Ffax %>">

<input type=hidden name=jungsan_bank value="<%= opartner.FOneItem.Fjungsan_bank %>">
<input type=hidden name=jungsan_acctno value="<%= opartner.FOneItem.Fjungsan_acctno %>">
<input type=hidden name=jungsan_acctname value="<%= opartner.FOneItem.Fjungsan_acctname %>">
<input type=hidden name=manager_name value="<%= opartner.FOneItem.Fmanager_name %>">
<input type=hidden name=manager_phone value="<%= opartner.FOneItem.Fmanager_phone %>">
<input type=hidden name=manager_email value="<%= opartner.FOneItem.Femail %>">
<input type=hidden name=manager_hp value="<%= opartner.FOneItem.Fmanager_hp %>">

<input type=hidden name=deliver_name value="<%= opartner.FOneItem.Fdeliver_name %>">
<input type=hidden name=deliver_phone value="<%= opartner.FOneItem.Fdeliver_phone %>">
<input type=hidden name=deliver_email value="<%= opartner.FOneItem.Fdeliver_email %>">
<input type=hidden name=deliver_hp value="<%= opartner.FOneItem.Fdeliver_hp %>">

<input type=hidden name=jungsan_name value="<%= opartner.FOneItem.Fjungsan_name %>">
<input type=hidden name=jungsan_phone value="<%= opartner.FOneItem.Fjungsan_phone %>">
<input type=hidden name=jungsan_email value="<%= opartner.FOneItem.Fjungsan_email %>">
<input type=hidden name=jungsan_hp value="<%= opartner.FOneItem.Fjungsan_hp %>">


</form>

<table width="600" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor=#000000>
<form name="frmupche" method="post" action="/admin/member/doupcheedit.asp">
<input type="hidden" name="mode" value="groupedit">
<input type="hidden" name="uid" value="<%= designer %>">
	<tr bgcolor="#DDDDFF">
		<td colspan=4><b>1.��ü��������</b></td>
		<!--<td colspan=2><<span align=right><input type=button value="�ӽ� - �귣�� �������� ����" onclick="CopyFromBrandInfo()"></span>-->
		</td>
	</tr>
	<tr >
		<td width="100" bgcolor="#DDDDFF">��ü�ڵ�</td>
		<td bgcolor="#FFFFFF" width="200">
		<input type="text" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" size="7" maxlength="5" style="background-color:#EEEEEE;" readonly>
		<input type="button" value="��ü����" onClick="PopUpcheSelect('frmupche');">
		</td>
		<td width="100" bgcolor="#DDDDFF">��ü��</td>
		<td bgcolor="#FFFFFF" width="200">
		<%= ogroup.FOneItem.FCompany_name %>
		</td>
	</tr>
	<tr >
		<td width="100" bgcolor="#DDDDFF">�����귣��ID</td>
		<td colspan="3" bgcolor="#FFFFFF"><%= ogroup.FOneItem.getBrandListHTML %></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**����ڵ������**</td>
	</tr>

	<tr>
		<td width="100" bgcolor="#DDDDFF">ȸ���(��ȣ)</td>
		<td bgcolor="#FFFFFF"><input type="text" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="24" maxlength="32"></td>
		<td width="100" bgcolor="#DDDDFF">��ǥ��</td>
		<td bgcolor="#FFFFFF"><input type="text" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td width="100" bgcolor="#DDDDFF">����ڹ�ȣ</td>
		<td bgcolor="#FFFFFF"><input type="text" name="company_no" value="<%= ogroup.FOneItem.Fcompany_no %>" size="16" maxlength="20"></td>
		<td width="100" bgcolor="#DDDDFF">��������</td>
		<td bgcolor="#FFFFFF">
			<select name="jungsan_gubun">
			<option value="�Ϲݰ���" <% if ogroup.FOneItem.Fjungsan_gubun="�Ϲݰ���" then response.write "selected" %> >�Ϲݰ���</option>
			<option value="���̰���" <% if ogroup.FOneItem.Fjungsan_gubun="���̰���" then response.write "selected" %> >���̰���</option>
			<option value="��õ¡��" <% if ogroup.FOneItem.Fjungsan_gubun="��õ¡��" then response.write "selected" %> >��õ¡��</option>
			<option value="�鼼" <% if ogroup.FOneItem.Fjungsan_gubun="�鼼" then response.write "selected" %> >�鼼</option>
			<option value="����(�ؿ�)" <% if ogroup.FOneItem.Fjungsan_gubun="����(�ؿ�)" then response.write "selected" %> >����(�ؿ�)</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">����������</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" name="company_zipcode" value="<%= ogroup.FOneItem.Fcompany_zipcode %>" size="7" maxlength="7">(000-000)<br>
			<input type="text" name="company_address" value="<%= ogroup.FOneItem.Fcompany_address %>" size="16" maxlength="64">&nbsp;
			<input type="text" name="company_address2" value="<%= ogroup.FOneItem.Fcompany_address2 %>" size="28" maxlength="64">
		</td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">����</td>
		<td bgcolor="#FFFFFF"><input type="text" name="company_uptae" value="<%= ogroup.FOneItem.Fcompany_uptae %>" size="24" maxlength="32"></td>
		<td bgcolor="#DDDDFF">����</td>
		<td bgcolor="#FFFFFF"><input type="text" name="company_upjong" value="<%= ogroup.FOneItem.Fcompany_upjong %>" size="24" maxlength="32"></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**��ü�⺻����**</td>
	</tr>

	<tr>
		<td bgcolor="#DDDDFF">��ǥ��ȭ</td>
		<td bgcolor="#FFFFFF"><input type="text" name="company_tel" value="<%= ogroup.FOneItem.Fcompany_tel %>" size="16" maxlength="16"></td>
		<td bgcolor="#DDDDFF">�ѽ�</td>
		<td bgcolor="#FFFFFF"><input type="text" name="company_fax" value="<%= ogroup.FOneItem.Fcompany_fax %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">��ǰ �ּ�</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>" size="7" maxlength="7">(000-000)<br>
			<input type="text" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>" size="16" maxlength="64">&nbsp;
			<input type="text" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>" size="28" maxlength="64">
		</td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**������������**</td>
	</tr>

	<tr>
		<td bgcolor="#DDDDFF">�ŷ�����</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		<% DrawBankCombo "jungsan_bank", ogroup.FOneItem.Fjungsan_bank %>
		</td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">���¹�ȣ</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" name="jungsan_acctno" value="<%= ogroup.FOneItem.Fjungsan_acctno %>" size="16" maxlength="32">
		&nbsp;&nbsp; '-'�� ���� ��ȣ�� �Է����ֽñ� �ٶ��ϴ�.
		</td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">�����ָ�</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" name="jungsan_acctname" value="<%= ogroup.FOneItem.Fjungsan_acctname %>" size="24" maxlength="16">
		&nbsp;&nbsp; ���� ���� ���ñ� �ٶ��ϴ�.
		</td>
	</tr>
    <tr>
		<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		�¶��� : <% DrawJungsanDateCombo "jungsan_date", ogroup.FOneItem.Fjungsan_date %>
		&nbsp;
		�������� : <% DrawJungsanDateCombo "jungsan_date_off", ogroup.FOneItem.Fjungsan_date_off %>
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
	<% if (session("ssBctId") = "icommang") or (session("ssBctId") = "coolhas") or (session("ssBctId") = "aimcta") or (session("ssBctId") = "tozzinet") or (session("ssAdminPsn") = "8") then %>
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
	<% if (session("ssBctId") = "icommang") or (session("ssBctId") = "coolhas") or (session("ssBctId") = "tozzinet") or (session("ssAdminPsn") = "8") then %>
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
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**���������**</td>
	</tr>

	<tr>
		<td bgcolor="#DDDDFF">����ڸ�</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>" size="24" maxlength="32"></td>
		<td bgcolor="#DDDDFF">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>" size="24" maxlength="64"></td>
		<td bgcolor="#DDDDFF">�ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">��۴���ڸ�</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="deliver_name" value="<%= ogroup.FOneItem.Fdeliver_name %>" size="24" maxlength="32"></td>
		<td bgcolor="#DDDDFF">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="deliver_phone" value="<%= ogroup.FOneItem.Fdeliver_phone %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" name="deliver_email" value="<%= ogroup.FOneItem.Fdeliver_email %>" size="24" maxlength="64"></td>
		<td bgcolor="#DDDDFF">�ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" name="deliver_hp" value="<%= ogroup.FOneItem.Fdeliver_hp %>" size="16" maxlength="16"></td>
	</tr>

	<tr>
		<td width="80" bgcolor="#DDDDFF">�������ڸ�</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>" size="24" maxlength="32"></td>
		<td width="80" bgcolor="#DDDDFF">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td width="60" bgcolor="#DDDDFF">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>" size="24" maxlength="64"></td>
		<td width="60" bgcolor="#DDDDFF">�ڵ���</td>
		<td bgcolor="#FFFFFF"><input type="text" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" value="��ü���� ����" onclick="SaveUpcheInfo(frmupche);"></td>
	</tr>
</form>
</table>

<br>
<table width="600" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor=#000000>
<form name="frmbrand" method="post" action="doupcheedit.asp">
<input type="hidden" name="uid" value="<%= opartner.FOneItem.FID %>">
<input type="hidden" name="mode" value="brandedit">
<tr>
	<td bgcolor="#FFDDDD" colspan=6><b>2.�귣���������</b></td>
</tr>
<tr>
	<td width="110" bgcolor="#FFDDDD">ȸ���</td>
	<td bgcolor="#FFFFFF" colspan=2><%= opartner.FOneItem.FCompany_name %></td>
	<td width="110" bgcolor="#FFDDDD" >����ȣ</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="prtidx" value="<%= opartner.FOneItem.getRackCode %>" size="4" maxlength="4">
	(�⺻�� : 9999)</td>
	</td>
</tr>
<tr>
	<td width="110" bgcolor="#FFDDDD">�귣��ID</td>
	<td bgcolor="#FFFFFF" colspan=2><%= opartner.FOneItem.FID %></td>
	<td width="110" bgcolor="#FFDDDD">�н�����</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="password" value="<%= opartner.FOneItem.Fppass %>">
	</td>
</tr>
<tr>
	<td width="110" bgcolor="#FFDDDD">��Ʈ��Ʈ��<br>(�ѱ�)</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="socname_kor" value="<%= opartner.FOneItem.Fsocname_kor %>">
	</td>
	<td width="110" bgcolor="#FFDDDD">��Ʈ��Ʈ��<br>(����)</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="socname" value="<%= opartner.FOneItem.Fsocname %>">
	</td>
</tr>
<tr>
	<td rowspan="3" width="110" bgcolor="#FFDDDD">�귣��<br>��뿩��<br>(ī�װ�����)</td>
	<td bgcolor="#FFFFFF">�ٹ�����</td>
	<td bgcolor="#FFFFFF"><input type=radio name="isusing" value="Y" <% if opartner.FOneItem.Fisusing="Y" then response.write "checked" %> >��� <input type=radio name="isusing" value="N" <% if opartner.FOneItem.Fisusing="N" then response.write "checked" %> >������</td>
	<td rowspan="3" width="110" bgcolor="#FFDDDD">��Ʈ��Ʈ<br>ǥ�ÿ���<br>(�귣������)</td>
	<td bgcolor="#FFFFFF">�ٹ�����</td>
	<td bgcolor="#FFFFFF"><input type=radio name="streetusing" value="Y" <% if opartner.FOneItem.Fstreetusing="Y" then response.write "checked" %> >��� <input type=radio name="streetusing" value="N" <% if opartner.FOneItem.Fstreetusing="N" then response.write "checked" %> >������</td>
</tr>
<tr >
	<td bgcolor="#FFFFFF">���޸�</td>
	<td bgcolor="#FFFFFF"><input type=radio name="isextusing" value="Y" <% if opartner.FOneItem.Fisextusing="Y" then response.write "checked" %> >��� <input type=radio name="isextusing" value="N" <% if opartner.FOneItem.Fisextusing="N" then response.write "checked" %> >������	</td>
	<td bgcolor="#FFFFFF">���޸�</td>
	<td bgcolor="#FFFFFF"><input type=radio name="extstreetusing" value="Y" <% if opartner.FOneItem.Fextstreetusing="Y" then response.write "checked" %> >��� <input type=radio name="extstreetusing" value="N" <% if opartner.FOneItem.Fextstreetusing="N" then response.write "checked" %> >������	</td>
</tr>
<tr >
	<td bgcolor="#FFFFFF" colspan=2></td>
	<td bgcolor="#FFFFFF">Ŀ�´�Ƽ</td>
	<td bgcolor="#FFFFFF"><input type=radio name="specialbrand" value="Y" <% if opartner.FOneItem.Fspecialbrand="Y" then response.write "checked" %>>��� <input type=radio name="specialbrand" value="N" <% if opartner.FOneItem.Fspecialbrand="N" then response.write "checked" %>>������</td>
</tr>
<tr >
	<td bgcolor="#FFDDDD">��ü����</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<% DrawBrandGubunCombo "userdiv", opartner.FOneItem.Fuserdiv %>
	</td>
	<td bgcolor="#FFDDDD">��ǰ��</td>
	<td bgcolor="#FFFFFF" colspan=2><%= opartner.FOneItem.FTotalitemcount %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=6>**�����û���**</td>
</td>
<tr >
	<td width="110" bgcolor="#FFDDDD" >�⺻����</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<% DrawBrandMWUCombo "maeipdiv",opartner.FOneItem.Fmaeipdiv %>
	<input type="text" name="defaultmargine" value="<%= opartner.FOneItem.Fdefaultmargine %>" size="4" style="text-align:right"> %
	</td>
	<td width="110" bgcolor="#FFDDDD" >������ :</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<% DrawJungsanDateCombo "jungsan_date", opartner.FOneItem.Fjungsan_date %>
	</td>
</tr>
<tr>
	<td width="110" bgcolor="#FFDDDD" >���MD</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<% drawSelectBoxCoWorker "mduserid", opartner.FOneItem.Fmduserid %>
	</td>
</tr>
<!--
<% if ogroup.FOneItem.FGroupId<>"" then %>
<tr>
	<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" value="�귣������ ����" onclick="SaveBrandInfo(frmbrand);"></td>
</tr>
<% else %>
<tr>
	<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" value="�귣������ ����" onclick="alert('��ü������ ���� ���� �Ͻ��� �귣�������� ���� �� �� �ֽ��ϴ�.');"></td>
</tr>
<% end if %>
-->
</form>
</table>

<% end if %>
<%
set opartner = Nothing
set oJungsanDiff = Nothing
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
