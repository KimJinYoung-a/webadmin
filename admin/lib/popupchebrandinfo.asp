<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->

��� ���ϴ� �޴� �Դϴ�.<br>
<br>
�� �������� ���ϰ�� �����ڿ��� �Ű� �ϼ���. <br>
� �޴����� Ŭ���� ���������� �Ǵ� � �׼��߿� ���� ����������.
<br>

<br><br>
<a href="/admin/member/popupchebrandinfo.asp?designer=<%= request("designer") %>"><font color="blue">���ο� �޴��� �̵� &gt;&gt;</font></a>

<br><br>
<font color="#999999">�ű����� Fnc :  javascript:PopUpcheBrandInfoEdit('makerid')</font> <br>
<%
'' ��� ���ϴ� �޴�
dbget.close()	:	response.End

dim ogroup,opartner,i
dim designer
dim groupid

designer = request("designer")

set opartner = new CPartnerUser
opartner.FRectDesignerID = designer
opartner.GetOnePartnerNUser


set ogroup = new CPartnerGroup
ogroup.FRectGroupid = opartner.FOneItem.FGroupid
ogroup.GetOneGroupInfo


dim ooffontract
set ooffontract = new COffContractInfo
ooffontract.FRectDesignerID = designer
ooffontract.GetPartnerOffContractInfo


dim returnsongjangStr

returnsongjangStr = returnsongjangStr + "10x10" & chr(9)
returnsongjangStr = returnsongjangStr + "(��)�ٹ�����" & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.FCompany_name  & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Fdeliver_phone  & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Fdeliver_hp  & chr(9)
returnsongjangStr = returnsongjangStr + replace(ogroup.FOneItem.Freturn_zipcode,"-","") & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Freturn_address  & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Freturn_address2  & chr(9)
returnsongjangStr = returnsongjangStr + "10x10 ��ǰ" & chr(9)
returnsongjangStr = returnsongjangStr + "��ǰ��ǰ" & chr(9)
returnsongjangStr = returnsongjangStr + opartner.FOneItem.FID
%>

<!-- returnsongjangStr = FormatDate(now(),"0000.00.00 00:00:00")
returnsongjangStr = Replace(returnsongjangStr,".","")
returnsongjangStr = Replace(returnsongjangStr,":","")
returnsongjangStr = Replace(returnsongjangStr," ","")
returnsongjangStr = returnsongjangStr & chr(9)
-->

<script language='javascript'>
function copyComp(comp) {
	comp.focus()
	comp.select()
	therange=comp.createTextRange()
	therange.execCommand("Copy")
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

function SaveBrandInfo(frm){
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
		frm.maeipdiv.focus();
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

	var ret = confirm('�귣�� ������ ���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}

function SaveUpcheInfo(frm){
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
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		//frm.submit();
	}

}

function PopUpcheList(frmname){
	var popwin = window.open("/admin/lib/popupchelist.asp?frmname=" + frmname,"popupchelist","width=700 height=540 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>
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

<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	�귣�� ID : <input type="text" name="designer" value="<%= designer %>" Maxlength="32" size="16">
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>


<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#000000>
<form name="frmupche" method="post" action="doupcheedit.asp">
<input type="hidden" name="mode" value="groupedit">
<input type="hidden" name="uid" value="<%= designer %>">
	<tr bgcolor="#DDDDFF">
		<td colspan=4><b>1.��ü��������</b></td>
	</tr>
	<tr >
		<td width="100" bgcolor="#DDDDFF">��ü�ڵ�</td>
		<td bgcolor="#FFFFFF" width="200">
		<input type="text" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" size="7" maxlength="5" style="background-color:#EEEEEE;" readonly>
		<% if (C_ADMIN_AUTH=true) then %>
		<input type="button" value="��ü����" onClick="PopUpcheList('frmupche');">
		<% end if %>
		</td>
		<td width="100" bgcolor="#DDDDFF">��ü��</td>
		<td bgcolor="#FFFFFF" width="200">
		<%= ogroup.FOneItem.FCompany_name %>
		</td>
	</tr>
	<tr >
		<td width="100" bgcolor="#DDDDFF">�����귣��ID</td>
		<td colspan="3" bgcolor="#FFFFFF"><%= ogroup.FOneItem.getBrandList %></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**����ڵ������**</td>
	</tr>

	<tr>
		<td width="100" bgcolor="#DDDDFF">ȸ���(��ȣ)</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH=true) then %>
			<input type="text" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="28" maxlength="32">
			<% else %>
			<input type="text" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="28" maxlength="32" style="background-color:#EEEEEE;" readonly>
			<% end if %>
		</td>
		<td width="100" bgcolor="#DDDDFF">��ǥ��</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH=true) then %>
			<input type="text" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="16" maxlength="16">
			<% else %>
			<input type="text" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="16" maxlength="16" style="background-color:#EEEEEE;" readonly>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td width="100" bgcolor="#DDDDFF">����ڹ�ȣ</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH=true) then %>
			<input type="text" name="company_no" value="<%= ogroup.FOneItem.Fcompany_no %>" size="16" maxlength="20">
			<% else %>
			<input type="text" name="company_no" value="<%= ogroup.FOneItem.Fcompany_no %>" size="16" maxlength="20" style="background-color:#EEEEEE;" readonly>
			<% end if %>
		</td>
		<td width="100" bgcolor="#DDDDFF">��������</td>
		<td bgcolor="#FFFFFF">
			<select name="jungsan_gubun">
			<option value="�Ϲݰ���" <% if ogroup.FOneItem.Fjungsan_gubun="�Ϲݰ���" then response.write "selected" %> >�Ϲݰ���</option>
			<option value="���̰���" <% if ogroup.FOneItem.Fjungsan_gubun="���̰���" then response.write "selected" %> >���̰���</option>
			<option value="��õ¡��" <% if ogroup.FOneItem.Fjungsan_gubun="��õ¡��" then response.write "selected" %> >��õ¡��</option>
			<option value="�鼼" <% if ogroup.FOneItem.Fjungsan_gubun="�鼼" then response.write "selected" %> >�鼼</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">����������</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" name="company_zipcode" value="<%= ogroup.FOneItem.Fcompany_zipcode %>" size="7" maxlength="7"><a href="javascript:popZip('s');"><img src="http://www.10x10.co.kr/images/zip_search.gif" border=0 align="absmiddle"></a><br>
			<input type="text" name="company_address" value="<%= ogroup.FOneItem.Fcompany_address %>" size="16" maxlength="64">&nbsp;
			<input type="text" name="company_address2" value="<%= ogroup.FOneItem.Fcompany_address2 %>" size="42" maxlength="64">
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
		<input type="text" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>" size="7" maxlength="7"><a href="javascript:popZip('m');"><img src="http://www.10x10.co.kr/images/zip_search.gif" border=0 align="absmiddle"></a>
		<input type=checkbox name=samezip onclick="SameReturnAddr(this.checked)">��
		<br>
			<input type="text" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>" size="16" maxlength="64">&nbsp;
			<input type="text" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>" size="42" maxlength="64">
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

<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#000000>
<form name="frmbrand" method="post" action="doupcheedit.asp">
<input type="hidden" name="uid" value="<%= opartner.FOneItem.FID %>">
<input type="hidden" name="mode" value="brandedit">
<tr>
	<td bgcolor="#FFDDDD" colspan=6><b>2.�귣���������</b></td>
</tr>
<tr>
	<td width="100" bgcolor="#FFDDDD">ȸ���</td>
	<td bgcolor="#FFFFFF" colspan=2><%= opartner.FOneItem.FCompany_name %></td>
	<td width="100" bgcolor="#FFDDDD" >����ȣ</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="prtidx" value="<%= opartner.FOneItem.getRackCode %>" size="4" maxlength="4">
	(�⺻�� : 9999)</td>
	</td>
</tr>
<tr>
	<td bgcolor="#FFDDDD">�귣��ID</td>
	<td bgcolor="#FFFFFF" colspan=2><%= opartner.FOneItem.FID %></td>
	<td bgcolor="#FFDDDD">�н�����</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="password" value="<%= opartner.FOneItem.Fppass %>">
	</td>
</tr>
<tr>
	<td bgcolor="#FFDDDD">��Ʈ��Ʈ��<br>(�ѱ�)</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="socname_kor" value="<%= opartner.FOneItem.Fsocname_kor %>">
	</td>
	<td bgcolor="#FFDDDD">��Ʈ��Ʈ��<br>(����)</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="socname" value="<%= opartner.FOneItem.Fsocname %>">
	</td>
</tr>
<tr>
	<td rowspan="3" bgcolor="#FFDDDD">�귣��<br>��뿩��<br>(ī�װ�����)</td>
	<td bgcolor="#FFFFFF">�ٹ�����</td>
	<td bgcolor="#FFFFFF"><input type=radio name="isusing" value="Y" <% if opartner.FOneItem.Fisusing="Y" then response.write "checked" %> >��� <input type=radio name="isusing" value="N" <% if opartner.FOneItem.Fisusing="N" then response.write "checked" %> >������</td>
	<td rowspan="3" bgcolor="#FFDDDD">��Ʈ��Ʈ<br>ǥ�ÿ���<br>(�귣������)</td>
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
	<td bgcolor="#FFFFFF" colspan=2><% DrawBrandGubunCombo "userdiv", opartner.FOneItem.Fuserdiv %></td>
	<td bgcolor="#FFDDDD">�����</td>
	<td bgcolor="#FFFFFF" colspan=2><%= opartner.FOneItem.Fregdate %></td>
</tr>
<tr >
	<td bgcolor="#FFDDDD">ī�װ�</td>
	<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "catecode", opartner.FOneItem.Fcatecode %></td>
	<td bgcolor="#FFDDDD" >���MD</td>
	<td bgcolor="#FFFFFF" colspan=2><% drawSelectBoxCoWorker "mduserid", opartner.FOneItem.Fmduserid %></td>
</tr>
<tr>
	<td bgcolor="#FFDDDD" >��ǰ����</td>
	<td bgcolor="#FFFFFF" colspan=5>

	<input type=text name=brandsongjang value="<%= returnsongjangStr %>" size=50 > <a href="javascript:copyComp(frmbrand.brandsongjang);">����</a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=6>**�����û���**</td>
</td>
<tr >
	<td bgcolor="#FFDDDD" >�⺻����</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<% DrawBrandMWUCombo "maeipdiv",opartner.FOneItem.Fmaeipdiv %>
	<input type="text" name="defaultmargine" value="<%= opartner.FOneItem.Fdefaultmargine %>" size="4" style="text-align:right"> %
	</td>
	<td bgcolor="#FFDDDD" >������ :</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<% DrawJungsanDateCombo "jungsan_date", opartner.FOneItem.Fjungsan_date %>
	</td>
</tr>
<tr >
	<td bgcolor="#FFDDDD" >��������(������)</td>
	<td bgcolor="#FFFFFF" colspan=2>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=100%>
		<tr>
			<td width="100"><b>��������ǥ</b></td>
			<td><%= ooffontract.GetSpecialChargeDivName("streetshop000") %></td>
			<td width="40"><%= ooffontract.GetSpecialDefaultMargin("streetshop000") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv<>"3") and (ooffontract.FItemList(i).Fshopid<>"streetshop000") then %>
		<tr>
			<td><%= ooffontract.FItemList(i).Fshopname %></td>
			<td><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
	<td bgcolor="#FFDDDD" >������ </td>
	<td bgcolor="#FFFFFF" colspan=2>
	<% DrawJungsanDateCombo "jungsan_date_off", opartner.FOneItem.Fjungsan_date_off %>
	</td>
</tr>
<tr >
	<td bgcolor="#FFDDDD" >��������(������)</td>
	<td bgcolor="#FFFFFF" colspan=2>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=100%>
		<tr>
			<td width="100"><b>����������ǥ</b></td>
			<td><%= ooffontract.GetSpecialChargeDivName("streetshop800") %></td>
			<td width="40"><%= ooffontract.GetSpecialDefaultMargin("streetshop800") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="3") and (ooffontract.FItemList(i).Fshopid<>"streetshop800") then %>
		<tr>
			<td ><%= ooffontract.FItemList(i).Fshopname %></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
	<td bgcolor="#FFDDDD" >������ </td>
	<td bgcolor="#FFFFFF" colspan=2>
	<% DrawJungsanDateCombo "jungsan_date_frn", opartner.FOneItem.Fjungsan_date_frn %>
	</td>
</tr>


<!--
<tr>
	<td bgcolor="#FFDDDD" >���� ���¿���</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<% if opartner.FOneItem.Fpartnerusing="Y" then %>
		<input type="radio" name="partnerusing" value="Y" checked >�����
		<input type="radio" name="partnerusing" value="N" >������
		<% else %>
		<input type="radio" name="partnerusing" value="Y"  >�����
		<input type="radio" name="partnerusing" value="N" checked ><font color=red>������</font>
		<% end if %>
	</td>
</tr>
-->

<% if ogroup.FOneItem.FGroupId<>"" then %>
<tr>
	<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" value="�귣������ ����" onclick="SaveBrandInfo(frmbrand);"></td>
</tr>
<% else %>
<tr>
	<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" value="�귣������ ����" onclick="alert('��ü������ ���� ���� �Ͻ��� �귣�������� ���� �� �� �ֽ��ϴ�.');"></td>
</tr>
<% end if %>
</form>
</table>

<br>

<table  width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#000000>
<form name="frmetc" method=post action="http://partner.10x10.co.kr/linkweb/doprofileimageadmin.asp" enctype="multipart/form-data">
<input type=hidden name=designerid value="<%= opartner.FOneItem.FID %>">
	<tr>
		<td bgcolor="#DDDDFF" colspan=4><b>3.�귣�� ��Ÿ����</b></td>
	</tr>
	<tr>
		<td width="110" bgcolor="#DDDDFF" >�ΰ�</td>
		<td bgcolor="#FFFFFF">
		<img name=logoimg src="<%= opartner.FOneItem.getSocLogoUrl %>" width=150 height=100><br>
		(�귣�� �ΰ�� 150x100 �ȼ��� ������ �ֽ��ϴ�.)<br>
		<input type=file name=file1 size=40 onchange="ChangeLogo(this,frmetc.logoimg);">
		</td>
	</tr>
	<tr>
		<td width="110" bgcolor="#DDDDFF" >Ÿ��Ʋ</td>
		<td bgcolor="#FFFFFF">
		<img name=titleimg src="<%= opartner.FOneItem.getTitleImgUrl %>" width=300 height=75><br>
		(Ÿ��Ʋ�̹����� 600x150 �ȼ��� ������ �ֽ��ϴ�.)(600x150)<br>
		<input type=file name=file2 size=40 onchange="ChangeTitle(this,frmetc.titleimg);">
		</td>
	</tr>
	<tr>
		<td width="110" bgcolor="#DDDDFF" >�����̳�<br>�ڸ�Ʈ</td>
		<td bgcolor="#FFFFFF">
		<textarea name="dgncomment" cols=64 rows=6><%= opartner.FOneItem.Fdgncomment %></textarea>
		</td>
	</tr>
	<tr>
		<td colspan="2" align=center bgcolor="#FFFFFF"><input type="button" value="�귣�� ��Ÿ���� ����" onclick="SaveBrandEtcInfo(frmetc);"></td>
	</tr>
</form>
</table>
<%
set opartner = Nothing
set ogroup = Nothing
set ooffontract = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->