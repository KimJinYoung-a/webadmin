<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%

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


	var ret = confirm('�귣�� ������ ���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}

function SaveUpcheInfo(frm){
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
		<input type="button" value="��ü����" onClick="PopUpcheSelect('frmupche');">
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
		<td height="30" colspan="4" bgcolor="#FFFFFF" align="center"><input class="button" type="button" value="��ü���� ����" onclick="PopUpcheInfoEdit('<%= ogroup.FOneItem.FGroupId %>');"></td>
	</tr>
</form>
</table>

<br>
<br>
<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#000000>
<form name="frmbrand" method="post" action="doupcheedit.asp">
<input type="hidden" name="uid" value="<%= opartner.FOneItem.FID %>">
<input type="hidden" name="mode" value="brandedit">
<tr>
	<td bgcolor="#FFDDDD" colspan=4><b>2.�귣���������</b></td>
</tr>
<tr>
	<td width="100" bgcolor="#FFDDDD">ȸ���</td>
	<td bgcolor="#FFFFFF"><%= opartner.FOneItem.FCompany_name %></td>
	<td width="100" bgcolor="#FFDDDD" >�귣��ID</td>
	<td bgcolor="#FFFFFF"><%= opartner.FOneItem.FID %></td>
</tr>

<tr>
	<td bgcolor="#FFDDDD">��Ʈ��Ʈ��<br>(�ѱ�)</td>
	<td bgcolor="#FFFFFF">
		<%= opartner.FOneItem.Fsocname_kor %>
	</td>
	<td bgcolor="#FFDDDD">��Ʈ��Ʈ��<br>(����)</td>
	<td bgcolor="#FFFFFF">
		<%= opartner.FOneItem.Fsocname %>
	</td>
</tr>

<tr>
	<td bgcolor="#FFDDDD" >���� ���¿���</td>
	<td bgcolor="#FFFFFF" colspan=2>
		<% if opartner.FOneItem.Fpartnerusing="Y" then %>
		�����
		<% else %>
		<font color=red>������</font>
		<% end if %>
	</td>
	<td bgcolor="#FFFFFF" align="right"><a href="javascript:PopBrandAdminUsingChange('<%= opartner.FOneItem.FID %>');"><img src="/images/icon_modify.gif" border="0"></td>
</tr>
<tr>
	<td height="30" colspan="4" bgcolor="#FFFFFF" align="center"><input class="button" type="button" value="�귣������ ����" onclick="PopBrandInfoEdit('<%= opartner.FOneItem.FID %>');"></td>
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

<br>
<br>

<%
set opartner = Nothing
set ogroup = Nothing
set ooffontract = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->