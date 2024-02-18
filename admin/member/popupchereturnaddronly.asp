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
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->

<%
dim ogroup,i, groupid
	groupid = request("groupid")

'groupid = "G00240"

set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = groupid
	ogroup.GetOneGroupInfo

dim OReturnAddr
set OReturnAddr = new CCSReturnAddress
	OReturnAddr.FRectGroupCode = groupid
	OReturnAddr.FPageSize = 200
	OReturnAddr.GetReturnAddressList

%>

<script type="text/javascript">

function CopyZip(flag,post1,post2,add,dong){
	var frm = eval(flag);

	frm.return_zipcode.value= post1 + "-" + post2;
	frm.return_address.value= add;
	frm.return_address2.value= dong;
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function SameReturnAddr(frm, bool){
	if (bool){
		frm.return_zipcode.value = document.frmgroup.return_zipcode.value;
		frm.return_address.value = document.frmgroup.return_address.value;
		frm.return_address2.value = document.frmgroup.return_address2.value;
	}else{
		frm.return_zipcode.value = "";
		frm.return_address.value = "";
		frm.return_address2.value = "";
	}
}

function SameReturnName(frm, bool){
	if (bool){
		frm.deliver_name.value = document.frmgroup.deliver_name.value;
		frm.deliver_phone.value = document.frmgroup.deliver_phone.value;
		frm.deliver_hp.value = document.frmgroup.deliver_hp.value;
	}else{
		frm.deliver_name.value = "";
		frm.deliver_phone.value = "";
		frm.deliver_hp.value = "";
	}
}

function ModifyReturnAddress(frm){
	if (frm.return_zipcode.value.length < 1){
		alert('�����ȣ�� �����ϼ���.');
		frm.return_zipcode.focus();
		return;
	}

	if (frm.return_address.value.length < 1){
		alert('�ּҸ� ��Ȯ�� �Է��ϼ���.');
		frm.return_address.focus();
		return;
	}

	if (frm.return_address2.value.length < 1){
		alert('�ּҸ� ��Ȯ�� �Է��ϼ���.');
		frm.return_address2.focus();
		return;
	}

	if (frm.deliver_name.value.length < 1){
		alert('��۴���� �̸��� �Է��ϼ���.');
		frm.deliver_name.focus();
		return;
	}

	if (frm.deliver_phone.value.length < 1){
		alert('��۴���� ��ȭ��ȣ�� �Է��ϼ���.');
		frm.deliver_phone.focus();
		return;
	}

	if (frm.deliver_hp.value.length < 1){
		alert('��۴���� �ڵ�����ȣ�� �Է��ϼ���.');
		frm.deliver_hp.focus();
		return;
	}

	var ret = confirm('�귣�� ��ǰ������ ���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}


function PopUpcheList(frmname){
	var popwin = window.open("/admin/lib/popupchelist.asp?frmname=" + frmname,"popupchelist","width=700 height=540 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=4><b>* �귣�庰 ��ǰ���� �� �ù�� ����</b></td>
</tr>

<!--
<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**�繫�� �ּ�**</td>
</tr>
-->
<tr>
	<td height="25" width="150" bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
	<td width="250" bgcolor="#FFFFFF"><b><%= ogroup.FOneItem.FCompany_name %></b></td>
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">�׷��ڵ�</td>
	<td bgcolor="#FFFFFF"><b><%= ogroup.FOneItem.FGroupId %></b></td>
</tr>
<!--
<tr>
	<td height="25" bgcolor="<%= adminColor("tabletop") %>">��۴���ڸ�</td>
	<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_name %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">��۴���� ��ȭ��ȣ</td>
	<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_phone %></td>
</tr>
<tr>
	<td height="25" bgcolor="<%= adminColor("tabletop") %>">��۴���� �̸���</td>
	<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_email %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">��۴���� �ڵ�����ȣ</td>
	<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_hp %></td>
</tr>
<tr>
	<form name=frmgroup>
	<input type=hidden name=return_zipcode value="<%= ogroup.FOneItem.Freturn_zipcode %>">
	<input type=hidden name=return_address value="<%= ogroup.FOneItem.Freturn_address %>">
	<input type=hidden name=return_address2 value="<%= ogroup.FOneItem.Freturn_address2 %>">
	<input type=hidden name=deliver_name value="<%= ogroup.FOneItem.Fdeliver_name %>">
	<input type=hidden name=deliver_phone value="<%= ogroup.FOneItem.Fdeliver_phone %>">
	<input type=hidden name=deliver_hp value="<%= ogroup.FOneItem.Fdeliver_hp %>">
	</form>
	<td height="25" bgcolor="<%= adminColor("tabletop") %>">�ּ�</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		[<%= ogroup.FOneItem.Freturn_zipcode %>] <%= ogroup.FOneItem.Freturn_address %> <%= ogroup.FOneItem.Freturn_address2 %>
	</td>
</tr>
-->

<!--


<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">* [ ��ü �繫������ ]</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="<%= ogroup.FOneItem.Fdeliver_name %>" size="16" maxlength="16"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�̸���</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="<%= ogroup.FOneItem.Fdeliver_email %>" size="16" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">��ȭ��ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="<%= ogroup.FOneItem.Fdeliver_phone %>" size="16" maxlength="16"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="<%= ogroup.FOneItem.Fdeliver_hp %>" size="16" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�ּ�</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<input type="text" class="text" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>" size="7" maxlength="7"><a href="javascript:popZip('m');"><img src="http://fiximage.10x10.co.kr/images/zip_search.gif" border=0 align="absmiddle"></a>
	<input type=checkbox name=samezip onclick="SameReturnAddr(this.checked)">��
	<br>
		<input type="text" class="text" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>" size="30" maxlength="64">&nbsp;
		<input type="text" class="text" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>" size="46" maxlength="64">
	</td>
</tr>

-->
</table>

<br>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		**�귣�庰 ��ǰ���� ����**
	</td>
</tr>
<% for i=0 to OReturnAddr.FResultCount - 1%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">�귣��ID</td>
	<td>�귣���</td>
	<td width="200">��ǰ�ּ�</td>
	<td width="130">��۴����</td>

	<td width="130">����ù��</td>
	<td width="50">����</td>
</tr>
<tr bgcolor="#FFFFFF">
<form name="frm<%= i %>" method="post" action="/admin/member/doupcheedit.asp" onsubmit="return false;">
<input type=hidden name=uid value="<%= OReturnAddr.FItemList(i).Fbrandid %>">
<input type=hidden name=mode value="modifyreturnaddress">
	<td height="25" align="center"><%= OReturnAddr.FItemList(i).Fbrandid %></td>
	<td align="center"><%= OReturnAddr.FItemList(i).Fstreetname_kor %><br><%= OReturnAddr.FItemList(i).Fstreetname_eng %></td>
	<td>
		<input type="text" class="text" name="return_zipcode" value="<%= OReturnAddr.FItemList(i).FreturnZipcode %>" size="7" maxlength="7">
	    <input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frm<%= i %>','D')">
		<input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('frm<%= i %>','D')">
	    <!--<input type="button" class="button" value="�˻�(��)" onClick="popZip('frm<%'= i %>');">-->
		<!-- <input type=checkbox name=samezip onclick="SameReturnAddr(frm<%= i %>, this.checked)">�� -->
		<br>
		<input type="text" class="text" name="return_address" value="<%= OReturnAddr.FItemList(i).FreturnZipaddr %>" size="30" maxlength="64">
		<br>
		<input type="text" class="text" name="return_address2" value="<%= OReturnAddr.FItemList(i).FreturnEtcaddr %>" size="30" maxlength="64">
	</td>
	<td align="left">
		����� <input type="text" class="text" name="deliver_name" value="<%= OReturnAddr.FItemList(i).FreturnName %>" size="8" maxlength="32">
		<!-- <input type=checkbox name=samename onclick="SameReturnName(frm<%= i %>, this.checked)">�� -->
		<br>
		��ȭ <input type="text" class="text" name="deliver_phone" value="<%= OReturnAddr.FItemList(i).FreturnPhone %>" size="16" maxlength="16"><br>
		HP <input type="text" class="text" name="deliver_hp" value="<%= OReturnAddr.FItemList(i).Freturnhp %>" size="16" maxlength="16">
		Email <input type="text" class="text" name="deliver_email" value="<%= OReturnAddr.FItemList(i).FreturnEmail %>" size="30" maxlength="100">
	</td>
	<td align="center"><% drawSelectBoxDeliverCompany "defaultsongjangdiv",OReturnAddr.FItemList(i).Fsongjangdiv %></td>
	<td align="center">
		<input type="button" class="button" value="����" onclick="ModifyReturnAddress(frm<%= i %>)">
	</td>
</form>
</tr>
<% next %>
</table>

<%
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
