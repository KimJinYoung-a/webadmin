<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%


dim makerid
makerid = requestCheckVar(request("makerid"),32)


dim opartner
set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid

if makerid<>"" then
	opartner.GetOnePartnerNUser
end if

dim ooffontract
set ooffontract = new COffContractInfo
ooffontract.FRectDesignerID = makerid

if makerid<>"" then
	ooffontract.GetPartnerOffContractInfo
end if

dim i

''DefaultCenterMwdiv  
dim DefaultCenterMwdiv
DefaultCenterMwdiv = GetDefaultItemMwdivByBrand(makerid)
%>
<script language='javascript'>
function AttachImage(comp,filecomp){
	comp.src=filecomp.value;
}

function ChangeBrand(comp){
	location.href="?makerid=" + comp.value;
}

function CheckAddItem(frm){
/*
	if ((frm.itemgubun[0].checked==false)&&(frm.itemgubun[1].checked==false)){
		alert('��ǰ������ �����ϼ���.');
		return;
	}
*/

	if (frm.makerid.value.length<1){
		alert('�귣�带 �����ϼ���.');
		return;
	}
/*
	if (frm.cd1.value.length<1){
		alert('ī�װ��� �����ϼ���.');
		return;
	}
*/
	if (frm.shopitemname.value.length<1){
		alert('��ǰ���� �Է��ϼ���.');
		frm.shopitemname.focus();
		return;
	}

	if ((frm.extbarcode.value.length>0) && (frm.extbarcode.value.length<10)){
		alert('���ڵ� ���̰� �ʹ� ª���ϴ�. ���� ���ڵ尡 �ִ°�츸 �Է��� �ּ���' );
		frm.extbarcode.focus();
		return;
	}

	if (!IsDigit(frm.shopitemprice.value)){
		alert('�ǸŰ��� ���ڸ� �����մϴ�.');
		frm.shopitemprice.focus();
		return;
	}


//	if (!IsDigit(frm.discountsellprice.value)){
//		alert('���� �ǸŰ��� ���ڸ� �����մϴ�.');
//		frm.discountsellprice.focus();
//		return;
//	}


	if (!IsDigit(frm.shopsuplycash.value)){
		alert('��ü ���԰��� ���ڸ� �����մϴ�.');
		frm.shopsuplycash.focus();
		return;
	}

	if (!IsDigit(frm.shopbuyprice.value)){
		alert('�� ���ް��� ���ڸ� �����մϴ�.');
		frm.shopbuyprice.focus();
		return;
	}

	if (((frm.shopsuplycash.value!=0)||(frm.shopbuyprice.value!=0))){
		if (!confirm('!! �⺻ ��� ������ �ٸ� ��쿡�� ���԰� ���ް��� �Է� �ϼž� �մϴ�. \n\n��� �Ͻðڽ��ϱ�?')){
			return;
		}
	}
/*
	if (frm.file1.value.length<1){
		alert('�̹����� �Է��� �ּ��� - �ʼ� �����Դϴ�.');
		frm.file1.focus();
		return;
	}
*/

	var ret = confirm('�߰��Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}

// ============================================================================
// ī�װ����
function editCategory(cdl,cdm,cdn){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cdn=" + cdn ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.frmedit;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;
}
</script>
<table border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#FFFFFF>
<tr>
	<td>&gt;&gt;�������� ��ǰ ���</td>
</tr>
</table>

<table border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#3d3d3d>
<form name="frmedit" method=post action="shopitem_process.asp" >
<input type=hidden name=mode value="addetcoffitem">
<input type=hidden name=makerid value="<%= makerid %>">
<tr bgcolor="#FFDDDD">
	<td width=100>�귣�� ID</td>
	<td bgcolor="#FFFFFF" colspan=5><%= makerid %>
	</td>
</tr>
<% if makerid<>"" then %>

<tr bgcolor="#DDDDFF">
	<td width=100>��ǰ����</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type="radio" name="itemgubun" value="00" checked >�޴���ǰ(00) &nbsp;
	
	</td>
</tr>
<!--
<tr bgcolor="#DDDDFF" >
	<td width=100 >ī�װ�</td>
	<td bgcolor="#FFFFFF" colspan=5>
	  <input type="hidden" name="cd1" value="">
	  <input type="hidden" name="cd2" value="">
	  <input type="hidden" name="cd3" value="">

      <input type="text" name="cd1_name" value="" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" value="" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" value="" size="12" readonly style="background-color:#E6E6E6">

      <input type="button" value="����" onclick="editCategory(frmedit.cd1.value,frmedit.cd2.value,frmedit.cd3.value);">
	</td>
</tr>
-->
<tr bgcolor="#DDDDFF">
	<td width=100>��ǰ��</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=text name="shopitemname" value="" size=40 maxlength=40 class="input_01" >
	</td>
</tr>
<!--
<tr bgcolor="#DDDDFF">
	<td width=100>�ɼǸ�</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=hidden name="shopitemoptionname" value="">
	</td>
</tr>
-->
<tr bgcolor="#DDDDFF">
	<td width=100>������ڵ�</td>
	<td bgcolor="#FFFFFF" colspan=5><input type=text name="extbarcode" value="" size=20 maxlength=20 class="input_01" >(�ִ� ��츸 ���)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>�������</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=radio name=isusing value="Y" checked >�����
	<input type=radio name=isusing value="N">������
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>���͸��Ա���</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=radio name=centermwdiv value="W" disabled >Ư��
	<input type=radio name=centermwdiv value="M" checked >����
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td width=100 >��������</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=radio name=vatinclude value="Y" checked >����
	<input type=radio name=vatinclude value="N">�鼼
	</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td width=100 align="left" rowspan="3">���ݼ���</td>
	<td bgcolor="#FFFFFF" >�ǸŰ�</td>
	<td bgcolor="#FFFFFF" >���԰�</td>
	<td bgcolor="#FFFFFF" >���ް�</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td bgcolor="#FFFFFF"><input type=text name="shopitemprice" value="" size=8 maxlength=9 class="input_right" ></td>
	<td bgcolor="#FFFFFF"><input type=text name="shopsuplycash" value="0" size=8 maxlength=9 class="input_right" ></td>
	<td bgcolor="#FFFFFF" ><input type=text name="shopbuyprice" value="0" size=8 maxlength=9 class="input_right" ></td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td bgcolor="#FFFFFF" ></td>
	<td bgcolor="#FFFFFF" colspan="2">(0�ΰ�� �⺻���� ���� ������)</td>
</tr>

</tr>

</form>
<tr bgcolor="#FFFFFF">
	<td colspan=6 align=center><input type=button value=" ��  �� " onclick="CheckAddItem(frmedit)" class="input_01"></td>
</tr>
<% end if %>
</table>

<%
set opartner = Nothing
set ooffontract = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->