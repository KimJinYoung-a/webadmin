<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim makerid, shopid, gubuncd, masteridx
makerid     = request("makerid")
shopid      = request("shopid")
gubuncd     = request("gubuncd")
masteridx   = request("masteridx")
%>
<script language='javascript'>
function AddThis(frm){
	if (frm.sellprice.value.length<1){
		alert('�ǸŰ��� �Է��ϼ���.');
		frm.sellprice.focus();
		return;
	}

	if (frm.suplyprice.value.length<1){
		alert('���԰��� �Է��ϼ���.');
		frm.suplyprice.focus();
		return;
	}

	if (frm.commission.value.length<1){
		alert('�����Ḧ �Է��ϼ���.');
		frm.commission.focus();
		return;
	}


	if (frm.itemno.value.length<1){
		alert('������ �Է��ϼ���.');
		frm.itemno.focus();
		return;
	}

	if (confirm('��Ÿ ������ �߰��Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function calCommission() {
	var frm = document.frm;
	var sellprc = frm.sellprice.value;
	var suplyprc = frm.suplyprice.value;
	if(!sellprc){sellprc=0;}
	if(!suplyprc){suplyprc=0;}
	frm.commission.value = parseInt(sellprc)-parseInt(suplyprc);
	frm.sellprice.value = parseInt(sellprc);
	frm.suplyprice.value = parseInt(suplyprc);
}
</script>
<table border=0 cellspacing="1" class="a"  width=500 bgcolor=#3d3d3d>
<form name=frm method=post action="off_jungsan_process.asp">
<input type=hidden name=mode value="addetcdetail">
<input type=hidden name=gubuncd value="B999">
<input type=hidden name=shopid value="<%= shopid %>">
<input type=hidden name=masteridx value="<%= masteridx %>">
<tr>
	<td width=120 bgcolor="#DDDDFF">��ǰ�ڵ�</td>
	<td bgcolor="#FFFFFF">
	<input type=text name="itemgubun" value="00" size=2 maxlength=2 >
	<input type=text name="itemid" value="000000" size=9 maxlength=9 >
	<input type=text name="itemoption" value="0000" size=4 maxlength=4 >
	</td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">��ǰ��</td>
	<td bgcolor="#FFFFFF"><input type=text name="itemname" value="" size=26 maxlength=40></td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">�ɼǸ�</td>
	<td bgcolor="#FFFFFF"><input type=text name="itemoptionname" value="" size=26 maxlength=40></td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">������̵�</td>
	<td bgcolor="#FFFFFF"><input type=text name="makerid" value="<%= makerid %>" size=26 maxlength=32></td>
</tr>

<tr>
	<td width=120 bgcolor="#DDDDFF">�ǸŰ�</td>
	<td bgcolor="#FFFFFF">
	<input type=text name="sellprice" value="" size=9 maxlength=9 style="text-align:right" onkeyup="calCommission();" />(�ǸŻ�ǰ�� �ƴѰ�� 0��)
	</td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">������</td>
	<td bgcolor="#FFFFFF">
	<input type=text name="commission" value="0" size=9 maxlength=9 style="text-align:right" readOnly >
	</td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">���԰�(�����)</td>
	<td bgcolor="#FFFFFF"><input type=text name="suplyprice" value="" size=9 maxlength=9 style="text-align:right" onkeyup="calCommission();" /></td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><input type=text name="itemno" value="" size=3 maxlength=5 ></td>
</tr>
<tr>
	<td colspan=2 align=center bgcolor="#FFFFFF"><input type=button value=" �� �� " onclick="AddThis(frm)"></td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->