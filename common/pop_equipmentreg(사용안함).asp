<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	2009�� 01�� 19�� �ѿ�� ����
'#######################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/bscclass/equipmentcls.asp"-->

<%
dim idx
	idx = request("idx")

if idx="" then idx=0
dim oequip
set oequip = new CEquipment
	oequip.FRectIdx = idx
	oequip.getOneEquipment
%>

<script language="javascript">

//��밡����ip���ý���
function checkip(frm)
{
	if (document.frmreg.checkipform.value!="")
	{
		document.frmreg.detail_ip.value = ""
		document.frmreg.detail_ip.value = document.frmreg.checkipform.value;
	}	
}

//����
function regEquip(frm){
	//�ʼ��Է�üũ
	if (frm.equip_gubun.value.length<1){
		alert('��񱸺��� �����ϼ���.');
		frm.equip_gubun.focus();
		return;
	}

	if (frm.part_code.value.length<1){
		alert('��뱸���� �����ϼ���.');
		frm.part_code.focus();
		return;
	}


	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

//�ɼ� �� ���úκн���
function selectChange(comp){
	if (comp.name=="equip_gubun"){
		//�ɼ�1
		if ((comp.value=="PC")||(comp.value=="NB")||(comp.value=="MO")||(comp.value=="SV")||(comp.value=="FS")){
			div_detail_quality1.style.display="inline";
		}else{
			div_detail_quality1.style.display="none";
		}

		if ((comp.name=="equip_gubun")&&(comp.value=="MO")){
			detail_quality1_name.innerText = "����ͻ�� :";
			detail_quality1_etc.innerText = "(LCD 17, CRT 19)";
		}else{
			detail_quality1_name.innerText = "CPU :";
			detail_quality1_etc.innerText = "(P2.8, C2.4, AMD 1800 ..)";
		}

		//�ɼ�2
		if ((comp.value=="PC")||(comp.value=="NB")||(comp.value=="SV")||(comp.value=="FS")){
			div_detail_quality2.style.display="inline";
		}else{
			div_detail_quality2.style.display="none";
		}
		if ((comp.value=="SC")||(comp.value=="PR")||(comp.value=="CX")){
			div_detail_quality3.style.display="inline";
		}else{
			div_detail_quality3.style.display="none";
		}
		if (comp.value=="NE"){
			div_detail_quality4.style.display="inline";
		}else{
			div_detail_quality4.style.display="none";
		}
		if (comp.value=="UP"){
			div_detail_quality5.style.display="inline";
		}else{
			div_detail_quality5.style.display="none";
		}
		
	<!--}else if (comp.name=="part_code"){
		if ((comp.value=="10")&&(frmreg.usinguserid.value.length<1)){
			frmreg.usinguserid.value = frmreg.curruserid.value;
		}else{
			//frmreg.usinguserid.value = "";
		}-->
	}
}	

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		�� ��� �ڻ� ����Ʈ �߰� </strong> / �ǵ��� �ڼ��� �Է��� �ּ���. 	
	</td>
	<td align="right">		
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<!--�ϴ����̺����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
<form name="frmreg" method="post" action="do_equipment.asp">
<input type="hidden" name="idx" value="<%= oequip.FOneItem.Fidx %>">
<input type="hidden" name="curruserid" value="<%= session("ssBctId") %>">
<input type="hidden" name="currusername" value="<%= session("ssBctCname") %>">

<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">����ڵ�</td>
	<td colspan="2"><%= oequip.FOneItem.getEquipCode %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">��񱸺�</td>
	<td >
		<% DrawEquipMentGubun "10","equip_gubun",oequip.FOneItem.Fequip_gubun ," onchange='selectChange(frmreg.equip_gubun)'" %>
	</td>
	
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">��뱸��</td>
	<td >
		<% DrawEquipMentGubun "20","part_code",oequip.FOneItem.Fpart_code ,"" %>
	</td>
	
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">��ǰ��</td>
	<td colspan="2">
		<input type="text" name="equip_name" value="<%= oequip.FOneItem.Fequip_name %>" size="60" maxlength="60">
		(ex : �ﺸ �帲�ý� 74SC)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">�ø����ȣ</td>
	<td colspan="2">
		<input type="text" name="model_name" value="<%= oequip.FOneItem.Fmodel_name %>" size="60" maxlength="60">
		(ex : PN17AS)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">������</td>
	<td colspan="2">
		<input type="text" name="manufacture_company" value="<%= oequip.FOneItem.Fmanufacture_company %>" size="60" maxlength="60">
		(ex : �Ｚ����, LG����)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">����ó</td>
	<td colspan="2">
		<input type="text" name="buy_company_name" value="<%= oequip.FOneItem.Fbuy_company_name %>" size="60" maxlength="60">
		(ex : �Ｚ��, ������ũ, DELL�ڸ���)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">������</td>
	<td colspan="2">
		<input type="text" name="buy_date" value="<%= oequip.FOneItem.Fbuy_date %>" size="10" maxlength="10" readonly>
		<a href="javascript:calendarOpen3(frmreg.buy_date,'������',frmreg.buy_date.value)"><img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">���Ű���</td>
	<td colspan="2">
		<input type="text" name="buy_sum" value="<%= oequip.FOneItem.Fbuy_sum %>" size="10" maxlength="9">
		(�ΰ��� ���԰�)
		<!-- <input type="checkbox" name="" value="">�ΰ������� -->
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">�󼼻��</td>
	<td colspan="2">
		
	<!-- ��񱸺п� ���� �ѷ���-->	
		<div id="div_detail_quality1" style="display:none">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80"><span id="detail_quality1_name">CPU :</span></td>
			<td>
				<input type="text" name="detail_quality1" value="<%= oequip.FOneItem.Fdetail_quality1 %>" size="50" maxlength="50">
				<span id="detail_quality1_etc">(ex: P2.8, C2.4, AMD 1800)</span>
			</td>
		</tr>
		</table>
		</div>

		<div id="div_detail_quality2" style="display:none">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80">Memory :</td>
			<td>
				<input type="text" name="detail_quality2" value="<%= oequip.FOneItem.Fdetail_quality2 %>" size="50" maxlength="50">
				(ex: 512M, 1G)
			</td>
		</tr>
		</table>
		</div>
		
		<div id="div_detail_quality3" style="display:none">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80">�ػ� :</td>
			<td>
				<input type="text" name="detail_quality3" value="<%= oequip.FOneItem.Fdetail_quality2 %>" size="50" maxlength="50">
				(ex: 600DPI, 1200DPI)
			</td>
		</tr>
		</table>
		</div>
		
		<div id="div_detail_quality4" style="display:none">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80">���� :</td>
			<td>
				<input type="text" name="detail_quality4" value="<%= oequip.FOneItem.Fdetail_quality2 %>" size="50" maxlength="50">
				(ex: 15��Ʈ���, 5��ƮIP������)
			</td>
		</tr>
		</table>
		</div>
		
		<div id="div_detail_quality5" style="display:none">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80">�뵵 :</td>
			<td>
				<input type="text" name="detail_quality5" value="<%= oequip.FOneItem.Fdetail_quality2 %>" size="50" maxlength="50">
				(ex: ��ǻ�ͺ�ǰ, å)
			</td>
		</tr>
		</table>
		</div>

		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80">��� :</td>
			<td>
				<textarea cols="60" rows="4" name="detail_qualityetc"><%= oequip.FOneItem.Fdetail_qualityetc %></textarea>
			</td>
		</tr>
		<!--<tr>
			<td width="80">IP :</td>
			<td>
				<input type="text" name="detail_ip" value="<%= oequip.FOneItem.Fdetail_ip %>" size="16" maxlength="16">		
				<%' DrawipGubun "equip_gubun" %>			
			</td>
		</tr>-->
		<tr>
			<td></td>
			<td>
				<%' DrawipGubun2 "equip_gubun" %>							
			</td>
		</tr>
		</table>

	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">����� ID</td>
	<td colspan="2">
		<% drawpartneruser "usinguserid", oequip.FOneItem.Fusinguserid ,"" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">��Ÿ����<br>��ǰ��ġ</td>
	<td colspan="2">
		<textarea cols="80" rows="5" name="etc_str"><%= oequip.FOneItem.Fetc_str %></textarea><br>
		<font size="2">(ex : 3�� ������ڸ� ���� �����)</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="3" align="center"><input type="button" value="����" onclick="regEquip(frmreg);" class="button"></td>
</tr>
</form>
</table>

<%
set oequip = Nothing
%>

<script>
	selectChange(frmreg.equip_gubun);
	selectChange(frmreg.part_code);
</script>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->