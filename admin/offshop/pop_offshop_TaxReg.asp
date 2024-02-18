<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : ���ݰ�꼭
' History : ������ ����
'			2017.04.12 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim idx
idx = requestCheckVar(request("idx"),10)

dim obj
set obj = new CFranjungsan
obj.FRectidx = idx
obj.getOneFranJungsan


if obj.FResultCount<1 then
	response.write "<script type='text/javascript'>alert('���������� �����ϴ�.');</script>"
	response.write "<script type='text/javascript'>window.close()</script>"
	dbget.close()	:	response.End
end if

if (obj.FoneItem.FStateCd>"0") and (obj.FoneItem.FStateCd<"4") then
	stypename = "���ݰ�꼭"
else
	response.write "<script type='text/javascript'>alert('���ݰ�꼭 Ȥ�� ��꼭�� ���� �����մϴ�. - �̹� ���� �Ͽ��ų� ������ ������ �����ϴ�.');</script>"
	response.write "<script type='text/javascript'>window.close()</script>"
	dbget.close()	:	response.End
end if

dim objShop, ogroup
dim stypename

set objShop = new COffShopChargeUser
objShop.FRectShopID = obj.FoneItem.Fshopid
objShop.GetOffShopList

set ogroup = new CPartnerGroup
ogroup.FRectGroupid = objShop.FItemList(0).Fgroupid
ogroup.GetOneGroupInfo

'rw objShop.FItemList(0).Fgroupid

dim jungsan_hpall, jungsan_hp1,jungsan_hp2,jungsan_hp3
jungsan_hpall = ogroup.FOneItem.Fjungsan_hp
If Not IsNull(jungsan_hpall) Then 
	jungsan_hpall = split(jungsan_hpall,"-")
	if UBound(jungsan_hpall) >= 2 then
		jungsan_hp1 = jungsan_hpall(0)
		jungsan_hp2 = jungsan_hpall(1)
		jungsan_hp3 = jungsan_hpall(2)
	end if
End If 

If IsNull(ogroup.FOneItem.Fcompany_no) Then 
	ogroup.FOneItem.Fcompany_no = ""
End If 

Dim totalCost, supplyCost, vatCost

totalCost	= CDbl(obj.FoneItem.Ftotalsum)
supplyCost	= Round(totalCost / 1.1)
vatCost		= totalCost - supplyCost

%>

<script type='text/javascript'>

function ActTaxReg(frm){
//alert('�������Դϴ�');
//return;
	if (frm.biz_no.value.length!=10){
		alert('����� ��� ��ȣ�� �ùٸ��� �ʰų� ��ϵǾ� ���� �ʽ��ϴ�. - ���������� ������ ����ϼ���.');
		return;
	}

	if (frm.corp_nm.value.length<1){
		alert('����� ���� ��ϵǾ� ���� �ʽ��ϴ�. - ���������� ������ ����ϼ���.');
		return;
	}

	if (frm.ceo_nm.value.length<1){
		alert('��ǥ�� ���� ��ϵǾ� ���� �ʽ��ϴ�. - ���������� ������ ����ϼ���.');
		return;
	}

	if (frm.biz_status.value.length<1){
		alert('���°� ��ϵǾ� ���� �ʽ��ϴ�. - ���������� ������ ����ϼ���.');
		return;
	}

	if (frm.biz_type.value.length<1){
		alert('������ ��ϵǾ� ���� �ʽ��ϴ�. - ���������� ������ ����ϼ���.');
		return;
	}

	if (frm.addr.value.length<1){
		alert('����� �ּҰ� ��ϵǾ� ���� �ʽ��ϴ�. - ���������� ������ ����ϼ���.');
		return;
	}

	if (frm.dam_nm.value.length<1){
		alert('����� ������ ��ϵǾ� ���� �ʽ��ϴ�. - ���������� ������ ����ϼ���.');
		return;
	}

	if (frm.email.value.length<1){
		alert('����� �̸����� ��ϵǾ� ���� �ʽ��ϴ�. - ���������� ������ ����ϼ���.');
		return;
	}

	if (frm.write_date.value.length<1){
		alert('��꼭 ������ �Է� �� ����ϼ���.');
		return;
	}


	if (confirm('<%= stypename %> �� ���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" width="16" height="16" align="absbottom">
        	<strong>���� <%= stypename %> ����</strong>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="pop_offshop_TaxReg_Proc.asp">
	<input type=hidden name=jungsanid value="<%=obj.FoneItem.FIdx%>">
	<input type=hidden name=jungsanname value="<%=obj.FoneItem.Ftitle%>">
	<input type=hidden name=jungsangubun value="OFFSHOP">
	<input type=hidden name=makerid value="<%=obj.FoneItem.Fshopid%>">
	
	<input type=hidden name=biz_no value="<%= replace(replace(ogroup.FOneItem.Fcompany_no,"-","")," ","") %>" >
	<input type=hidden name=corp_nm value="<%= ogroup.FOneItem.FCompany_name %>">
	<input type=hidden name=ceo_nm value="<%= ogroup.FOneItem.Fceoname %>">
	<input type=hidden name=biz_status value="<%= ogroup.FOneItem.Fcompany_uptae %>">
	<input type=hidden name=biz_type value="<%= ogroup.FOneItem.Fcompany_upjong %>">
	
	
	<input type=hidden name=addr value="<%= ogroup.FOneItem.Fcompany_address %> <%= ogroup.FOneItem.Fcompany_address2 %>">
	<input type=hidden name=dam_nm value="<%= ogroup.FOneItem.Fjungsan_name %>">
	<input type=hidden name=email value="<%= ogroup.FOneItem.Fjungsan_email %>">
	<input type=hidden name=hp_no1 value="<%= jungsan_hp1 %>">
	<input type=hidden name=hp_no2 value="<%= jungsan_hp2 %>">
	<input type=hidden name=hp_no3 value="<%= jungsan_hp3 %>">
	
	<input type=hidden name=sb_type value="01"> <!-- ���� 01 ���� 02 -->
	<input type=hidden name=tax_type value="01"> <!-- ���ݰ�꼭 01 -->
	<input type=hidden name=bill_type value="18"> <!-- ���� 01 û�� 18 -->
	<input type=hidden name=pc_gbn value="C"> <!-- ���� P ��� C -->
	
	<input type=hidden name=item_count value="1">
	<input type=hidden name=item_nm value="<%=obj.FoneItem.Ftitle%>">
	<input type=hidden name=item_qty value="1">
	<input type=hidden name=item_price value="<%=supplyCost%>">
	<input type=hidden name=item_amt value="<%=supplyCost%>">
	<input type=hidden name=item_vat value="<%=vatCost%>">
	<input type=hidden name=item_remark value="">
	
	<input type=hidden name=credit_amt value="<%=totalCost%>">

	<!-- DEV 1000394, REAL 244730, ON 261744 -->
<!-- 
	<input type=hidden name=cur_u_user_no value="261744"> 
	<input type=hidden name=cur_dam_nm value="�̹���">
	<input type=hidden name=cur_email value="moon@10x10.co.kr">
	<input type=hidden name=cur_hp_no1 value="000">
	<input type=hidden name=cur_hp_no2 value="000">
	<input type=hidden name=cur_hp_no3 value="0000">
 -->

	<input type=hidden name=cur_u_user_no value="261748">
	<input type=hidden name=cur_dam_nm value="����">
	<input type=hidden name=cur_email value="shyoung@10x10.co.kr">
	<input type=hidden name=cur_hp_no1 value="010">
	<input type=hidden name=cur_hp_no2 value="4260">
	<input type=hidden name=cur_hp_no3 value="0622">

    <tr align="center" bgcolor="#FFFFFF">
		<td colspan="2">
		* 2005�� 3���� �����(������ 3�� 31��)���ʹ� ���� <%= stypename %> ������ ����ϼž� �մϴ�.
		</td>
	</tr>
    <tr bgcolor="<%= adminColor("tabletop") %>">
   		<td height="20" colspan="2">
	   		<img src="/images/icon_arrow_down.gif" width="16" height="16" align="absbottom">
	   		<strong>���� <%= stypename %> ������</strong>
	   		&nbsp;&nbsp;&nbsp;&nbsp;
	   		<a href="http://www.neoport.net" target="_blank"><font color="blue">>>�׿���Ʈ ȸ�������ϱ�</font></a>
   		</td>
 	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			<img src="/images/icon_num01.gif" width="16" height="16" align="absbottom">
			<font color="red"><b>www.neoport.net�� ȸ������(ȸ�����Թ���)</b></font><br>
				&nbsp;&nbsp;1.�׿���Ʈ�� ���ȸ������ ���ᰡ���ϱ�� �ٶ��ϴ�.(����ڹ�ȣ ��Ȯ�� �Է�)<br>
				&nbsp;&nbsp;2.�������� �����Ͻ� �ʿ� �����ϴ�.<br>
			<img src="/images/icon_num02.gif" width="16" height="16" align="absbottom">
			<font color="red"><b>�̿�� ����(�Ǽ� ����)</b></font><br>
				&nbsp;&nbsp;1.�׿���Ʈ�� �α��� ��, �̿�Ḧ �����ϼ���.(�Ǽ� ����)<br>
				&nbsp;&nbsp;2.����ȭ�� ������ ���̴� "����/��ǰ����"�� ���ø� �˴ϴ�.<br>
				&nbsp;&nbsp;3.�Ǵ� �̿��� 200���̸�, ���Ͻô� �Ǽ��� �̸� �����Ͻø� �˴ϴ�.<br>
			<img src="/images/icon_num03.gif" width="16" height="16" align="absbottom">
			<font color="red"><b>����(����)��꼭 ����</b></font><br>
				&nbsp;&nbsp;1.1���� 2���� �Ϸ��Ͻø�, ����(����)��꼭 ������ �����մϴ�.<br>
				&nbsp;&nbsp;2.������ �� �ٹ����� ���ο��� ���ּž� �ڵ�ó���� �˴ϴ�.
		</td>
	</tr>
    <tr bgcolor="<%= adminColor("tabletop") %>">
   		<td colspan="2" height="20" valign="middle">
	   		<img src="/images/icon_arrow_down.gif" width="16" height="16" align="absbottom">
	   		<strong>��ϵ� ��������� Ȯ��</strong>
   		</td>
 	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF" width="30%">����ڸ�</td>
		<td><%= ogroup.FOneItem.FCompany_name %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">��ǥ�ڸ�</td>
		<td><%= ogroup.FOneItem.Fceoname %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">����ڹ�ȣ</td>
		<td><%= ogroup.FOneItem.Fcompany_no %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">��������</td>
		<td><%= ogroup.FOneItem.Fjungsan_gubun %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">����������</td>
		<td><%= ogroup.FOneItem.Fcompany_address %>&nbsp;<%= ogroup.FOneItem.Fcompany_address2 %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">����</td>
		<td><%= ogroup.FOneItem.Fcompany_uptae %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">����</td>
		<td><%= ogroup.FOneItem.Fcompany_upjong %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFFFFF">��꼭������</td>
	<% if (obj.FoneItem.FStateCd>"0") and (obj.FoneItem.FStateCd<"4") then %>
		<td><input type=text name=write_date value="<%=Left(obj.FoneItem.Ftaxdate,10)%>" size="10" maxlength=10 readonly ><a href="javascript:calendarOpen(frm.write_date);"><img src="/images/calicon.gif" border=0 align=absmiddle></a></td>
	<% else %>
		<td><input type=text name=write_date value="<%=Left(obj.FoneItem.Ftaxdate,10)%>" size="10" maxlength=10 readonly style="border:0"></td>
	<% end if %>

	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFFFFF">����ݾ�</td>
		<td><b><%=formatNumber(totalCost,0)%></b> (���ް� : <%=FormatNumber(supplyCost,0) %> �ΰ���: <%=FormatNumber(vatCost,0) %>)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			&nbsp;&nbsp;<b>* �ſ� 12�� ���� ����� : �������</b><br>
			&nbsp;&nbsp;<b>* �ſ� 13�� ���� ����� : �̿�����(�Ա�ó���� �̿�(15��)�˴ϴ�.)</b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">�������ڸ�</td>
		<td><%= ogroup.FOneItem.Fjungsan_name %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">��������E-mail</td>
		<td><%= ogroup.FOneItem.Fjungsan_email %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">�������� �ڵ�����ȣ</td>
		<td><%= ogroup.FOneItem.Fjungsan_hp %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			&nbsp;&nbsp;* ������������ Ȯ���Ͻð�, ���Էµ� ������ ���� ��ü������������ ������ �����Ͻñ� �ٶ��ϴ�.<br>
			&nbsp;&nbsp;* ���������� ������ �Է��Ͻø�, ���ݰ�꼭�� �����Ȳ�� E-mail�� ���ڼ��񽺷� �˷��帳�ϴ�.
		</td>
	</tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<input type=button value="���� <%= stypename %> ����" onClick="ActTaxReg(frm)">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</form>
</table>
<!-- ǥ �ϴܹ� ��-->



<%
set obj = Nothing
set objShop = Nothing
set ogroup = Nothing
%>

<script language=javascript>
function SvcErrMsg(){
    //alert('�̹��� ��꼭 ������ 4�� 14��(��) ���� �����մϴ�. ');
}
window.onload = SvcErrMsg;
</script>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
