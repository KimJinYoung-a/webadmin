<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���� ����� ����Ʈ ���º���
' History : 2009.04.07 ������ ����
'			2011.05.16 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim idx ,oipchulmaster
	idx = requestCheckVar(request("idx"),10)

set oipchulmaster = new CShopIpChul
	oipchulmaster.FRectIdx = idx
	oipchulmaster.GetIpChulMasterList
%>

<script type='text/javascript'>
	
function ModiMaster(frm){

	if (frm.statecd[3].checked){
		if (!calendarOpen4(frm.execdate,'�԰���',frm.execdate.value)) return;
		var ret = confirm('�԰��� : ' + frm.execdate.value + '\n�԰� Ȯ�� �Ͻðڽ��ϱ�?');
		if (ret) {
			frm.submit();
			return;
		}
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

</script>

<table width="100%" cellspacing="1" cellpadding="2" class="a" bgcolor=#3d3d3d>
<form name="frmMaster" method="post" action="/common/offshop/shopipchul_process.asp">
<input type="hidden" name="mode" value="modistate">
<input type="hidden" name="execdate" value="<%= oipchulmaster.FItemList(0).FexecDt %>">
<input type="hidden" name="idx" value="<%= idx %>">
<tr>
	<td width="100" bgcolor="#DDDDFF">����ó</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="chargeid" value="<%= oipchulmaster.FItemList(0).FChargeid %>">
		<%= oipchulmaster.FItemList(0).FChargeid %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">������ </td>
	<td bgcolor="#FFFFFF">
	<input type="hidden" name="shopid" value="<%= oipchulmaster.FItemList(0).FShopid %>">
		<%= oipchulmaster.FItemList(0).FShopid %> (<%= oipchulmaster.FItemList(0).FShopname %>)
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">���ǸŰ�</td>
	<td bgcolor="#FFFFFF"><%= FormatNumber(oipchulmaster.FItemList(0).FTotalSellCash,0) %></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">�Ѱ��ް�</td>
	<td bgcolor="#FFFFFF"><%= FormatNumber(oipchulmaster.FItemList(0).FTotalSuplyCash,0) %></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">�԰�����</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).FScheduleDt %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">�԰���</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).FexecDt %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">������Ȯ����</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).Fshopconfirmdate %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">��üȮ����</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).Fupcheconfirmdate %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">�����</td>
	<td bgcolor="#FFFFFF"><%= oipchulmaster.FItemList(0).FRegDate %></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">�԰����</td>
	<td bgcolor="#FFFFFF">
	<input type="radio" name="statecd" value="-2" <% if oipchulmaster.FItemList(0).Fstatecd="-2" then response.write "checked" %> >�԰��û
	<input type="radio" name="statecd" value="-1" <% if oipchulmaster.FItemList(0).Fstatecd="-1" then response.write "checked" %> >�԰��ûȮ�� 
	<input type="radio" name="statecd" value="0" <% if oipchulmaster.FItemList(0).Fstatecd="0" then response.write "checked" %> >�԰���
	<input type="radio" name="statecd" value="7" <% if oipchulmaster.FItemList(0).Fstatecd="7" then response.write "checked" %> >���� �԰�Ȯ��
	<input type="radio" name="statecd" value="8" <% if oipchulmaster.FItemList(0).Fstatecd="8" then response.write "checked" %> <% if oipchulmaster.FItemList(0).Fstatecd="0" then response.write "disabled" %> >�԰�Ȯ��
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="button" value="�԰���� ����" onClick="ModiMaster(frmMaster)"></td>
</tr>
</form>
</table>

<%
set oipchulmaster = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->