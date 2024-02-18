<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim idx, isfixed
idx = request("idx")
if idx="" then idx=0

dim ojumunmaster, ojumundetail

set ojumunmaster = new COrderSheet
ojumunmaster.FRectIdx = idx
ojumunmaster.GetOneOrderSheetMaster
isfixed = ojumunmaster.FOneItem.IsFixed

set ojumundetail= new COrderSheet
ojumundetail.FRectIdx = idx
ojumundetail.GetOrderSheetDetail

dim yyyymmdd
yyyymmdd = Left(ojumunmaster.FOneItem.Fscheduledate,10)
%>
<script language='javascript'>

<% if (ojumunmaster.FOneItem.FStatecd="0") or (ojumunmaster.FOneItem.FStatecd=" ") then %>
var jumunwait = true;
<% else %>
var jumunwait = false;
<% end if %>

<% if (Left(ojumunmaster.FOneItem.Fbaljucode,2) = "RJ") then %>
var rejumun = true;
<% else %>
var rejumun = false;
<% end if %>

function AddItems(frm){
	if (jumunwait!=true){
		alert('�ֹ����� ���°� �ƴϸ� �����Ͻ� �� �����ϴ�.');
		return;
	}

	if (rejumun == true){
		alert('���ۼ��� �ֹ������� ��ǰ�� �߰��Ҽ� �����ϴ�.(��ǰ�غ����Դϴ�.) \n�ٸ� ��ǰ�� �ֹ��Ͻ÷��� ������ �ֹ����� �ۼ��ϼ���.');
		return;
	}

	var popwin;
	var suplyer;

	if (frm.suplyer.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;

	popwin = window.open('popshopjumunitem.asp?suplyer=' + suplyer + '&idx=' + frm.masteridx.value ,'offjumuninputeditadd','width=880,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ModiThis(frm){
	if (jumunwait!=true){
		alert('�ֹ����� ���°� �ƴϸ� �����Ͻ� �� �����ϴ�.');
		return;
	}

	if (rejumun == true){
		alert('���ۼ��� �ֹ������� ��ǰ�� ���� �Ͻ� �� �����ϴ�.(��ǰ�غ����Դϴ�.) \n�ٸ� ��ǰ�� �ֹ��Ͻ÷��� ������ �ֹ����� �ۼ��ϼ���.');
		return;
	}


	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="modidetail";
		frm.submit();
	}
}

function DelThis(frm){
	if (jumunwait!=true){
		alert('�ֹ����� ���°� �ƴϸ� �����Ͻ� �� �����ϴ�.');
		return;
	}

	if (rejumun == true){
		alert('���ۼ��� �ֹ������� ��ǰ�� ���� �Ͻ� �� �����ϴ�.(��ǰ�غ����Դϴ�.) \n�ٸ� ��ǰ�� �ֹ��Ͻ÷��� ������ �ֹ����� �ۼ��ϼ���.');
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="deldetail";
		frm.submit();
	}
}

function DelMaster(frm){
	if (jumunwait!=true){
		alert('�ֹ����� ���°� �ƴϸ� �����Ͻ� �� �����ϴ�.');
		return;
	}

	if (rejumun == true){
		alert('���ۼ��� �ֹ������� ��ǰ�� ���� �Ͻ� �� �����ϴ�.(��ǰ�غ����Դϴ�.) \n�ٸ� ��ǰ�� �ֹ��Ͻ÷��� ������ �ֹ����� �ۼ��ϼ���.');
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="delmaster";
		frm.submit();
	}
}

function ModiMaster(frm){
	if (jumunwait!=true){
		alert('�ֹ����� ���°� �ƴϸ� �����Ͻ� �� �����ϴ�.');
		return;
	}

	if (rejumun == true){
		alert('���ۼ��� �ֹ������� ��ǰ�� ���� �Ͻ� �� �����ϴ�.(��ǰ�غ����Դϴ�.) \n�ٸ� ��ǰ�� �ֹ��Ͻ÷��� ������ �ֹ����� �ۼ��ϼ���.');
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="modimaster";
		frm.submit();
	}
}

function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner){
	if (iidx!='<%= idx %>'){
		alert('�ֹ����� ��ġ���� �ʽ��ϴ�. �ٽýõ��� �ּ���.');
		return;
	}

	if (rejumun == true){
		alert('���ۼ��� �ֹ������� ��ǰ�� ���� �Ͻ� �� �����ϴ�.(��ǰ�غ����Դϴ�.) \n�ٸ� ��ǰ�� �ֹ��Ͻ÷��� ������ �ֹ����� �ۼ��ϼ���.');
		return;
	}

	frmadd.itemgubunarr.value = igubun;
	frmadd.itemarr.value = iitemid;
	frmadd.itemoptionarr.value = iitemoption;
	frmadd.sellcasharr.value = isellcash;
	frmadd.suplycasharr.value = isuplycash;
	frmadd.buycasharr.value = ibuycash;
	frmadd.itemnoarr.value = iitemno;

	frmadd.submit();
}
</script>
<table width="760" cellspacing="1" class="a" bgcolor=#3d3d3d>
<form name="frmMaster" method="post" action="shopjumun_process.asp">
<input type=hidden name="mode" value="">
<input type=hidden name="masteridx" value="<%= idx %>">
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>����ó</td>
	<td>
	<input type=hidden name="suplyer" value="<%= ojumunmaster.FOneItem.Ftargetid %>">
	<%= ojumunmaster.FOneItem.Ftargetid %>&nbsp;(<%= ojumunmaster.FOneItem.Ftargetname %>)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>����ó</td>
	<td>
	<%= ojumunmaster.FOneItem.Fbaljuid %>&nbsp;(<%= ojumunmaster.FOneItem.Fbaljuname %>)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>�����</td>
	<td><%= ojumunmaster.FOneItem.Fregdate %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>�԰��û��</td>
	<td><input type=text name="yyyymmdd" value="<%= yyyymmdd %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.yyyymmdd);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a> (���ϴ� �԰� ��¥�� �Է��ϼ���.)</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>�������</td>
	<td><font color="<%= ojumunmaster.FOneItem.GetStateColor %>"><%= ojumunmaster.FOneItem.GetStateName %></font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>�� �Һ��ڰ�(�ֹ�)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsellcash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>�� ���԰�(�ֹ�)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsuplycash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>�� �Һ��ڰ�(Ȯ��)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsellcash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>�� ���԰�(Ȯ��)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsuplycash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>��Ÿ��û����</td>
	<td>
	<textarea name=comment cols=80 rows=6><%= ojumunmaster.FOneItem.FComment %></textarea>
	</td>
</tr>
<tr  bgcolor="#FFFFFF">
	<td colspan="2">
	<br>
	* 5�ϳ� ��� : ��ü ��� ��ǰ (�������ͷ� �԰� �Ǵ´�� �������� �߼� �ص帮�ڽ��ϴ�.) <br>
	* ��� ���� : �������� ��� �������� ���� ��ü�� ���ְ� �� �ִ� �����Դϴ�. <br>
				2~3�� ���� �԰� �� �� �ִ� ��ǰ �Դϴ�. ���� �����帮�� ������, <B>���� �ֹ��� �߰�(���ֹ�)</B>�� �ּž� �մϴ�.<br>
	* �Ͻ�ǰ�� : ��ü ���������� ���� ��������� ��ǰ�Դϴ�.(�ܱⰣ ���� �԰� �Ǳ� ����� ��ǰ�Դϴ�.)
	<br>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align=center>
		<input type=button value="����" onclick="ModiMaster(frmMaster)">
		&nbsp;
		<input type=button value="��ü����" onclick="DelMaster(frmMaster)">
	</td>
</tr>
</form>
</table>
<br>
<%

dim i,selltotal, suplytotal
selltotal =0
suplytotal =0
%>
<table width="760" cellspacing="0" class="a" bgcolor=#ffffff>
<tr>
	<td align=right><input type=button value="��ǰ�߰�" onclick="AddItems(frmMaster)"></td>
</tr>
</table>
<table width="760" cellspacing="1" class="a" bgcolor=#3d3d3d>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="right">�ѰǼ�:  <%= ojumundetail.FResultCount %>&nbsp;</td>
	</tr>
	<tr bgcolor="#DDDDFF" align=center>
		<td width="100">���ڵ�</td>
		<td width="100">�귣��</td>
		<td width="200">��ǰ��</td>
		<td width="80">�ɼǸ�</td>
		<td width="80">�ǸŰ�</td>
		<td width="80">���ް�</td>
		<td width="60">�ֹ�����</td>
		<% if isfixed then %>
		<td width="60">Ȯ������</td>
		<td width="60">���</td>
		<% else %>
		<td width="40">����</td>
		<td width="40">����</td>
		<% end if %>
	</tr>
	<% for i=0 to ojumundetail.FResultCount-1 %>
	<%
	selltotal  = selltotal + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno
	suplytotal = suplytotal + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno
	%>
	<form name="frmBuyPrc_<%= i %>" method="post" action="shopjumun_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="masteridx" value="<%= idx %>">
	<input type="hidden" name="itemgubun" value="<%= ojumundetail.FItemList(i).FItemGubun %>">
	<input type="hidden" name="itemid" value="<%= ojumundetail.FItemList(i).FItemID %>">
	<input type="hidden" name="itemoption" value="<%= ojumundetail.FItemList(i).Fitemoption %>">
	<input type="hidden" name="desingerid" value="<%= ojumundetail.FItemList(i).Fmakerid %>">
	<input type="hidden" name="sellcash" value="<%= ojumundetail.FItemList(i).FSellcash %>">
	<input type="hidden" name="suplycash" value="<%= ojumundetail.FItemList(i).FSuplycash %>">
	<input type="hidden" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>">
	<tr bgcolor="#FFFFFF">
		<td ><%= ojumundetail.FItemList(i).FItemGubun %><%= CHKIIF(ojumundetail.FItemList(i).FItemID>=1000000,format00(8,ojumundetail.FItemList(i).FItemID),format00(6,ojumundetail.FItemList(i).FItemID)) %><%= ojumundetail.FItemList(i).Fitemoption %></td>
		<td ><%= ojumundetail.FItemList(i).Fmakerid %></td>
		<td ><%= ojumundetail.FItemList(i).Fitemname %></td>
		<td ><%= ojumundetail.FItemList(i).Fitemoptionname %></td>

		<td align=right><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojumundetail.FItemList(i).Fsuplycash,0) %></td>
		<td align=center><input type="text" name="baljuitemno" value="<%= ojumundetail.FItemList(i).Fbaljuitemno %>"  size="4" maxlength="4"></td>
		<% if isfixed then %>
		<td><%= ojumundetail.FItemList(i).Frealitemno %></td>
		<td><%= ojumundetail.FItemList(i).Fcomment %></td>
		<% else %>
		<td><input type=button value="����" onclick="ModiThis(frmBuyPrc_<%= i %>)"></td>
		<td><input type=button value="����" onclick="DelThis(frmBuyPrc_<%= i %>)"></td>
		<% end if %>
	</tr>
	</form>
	<% next %>

	<% if (ojumundetail.FResultCount>0) then %>
	<tr bgcolor="#FFFFFF">
		<td align="center">�Ѱ�</td>
		<td colspan="3" align="center">
		<td align=right><%= formatNumber(selltotal,0) %></td>
		<td align=right><%= formatNumber(suplytotal,0) %></td>
		<td></td>
		<td></td>
		<td></td>
		</td>
	</tr>
	<% end if %>
</table>
<%
set ojumunmaster = Nothing
set ojumundetail = Nothing
%>
<form name="frmadd" method=post action="shopjumun_process.asp">
<input type=hidden name="mode" value="shopjumunitemaddarr">
<input type=hidden name="masteridx" value="<%= idx %>">
<input type=hidden name="itemgubunarr" value="">
<input type=hidden name="itemarr" value="">
<input type=hidden name="itemoptionarr" value="">
<input type=hidden name="sellcasharr" value="">
<input type=hidden name="suplycasharr" value="">
<input type=hidden name="buycasharr" value="">
<input type=hidden name="itemnoarr" value="">
</form>
<script language='javascript'>
if (jumunwait!=true){
	alert('�ֹ����� ���°� �ƴϸ� �����Ͻ� �� �����ϴ�.');
}else if (rejumun==true){
	alert('���ۼ��� �ֹ����� �����Ͻ� �� �����ϴ�.');
}
</script>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->