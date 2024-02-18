<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemMaeipSaleMarginShareCls.asp"-->
<%

dim idx
idx     = requestCheckvar(request("idx"),32)


'==============================================================================
dim oMaster, oDetail

set oMaster = new CItemMaeipSaleMarginShare
oMaster.FRectIdx         = idx
oMaster.GetMasterOne

set oDetail = new CItemMaeipSaleMarginShare
oDetail.FPageSize 		= 500
oDetail.FRectIdx    	= idx
if (idx <> "") then
	oDetail.GetDetailList
end if


dim mode, i

mode = "modi"
if oMaster.FResultCount < 1 then
	mode = "ins"
end if

%>
<script language="javascript">
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function jsSubmitSave() {
	var frm = document.frm;

	if (frm.makerid.value == "") {
		alert("�귣�带 �Է��ϼ���.");
		return;
	}

	if (frm.saleCode.value == "") {
		alert("�����ڵ带 �Է��ϼ���.");
		return;
	}

	if (frm.saleCode.value*0 != 0) {
		alert("�����ڵ�� ���ڸ� �Է°��� �մϴ�.");
		return;
	}

	if ((frm.startDate.value == "") || (frm.endDate.value == "")) {
		alert("�Ⱓ�� �Է��ϼ���.");
		return;
	}

	if (frm.defaultMargin.value == "") {
		alert("�⺻������ �Է��ϼ���.");
		return;
	}

	if (frm.defaultMargin.value*0 != 0) {
		alert("�⺻������ ���ڸ� �Է°��� �մϴ�.");
		return;
	}

	if (frm.saleMargin.value == "") {
		alert("���θ����� �Է��ϼ���.");
		return;
	}

	if (frm.saleMargin.value*0 != 0) {
		alert("���θ����� ���ڸ� �Է°��� �մϴ�.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="maeipSaleMargin_process.asp" onSubmit="return false;">
	<input type="hidden" name="mode" value="<%= mode %>">
	<input type="hidden" name="idx" value="<%= oMaster.FOneItem.Fidx %>">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr height="25">
		<td width="7%" bgcolor="<%= adminColor("tabletop") %>" align="center">IDX</td>
		<td width="43%" bgcolor="#FFFFFF"><%= oMaster.FOneItem.Fidx %></td>
		<td width="7%" bgcolor="<%= adminColor("tabletop") %>" align="center"></td>
		<td bgcolor="#FFFFFF"></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����ڵ�</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="saleCode" size="14" maxlength="64" value="<%= oMaster.FOneItem.FsaleCode %>">
			<!--
			<input type="button" class="button" value="��������" onClick="alert('aa');">
			-->
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�귣��</td>
		<td bgcolor="#FFFFFF">
			<%	drawSelectBoxDesignerWithName "makerid", oMaster.FOneItem.Fmakerid %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�Ⱓ(��������)</td>
		<td bgcolor="#FFFFFF">
			������ : <input type="text" name="startDate" size="13" onClick="jsPopCal('startDate');" style="cursor:hand;" value="<%= oMaster.FOneItem.FstartDate %>"  class="text">
			~
			������ : <input type="text" name="endDate" size="13" onClick="jsPopCal('endDate');" style="cursor:hand;" value="<%= oMaster.FOneItem.FendDate %>"  class="text">
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">���ⱸ��</td>
		<td bgcolor="#FFFFFF">
			<select class="select" name="meachulGubun">
				<!-- ����� ������ �� ����ϴ�., skyer9, 2018-03-21
				<option value="1" <%= CHKIIF(oMaster.FOneItem.FmeachulGubun="1", "selected", "")%>>������ ����</option>
				-->
				<option value="2" <%= CHKIIF(oMaster.FOneItem.FmeachulGubun="2", "selected", "")%>>����� ����</option>
			</select>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">
			�⺻����
		</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="defaultMargin" size="4" maxlength="64" value="<%= oMaster.FOneItem.FdefaultMargin %>">% (�Һ��ڰ� ���)
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">
			���θ���
		</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="saleMargin" size="4" maxlength="64" value="<%= oMaster.FOneItem.FsaleMargin %>">% (���ΰ� ���)
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">��뿩��</td>
		<td bgcolor="#FFFFFF">
			<select class="select" name="useyn">
				<option value="Y" <%= CHKIIF(oMaster.FOneItem.Fuseyn="Y", "selected", "")%>>���</option>
				<option value="N" <%= CHKIIF(oMaster.FOneItem.Fuseyn="N", "selected", "")%>>������</option>
			</select>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����</td>
		<td bgcolor="#FFFFFF">
			<%= oMaster.FOneItem.Freguserid %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����</td>
		<td bgcolor="#FFFFFF">
			<%= oMaster.FOneItem.Fregdate %>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">��������</td>
		<td bgcolor="#FFFFFF">
			<%= oMaster.FOneItem.Flastupdate %>
		</td>
	</tr>
	<tr height="50">
		<td bgcolor="#FFFFFF" colspan="4" align="center">
			<input type="button" class="button" value="�����ϱ�" onClick="jsSubmitSave();">
			<input type="button" class="button" value="����ϱ�" onClick="location.href='maeipSaleMarginList.asp?menupos=<%= menupos %>';">
		</td>
	</tr>
	</form>
</table>

<p />

<% if (mode = "modi") then %>
�� ���Ի�ǰ�� ���밡���մϴ�.<br />
�� ���������� �ٹ����ٺδ㸸 ��� �����մϴ�.(�̹� ������ ��ǰ�̶� ������� �ȵǰ�, �Ǹ�������� �����մϴ�.)<br />
�� �����ڵ�(<%= oMaster.FOneItem.FsaleCode %>) �� ��ϵ� �귣��(<%= oMaster.FOneItem.Fmakerid %>) ��ǰ�� <font color="red">�⺻���� <%= oMaster.FOneItem.FdefaultMargin %>%</font> �̰�, <font color="red">�⺻���԰�:���θ��԰� ������</font> ��ǰ�� �߰��˴ϴ�.

<p />

<% end if %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td colspan="17" align="left">�˻���� : <b><%= oDetail.FResultCount %> / <%= oDetail.FTotalCount %></b></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" width="80">��ǰ�ڵ�</td>
		<td align="center" width="55">�̹���</td>
		<td align="center">�귣��</td>
		<td align="center">��ǰ��</td>
		<td align="center" width="60">���<br>����</td>

		<td align="center" width="80">�Һ��ڰ�</td>
		<td align="center" width="80"><b>�⺻<br />���԰�</b></td>
		<td align="center" width="80">�⺻<br />������</td>

		<td align="center" width="80">���ΰ�</td>
		<td align="center" width="80">���η�</td>
		<td align="center" width="80"><b>��������<br />���԰�</b></td>
		<td align="center" width="80">��������<br />������</td>
		<td align="center" width="80"><b>�Ǹ�<br />�����</b></td>
		<td>���</td>
	</tr>
<% if oDetail.FresultCount > 0 then %>
   	<% for i=0 to oDetail.FresultCount-1 %>
	<tr align="center">
		<td bgcolor="#FFFFFF"><%= oDetail.FItemList(i).Fitemid %></td>
		<td bgcolor="#FFFFFF">
			<img src="<%= oDetail.FItemList(i).Fsmallimage %>">
		</td>
		<td bgcolor="#FFFFFF"><%= oDetail.FItemList(i).Fmakerid %></td>
		<td bgcolor="#FFFFFF"><%= oDetail.FItemList(i).Fitemname %></td>
		<td bgcolor="#FFFFFF">
			<%= oDetail.FItemList(i).getMwDivName %>
			<% if (oDetail.FItemList(i).FmwDiv <> oDetail.FItemList(i).Fcurrmwdiv) then %>
			<br />(��:<%= oDetail.FItemList(i).getCurrMwDivName %>)
			<% end if %>
		</td>
		<td bgcolor="#FFFFFF"><%= FormatNumber(oDetail.FItemList(i).Forgprice,0) %></td>
		<td bgcolor="#FFFFFF"><b><%= FormatNumber(oDetail.FItemList(i).ForgBuyCash,0) %></b></td>
		<td bgcolor="#FFFFFF">
			<%= (oDetail.FItemList(i).Forgprice - oDetail.FItemList(i).ForgBuyCash) / oDetail.FItemList(i).Forgprice * 100 %>%
		</td>
		<td bgcolor="#FFFFFF"><%= FormatNumber(oDetail.FItemList(i).Fsaleprice,0) %></td>
		<td bgcolor="#FFFFFF">
			<%= (oDetail.FItemList(i).Forgprice - oDetail.FItemList(i).Fsaleprice) / oDetail.FItemList(i).Forgprice * 100 %>%
		</td>
		<td bgcolor="#FFFFFF"><b><%= FormatNumber(oDetail.FItemList(i).FsaleBuyCash,0) %></b></td>
		<td bgcolor="#FFFFFF">
			<%= (oDetail.FItemList(i).Fsaleprice - oDetail.FItemList(i).FsaleBuyCash) / oDetail.FItemList(i).Fsaleprice * 100 %>%
		</td>
		<td bgcolor="#FFFFFF"><b><%= FormatNumber((oDetail.FItemList(i).ForgBuyCash - oDetail.FItemList(i).FsaleBuyCash),0) %></b></td>
		<td bgcolor="#FFFFFF"></td>
	</tr>
	<% next %>
<% end if %>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
