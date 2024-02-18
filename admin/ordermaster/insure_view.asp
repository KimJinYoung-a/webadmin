<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/myorder_Insurecls.asp"-->
<%
	'// ���� ���� //
	dim OrderIdx
	dim page, searchDiv, searchKey, searchString, param

	dim oInsure, i, lp

	'// �Ķ���� ���� //
	OrderIdx = request("OrderIdx")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

	'// ���� ����
	set oInsure = new CInsure
	oInsure.FRectOrderIdx = OrderIdx

	oInsure.GetInsureRead

%>
<script language="javascript">
<!--
	// ������ ����
	function GotoInsureDel(){
		if (confirm('���ں����� ������ �����Ͻðڽ��ϱ�?\n\n�� 10x10�� ���������� �����Ǵ� ���̹Ƿ� ���� ó���� U-Safe���� �ݵ�� Ȯ�����ֽʽÿ�.')){
			document.frm_trans.mode.value="Del";
			document.frm_trans.submit();
		}
	}

	// ���ں����� �˾�
	function insurePrint(iorderserial, mallid)
	{
		var receiptUrl = "https://gateway.usafe.co.kr/esafe/ResultCheck.asp?oinfo=" + iorderserial + "|" + mallid
		window.open(receiptUrl,"insurePop","width=720,height=500,scrollbars=yes");
	}
//-->
</script>
<!-- ���ں����� ���� ���� -->
<table width="600" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4" height="24" align="left"><b>���ں����� �� ����</b></td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
		<td bgcolor="#FFFFFF" width="180"><%=oInsure.FInsureList(0).Forderserial %></td>
		<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td bgcolor="#FFFFFF"><%=FormatDate(oInsure.FInsureList(0).Fregdate,"0000.00.00")%></td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�ֹ� ǰ��</td>
		<td bgcolor="#F8F8FF" colspan="3">
			<%=db2html(oInsure.FInsureList(0).Fitemname)%>
			<% if Not(oInsure.FInsureList(0).Fipkumdate="" or isnull(oInsure.FInsureList(0).Fipkumdate)) then %>(�Ա��� : <%=FormatDate(oInsure.FInsureList(0).Fipkumdate,"0000.00.00")%>)<% end if %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� �ݾ�</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= CurrFormat(oInsure.FInsureList(lp).FsubtotalPrice) & "��"%></td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td bgcolor="#FFFFFF" colspan="3"><%=oInsure.FInsureList(0).Fbuyname %></td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">������ ��ȭ</td>
		<td bgcolor="#FFFFFF"><%=db2html(oInsure.FInsureList(0).Fbuyphone)%></td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">������ �޴���</td>
		<td bgcolor="#FFFFFF"><%=db2html(oInsure.FInsureList(0).Fbuyhp)%></td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">������ �̸���</td>
		<td bgcolor="#FFFFFF" colspan="3"><%=db2html(oInsure.FInsureList(0).Fbuyemail)%></td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�ֹ� ����</td>
		<td bgcolor="#F8F8FF" colspan="3"><%=NormalIpkumDivName(oInsure.FInsureList(0).Fipkumdiv)%></td>
	</tr>
	<tr><td height="1" colspan="4" bgcolor="#D0D0D0"></td></tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� ���</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<%
				'������
				if oInsure.FInsureList(0).FinsureCd="0" then
			%>
					<font color=darkblue>����</font>
					&nbsp;
					<input type="button" class="button" value="���ں����� ���" onClick="insurePrint('<%=oInsure.FInsureList(0).Forderserial%>','ZZcube1010')">
			<%	else %>
					<font color=darkred>����</font>
			<%	end if %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������ȣ(���)</td>
		<td bgcolor="#FFFFFF" colspan="3"><%=oInsure.FInsureList(0).FinsureMsg%></td>
	</tr>
	<tr><td height="1" colspan="4" bgcolor="#D0D0D0"></td></tr>
	<tr>
		<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
			<input type="button" class="button" value="����" onClick="GotoInsureDel()">
			&nbsp;
			<input type="button" class="button" value="���" onClick="self.location='Insure_list.asp?menupos=<%=menupos & param %>'">
		</td>
	</tr>
<form name="frm_trans" method="POST" action="doInsure.asp">
<input type="hidden" name="OrderIdx" value="<%=OrderIdx%>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
</form>
</table>
<!-- ���ݰ�� ��û�� ���� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
