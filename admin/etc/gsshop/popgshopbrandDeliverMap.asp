<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<%
Dim oGSShop, i, mode
Dim makerid
makerid	= request("makerid")

If makerid = "" Then
	Call Alert_Close("�귣��ID�� �����ϴ�.")
	dbget.Close: Response.End
End IF

Set oGSShop = new CGSShop
	oGSShop.FRectMakerid = makerid
	oGSShop.getTengsshopOneBrandDeliver
%>
<script language="javascript">
<!--
	// ��Ī �����ϱ�
	function fnSaveForm() {
		var frm = document.srcFrm;

		if(frm.deliveryCd.value=="") {
			alert("��Ī�� GSShop �ù���ڵ带 �������ּ���.");
			frm.deliveryCd.focus();
			return;
		}

		if(frm.deliveryAddrCd.value=="") {
			alert("��Ī�� GSShop ���/��ǰ���ڵ带 �Է����ּ���.");
			frm.deliveryAddrCd.focus();
			return;
		}

		if(frm.deliveryAddrCd.value.length < 4 ){
			alert("���/��ǰ���ڵ�� 4�ڸ� �Դϴ�. �ٽ� �Է��ϼ���");
			frm.deliveryAddrCd.focus();
			return;
		}

		if(frm.brandcd.value=="") {
			alert("��Ī�� �귣���ڵ带 �Է����ּ���.");
			frm.brandcd.focus();
			return;
		}

		if(frm.brandcd.value.length < 6 ){
			alert("�귣���ڵ�� 6�ڸ� �Դϴ�. �ٽ� �Է��ϼ���");
			frm.brandcd.focus();
			return;
		}

		if(confirm("�����Ͻðڽ��ϱ�?")) {
			frm.action="procgsshop3.asp";
			frm.submit();
		}
	}

	// â�ݱ�
	function fnCancel() {
		if(confirm("�۾��� ����ϰ� â�� �����ðڽ��ϱ�?")) {
			self.close();
		}
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>GSShop �귣�� �ù��/��ǰ�� ��Ī</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<p>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �ٹ����� �귣�� ����</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">�귣��ID</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FUserid%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">�ù��</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FDivname%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">�귣���(�ѱ�)</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FSocname%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">�귣���(����)</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FSocname_kor%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">�����</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FDeliver_name%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">�ּ�</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FReturn_zipcode%>&nbsp;<%=oGSShop.FItemList(0).FReturn_address%>&nbsp;<%=oGSShop.FItemList(0).FReturn_address2%></td>
</tr>
<tr height="25">
	<td width="120" align="center" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><%= ChkIIF(ogsshop.FItemList(0).FMaeipdiv="U","��ü���","�ٹ����ٹ��") %></td>
</tr>
</table>
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> GSShop ��Ī ����</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="srcFrm" method="GET" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="makerid" value="<%=oGSShop.FItemList(0).FUserid%>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td width="120" align="center" bgcolor="#DDDDFF">�ù���ڵ�</td>
	<td bgcolor="#FFFFFF" height="1">
		<select name="deliveryCd" class="select">
			<option value="">-Choice-</option>
			<option value="ZY" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="ZY","selected","") %>>��ü(��ġ)���</option>
			<option value="HF" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="HF","selected","") %>>����(�������չ��)</option>
			<option value="HJ" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="HJ","selected","") %>>�����ù�</option>
			<option value="DH" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="DH","selected","") %>>�������</option>
			<option value="HD" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="HD","selected","") %>>�����ù�</option>
			<option value="EP" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="EP","selected","") %>>��ü���ù�</option>
			<option value="ER" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="ER","selected","") %>>��ü�����</option>
			<option value="CJ" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="CJ","selected","") %>>CJ GLS</option>
			<option value="KG" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="KG","selected","") %>>�����ù�</option>
			<option value="KL" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="KL","selected","") %>>KGB�ù�</option>
			<option value="YC" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="YC","selected","") %>>���ο�ĸ</option>
			<option value="FA" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="FA","selected","") %>>�����ù��ͽ�������</option>
			<option value="SG" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="SG","selected","") %>>SC������(�簡��)</option>
			<option value="KR" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="KR","selected","") %>>�ϳ����ù�</option>
			<option value="IN" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="IN","selected","") %>>�̳������ù�</option>
			<option value="DS" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="DS","selected","") %>>����ù�</option>
			<option value="CI" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="CI","selected","") %>>õ���ù�</option>
			<option value="KD" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="KD","selected","") %>>�浿�ù�</option>
			<option value="HN" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="HN","selected","") %>>ȣ���ù�</option>
			<option value="YY" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="YY","selected","") %>>����ù�</option>
			<option value="99" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="99","selected","") %>>��Ÿ�ù�</option>
			<option value="LE" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="LE","selected","") %>>LG����</option>
			<option value="DZ" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="DZ","selected","") %>>LG����(�������)</option>
			<option value="SE" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="SE","selected","") %>>�Ｚ����</option>
			<option value="DM" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="DM","selected","") %>>�������</option>
			<option value="MB" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="MB","selected","") %>>GS����</option>
			<option value="IY" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="IY","selected","") %>>�Ͼ��ù�</option>
			<option value="GT" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="GT","selected","") %>>GTX�ù�</option>
			<option value="CV" <%= CHKIIF(oGSShop.FItemList(0).FDeliveryCd="CV","selected","") %>>�������ù�</option>
		</select>
	</td>
</tr>
<tr height="25">
    <td width="120" align="center" bgcolor="#DDDDFF">���/��ǰ���ڵ�</td>
	<td bgcolor="#FFFFFF" height="1">
		<input type="text" maxlength="4" name="deliveryAddrCd" value="<%= oGSShop.FItemList(0).FDeliveryAddrCd %>">
	</td>
</tr>
<tr height="25">
    <td width="120" align="center" bgcolor="#DDDDFF">�귣���ڵ�</td>
	<td bgcolor="#FFFFFF" height="1">
		<input type="text" maxlength="6" name="brandcd" value="<%= oGSShop.FItemList(0).FBrandcd %>">
	</td>
</tr>
</table>
</form>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"></td>
    <td valign="bottom" align="right">
		<img src="http://testwebadmin.10x10.co.kr/images/icon_cancel.gif" width="45" height="20" border="0" onclick="fnCancel()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<img src="http://testwebadmin.10x10.co.kr/images/icon_save.gif" width="45" height="20" border="0" onclick="fnSaveForm()" style="cursor:pointer" align="absmiddle">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td colspan="2" background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
</p>
<% Set oGSShop = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
