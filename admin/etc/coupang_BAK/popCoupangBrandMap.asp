<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<%
Dim oCoupang, i, mode, maeipdiv
Dim makerid
makerid	= request("id")

If makerid = "" AND maeipdiv = "" Then
	Call Alert_Close("�귣��ID or ��ü���а��� �����ϴ�.")
	dbget.Close: Response.End
End IF

Set oCoupang = new CCoupang
	oCoupang.FRectMakerid = makerid
	oCoupang.getTenCoupangOneBrandDeliver
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
	function fnSaveForm() {
		var frm = document.frm;
	<% If maeipdiv = "U" Then %>
		if($("#phoneNumber2").val() == ""){
			alert('�߰���ȭ��ȣ�� �Է��ϼ���');
			$("#phoneNumber2").focus();
			return false;
		}
		if($("#deliveryCode").val() == ""){
			alert('�ù�縦 �����ϼ���');
			$("#deliveryCode").focus();
			return false;
		}
		if($("#returnZipCode").val() == ""){
			alert('�����ȣ�� �Է��ϼ���');
			$("#returnZipCode").focus();
			return false;
		}
		if($("#returnAddress").val() == ""){
			alert('�ּ�1�� �Է��ϼ���');
			$("#returnAddress").focus();
			return false;
		}
		if($("#returnAddressDetail").val() == ""){
			alert('�ּ�2�� �Է��ϼ���');
			$("#returnAddressDetail").focus();
			return false;
		}
		if($("#jeju").val() == ""){
			alert('�����갣_���ָ� �Է��ϼ���');
			$("#jeju").focus();
			return false;
		}
		if($("#notJeju").val() == ""){
			alert('�����갣_���ֿܸ� �Է��ϼ���');
			$("#notJeju").focus();
			return false;
		}
	<% End If %>
		if(confirm("�����Ͻðڽ��ϱ�?")) {
			frm.action="procBrandMapping.asp";
			frm.submit();
		}
	}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="maeipdiv" value="<%= maeipdiv %>">
<input type="hidden" name="gubun" value="popup">
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= oCoupang.FItemList(0).FId %></td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">�귣���(�ѱ�)</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= oCoupang.FItemList(0).FSocname_kor %></td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">�귣���(����)</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= oCoupang.FItemList(0).FSocname %></td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">�߰���ȭ��ȣ</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" id="phoneNumber2" name="phoneNumber2" value="<%= oCoupang.FItemList(0).FDeliverPhone %>">
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">�ù��</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<select name="deliveryCode" id="deliveryCode" class="select">
			<option value=""></option>
			<option value="HYUNDAI" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "2", "selected", "") %> >�Ե��ù�</option>
			<option value="KGB" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "18", "selected", "") %> >�����ù�</option>
			<option value="EPOST" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "8", "selected", "") %> >��ü��</option>
			<option value="HANJIN" <%= Chkiif((oCoupang.FItemList(0).FDefaultSongjangDiv = "1") OR (oCoupang.FItemList(0).FDefaultSongjangDiv = "36"), "selected", "") %> >�����ù�</option>
			<option value="CJGLS" <%= Chkiif( (oCoupang.FItemList(0).FDefaultSongjangDiv = "3") OR (oCoupang.FItemList(0).FDefaultSongjangDiv = "4"), "selected", "") %> >CJ�������</option>
			<option value="KDEXP" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "21", "selected", "") %> >�浿�ù�</option>
			<option value="DONGBU" <%= Chkiif((oCoupang.FItemList(0).FDefaultSongjangDiv = "39") OR (oCoupang.FItemList(0).FDefaultSongjangDiv = "41"), "selected", "") %> >�帲�ù�(�� KG������)</option>
			<option value="ILYANG" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "26", "selected", "") %> >�Ͼ��ù�</option>
			<option value="CHUNIL" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "31", "selected", "") %> >õ���ù�</option>
			<option value="AJOU" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "10", "selected", "") %> >�����ù�</option>
			<option value="CSLOGIS" <%= Chkiif((oCoupang.FItemList(0).FDefaultSongjangDiv = "5") OR (oCoupang.FItemList(0).FDefaultSongjangDiv = "24"), "selected", "") %> >SC������</option>
			<option value="DAESIN" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "34", "selected", "") %> >����ù�</option>
			<option value="CVS" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "35", "selected", "") %> >CVS�ù�</option>
			<option value="HDEXP" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "37", "selected", "") %> >�յ��ù�</option>
			<option value="DADREAM">�ٵ帲</option>
			<option value="DHL" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "91", "selected", "") %> >DHL</option>
			<option value="UPS">UPS</option>
			<option value="FEDEX">FEDEX</option>
			<option value="REGISTPOST">������</option>
			<option value="DIRECT">��ü����</option>
			<option value="COUPANG">������ü���</option>
			<option value="EMS">��ü�� EMS</option>
			<option value="TNT">TNT</option>
			<option value="USPS">USPS</option>
			<option value="IPARCEL">i-parcel</option>
			<option value="GSMNTON">GSM NtoN</option>
			<option value="SWGEXP">�����۷ι�</option>
			<option value="PANTOS">�������佺</option>
			<option value="ACIEXPRESS">ACI Express</option>
			<option value="DAEWOON">���۷ι�</option>
			<option value="AIRBOY">������ͽ�������</option>
			<option value="KGLNET">KGL��Ʈ����</option>
			<option value="KUNYOUNG" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "29", "selected", "") %> >�ǿ��ù�</option>
			<option value="SLX">SLX�ù�</option>
			<option value="HONAM" <%= Chkiif(oCoupang.FItemList(0).FDefaultSongjangDiv = "33", "selected", "") %> >ȣ���ù�</option>
			<option value="LINEEXPRESS">LineExpress</option>
			<option value="SFEXPRESS">��ǳ�ù�</option>
			<option value="TWOFASTEXP">2FastsExpress</option>
			<option value="ECMS">ECMS�ͽ�������
		</select>
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">�����ȣ</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" name="returnZipCode" id="returnZipCode" value="<%= oCoupang.FItemList(0).FReturn_zipcode %>">
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">�ּ�1</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" name="returnAddress" id="returnAddress" value="<%= oCoupang.FItemList(0).FReturn_address %>">
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">�ּ�2</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" size="50" name="returnAddressDetail" id="returnAddressDetail" value="<%= oCoupang.FItemList(0).FReturn_address2 %>">
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">�����갣_����</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" name="jeju" id="jeju" value="<%= oCoupang.FItemList(0).FJeju %>">
	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="<%= adminColor("tabletop") %>">�����갣_���ֿ�</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<input type="text" name="notJeju" id="notJeju" value="<%= oCoupang.FItemList(0).FNotJeju %>">
	</td>
</tr>
<tr align="center">
	<td colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="����" onclick="fnSaveForm();">
		<input type="button" class="button" value="���" onclick="self.close();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
