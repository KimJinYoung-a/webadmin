<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%

'// http://webadmin.10x10.co.kr/cscenter/ordermaster/ordermaster_detail.asp?orderserial=12021576159 ���� �������� �߱� -> ���ݰ�꼭 ���࿡�� ���Ǵ� ������

'// ���� ���� //
dim mode

dim taxIdx
dim sdate, edate, chkTerm
dim page, searchDiv, searchKey, searchString, param

dim oTax, i, lp, strSql

dim orderserial

orderserial = request("orderserial")



'==============================================================================
dim ojumun
set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if

if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if

dim taxwritedate
if (ojumun.FResultCount > 0) then
	taxwritedate = getMayTaxDate(ojumun.FOneItem.Fipkumdate)
end if



'==============================================================================
set oTax = new CTax
oTax.FCurrPage = 1
oTax.FPageSize = 100
'oTax.FRectsearchDiv = "Y"					'����� ������
oTax.FRectsearchBilldiv = "01"				'�Һ��ڸ���
oTax.FRectsearchKey = "t1.userid"

if (ojumun.FOneItem.FUserID <> "") then
	oTax.FRectsearchString = ojumun.FOneItem.FUserID
else
	oTax.FRectsearchString = "----"
end if

oTax.GetTaxList



'==============================================================================
dim itemNames, totalRealPrice


strSql =	"select ( select " &_
		"			Case " &_
		"				When count(idx)>1 Then max(itemname) + '�� ' + Cast((count(idx)-1) as varchar) + '��' " &_
		"				Else max(itemname) " &_
		"			End " &_
		"		from db_order.[dbo].tbl_order_detail " &_
		"		where orderserial='" & orderserial & "' and itemid<>0 and cancelyn='N' group by orderserial " &_
		"	) as itemname " &_
		"	, subtotalprice, accountdiv, IsNull(sumPaymentEtc, 0) as sumPaymentEtc " &_
		"from db_order.[dbo].tbl_order_master " &_
		"Where orderserial = '" & orderserial & "'"
rsget.Open strSql, dbget, 1

if Not(rsget.EOF or rsget.BOF) then
	itemNames = rsget("itemname")

	if (CStr(rsget("accountdiv")) = "7") or (CStr(rsget("accountdiv")) = "20") then
		'������, �ǽð���ü : ��ü�ݾ�
		totalRealPrice = rsget("subtotalprice")
	else
		'���������ݾ׸�
		totalRealPrice = rsget("sumPaymentEtc")
	end if
end if
rsget.close



%>
<script language="javascript">
<!--
	function jsPopCal(fName,sName)
	{
		var fd = eval("document."+fName+"."+sName);

		if(fd.readOnly==false)
		{
			var winCal;
			winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
			winCal.focus();
		}
	}

	// ����ڵ���� Ȯ�� ó��
	function chkSheetOk(){
		if (confirm('����ڵ������ Ȯ���ϼ̽��ϱ�?')){
			document.frm_trans.mode.value="BusiOk";
			document.frm_trans.submit();
		}
	}

	// ��û�� ��� ó��
	function GotoTaxPrint(){
	    alert('�׿���Ʈ�� ���̻� �������� �ʽ��ϴ�.');
	    return;
		if (confirm('���ݰ�꼭�� �����Ͻðڽ��ϱ�?')){
			document.frm_trans.mode.value="sheetOk";
			document.frm_trans.submit();
		}
	}

	// ��û�� ����
	function GotoTaxDel(){
		if (confirm('��û���� ���� �Ͻðڽ��ϱ�?\n\n��꼭�� ����� ��� ������ ��ҵ��� �����Ͻñ� �ٶ��ϴ�.')){
			document.frm_trans.mode.value="sheetDel";
			document.frm_trans.submit();
		}
	}

	// ���ݰ�꼭 ����
	function goView(tax_no, b_biz_no, s_biz_no)
	{
		<% if (application("Svr_Info")="Dev") then %>
			// �׽�Ʈ
			window.open("http://ifs.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% else %>
			// �Ǽ���
			window.open("http://web1.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + tax_no + "&cur_biz_no="+b_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% end if %>
	}

	function goView2(tax_no, b_biz_no, s_biz_no){
		<% if (application("Svr_Info")="Dev") then %>
			// �׽�Ʈ
			window.open("http://ifs.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&cur_biz_no="+s_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% else %>
			// �Ǽ���
			window.open("http://web1.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+tax_no+"&cur_biz_no="+s_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
		<% end if %>
	}

	function goView_Bill36524(tax_no, b_biz_no)
	{
			window.open("http://www.bill36524.com/popupBillTax.jsp?NO_TAX=" + tax_no + "&NO_BIZ_NO="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
	}

	function setRegisterInfo()
	{
		// alert(frm.reg_div.value);

		if((frm.reg_div.value == "01") || (frm.reg_div.value == "03") || (frm.reg_div.value == "51")) {
			// ������ �ٹ�����
			// <option value="01">�Һ���(customer)</option>
			// <option value="03">���θ��(promotion)</option>
			// <option value="51">��Ÿ����(accounts)</option>

			// ================================================================
			// cs_taxsheetcls.asp ���� �����´�.
			// ================================================================
			frm.reg_socno.value = "<%= TENBYTEN_SOCNO %>";
			// frm.reg_subsocno.value = "<%= TENBYTEN_SUBSOCNO %>";
			frm.reg_socname.value = "<%= TENBYTEN_SOCNAME %>";
			frm.reg_ceoname.value = "<%= TENBYTEN_CEONAME %>";
			frm.reg_socaddr.value = "<%= TENBYTEN_SOCADDR %>";
			frm.reg_socstatus.value = "<%= TENBYTEN_SOCSTATUS %>";
			frm.reg_socevent.value = "<%= TENBYTEN_SOCEVENT %>";
			frm.reg_managername.value = "<%= TENBYTEN_MANAGERNAME %>";
			frm.reg_managerphone.value = "<%= TENBYTEN_MANAGERPHONE %>";
			frm.reg_managermail.value = "<%= TENBYTEN_MANAGERMAIL %>";
		}

		if(frm.reg_div.value == "52") {
			// ������ (��)���ش�����

			frm.reg_socno.value = "101-85-29011";
			frm.reg_socname.value = "(��)���ش�����";
			frm.reg_ceoname.value = "�̹���";
			frm.reg_socaddr.value = "���� ���α� ��ȭ�� 197-1 �̿������� 2��";
			frm.reg_socstatus.value = "����,���Ҹ�";
			frm.reg_socevent.value = "���ڻ�ŷ� ��";
			frm.reg_managername.value = "����";
			frm.reg_managerphone.value = "02-554-2033";
			frm.reg_managermail.value = "accounts@10x10.co.kr";
		}

		if(frm.reg_div.value == "53") {
			// ������ (��)���̶��

			frm.reg_socno.value = "101-85-36109";
			frm.reg_socname.value = "(��)���̶��";
			frm.reg_ceoname.value = "�̹���";
			frm.reg_socaddr.value = "���� ���α� ������ 1-45 �������� 4��";
			frm.reg_socstatus.value = "���Ҹ�";
			frm.reg_socevent.value = "�ҽÿ�ǰ";
			frm.reg_managername.value = "����";
			frm.reg_managerphone.value = "02-554-2033";
			frm.reg_managermail.value = "accounts@10x10.co.kr";
		}
	}

	function SearchSocno() {
		if (frm.socno.value == "") {
			alert("����ڹ�ȣ�� �Է��ϼ���.");
			return;
		}

		if (frm.socno.value.length != 12) {
			alert("����ڹ�ȣ�� �Ʒ��� ���� �������� �Է��ϼ���.\n\n000-00-00000");
			return;
		}

		icheckframe.location.href="isearchframe.asp?socno=" + frm.socno.value;
		// location.href="isearchframe.asp?socno=" + frm.socno.value;
	}

	function setCompanyInfo(socname, ceoname, socaddr, socstatus, socevent, managername, managerphone, managermail)
	{
		frm.socname.value = socname;
		frm.ceoname.value = ceoname;
		frm.socaddr.value = socaddr;
		frm.socstatus.value = socstatus;
		frm.socevent.value = socevent;
		frm.managername.value = managername;
		frm.managerphone.value = managerphone;
		frm.managermail.value = managermail;
	}

	function CalcPrice()
	{
		if (frm.totalsuply.value == "") { return; }

		if (frm.taxtype.value.length<1){
			alert('���������� �Է��ϼ���.');
			return;
		}

		if (frm.totalsuply.value*0 != 0) { alert("�߸��� ���� �Է��߽��ϴ�."); return; }

		frm.totalsuply2.value = frm.totalsuply.value;
		frm.totalsuplysum.value = frm.totalsuply.value;

		if (frm.taxtype.value == "Y") {
			frm.totaltax.value = parseInt(frm.totalsuply.value*0.1);
			frm.totaltaxsum.value = parseInt(frm.totalsuply.value*0.1);
		} else {
			frm.totaltax.value = 0;
			frm.totaltaxsum.value = 0;
		}

		frm.totalpricesum.value = parseInt(frm.totalsuply.value) + parseInt(frm.totaltaxsum.value);
		frm.totalpricesum2.value = parseInt(frm.totalsuply.value) + parseInt(frm.totaltaxsum.value);
		frm.totalpricesum3.value = parseInt(frm.totalsuply.value) + parseInt(frm.totaltaxsum.value);
	}

	function CalcPriceWithPrice()
	{
		if (frm.totalpricesum.value == "") { return; }

		if (frm.taxtype.value.length<1){
			alert('���������� �Է��ϼ���.');
			return;
		}

		if (frm.totalpricesum.value*0 != 0) { alert("�߸��� ���� �Է��߽��ϴ�."); return; }

		frm.totalpricesum2.value = frm.totalpricesum.value;
		frm.totalpricesum3.value = frm.totalpricesum.value;

		if (frm.taxtype.value == "Y") {
			// ������ ���ް��� ���ϰ� 0.1 �� �ݿø� ���ָ� �ȴ�.
			frm.totaltax.value = Math.round(1.0 * frm.totalpricesum.value / 1.1 / 10.0);
			frm.totaltaxsum.value = frm.totaltax.value;
		} else {
			frm.totaltax.value = 0;
			frm.totaltaxsum.value = 0;
		}

		frm.totalsuply.value = frm.totalpricesum.value - frm.totaltax.value;
		frm.totalsuply2.value = frm.totalsuply.value;
		frm.totalsuplysum.value = frm.totalsuply.value;
	}


	function doRegisterSheet(){

		if(frm.reg_div.value == "0") {
			alert('�����ڸ� �����ϼ���.');
			return;
		}

		if (frm.socname.value.length<1){
			alert('����� ��ϻ��� ȸ����� �Է��ϼ���.');
			frm.socname.focus();
			return;
		}

		if (frm.ceoname.value.length<1){
			alert('����� ��ϻ��� ��ǥ�ڸ��� �Է��ϼ���.');
			frm.ceoname.focus();
			return;
		}

		if (frm.socno.value.length<1){
			alert('����� ��� ��ȣ�� �Է��ϼ���.');
			frm.socno.focus();
			return;
		}

		if (frm.socno.value.length != 12) {
			alert("����ڹ�ȣ�� �Ʒ��� ���� �������� �Է��ϼ���.\n\n000-00-00000");
			return;
		}

		if (frm.socaddr.value.length<1){
			alert('����� ��ϻ��� �ּҸ� �Է��ϼ���.');
			frm.socaddr.focus();
			return;
		}

		if (frm.socstatus.value.length<1){
			alert('����� ��ϻ��� ���¸� �Է��ϼ���.');
			frm.socstatus.focus();
			return;
		}

		if (frm.socevent.value.length<1){
			alert('����� ��ϻ��� ������ �Է��ϼ���.');
			frm.socevent.focus();
			return;
		}

		if (frm.managername.value.length<1){
			alert('����� ������ �Է��ϼ���.');
			frm.managername.focus();
			return;
		}

		if (frm.managerphone.value.length<1){
			alert('����� ��ȭ��ȣ�� �Է��ϼ���.');
			frm.managerphone.focus();
			return;
		}

		if (frm.managermail.value.length<1){
			alert('����� �̸����ּҸ� �Է��ϼ���.');
			frm.managermail.focus();
			return;
		}

		if (frm.yyyymmdd_register.value.length<1){
			alert('�ۼ����� �Է��ϼ���.');
			return;
		}

		if (frm.itemname.value.length<1){
			alert('ǰ���� �Է��ϼ���.');
			return;
		}

		if (frm.totalsuply.value.length<1){
			alert('�ܰ��� �Է��ϼ���.');
			return;
		}

		if (frm.taxtype.value.length<1){
			alert('���������� �Է��ϼ���.');
			return;
		}

		if(frm.reg_div.value == "01") {
			if(frm.etcstring.value == "") {
				alert('��� �ֹ���ȣ �Ǵ� ����ڵ带 �Է��ϼ���.');
				return;
			}
		} else if (frm.etcstring.value != "") {
			alert('�Һ��ڸ��⿡�� ��� �ֹ���ȣ �Ǵ� ����ڵ带 ���� �� �ֽ��ϴ�.');
			return;
		}



	    if (confirm('���ݰ�꼭 �����û�� �Ͻðڽ��ϱ�?')){
	        document.frm.submit();
	    }
	}

function popListPreviousCustomerTaxSheet(userid){
    var popwin=window.open("/cscenter/taxsheet/popListPreviousCustomerTaxSheet.asp?userid=" + userid,"popListPreviousCustomerTaxSheet","width=700,height=400,scrollbars=yes,resizable=yes");
    popwin.focus();
}

//-->
</script>

<style type="text/css">
.Readonlybox { border:0px; }
.writebox { border:10px; background:#E6E6E6; }
</style>



<table width="800" border="0" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td>

		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
			<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td colspan="2" align="left">
					<b>���ݰ�꼭 �����û</b>
				</td>
			</tr>
			<tr height="25">
				<td align="center" width="120" bgcolor="#F0F0FD">��û��</td>
				<td bgcolor="#FFFFFF"><%= session("ssBctId") %></td>
			</tr>
		</table>

	</td>
</tr>
<tr height="20">
	<td>
	</td>
</tr>
<tr>
	<td>

		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			<tr valign="top">
		        <td width="49%">
		        	<!-- ���������� ���� -->
		        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    			<form name="frm" method="post" onsubmit="return false;" action="doTaxOrder.asp">
		    			<input type=hidden name=mode value="tax_register_new">
		    			<input type=hidden name=sellBizCd value="0000000101">
		    			<input type=hidden name=selltype value="20166">
		    			<input type=hidden name=taxissuetype value="C">
		    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		        			<td colspan="2" height="25"><b>������ ����</b></td>
		        			<td colspan="2">
		        				<select class="select" name="reg_div" onchange="setRegisterInfo()">
		        					<option value="01">�Һ���(customer)</option>
		        					<option value="03">���θ��(promotion)</option>
		        					<option value="51">��Ÿ����(accounts)</option>
		        					<option value="52">���Ʒ���</option>
		        					<option value="53">���̶��</option>
		        				</select>
		        			</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">��Ϲ�ȣ</td>
		        			<td colspan="3"><input type=text name="reg_socno" size=12 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td width="70" bgcolor="#F0F0FD" height="25">��ȣ</td>
		        			<td><input type=text name="reg_socname" size=14 value="" border=0 class="readonlybox" readonly></td>
		        			<td width="70" bgcolor="#F0F0FD">��ǥ��</td>
		        			<td><input type=text name="reg_ceoname" size=8 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">������ּ�</td>
		        			<td colspan="3"><input type=text name="reg_socaddr" size=40 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">����</td>
		        			<td><input type=text name="reg_socstatus" size=14 value="" class="readonlybox" readonly></td>
		        			<td bgcolor="#F0F0FD">����</td>
		        			<td><input type=text name="reg_socevent" size=14 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">�����</td>
		        			<td><input type=text name="reg_managername" size=14 value="" class="readonlybox" readonly></td>
		        			<td bgcolor="#F0F0FD">����ó</td>
		        			<td><input type=text name="reg_managerphone" size=14 value="" class="readonlybox" readonly></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">�̸���</td>
		        			<td colspan=3><input type=text name="reg_managermail" size=40 value="" class="readonlybox" readonly></td>
		        		</tr>
		        	</table>
		        	<!-- ���������� �� -->
		        </td>
		        <td>&nbsp;</td>
		        <td width="49%">
		        	<!-- ���޹޴������� ���� -->
		        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		        			<td colspan="4" height="25"><b>���޹޴��� ����</b></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">��Ϲ�ȣ</td>
		        			<td colspan="3">
		        				<input type=text name="socno" size=12 value="" class="writebox">
		        				<input type="button" class="button_s" value="��������(<%= oTax.FTotalCount %>)" onClick="popListPreviousCustomerTaxSheet('<%= ojumun.FOneItem.FUserID %>')" <% if (oTax.FTotalCount < 1) then %>disabled<% end if %>>
		        			</td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td width="70" bgcolor="#F0F0FD" height="25">��ȣ</td>
		        			<td align="left"><input type=text name="socname" size=14 value="" border=0 class="writebox"></td>
		        			<td width="70" bgcolor="#F0F0FD">��ǥ��</td>
		        			<td align="left"><input type=text name="ceoname" size=14 value="" class="writebox"></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">������ּ�</td>
		        			<td align="left" colspan="3"><input type=text name="socaddr" size=40 value="" class="writebox"></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">����</td>
		        			<td align="left"><input type=text name="socstatus" size=14 value="" class="writebox"></td>
		        			<td bgcolor="#F0F0FD">����</td>
		        			<td align="left"><input type=text name="socevent" size=14 value="" class="writebox"></td>
		        		</tr>
		        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">�����</td>
		        			<td align="left"><input type=text name="managername" size=14 value="" class="writebox"></td>
		        			<td bgcolor="#F0F0FD">����ó</td>
		        			<td align="left"><input type=text name="managerphone" size=14 value="" class="writebox"></td>
		        		</tr>
		        		<tr align="center" bgcolor="#FFFFFF">
		        			<td bgcolor="#F0F0FD" height="25">�̸���</td>
		        			<td align="left" colspan="3"><input type=text name="managermail" size=40 value="" class="writebox"></td>
		        		</tr>
		        	</table>
		        	<!-- ���޹޴������� �� -->
		        </td>
			</tr>
		</table>

	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    <tr align="center" bgcolor="#F0F0FD">
				<td width="120" height="25">�ۼ���</td>
				<td width="100">���ް���</td>
				<td width="100">��������</td>
				<td width="100">����</td>
				<td width="100">�հ�ݾ�</td>
				<td>���</td>
			</tr>
		    <tr align="center" bgcolor="#FFFFFF">
				<td height="25"><input type="text" size="10" name="yyyymmdd_register" value="<%= taxwritedate %>" onClick="jsPopCal('frm','yyyymmdd_register');" style="cursor:hand;" class="writebox"></td>
				<td><input type=text name="totalsuplysum" size=10 value="" class="readonlybox" readonly></td>
				<td>
					<select name=taxtype class="writebox" onchange="CalcPriceWithPrice()">
					<option value="Y">����</option>
					<option value="N">�鼼</option>
					<option value="0">����</option>
					</select>
				</td>
				<td><input type=text name="totaltaxsum" size=10 value="" class="readonlybox" readonly></td>
				<td><input type=text name="totalpricesum" size=10 value="<%= totalRealPrice %>" class="writebox" onkeyup="CalcPriceWithPrice()"></td>
				<td><input type=text name="etcstring" size=30 value="<%= orderserial %>" class="writebox">
				</td>
			</tr>
		</table>

	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    <tr align="center" bgcolor="#F0F0FD">
				<td width="30" height="25">��</td>
				<td width="30">��</td>
				<td>ǰ��</td>
				<td width="50">�԰�</td>
				<td width="50">����</td>
				<td width="100">�ܰ�</td>
				<td width="100">���ް���</td>
				<td width="100">����</td>
			</tr>
		    <tr align="center" bgcolor="#FFFFFF">
				<td height="25"></td>
				<td></td>
				<td><input type=text name="itemname" size=40 value="<%= itemNames %>" class="writebox"></td>
				<td></td>
				<td>1</td>
				<td><input type=text name="totalsuply" size=10 value="" class="writebox" onkeyup="CalcPrice()"></td>
				<td><input type=text name="totalsuply2" size=10 value="" class="readonlybox" readonly></td>
				<td><input type=text name="totaltax" size=10 value="" class="readonlybox" readonly></td>
			</tr>
		</table>

	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    <tr align="center" bgcolor="#F0F0FD">
				<td height="25"><b>�հ�ݾ�</b></td>
				<td width="100">����</td>
				<td width="100">��ǥ</td>
				<td width="100">����</td>
				<td width="100">�ܻ�̼���</td>
			</tr>
		    <tr align="center" bgcolor="#FFFFFF">
				<td height="25"><input type=text name="totalpricesum3" size=10 value="" class="readonlybox" readonly></td>
				<td>
				</td>
				<td></td>
				<td></td>
				<td>
					<input type=text name="totalpricesum2" size=10 value="" class="readonlybox" readonly>
				</td>

			</tr>
			<% if (C_ADMIN_AUTH) then %>
			<tr align="right" bgcolor="#FFFFFF">
			    <td height="20" colspan="5">
			    �����ڸ޴� <input type="checkbox" name="nocheckVal">�ݾ�üũ����
			    </td>
			</tr>
			<% end if %>
			</form>
		</table>

	</td>
</tr>
<tr height="5">
	<td>
	</td>
</tr>
<tr>
	<td>

		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
		    <tr align="center">
				<td align="center" height="25">
				  <input type="button" class="button" value="�ۼ�" onClick="doRegisterSheet()">
				  &nbsp;
				  <input type="button" class="button" value="���" onClick="self.location='Tax_list.asp'">


				</td>
			</tr>
		</table>

	</td>
</tr>
</table>

<p>

<iframe src="" name="icheckframe" width="0" height="0" frameborder="1" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>

<script>
function init() {
	setRegisterInfo();
	CalcPriceWithPrice();
}

window.onload = init;
</script>


<!-- ���ݰ�� ��û�� ���� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->