<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%

'������

1

'// [OFF]����_���Ͱ���>>�������������(����) ���� �����û �ϸ� ������ ������

	'// ���� ���� //
	dim mode

	dim taxIdx, account_idx
	dim sdate, edate, chkTerm
	dim page, searchDiv, searchKey, searchString, param

	dim ofranchulgomaster
	dim ofranchulgojungsan
	dim opartner
	dim ogroup

	dim oTax, i, lp

	dim Ftenten_manager_name
	dim Ftenten_manager_phone
	dim Ftenten_manager_email

	dim Fetcstring

	dim taxtype



	'==========================================================================
	account_idx = request("idx")
	mode = request("mode")
	taxtype = request("taxtype")


	if (mode = "") then
		mode = "02"
	end if

	if (taxtype = "") then
		taxtype = "Y"
	end if

	if (mode = "01") then
		Ftenten_manager_name = "����"
		Ftenten_manager_phone = "02-1644-6030"
		Ftenten_manager_email = "accounts@10x10.co.kr"
	elseif (mode = "02") then
		Ftenten_manager_name = "����"
		Ftenten_manager_phone = "02-554-2033"
		Ftenten_manager_email = "accounts@10x10.co.kr"
	else
		Ftenten_manager_name = "����"
		Ftenten_manager_phone = "02-554-2033"
		Ftenten_manager_email = "accounts@10x10.co.kr"
	end if



	'==========================================================================
	'��������
	set ofranchulgomaster = new CFranjungsan
	ofranchulgomaster.FRectidx = account_idx

	ofranchulgomaster.getOneFranJungsan

	'ofranchulgomaster.FOneItem.Ftotalsum '�� ����ݾ��� �Ѱ��ް��� ��.(�ΰ������Աݾ�)



	'==========================================================================
	'����̵𿡼� �׷��ڵ� ����
	set opartner = new CPartnerUser

	opartner.FCurrpage = 1
	opartner.FPageSize = 100
	opartner.FRectDesignerID = ofranchulgomaster.FOneItem.Fshopid

	opartner.GetPartnerNUserCList



	'==========================================================================
	'�׷��ڵ忡�� ���ݰ�꼭/�������� ���� ����
	set ogroup = new CPartnerGroup

	ogroup.FRectGroupid = opartner.FPartnerList(0).FGroupID

	ogroup.GetOneGroupInfo



	'==========================================================================
	Fetcstring = CStr(account_idx)



	'==========================================================================
	''����� ���ݰ�꼭���� üũ

	set oTax = new CTax

	oTax.FRectsearchKey = " t1.orderidx "
	oTax.FRectsearchString = CStr(account_idx)

	oTax.GetTaxList

	if oTax.FResultCount > 0 then
		if oTax.FTaxList(0).FisueYn="Y" then
			response.write "<script>alert('�̹� ����� ���ݰ�꼭�� �ֽ��ϴ�.\n\n����� �Ͻ÷��� �ŷ�ó�� [��ҿ�û]�� ���� ���ݰ�꼭 ��Ͽ��� [����]�� �����ϼž� �մϴ�');</script>"
		else
			response.write "<script>alert('���������� ���ݰ�꼭�� �ֽ��ϴ�.\n\n����� �Ͻ÷��� ���� ���ݰ�꼭 ��Ͽ��� [����]�� �����ϼž� �մϴ�.');</script>"
		end if
	end if

%>
<script language="javascript">

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

function doRegisterSheet(){
<% if (ogroup.FResultCount = 1) then %>
	<% if oTax.FResultCount > 0 then %>
		<% if oTax.FTaxList(0).FisueYn="Y" then %>
			alert('�̹� ����� ���ݰ�꼭�� �ֽ��ϴ�.\n\n����� �Ͻ÷��� �ŷ�ó�� [��ҿ�û]�� ���� ���ݰ�꼭 ��Ͽ��� [����]�� �����ϼž� �մϴ�');
		<% else %>
			alert('���������� ���ݰ�꼭�� �ֽ��ϴ�.\n\n����� �Ͻ÷��� ���� ���ݰ�꼭 ��Ͽ��� [����]�� �����ϼž� �մϴ�.');
		<% end if %>
	<% else %>
    if (document.frm.yyyymmdd_register.value == "") {
    	alert("�ۼ����� �Է��ϼ���.");
    	return;
    }

    if (confirm('���ݰ�꼭�� �ۼ��Ͻðڽ��ϱ�?')){
        document.frm.submit();
    }
	<% end if %>
<% else %>
	alert("�׷��ڵ尡 �����Ǿ� ���� ���� ��ü�Դϴ�.");
<% end if %>
}

function ChangePage(frm){
    location.href = "?mode=" + frm.mode.value + "&idx=" + <%= account_idx %> + "&taxtype=" + frm.taxtype.value;
}

function CalcPriceWithPrice111()
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
</script>

<!-- ���ݰ�� ��û�� ���� ���� -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<form name="frm" method="post" action="doTaxOrder.asp">
	<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<input type="hidden" name="idx" value="<%= account_idx %>">
		<td colspan="4" align="left">
			<b>������ ���ݰ�꼭 ����</b>
		</td>
	</tr>
</table>

<br>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr valign="top">
        <td width="49%">
        	<!-- ���������� ���� -->
        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4" height="25"><b>������ ����</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">����ڹ�ȣ</td>
        			<td colspan="3"><b>211-87-00620</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">��ȣ</td>
        			<td><b>(��)�ٹ�����</b></td>
        			<td width="70" bgcolor="#F0F0FD">��ǥ��</td>
        			<td><b>�̹���</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">������ּ�</td>
        			<td colspan="3">����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ�����</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">����</td>
        			<td>����,���Ҹ� ��</td>
        			<td bgcolor="#F0F0FD">����</td>
        			<td>���ڻ�ŷ� ��</td>
        		</tr>
        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">�����</td>
        			<td><%= Ftenten_manager_name %></td>
        			<td bgcolor="#F0F0FD">����ó</td>
        			<td><%= Ftenten_manager_phone %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">�̸���</td>
        			<td><%= Ftenten_manager_email %></td>
        			<td bgcolor="#F0F0FD">BILL���̵�</td>
        			<td>
        			 	<select class="select" name="mode" onchange="ChangePage(frm)">
							<option value="02" <% if (mode = "02") then %>selected<% end if %>>������(ACCOUNTS)</option>
							<option value="03" <% if (mode = "03") then %>selected<% end if %>>���θ��(PROMOTION)</option>
						</select>
        			</td>
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
        			<td bgcolor="#F0F0FD" height="25">����ڹ�ȣ</td>
        			<td colspan="3"><b><%= ogroup.FOneItem.Fcompany_no %></b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">��ȣ</td>
        			<td><b><%= ogroup.FOneItem.FCompany_name %></b></td>
        			<td width="70" bgcolor="#F0F0FD">��ǥ��</td>
        			<td><b><%= ogroup.FOneItem.Fceoname %></b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">������ּ�</td>
        			<td colspan="3"><%= ogroup.FOneItem.Fcompany_address %> <%= ogroup.FOneItem.Fcompany_address2 %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">����</td>
        			<td><%= ogroup.FOneItem.Fcompany_uptae %></td>
        			<td bgcolor="#F0F0FD">����</td>
        			<td><%= ogroup.FOneItem.Fcompany_upjong %></td>
        		</tr>
        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">�����</td>
        			<td><%= ogroup.FOneItem.Fjungsan_name %></td>
        			<td bgcolor="#F0F0FD">����ó</td>
        			<td><%= ogroup.FOneItem.Fjungsan_hp %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">�̸���</td>
        			<td colspan="3"><%= ogroup.FOneItem.Fjungsan_email %></td>
        		</tr>
        	</table>
        	<!-- ���޹޴������� �� -->
        </td>
	</tr>
</table>

<p>
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
		<td height="25"><input type="text" size="10" name="yyyymmdd_register" value="" onClick="jsPopCal('frm','yyyymmdd_register');" style="cursor:hand;"></td>
<% if (taxtype = "Y") then %>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum/1.1),0) %></td>
<% else %>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
<% end if %>
		<td>
			<select name=taxtype class="writebox" onchange="ChangePage(frm)">
			<option value="Y" <% if (taxtype = "Y") then %>selected<% end if %>>����</option>
			<option value="N" <% if (taxtype = "N") then %>selected<% end if %>>�鼼</option>
			<option value="0" <% if (taxtype = "0") then %>selected<% end if %>>����</option>
			</select>
		</td>
<% if (taxtype = "Y") then %>
		<td><%= FormatNumber(((ofranchulgomaster.FOneItem.Ftotalsum/1.1)*0.1),0) %></td>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
<% else %>
		<td>0</td>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
<% end if %>
		<td>�ε����ڵ� : <%= Fetcstring %></td>
	</tr>
</table>

<p>

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
		<td><%= ofranchulgomaster.FOneItem.Ftitle %></td>
		<td></td>
		<td>1</td>
<% if (taxtype = "Y") then %>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum/1.1),0) %></td>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum/1.1),0) %></td>
		<td><%= FormatNumber(((ofranchulgomaster.FOneItem.Ftotalsum/1.1)*0.1),0) %></td>
<% else %>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
		<td>0</td>
<% end if %>
	</tr>
</table>

<p>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td height="25"><b>�հ�ݾ�</b></td>
		<td width="100">����</td>
		<td width="100">��ǥ</td>
		<td width="100">����</td>
		<td width="100">�ܻ�̼���</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25"><b><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></b></td>
		<td></td>
		<td></td>
		<td></td>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
	</tr>
	</form>
</table>
<br>

<p>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="25">
		  <input type="button" class="button" value="�ۼ�" onClick="doRegisterSheet()">
		  &nbsp;
		  <input type="button" class="button" value="���" onClick="self.location='Tax_list.asp'">
		</td>
	</tr>
</table>





<!-- ���ݰ�� ��û�� ���� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->