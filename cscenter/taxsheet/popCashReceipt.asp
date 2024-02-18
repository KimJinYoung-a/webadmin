<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ������ ���ݿ�����,���ݰ�꼭 ����
' History : �̻� ����
'			2023.07.31 �ѿ�� ����(10x10_cs �ֹ��ǵ� ���ݿ����� ���డ���ϰ� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_cashreceiptcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetReqCls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/cscenter/action/incNaverpayCommon.asp"-->
<%
dim i, idx, orderSerial, ojumun, IsCancelOrder, oreceipt, IsCashReceipExists, IsCashReceipListExists, oreceiptList
dim IsTaxSheetExists, taxSheetRequestType, MxDlvDate, taxidx, oTax
dim sqlStr, IsTaxIdxExist, IsAllTaxReqExist, oCTaxRequest, authcode, cashreceiptReq, accountno
dim IsCreateNewPaperOK, IsDacomCyberAccountPay, IsAcademy, IsOldRealTimePay, IsOldOrder
	idx	= req("idx","")
	orderSerial	= req("orderSerial","")

set ojumun = new COrderMaster
if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if

if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if

If ojumun.FResultCount>0 then
    IsCancelOrder = ojumun.FOneItem.FCancelyn<>"N"
end if

'// ���ݿ����� ����� ���� �ִ��� //�ֱ� 3�� ������ ����..
IsCashReceipExists = False
IsCashReceipListExists = False

set oreceipt = new CCashReceipt
oreceipt.FRectIdx = idx
oreceipt.FRectorderSerial = orderSerial

if (idx<>"") then
    oreceipt.GetOneCashReceipt
else
    if (IsCancelOrder) then
    	''D�÷��� ���� ��ü ����
    else
        oreceipt.FRectCancelyn = "N"
        oreceipt.FRectExcFailData = "Y"			'// ������ ������, 2021-04-27, skyer9
    end if

    oreceipt.GetReceiptByOrderSerial

    ''2015/08/10 �߰� (���ݿ����� ���ų���)
    if (oreceipt.FResultcount<1) then
        if (ojumun.FResultCount>0) then
            if (ojumun.FOneItem.GetPaperType="R") then
                oreceipt.GetReceiptByOrderSerial_OLD
            end if
        end if
    end if
end if

IsCashReceipExists = oreceipt.FResultCount > 0

'' �����/���/��û ���ݿ����� ���� (�ش��ֹ��� ��ü ��ȸ)
set oreceiptList = new CCashReceipt
oreceiptList.FRectorderSerial = orderSerial
oreceiptList.FPageSize = 20
if (IsCashReceipExists) then
    oreceiptList.FRectExceptIdx = oreceipt.FOneItem.Fidx        ''���� �������°��� ǥ�� ����.
end if
if (oreceiptList.FRectorderSerial<>"") then
    oreceiptList.GetReceiptLogList
end if

IsCashReceipListExists = (oreceiptList.Ftotalcount>0)

'// ���ݰ�꼭 ����� ���� �ִ���
IsTaxSheetExists = False
taxSheetRequestType = ""			'// 01 : 2013�� ������ �����, 11 : 2014�� ���� �����(db_order.dbo.tbl_taxSheet ���̺� billdiv ����)
MxDlvDate = ""

taxidx = 0

if (orderSerial <> "") then

	'// ��Ҿȵ� ��ǰ�� ������ ����� ��������
	'// �ٹ����ٸ��� �Ǵ� ��ü������� ��꼭 �߱��Ѵ�.
	sqlStr = " select MAX(convert(Varchar(10),IsNull(d.beasongdate, getdate()),21)) as MxDlvDate "
	sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m with (nolock)"
	sqlStr = sqlStr & " join db_order.dbo.tbl_order_detail d with (nolock)"
	sqlStr = sqlStr & " 	on "
	sqlStr = sqlStr & " 		m.orderserial = d.orderserial "
	sqlStr = sqlStr & " where "
	sqlStr = sqlStr & " 	1 = 1 "
	sqlStr = sqlStr & " 	and m.orderserial = '" + CStr(orderSerial) + "' "
	sqlStr = sqlStr & " 	and d.itemid <> 0 "
	sqlStr = sqlStr & " 	and d.cancelyn<>'Y' "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if Not (rsget.EOF or rsget.BOF) then
			MxDlvDate = rsget("MxDlvDate")
		end if
	rsget.Close

	if (MxDlvDate = "") or IsNull(MxDlvDate) then
		MxDlvDate = Left(Now(), 10)
	end if

	if (MxDlvDate >= "2014-01-01") or (MxDlvDate = "") then
		'// 2014�� ���� : ��ü������
		taxSheetRequestType = "11"

		sqlStr = " select 1 idx "
		sqlStr = sqlStr & " from db_log.dbo.tbl_tax_issue_request with (nolock)"
		sqlStr = sqlStr & " where orderserial = '" + CStr(orderSerial) + "' and useYN = 'Y' "

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if Not (rsget.EOF or rsget.BOF) then
			IsTaxSheetExists = True
		end if
		rsget.Close
	else
		'// 2013�� ���� : �ٹ����ٸ���
		taxSheetRequestType = "01"

		sqlStr = "select taxIdx From db_order.[dbo].tbl_taxSheet with (nolock)"
		sqlStr = sqlStr & " where orderserial = '" + CStr(orderSerial) + "' "
		sqlStr = sqlStr & " and delYn='N'"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if Not (rsget.EOF or rsget.BOF) then
			IsTaxSheetExists = True
			taxidx = CLng(rsget("taxIdx"))
		end if
		rsget.Close
	end if
end if

set oTax = new CTax
oTax.FRecttaxIdx = taxIdx
if (CLng(taxidx)<>0) then
	oTax.GetTaxRead

	if oTax.FREsultCount>0 then
	    taxIdx = CLng(oTax.FOneItem.FtaxIdx)
	end if
end if

'// �����û ��꼭 �ִ���(2014)
IsTaxIdxExist = False
IsAllTaxReqExist = True

set oCTaxRequest = new CTaxRequest
oCTaxRequest.FRectOrderserial = orderSerial
if (taxSheetRequestType = "11") and (IsTaxSheetExists = True) then
	oCTaxRequest.FPageSize = 100
	oCTaxRequest.FRectOrderserial = orderSerial
	oCTaxRequest.GetTaxRequestOneOrder
end if

IsCreateNewPaperOK = false
IsDacomCyberAccountPay=false
IsAcademy = false
IsOldRealTimePay = false			'�ǽð���ü : ���ų����� �ֹ��������� paygatetid ��, 2011�� 4�� 15�� ���Ĵ�

IsOldOrder = false
if (orderSerial<>"") then
    IsCreateNewPaperOK = False

    '(������ or �ǽð�) + �����Ϸ� + ������ ���� ���
    sqlStr = " select orderserial, IsNULL(authcode,'') as authcode, accountno, IsNULL(cashreceiptReq,'') as cashreceiptReq, ipkumdiv from db_order.dbo.tbl_order_master with (nolock)"
    sqlStr = sqlStr & " where orderserial='"&orderSerial&"'"
    sqlStr = sqlStr & " and ipkumdiv>=2"
    ''sqlStr = sqlStr & " and cancelyn='N'"
    sqlStr = sqlStr & " and accountdiv in ('7','20') "
   ''''sqlStr = sqlStr & " and jumundiv<>9"

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    If Not rsget.Eof then
        IsCreateNewPaperOK = rsget("ipkumdiv")>3
        authcode = rsget("authcode")
        cashreceiptReq  = rsget("cashreceiptReq")
        accountno = rsget("accountno")

        '''IsDacomCyberAccountPay = (authcode <> "")
    end if
    rsget.Close

    '(������+�ǽð� �̿�)+�����Ϸ�+������������+������ ���� ���
    sqlStr = " select orderserial, IsNULL(authcode,'') as authcode, accountno, IsNULL(cashreceiptReq,'') as cashreceiptReq, ipkumdiv, IsNull(sumPaymentEtc, 0) as sumPaymentEtc from db_order.dbo.tbl_order_master with (nolock)"
    sqlStr = sqlStr & " where orderserial='"&orderSerial&"'"
    sqlStr = sqlStr & " and ipkumdiv>=2"
    'sqlStr = sqlStr & " and cancelyn='N'"
    sqlStr = sqlStr & " and accountdiv not in ('7', '20') "
  ''''sqlStr = sqlStr & " and jumundiv<>9"

	''response.write sqlStr
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    If Not rsget.Eof then
        IsCreateNewPaperOK = rsget("ipkumdiv")>3
        authcode = rsget("authcode")
        cashreceiptReq  = rsget("cashreceiptReq")
        accountno = rsget("accountno")
    end if
    rsget.Close

    ''' ��ī���� DIY
    '������+�����Ϸ�+������ ���� ���
    IF (Not IsCreateNewPaperOK) and (LEFT(orderSerial,1)="Y") THEN
        sqlStr = " select orderserial, IsNULL(authcode,'') as authcode, accountno, IsNULL(cashreceiptReq,'') as cashreceiptReq, ipkumdiv from db_academy.dbo.tbl_academy_order_master with (nolock)"
        sqlStr = sqlStr & " where orderserial='"&orderSerial&"'"
        sqlStr = sqlStr & " and ipkumdiv>=2"
        sqlStr = sqlStr & " and cancelyn='N'"
        sqlStr = sqlStr & " and accountdiv='7'"
       ''''sqlStr = sqlStr & " and jumundiv<>9"

		rsACADEMYget.CursorLocation = adUseClient
    	rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly, adLockReadOnly
        If Not rsACADEMYget.Eof then
            IsCreateNewPaperOK = rsACADEMYget("ipkumdiv")>3
            authcode = rsACADEMYget("authcode")
            cashreceiptReq  = rsACADEMYget("cashreceiptReq")
            accountno = rsACADEMYget("accountno")

            IsAcademy = true
            IsDacomCyberAccountPay = (authcode <> "")
        end if
        rsACADEMYget.Close
    ENd IF

    IF (not IsCreateNewPaperOK) then
        '�ǽð�+�����Ϸ�+������ ���� ���
        sqlStr = " select orderserial, IsNULL(authcode,'') as authcode, accountno, IsNULL(cashreceiptReq,'') as cashreceiptReq, ipkumdiv from db_log.dbo.tbl_old_order_master_2003 with (nolock)"
        sqlStr = sqlStr & " where orderserial='"&orderSerial&"'"
        sqlStr = sqlStr & " and ipkumdiv>=2"
        'sqlStr = sqlStr & " and cancelyn='N'"
        sqlStr = sqlStr & " and accountdiv in ('7','20')"
       ''''sqlStr = sqlStr & " and jumundiv<>9"

	   ''response.write sqlStr
        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        If Not rsget.Eof then
            IsCreateNewPaperOK = rsget("ipkumdiv")>3
            authcode = rsget("authcode")
            cashreceiptReq  = rsget("cashreceiptReq")
            accountno = rsget("accountno")

			IsOldOrder = true
        end if
        rsget.Close
    End IF
'rw IsCreateCashReceiptOK
'rw isOLDORDER

    '''???
    if (accountno="���� 470301-01-014754") _
            or (accountno="���� 100-016-523130") _
            or (accountno="�츮 092-275495-13-001") _
            or (accountno="�ϳ� 146-910009-28804") _
            or (accountno="��� 277-028182-01-046") _
            or (accountno="���� 029-01-246118") then
        IsDacomCyberAccountPay = false
    end if
end if

'''�����ֹ�����
'IF orderserial="11072931537" or orderserial="11071271646" or orderserial="11070961322" or orderserial="11070859345" or orderserial="11041320639" then
'    IsCreateCashReceiptOK = true
'ENd IF
%>
<html>
<head>
<title>���ݿ�����</title>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script type="text/javascript">

function jsCancel(){
	var f = document.frmWrite;
	if (f.chkPrint.value){
		if (confirm("����� ���ݿ������� ����Ͻðڽ��ϱ�?"))
		{
		    frmWrite.Atype.value='C2';
		    <% if (IsAcademy) then %>
		    frmWrite.Atype.value='CA';
		    <% end if %>
			f.submit();
		}
	}
}

function jsReEvalNCancel(mayPrc){

    var f = document.frmWrite;
	if (f.chkPrint.value){
		if (confirm("�� ����� ���ݿ������� ����� ����� �Ͻðڽ��ϱ�?"))
		{
		    frmWrite.mayPrc.value=mayPrc;
		    frmWrite.Atype.value='RNC';
		    <% if (IsAcademy) then %>
		    frmWrite.Atype.value='RNCA';
		    <% end if %>
			f.submit();
		}
	}
}

function jsReCalcuEvalPrc(){
    var f = document.frmWrite;
	if (f.chkPrint.value){
		if (confirm("���ݿ����� ��û�ݾ��� ���� �Ͻðڽ��ϱ�?"))
		{
		    frmWrite.Atype.value='Recalcu';
		    <% if (IsAcademy) then %>
		    frmWrite.Atype.value='RecalcuA';
		    <% end if %>
			f.submit();
		}
	}
}

function popReceipt(tid){

    var receiptUrl = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?noTid="+tid+"&noMethod=1";

	var popwin = window.open(receiptUrl,"CashreceiptPrt","width=380,height=750,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function jsCashEval(iidx){
    var f = document.frmWrite;
	if (f.chkPrint.value)
	{
		if (confirm("���ݿ������� �����Ͻðڽ��ϱ�?"))
		{
		    frmWrite.Atype.value='R';
		    <% if (IsAcademy) then %>
		    frmWrite.Atype.value='RA';
		    <% end if %>
			f.submit();
		}
	}
}

function popEvalCashRecipt(orderserial,sitenametype){
    var popwin=window.open("/cscenter/receipt/INIreceiptReq.asp?orderserial="+orderserial+"&sitenametype="+sitenametype,"INIreceiptReq","width=900,height=600,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function popEvalCashReciptHand(orderserial){
    var popwin=window.open("/cscenter/receipt/INIreceiptReq.asp?issuetype=orderserial&orderserial=" + orderserial+"&hand=on","INIreceiptReq","width=680,height=480,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function popWriteCustomerTaxSheet(orderserial){
	<%
if (ojumun.FResultCount>0) then
	if (ojumun.FOneItem.Fipkumdiv < 4) then
	''if (ojumun.FOneItem.Fipkumdiv < 7) then
	%>
	    alert("�����Ϸ� ���� �ֹ��Դϴ�.");
	    return;
	<%
	else
		if (MxDlvDate >= "2014-01-01") then
			%>
    var popwin=window.open("/cscenter/taxsheet/tax_view.asp?orderserial=" + orderserial + "&chulgoyear=<%= Left(MxDlvDate, 4) %>","popWriteCustomerTaxSheet","width=850,height=600,scrollbars=yes,resizable=yes");
    popwin.focus();
			<%
		else
			%>
	alert("2013�� ������Դϴ�.\n\n�ٹ����� ���� ���ݰ�꼭�� �����մϴ�.");
    var popwin=window.open("/cscenter/taxsheet/tax_view.asp?orderserial=" + orderserial,"popWriteCustomerTaxSheet","width=850,height=600,scrollbars=yes,resizable=yes");
    popwin.focus();
			<%
		end if
	end if
end if
	%>
}

// �ſ�ī�� ������ǥ �˾�_�̴Ͻý�
function receiptinicis(tid){
	var receiptUrl = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?" +
		"noTid=" + tid + "&noMethod=1";
	var popwin = window.open(receiptUrl,"INIreceipt","width=415,height=600");
	popwin.focus();
}

function fnDelIssueReq() {
	if (confirm('���ݰ�꼭 �����û�� ��� �Ͻðڽ��ϱ�?\n\n����� ��꼭�� �̹� �ִ� ��� ���� ��꼭�� �����ϼ���.')) {
		document.frm.mode.value="delIssueReq";
		document.frm.submit();
	}
}

function fnFinishIssueReq() {
	if (confirm('�Ϸ�ó�� �Ͻðڽ��ϱ�?')) {
		document.frm.mode.value="finishIssueReq";
		document.frm.submit();
	}
}

function fnRegUpcheTax(groupID, itemname, chulgoPrice, taxtype, busiIdx) {
	location.href = '/cscenter/taxsheet/tax_view.asp?orderserial=<%= orderserial %>&groupID=' + groupID + '&itemname=' + itemname + '&chulgoPrice=' + chulgoPrice + '&taxtype=' + taxtype + '&busiIdx=' + busiIdx + '&chulgoyear=2014';
}

function fnViewUpcheTax(taxIdx) {
	location.href = '/cscenter/taxsheet/tax_view.asp?taxIdx=' + taxIdx;
}

</script>
</head>
<body>

<div align="center">
	<b>���ݿ�����/���ݰ�꼭 ���� ����</b>
</div>
<br>
<% if (IsCashReceipListExists) then %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="<%= adminColor("tabletop") %>">
        <td>*����� �α�</td>
        <td>�ݾ�</td>
        <td>����</td>
        <td>����</td>
        <td>������</td>
        <td>����</td>
    </tr>
    <% for i=0 to oreceiptList.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
        <td><%=oreceiptList.FItemList(i).Fidx%></td>
        <td><%=FormatNumber(oreceiptList.FItemList(i).Fcr_price,0)%></td>
        <td><%=LEFT(oreceiptList.FItemList(i).getReceiptType,2)%></td>
        <td><%=oreceiptList.FItemList(i).getStateName%></td>
        <td><%=oreceiptList.FItemList(i).getMayEvalDT%></td>
        <td>
        <% if (Not isNULL(oreceiptList.FItemList(i).Ftid)) then %>
        <input type="button" value="����" onClick="popReceipt('<%=oreceiptList.FItemList(i).Ftid%>');">
        <% end if %>
        </td>
    </tr>
    <% next %>
</table>
<p>
<% end if %>

<% if (IsCashReceipExists = True) then %>
<%
'' ���೻�� ���� 2016/08/12 -----------------------------------------------------------
dim orginPrc, minusSubtotalprice, mayReqPrc
if (ojumun.FOneItem.FAccountDiv="7") or (ojumun.FOneItem.FAccountDiv="20") then   ''2016/09/19
    orginPrc = ojumun.FOneItem.FsubtotalPrice
else
    orginPrc = ojumun.FOneItem.FsumpaymentEtc
end if

if (ojumun.FOneItem.FCancelyn<>"N") then orginPrc=0 '' ���

minusSubtotalprice = GetReceiptMinusOrderSUM(orderserial)

mayReqPrc = orginPrc+minusSubtotalprice

dim isNaverPay, NPay_Result, NpayCashAmt, NpayCashAmt_Only, NpaySuplyAmt, NpaySuplyAmt_Only
isNaverPay = (ojumun.FOneItem.Fpggubun="NP")

if (isNaverPay) then
    Set NPay_Result = fnCallNaverPayCashAmt(ojumun.FOneItem.Fpaygatetid)
    if NPay_Result.code="Success" then
		NpayCashAmt_Only = CLng(NPay_Result.body.totalCashAmount)
        NpayCashAmt    = NpayCashAmt_Only + ojumun.FOneItem.FsumPaymentEtc	'// �� ���ݾ�
		NpaySuplyAmt_Only = CLng(NPay_Result.body.supplyCashAmount)
		NpaySuplyAmt   = NpaySuplyAmt_Only + CLng(ojumun.FOneItem.FsumPaymentEtc*10/11)	'// �� ���ް�
		''i_sup_price   = CLng(NPay_Result.body.supplyCashAmount) + CLng(myorder.FMasterItem.FsumPaymentEtc*10/11)	'// ���ݼ� ���ް�
		''i_tax         = i_cr_price - i_sup_price													'// ���ݼ� ������
    end if
    Set NPay_Result = Nothing
end if

''��Ұ��� ���� �Ұ�.
IsCreateNewPaperOK = IsCreateNewPaperOK AND (ojumun.FOneItem.FCancelyn="N")
''��û���� �����. ����ݾ� üũ

'if (isNaverPay) then
'    IsCreateNewPaperOK = IsCreateNewPaperOK AND (oreceipt.FOneItem.Fcr_price=NpayCashAmt)
'else
'    IsCreateNewPaperOK = IsCreateNewPaperOK AND (oreceipt.FOneItem.Fcr_price=mayReqPrc)
'end if
'' -------------------------------------------------------------------------------------
%>
	<!-- ���ݿ����� -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center">
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
			<td align="left" bgcolor="#FFFFFF"><%=oreceipt.FOneItem.Fgoodname%></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">��û�ݾ�</td>
			<td align="left" bgcolor="#FFFFFF"><%=CurrFormat(oreceipt.FOneItem.Fcr_price)%></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">�����ڸ�</td>
			<td align="left" bgcolor="#FFFFFF"><%=oreceipt.FOneItem.Fbuyername%></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">�����ڸ���</td>
			<td align="left" bgcolor="#FFFFFF"><a href="mailto:<%=oreceipt.FOneItem.Fbuyeremail%>"><%=oreceipt.FOneItem.Fbuyeremail%></a></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">�ĺ���ȣ</td>
			<td align="left" bgcolor="#FFFFFF">
			<% if (oreceipt.FOneItem.Freg_num="0100001234") then %>
			    <%= oreceipt.FOneItem.Freg_num %> (�����߱�)
			<% else %>
				<%= Left(oreceipt.FOneItem.Freg_num,6)   %>

				<% if Len(oreceipt.FOneItem.Freg_num)=13 then %>
				    *******
				<% elseif Len(oreceipt.FOneItem.Freg_num)=10 then %>
				    ****
				<% elseif Len(oreceipt.FOneItem.Freg_num)=11 then %>
				    *****
				<% else %>

				<% end if %>
			<% end if %>
			</td>
		</tr>
		<% if (oreceipt.FOneItem.Freg_num="0100001234") then %>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">���ι�ȣ</td>
			<td align="left" bgcolor="#FFFFFF"><b><%=oreceipt.FOneItem.Fresultcashnoappl%></b></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">����(�ŷ�)����</td>
			<td align="left" bgcolor="#FFFFFF"><b><%= CHKIIF(IsNull(oreceipt.FOneItem.FEvalDT),"",Left(oreceipt.FOneItem.FEvalDT,10)) %></b></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">�ݾ�</td>
			<td align="left" bgcolor="#FFFFFF"><b><%= FormatNumber(oreceipt.FOneItem.Fcr_price,0) %></b></td>
		</tr>
		<% end if %>

		<tr>
			<td align="center"height="25"  bgcolor="<%= adminColor("tabletop") %>">����뵵</td>
			<td align="left" bgcolor="#FFFFFF"><%=oreceipt.FOneItem.getReceiptType%></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">�������</td>
			<td align="left" bgcolor="#FFFFFF">
				<% if oreceipt.FOneItem.Fresultcode="00" then %>
				    <% if (oreceipt.FOneItem.Fcancelyn="Y") and (oreceipt.FResultcount=1) then %>
				    <font color="red"><a href="javascript:popReceipt('<%=oreceipt.FOneItem.Ftid%>');">���� �� ���</a></font>

				    <input type="button" class="button" value="���ݿ����� ����" onClick="popEvalCashRecipt('<%= orderserial %>','');">
				    <% else %>
					<font color=darkblue>����Ϸ�</font>
					&nbsp;&nbsp;
					<input type="button" class="button" value="����������" onClick="popReceipt('<%=oreceipt.FOneItem.Ftid%>');">
					&nbsp;&nbsp;
					<input type="button" class="button" value="�������" onClick="jsCancel();">

					<form name="frmWrite" method="post" action="/cscenter/taxSheet/receipt_process.asp" style="margin:0px;">
	                <input type="hidden" name="chkPrint" value="<%=oreceipt.FOneItem.Fidx%>">
	                <input type="hidden" name="Atype" value="C2">
	                <input type="hidden" name="pggubun" value="<%=ojumun.FOneItem.Fpggubun%>">
	                <input type="hidden" name="mayPrc" value="0">
	                </form>
	                <% end if %>
				<% else %>
					<font color=darkred>�̹���</font>
					&nbsp;&nbsp;
					<%= oreceipt.FOneItem.FIpkumdiv %>
					<% if (IsCreateNewPaperOK) then %>
						<input type="button" class="button" value="���ݿ���������" onClick="jsCashEval('<%=oreceipt.FOneItem.Ftid%>');">
					<% else %>
						<br>���ݿ����� ���� ���� ���°� �ƴմϴ�.
						<br>(������ �Ǵ� ��� �Ǵ� �ݾ� ����)
					<% end if %>
					<form name="frmWrite" method="post" action="/cscenter/taxSheet/receipt_process.asp" style="margin:0px;">
	                <input type="hidden" name="chkPrint" value="<%=oreceipt.FOneItem.Fidx%>">
	                <input type="hidden" name="Atype" value="R">
	                <input type="hidden" name="pggubun" value="<%=ojumun.FOneItem.Fpggubun%>">
	                </form>
				<% end if %>
			</td>
		</tr>
	</table>
	<% if (oreceipt.FOneItem.Fcr_price<>mayReqPrc) or (isNaverPay and oreceipt.FOneItem.Fcr_price<>NpayCashAmt) or (isNaverPay and oreceipt.FOneItem.Fsup_price<>NpaySuplyAmt) then %>
    <p>
    <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center">
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">���ֹ�������(A)</td>
			<td align="left" bgcolor="#FFFFFF"><%=FormatNumber(orginPrc,0)%></td>
		</tr>
		<tr align="center">
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">��ǰ������(B)</td>
			<td align="left" bgcolor="#FFFFFF"><%=FormatNumber(minusSubtotalprice,0)%></td>
		</tr>
		<tr align="center">
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">�����ʿ��(A+B)</td>
			<td align="left" bgcolor="#FFFFFF"><%=FormatNumber(mayReqPrc,0)%></td>
		</tr>
		<% if (isNaverPay) then %>
		<tr align="center">
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">���̹�������ȸ��</td>
			<td align="left" bgcolor="#FFFFFF"><%=FormatNumber(NpayCashAmt,0)%>
			<% if (oreceipt.FOneItem.Fcr_price<>NpayCashAmt) then %>
				<% if (NpayCashAmt<>NpayCashAmt_Only) then %>
					(<%=FormatNumber(oreceipt.FOneItem.Fcr_price,0)%> :: <%=FormatNumber(NpayCashAmt_Only,0)%> :: <%=FormatNumber(NpayCashAmt-NpayCashAmt_Only,0)%>)
				<% else %>
					(<%=FormatNumber(oreceipt.FOneItem.Fcr_price,0)%>)
				<% end if %>
			<% end if %>
			</td>
		</tr>
		<tr align="center">
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">���̹�������ȸ��<br>(���ް�)</td>
			<td align="left" bgcolor="#FFFFFF"><%=FormatNumber(NpaySuplyAmt,0)%>
			<% if (oreceipt.FOneItem.Fsup_price<>NpaySuplyAmt) then %>
				<% if (NpaySuplyAmt<>NpaySuplyAmt_Only) then %>
					(<%=FormatNumber(oreceipt.FOneItem.Fsup_price,0)%> :: <%=FormatNumber(NpaySuplyAmt_Only,0)%> :: <%=FormatNumber(NpaySuplyAmt-NpaySuplyAmt_Only,0)%>)
				<% else %>
					(<%=FormatNumber(oreceipt.FOneItem.Fsup_price,0)%>)
				<% end if %>
			<% end if %>
			</td>
		</tr>
	    <% end if %>

	    <% if oreceipt.FOneItem.Fresultcode="00" then %>
	    <tr align="center">
			<td align="center" bgcolor="#FFFFFF" colspan="2">
			<% if (Not isNaverPay and (mayReqPrc=0)) or (isNaverPay and (NpayCashAmt=0)) then %>
			<input type="button" class="button" value="������� �ʿ�" onClick="jsCancel();">
		    <% elseif (Not isNaverPay and (oreceipt.FOneItem.Fcr_price<>mayReqPrc)) or (isNaverPay and (oreceipt.FOneItem.Fcr_price<>NpayCashAmt)) then %>
			<input type="button" class="button" value="����� ����� ����� �ʿ�" onClick="jsReEvalNCancel('<%=CHKIIF(isNaverPay,NpayCashAmt,mayReqPrc)%>');">
		    <% end if %>
			</td>
		</tr>
	    <% else %> <!-- ���� ���� ���� -->
	    <tr align="center">
			<td align="center" bgcolor="#FFFFFF" colspan="2">
			<% if (isNaverPay and oreceipt.FOneItem.Fcr_price<>NpayCashAmt) then %>
			NPay ���� �ݾ� ���� ��
		    <% end if %>

		    <% if (Not isNaverPay and oreceipt.FOneItem.Fcr_price<>mayReqPrc) then %>
			���� �ݾ� ���� ��
		    <% end if %>
			</td>
		</tr>
		<% end if %>
    </table>
    <% end if %>
<% end if %>

<!--
<% if (IsCashReceipExists = False) and (IsTaxSheetExists = False) and (Not IsCreateNewPaperOK) then %>
<div align="center">
	<br><br>���ݿ����� �� ��꼭�� ������ �� �����ϴ�.(��������)  <%  rw IsCreateNewPaperOK %> <% rw taxidx %>
</div>
<% end if %>
-->

<% if (IsCashReceipExists = False) and (IsTaxSheetExists = False) and (Not IsDacomCyberAccountPay) then %>
<!-- ���� ��û������ ���°��(������ ��� ����) -->

    <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center">
        <td bgcolor="#FFFFFF" colspan="2">
	<% if (False)  then %>
		<!-- �ֹ��� �ǽð���ü ����������(2011�� ������ �����ֹ�) -->
		<input type="button" class="button" value="INICIS������" onclick="javascript:receiptinicis('<%= ojumun.FOneItem.Fpaygatetid %>')">

	<% else %>
	    �����û������ �����ϴ�.<br><br>
	    <% if (isOLDORDER) then %>
	    	���� �ֹ� ���� ������ ���� ���
			<% if (orderserial="17022880499") then %>
				<input type="button" class="button" value="���ݿ����� ����" onClick="popEvalCashRecipt('<%= orderserial %>','');">
			<% end if %>
	    	<!-- <input type="button" class="button" value="���ݰ�꼭 ����" onClick="popWriteCustomerTaxSheet('<%= orderserial %>');"> -->
	    <% else %>
        	<input type="button" class="button" value="���ݿ����� ����" onClick="popEvalCashRecipt('<%= orderserial %>','');">

			<% if ojumun.FOneItem.Fsitename="10x10_cs" then %>
				<% if C_ADMIN_AUTH or C_CSPowerUser then %>
					<input type="button" class="button" value="���ݿ����� ����(������ 10x10_cs)" onClick="popEvalCashRecipt('<%= orderserial %>','10x10_cs');">
				<% end if %>
			<% end if %>
			<% if ojumun.FResultCount>0 then %>
				<% if (ojumun.FOneItem.Fjumundiv = "3") then %>
					<br><br>���ݰ�꼭 ����Ұ�(����翡 ������û�ؾ� �մϴ�.)
				<% else %>
					<input type="button" class="button" value="���ݰ�꼭 ����" onClick="popWriteCustomerTaxSheet('<%= orderserial %>');">
				<% end if %>
			<% end if %>
			<% if (orderserial="13031577231") then %>
				<br><br>
				<input type="button" class="button" value="���ݿ����� ����(�ݾ�����)" onClick="popEvalCashReciptHand('<%= orderserial %>');">
			<% end if %>
        <% end if %>
	<% end if %>
        </td>
    </tr>
    </table>

<% end if %>

<% if (IsCashReceipExists = False) and (IsTaxSheetExists = False) and (IsDacomCyberAccountPay = True) then %>

	<!-- ������ -->
    <script language='javascript'>
        location.replace('http://pg.dacom.net/transfer/cashreceipt.jsp?orderid=<%= orderserial %>&mid=tenbyten01&servicetype=SC0040&seqno=001');
    </script>

<% end if %>

	<% if (IsTaxSheetExists = True) then %>

		<% if (taxSheetRequestType = "01") then %>

			<!-- ��꼭 ���� -->
			<script language='javascript'>
			location.replace('/cscenter/taxsheet/Tax_view.asp?taxIdx=<%=oTax.FOneItem.FtaxIdx%>&searchDiv=N&page=1&menupos=861');
			</script>

		<% end if %>

		<% if (taxSheetRequestType = "11") then %>

			<!-- ���ݰ�꼭(2014�� ����) -->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" height="25">
					<td width="80" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
					<!--
					<td width="100" bgcolor="<%= adminColor("tabletop") %>">���̵�</td>
					-->
					<td width="50" bgcolor="<%= adminColor("tabletop") %>">������</td>
					<td width="40" bgcolor="<%= adminColor("tabletop") %>">�ֹ�<br>����</td>
					<td width="40" bgcolor="<%= adminColor("tabletop") %>">���<br>����</td>
					<td width="70" bgcolor="<%= adminColor("tabletop") %>">���ݾ�</td>
					<!--
					<td width="60" bgcolor="<%= adminColor("tabletop") %>">������<br>�׷��ڵ�</td>
					-->
					<td bgcolor="<%= adminColor("tabletop") %>">������<br>��ü��</td>
					<td width="80" bgcolor="<%= adminColor("tabletop") %>">������<br>����ڹ�ȣ</td>
					<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
					<td width="40" bgcolor="<%= adminColor("tabletop") %>">����<br>����</td>
					<td width="120" bgcolor="<%= adminColor("tabletop") %>">��꼭<br>IDX</td>
					<td width="50" bgcolor="<%= adminColor("tabletop") %>">������</td>
					<td width="80" bgcolor="<%= adminColor("tabletop") %>">������<br>�����</td>
					<td bgcolor="<%= adminColor("tabletop") %>">���</td>
				</tr>
				<% for i = 0 to oCTaxRequest.FResultCount - 1 %>
				<%
				if CLng(oCTaxRequest.FTaxList(i).FtaxIdx) > 0 then
					IsTaxIdxExist = True
				else
					IsAllTaxReqExist = False
				end if
				%>
				<tr align="center" height="25">
					<td bgcolor="#FFFFFF"><%= oCTaxRequest.FTaxList(i).Forderserial %></td>
					<!--
					<td bgcolor="#FFFFFF"><%= oCTaxRequest.FTaxList(i).Fuserid %></td>
					-->
					<td bgcolor="#FFFFFF"><%= oCTaxRequest.FTaxList(i).Fbuyname %></td>
					<td bgcolor="#FFFFFF"><%= oCTaxRequest.FTaxList(i).FOrdItemCNT %></td>
					<td bgcolor="#FFFFFF"><%= oCTaxRequest.FTaxList(i).FCHuLItemCNT %></td>
					<td bgcolor="#FFFFFF" align="right"><%= FormatNumber(oCTaxRequest.FTaxList(i).FchulgoPriceSum, 0) %></td>
					<!--
					<td bgcolor="#FFFFFF"><%= oCTaxRequest.FTaxList(i).Fgroupid %></td>
					-->
					<td bgcolor="#FFFFFF" align="left"><%= oCTaxRequest.FTaxList(i).Fcompany_name %><!--<br><%= oCTaxRequest.FTaxList(i).Fgroupid %>--></td>
					<td bgcolor="#FFFFFF"><%= oCTaxRequest.FTaxList(i).FbusiNO %></td>
					<td bgcolor="#FFFFFF" align="left"><%= oCTaxRequest.FTaxList(i).Fgoodname %></td>
					<td bgcolor="#FFFFFF"><%= oCTaxRequest.FTaxList(i).Fvatinclude %></td>
					<td bgcolor="#FFFFFF">
						<% if (oCTaxRequest.FTaxList(i).FOrdItemCNT <> 0) and (oCTaxRequest.FTaxList(i).FOrdItemCNT = oCTaxRequest.FTaxList(i).FCHuLItemCNT) then %>
							<% if (CLng(oCTaxRequest.FTaxList(i).FtaxIdx) = -1) then %>
								<input type="button" class="button" value="�����û" onclick="fnRegUpcheTax('<%= oCTaxRequest.FTaxList(i).Fgroupid %>', '<%= oCTaxRequest.FTaxList(i).GetGoodNameStr %>', '<%= oCTaxRequest.FTaxList(i).FchulgoPriceSum %>', '<%= oCTaxRequest.FTaxList(i).Fvatinclude %>', '<%= oCTaxRequest.FTaxList(i).FbusiIdx %>');">
							<% else %>
								<input type="button" class="button" value="��ȸ(<%= oCTaxRequest.FTaxList(i).FtaxIdx %>)" onclick="fnViewUpcheTax(<%= oCTaxRequest.FTaxList(i).FtaxIdx %>);">
							<% end if %>
						<% end if %>
					</td>
					<td bgcolor="#FFFFFF"><font color="<%= oCTaxRequest.FTaxList(i).GetChulgoStateColor %>"><%= oCTaxRequest.FTaxList(i).GetChulgoState %></font></td>
					<td bgcolor="#FFFFFF">
						<% if (oCTaxRequest.FTaxList(i).FlastChulgoDate <> "2000-01-01") then %>
						<%= oCTaxRequest.FTaxList(i).FlastChulgoDate %>
						<% end if %>
					</td>
					<td bgcolor="#FFFFFF"></td>
				</tr>
				<% next %>
			</table>

			<form name="frm" method="POST" action="doTaxOrder.asp" style="margin:0px;">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="orderserial" value="<%= orderserial %>">
			</form>

		<% end if %>

	<% end if %>

<br>
<div align="center">
	<% if (taxSheetRequestType = "11") and (IsTaxSheetExists = True) then %>
	<input type="button" class="button" value="��û ���" onclick="fnDelIssueReq();" <% if IsTaxIdxExist then %>disabled<% end if %> >
	&nbsp;
	<input type="button" class="button" value="�Ϸ�ó��" onclick="fnFinishIssueReq();" <% if Not IsAllTaxReqExist then %>disabled<% end if %> >
	&nbsp;
	<% end if %>
	<input type="button" class="button" value="â�ݱ�" onclick="window.close();">
</div>
</body>
</html>

<%
set ojumun = Nothing
set oreceiptList = Nothing
set oreceipt = Nothing
set  oTax = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->
