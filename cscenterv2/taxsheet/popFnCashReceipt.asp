<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_cashreceiptcls.asp"-->
<%

dim i
Dim idx				: idx	= req("idx","")
Dim orderSerial		: orderSerial	= req("orderSerial","")



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


Dim IsCancelOrder
If ojumun.FResultCount>0 then
    IsCancelOrder = ojumun.FOneItem.FCancelyn<>"N"
end if

''rw ojumun.FOneItem.GetPaperType
'==============================================================================
'// ���ݿ����� ����� ���� �ִ��� //�ֱ� 3�� ������ ����..
dim IsCashReceipExists : IsCashReceipExists = False
dim IsCashReceipListExists : IsCashReceipListExists = False
Dim oreceipt
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
    end if

    oreceipt.GetReceiptByOrderSerial
    
    if (oreceipt.FResultcount>0) then
        idx = oreceipt.FOneItem.Fidx
    end if
    ''2015/08/10 �߰� (���ݿ����� ���ų���)
'    if (oreceipt.FResultcount<1) then
'        if (ojumun.FResultCount>0) then
'            if (ojumun.FOneItem.GetPaperType="R") then
'                oreceipt.GetReceiptByOrderSerial_OLD
'            end if
'        end if
'    end if
end if

IsCashReceipExists = oreceipt.FResultCount > 0

'' �����/���/��û ���ݿ����� ���� (�ش��ֹ��� ��ü ��ȸ)
Dim oreceiptList
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

'''==============================================================================
'''// ���ݰ�꼭 ����� ���� �ִ���
''dim IsTaxSheetExists : IsTaxSheetExists = False
''dim taxSheetRequestType : taxSheetRequestType = ""			'// 01 : 2013�� ������ �����, 11 : 2014�� ���� �����(db_order.dbo.tbl_taxSheet ���̺� billdiv ����)
''dim MxDlvDate : MxDlvDate = ""
''
''Dim taxidx : taxidx = 0
''
''if (orderSerial <> "") then
''
''	'// ��Ҿȵ� ��ǰ�� ������ ����� ��������
''	'// �ٹ����ٸ��� �Ǵ� ��ü������� ��꼭 �߱��Ѵ�.
''	sqlStr = " select MAX(convert(Varchar(10),IsNull(d.beasongdate, getdate()),21)) as MxDlvDate "
''	sqlStr = sqlStr & " from "
''	sqlStr = sqlStr & " 	db_order.dbo.tbl_order_master m "
''	sqlStr = sqlStr & " 	join db_order.dbo.tbl_order_detail d "
''	sqlStr = sqlStr & " 	on "
''	sqlStr = sqlStr & " 		m.orderserial = d.orderserial "
''	sqlStr = sqlStr & " where "
''	sqlStr = sqlStr & " 	1 = 1 "
''	sqlStr = sqlStr & " 	and m.orderserial = '" + CStr(orderSerial) + "' "
''	sqlStr = sqlStr & " 	and d.itemid <> 0 "
''	sqlStr = sqlStr & " 	and d.cancelyn<>'Y' "
''
''    rsget.Open sqlStr, dbget, 1
''		if Not (rsget.EOF or rsget.BOF) then
''			MxDlvDate = rsget("MxDlvDate")
''		end if
''	rsget.Close
''
''	if (MxDlvDate >= "2014-01-01") or (MxDlvDate = "") then
''		'// 2014�� ���� : ��ü������
''		taxSheetRequestType = "11"
''
''		sqlStr = " select 1 idx "
''		sqlStr = sqlStr & " from "
''		sqlStr = sqlStr & " db_log.dbo.tbl_tax_issue_request "
''		sqlStr = sqlStr & " where orderserial = '" + CStr(orderSerial) + "' and useYN = 'Y' "
''
''		rsget.Open sqlStr, dbget, 1
''		if Not (rsget.EOF or rsget.BOF) then
''			IsTaxSheetExists = True
''		end if
''		rsget.Close
''	else
''		'// 2013�� ���� : �ٹ����ٸ���
''		taxSheetRequestType = "01"
''
''		sqlStr = "select taxIdx From db_order.[dbo].tbl_taxSheet"
''		sqlStr = sqlStr & " where orderserial = '" + CStr(orderSerial) + "' "
''		sqlStr = sqlStr & " and delYn='N'"
''
''		rsget.Open sqlStr, dbget, 1
''		if Not (rsget.EOF or rsget.BOF) then
''			IsTaxSheetExists = True
''			taxidx = CLng(rsget("taxIdx"))
''		end if
''		rsget.Close
''	end if
''end if


'==============================================================================
''Dim oTax
''set oTax = new CTax
''oTax.FRecttaxIdx = taxIdx
''if (CLng(taxidx)<>0) then
''	oTax.GetTaxRead
''
''	if oTax.FREsultCount>0 then
''	    taxIdx = CLng(oTax.FOneItem.FtaxIdx)
''	end if
''end if
''
'''// �����û ��꼭 �ִ���(2014)
''dim IsTaxIdxExist : IsTaxIdxExist = False
''dim IsAllTaxReqExist : IsAllTaxReqExist = True
''
''Dim oCTaxRequest
''set oCTaxRequest = new CTaxRequest
''oCTaxRequest.FRectOrderserial = orderSerial
''if (taxSheetRequestType = "11") and (IsTaxSheetExists = True) then
''	oCTaxRequest.FPageSize = 100
''	oCTaxRequest.FRectOrderserial = orderSerial
''	oCTaxRequest.GetTaxRequestOneOrder
''end if


'==============================================================================
dim sqlStr

Dim IsCreateNewPaperOK 			: IsCreateNewPaperOK = false
Dim IsDacomCyberAccountPay 		: IsDacomCyberAccountPay=false
Dim IsAcademy 					: IsAcademy = true
Dim IsOldRealTimePay 			: IsOldRealTimePay = false			'�ǽð���ü : ���ų����� �ֹ��������� paygatetid ��, 2011�� 4�� 15�� ���Ĵ�

Dim authcode, cashreceiptReq, accountno

Dim IsOldOrder : IsOldOrder = false
if (orderSerial<>"") then
    IsCreateNewPaperOK = False

    '(������ or �ǽð�) + �����Ϸ� + ������ ���� ���
    sqlStr = " select orderserial, IsNULL(authcode,'') as authcode, accountno, IsNULL(cashreceiptReq,'') as cashreceiptReq, ipkumdiv from db_academy.dbo.tbl_academy_order_master"
    sqlStr = sqlStr & " where orderserial='"&orderSerial&"'"
    sqlStr = sqlStr & " and ipkumdiv>=2"
    sqlStr = sqlStr & " and accountdiv in ('7','20') "

    rsget.Open sqlStr, dbget, 1
    If Not rsget.Eof then
        IsCreateNewPaperOK = rsget("ipkumdiv")>3
        authcode = rsget("authcode")
        cashreceiptReq  = rsget("cashreceiptReq")
        accountno = rsget("accountno")

    end if
    rsget.Close

    '(������+�ǽð� �̿�)+�����Ϸ�+������������+������ ���� ���
    sqlStr = " select orderserial, IsNULL(authcode,'') as authcode, accountno, IsNULL(cashreceiptReq,'') as cashreceiptReq, ipkumdiv, IsNull(sumPaymentEtc, 0) as sumPaymentEtc from db_academy.dbo.tbl_academy_order_master"
    sqlStr = sqlStr & " where orderserial='"&orderSerial&"'"
    sqlStr = sqlStr & " and ipkumdiv>=2"
    sqlStr = sqlStr & " and accountdiv not in ('7', '20') "

	''response.write sqlStr
    rsget.Open sqlStr, dbget, 1
    If Not rsget.Eof then
        IsCreateNewPaperOK = rsget("ipkumdiv")>3
        authcode = rsget("authcode")
        cashreceiptReq  = rsget("cashreceiptReq")
        accountno = rsget("accountno")
    end if
    rsget.Close

end if



%>
<html>
<head>
<title>���ݿ�����</title>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script>
function jsCancel(){
	var f = document.frmWrite;
	if (f.chkPrint.value){
		if (confirm("����� ���ݿ������� ����Ͻðڽ��ϱ�?"))
		{
		    frmWrite.Atype.value='CA';
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
		    frmWrite.Atype.value='RNCA';
			f.submit();
		}
	}
}

function jsReCalcuEvalPrc(){
    var f = document.frmWrite;
	if (f.chkPrint.value){
		if (confirm("���ݿ����� ��û�ݾ��� ���� �Ͻðڽ��ϱ�?"))
		{
		    frmWrite.Atype.value='RecalcuA';
			f.submit();
		}
	}
}

function popReceipt(tid){
    //var receiptUrl = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?noTid="+tid+"&noMethod=1";
    //https://admin.kcp.co.kr/Modules/Service/Cash/Cash_Bill_Common_View.jsp?term_id=PGNWT0000&orderid=TESTSHOP_080101&bill_yn=N&authno=560098441  ''�Ŵ���.
    //var receiptUrl = "https://admin.kcp.co.kr/Modules/Service/Cash/Cash_Bill_Common_View.jsp?term_id=T0000&orderid=&bill_yn=N&authno=";
    if (tid.length=14){
        var receiptUrl = "https://<%=chkIIF((application("Svr_Info")="Dev"),"dev","")%>admin.kcp.co.kr/Modules/Service/Cash/Cash_Bill_Common_View.jsp?cash_no="+tid;    //KCP
    }else{
	    var receiptUrl = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?noTid="+tid+"&noMethod=1";
	}
    
	var popwin = window.open(receiptUrl,"CashreceiptPrtFn","width=420,height=750,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function jsCashEval(iidx){
    var f = document.frmWrite;
	if (f.chkPrint.value){
		if (confirm("���ݿ������� �����Ͻðڽ��ϱ�?"))
		{
		    frmWrite.Atype.value='RA';
			f.submit();
		}
	}
}

function popEvalCashRecipt(orderserial){
    var popwin=window.open("/cscenter/receipt/INIreceiptReq.asp?orderserial=" + orderserial,"INIreceiptReq","width=680,height=480,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function popEvalCashReciptHand(orderserial){
    var popwin=window.open("/cscenter/receipt/INIreceiptReq.asp?issuetype=orderserial&orderserial=" + orderserial+"&hand=on","INIreceiptReq","width=680,height=480,scrollbars=yes,resizable=yes");
    popwin.focus();
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



<p>
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
orginPrc = ojumun.FOneItem.FsubtotalPrice
if (ojumun.FOneItem.FCancelyn<>"N") then orginPrc=0 '' ���
    
minusSubtotalprice = GetReceiptMinusOrderSUM(orderserial)

mayReqPrc = orginPrc+minusSubtotalprice

dim isNaverPay, NPay_Result, NpayCashAmt
isNaverPay = (ojumun.FOneItem.Fpggubun="NP")

if (isNaverPay) then 
    Set NPay_Result = fnCallNaverPayCashAmt(ojumun.FOneItem.Fpaygatetid)
    if NPay_Result.code="Success" then
        NpayCashAmt    = CLng(NPay_Result.body.totalCashAmount) + ojumun.FOneItem.FsumPaymentEtc	'// �� ���ݾ�
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
			    <% if LEN(oreceipt.FOneItem.Freg_num)=18 then %>
			        <%= oreceipt.FOneItem.Freg_num %>
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

				    <input type="button" class="button" value="���ݿ����� ����" onClick="popEvalCashRecipt('<%= orderserial %>');">
				    <% else %>
					<font color=darkblue>����Ϸ�</font>
					&nbsp;&nbsp;
					<input type="button" class="button" value="����������" onClick="popReceipt('<%=oreceipt.FOneItem.Ftid%>');">
					&nbsp;&nbsp;
					<input type="button" class="button" value="�������" onClick="jsCancel();">

					<form name="frmWrite" method="post" action="/cscenterv2/taxSheet/receipt_Fnprocess.asp">
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
					<form name="frmWrite" method="post" action="/cscenterv2/taxSheet/receipt_Fnprocess.asp">
	                <input type="hidden" name="chkPrint" value="<%=oreceipt.FOneItem.Fidx%>">
	                <input type="hidden" name="Atype" value="R">
	                <input type="hidden" name="pggubun" value="<%=ojumun.FOneItem.Fpggubun%>">
	                </form>
				<% end if %>
			</td>
		</tr>
	</table>
	<% if (oreceipt.FOneItem.Fcr_price<>mayReqPrc) or (isNaverPay and oreceipt.FOneItem.Fcr_price<>NpayCashAmt) then %>
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
			<td align="left" bgcolor="#FFFFFF"><%=FormatNumber(NpayCashAmt,0)%></td>
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


<% if (IsCashReceipExists = False) and  (Not IsDacomCyberAccountPay) then %>
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
	    <% if (orderserial="14102905760") then %>
	    <input type="button" class="button" value="���ݿ����� ����" onClick="popEvalCashRecipt('<%= orderserial %>');">
	    <% end if %>
	    
	    <% else %>
        <input type="button" class="button" value="���ݿ����� ����" onClick="popEvalCashRecipt('<%= orderserial %>');">
    
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


<br>
<div align="center">
	
	<input type="button" class="button" value="â�ݱ�" onclick="window.close();">
</div>
</body>
</html>

<%
set ojumun = Nothing
set oreceiptList = Nothing
set oreceipt = Nothing
%>
<!-- #include virtual="/cscenterv2/lib/poptail.asp"-->
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
