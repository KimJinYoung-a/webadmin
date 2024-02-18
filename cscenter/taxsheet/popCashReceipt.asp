<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 고객센터 현금영수증,세금계산서 밸행
' History : 이상구 생성
'			2023.07.31 한용민 생성(10x10_cs 주문건도 현금영수증 발행가능하게 추가)
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

'// 현금영수증 기발행 내역 있는지 //최근 3년 내역만 있음..
IsCashReceipExists = False
IsCashReceipListExists = False

set oreceipt = new CCashReceipt
oreceipt.FRectIdx = idx
oreceipt.FRectorderSerial = orderSerial

if (idx<>"") then
    oreceipt.GetOneCashReceipt
else
    if (IsCancelOrder) then
    	''D플래그 제외 전체 쿼리
    else
        oreceipt.FRectCancelyn = "N"
        oreceipt.FRectExcFailData = "Y"			'// 성공한 내역만, 2021-04-27, skyer9
    end if

    oreceipt.GetReceiptByOrderSerial

    ''2015/08/10 추가 (현금영수증 과거내역)
    if (oreceipt.FResultcount<1) then
        if (ojumun.FResultCount>0) then
            if (ojumun.FOneItem.GetPaperType="R") then
                oreceipt.GetReceiptByOrderSerial_OLD
            end if
        end if
    end if
end if

IsCashReceipExists = oreceipt.FResultCount > 0

'' 기발행/취소/요청 현금영수증 내역 (해당주문건 전체 조회)
set oreceiptList = new CCashReceipt
oreceiptList.FRectorderSerial = orderSerial
oreceiptList.FPageSize = 20
if (IsCashReceipExists) then
    oreceiptList.FRectExceptIdx = oreceipt.FOneItem.Fidx        ''현재 보여지는것은 표시 안함.
end if
if (oreceiptList.FRectorderSerial<>"") then
    oreceiptList.GetReceiptLogList
end if

IsCashReceipListExists = (oreceiptList.Ftotalcount>0)

'// 세금계산서 기발행 내역 있는지
IsTaxSheetExists = False
taxSheetRequestType = ""			'// 01 : 2013년 까지의 출고내역, 11 : 2014년 이후 출고내역(db_order.dbo.tbl_taxSheet 테이블에 billdiv 참조)
MxDlvDate = ""

taxidx = 0

if (orderSerial <> "") then

	'// 취소안된 상품의 마지막 출고일 기준으로
	'// 텐바이텐매출 또는 업체별매출로 계산서 발급한다.
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
		'// 2014년 이후 : 업체별매출
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
		'// 2013년 까지 : 텐바이텐매출
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

'// 발행요청 계산서 있는지(2014)
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
IsOldRealTimePay = false			'실시간이체 : 과거내역은 주문마스터의 paygatetid 로, 2011년 4월 15일 이후는

IsOldOrder = false
if (orderSerial<>"") then
    IsCreateNewPaperOK = False

    '(무통장 or 실시간) + 결제완료 + 기발행분 없을 경우
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

    '(무통장+실시간 이외)+결제완료+보조결제있음+기발행분 없을 경우
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

    ''' 아카데미 DIY
    '무통장+결제완료+기발행분 없을 경우
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
        '실시간+결제완료+기발행분 없을 경우
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
    if (accountno="국민 470301-01-014754") _
            or (accountno="신한 100-016-523130") _
            or (accountno="우리 092-275495-13-001") _
            or (accountno="하나 146-910009-28804") _
            or (accountno="기업 277-028182-01-046") _
            or (accountno="농협 029-01-246118") then
        IsDacomCyberAccountPay = false
    end if
end if

'''과거주문내역
'IF orderserial="11072931537" or orderserial="11071271646" or orderserial="11070961322" or orderserial="11070859345" or orderserial="11041320639" then
'    IsCreateCashReceiptOK = true
'ENd IF
%>
<html>
<head>
<title>현금영수증</title>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script type="text/javascript">

function jsCancel(){
	var f = document.frmWrite;
	if (f.chkPrint.value){
		if (confirm("발행된 현금영수증을 취소하시겠습니까?"))
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
		if (confirm("기 발행된 현금영수증을 취소후 재발행 하시겠습니까?"))
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
		if (confirm("현금영수증 요청금액을 수정 하시겠습니까?"))
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
		if (confirm("현금영수증을 발행하시겠습니까?"))
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
	    alert("결제완료 이전 주문입니다.");
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
	alert("2013년 출고내역입니다.\n\n텐바이텐 매출 세금계산서를 발행합니다.");
    var popwin=window.open("/cscenter/taxsheet/tax_view.asp?orderserial=" + orderserial,"popWriteCustomerTaxSheet","width=850,height=600,scrollbars=yes,resizable=yes");
    popwin.focus();
			<%
		end if
	end if
end if
	%>
}

// 신용카드 매출전표 팝업_이니시스
function receiptinicis(tid){
	var receiptUrl = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?" +
		"noTid=" + tid + "&noMethod=1";
	var popwin = window.open(receiptUrl,"INIreceipt","width=415,height=600");
	popwin.focus();
}

function fnDelIssueReq() {
	if (confirm('세금계산서 발행요청을 취소 하시겠습니까?\n\n발행된 계산서가 이미 있는 경우 먼저 계산서를 삭제하세요.')) {
		document.frm.mode.value="delIssueReq";
		document.frm.submit();
	}
}

function fnFinishIssueReq() {
	if (confirm('완료처리 하시겠습니까?')) {
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
	<b>현금영수증/세금계산서 발행 상태</b>
</div>
<br>
<% if (IsCashReceipListExists) then %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="<%= adminColor("tabletop") %>">
        <td>*기발행 로그</td>
        <td>금액</td>
        <td>구분</td>
        <td>상태</td>
        <td>발행일</td>
        <td>보기</td>
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
        <input type="button" value="보기" onClick="popReceipt('<%=oreceiptList.FItemList(i).Ftid%>');">
        <% end if %>
        </td>
    </tr>
    <% next %>
</table>
<p>
<% end if %>

<% if (IsCashReceipExists = True) then %>
<%
'' 발행내역 검토 2016/08/12 -----------------------------------------------------------
dim orginPrc, minusSubtotalprice, mayReqPrc
if (ojumun.FOneItem.FAccountDiv="7") or (ojumun.FOneItem.FAccountDiv="20") then   ''2016/09/19
    orginPrc = ojumun.FOneItem.FsubtotalPrice
else
    orginPrc = ojumun.FOneItem.FsumpaymentEtc
end if

if (ojumun.FOneItem.FCancelyn<>"N") then orginPrc=0 '' 취소

minusSubtotalprice = GetReceiptMinusOrderSUM(orderserial)

mayReqPrc = orginPrc+minusSubtotalprice

dim isNaverPay, NPay_Result, NpayCashAmt, NpayCashAmt_Only, NpaySuplyAmt, NpaySuplyAmt_Only
isNaverPay = (ojumun.FOneItem.Fpggubun="NP")

if (isNaverPay) then
    Set NPay_Result = fnCallNaverPayCashAmt(ojumun.FOneItem.Fpaygatetid)
    if NPay_Result.code="Success" then
		NpayCashAmt_Only = CLng(NPay_Result.body.totalCashAmount)
        NpayCashAmt    = NpayCashAmt_Only + ojumun.FOneItem.FsumPaymentEtc	'// 총 대상금액
		NpaySuplyAmt_Only = CLng(NPay_Result.body.supplyCashAmount)
		NpaySuplyAmt   = NpaySuplyAmt_Only + CLng(ojumun.FOneItem.FsumPaymentEtc*10/11)	'// 총 공급가
		''i_sup_price   = CLng(NPay_Result.body.supplyCashAmount) + CLng(myorder.FMasterItem.FsumPaymentEtc*10/11)	'// 현금성 공급가
		''i_tax         = i_cr_price - i_sup_price													'// 현금성 과세액
    end if
    Set NPay_Result = Nothing
end if

''취소건은 발행 불가.
IsCreateNewPaperOK = IsCreateNewPaperOK AND (ojumun.FOneItem.FCancelyn="N")
''요청내역 발행시. 발행금액 체크

'if (isNaverPay) then
'    IsCreateNewPaperOK = IsCreateNewPaperOK AND (oreceipt.FOneItem.Fcr_price=NpayCashAmt)
'else
'    IsCreateNewPaperOK = IsCreateNewPaperOK AND (oreceipt.FOneItem.Fcr_price=mayReqPrc)
'end if
'' -------------------------------------------------------------------------------------
%>
	<!-- 현금영수증 -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center">
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">상품명</td>
			<td align="left" bgcolor="#FFFFFF"><%=oreceipt.FOneItem.Fgoodname%></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">요청금액</td>
			<td align="left" bgcolor="#FFFFFF"><%=CurrFormat(oreceipt.FOneItem.Fcr_price)%></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">구매자명</td>
			<td align="left" bgcolor="#FFFFFF"><%=oreceipt.FOneItem.Fbuyername%></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">구매자메일</td>
			<td align="left" bgcolor="#FFFFFF"><a href="mailto:<%=oreceipt.FOneItem.Fbuyeremail%>"><%=oreceipt.FOneItem.Fbuyeremail%></a></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">식별번호</td>
			<td align="left" bgcolor="#FFFFFF">
			<% if (oreceipt.FOneItem.Freg_num="0100001234") then %>
			    <%= oreceipt.FOneItem.Freg_num %> (자진발급)
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
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">승인번호</td>
			<td align="left" bgcolor="#FFFFFF"><b><%=oreceipt.FOneItem.Fresultcashnoappl%></b></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">발행(거래)일자</td>
			<td align="left" bgcolor="#FFFFFF"><b><%= CHKIIF(IsNull(oreceipt.FOneItem.FEvalDT),"",Left(oreceipt.FOneItem.FEvalDT,10)) %></b></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">금액</td>
			<td align="left" bgcolor="#FFFFFF"><b><%= FormatNumber(oreceipt.FOneItem.Fcr_price,0) %></b></td>
		</tr>
		<% end if %>

		<tr>
			<td align="center"height="25"  bgcolor="<%= adminColor("tabletop") %>">발행용도</td>
			<td align="left" bgcolor="#FFFFFF"><%=oreceipt.FOneItem.getReceiptType%></td>
		</tr>
		<tr>
			<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">발행상태</td>
			<td align="left" bgcolor="#FFFFFF">
				<% if oreceipt.FOneItem.Fresultcode="00" then %>
				    <% if (oreceipt.FOneItem.Fcancelyn="Y") and (oreceipt.FResultcount=1) then %>
				    <font color="red"><a href="javascript:popReceipt('<%=oreceipt.FOneItem.Ftid%>');">발행 후 취소</a></font>

				    <input type="button" class="button" value="현금영수증 발행" onClick="popEvalCashRecipt('<%= orderserial %>','');">
				    <% else %>
					<font color=darkblue>발행완료</font>
					&nbsp;&nbsp;
					<input type="button" class="button" value="영수증보기" onClick="popReceipt('<%=oreceipt.FOneItem.Ftid%>');">
					&nbsp;&nbsp;
					<input type="button" class="button" value="발행취소" onClick="jsCancel();">

					<form name="frmWrite" method="post" action="/cscenter/taxSheet/receipt_process.asp" style="margin:0px;">
	                <input type="hidden" name="chkPrint" value="<%=oreceipt.FOneItem.Fidx%>">
	                <input type="hidden" name="Atype" value="C2">
	                <input type="hidden" name="pggubun" value="<%=ojumun.FOneItem.Fpggubun%>">
	                <input type="hidden" name="mayPrc" value="0">
	                </form>
	                <% end if %>
				<% else %>
					<font color=darkred>미발행</font>
					&nbsp;&nbsp;
					<%= oreceipt.FOneItem.FIpkumdiv %>
					<% if (IsCreateNewPaperOK) then %>
						<input type="button" class="button" value="현금영수증발행" onClick="jsCashEval('<%=oreceipt.FOneItem.Ftid%>');">
					<% else %>
						<br>현금영수증 발행 가능 상태가 아닙니다.
						<br>(결제전 또는 취소 또는 금액 상이)
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
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">원주문결제액(A)</td>
			<td align="left" bgcolor="#FFFFFF"><%=FormatNumber(orginPrc,0)%></td>
		</tr>
		<tr align="center">
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">반품결제액(B)</td>
			<td align="left" bgcolor="#FFFFFF"><%=FormatNumber(minusSubtotalprice,0)%></td>
		</tr>
		<tr align="center">
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">발행필요액(A+B)</td>
			<td align="left" bgcolor="#FFFFFF"><%=FormatNumber(mayReqPrc,0)%></td>
		</tr>
		<% if (isNaverPay) then %>
		<tr align="center">
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">네이버페이조회액</td>
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
			<td width="100" height="25" bgcolor="<%= adminColor("tabletop") %>">네이버페이조회액<br>(공급가)</td>
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
			<input type="button" class="button" value="발행취소 필요" onClick="jsCancel();">
		    <% elseif (Not isNaverPay and (oreceipt.FOneItem.Fcr_price<>mayReqPrc)) or (isNaverPay and (oreceipt.FOneItem.Fcr_price<>NpayCashAmt)) then %>
			<input type="button" class="button" value="기발행 취소후 재발행 필요" onClick="jsReEvalNCancel('<%=CHKIIF(isNaverPay,NpayCashAmt,mayReqPrc)%>');">
		    <% end if %>
			</td>
		</tr>
	    <% else %> <!-- 아직 발행 이전 -->
	    <tr align="center">
			<td align="center" bgcolor="#FFFFFF" colspan="2">
			<% if (isNaverPay and oreceipt.FOneItem.Fcr_price<>NpayCashAmt) then %>
			NPay 발행 금액 수정 됨
		    <% end if %>

		    <% if (Not isNaverPay and oreceipt.FOneItem.Fcr_price<>mayReqPrc) then %>
			발행 금액 수정 됨
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
	<br><br>현금영수증 및 계산서를 발행할 수 없습니다.(기발행없음)  <%  rw IsCreateNewPaperOK %> <% rw taxidx %>
</div>
<% end if %>
-->

<% if (IsCashReceipExists = False) and (IsTaxSheetExists = False) and (Not IsDacomCyberAccountPay) then %>
<!-- 발행 요청내역이 없는경우(발행후 취소 포함) -->

    <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center">
        <td bgcolor="#FFFFFF" colspan="2">
	<% if (False)  then %>
		<!-- 주문시 실시간이체 영수증발행(2011년 리뉴얼 이전주문) -->
		<input type="button" class="button" value="INICIS영수증" onclick="javascript:receiptinicis('<%= ojumun.FOneItem.Fpaygatetid %>')">

	<% else %>
	    발행요청내역이 없습니다.<br><br>
	    <% if (isOLDORDER) then %>
	    	과거 주문 내역 관리자 문의 요망
			<% if (orderserial="17022880499") then %>
				<input type="button" class="button" value="현금영수증 발행" onClick="popEvalCashRecipt('<%= orderserial %>','');">
			<% end if %>
	    	<!-- <input type="button" class="button" value="세금계산서 발행" onClick="popWriteCustomerTaxSheet('<%= orderserial %>');"> -->
	    <% else %>
        	<input type="button" class="button" value="현금영수증 발행" onClick="popEvalCashRecipt('<%= orderserial %>','');">

			<% if ojumun.FOneItem.Fsitename="10x10_cs" then %>
				<% if C_ADMIN_AUTH or C_CSPowerUser then %>
					<input type="button" class="button" value="현금영수증 발행(관리자 10x10_cs)" onClick="popEvalCashRecipt('<%= orderserial %>','10x10_cs');">
				<% end if %>
			<% end if %>
			<% if ojumun.FResultCount>0 then %>
				<% if (ojumun.FOneItem.Fjumundiv = "3") then %>
					<br><br>세금계산서 발행불가(여행사에 직접요청해야 합니다.)
				<% else %>
					<input type="button" class="button" value="세금계산서 발행" onClick="popWriteCustomerTaxSheet('<%= orderserial %>');">
				<% end if %>
			<% end if %>
			<% if (orderserial="13031577231") then %>
				<br><br>
				<input type="button" class="button" value="현금영수증 발행(금액조정)" onClick="popEvalCashReciptHand('<%= orderserial %>');">
			<% end if %>
        <% end if %>
	<% end if %>
        </td>
    </tr>
    </table>

<% end if %>

<% if (IsCashReceipExists = False) and (IsTaxSheetExists = False) and (IsDacomCyberAccountPay = True) then %>

	<!-- 데이콤 -->
    <script language='javascript'>
        location.replace('http://pg.dacom.net/transfer/cashreceipt.jsp?orderid=<%= orderserial %>&mid=tenbyten01&servicetype=SC0040&seqno=001');
    </script>

<% end if %>

	<% if (IsTaxSheetExists = True) then %>

		<% if (taxSheetRequestType = "01") then %>

			<!-- 계산서 발행 -->
			<script language='javascript'>
			location.replace('/cscenter/taxsheet/Tax_view.asp?taxIdx=<%=oTax.FOneItem.FtaxIdx%>&searchDiv=N&page=1&menupos=861');
			</script>

		<% end if %>

		<% if (taxSheetRequestType = "11") then %>

			<!-- 세금계산서(2014년 이후) -->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" height="25">
					<td width="80" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
					<!--
					<td width="100" bgcolor="<%= adminColor("tabletop") %>">아이디</td>
					-->
					<td width="50" bgcolor="<%= adminColor("tabletop") %>">구매자</td>
					<td width="40" bgcolor="<%= adminColor("tabletop") %>">주문<br>수량</td>
					<td width="40" bgcolor="<%= adminColor("tabletop") %>">출고<br>수량</td>
					<td width="70" bgcolor="<%= adminColor("tabletop") %>">출고금액</td>
					<!--
					<td width="60" bgcolor="<%= adminColor("tabletop") %>">공급자<br>그룹코드</td>
					-->
					<td bgcolor="<%= adminColor("tabletop") %>">공급자<br>업체명</td>
					<td width="80" bgcolor="<%= adminColor("tabletop") %>">공급자<br>사업자번호</td>
					<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
					<td width="40" bgcolor="<%= adminColor("tabletop") %>">과세<br>구분</td>
					<td width="120" bgcolor="<%= adminColor("tabletop") %>">계산서<br>IDX</td>
					<td width="50" bgcolor="<%= adminColor("tabletop") %>">출고상태</td>
					<td width="80" bgcolor="<%= adminColor("tabletop") %>">마지막<br>출고일</td>
					<td bgcolor="<%= adminColor("tabletop") %>">비고</td>
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
								<input type="button" class="button" value="발행요청" onclick="fnRegUpcheTax('<%= oCTaxRequest.FTaxList(i).Fgroupid %>', '<%= oCTaxRequest.FTaxList(i).GetGoodNameStr %>', '<%= oCTaxRequest.FTaxList(i).FchulgoPriceSum %>', '<%= oCTaxRequest.FTaxList(i).Fvatinclude %>', '<%= oCTaxRequest.FTaxList(i).FbusiIdx %>');">
							<% else %>
								<input type="button" class="button" value="조회(<%= oCTaxRequest.FTaxList(i).FtaxIdx %>)" onclick="fnViewUpcheTax(<%= oCTaxRequest.FTaxList(i).FtaxIdx %>);">
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
	<input type="button" class="button" value="요청 취소" onclick="fnDelIssueReq();" <% if IsTaxIdxExist then %>disabled<% end if %> >
	&nbsp;
	<input type="button" class="button" value="완료처리" onclick="fnFinishIssueReq();" <% if Not IsAllTaxReqExist then %>disabled<% end if %> >
	&nbsp;
	<% end if %>
	<input type="button" class="button" value="창닫기" onclick="window.close();">
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
