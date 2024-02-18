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
'// 현금영수증 기발행 내역 있는지 //최근 3년 내역만 있음..
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
    	''D플래그 제외 전체 쿼리
    else
        oreceipt.FRectCancelyn = "N"
    end if

    oreceipt.GetReceiptByOrderSerial
    
    if (oreceipt.FResultcount>0) then
        idx = oreceipt.FOneItem.Fidx
    end if
    ''2015/08/10 추가 (현금영수증 과거내역)
'    if (oreceipt.FResultcount<1) then
'        if (ojumun.FResultCount>0) then
'            if (ojumun.FOneItem.GetPaperType="R") then
'                oreceipt.GetReceiptByOrderSerial_OLD
'            end if
'        end if
'    end if
end if

IsCashReceipExists = oreceipt.FResultCount > 0

'' 기발행/취소/요청 현금영수증 내역 (해당주문건 전체 조회)
Dim oreceiptList
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

'''==============================================================================
'''// 세금계산서 기발행 내역 있는지
''dim IsTaxSheetExists : IsTaxSheetExists = False
''dim taxSheetRequestType : taxSheetRequestType = ""			'// 01 : 2013년 까지의 출고내역, 11 : 2014년 이후 출고내역(db_order.dbo.tbl_taxSheet 테이블에 billdiv 참조)
''dim MxDlvDate : MxDlvDate = ""
''
''Dim taxidx : taxidx = 0
''
''if (orderSerial <> "") then
''
''	'// 취소안된 상품의 마지막 출고일 기준으로
''	'// 텐바이텐매출 또는 업체별매출로 계산서 발급한다.
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
''		'// 2014년 이후 : 업체별매출
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
''		'// 2013년 까지 : 텐바이텐매출
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
'''// 발행요청 계산서 있는지(2014)
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
Dim IsOldRealTimePay 			: IsOldRealTimePay = false			'실시간이체 : 과거내역은 주문마스터의 paygatetid 로, 2011년 4월 15일 이후는

Dim authcode, cashreceiptReq, accountno

Dim IsOldOrder : IsOldOrder = false
if (orderSerial<>"") then
    IsCreateNewPaperOK = False

    '(무통장 or 실시간) + 결제완료 + 기발행분 없을 경우
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

    '(무통장+실시간 이외)+결제완료+보조결제있음+기발행분 없을 경우
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
<title>현금영수증</title>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script>
function jsCancel(){
	var f = document.frmWrite;
	if (f.chkPrint.value){
		if (confirm("발행된 현금영수증을 취소하시겠습니까?"))
		{
		    frmWrite.Atype.value='CA';
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
		    frmWrite.Atype.value='RNCA';
			f.submit();
		}
	}
}

function jsReCalcuEvalPrc(){
    var f = document.frmWrite;
	if (f.chkPrint.value){
		if (confirm("현금영수증 요청금액을 수정 하시겠습니까?"))
		{
		    frmWrite.Atype.value='RecalcuA';
			f.submit();
		}
	}
}

function popReceipt(tid){
    //var receiptUrl = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?noTid="+tid+"&noMethod=1";
    //https://admin.kcp.co.kr/Modules/Service/Cash/Cash_Bill_Common_View.jsp?term_id=PGNWT0000&orderid=TESTSHOP_080101&bill_yn=N&authno=560098441  ''매뉴얼.
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
		if (confirm("현금영수증을 발행하시겠습니까?"))
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



<p>
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
orginPrc = ojumun.FOneItem.FsubtotalPrice
if (ojumun.FOneItem.FCancelyn<>"N") then orginPrc=0 '' 취소
    
minusSubtotalprice = GetReceiptMinusOrderSUM(orderserial)

mayReqPrc = orginPrc+minusSubtotalprice

dim isNaverPay, NPay_Result, NpayCashAmt
isNaverPay = (ojumun.FOneItem.Fpggubun="NP")

if (isNaverPay) then 
    Set NPay_Result = fnCallNaverPayCashAmt(ojumun.FOneItem.Fpaygatetid)
    if NPay_Result.code="Success" then
        NpayCashAmt    = CLng(NPay_Result.body.totalCashAmount) + ojumun.FOneItem.FsumPaymentEtc	'// 총 대상금액
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

				    <input type="button" class="button" value="현금영수증 발행" onClick="popEvalCashRecipt('<%= orderserial %>');">
				    <% else %>
					<font color=darkblue>발행완료</font>
					&nbsp;&nbsp;
					<input type="button" class="button" value="영수증보기" onClick="popReceipt('<%=oreceipt.FOneItem.Ftid%>');">
					&nbsp;&nbsp;
					<input type="button" class="button" value="발행취소" onClick="jsCancel();">

					<form name="frmWrite" method="post" action="/cscenterv2/taxSheet/receipt_Fnprocess.asp">
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
			<td align="left" bgcolor="#FFFFFF"><%=FormatNumber(NpayCashAmt,0)%></td>
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


<% if (IsCashReceipExists = False) and  (Not IsDacomCyberAccountPay) then %>
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
	    <% if (orderserial="14102905760") then %>
	    <input type="button" class="button" value="현금영수증 발행" onClick="popEvalCashRecipt('<%= orderserial %>');">
	    <% end if %>
	    
	    <% else %>
        <input type="button" class="button" value="현금영수증 발행" onClick="popEvalCashRecipt('<%= orderserial %>');">
    
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


<br>
<div align="center">
	
	<input type="button" class="button" value="창닫기" onclick="window.close();">
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
