<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 매입상품원가관리
' History : 2022.01.17 이상구 생성
'           2022.08.18 한용민 수정(세금계산서 내용 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/PurchasedProductCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim totBuyPrice, totSuplyPrice, totVatPrice, forceRed, lastYYYYMM, lastYYYYMM2, menupos, displayRowCount
dim reportNo, reportPrice, orderNo, orderPrice, ipgoNo, ipgoPrice, totalPrice, orgprice, oCPurchasedProductPay
dim lastReportNo, lastReportPrice, lastOrderNo, lastOrderPrice, lastIpgoNo, lastIpgoPrice, lastTotalPrice
dim idx, yyyymmArray, i, j, k, arrcodeList, oCPurchasedProduct, oCPurchasedProductItem, INSERT_NODE, oCPurchasedProductSheet
dim eappReportUpdateYN, parameterStartdate, parameterEnddate, totalPayRequestPrice
    idx = requestCheckVar(getNumeric(request("idx")),10)
    menupos = requestCheckVar(getNumeric(request("menupos")),10)

INSERT_NODE = True
eappReportUpdateYN="Y"
if (idx <> "") then
    if Not IsNumeric(idx) then
        idx = ""
    end if
end if

if (idx <> "") then
    INSERT_NODE = False
end if

'// 품의정보
set oCPurchasedProduct = new CPurchasedProduct
    oCPurchasedProduct.FRectIdx = CHKIIF(idx="", "-1", idx)
    oCPurchasedProduct.GetPurchasedProductMaster

'// 상품정보
set oCPurchasedProductItem = new CPurchasedProduct
    oCPurchasedProductItem.FRectIdx = CHKIIF(idx="", "-1", idx)
    oCPurchasedProductItem.FPageSize = 1500
    oCPurchasedProductItem.FRectExcDel = "Y"
    oCPurchasedProductItem.GetPurchasedProductItemList

'// 원가정보
set oCPurchasedProductSheet = new CPurchasedProduct
    oCPurchasedProductSheet.FRectMasterIdx = CHKIIF(idx="", "-1", idx)
    oCPurchasedProductSheet.FPageSize = 1500
    oCPurchasedProductSheet.FRectExcDel = "Y"
    oCPurchasedProductSheet.GetPurchasedProductSheetMasterList

if oCPurchasedProductSheet.FResultCount>0 then
    for i=0 to oCPurchasedProductSheet.FResultCount-1
        ' 원가정보에 정산월을 모두 저장함
        if instr(yyyymmArray,oCPurchasedProductSheet.FItemList(i).Fyyyymm)<1 then
            yyyymmArray = yyyymmArray & oCPurchasedProductSheet.FItemList(i).Fyyyymm & ","
        end if
    next
    if right(yyyymmArray,1)="," then yyyymmArray = left(yyyymmArray,len(yyyymmArray)-1)
    parameterStartdate = dateadd("yyyy",-1,dateserial(left(yyyymmArray,4),right(yyyymmArray,2),"01"))
    parameterEnddate = date()
else
    parameterStartdate = dateadd("yyyy",-5,date())
    parameterEnddate = date()
end if

'// 결제정보
set oCPurchasedProductPay = new CPurchasedProduct
    oCPurchasedProductPay.FRectIdx = idx
    oCPurchasedProductPay.FPageSize = 50
    oCPurchasedProductPay.GetPurchasedProductItemPayList

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function ModiMaster(frm) {
	var ret = confirm('저장 하시겠습니까?');

	if (ret) {
		frm.submit();
	}
}

function jsRemoveOrder(frm) {
    if (frm.ordercode.value == '') {
        alert('먼저 삭제할 주문서를 입력하세요.');
        frm.ordercode.focus();
        return;
    }

	var ret = confirm('저장 하시겠습니까?');

	if (ret) {
        frm.mode.value = 'rmordr';
		frm.submit();
	}
}

function jsDelMaster(frm) {

    var ret = confirm('정말로 삭제 하시겠습니까?');

	if (ret) {
        frm.mode.value = 'delmaster';
		frm.submit();
	}
}

function jsApplyBuyPrice(frm) {
    var ret = confirm('원가 주문서 반영합니다. 진행 하시겠습니까?');

	if (ret) {
        frm.mode.value = 'doapplycogs';
		frm.submit();
	}
}

function jsApplyIpgoBuyPrice(frm) {
    var ret = confirm('원가 입고내역 반영합니다. 진행 하시겠습니까?');

	if (ret) {
        frm.mode.value = 'doapplyipgocogs';
		frm.submit();
	}
}

function jsApplyIpgoToOrder(frm) {
    var ret = confirm('입고월/주량 주문서 반영합니다. 진행 하시겠습니까?');

	if (ret) {
        frm.mode.value = 'doapplyipgotoorder';
		frm.submit();
	}
}

function jsCancel() {
    history.back();
}

function jsAddSheet(frm, lastYYYYMM) {

    if (frm.idx.value == '') {
        alert('먼저 품의자료를 저장하세요.');
        return;
    }

    var popwin = window.open("PurchasedProductSheetModify.asp?ppMasterIdx=" + frm.idx.value + "&lastYYYYMM=" + lastYYYYMM,"jsAddSheet","width=1200 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

function jsViewSheet(idx) {
    var popwin = window.open("PurchasedProductSheetModify.asp?idx=" + idx,"jsViewSheet","width=1200 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

function jsWriteReport() {
    var frm = document.frmEapp;

    if (frm.codeList.value == '') {
        alert('주문서 등록 후 품의서 작성가능합니다.');
        return;
    }

	var winEapp = window.open("","popE","width=1400,height=768,scrollbars=yes,resizable=yes");
	document.frmEapp.target = "popE";
	document.frmEapp.submit();
	winEapp.focus();
}

function jsViewEapp(reportidx,reportstate){
	var winEapp = window.open("/admin/approval/eapp/modeapp.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	winEapp.focus();
}

function PopPurchasedTaxPrintReDirect(itax_no, groupcode){
	var popPurchasedwinsub = window.open("/admin/newstorage/red_Purchasedtaxprint.asp?tax_no=" + itax_no + "&groupcode="+groupcode ,"Purchasedtaxview","width=1200,height=768,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popPurchasedwinsub.focus();
}

function popOrderlist(baljucode){
	var popwinOrderlist = window.open('/admin/newstorage/orderlist.asp?baljucode='+baljucode+'&menupos=537&yyyy1=<%= year(parameterStartdate) %>&mm1=<%= month(parameterStartdate) %>&dd1=<%= day(parameterStartdate) %>&yyyy2=<%= year(parameterEnddate) %>&mm2=<%= month(parameterEnddate) %>&dd2=<%= day(parameterEnddate) %>','addregOrderlist','width=1400,height=768,scrollbars=yes,resizable=yes');
	popwinOrderlist.focus();
}

function eappReportChgProcess(){
    if ($('#frmMaster input[name="reportIdx"]').val()=="" || $('#frmMaster input[name="reportIdx"]').val()=="0"){
        alert('변경할 품의번호가 없습니다.');
        frmMaster.reportIdx.focus();
        return;
    }
    $('#frmupdate input[name="reportIdx"]').val($('#frmMaster input[name="reportIdx"]').val());
    $('#frmupdate input[name="productidx"]').val('<%= idx %>');
    $('#frmupdate input[name="mode"]').val('ReportIdxEdit');
	frmupdate.action="/admin/newstorage/PurchasedProduct_process.asp";

	var ret = confirm('품의번호를 변경 하시겠습니까?');
	if(ret){
		frmupdate.submit();
	}
}

function eappReportDelProcess(){
    if ($('#frmMaster input[name="reportIdx"]').val()=="" || $('#frmMaster input[name="reportIdx"]').val()=="0"){
        alert('삭제할 품의번호가 없습니다.');
        frmMaster.reportIdx.focus();
        return;
    }
    $('#frmupdate input[name="reportIdx"]').val($('#frmMaster input[name="reportIdx"]').val());
    $('#frmupdate input[name="productidx"]').val('<%= idx %>');
    $('#frmupdate input[name="mode"]').val('ReportIdxDel');
	frmupdate.action="/admin/newstorage/PurchasedProduct_process.asp";

	var ret = confirm('품의번호를 삭제 하시겠습니까?');
	if(ret){
		frmupdate.submit();
	}
}

</script>
<form name="frmupdate" id="frmupdate" method="post" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="productidx" value="">
<input type="hidden" name="reportIdx" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<form name="frmMaster" id="frmMaster" method="post" action="/admin/newstorage/PurchasedProduct_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%= CHKIIF(INSERT_NODE, "insmaster", "modimaster") %>">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<!-- 상단바 시작 -->
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td colspan="4">
        ※ <font color="red"><strong>품의자료 <%= CHKIIF(INSERT_NODE, "작성", "수정") %></strong></font>
    </td>
</tr>
<!-- 상단바 끝 -->
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">원가마스터IDX</td>
    <td>
        <%= idx %>
        <input type="hidden" name="idx" value="<%= idx %>">
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">적요</td>
    <td>
        <input type="text" class="text" name="title" value="<%= oCPurchasedProduct.FOneItem.ftitle %>" size="50" maxlength=128>
    </td>
</tr>

<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">주문서</td>
    <td>
        <%
        if oCPurchasedProduct.FOneItem.FcodeList<>"" and not(isnull(oCPurchasedProduct.FOneItem.FcodeList)) then
            arrcodeList = split(oCPurchasedProduct.FOneItem.FcodeList,",")
            if isarray(arrcodeList) then
                for i = 0 to ubound(arrcodeList)
        %>
            <a href="#" onclick="popOrderlist('<%= arrcodeList(i) %>'); return false;"><%= arrcodeList(i) %></a>
            <% if i<>ubound(arrcodeList) then %>,<% end if %>
        <%
                next
            end if
        end if
        %>
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">주문서 추가</td>
    <td>
        <input type="text" class="text" name="ordercode" value="" size="10" autocomplete="off">
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">품의번호</td>
    <td width="40%">
        <% if eappReportUpdateYN="Y" then %>
            <input type="text" class="text" name="reportIdx" value="<%= oCPurchasedProduct.FOneItem.FreportIdx %>" size="6" maxlength="10" autocomplete="off">
            <input type="button" value="품의번호변경" onClick="eappReportChgProcess();" class="button" >
            <% if oCPurchasedProduct.FOneItem.FreportIdx<>"" and not(isnull(oCPurchasedProduct.FOneItem.FreportIdx)) then %>
                <input type="button" value="품의번호삭제" onClick="eappReportDelProcess();" class="button" >
            <% end if %>
        <% else %>
            <%= oCPurchasedProduct.FOneItem.FreportIdx %>
            <input type="hidden" name="reportIdx" value="<%= oCPurchasedProduct.FOneItem.FreportIdx %>">
        <% end if %>
        <br>
        <% if Not INSERT_NODE and oCPurchasedProduct.FOneItem.FreportIdx = 0 then %>
        <input type="button" class="button" value="품의서 작성" onClick="jsWriteReport()">
        <% elseif Not INSERT_NODE and oCPurchasedProduct.FOneItem.FreportIdx <> 0 then %>
        <%
        select case oCPurchasedProduct.FOneItem.FreportState
            case "7":
                response.write "품의완료"
            case "5":
                response.write "품의반려"
            case else:
                response.write "품의 진행중"
        end select
        %>
        <input type="button" class="button" value="품의서 보기" onClick="jsViewEapp(<%= oCPurchasedProduct.FOneItem.FreportIdx %>, '<%= oCPurchasedProduct.FOneItem.FreportState %>')">
        <% end if %>
        <input type="button" class="button" value="품의서 보기(TEST)" onClick="javascript:jsViewEapp('78907','8');">
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">품의금액</td>
    <td>
        <% if oCPurchasedProduct.FOneItem.FrealReportPrice <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.FrealReportPrice, 0) %> 원
        <% end if %>
        (
        <% if oCPurchasedProduct.FOneItem.FreportPrice <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.FreportPrice, 0) %> 원
        <% end if %>
        )
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">주문수량</td>
    <td>
        <% if oCPurchasedProduct.FOneItem.ForderNo <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.ForderNo, 0) %> 개
        <% end if %>
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">주문금액</td>
    <td>
        <% if oCPurchasedProduct.FOneItem.ForderPrice <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.ForderPrice, 0) %> 원
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">입고수량</td>
    <td>
        <% if oCPurchasedProduct.FOneItem.FipgoNo <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.FipgoNo, 0) %> 개
        <% end if %>
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">입고금액</td>
    <td>
        <% if oCPurchasedProduct.FOneItem.FipgoPrice <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.FipgoPrice, 0) %> 원
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">등록자</td>
    <td colspan="3">
        <% if INSERT_NODE then %>
        <%= html2db(session("ssBctCname")) %>(<%= session("ssBctId") %>)
        <% else %>
        <%= oCPurchasedProduct.FOneItem.Fregusername %>(<%= oCPurchasedProduct.FOneItem.Freguserid %>)
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">등록일</td>
    <td colspan="3">
        <% if Not INSERT_NODE then %>
        <%= oCPurchasedProduct.FOneItem.Findt %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">수정일</td>
    <td colspan="3">
        <% if Not INSERT_NODE then %>
        <%= oCPurchasedProduct.FOneItem.Fupdt %>
        <% end if %>
    </td>
</tr>
<% if oCPurchasedProduct.FOneItem.Fdeldt <> "" then %>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">삭제일</td>
    <td colspan="3">
        <%= oCPurchasedProduct.FOneItem.Fdeldt %>
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" align="center">
    <td align="center" colspan="15">
        <input type="button" class="button" value=" 저장하기 " onclick="ModiMaster(frmMaster)">
        <input type="button" class="button" value=" 취소 " onclick="jsCancel();">

        <% if (idx <> "") then %>
            &nbsp;
            &nbsp;
            <input type="button" class="button" value=" 주문서 삭제 " onClick="jsRemoveOrder(frmMaster)">
            <input type="button" class="button" value=" 삭제하기 " onclick="jsDelMaster(frmMaster);">
        <% end if %>
    </td>
</tr>
</table>

<br />

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<!-- 상단바 시작 -->
<tr height="25" bgcolor="FFFFFF">
    <td colspan="17">
        <table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
            <tr>
                <td>
                    ※ <font color="red"><strong>상품정보</strong></font>
                    월별 입고내역과 일치해야 합니다.
                </td>
                <td align="right">
                    총건수:  <%= oCPurchasedProductItem.FResultCount %>
                </td>
            </tr>
        </table>
    </td>
</tr>
<!-- 상단바 끝 -->

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width="100">
        입고예정월<br />
        (주문서)
    </td>
    <td width="150">브랜드</td>
    <td width="120">상품코드</td>
    <td>상품명</td>
    <td>옵션명</td>
    <td width="80">품의수량</td>
    <td width="80">품의금액</td>
    <td width="80">소비자가</td>
    <td width="80">원가</td>
    <td width="80">주문수량<br />(저장값)</td>
    <td width="80">원가총액</td>
    <td width="80">주문수량<br />(실시간)</td>
    <td width="80">주문서금액<br />(실시간)</td>
    <td width="80">입고수량<br />(실시간)</td>
    <td width="80">입고금액<br />(실시간)</td>
    <td width="200">비고</td>
</tr>
<%
orgprice=0
reportNo = 0
reportPrice = 0
orderNo = 0
orderPrice = 0
ipgoNo = 0
ipgoPrice = 0
totalPrice = 0
''lastReportNo, lastReportPrice, lastOrderNo, lastOrderPrice, lastIpgoNo, lastIpgoPrice, lastTotalPrice
lastReportNo = 0
lastReportPrice = 0
lastOrderNo = 0
lastOrderPrice = 0
lastIpgoNo = 0
lastIpgoPrice = 0
lastTotalPrice = 0
lastYYYYMM = ""
%>
<%
displayRowCount=0
for i=0 to oCPurchasedProductItem.FResultCount-1

' 입고예정월이 원가정보에 정산월에 동일월이 없는 경우 노출하지 않음
'if instr(yyyymmArray,oCPurchasedProductItem.FItemList(i).Fyyyymm)>0 or yyyymmArray="" then
' 주문수량 0 은 노출하지 않음
if oCPurchasedProductItem.FItemList(i).ForderNo<>"0" then
    if (i <> 0) and (lastYYYYMM <> oCPurchasedProductItem.FItemList(i).Fyyyymm) and displayRowCount>0 then	'// 전달 합계 표시
%>
<tr bgcolor="#FFFFFF" height="25" align="left">
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="right"><%= FormatNumber(lastReportNo, 0) %></td>
    <td align="right"><%= FormatNumber(lastReportPrice, 0) %></td>
    <td align="right"><%= FormatNumber(orgprice, 0) %></td>
    <td></td>
    <td></td>
    <td align="right"><b><%= FormatNumber(lastTotalPrice, 0) %></b></td>
    <td align="right"><%= FormatNumber(lastOrderNo, 0) %></td>
    <td align="right">
        <%
        forceRed = (lastTotalPrice > lastOrderPrice) or (lastReportPrice < lastOrderPrice)
        %>
        <%= CHKIIF(forceRed, "<b><font color=red>", "") %>
        <%= FormatNumber(lastOrderPrice, 0) %>
        <%= CHKIIF(forceRed, "</font></b>", "") %>
    </td>
    <td align="right"><%= FormatNumber(lastIpgoNo, 0) %></td>
    <td align="right"><%= FormatNumber(lastIpgoPrice, 0) %></td>
    <td></td>
    <%
    lastReportNo = 0
    lastReportPrice = 0
    lastOrderNo = 0
    lastOrderPrice = 0
    lastIpgoNo = 0
    lastIpgoPrice = 0
    lastTotalPrice = 0
    %>
</tr>
<%
    end if

    lastYYYYMM = oCPurchasedProductItem.FItemList(i).Fyyyymm
    orgprice = orgprice + oCPurchasedProductItem.FItemList(i).Forgprice
    reportNo = reportNo + oCPurchasedProductItem.FItemList(i).FreportNo
    reportPrice = reportPrice + oCPurchasedProductItem.FItemList(i).FreportPrice
    orderNo = orderNo + oCPurchasedProductItem.FItemList(i).ForderNo
    orderPrice = orderPrice + oCPurchasedProductItem.FItemList(i).ForderPrice
    ipgoNo = ipgoNo + oCPurchasedProductItem.FItemList(i).FipgoNo
    ipgoPrice = ipgoPrice + oCPurchasedProductItem.FItemList(i).FipgoPrice
    totalPrice = totalPrice + oCPurchasedProductItem.FItemList(i).FtotalPrice

    lastReportNo = lastReportNo + oCPurchasedProductItem.FItemList(i).FreportNo
    lastReportPrice = lastReportPrice + oCPurchasedProductItem.FItemList(i).FreportPrice
    lastOrderNo = lastOrderNo + oCPurchasedProductItem.FItemList(i).ForderNo
    lastOrderPrice = lastOrderPrice + oCPurchasedProductItem.FItemList(i).ForderPrice
    lastIpgoNo = lastIpgoNo + oCPurchasedProductItem.FItemList(i).FipgoNo
    lastIpgoPrice = lastIpgoPrice + oCPurchasedProductItem.FItemList(i).FipgoPrice
    lastTotalPrice = lastTotalPrice + oCPurchasedProductItem.FItemList(i).FtotalPrice
    displayRowCount = displayRowCount + 1
%>
<tr bgcolor="#FFFFFF" height="25" align="left">
    <td align="center">
        <%= oCPurchasedProductItem.FItemList(i).Fyyyymm %>
    </td>
    <td align="center"><%= oCPurchasedProductItem.FItemList(i).Fmakerid %></td>
    <td align="center">
        <%= oCPurchasedProductItem.FItemList(i).FItemGubun %>-<%= BF_GetFormattedItemId(oCPurchasedProductItem.FItemList(i).FItemID) %>-<%= oCPurchasedProductItem.FItemList(i).Fitemoption %>
    </td>

    <td><%= oCPurchasedProductItem.FItemList(i).Fitemname %></td>
    <td><%= oCPurchasedProductItem.FItemList(i).Fitemoptionname %></td>
    <td align="right">
        <input type="text" class="text" name="reportNo" value="<%= oCPurchasedProductItem.FItemList(i).FreportNo %>" size="7">
        <input type="hidden" name="didx" value="<%= oCPurchasedProductItem.FItemList(i).Fidx %>">
    </td>
    <td align="right">
        <input type="text" class="text" name="reportPrice" value="<%= oCPurchasedProductItem.FItemList(i).FreportPrice %>" size="7">
    </td>
    <td align="right"><%= FormatNumber(oCPurchasedProductItem.FItemList(i).Forgprice, 0) %></td>
    <td align="right"><%= FormatNumber(oCPurchasedProductItem.FItemList(i).Fcogs, 0) %></td>
    <td align="right"><%= FormatNumber(oCPurchasedProductItem.FItemList(i).ForderNo, 0) %></td>
    <td align="right"><%= FormatNumber(oCPurchasedProductItem.FItemList(i).FtotalPrice, 0) %></td>
    <td align="right">
        <%= FormatNumber(oCPurchasedProductItem.FItemList(i).ForderNo, 0) %>
        <% if (oCPurchasedProductItem.FItemList(i).ForderNo <> oCPurchasedProductItem.FItemList(i).Fbaljuitemno) then %>
        <font color="red">(<%= FormatNumber(oCPurchasedProductItem.FItemList(i).Fbaljuitemno, 0) %>)</font>
        <% end if %>
    </td>
    <td align="right">
        <%
        forceRed = (oCPurchasedProductItem.FItemList(i).FtotalPrice > oCPurchasedProductItem.FItemList(i).ForderPrice) or (oCPurchasedProductItem.FItemList(i).FreportPrice < oCPurchasedProductItem.FItemList(i).ForderPrice)
        %>
        <%= CHKIIF(forceRed, "<b><font color=red>", "") %>
        <%= FormatNumber(oCPurchasedProductItem.FItemList(i).ForderPrice, 0) %>
        <%= CHKIIF(forceRed, "</font></b>", "") %>
        <% if (oCPurchasedProductItem.FItemList(i).ForderPrice <> oCPurchasedProductItem.FItemList(i).Fbaljubuycash) then %>
        <font color="red">(<%= FormatNumber(oCPurchasedProductItem.FItemList(i).ForderPrice, 0) %>)</font>
        <% end if %>
    </td>
    <td align="right">
        <%= FormatNumber(oCPurchasedProductItem.FItemList(i).FipgoNo, 0) %>
        <% if (oCPurchasedProductItem.FItemList(i).FipgoNo <> oCPurchasedProductItem.FItemList(i).Fitemno) then %>
        <font color="red">(<%= FormatNumber(oCPurchasedProductItem.FItemList(i).Fitemno, 0) %>)</font>
        <% end if %>
    </td>
    <td align="right">
        <%= FormatNumber(oCPurchasedProductItem.FItemList(i).FipgoPrice, 0) %>
        <% if (oCPurchasedProductItem.FItemList(i).FipgoPrice <> oCPurchasedProductItem.FItemList(i).FrealItemPrice) then %>
        <font color="red">(<%= FormatNumber(oCPurchasedProductItem.FItemList(i).FrealItemPrice, 0) %>)</font>
        <% end if %>
    </td>
    <td>
        <% if forceRed then %>
            <% 'if (oCPurchasedProductItem.FItemList(i).FtotalPrice > oCPurchasedProductItem.FItemList(i).ForderPrice) then %>
            <%
            ' 주문서 저장시 round 0 으로 처리 되어 있어서 동일하게 처리함.
            if (FormatNumber(oCPurchasedProductItem.FItemList(i).Fcogs,0) * oCPurchasedProductItem.FItemList(i).ForderNo) <> oCPurchasedProductItem.FItemList(i).ForderPrice then
            %>
                * 주문서금액 원가 반영 필요
                <br>원가(<%= (FormatNumber(oCPurchasedProductItem.FItemList(i).Fcogs,0) * oCPurchasedProductItem.FItemList(i).ForderNo) %>)
                <br>주문서(<%= oCPurchasedProductItem.FItemList(i).ForderPrice %>)
            <% elseif False and (oCPurchasedProductItem.FItemList(i).FreportPrice < oCPurchasedProductItem.FItemList(i).ForderPrice) then %>
                <!--* 품의금액 초과-->
            <% end if %>
        <% end if %>
    </td>
</tr>
<% 'end if %>
<% end if %>
<% next %>
<% if (i >= (oCPurchasedProductItem.FResultCount-1)) and (lastOrderNo > 0) then %>
<tr bgcolor="#FFFFFF" height="25" align="left">
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="right"><%= FormatNumber(lastReportNo, 0) %></td>
    <td align="right"><%= FormatNumber(lastReportPrice, 0) %></td>
    <td align="right"><%= FormatNumber(orgprice, 0) %></td>
    <td></td>
    <td></td>
    <td align="right"><b><%= FormatNumber(lastTotalPrice, 0) %></b></td>
    <td align="right"><%= FormatNumber(lastOrderNo, 0) %></td>
    <td align="right">
        <%
        forceRed = (lastTotalPrice > lastOrderPrice) or (lastReportPrice < lastOrderPrice)
        %>
        <%= CHKIIF(forceRed, "<b><font color=red>", "") %>
        <%= FormatNumber(lastOrderPrice, 0) %>
        <%= CHKIIF(forceRed, "</font></b>", "") %>
    </td>
    <td align="right"><%= FormatNumber(lastIpgoNo, 0) %></td>
    <td align="right"><%= FormatNumber(lastIpgoPrice, 0) %></td>
    <td></td>
<% end if %>
<% if displayRowCount>0 then %>
<tr bgcolor="#FFFFFF" height="25" align="left">
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="right"><%= FormatNumber(reportNo, 0) %></td>
    <td align="right"><%= FormatNumber(reportPrice, 0) %></td>
    <td align="right"><%= FormatNumber(orgprice, 0) %></td>
    <td></td>
    <td></td>
    <td align="right"><%= FormatNumber(totalPrice, 0) %></td>
    <td align="right"><%= FormatNumber(orderNo, 0) %></td>
    <td align="right">
        <%
        forceRed = (totalPrice > orderPrice) or (reportPrice < orderPrice)
        %>
        <%= CHKIIF(forceRed, "<b><font color=red>", "") %>
        <%= FormatNumber(orderPrice, 0) %>
        <%= CHKIIF(forceRed, "</font></b>", "") %>
    </td>
    <td align="right"><%= FormatNumber(ipgoNo, 0) %></td>
    <td align="right"><%= FormatNumber(ipgoPrice, 0) %></td>
    <td>
        <% if forceRed then %>
        <% if False and (totalPrice > orderPrice) then %>
        * 주문서금액 원가 반영 필요
        <% elseif (abs(reportPrice) < abs(orderPrice)) then %>
        <!--* 품의금액 초과-->
        <% end if %>
        <% end if %>
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" align="center">
    <td align="center" colspan="16">
        <input type="button" class="button" value=" 저장하기 " onclick="ModiMaster(frmMaster)">
        &nbsp;
        <input type="button" class="button" value=" 원가 주문서반영 " onclick="jsApplyBuyPrice(frmMaster)">
        &nbsp;
        <input type="button" class="button" value=" 원가 입고반영 " onclick="jsApplyIpgoBuyPrice(frmMaster)">
        &nbsp;
        <input type="button" class="button" value=" 입고월/수량 주문서반영 " onclick="jsApplyIpgoToOrder(frmMaster)">
    </td>
</tr>
</table>
</form>

<br />

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="17">
        <table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
            <tr>
                <td>
                    ※ <font color="red"><strong>원가정보</strong></font>
                    세금계산서와 일치해야 합니다.
                </td>
                <td align="right">
                    총건수:  <%= oCPurchasedProductSheet.FResultCount %>
                </td>
            </tr>
        </table>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width="60">원가상세<Br>IDX</td>
    <td width="50">정산월</td>
    <td width="50">그룹코드</td>
    <td width="150">사업자명</td>
    <td width="120">비용구분</td>
    <td width="70">원가총액<br>(매입가)</td>
    <td width="60">관련품의IDX</td>
    <td width="60">세금계산서<Br>상태</td>
    <td width="80">세금계산서<Br>등록일</td>
	<td width="70">발행일</td>
    <!--
    <td width="80">공급가</td>
    <td width="80">부가세</td>
    -->
    <td>비고</td>
</tr>
<%
totBuyPrice = 0
totSuplyPrice = 0
totVatPrice = 0

lastTotalPrice = 0
lastYYYYMM2 = ""
%>
<%
for i=0 to oCPurchasedProductSheet.FResultCount-1
    if (i <> 0) and (lastYYYYMM2 <> oCPurchasedProductSheet.FItemList(i).Fyyyymm) then	'// 전달 합계 표시
%>
<tr bgcolor="#FFFFFF" height="25" align="left">
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="Right"><b><%= FormatNumber(lastTotalPrice, 0) %></b></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <!--
    <td align="Right"></td>
    <td align="Right"></td>
    -->
    <td></td>
    <% lastTotalPrice = 0 %>
</tr>
<%
    end if
    lastYYYYMM2 = oCPurchasedProductSheet.FItemList(i).Fyyyymm

    totBuyPrice = totBuyPrice + oCPurchasedProductSheet.FItemList(i).FbuyPrice
    totSuplyPrice = totSuplyPrice + oCPurchasedProductSheet.FItemList(i).FsuplyPrice
    totVatPrice = totVatPrice + oCPurchasedProductSheet.FItemList(i).FvatPrice

    lastTotalPrice = lastTotalPrice + oCPurchasedProductSheet.FItemList(i).FbuyPrice
%>
<tr bgcolor="#FFFFFF" height="25" align="center">
    <td><a href="javascript:jsViewSheet(<%= oCPurchasedProductSheet.FItemList(i).Fidx %>)"><%= oCPurchasedProductSheet.FItemList(i).Fidx %></a></td>
    <td><a href="javascript:jsViewSheet(<%= oCPurchasedProductSheet.FItemList(i).Fidx %>)"><%= oCPurchasedProductSheet.FItemList(i).Fyyyymm %></a></td>
    <td><a href="javascript:jsViewSheet(<%= oCPurchasedProductSheet.FItemList(i).Fidx %>)"><%= oCPurchasedProductSheet.FItemList(i).FgroupCode %></a></td>
    <td><a href="javascript:jsViewSheet(<%= oCPurchasedProductSheet.FItemList(i).Fidx %>)"><%= oCPurchasedProductSheet.FItemList(i).Fcompany_name %></a></td>
    <td><%= oCPurchasedProductSheet.FItemList(i).FppGubunName %></td>
    <td align="Right"><%= FormatNumber(oCPurchasedProductSheet.FItemList(i).FbuyPrice, 0) %></td>
    <td><%= oCPurchasedProductSheet.FItemList(i).freportIdx %></td>
    <td><%= GetStateName(oCPurchasedProductSheet.FItemList(i).ffinishflag) %></td>
    <td>
        <% if oCPurchasedProductSheet.FItemList(i).ftaxinputdate<>"" and not(isnull(oCPurchasedProductSheet.FItemList(i).ftaxinputdate)) then %>
            <%= left(oCPurchasedProductSheet.FItemList(i).ftaxinputdate,10) %>
            <Br><%= mid(oCPurchasedProductSheet.FItemList(i).ftaxinputdate,11,20) %>
        <% end if %>
    </td>
    <td><%= oCPurchasedProductSheet.FItemList(i).Ftaxregdate %></td>
    <!--
    <td align="Right"><%= FormatNumber(oCPurchasedProductSheet.FItemList(i).FsuplyPrice, 0) %></td>
    <td align="Right"><%= FormatNumber(oCPurchasedProductSheet.FItemList(i).FvatPrice, 0) %></td>
    -->
	<td>
		<% if IsElecTaxExists(oCPurchasedProductSheet.FItemList(i).fTaxLinkidx,oCPurchasedProductSheet.FItemList(i).ffinishflag) then %>
			<a href="#" onclick="PopPurchasedTaxPrintReDirect('<%= oCPurchasedProductSheet.FItemList(i).Fneotaxno %>','<%= oCPurchasedProductSheet.FItemList(i).fgroupCode %>'); return false;" class="btn3 btnIntb">출력</a>
		<% else %>
			<%= oCPurchasedProductSheet.FItemList(i).Fbillsitecode %>
		<% end if %>
	</td>
</tr>
<% next %>
<% if (i >= (oCPurchasedProductSheet.FResultCount-1)) and (lastTotalPrice > 0) then %>
<tr bgcolor="#FFFFFF" height="25" align="left">
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="Right"><b><%= FormatNumber(lastTotalPrice, 0) %></b></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <!--
    <td align="Right"></td>
    <td align="Right"></td>
    -->
    <td></td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" height="25" align="left">
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="Right"><%= FormatNumber(totBuyPrice, 0) %></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <td align="center"></td>
    <!--
    <td align="Right"><%= FormatNumber(totSuplyPrice, 0) %></td>
    <td align="Right"><%= FormatNumber(totVatPrice, 0) %></td>
    -->
    <td></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td align="center" colspan="13">
        <input type="button" class="button" value=" 추가하기 " onclick="jsAddSheet(frmMaster, '<%= lastYYYYMM %>')">
    </td>
</tr>
</table>

<% if oCPurchasedProductPay.FResultCount>0 then %>
    <br />
    <table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
        <td colspan="11">
            <table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
                <tr>
                    <td>
                        ※ <font color="red"><strong>결제정보</strong></font>
                    </td>
                    <td align="right">
                        총건수:  <%= oCPurchasedProductPay.FResultCount %>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
        <td width="60">품의번호</td>
        <td width="70">품의금액</td>
        <td width="80">결제요청서IDX</td>
        <td width="100">결제요청일</td>
        <td width="100">결제일</td>
        <td width="100">결제요청금액(원)</td>
        <td width="70">결제방법</td>
        <td>자금용도</td>
        <td>거래처</td>
        <td width="70">상태</td>
        <td width="50">비고</td>
    </tr>
    <% if oCPurchasedProductPay.FResultCount>0 then %>
        <%
        totalPayRequestPrice=0
        for i=0 to oCPurchasedProductPay.FResultCount-1
        totalPayRequestPrice=totalPayRequestPrice+oCPurchasedProductPay.FItemList(i).fpayRequestPrice
        %>
        <tr bgcolor="#FFFFFF" align="center">
            <td align="center"><%= oCPurchasedProductPay.FItemList(i).freportIdx %></td>
            <td align="right"><%= FormatNumber(oCPurchasedProductPay.FItemList(i).freportPrice, 0) %></td>
            <td align="center"><%= oCPurchasedProductPay.FItemList(i).fpayRequestidx %></td>
            <td align="center"><%= oCPurchasedProductPay.FItemList(i).fpayRequestdate %></td>
            <td align="center"><%= oCPurchasedProductPay.FItemList(i).fpaydate %></td>
            <td align="right"><%= FormatNumber(oCPurchasedProductPay.FItemList(i).fpayRequestPrice, 0) %></td>
            <td align="center"><%= fnGetPayType(oCPurchasedProductPay.FItemList(i).fpaytype) %></td>
            <td align="center"><%= oCPurchasedProductPay.FItemList(i).fpayRequestTitle %></td>
            <td align="center"><%= oCPurchasedProductPay.FItemList(i).fcust_nm %></td>
            <td align="center"><%= fnGetPayRequestState(oCPurchasedProductPay.FItemList(i).fpayrequeststate) %></td>
            <td align="center"></td>
        </tr>

        <% next %>
        <tr bgcolor="#FFFFFF">
            <td colspan="5" align="center">합계</td>
            <td align="right"><%= FormatNumber(totalPayRequestPrice, 0) %></td>
            <td colspan="5" align="center"></td>
        </tr>
    <% else %>
        <tr bgcolor="#FFFFFF">
            <td colspan="11" align="center" class="page_link">[검색결과가 없습니다.]</td>
        </tr>
    <% end if %>
    </table>
<% end if %>

<form name="frmEapp" method="post" action="PurchasedProduct_regeapp.asp" style="margin:0px;">
	<input type="hidden" name="idx" value="<%= idx %>">
    <input type="hidden" name="codeList" value="<%= oCPurchasedProduct.FOneItem.FcodeList %>">
</form>

<%
set oCPurchasedProduct=nothing
set oCPurchasedProductItem=nothing
set oCPurchasedProductSheet=nothing
set oCPurchasedProductPay=nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
