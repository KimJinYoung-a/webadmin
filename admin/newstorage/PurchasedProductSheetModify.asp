<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���Ի�ǰ��������
' History : 2022.01.17 �̻� ����
'           2022.08.19 �ѿ�� ����(���ݰ�꼭 ���� �߰�, ȯ����� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/jungsan_function.asp"-->
<!-- #include virtual="/lib/classes/PurchasedProductCls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim INSERT_NODE, idx, ppMasterIdx, lastYYYYMM, i, j, k, oCPurchasedProduct, oCPurchasedProductDetail
dim menupos, totalSumDbaljuitemno, totalSumDbuycash, totalSumAnbunBuyPrice, diffyyyymm
    idx = requestCheckVar(getNumeric(request("idx")),10)    '// �δ��� ������ IDX
    ppMasterIdx = requestCheckVar(getNumeric(request("ppMasterIdx")),10)    '// ǰ���ڷ� ������ IDX
    menupos = requestCheckVar(getNumeric(request("menupos")),10)

INSERT_NODE = True
totalSumDbaljuitemno=0
totalSumDbuycash=0
totalSumAnbunBuyPrice=0
diffyyyymm=""

if (idx <> "") then
    if Not IsNumeric(idx) then
        idx = ""
    end if
end if
if (ppMasterIdx <> "") then
    if Not IsNumeric(ppMasterIdx) then
        ppMasterIdx = ""
    end if
end if

lastYYYYMM = request("lastYYYYMM")
if lastYYYYMM = "" then
    lastYYYYMM = Left(Now, 7)
end if

if (idx <> "") then
    INSERT_NODE = False
end if

'// �δ��� ������
set oCPurchasedProduct = new CPurchasedProduct
    oCPurchasedProduct.FRectIdx = CHKIIF(idx="", "-1", idx)
    oCPurchasedProduct.GetPurchasedProductSheetMaster
    if oCPurchasedProduct.FResultCount>0 then
        ppMasterIdx = oCPurchasedProduct.FOneItem.FppMasterIdx

        diffyyyymm=CHKIIF(oCPurchasedProduct.FOneItem.Fyyyymm="", lastYYYYMM, oCPurchasedProduct.FOneItem.Fyyyymm)
    end if


set oCPurchasedProductDetail = new CPurchasedProduct
    oCPurchasedProductDetail.FPageSize = 1500
    oCPurchasedProductDetail.FCurrPage = 1
    oCPurchasedProductDetail.FRectMasterIdx = CHKIIF(idx="", "-1", idx)
    oCPurchasedProductDetail.GetPurchasedProductSheetDetailList

if (oCPurchasedProduct.FOneItem.FcodeList = "") and INSERT_NODE = False then
	oCPurchasedProductDetail.Fyyyymm = oCPurchasedProduct.FOneItem.Fyyyymm
	oCPurchasedProductDetail.GetPurchasedProductSheetDetailListByMonth
end if

%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type='text/javascript'>

function ModiMaster(frm) {
    var suplyPrice, vatPrice;
    var dsuplyPrice, dvatPrice;
    var dsuplyPriceSum, dvatPriceSum;

    dsuplyPriceSum = 0;
    dvatPriceSum = 0;

    /*
    if (frm.orderCode.value == '') {
        alert('���� �ֹ����ڵ带 �Է��ϼ���.');
        frm.orderCode.focus();
        return;
    }
    */

    suplyPrice = frm.suplyPrice.value.replace(/,/gi, '');
    vatPrice = frm.vatPrice.value.replace(/,/gi, '');

    if ((validNumber(suplyPrice) != true) || (validNumber(vatPrice) != true)) {
        alert('�ݾ��� �߸� �ԷµǾ����ϴ�.');
        return;
    }

    <% if False and (oCPurchasedProduct.FOneItem.FanbunType = "G203") then '' �����Է� %>
    for (var i = 0; ; i++) {
        dsuplyPrice = document.getElementById("anbunSuplyPrice" + i);
        dvatPrice = document.getElementById("anbunVatPrice" + i);

        if (dsuplyPrice == undefined) { break; }

        dsuplyPrice = dsuplyPrice.value.replace(/,/gi, '');
        dvatPrice = dvatPrice.value.replace(/,/gi, '');

        if ((validNumber(dsuplyPrice) != true) || (validNumber(dvatPrice) != true)) {
            alert('�ݾ��� �߸� �ԷµǾ����ϴ�.[�Ⱥбݾ� ����]');
            return;
        }

        document.getElementById("anbunSuplyPrice" + i).value = dsuplyPrice;
        document.getElementById("anbunVatPrice" + i).value = dvatPrice;

        dsuplyPriceSum = dsuplyPriceSum + dsuplyPrice*1;
        dvatPriceSum = dvatPriceSum + dvatPrice*1;
    }

    if ((dsuplyPriceSum*1 != suplyPrice*1) || (dsuplyPriceSum*1 != suplyPrice*1)) {
        alert('�ݾ��� �߸� �ԷµǾ����ϴ�.[�հ�ݾ� ����ġ]');
        alert(dsuplyPriceSum);
        alert(suplyPrice);
        return;
    }
    <% end if %>

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
        frm.suplyPrice.value = suplyPrice;
        frm.vatPrice.value = vatPrice;
        frm.mode.value = '<%= CHKIIF(INSERT_NODE, "inssheetmaster", "modisheetmaster") %>'
		frm.submit();
	}
}

function jsDelMaster(frm) {

    var ret = confirm('������ ���� �Ͻðڽ��ϱ�?');

	if (ret) {
        frm.mode.value = 'delsheetmaster';
		frm.submit();
	}
}

function validNumber(e) {
    var pattern = /^[0-9-]+$/;
    return pattern.test(e);
}

function calcBuyPrice() {
    var frm = document.frmMaster;
    var suplyPrice, vatPrice;

    suplyPrice = frm.suplyPrice.value.replace(/,/gi, '');
    vatPrice = frm.vatPrice.value.replace(/,/gi, '');

    if ((validNumber(suplyPrice) != true) || (validNumber(vatPrice) != true)) {
        return false;
    }

    frm.buyPrice.value = (suplyPrice*1 + vatPrice*1).format();
}

function calcDetailBuyPrice(detailidx) {
    var frm = document.frmMaster;

    var anbunSuplyPrice, anbunVatPrice, anbunBuyPrice;
    var valAnbunSuplyPrice, valAnbunVatPrice;

    anbunSuplyPrice = document.getElementById("anbunSuplyPrice" + detailidx);
    anbunVatPrice = document.getElementById("anbunVatPrice" + detailidx);
    anbunBuyPrice = document.getElementById("anbunBuyPrice" + detailidx);

    valAnbunSuplyPrice = anbunSuplyPrice.value.replace(/,/gi, '');
    valAnbunVatPrice = anbunVatPrice.value.replace(/,/gi, '');

    if ((validNumber(valAnbunSuplyPrice) != true) || (validNumber(valAnbunVatPrice) != true)) {
        return false;
    }

    anbunBuyPrice.value = (valAnbunSuplyPrice*1 + valAnbunVatPrice*1);
}

// ���� Ÿ�Կ��� �� �� �ֵ��� format() �Լ� �߰�
Number.prototype.format = function(){
    if(this==0) return 0;

    var reg = /(^[+-]?\d+)(\d{3})/;
    var n = (this + '');

    while (reg.test(n)) n = n.replace(reg, '$1' + ',' + '$2');

    return n;
};

// ���ڿ� Ÿ�Կ��� �� �� �ֵ��� format() �Լ� �߰�
String.prototype.format = function(){
    var num = parseFloat(this);
    if( isNaN(num) ) return "0";

    return num.format();
};

Date.prototype.yyyymmdd = function() {
  var mm = this.getMonth() + 1; // getMonth() is zero-based
  var dd = this.getDate();

  return [this.getFullYear(),
          (mm>9 ? '' : '0') + mm,
          (dd>9 ? '' : '0') + dd
         ].join('-');
};

function jsSetYYYYMM(diff) {
    var frm = document.frmMaster;
    var yyyymm, yyyy, mm;

    yyyymm = frm.yyyymm.value;

    var dt;
    try {
        if (isNaN(left(yyyymm, 4)*1) || isNaN(right(yyyymm, 2)*1)) {
            dt = new Date();
        } else {
            dt = new Date(left(yyyymm, 4)*1, right(yyyymm, 2)*1 - 1, 1);
        }

    } catch (e) {
        dt = new Date();
    }

    dt.setMonth(dt.getMonth() + diff);
    frm.yyyymm.value = left(dt.yyyymmdd(), 7);
}

function left(s,c) {
    return s.substr(0,c);
}

function right(s,c) {
    return s.substr(-c);
}

function jsRemoveOrder(frm) {
    if (frm.orderCode.value == '') {
        alert('���� ������ �ֹ����� �Է��ϼ���.');
        frm.orderCode.focus();
        return;
    }

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
        frm.mode.value = 'rmsheetordr';
		frm.submit();
	}
}

function jsGetBuyPrice() {
    var frm = document.frmMaster;
    var buyPrice, suplyPrice, vatPrice;

    var totPrice = frm.totPrice.value*1;
    buyPrice = totPrice;
    vatPrice = Math.round(1.0 * buyPrice / 11);
    //vatPrice = (1.0 * buyPrice / 11);
    suplyPrice = buyPrice - vatPrice;

    frm.buyPrice.value = buyPrice.format();
    frm.suplyPrice.value = suplyPrice.format();
    frm.vatPrice.value = vatPrice.format();
}

function jsCalcSuplyPrice() {
    var frm = document.frmMaster;
    var buyPrice, suplyPrice, vatPrice;

    buyPrice = frm.buyPrice.value.replace(/,/gi, '')*1;
    vatPrice = Math.round(1.0 * buyPrice / 11);
    //vatPrice = (1.0 * buyPrice / 11);
    suplyPrice = buyPrice - vatPrice;

    frm.suplyPrice.value = suplyPrice.format();
    frm.vatPrice.value = vatPrice.format();
}

<% '�̹����̻�� ��û���� �ּ� ó���� %>
<% 'function jsGetDetailBuyPrice(i) { %>
<% '    var dbuycash, anbunSuplyPrice, anbunVatPrice, anbunBuyPrice; %>
<% '    dbuycash = document.getElementById("dbuycash" + i).value*1; %>
<% '    anbunBuyPrice = dbuycash; %>
<% '    anbunVatPrice = Math.round(1.0 * anbunBuyPrice / 11); %>
<% '    anbunSuplyPrice = anbunBuyPrice - anbunVatPrice; %>
<% '    document.getElementById("anbunSuplyPrice" + i).value = anbunSuplyPrice; %>
<% '    document.getElementById("anbunVatPrice" + i).value = anbunVatPrice; %>
<% '    document.getElementById("anbunBuyPrice" + i).value = anbunBuyPrice; %>
<% '} %>
<% 'function jsGetDetailBuyPriceAll() { %>
<% '    for (var i = 0; ; i++) { %>
<% '        if (document.getElementById("dbuycash" + i) == undefined) { break; } %>
<% '        jsGetDetailBuyPrice(i); %>
<% '    } %>
<% '} %>

function finishflagoneChgProcess(sheetidx){
    var finishflagvar = '';

    for(var i=0; i<frmMaster.finishflag.length; i++){
        if (frmMaster.finishflag[i].checked){
            finishflagvar=frmMaster.finishflag[i].value;
        }
    }
    if (finishflagvar==''){
        alert('���õ� ���ݰ�꼭 ���°��� �����ϴ�.');
        return;
    }

    frmupdate.finishflag.value=finishflagvar;
    frmupdate.sheetidx.value=sheetidx;
	frmupdate.mode.value='finishflagone';
	frmupdate.action="/admin/newstorage/PurchasedProductJungsanProcess.asp";

	var ret = confirm('��꼭���� ���¸� ���� �Ͻðڽ��ϱ�?');
	if(ret){
		frmupdate.submit();
	}
}

function savetaxReg(frm){
    if (frm.taxregdate.value.length<1){
        alert('�������� �������� �ʾҽ��ϴ�. ');
        return;
    }

    if (frm.billsiteCode.value.length<1){
        alert('���� ��ü�� �������� �ʾҽ��ϴ�. ��� �Ͻðڽ��ϱ�?');
        return;
    }

	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret){
		frm.mode.value="taxregchange";
		frm.submit();
	}
}

function jsGetTax(ibizNo, itotSum){
	var sSearchText = ibizNo;
	var itotSum = itotSum;
	var winTax = window.open("/admin/tax/popSetEseroTax.asp?sST="+sSearchText+"&totSum="+itotSum+"&tgType=NRM","popGetTaxInfo","width=1400, height=768, resizable=yes, scrollbars=yes");
	winTax.focus();
}

function fillTaxInfo(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP){
    var frm = document.frmMaster;
    frm.taxregdate.value = dID;
    frm.eseroEvalSeq.value = eTax;

    //�����ü ����
    var mayApCd = eTax.substring(8,16);
    if (mayApCd=="10000000"){
        //����û
        frm.billsiteCode.value = 'E';
    }else if(mayApCd=="10000966"){
        //��365
        //frm.billsiteCode.value = 'B';
        // ���ϰ�
        frm.billsiteCode.value = 'WE';
    }else{
        //��Ÿ
        frm.billsiteCode.value = 'Y';
    }
}

function delTaxInfo(frm){
	var ret = confirm('��꼭 ���������� ���� �Ͻðڽ��ϱ�?');
	if (ret){
		frm.mode.value="delTaxInfo";
		frm.submit();
	}
}

function jsNewRegHand(){
    var winD = window.open("/admin/tax/popRegfileHand.asp","popDHand","width=860, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

function jsNewRegXML(){
    var winD = window.open("/admin/tax/popRegfileXML.asp","popDXML","width=600, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

// ���� ȯ�����
function CheckexchangeRateBuyPrice(i){
    var exchangeRate = 1;
    var orgBuyItemPrice = 0;
    var dbaljuitemno = 0;
    var buyItemPrice = 0;
    var currencyunit = '';
    var totalBuyPrice = 0;
    var vatPrice = 0;
    var suplyPrice = 0;

    exchangeRate = $('#frmMaster input[name="exchangerate"]').val();		// ȯ��
    if(exchangeRate!=''){
        if (!IsDouble(exchangeRate) || exchangeRate=='0'){
            alert('ȯ���� ���ڸ� �Է� ���� �մϴ�.');
            $('#frmMaster input[name="exchangerate"]').focus();
            return;
        }
    }else if(exchangeRate==''){
        alert('ȯ���� �Է��� �ּ���.');
        $('#frmMaster input[name="exchangerate"]').focus();
        return;
    }

    currencyunit = $('#currencyunit'+i).val();		// ��ȭȭ��
    if(currencyunit==''){
        alert('��ȭȭ�� �������� ���� ��ǰ�� �ֽ��ϴ�.\n��ǰ�ڵ� '+ $('#itemid'+i).val());
        $('#buyitemprice'+i).focus();
        return;
    }

    orgBuyItemPrice = $('#buyitemprice'+i).val();		// ���԰�
    if(orgBuyItemPrice=='' || orgBuyItemPrice=='0'){
        alert('FOB�� �������� ���� ��ǰ�� �ֽ��ϴ�.\n��ǰ�ڵ� '+ $('#itemid'+i).val());
        $('#buyitemprice'+i).focus();
        return;
    }

    dbaljuitemno = $('#dbaljuitemno'+i).val();		// �ֹ�����
    if(dbaljuitemno!=''){
        //if (!IsDouble(dbaljuitemno) || dbaljuitemno=='0'){
        //    alert('�ֹ������� ���ڸ� �Է� ���� �մϴ�.');
        //    $('#dbaljuitemno'+i).focus();
        //    return;
        //}
    }else if(dbaljuitemno==''){
        alert('�ֹ������� �Է��� �ּ���.');
        $('#dbaljuitemno'+i).focus();
        return;
    }

    if (currencyunit=='JPY'){
        buyItemPrice = Math.round( (orgBuyItemPrice*dbaljuitemno)*(exchangeRate/100) ).toFixed(0);
        //buyItemPrice = ( (orgBuyItemPrice*dbaljuitemno)*(exchangeRate/100) );
    }else{
        buyItemPrice = Math.round( (orgBuyItemPrice*dbaljuitemno)*exchangeRate ).toFixed(0);
        //buyItemPrice = ( (orgBuyItemPrice*dbaljuitemno)*exchangeRate );
    }

    vatPrice = Math.round(1.0 * buyItemPrice / 11).toFixed(0);
    //vatPrice = (1.0 * buyItemPrice / 11);
    suplyPrice = buyItemPrice - vatPrice;

    // ����Ʈ �󼼻�ǰ�� �� �Է�
    $('#anbunBuyPrice'+i).val(buyItemPrice);    // ���԰�
    $('#anbunVatPrice'+i).val(vatPrice);    // �ΰ���
    $('#anbunSuplyPrice'+i).val(suplyPrice);    // ���ް�

    // ���ջ�ݾ� ���
	var loopCount = '<%= oCPurchasedProductDetail.FResultCount %>';
	for (var j=0; j<loopCount; j++){
        totalBuyPrice = totalBuyPrice + parseInt($('#anbunBuyPrice'+j).val());
    }

    // ������ �� �Է�
    $('#buyPrice').val(totalBuyPrice);
    // ������ ���ް��� �ΰ��� ����
    jsCalcSuplyPrice()

    // ����Ʈ �� ���հ� �Է�
    $('#disptotalSumAnbunBuyPrice').html(plusComma(totalBuyPrice));
}

// ��� ȯ�����
function CheckexchangeRateBuyPriceAuto(){
    var exchangeRate = 1;
    var orgBuyItemPrice = 0;
    var dbaljuitemno = 0;
    var buyItemPrice = 0;
    var currencyunit = '';
    var totalBuyPrice = 0;
    var vatPrice = 0;
    var suplyPrice = 0;

    exchangeRate = $('#frmMaster input[name="exchangerate"]').val();		// ȯ��
    if(exchangeRate!=''){
        if (!IsDouble(exchangeRate) || exchangeRate=='0'){
            alert('ȯ���� ���ڸ� �Է� ���� �մϴ�.');
            $('#frmMaster input[name="exchangerate"]').focus();
            return;
        }
    }else if(exchangeRate==''){
        alert('ȯ���� �Է��� �ּ���.');
        $('#frmMaster input[name="exchangerate"]').focus();
        return;
    }

	var loopCount = '<%= oCPurchasedProductDetail.FResultCount %>';
	for (var i=0; i<loopCount; i++){
        currencyunit = $('#currencyunit'+i).val();		// ��ȭȭ��
        if(currencyunit==''){
            alert('��ȭȭ�� �������� ���� ��ǰ�� �ֽ��ϴ�.\n��ǰ�ڵ� '+ $('#itemid'+i).val());
            $('#buyitemprice'+i).focus();
            return;
        }

        orgBuyItemPrice = $('#buyitemprice'+i).val();		// ���԰�
        if(orgBuyItemPrice=='' || orgBuyItemPrice=='0'){
            alert('FOB�� �������� ���� ��ǰ�� �ֽ��ϴ�.\n��ǰ�ڵ� '+ $('#itemid'+i).val());
            $('#buyitemprice'+i).focus();
            return;
        }

        dbaljuitemno = $('#dbaljuitemno'+i).val();		// �ֹ�����
        if(dbaljuitemno!=''){
            //if (!IsDouble(dbaljuitemno) || dbaljuitemno=='0'){
            //    alert('�ֹ������� ���ڸ� �Է� ���� �մϴ�.');
            //    $('#dbaljuitemno'+i).focus();
            //    return;
            //}
        }else if(dbaljuitemno==''){
            alert('�ֹ������� �Է��� �ּ���.');
            $('#dbaljuitemno'+i).focus();
            return;
        }

        if (currencyunit=='JPY'){
            buyItemPrice = Math.round( (orgBuyItemPrice*dbaljuitemno)*(exchangeRate/100) ).toFixed(0);
            //buyItemPrice = ( (orgBuyItemPrice*dbaljuitemno)*(exchangeRate/100) );
        }else{
            buyItemPrice = Math.round( (orgBuyItemPrice*dbaljuitemno)*exchangeRate ).toFixed(0);
            //buyItemPrice = ( (orgBuyItemPrice*dbaljuitemno)*exchangeRate );
        }

        vatPrice = Math.round(1.0 * buyItemPrice / 11).toFixed(0);
        //vatPrice = (1.0 * buyItemPrice / 11);
        suplyPrice = buyItemPrice - vatPrice;

        // ����Ʈ �󼼻�ǰ�� �� �Է�
        $('#anbunBuyPrice'+i).val(buyItemPrice);    // ���԰�
        $('#anbunVatPrice'+i).val(vatPrice);    // �ΰ���
        $('#anbunSuplyPrice'+i).val(suplyPrice);    // ���ް�

        // ���ջ�ݾ� ���
        totalBuyPrice = totalBuyPrice + parseInt($('#anbunBuyPrice'+i).val());
	}

    // ������ �� �Է�
    $('#buyPrice').val(totalBuyPrice);
    // ������ ���ް��� �ΰ��� ����
    jsCalcSuplyPrice()

    // ����Ʈ �� ���հ� �Է�
    $('#disptotalSumAnbunBuyPrice').html(plusComma(totalBuyPrice));
}

// ���� ���ް��� �ΰ��� ���
function CheckSuplyPrice(i){
    var anbunBuyPrice = 0;
    var totalBuyPrice = 0;
    var vatPrice = 0;
    var suplyPrice = 0;

    anbunBuyPrice=$('#anbunBuyPrice'+i).val();    // ���԰�
    if(anbunBuyPrice!=''){
        //if (!IsDouble(anbunBuyPrice)){      //  || anbunBuyPrice=='0'
        //    alert('���԰��Ѿ��� ���ڸ� �Է� ���� �մϴ�.');
        //    //$('#anbunBuyPrice'+i).focus();
        //    return;
        //}
    }else if(anbunBuyPrice==''){
        alert('���԰��Ѿ��� �Է��� �ּ���.');
        //$('#anbunBuyPrice'+i).focus();
        return;
    }
    //vatPrice = Math.round(1.0 * anbunBuyPrice / 11).toFixed(0);
    vatPrice = (1.0 * anbunBuyPrice / 11);
    suplyPrice = anbunBuyPrice - vatPrice;

    // ����Ʈ �󼼻�ǰ�� �� �Է�
    $('#anbunVatPrice'+i).val(vatPrice);    // �ΰ���
    $('#anbunSuplyPrice'+i).val(suplyPrice);    // ���ް�

    // ���ջ�ݾ� ���
	var loopCount = '<%= oCPurchasedProductDetail.FResultCount %>';
	for (var j=0; j<loopCount; j++){
        totalBuyPrice = totalBuyPrice + parseInt($('#anbunBuyPrice'+j).val());
    }

    // ������ �� �Է�
    $('#buyPrice').val(totalBuyPrice);
    // ������ ���ް��� �ΰ��� ����
    jsCalcSuplyPrice()

    // ����Ʈ �� ���հ� �Է�
    $('#disptotalSumAnbunBuyPrice').html(plusComma(totalBuyPrice));
}

</script>

<form name="frmMaster" id="frmMaster" method="post" action="/admin/newstorage/PurchasedProduct_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td colspan="4">
        �� <font color="red"><strong>�δ��� <%= CHKIIF(INSERT_NODE, "�ۼ�", "����") %></strong></font>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="12%">������IDX</td>
    <td width="40%">
        <%= idx %>
        <input type="hidden" name="idx" value="<%= idx %>">
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="12%">����IDX</td>
    <td width="40%">
        <%= ppMasterIdx %>
        <input type="hidden" name="ppMasterIdx" value="<%= ppMasterIdx %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="12%">�԰���</td>
    <td width="40%">
        <input type="text" name="yyyymm" value="<%= CHKIIF(oCPurchasedProduct.FOneItem.Fyyyymm="", lastYYYYMM, oCPurchasedProduct.FOneItem.Fyyyymm) %>" size="7">
        <input type="button" class="button" value="������" onClick="jsSetYYYYMM(-1)">
        <input type="button" class="button" value="������" onClick="jsSetYYYYMM(1)">
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="12%">�����</td>
    <td width="40%">
        <%= oCPurchasedProduct.FOneItem.Fyyyymm %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">�׷��ڵ�</td>
    <td>
        <input type="text" name="groupCode" value="<%= oCPurchasedProduct.FOneItem.FgroupCode %>" size="10">
        <input type="button" class="button" value="��ü����" onClick="PopUpcheSelect('frmMaster', 'mode=cogs&pcuserdiv=902_21');">
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">����ڸ�</td>
    <td>
        <%= oCPurchasedProduct.FOneItem.Fcompany_name %>
    </td>
</tr>
<!--
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">�ֹ���</td>
    <td width="40%">
        <%= oCPurchasedProduct.FOneItem.FcodeList %>
        <% if (oCPurchasedProduct.FOneItem.FcodeList = "") then %>
        * �ֹ����� �߰����� ������ �԰������� ��� ��ǰ�� �߰��˴ϴ�.
        <% end if %>
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">�ֹ��� �߰�</td>
    <td>
        <input type="text" class="text" name="orderCode" value="" size="10">
    </td>
</tr>
-->
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">��뱸��</td>
    <td>
        <% Call drawCSCommCodeBox(True, "Z501", "ppGubun", oCPurchasedProduct.FOneItem.FppGubun, "") %>
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">�Ⱥй��</td>
    <td>
        <% Call drawCSCommCodeBox(True, "Z502", "anbunType", oCPurchasedProduct.FOneItem.FanbunType, "") %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">���԰�</td>
    <td>
        <input type="text" class="text" name="buyPrice" id="buyPrice" value="<%= FormatNumber(oCPurchasedProduct.FOneItem.FbuyPrice, 0) %>">
        <!--
        <input type="button" class="button" value="��������" onClick="jsGetBuyPrice()">
        -->
        <input type="button" class="button" value="�ڵ����" onClick="jsCalcSuplyPrice()">
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">���ް�</td>
    <td>
        <input type="text" class="text" name="suplyPrice" value="<%= FormatNumber(oCPurchasedProduct.FOneItem.FsuplyPrice, 0) %>" onFocusOut="calcBuyPrice()">
    </td>
</tr>

<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">�ΰ���</td>
    <td>
        <input type="text" class="text" name="vatPrice" value="<%= FormatNumber(oCPurchasedProduct.FOneItem.FvatPrice, 0) %>" onFocusOut="calcBuyPrice()">
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">����ǰ��IDX</td>
    <td>
        <%= oCPurchasedProduct.FOneItem.freportIdx %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">�ֹ����Ѿ�</td>
    <td>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.FtotPrice, 0) %>
        <input type="hidden" name="totPrice" value="<%= oCPurchasedProduct.FOneItem.FtotPrice %>">
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">�ֹ����Ѽ���</td>
    <td>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.FtotNo, 0) %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">���ʵ��</td>
    <td>
        <%= oCPurchasedProduct.FOneItem.Findt %>
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
    <td>
        <%= oCPurchasedProduct.FOneItem.Fupdt %>
    </td>
</tr>

<% if oCPurchasedProduct.FOneItem.Fdeldt <> "" then %>
    <tr bgcolor="#FFFFFF" height="25">
        <td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
        <td colspan="3">
            <%= oCPurchasedProduct.FOneItem.Fdeldt %>
        </td>
    </tr>
<% end if %>

</table>

<br />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td colspan="4">
        �� <font color="red"><strong>���ݰ�꼭����</strong></font>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="12%">����</td>
    <td width="40%">
        <input type="radio" name="finishflag" value="0" <% if oCPurchasedProduct.FOneItem.Ffinishflag="0" then response.write " checked" %>>�ۼ���
        <input type="radio" name="finishflag" value="1" <% if oCPurchasedProduct.FOneItem.Ffinishflag="1" then response.write " checked" %>>��꼭�����û
        <input type="radio" name="finishflag" value="3" <% if oCPurchasedProduct.FOneItem.Ffinishflag="3" then response.write " checked" %>>����Ϸ�

        <% if (idx <> "") then %>
            <input type="button" value="���º���" onClick="finishflagoneChgProcess('<%= idx %>');" class="button" >
        <% end if %>
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="12%">�����</td>
    <td width="40%">
        <% if oCPurchasedProduct.FOneItem.ftaxinputdate<>"" and not(isnull(oCPurchasedProduct.FOneItem.ftaxinputdate)) then %>
            <%= left(oCPurchasedProduct.FOneItem.ftaxinputdate,10) %>
            <Br><%= mid(oCPurchasedProduct.FOneItem.ftaxinputdate,11,20) %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
    <td colspan=3>
        <% if (oCPurchasedProduct.FOneItem.Ffinishflag="1") or (oCPurchasedProduct.FOneItem.Ffinishflag="3") or (oCPurchasedProduct.FOneItem.Ffinishflag="7") then %>
            <input type="text" name="taxregdate" value="<%= oCPurchasedProduct.FOneItem.Ftaxregdate %>" size="7" maxlength=10>
            <a href="#" onclick="calendarOpen(frmMaster.taxregdate); return false;"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
            <input type="button" value="��꼭��������" onclick="savetaxReg(frmMaster);" class="button">

      	    <% If ISNULL(oCPurchasedProduct.FOneItem.Ftaxlinkidx) then %>
                &nbsp;
                <input type="button" value="�����Է�" onClick="jsGetTax('<%= REplace(oCPurchasedProduct.FOneItem.Fcompany_no,"-","") %>','<%= oCPurchasedProduct.FOneItem.GetTotalSuplycash %>');" class="button">
                <input type="button" value="XML" onClick="jsNewRegXML();">
                <input type="button" value="���̰�꼭�Է�" onClick="jsNewRegHand();" class="button">
          	<% end if %>
            <br>
            <input type="hidden" name="taxlinkidx" value="<%= oCPurchasedProduct.FOneItem.Ftaxlinkidx %>">
            <% if isNULL(oCPurchasedProduct.FOneItem.Ftaxlinkidx) then %>
                <% call DrawBillSiteCombo("billsiteCode",oCPurchasedProduct.FOneItem.FbillsiteCode) %>
            <% else %>
                <input type="hidden" name="billsiteCode" value="<%= oCPurchasedProduct.FOneItem.FbillsiteCode %>">
                <%= oCPurchasedProduct.FOneItem.FbillSiteName %>
            <% end if %>
            <input type="text" name="neotaxno" value="<%= oCPurchasedProduct.FOneItem.Fneotaxno %>" size="20" maxlength="32" <%= CHKIIF(ISNULL(oCPurchasedProduct.FOneItem.Ftaxlinkidx),"","class='text_ro' READONLY") %>>(TAXNO)
            <br>
            <input type="text" name="eseroEvalSeq" value="<%= oCPurchasedProduct.FOneItem.FeseroEvalSeq %>" size="30" maxlength="24" <%= CHKIIF(ISNULL(oCPurchasedProduct.FOneItem.Ftaxlinkidx),"","class='text_ro' READONLY") %> >(�̼��� ���ι�ȣ '-' �����Է� 24�ڸ�)

            <% 'If ISNULL(oCPurchasedProduct.FOneItem.Ftaxlinkidx) then %>
                <% if (oCPurchasedProduct.FOneItem.Ffinishflag="0" or oCPurchasedProduct.FOneItem.Ffinishflag="1") then %>
                <br><input type="button" value="��꼭������������" onClick="delTaxInfo(frmMaster);" class="button">
                <% end if %>
            <% 'end if %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td align="center" colspan="4">
        <input type="button" class="button" value=" �����ϱ� " onclick="ModiMaster(frmMaster)">

        <% if (idx <> "") then %>
            <% '&nbsp;<input type="button" class="button" value=" �ֹ��� �����ϱ� " onClick="jsRemoveOrder(frmMaster)"> %>
            &nbsp;
            <input type="button" class="button" value=" �����ϱ� " onclick="jsDelMaster(frmMaster);">
        <% end if %>
    </td>
</tr>
</table>

<br />

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="17">
        <table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
            <tr>
                <td>
                    �� <font color="red"><strong>��ǰ����</strong></font>
                </td>
                <td align="right">
                    �ѰǼ�:  <%= oCPurchasedProductDetail.FResultCount %>
                </td>
            </tr>
        </table>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width="60">�������</td>
	<td width=100>�귣��ID</td>
    <td width="110">��ǰ�ڵ�</td>
    <td>��ǰ��</td>
    <td>�ɼǸ�</td>
    <td width="50">�ֹ�����</td>

    <%
    ' ��뱸���� ��ǰ��� �ϰ��
    if (oCPurchasedProduct.FOneItem.FppGubun = "G101") then
    %>
        <td width="50">��ȭȭ��</td>
        <td width="60">FOB</td>
    <% end if %>

    <td width="60">
        <!--�ֹ���-->
        ���԰�
    </td>
    <%
    ' �Ⱥй���� �����Է� �ϰ��
    if (false and oCPurchasedProduct.FOneItem.FanbunType = "G203") then
    %>
    <td width="90">���ް��Ѿ�</td>
    <td width="90">�ΰ����Ѿ�</td>
    <% end if %>
    <td width="80">
        <!--���԰�-->
        ���԰��Ѿ�
    </td>
    <td width="60">
        <%
        ' ��뱸���� ��ǰ��� �ϰ��
        if (oCPurchasedProduct.FOneItem.FppGubun = "G101") then
        %>
            <%
            ' �Ⱥй���� �����Է� �ϰ��
            if (oCPurchasedProduct.FOneItem.FanbunType = "G203") then
            %>
                ȯ��
                <input type="text" name="exchangerate" value="1000" maxlength=10 size=2 />
                <input type="button" class="button" value="�ϰ�����" onClick="CheckexchangeRateBuyPriceAuto()">
            <% end if %>
        <% end if %>

        <% '�̹����̻�� ��û���� �ּ� ó���� %>
        <% 'if (oCPurchasedProduct.FOneItem.FanbunType = "G203") then %>
            <% '<input type="button" class="button" value="�ϰ�ó��" onClick="jsGetDetailBuyPriceAll()"> %>
        <% 'end if %>
    </td>
    <td width="50">���</td>
</tr>
<% if oCPurchasedProductDetail.FResultCount>0 then %>
<%
for i=0 to oCPurchasedProductDetail.FResultCount-1
' ���꿹������ �δ��뿡 �԰����� ���Ͽ��� ���� ��� �������� ����
if diffyyyymm=oCPurchasedProductDetail.FItemList(i).ForderCode then
totalSumDbaljuitemno = totalSumDbaljuitemno + oCPurchasedProductDetail.FItemList(i).Fdbaljuitemno
totalSumDbuycash = totalSumDbuycash + oCPurchasedProductDetail.FItemList(i).Fdbuycash
'totalSumAnbunBuyPrice = totalSumAnbunBuyPrice + Round(oCPurchasedProductDetail.FItemList(i).FanbunBuyPrice, 0)
totalSumAnbunBuyPrice = totalSumAnbunBuyPrice + cdbl(oCPurchasedProductDetail.FItemList(i).FanbunBuyPrice)
%>
<tr bgcolor="#FFFFFF" height="25" align="left">
    <td align="center"><%= oCPurchasedProductDetail.FItemList(i).ForderCode %></td>
    <td align="center"><%= oCPurchasedProductDetail.FItemList(i).fmakerid %></td>
    <td align="center">
        <%= oCPurchasedProductDetail.FItemList(i).FItemGubun %>-<%= BF_GetFormattedItemId(oCPurchasedProductDetail.FItemList(i).FItemID) %>-<%= oCPurchasedProductDetail.FItemList(i).Fitemoption %>
        <input type="hidden" name="itemid" id="itemid<%= i %>" value="<%= oCPurchasedProductDetail.FItemList(i).FItemID %>">
    </td>

    <td><%= oCPurchasedProductDetail.FItemList(i).Fitemname %></td>
    <td><%= oCPurchasedProductDetail.FItemList(i).Fitemoptionname %></td>
    <td align="right">
        <%= FormatNumber(oCPurchasedProductDetail.FItemList(i).Fdbaljuitemno, 0) %>
        <input type="hidden" name="dbaljuitemno" id="dbaljuitemno<%= i %>" value="<%= oCPurchasedProductDetail.FItemList(i).Fdbaljuitemno %>">
    </td>

    <%
    ' ��뱸���� ��ǰ��� �ϰ��
    if (oCPurchasedProduct.FOneItem.FppGubun = "G101") then
    %>
        <td align="center">
            <%= oCPurchasedProductDetail.FItemList(i).fcurrencyunit %>
            <input type="hidden" name="currencyunit" id="currencyunit<%= i %>" value="<%= oCPurchasedProductDetail.FItemList(i).fcurrencyunit %>">
        </td>
        <td align="right">
            <%= oCPurchasedProductDetail.FItemList(i).fbuyitemprice %>
            <input type="hidden" name="buyitemprice" id="buyitemprice<%= i %>" value="<%= oCPurchasedProductDetail.FItemList(i).fbuyitemprice %>">
        </td>
    <% end if %>

    <td align="right">
        <% if cdbl(oCPurchasedProductDetail.FItemList(i).FanbunBuyPrice)<>0 and oCPurchasedProductDetail.FItemList(i).Fdbaljuitemno<>0 then %>
            <%= FormatNumber(cdbl(oCPurchasedProductDetail.FItemList(i).FanbunBuyPrice)/oCPurchasedProductDetail.FItemList(i).Fdbaljuitemno, 0) %>
        <% else %>
            0
        <% end if %>
        <%'= FormatNumber(oCPurchasedProductDetail.FItemList(i).Fdbuycash, 0) %>
        <input type="hidden" id="dbuycash<%= i %>" name="dbuycash" value="<%= oCPurchasedProductDetail.FItemList(i).Fdbuycash %>">
    </td>

    <%
    ' �Ⱥй���� �����Է� �ϰ��
    if (false and oCPurchasedProduct.FOneItem.FanbunType = "G203") then
    %>
        <td align="right">
            <input type="text" id="anbunSuplyPrice<%= i %>" name="anbunSuplyPrice" value="<%= cdbl(oCPurchasedProductDetail.FItemList(i).FanbunSuplyPrice) %>" class="text" size="8">
        </td>
        <td align="right">
            <% if (oCPurchasedProduct.FOneItem.FanbunType = "G203") then %>
                <input type="text" id="anbunVatPrice<%= i %>" name="anbunVatPrice" value="<%= cdbl(oCPurchasedProductDetail.FItemList(i).FanbunVatPrice) %>" class="text" size="8">
            <% else %>
                <input type="text" id="anbunVatPrice<%= i %>" name="anbunVatPrice" value="<%= cdbl(oCPurchasedProductDetail.FItemList(i).FanbunBuyPrice) - cdbl(oCPurchasedProductDetail.FItemList(i).FanbunSuplyPrice) %>" class="text" size="8">
            <% end if %>
        </td>
    <% else %>
        <input type="hidden" id="anbunSuplyPrice<%= i %>" name="anbunSuplyPrice" value="<%= cdbl(oCPurchasedProductDetail.FItemList(i).FanbunSuplyPrice) %>" class="text" size="8">
        <% if (oCPurchasedProduct.FOneItem.FanbunType = "G203") then %>
            <input type="hidden" id="anbunVatPrice<%= i %>" name="anbunVatPrice" value="<%= cdbl(oCPurchasedProductDetail.FItemList(i).FanbunVatPrice) %>" class="text" size="8">
        <% else %>
            <input type="hidden" id="anbunVatPrice<%= i %>" name="anbunVatPrice" value="<%= cdbl(oCPurchasedProductDetail.FItemList(i).FanbunBuyPrice) - cdbl(oCPurchasedProductDetail.FItemList(i).FanbunSuplyPrice) %>" class="text" size="8">
        <% end if %>
    <% end if %>

    <td align="right">
        <input type="hidden" name="detailidx" value="<%= oCPurchasedProductDetail.FItemList(i).Fidx %>">
        <%
        ' �Ⱥй���� �����Է� �ϰ��
        if (oCPurchasedProduct.FOneItem.FanbunType = "G203") then
        %>
            <input type="text" class="text" id="anbunBuyPrice<%= i %>" name="anbunBuyPrice" value="<%= cdbl(oCPurchasedProductDetail.FItemList(i).FanbunBuyPrice) %>" size="8" onFocusOut="CheckSuplyPrice('<%= i %>');">
        <% else %>
            <input type="text" class="text_ro" id="anbunBuyPrice<%= i %>" name="anbunBuyPrice" value="<%= cdbl(oCPurchasedProductDetail.FItemList(i).FanbunBuyPrice) %>" size="8" readOnly>
        <% end if %>
    </td>
    <td align="center">
        <%
        ' ��뱸���� ��ǰ��� �ϰ��
        if (oCPurchasedProduct.FOneItem.FppGubun = "G101") then
        %>
            <%
            ' �Ⱥй���� �����Է� �ϰ��
            if (oCPurchasedProduct.FOneItem.FanbunType = "G203") then
            %>
                <input type="button" class="button" value="ȯ�����" onClick="CheckexchangeRateBuyPrice('<%= i %>');">
            <% end if %>
        <% end if %>

        <% '�̹����̻�� ��û���� �ּ� ó���� %>
        <% 'if (oCPurchasedProduct.FOneItem.FanbunType = "G203") then %>
            <!--<input type="button" class="button" value="��������" onClick="jsGetDetailBuyPrice(<%'= i %>)">-->
        <% 'end if %>
    </td>
    <td>

    </td>
</tr>
<%
end if
next
%>
<tr bgcolor="#FFFFFF" height="25" align="left">
    <td align="center" colspan=5></td>
    <td align="right"><%= FormatNumber(totalSumDbaljuitemno, 0) %></td>

    <%
    ' ��뱸���� ��ǰ��� �ϰ��
    if (oCPurchasedProduct.FOneItem.FppGubun = "G101") then
    %>
        <td></td>
        <td></td>
    <% end if %>

    <td align="right">
        <%'= FormatNumber(totalSumDbuycash, 0) %>
    </td>

    <%
    ' �Ⱥй���� �����Է� �ϰ��
    if (false and oCPurchasedProduct.FOneItem.FanbunType = "G203") then
    %>
        <td></td>
        <td></td>
    <% end if %>

    <td align="right"><div id="disptotalSumAnbunBuyPrice"><%= FormatNumber(totalSumAnbunBuyPrice, 0) %><div></td>
    <td align="center" colspan=2></td>
</tr>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="17" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
</table>
</form>
<form action="" name="frmupdate" method="post" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sheetidx" value="">
<input type="hidden" name="finishflag" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<%
set oCPurchasedProduct=nothing
set oCPurchasedProductDetail=nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
