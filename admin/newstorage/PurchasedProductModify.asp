<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���Ի�ǰ��������
' History : 2022.01.17 �̻� ����
'           2022.08.18 �ѿ�� ����(���ݰ�꼭 ���� �߰�)
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

'// ǰ������
set oCPurchasedProduct = new CPurchasedProduct
    oCPurchasedProduct.FRectIdx = CHKIIF(idx="", "-1", idx)
    oCPurchasedProduct.GetPurchasedProductMaster

'// ��ǰ����
set oCPurchasedProductItem = new CPurchasedProduct
    oCPurchasedProductItem.FRectIdx = CHKIIF(idx="", "-1", idx)
    oCPurchasedProductItem.FPageSize = 1500
    oCPurchasedProductItem.FRectExcDel = "Y"
    oCPurchasedProductItem.GetPurchasedProductItemList

'// ��������
set oCPurchasedProductSheet = new CPurchasedProduct
    oCPurchasedProductSheet.FRectMasterIdx = CHKIIF(idx="", "-1", idx)
    oCPurchasedProductSheet.FPageSize = 1500
    oCPurchasedProductSheet.FRectExcDel = "Y"
    oCPurchasedProductSheet.GetPurchasedProductSheetMasterList

if oCPurchasedProductSheet.FResultCount>0 then
    for i=0 to oCPurchasedProductSheet.FResultCount-1
        ' ���������� ������� ��� ������
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

'// ��������
set oCPurchasedProductPay = new CPurchasedProduct
    oCPurchasedProductPay.FRectIdx = idx
    oCPurchasedProductPay.FPageSize = 50
    oCPurchasedProductPay.GetPurchasedProductItemPayList

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function ModiMaster(frm) {
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}

function jsRemoveOrder(frm) {
    if (frm.ordercode.value == '') {
        alert('���� ������ �ֹ����� �Է��ϼ���.');
        frm.ordercode.focus();
        return;
    }

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
        frm.mode.value = 'rmordr';
		frm.submit();
	}
}

function jsDelMaster(frm) {

    var ret = confirm('������ ���� �Ͻðڽ��ϱ�?');

	if (ret) {
        frm.mode.value = 'delmaster';
		frm.submit();
	}
}

function jsApplyBuyPrice(frm) {
    var ret = confirm('���� �ֹ��� �ݿ��մϴ�. ���� �Ͻðڽ��ϱ�?');

	if (ret) {
        frm.mode.value = 'doapplycogs';
		frm.submit();
	}
}

function jsApplyIpgoBuyPrice(frm) {
    var ret = confirm('���� �԰��� �ݿ��մϴ�. ���� �Ͻðڽ��ϱ�?');

	if (ret) {
        frm.mode.value = 'doapplyipgocogs';
		frm.submit();
	}
}

function jsApplyIpgoToOrder(frm) {
    var ret = confirm('�԰��/�ַ� �ֹ��� �ݿ��մϴ�. ���� �Ͻðڽ��ϱ�?');

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
        alert('���� ǰ���ڷḦ �����ϼ���.');
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
        alert('�ֹ��� ��� �� ǰ�Ǽ� �ۼ������մϴ�.');
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
        alert('������ ǰ�ǹ�ȣ�� �����ϴ�.');
        frmMaster.reportIdx.focus();
        return;
    }
    $('#frmupdate input[name="reportIdx"]').val($('#frmMaster input[name="reportIdx"]').val());
    $('#frmupdate input[name="productidx"]').val('<%= idx %>');
    $('#frmupdate input[name="mode"]').val('ReportIdxEdit');
	frmupdate.action="/admin/newstorage/PurchasedProduct_process.asp";

	var ret = confirm('ǰ�ǹ�ȣ�� ���� �Ͻðڽ��ϱ�?');
	if(ret){
		frmupdate.submit();
	}
}

function eappReportDelProcess(){
    if ($('#frmMaster input[name="reportIdx"]').val()=="" || $('#frmMaster input[name="reportIdx"]').val()=="0"){
        alert('������ ǰ�ǹ�ȣ�� �����ϴ�.');
        frmMaster.reportIdx.focus();
        return;
    }
    $('#frmupdate input[name="reportIdx"]').val($('#frmMaster input[name="reportIdx"]').val());
    $('#frmupdate input[name="productidx"]').val('<%= idx %>');
    $('#frmupdate input[name="mode"]').val('ReportIdxDel');
	frmupdate.action="/admin/newstorage/PurchasedProduct_process.asp";

	var ret = confirm('ǰ�ǹ�ȣ�� ���� �Ͻðڽ��ϱ�?');
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
<!-- ��ܹ� ���� -->
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td colspan="4">
        �� <font color="red"><strong>ǰ���ڷ� <%= CHKIIF(INSERT_NODE, "�ۼ�", "����") %></strong></font>
    </td>
</tr>
<!-- ��ܹ� �� -->
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">����������IDX</td>
    <td>
        <%= idx %>
        <input type="hidden" name="idx" value="<%= idx %>">
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">����</td>
    <td>
        <input type="text" class="text" name="title" value="<%= oCPurchasedProduct.FOneItem.ftitle %>" size="50" maxlength=128>
    </td>
</tr>

<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�ֹ���</td>
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
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�ֹ��� �߰�</td>
    <td>
        <input type="text" class="text" name="ordercode" value="" size="10" autocomplete="off">
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">ǰ�ǹ�ȣ</td>
    <td width="40%">
        <% if eappReportUpdateYN="Y" then %>
            <input type="text" class="text" name="reportIdx" value="<%= oCPurchasedProduct.FOneItem.FreportIdx %>" size="6" maxlength="10" autocomplete="off">
            <input type="button" value="ǰ�ǹ�ȣ����" onClick="eappReportChgProcess();" class="button" >
            <% if oCPurchasedProduct.FOneItem.FreportIdx<>"" and not(isnull(oCPurchasedProduct.FOneItem.FreportIdx)) then %>
                <input type="button" value="ǰ�ǹ�ȣ����" onClick="eappReportDelProcess();" class="button" >
            <% end if %>
        <% else %>
            <%= oCPurchasedProduct.FOneItem.FreportIdx %>
            <input type="hidden" name="reportIdx" value="<%= oCPurchasedProduct.FOneItem.FreportIdx %>">
        <% end if %>
        <br>
        <% if Not INSERT_NODE and oCPurchasedProduct.FOneItem.FreportIdx = 0 then %>
        <input type="button" class="button" value="ǰ�Ǽ� �ۼ�" onClick="jsWriteReport()">
        <% elseif Not INSERT_NODE and oCPurchasedProduct.FOneItem.FreportIdx <> 0 then %>
        <%
        select case oCPurchasedProduct.FOneItem.FreportState
            case "7":
                response.write "ǰ�ǿϷ�"
            case "5":
                response.write "ǰ�ǹݷ�"
            case else:
                response.write "ǰ�� ������"
        end select
        %>
        <input type="button" class="button" value="ǰ�Ǽ� ����" onClick="jsViewEapp(<%= oCPurchasedProduct.FOneItem.FreportIdx %>, '<%= oCPurchasedProduct.FOneItem.FreportState %>')">
        <% end if %>
        <input type="button" class="button" value="ǰ�Ǽ� ����(TEST)" onClick="javascript:jsViewEapp('78907','8');">
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">ǰ�Ǳݾ�</td>
    <td>
        <% if oCPurchasedProduct.FOneItem.FrealReportPrice <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.FrealReportPrice, 0) %> ��
        <% end if %>
        (
        <% if oCPurchasedProduct.FOneItem.FreportPrice <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.FreportPrice, 0) %> ��
        <% end if %>
        )
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�ֹ�����</td>
    <td>
        <% if oCPurchasedProduct.FOneItem.ForderNo <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.ForderNo, 0) %> ��
        <% end if %>
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�ֹ��ݾ�</td>
    <td>
        <% if oCPurchasedProduct.FOneItem.ForderPrice <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.ForderPrice, 0) %> ��
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�԰����</td>
    <td>
        <% if oCPurchasedProduct.FOneItem.FipgoNo <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.FipgoNo, 0) %> ��
        <% end if %>
    </td>
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�԰�ݾ�</td>
    <td>
        <% if oCPurchasedProduct.FOneItem.FipgoPrice <> "" then %>
        <%= FormatNumber(oCPurchasedProduct.FOneItem.FipgoPrice, 0) %> ��
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�����</td>
    <td colspan="3">
        <% if INSERT_NODE then %>
        <%= html2db(session("ssBctCname")) %>(<%= session("ssBctId") %>)
        <% else %>
        <%= oCPurchasedProduct.FOneItem.Fregusername %>(<%= oCPurchasedProduct.FOneItem.Freguserid %>)
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�����</td>
    <td colspan="3">
        <% if Not INSERT_NODE then %>
        <%= oCPurchasedProduct.FOneItem.Findt %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">������</td>
    <td colspan="3">
        <% if Not INSERT_NODE then %>
        <%= oCPurchasedProduct.FOneItem.Fupdt %>
        <% end if %>
    </td>
</tr>
<% if oCPurchasedProduct.FOneItem.Fdeldt <> "" then %>
<tr bgcolor="#FFFFFF" height="25">
    <td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">������</td>
    <td colspan="3">
        <%= oCPurchasedProduct.FOneItem.Fdeldt %>
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" align="center">
    <td align="center" colspan="15">
        <input type="button" class="button" value=" �����ϱ� " onclick="ModiMaster(frmMaster)">
        <input type="button" class="button" value=" ��� " onclick="jsCancel();">

        <% if (idx <> "") then %>
            &nbsp;
            &nbsp;
            <input type="button" class="button" value=" �ֹ��� ���� " onClick="jsRemoveOrder(frmMaster)">
            <input type="button" class="button" value=" �����ϱ� " onclick="jsDelMaster(frmMaster);">
        <% end if %>
    </td>
</tr>
</table>

<br />

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<!-- ��ܹ� ���� -->
<tr height="25" bgcolor="FFFFFF">
    <td colspan="17">
        <table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
            <tr>
                <td>
                    �� <font color="red"><strong>��ǰ����</strong></font>
                    ���� �԰����� ��ġ�ؾ� �մϴ�.
                </td>
                <td align="right">
                    �ѰǼ�:  <%= oCPurchasedProductItem.FResultCount %>
                </td>
            </tr>
        </table>
    </td>
</tr>
<!-- ��ܹ� �� -->

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width="100">
        �԰�����<br />
        (�ֹ���)
    </td>
    <td width="150">�귣��</td>
    <td width="120">��ǰ�ڵ�</td>
    <td>��ǰ��</td>
    <td>�ɼǸ�</td>
    <td width="80">ǰ�Ǽ���</td>
    <td width="80">ǰ�Ǳݾ�</td>
    <td width="80">�Һ��ڰ�</td>
    <td width="80">����</td>
    <td width="80">�ֹ�����<br />(���尪)</td>
    <td width="80">�����Ѿ�</td>
    <td width="80">�ֹ�����<br />(�ǽð�)</td>
    <td width="80">�ֹ����ݾ�<br />(�ǽð�)</td>
    <td width="80">�԰����<br />(�ǽð�)</td>
    <td width="80">�԰�ݾ�<br />(�ǽð�)</td>
    <td width="200">���</td>
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

' �԰������� ���������� ������� ���Ͽ��� ���� ��� �������� ����
'if instr(yyyymmArray,oCPurchasedProductItem.FItemList(i).Fyyyymm)>0 or yyyymmArray="" then
' �ֹ����� 0 �� �������� ����
if oCPurchasedProductItem.FItemList(i).ForderNo<>"0" then
    if (i <> 0) and (lastYYYYMM <> oCPurchasedProductItem.FItemList(i).Fyyyymm) and displayRowCount>0 then	'// ���� �հ� ǥ��
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
            ' �ֹ��� ����� round 0 ���� ó�� �Ǿ� �־ �����ϰ� ó����.
            if (FormatNumber(oCPurchasedProductItem.FItemList(i).Fcogs,0) * oCPurchasedProductItem.FItemList(i).ForderNo) <> oCPurchasedProductItem.FItemList(i).ForderPrice then
            %>
                * �ֹ����ݾ� ���� �ݿ� �ʿ�
                <br>����(<%= (FormatNumber(oCPurchasedProductItem.FItemList(i).Fcogs,0) * oCPurchasedProductItem.FItemList(i).ForderNo) %>)
                <br>�ֹ���(<%= oCPurchasedProductItem.FItemList(i).ForderPrice %>)
            <% elseif False and (oCPurchasedProductItem.FItemList(i).FreportPrice < oCPurchasedProductItem.FItemList(i).ForderPrice) then %>
                <!--* ǰ�Ǳݾ� �ʰ�-->
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
        * �ֹ����ݾ� ���� �ݿ� �ʿ�
        <% elseif (abs(reportPrice) < abs(orderPrice)) then %>
        <!--* ǰ�Ǳݾ� �ʰ�-->
        <% end if %>
        <% end if %>
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" align="center">
    <td align="center" colspan="16">
        <input type="button" class="button" value=" �����ϱ� " onclick="ModiMaster(frmMaster)">
        &nbsp;
        <input type="button" class="button" value=" ���� �ֹ����ݿ� " onclick="jsApplyBuyPrice(frmMaster)">
        &nbsp;
        <input type="button" class="button" value=" ���� �԰�ݿ� " onclick="jsApplyIpgoBuyPrice(frmMaster)">
        &nbsp;
        <input type="button" class="button" value=" �԰��/���� �ֹ����ݿ� " onclick="jsApplyIpgoToOrder(frmMaster)">
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
                    �� <font color="red"><strong>��������</strong></font>
                    ���ݰ�꼭�� ��ġ�ؾ� �մϴ�.
                </td>
                <td align="right">
                    �ѰǼ�:  <%= oCPurchasedProductSheet.FResultCount %>
                </td>
            </tr>
        </table>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width="60">������<Br>IDX</td>
    <td width="50">�����</td>
    <td width="50">�׷��ڵ�</td>
    <td width="150">����ڸ�</td>
    <td width="120">��뱸��</td>
    <td width="70">�����Ѿ�<br>(���԰�)</td>
    <td width="60">����ǰ��IDX</td>
    <td width="60">���ݰ�꼭<Br>����</td>
    <td width="80">���ݰ�꼭<Br>�����</td>
	<td width="70">������</td>
    <!--
    <td width="80">���ް�</td>
    <td width="80">�ΰ���</td>
    -->
    <td>���</td>
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
    if (i <> 0) and (lastYYYYMM2 <> oCPurchasedProductSheet.FItemList(i).Fyyyymm) then	'// ���� �հ� ǥ��
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
			<a href="#" onclick="PopPurchasedTaxPrintReDirect('<%= oCPurchasedProductSheet.FItemList(i).Fneotaxno %>','<%= oCPurchasedProductSheet.FItemList(i).fgroupCode %>'); return false;" class="btn3 btnIntb">���</a>
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
        <input type="button" class="button" value=" �߰��ϱ� " onclick="jsAddSheet(frmMaster, '<%= lastYYYYMM %>')">
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
                        �� <font color="red"><strong>��������</strong></font>
                    </td>
                    <td align="right">
                        �ѰǼ�:  <%= oCPurchasedProductPay.FResultCount %>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
        <td width="60">ǰ�ǹ�ȣ</td>
        <td width="70">ǰ�Ǳݾ�</td>
        <td width="80">������û��IDX</td>
        <td width="100">������û��</td>
        <td width="100">������</td>
        <td width="100">������û�ݾ�(��)</td>
        <td width="70">�������</td>
        <td>�ڱݿ뵵</td>
        <td>�ŷ�ó</td>
        <td width="70">����</td>
        <td width="50">���</td>
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
            <td colspan="5" align="center">�հ�</td>
            <td align="right"><%= FormatNumber(totalPayRequestPrice, 0) %></td>
            <td colspan="5" align="center"></td>
        </tr>
    <% else %>
        <tr bgcolor="#FFFFFF">
            <td colspan="11" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
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
