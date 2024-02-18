<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%
dim orderserial, oaslist, totalascount, ix
	orderserial = RequestCheckVar(trim(request("orderserial")),11)

totalascount = 0

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

if ojumun.FTotalCount < 1 then
	response.write "�ش�Ǵ� �ֹ����� �����ϴ�."
	dbget.close() : response.end
end if

dim ojumunitemsummary
set ojumunitemsummary = new COrderMaster
	ojumunitemsummary.FRectOldOrder = ojumun.FRectOldOrder
	ojumunitemsummary.FRectOrderSerial = orderserial
	ojumunitemsummary.getOrderItemSummary

set oaslist = new CCSASList
	if (orderserial <> "") then
	    oaslist.FRectOrderSerial = orderserial
	    oaslist.GetCSASTotalCount

	    totalascount = oaslist.FResultCount
	end if

	if (orderserial<>"") then
	    if ojumun.FOneItem.IsForeignDeliver then
	        ojumun.getEmsOrderInfo
	    end if
	end if

dim oetcpayment
set oetcpayment = new COrderMaster
	if (orderserial<>"") then
		oetcpayment.FRectOrderSerial = orderserial
		oetcpayment.FRectIncMainPayment = "Y"
		oetcpayment.getEtcPaymentList
	end if

Dim oUniPassNumber
If orderserial <> "" And Not isnull(orderserial) Then
	oUniPassNumber = fnUniPassNumber(orderserial)
end if

dim csorderserial
if (orderserial<>"") then
    csorderserial = GetCsOrderSerial(orderserial)
end if

%>
<link rel="stylesheet" href="/cscenter/css/cs.css" type="text/css">
<style>
body {
	overflow: auto;
}
</style>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/cscenter/ippbxmng/ippbxClick2Call.js"></script>
<script type="text/javascript">

function misendmaster(v){
	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + v,"misendmaster","width=1200 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cs_mileage(v){
	var popwin = window.open("/cscenter/mileage/cs_mileage.asp?userid=" + v,"cs_mileage","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cs_deposit(v){
	var popwin = window.open("/cscenter/deposit/cs_deposit.asp?userid=" + v,"cs_deposit","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cs_coupon(v){
	var popwin = window.open("/cscenter/coupon/cs_coupon.asp?userid=" + v,"cs_coupon","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function pop_cs_register(v){
	// var popwin = window.showModalDialog("/cscenter/action/pop_cs_register.asp?orderserial=" + v,"misendmaster","resizable:yes; scroll:yes; dialogWidth:825px; dialogHeight:800px ");
	var popwin = window.open("/cscenter/action/pop_cs_register.asp?orderserial=" + v,"misendmaster","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function jsPopBeasongDate(v) {
    var popwin = window.open("/cscenter/delivery/DeliveryTrackingSummaryOne.asp?orderserial=" + v,"jsPopBeasongDate","width=1400 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}
function popBeasongCompensation(v) {
    var popwin = window.open("/cscenter/delivery/deliverytcompensation.asp?orderserial=" + v,"popBeasongCompensation","width=1280 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function order_receiver_info(v){
	var popwin = window.showModalDialog("/cscenter/ordermaster/order_receiver_info.asp?orderserial=" + v,"order_reciever_info","resizable:no; scroll:no; dialogWidth:250px; dialogHeight:480px");
	popwin.focus();
}

function order_buyer_info(v){
	var popwin = window.showModalDialog("/cscenter/ordermaster/order_buyer_info.asp?orderserial=" + v,"order_buyer_info","resizable:no; scroll:no; dialogWidth:250px; dialogHeight:270px");
	popwin.focus();
}

// ============================================================================
// CS��ϰ���

// �ֹ����
function PopupCancelOrder(orderserial){
	var mode, divcd;

	mode = "";
	divcd = "A008";

	if (orderserial == "") {
	        alert("���� �ֹ��� �����ϼ���.");
	        return;
	}

	var popwin = window.open("/cscenter/action/pop_cs_action_new.asp?mode=" + mode + "&divcd=" + divcd + "&orderserial=" + orderserial,"PopupCancelOrder","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// �ֹ�����
function PopOpenModifyOrder(orderserial) {
	// var popwin = window.open("orderdetail_editorder.asp?orderserial=" + orderserial,"PopOpenModifyOrder","width=1400 height=800 scrollbars=yes resizable=yes");
	var popwin = window.open("orderdetail_simple_editorder.asp?orderserial=" + orderserial,"PopOpenModifyOrder","width=1400 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

// ��ǰ
function PopupReturnOrder(orderserial){
	var mode, divcd;

	mode = "";
	divcd = "A010";

	if (orderserial == "") {
	        alert("���� �ֹ��� �����ϼ���.");
	        return;
	}

	var popwin = window.open("/cscenter/action/pop_cs_action_new.asp?mode=" + mode + "&divcd=" + divcd + "&orderserial=" + orderserial,"PopupReturnOrder","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

//ī�����
function PopOpenCancelCard(orderserial){
	if (orderserial == "") {
	        alert("���� �ֹ��� �����ϼ���.");
	        return;
        }
	var popwin = window.open("/cscenter/action/pop_cs_write_repay.asp?divcd=7&orderserial=" + orderserial,"PopOpenCancelCard","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

//�ܺθ���ҿ�û
function PopOpenCancelOtherSite(orderserial){
	if (orderserial == "") {
	        alert("���� �ֹ��� �����ϼ���.");
	        return;
        }
	var popwin = window.open("/cscenter/action/pop_cs_write_repay.asp?divcd=5&orderserial=" + orderserial,"PopOpenCancelOtherSite","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

// ============================================================================
// �����丮 ����
var selected_history_menu = "";

function ChangeWriteButton(menuname) {
    selected_history_menu = menuname;

	<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
	    if (menuname == "memo") {
	        document.frmhistory.writebutton.value = "MEMO�ۼ�";
	    } else if (menuname == "cs") {
	        document.frmhistory.writebutton.value = "CS����Ʈ";
	    } else if (menuname == "mileage") {
	        document.frmhistory.writebutton.value = "���ϸ�������";
	    } else if (menuname == "deposit") {
	        document.frmhistory.writebutton.value = "��ġ�ݰ���";
	    } else if (menuname == "coupon") {
	        document.frmhistory.writebutton.value = "��������";
	    } else if (menuname == "qna") {
	        document.frmhistory.writebutton.value = "1:1������";
	    }
	<% end if %>
}

function OpenHistoryWindow(userid, orderserial) {
        if (selected_history_menu == "memo") {
                GotoHistoryMemoWrite(userid, orderserial);
        } else if (selected_history_menu == "cs") {
                Cscenter_Action_List(orderserial,'','')
        } else if (selected_history_menu == "mileage") {
                cs_mileage(userid)
        } else if (selected_history_menu == "deposit") {
                cs_deposit(userid)
        } else if (selected_history_menu == "coupon") {
                cs_coupon(userid)
        }
}

function GotoHistoryMemo(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("memo");

                document.history.location.href = "/cscenter/history/history_memo.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}

function GotoHistoryCS(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("cs");

                document.history.location.href = "/cscenter/history/history_cs.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}

function GotoHistoryMileage(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("mileage");

                document.history.location.href = "/cscenter/history/history_mileage.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}

function GotoHistoryDeposit(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("deposit");

                document.history.location.href = "/cscenter/history/history_deposit.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}

function GotoHistoryCoupon(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("coupon");

                document.history.location.href = "/cscenter/history/history_coupon.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}
function GotoHistoryQna(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("qna");

                document.history.location.href = "/cscenter/history/history_qna.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}

function GotoHistoryMemoWrite(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
            if (top.callring){
                top.document.all.callring.src = "/cscenter/ippbxmng/CallRingWithOrderFrame.asp?orderserial=" + orderserial + '&userid=' + userid;
            }else{
                top.opener.top.header.i_ippbxmng.popCallRing('','','','',orderserial,userid);
            }
            /*
        	try{
        		top.opener.top.header.i_ippbxmng.popCallRing('','','','',orderserial,userid);
            }catch(e){
            	top.opener.opener.popCallRing('','','','',orderserial,userid);
            }
            */
        	//var popwin = window.open("/cscenter/history/history_memo_write.asp?userid=" + userid + "&orderserial=" + orderserial + "&backwindow=" + "opener.document.history","GotoHistoryMemoWrite","width=600 height=600 scrollbars=yes resizable=no");
        	//popwin.focus();
        }
}

function FindByIpkumname(){
    var accountname;
    accountname = frmbuyerinfo.accountname.value;

    var gourl = "/cscenter/ordermaster/ordermaster_list.asp?searchfield=etcfield&etcfield=04&etcstring=" + accountname;

    top.listFrame.location.href = gourl;
}

// �þ�ī�� ������ǥ �˾�
function receiptallat(tid){
	var receiptUrl = "http://www.allatpay.com/servlet/AllatBizPop/member/pop_card_receipt.jsp?" +
		"shop_id=10x10_2&order_no=" + tid;
	window.open(receiptUrl,"app","width=410,height=650,scrollbars=0");
}

// �ſ�ī�� ������ǥ �˾�_�̴Ͻý�
function receiptCardRedirect(iorderserial, tid){
	var receiptUrl = "/cscenter/taxsheet/popCardReceipt.asp?orderserial=" + iorderserial +"&tid=" + tid;
	var popwin = window.open(receiptUrl,"receiptCardRedirect","width=415,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function receiptinicis(tid){
	var receiptUrl = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?" + "noTid=" + tid + "&noMethod=1";
	var popwin = window.open(receiptUrl,"INIreceipt","width=415,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// �ſ�ī�� ������ǥ �˾�_KCP
function receiptkcp(tid){
	var receiptUrl = "https://admin.kcp.co.kr/Modules/Sale/CARD/ADSA_CARD_BILL_Receipt.jsp?" +
		"c_trade_no=" + tid + "&mnu_no=AA000001";
	var popwin = window.open(receiptUrl,"KCPreceipt","width=415,height=600");
	popwin.focus();
}

// ���ں����� �˾�
function insurePrint(orderserial, mallid){
	var receiptUrl = "https://gateway.usafe.co.kr/esafe/ResultCheck.asp?oinfo=" + orderserial + "|" + mallid
	var popwin = window.open(receiptUrl,"insurePop","width=518,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

//��ũ���� ���ݿ�����
function receiptbankpay(tid){
	var receiptUrl = "http://www.bankpay.or.kr/pgmember/customcashreceipt.jsp?bill_key1=" + tid;
	var popwin = window.open(receiptUrl,"BankPayreceipt","width=400,height=560");
	popwin.focus();
}

//���ݿ����� ��û or PopUp - �̴Ͻý� �ǽð���ü or ������
function cashreceipt(iorderserial)
{
    cashreceiptInfo(iorderserial);
    /*
	var receiptUrl = "popcheckreceiptRedirect.asp?orderserial=" + iorderserial;
	var popwin = window.open(receiptUrl,"Cashreceipt","width=380,height=750,scrollbars=yes,resizable=yes");
	popwin.focus();
	*/
}

//�̴Ϸ�Ż ������ǥ PopUp
function receiptinirental(tid, mid){
	var receiptUrl = "https://inirt.inicis.com/statement/v1/statement?mid=" + mid +"&encdata=" + tid;
	var popwin = window.open(receiptUrl,"receiptinirental","width=670,height=670,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function cashreceiptInfo(iorderserial){
	var receiptUrl = "/cscenter/taxsheet/popCashReceipt.asp?orderserial=" + iorderserial;
	var popwin = window.open(receiptUrl,"Cashreceipt","width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function popMileageRequest(userid, orderserial, mileage, jukyo) {
	// �ʼ� : ���̵�
	// �ɼ� : �ֹ���ȣ, ���ϸ���, ���䳻��

	if (userid == "") {
		alert("���̵� �����ϴ�.");
		return;
	}

    var popwin = window.open('/cscenter/mileage/pop_mileage_request.asp?userid=' + userid + '&orderserial=' + orderserial + '&mileage=' + mileage + '&jukyo=' + jukyo,'popMileageRequest','width=1000,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//������� �����.
function popDacomCyberPayEdit(iorderserial){
    var popUrl = "/cscenter/cyberAcct/popCyberAcctChange.asp?orderserial=" + iorderserial;
	var popwin = window.open(popUrl,"DcCyberAcct","width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function jsResizeTextArea(obj, targetid) {
	var target = document.getElementById(targetid);

	if (target.rows == 1) {
		target.rows = 2;
		obj.value = "��";
	} else {
		target.rows = 1;
		obj.value = "��";
	}
}

function resizeTextArea(textarea, textareawidth) {
	var lines = textarea.value.split("\n");

	var textareaheight = 1;
	for (x = 0; x < lines.length; x++) {
		c = lines[x].length;

		if (c >= textareawidth) {
			textareaheight += (Math.ceil(c / textareawidth) - 1);
		}
	}
	textareaheight += (lines.length - 1);

	textarea.rows = textareaheight;
}

window.onload = function() {
	if (document.getElementById("idReqZipAddr")) {
		resizeTextArea(document.getElementById("idReqZipAddr"), 35);
		resizeTextArea(document.getElementById("idComment"), 35);
	}
}

//������
function popEstimateReceipt(orderserial){
    var window_width = 925;
    var window_height = 800;
    var popwin=window.open("/common/pop_estimate_receipt.asp?orderserial=" + orderserial ,"popOrderReceipt","width=" + window_width + " height=" + window_height + "  left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=yes resizable=yes");
    popwin.focus();
}

function foreigndirectpurchaseedit(orderserial){
	var popwin = window.open('/cscenter/ordermaster/order_foreigndirectpurchase.asp?orderserial='+orderserial,'addreg','width=400,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopOpenSongjangLog(orderserial){
	var popwin = window.open('/cscenter/delivery/DeliveryTrackingSummaryOne.asp?orderserial='+orderserial,'PopOpenSongjangLog','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jsSetExtOrder(orderserial) {
    if (confirm("���޸� �ֹ����� ��ȯ�Ͻðڽ��ϱ�?\n\n���ֹ��� ���޸� �ֹ��̾�� �մϴ�.") == true) {
        var popwin = window.open('order_info_edit_process.asp?mode=chgtoextordr&orderserial='+orderserial,'jsSetExtOrder','width=400,height=300,scrollbars=yes,resizable=yes');
	    popwin.focus();
    }
}

function jsSetTenOrder(orderserial) {
    if (confirm("���� �ֹ����� ��ȯ�Ͻðڽ��ϱ�?\n\n���ֹ��� ���޸� �ֹ��̾�� �մϴ�.") == true) {
        var popwin = window.open('order_info_edit_process.asp?mode=chgtotenordr&orderserial='+orderserial,'jsSetTenOrder','width=400,height=300,scrollbars=yes,resizable=yes');
	    popwin.focus();
    }
}

</script>

<% if (orderserial<>"") then %>
	<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="FFFFFF">
		<tr height="25">
			<td align="left">
			    &nbsp;
				<input type="button" class="button" value="��ü���" class="csbutton" style="width:60px;" onclick="javascript:PopOpenCancelOrder('<%= orderserial %>');">
				&nbsp;
				<input type="button" class="button" value="�κ����" class="csbutton" style="width:60px;" onclick="javascript:PopOpenCancelItem('<%= orderserial %>');">
				<!--
				&nbsp;
				<input type="button" class="button" value="��ǰ����" class="csbutton" style="width:60px;" onclick="javascript:PopOpenModifyOrder('<%= orderserial %>');">
				-->
				&nbsp;|&nbsp;
				<!--
				<input type="button" class="button" value="ȸ����û(�ٹ�)" class="csbutton" style="width:100px;" onclick="javascript:PopOpenReceiveItemByTenTen('<%= orderserial %>');">
				&nbsp;
				-->
				<input type="button" class="button" value="��ǰ����" class="csbutton" style="width:70px;" onclick="javascript:PopOpenReceiveItemByUpche('<%= orderserial %>');">
				&nbsp;
				<input type="button" class="button" value="��ȯ���" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemChange('<%= orderserial %>');">
				&nbsp;
				<input type="button" class="button" value="������߼�" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemOmit('<%= orderserial %>');">
				&nbsp;
				<input type="button" class="button" value="���񽺹߼�" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemMore('<%= orderserial %>');">
		        &nbsp;|&nbsp;
				<input type="button" class="button" value="��Ÿȸ��" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceRecvItemMore('<%= orderserial %>');">
		        &nbsp;|&nbsp;
				<input type="button" class="button" value="��ü��޹���" class="csbutton" style="width:90px;" onclick="javascript:PopOpenNowReadMe('<%= orderserial %>');">
				&nbsp;
				<input type="button" class="button" value="�������ǻ���" class="csbutton" style="width:90px;" onclick="javascript:PopOpenReadMe('<%= orderserial %>');">
				&nbsp;
				<input type="button" class="button" value="��ü�߰�����" class="csbutton" style="width:90px;" onclick="javascript:PopOpenUpcheAddJungsan('<%= orderserial %>');">
				&nbsp;|&nbsp;
				<!--
				&nbsp;|&nbsp;
				<input type="button" class="button" value="�ſ�ī�����" class="csbutton" style="width:90px;" onclick="javascript:PopOpenCancelCard('<%= orderserial %>');">
				-->
				<% if (C_CSPowerUser) or (TRUE) then %>
				<input type="button" class="button" value="ȯ������" class="csbutton" style="width:70px;" onclick="javascript:PopCSActionCom('','<%= orderserial %>','regcsas','A003','');">
				&nbsp;
				<% end if %>
				<input type="button" class="button" value="���ϸ�������" class="csbutton" style="width:90px;" onclick="javascript:popMileageRequest('<%= ojumun.FOneItem.FUserID %>','<%= orderserial %>',0,'');">
				&nbsp;
				<input type="button" class="button" value="���߰�����" class="csbutton" style="width:90px;" onclick="javascript:PopOpenAddPayment('<%= orderserial %>');">
		    </td>
		    <td align="right">
				<!--<input type="button" class="button" value="��Ÿ���׵��" class="csbutton" style="width:90px;" onclick="javascript:PopOpenEtcNote('<%= orderserial %>');">-->
				<!--
				<input type="button" class="button" value="�ֹ�������߼�" class="csbutton" style="width:90px;" onclick="javascript:PopCSMailSendOrder('<%= orderserial %>');">
				&nbsp;
				-->
                <input type="button" class="button" value="���庯��α�" class="csbutton" style="width:90px;" onclick="PopOpenSongjangLog('<%= orderserial %>');">
                &nbsp;
				<input type="button" class="button" value="��ǰ�̹���ON/OFF" style="width:120px;" onclick="javascript:document.orderdetail.ReloadThisPage();">
				&nbsp;
				<input type="button" class="button" value="�����������" style="width:90px;" onclick="javascript:popOrderReceipt('<%= orderserial %>');">
				&nbsp;
				<input type="button" class="button" value="������" style="width:90px;" onclick="javascript:popEstimateReceipt('<%= orderserial %>');">
			</td>
		</tr>
		</table>
	<% end if %>

	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
	<tr>
		<td width="500" align="left">
			<!-- ���������� -->
			<form name="frmbuyerinfo" onsubmit="return false;" style="margin:0px;">
			<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="24" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="5">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td>
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ ����</b>
							    	[<b><%= orderserial %></b>]
									<input type="button" class="button" value="���������۰���" class="csbutton" onclick="PopCSfileSend('','<%= orderserial %>','','');" style="width:120px;">
		    				    </td>
		    				    <td align="right">
									<% if C_CriticInfoUserLV1 then %>
		    				    	<input type="button" class="button" value="��������������" class="csbutton" onclick="javascript:PopBuyerInfo('<%= orderserial %>');" style="width:120px;">
									<% end if %>
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr height="24">
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">������ID</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 then %>
						<%= ojumun.FOneItem.FUserID %>
						<% else %>
						xxxxxxxxx
						<% end if %>
					</td>
					<td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 then %>
						<%= ojumun.FOneItem.FBuyPhone %>
						<% else %>
						XXX-XXX-XXXX
						<% end if %>
					</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">�����ڸ�</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
						<%= ojumun.FOneItem.FBuyName %>
						<% else %>
						XXX
						<% end if %>
					</td>
				    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 then %>
						[<%= ojumun.FOneItem.FBuyHp %>]<input type="hidden" name="buyhp" value="<%= ojumun.FOneItem.FBuyHp %>">
						<% elseif C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
						XXX-XXX-<%= right(ojumun.FOneItem.FBuyHp, 4) %>
						<% else %>
						XXX-XXX-XXXX
						<% end if %>
					</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 then %>
				    	<a href="javascript:PopCSSMSSend('<%= ojumun.FOneItem.FBuyHp %>','<%= ojumun.FOneItem.Forderserial %>','<%= ojumun.FOneItem.Fuserid %>','');"><font color="blue">[SMS]</font></a>
						&nbsp;
						<a href="javascript:fnClick2Call(frmbuyerinfo.buyhp);"><font color="red">[CALL]</font></a>
						<% end if %>
				    </td>
				</tr>
				<tr height="24">
					<td bgcolor="<%= adminColor("topbar") %>">
						<% if (C_InspectorUser = False) then %>
							ȸ�����
						<% end if %>
					</td>
				    <td bgcolor="#FFFFFF">
						<font color="<%= getUserLevelColorByDate(ojumun.FOneItem.fUserLevel, Left(ojumun.FOneItem.FRegDate,10)) %>">
						<%= getUserLevelStrByDate(ojumun.FOneItem.fUserLevel, Left(ojumun.FOneItem.FRegDate,10)) %></font>
					</td>
					<td bgcolor="<%= adminColor("topbar") %>">�̸���</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 then %>
							<%= ojumun.FOneItem.FBuyEmail %>
						<% else %>
							xxxxxx@xxxxxx.com
						<% end if %>
					</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 then %>
				    	<a href="javascript:PopCSMailSend('<%= ojumun.FOneItem.FBuyEmail %>','<%= ojumun.FOneItem.Forderserial %>','<%= ojumun.FOneItem.Fuserid %>');"><font color="blue">[MAIL]</font></a>
						<% end if %>
					</td>
				</tr>
				<tr height="24">
				    <td bgcolor="<%= adminColor("topbar") %>">��������</td>
				    <td colspan="4" bgcolor="#FFFFFF">

						<%

						'������, �ǽð���ü : �ݾ� : ��ü�ݾ� ���������߱� ����
						'                     ���� : cashreceiptReq : R - ���ݿ����� / T - ���ݰ�꼭
						'                     ���� : AuthCode ���� ������ ����Ϸ�
						'                     �ǽð���ü ����Ϸ� : [db_log].[dbo].tbl_cash_receipt �� ����Ÿ������ �� ���� ������ paygatetid

						'�������������� : �ְ������� : �ڵ����� �Ǵ� ����Ұ�
						'����  �������� : ���� : cashreceiptReq : R - ���ݿ����� / T - ���ݰ�꼭
						'                 ���� : R/T ���� S/U �Ϸ�

						%>

				        <!-- �������� : All@ üũī�� -->
	                  	<% if (trim(ojumun.FOneItem.Faccountdiv)="80") and (ojumun.FOneItem.FIpkumDiv >= 4) then %>
	                  	    <!-- �ÿ� ����Ʈ ���� -->
	                  	    <input type="button" class="button" value="All@��ǥ" onclick="javascript:receiptallat('<%= ojumun.FOneItem.Fpaygatetid %>');">
	                  	<% end if %>

				        <!-- �������� : �ſ�ī�� -->
	                  	<% if (ojumun.FOneItem.FAccountDiv="100") and (ojumun.FOneItem.FIpkumDiv >= 4) then %>
							<% if ojumun.FOneItem.Fpaygatetid<>"" then %>
								<% if (ojumun.FOneItem.Fpggubun = "KA") then %>
		                  		    <!-- īī������ ��ǥ -->
		                  		    <input type="button" class="button" value="KAKAO��ǥ" onclick="javascript:receiptCardRedirect('<%= orderserial %>','<%= ojumun.FOneItem.FPaygatetID %>');">
		                  		<% elseif (ojumun.FOneItem.Fpggubun = "NP") then %>
		                  		    <input type="button" class="button" value="NAVERPAY��ǥ" onclick="javascript:receiptCardRedirect('<%= orderserial %>','<%= ojumun.FOneItem.FPaygatetID %>');">
		                  		<% elseif (ojumun.FOneItem.Fpggubun = "PY") then %>
		                  		    <input type="button" class="button" value="��������ǥ" onclick="javascript:receiptCardRedirect('<%= orderserial %>','<%= ojumun.FOneItem.FPaygatetID %>');">
		                  		<% elseif (Left(ojumun.FOneItem.Fpaygatetid,9)="IniTechPG") or (Left(ojumun.FOneItem.Fpaygatetid,9)="INIMX_CAR") or (Left(ojumun.FOneItem.Fpaygatetid,9)="INIMX_ISP") or (Left(ojumun.FOneItem.Fpaygatetid,6)="Stdpay") or (Left(ojumun.FOneItem.Fpaygatetid,10)="INIAPICARD") then %>
		                  		    <!-- �̴Ͻý� ��ǥ -->
		                  		    <input type="button" class="button" value="INICIS��ǥ" onclick="javascript:receiptCardRedirect('<%= orderserial %>','<%= ojumun.FOneItem.FPaygatetID %>');">
								<% elseif (ojumun.FOneItem.Fpggubun = "TS") then %>
									<input type="button" class="button" value="TOSS��ǥ" onclick="javascript:receiptCardRedirect('<%= orderserial %>','<%= ojumun.FOneItem.FPaygatetID %>');">
		                  		<% else %>
		                  			<!-- KCP ��ǥ -->
		                  		    <input type="button" class="button" value="KCP��ǥ" onclick="javascript:receiptkcp('<%= ojumun.FOneItem.FPaygatetID %>')">
		                  		<% end if %>
		                  	<% end if %>
		                <% end if %>

				        <!-- �������� : OK+�ſ� -->
	                  	<% if (ojumun.FOneItem.FAccountDiv="110") and (ojumun.FOneItem.FIpkumDiv >= 4) then %>
	                  		<% if ojumun.FOneItem.Fpaygatetid<>"" then %>
		                  		<% if (Left(ojumun.FOneItem.Fpaygatetid,9)="IniTechPG") or (Left(ojumun.FOneItem.Fpaygatetid,9)="INIMX_CAR") or (Left(ojumun.FOneItem.Fpaygatetid,9)="INIMX_ISP") or (Left(ojumun.FOneItem.Fpaygatetid,6)="Stdpay") then %>
		                  		    <!-- �̴Ͻý� ��ǥ -->
		                  		    <input type="button" class="button" value="INICIS��ǥ(ī���)" onclick="javascript:receiptCardRedirect('<%= orderserial %>','<%= ojumun.FOneItem.FPaygatetID %>');">
		                  		<% else %>
		                  			<!-- KCP ��ǥ -->
		                  		    <input type="button" class="button" value="KCP��ǥ" onclick="javascript:receiptkcp('<%= ojumun.FOneItem.FPaygatetID %>')">
		                  		<% end if %>
		                  	<% end if %>
		                <% end if %>

		                <!-- �������� : �ǽð���ü-->
						<% if (ojumun.FOneItem.FAccountDiv="20") and (ojumun.FOneItem.FIpkumDiv >= 4) then %>
							<% if (Left(ojumun.FOneItem.Fpaygatetid,9)="IniTechPG") or (Left(ojumun.FOneItem.Fpaygatetid,10)="StdpayDBNK") or (ojumun.FOneItem.Fjumundiv="9") or (ojumun.FOneItem.Fpggubun = "NP") or (ojumun.FOneItem.Fpggubun = "PY") then %>

						        <% if ojumun.FOneItem.IsPaperRequestExist then %>
						        	<% if ojumun.FOneItem.IsPaperFinished then %>
						        		<% if ojumun.FOneItem.GetPaperType = "R" then %>
											<!-- INICIS���ݿ����� : ����Ϸ� -->
											<input type="button" class="button" value="���ݿ�����" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
								        <% else %>
								        	<input type="button" class="button" value="���ݰ�꼭" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
								        <% end if %>
							        <% else %>
						        		<% if ojumun.FOneItem.GetPaperType = "R" then %>
											<input type="button" class="button" value="������ ��û" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
								        <% else %>
								        	<input type="button" class="button" value="��꼭 ��û" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
								        <% end if %>
							        <% end if %>
						        <% else %>
						        	<a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">���º���(��û����)</a>
						        <% end if %>
							<% elseif (ojumun.FOneItem.Fpggubun = "KK") then %>
								īī������ �ڵ�����(<a href="https://m.blog.naver.com/careery/221328152248" target="_blank">����</a>)
							<% elseif (ojumun.FOneItem.Fpggubun = "TS") then %>
								�佺�ۿ��� ����
							<% else %>
						        �������..
						        <!-- BANKPAY���ݿ����� -->
						        <!-- input type="button" class="button" value="BANKPAY������" onclick="javascript:receiptbankpay('<%= ojumun.FOneItem.Fpaygatetid %>')" -->
						    <% end if %>
						<% end if %>

						<!-- �������� : ������ -->
						<% if (ojumun.FOneItem.FAccountDiv="7") then %>

					        <% if ojumun.FOneItem.IsPaperRequestExist then %>
					        	<% if ojumun.FOneItem.IsPaperFinished then %>
					        		<% if ojumun.FOneItem.GetPaperType = "R" then %>
										<!-- INICIS���ݿ����� : ����Ϸ� -->
										<input type="button" class="button" value="���ݿ�����" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
							        <% else %>
							        	<input type="button" class="button" value="���ݰ�꼭" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
							        <% end if %>
						        <% else %>
					        		<% if ojumun.FOneItem.GetPaperType = "R" then %>
										<input type="button" class="button" value="������ ��û" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
							        <% else %>
							        	<input type="button" class="button" value="��꼭 ��û" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
							        <% end if %>
						        <% end if %>
					        <% else %>
					        	<a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">���º���(��û����)</a>
					        <% end if %>

	                    <% end if %>

						<!-- �������� : �ڵ������� -->
						<% if (ojumun.FOneItem.FAccountDiv="400") then %>
						����Ұ� : ��Ż� ���û������ ����
						<% end if %>

						<!-- �������� : ������ -->
						<% if (ojumun.FOneItem.FAccountDiv="550") then %>
						����Ұ� : ������ ����翡�� �߱�
						<% end if %>

						<!-- �������� : ����Ƽ�� -->
						<% if (ojumun.FOneItem.FAccountDiv="560") then %>
						����Ұ� : ����Ƽ�� ����翡�� �߱�
						<% end if %>

						<!-- �������� : ���������� -->
						<% if (ojumun.FOneItem.FAccountDiv="50") then %>
						����Ұ� : ������ ����
						<% end if %>

						<!-- ���������κ� ��꼭/���������� - �ǽð���ü/������ ���������ϳ��� �ջ�  -->
						<% if (ojumun.FOneItem.FAccountDiv <> "7") and (ojumun.FOneItem.FAccountDiv <> "20" or (ojumun.FOneItem.FAccountDiv = "20" and ojumun.FOneItem.Fpggubun="KK")) and (ojumun.FOneItem.FIpkumDiv >= 4) and (ojumun.FOneItem.FsumPaymentEtc > 0 or ((ojumun.FOneItem.Fpggubun="NP" or ojumun.FOneItem.Fpggubun="PY") and (ojumun.FOneItem.GetPaperType = "R" or ojumun.FOneItem.GetPaperType = "S"))) then %>
					        <% if ojumun.FOneItem.IsPaperRequestExist then %>
					        	<% if ojumun.FOneItem.IsPaperFinished then %>
					        		<% if ojumun.FOneItem.GetPaperType = "R" then %>
										<!-- INICIS���ݿ����� : ����Ϸ� -->
										<input type="button" class="button" value="���ݿ�����(����)" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
							        <% else %>
							        	<input type="button" class="button" value="���ݰ�꼭(����)" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
							        <% end if %>
						        <% else %>
					        		<% if ojumun.FOneItem.GetPaperType = "R" then %>
										<input type="button" class="button" value="������ ��û(����)" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
							        <% else %>
							        	<input type="button" class="button" value="��꼭 ��û(����)" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
							        <% end if %>
						        <% end if %>
					        <% else %>
					        	<a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">���º���(����)</a>
					        <% end if %>

						<% end if %>

						<!-- ���ں������� -->
	            		<% if (ojumun.FOneItem.FInsureCd = "0") then %>
	            		    <input type="button" class="button" value="����" onclick="javascript:insurePrint('<%= ojumun.FOneItem.ForderSerial %>','ZZcube1010')">
						<% end if %>

						<!-- �̴Ϸ�Ż -->
						<% if (ojumun.FOneItem.FAccountDiv="150") then %>
							<%
								Dim iniRentalAesKey, iniRentalAesIv, iniRentalAestext, getdata, xmlHttp
								Dim iniRentalAesEncodeTid, oJSON, strData, iniRentalMid
								if (application("Svr_Info")="Dev") then
									iniRentalMid = "teenxtest1"
									iniRentalAesKey = "A2xnAKKwJpeEPg5o"
									iniRentalAesIv = "NLT8pV02NQ3zaO=="
								Else
									iniRentalMid = "teenxteenr"
									iniRentalAesKey = "TkeKg0IccDtwJACZ"
									iniRentalAesIv = "JMLi2Nnh6GL4UE=="
								End If

								iniRentalAestext = "{""tid"":"""&ojumun.FOneItem.Fpaygatetid&"""}"

								getdata = "iv="&Server.URLEncode(CStr(iniRentalAesIv))
								getdata = getdata&"&key="&Server.URLEncode(CStr(iniRentalAesKey))
								getdata = getdata&"&text="&Server.URLEncode(Cstr(iniRentalAestext))

								Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
								xmlHttp.open "GET","https://fapi.10x10.co.kr/api/web/v1/encode/aes128?"&getdata, False
								xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"  ''UTF-8 charset �ʿ�.
								xmlHttp.setTimeouts 90000,90000,90000,90000 ''2013/03/14 �߰�
								xmlHttp.Send
								strData = BinaryToText(xmlHttp.responseBody, "UTF-8")
								Set xmlHttp = Nothing

								Set oJSON = New aspJSON
								oJSON.loadJSON(strData)
								iniRentalAesEncodeTid = oJSON.data("output")
								Set oJSON = Nothing

								'// �̴Ͻý��� �����ϱ� ���ؼ� urlencode�� ��
								iniRentalAesEncodeTid = Server.URLEncode(iniRentalAesEncodeTid)
							%>
							<input type="button" class="button" value="��Ż ��༭" onclick="receiptinirental('<%=iniRentalAesEncodeTid%>', '<%=iniRentalMid%>');return false;">
						<% End If %>

				    </td>
				</tr>
			</table>
			</form>
			<!-- ���������� -->
		</td>
	    <td width="5"></td>
		<td align="left">
			<!-- ������� -->
			<form name="frmreqinfo" onsubmit="return false;" style="margin:0px;">
			<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="24" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="5">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td width="200">
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��� ����</b>
									<% If ojumun.FOneItem.FisSendGift Then %>
										<span style="color:#F2D;">[�����ϱ� �ֹ�]</span>
									<% end if %>
									<% If oUniPassNumber <> "" And Not isnull(oUniPassNumber) Then %>
										<a href="#" onclick="foreigndirectpurchaseedit('<%= orderserial %>'); return false;" target="_blank">[�ؿ�������������]</a>
									<% end if %>
		    				    </td>
		    				    <td align="right">
									<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
		    				    	<input type="button" class="button" value="�������������" class="csbutton" onclick="javascript:PopReceiverInfo('<%= orderserial %>');" style="width:120px;">
									<% end if %>
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr height="24">
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">�����θ�</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
						<%= ojumun.FOneItem.FReqName %>
						<% else %>
						XXX
						<% end if %>
					</td>
				    <td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 then %>
						<%= ojumun.FOneItem.FReqPhone %>
						<% else %>
						XXX-XXX-XXXX
						<% end if %>
					</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="24">
					<td bgcolor="<%= adminColor("topbar") %>">�����ȣ</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
						<%= ojumun.FOneItem.FReqZipCode %>
						<% end if %>
					</td>
				    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 then %>
						[<%= ojumun.FOneItem.FReqHp %>]<input type="hidden" name="reqhp" value="<%= ojumun.FOneItem.FReqHp %>">
						<% elseif C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
						XXX-XXX-<%= right(ojumun.FOneItem.FReqHp, 4) %>
						<% else %>
						XXX-XXX-XXXX
						<% end if %>
					</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 then %>
				    	<a href="javascript:PopCSSMSSend('<%= ojumun.FOneItem.FReqHp %>','<%= ojumun.FOneItem.Forderserial %>','<%= ojumun.FOneItem.Fuserid %>','');"><font color="blue">[SMS]</font></a>
						&nbsp;
						<a href="javascript:fnClick2Call(frmreqinfo.reqhp);"><font color="red">[CALL]</font></a>
						<% end if %>
				    </td>
				</tr>
				<tr height="24">
				    <td bgcolor="<%= adminColor("topbar") %>">����ּ�</td>
				    <td colspan="4" bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
				    	<textarea id="idReqZipAddr" class="textarea_ro" rows="1" cols="60" readonly><%= ojumun.FOneItem.FReqZipAddr %>&nbsp;<%= ojumun.FOneItem.FReqAddress %></textarea>
						<% end if %>
	                </td>
				</tr>
				<tr height="24">
				    <td bgcolor="<%= adminColor("topbar") %>">��Ÿ����</td>
				    <td colspan="4" bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
				        <textarea id="idComment" class="textarea_ro" rows="1" cols="60" readonly><%= ojumun.FOneItem.FComment %></textarea>
                        <% if Not IsNull(ojumun.FOneItem.Flinkorderserial) then %>
                        �����ֹ� : <%= ojumun.FOneItem.Flinkorderserial %>
                        <% end if %>
                        <% if csorderserial <> "" then %>
                        CS�ֹ� : <%= csorderserial %>
                        <% end if %>
						<% end if %>
				    </td>
				</tr>
			</table>
			</form>
			<!-- ������� -->
		</td>
	    <td width="5"></td>
		<td width="350" align="left">
			<!-- �ؿܹ���� ��� �ؿܹ�� ���� �ƴҰ��, �ö���ֹ����� -->

			<% if ojumun.FOneItem.IsForeignDeliver=true then %>
				<!-- �ؿܹ�� ���� -->
				<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr height="24" bgcolor="<%= adminColor("topbar") %>">
					    <td colspan="4">
					    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
					    		<tr>
					    			<td width="100">
					    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ؿܹ��</b>
			    				    </td>
			    				    <td align="right">
								    	<input type="button" class="button" value="�������߼�����" class="csbutton" style="width:120px;" onclick="popForeignDeliverInfo('<%= ojumun.FOneItem.FDlvcountryCode %>');">
			    				    </td>
			    				</tr>
			    			</table>
			    		</td>
					</tr>
					<tr height="24">
					    <td width="50" bgcolor="<%= adminColor("topbar") %>">��ǰ�߷�</td>
					    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FitemWeigth %>(g)</td>
					    <td width="50" bgcolor="<%= adminColor("topbar") %>">�ڽ��߷�</td>
					    <td bgcolor="#FFFFFF">200(g)</td>
					</tr>
					<tr height="24">
					    <td bgcolor="<%= adminColor("topbar") %>">��۱���</td>
					    <td colspan="3" bgcolor="#FFFFFF"><%= ojumun.FOneItem.FcountryNameEn %></td>
					</tr>
					<tr height="24">
					    <td bgcolor="<%= adminColor("topbar") %>">�����ڵ�</td>
					    <td bgcolor="#FFFFFF">
					    	<%= ojumun.FOneItem.FDlvcountryCode %>
					    	&nbsp;/&nbsp;
					    	<%= ojumun.FOneItem.FemsAreaCode %> ����
					    </td>
					    <td colspan="2" bgcolor="#FFFFFF">
							<input type="button" class="button" value="���ǥ����" class="csbutton" style="width:100px;" onclick="popForeignDeliverPay('<%= ojumun.FOneItem.FemsAreaCode %>');">
					    </td>
					</tr>
					<tr height="24">
					    <td bgcolor="<%= adminColor("topbar") %>">EMS���</td>
					    <td bgcolor="#FFFFFF"><%= FormatNumber(ojumun.FOneItem.FemsDlvCost,0) %>��</td>
					    <td bgcolor="<%= adminColor("topbar") %>">���谡��</td>
					    <td bgcolor="#FFFFFF">
					    	<%= ojumun.FOneItem.FemsInsureYn %>
					    	&nbsp;
					    	<% If ojumun.FOneItem.FemsInsureYn = "Y" Then %>
					    	<%=FormatNumber(ojumun.FOneItem.FemsInsurePrice,0)%>��
					    	<% End If %>
					    </td>
					</tr>
				</table>
			<% else %>
				<!-- �ö�� �ֹ�  -->
				<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr height="24" bgcolor="<%= adminColor("topbar") %>">
					    <td colspan="5">
					    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
					    		<tr>
					    			<td width="100">
					    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ö������</b>
			    				    </td>
			    				    <td align="right">
										<% if C_CriticInfoUserLV1 then %>
			    				    	<input type="button" class="button" value="�ö���޼�������" class="csbutton" onclick="javascript:PopFlowerDeliverInfo('<%= orderserial %>');" style="width:120px;">
										<% end if %>
			    				    </td>
			    				</tr>
			    			</table>
			    		</td>
					</tr>
					<tr height="24">
						<td width="40" bgcolor="<%= adminColor("topbar") %>">FROM</td>
					    <td width="100" bgcolor="#FFFFFF"><%= ojumun.FOneItem.Ffromname %>&nbsp;</td>
					    <td width="40" bgcolor="<%= adminColor("topbar") %>">����</td>
					    <td colspan="2" bgcolor="#FFFFFF">
					        <input type="radio" name="cardribbon" value="1" <% if ojumun.FOneItem.Fcardribbon="1" then response.write "checked" %> >ī��
					        <input type="radio" name="cardribbon" value="2" <% if ojumun.FOneItem.Fcardribbon="2" then response.write "checked" %> >����
					        <input type="radio" name="cardribbon" value="3" <% if ojumun.FOneItem.Fcardribbon="3" then response.write "checked" %> >����
					    </td>
					</tr>
					<tr height="48">
					    <td colspan="5" bgcolor="#FFFFFF">
					        <textarea class="textarea_ro" name="message" rows="2" cols="50" readonly><%= ojumun.FOneItem.Fmessage %></textarea>
					    </td>
					</tr>
					<tr height="24">
					    <td bgcolor="<%= adminColor("topbar") %>">�����</td>
					    <td colspan="4" bgcolor="#FFFFFF">
					        <%= ojumun.FOneItem.Freqdate %> ��
					        <%= ojumun.FOneItem.GetReqTimeText %>
					    </td>
					</tr>
				</table>
			<% end if %>
		</td>
	</tr>
	</table>

	<div style="line-height:40%;">
		<br />
	</div>

	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr valign="top">
		<td>
			<!-- ���Ż�ǰ���� -->
			<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
				    <td colspan="10">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td>
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���Ż�ǰ����</b>
							    	&nbsp;
							    	[<b><%= orderserial %></b>]
							    	&nbsp;
							    	<input type="button" class="button" value="����CS <%= totalascount %>��" class="csbutton" style="width:90px;" onclick="javascript:Cscenter_Action_List('<%= orderserial %>','','');">
									&nbsp;
									�ٹ�:��ǰ��[<%= ojumunitemsummary.FOneItem.Ftenbeacnt %> ��] / ����:�귣���[<%= ojumunitemsummary.FOneItem.Fbrandcnt %> ��] ��ǰ��:[<%= ojumunitemsummary.FOneItem.Fupbeacnt %> ��]
		    				    </td>
		    				    <td align="right"  width="400">
									<!--<input type="button" class="button" value="��ۺ���" class="csbutton" style="width:90px;" onclick="popBeasongCompensation('<%= orderserial %>');">-->
		    				    	<input type="button" class="button" value="��ۿϷ��Ϻ���" class="csbutton" style="width:90px;" onclick="jsPopBeasongDate('<%= orderserial %>');">
                                    <input type="button" class="button" value="������ǰ����" class="csbutton" style="width:90px;" onclick="misendmaster('<%= orderserial %>');">
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr height="295" bgcolor="#FFFFFF">
				    <td valign="top">
				        <table height="25" width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#BABABA">
				            <tr align="center" bgcolor="<%= adminColor("topbar") %>">
	                        	<td width="30">����</td>
	                        	<td width="50">�������</td>
	                        	<td width="80">CODE</td>
	                          	<td width="50">�̹���</td>
	                            <td width="200">�귣��ID</td>
	                        	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	                        	<td width="30">����</td>
								<% if (C_InspectorUser = False) then %>
	                        	<td width="60">�Һ��ڰ�<br>(+�ɼǰ�)</td>
	                        	<td width="60">�ǸŰ�<br>(���ΰ�)</td>
	                        	<td width="60">���Ű�<br>(��ǰ����)</td>
								<% end if %>
	                        	<td width="60">
									<% if (C_InspectorUser = False) then %>
									���ʽ�����<br>���밡
									<% else %>
										�ǰ�����
									<% end if %>
								</td>
								<td width="60">
									<% if (C_InspectorUser = False) then %>
										��Ÿ����<br>���밡
									<% else %>
										�ǰ�����
									<% end if %>
								</td>
								<td width="60">����<br>���ϸ���</td>
								<td width="60">���԰�</td>

	                        	<td width="70">�뺸��<br>Ȯ����</td>
	                        	<td width="125">�����<br>�������</td>
	                        </tr>
	                        <tr>
	                            <td height="1" colspan="13" bgcolor="#BABABA"></td>
	                        </tr>
	                     </table>
	                     <table height="270" width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
	                        <tr height="100%">
	                            <td colspan="13" style="vertical-align: text-top;">
	                    	        <iframe name="orderdetail" src="orderitemmaster.asp?orderserial=<%= orderserial %>" border="0" frameborder="no" frameSpacing=0  width="100%" height="100%" leftmargin="0"></iframe>
	                            </td>
	                        <tr>
	                    </table>
				    </td>
				</tr>
			</table>
			<!-- ���Ż�ǰ���� -->
		</td>
	</tr>
	</table>

	<div style="line-height:40%;">
		<br />
	</div>

	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr height="100">
		<td valign="top">
		    <!-- �ֹ��� History -->
		    <form name="frmhistory" onsubmit="return false;" style="margin:0px;">
		    <table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
				    <td>
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td>
									<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryMemo('','<%= orderserial %>')"><b>MEMO</b></a>
		    				    	[<b><%= orderserial %></b>]
		    				    	|
		    				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryCS('','<%= orderserial %>')"><b>CS ó����</b></a>
		                            |
		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryMileage('<%= ojumun.FOneItem.FUserID %>','')"><b>���ϸ���</b></a>
		                            [<b><%= ojumun.FOneItem.FUserID %></b>]
		                            |
		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryDeposit('<%= ojumun.FOneItem.FUserID %>','')"><b>��ġ��</b></a>
		                            [<b><%= ojumun.FOneItem.FUserID %></b>]
		                            |
									<% if (C_InspectorUser = False) then %>
		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryCoupon('<%= ojumun.FOneItem.FUserID %>','')"><b>����</b></a>
		                            |
									<% end if %>
		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryQna('<%=ojumun.FOneItem.FUserID%>','<%If ojumun.FOneItem.FUserID = "" Then response.write orderserial End If %>')"><b>1:1���</b></a>
									<% end if %>
		    				    </td>
		    				    <td width="100" align="right">
									<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
		    				    	<input type="button" class="button" name="writebutton" value="History���" class="csbutton" onclick="OpenHistoryWindow('<%= ojumun.FOneItem.FUserID %>','<%= orderserial %>');">
									<% end if %>
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr>
				    <td style="background-color:#FFFFFF;">
				        <iframe name="history" src="blank.asp" border="0" frameSpacing=0 frameborder="no" width="100%" height="100%" leftmargin="0"></iframe>
	`			    </td>
				</tr>
			</table>
			</form>
			<!-- �ֹ��� History-->
		</td>
		<td width="5"></td>
		<td width="250" align="left" valign="top">
		    <!-- �ֹ����� -->
		    <table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="24" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="3">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td width="100">
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ� ����</b>
		    				    </td>
		    				    <td align="right">
		    				    	<input type="button" class="button" value="������������" class="csbutton" onclick="PopNextIpkumDiv('<%= orderserial %>');">
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr height="22">
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">�������</td>
				    <td bgcolor="#FFFFFF">
				        <%= ojumun.FOneItem.JumunMethodName %>
				        [<font color="<%= ojumun.FOneItem.IpkumDivColor %>"><%= ojumun.FOneItem.IpkumDivName %></font>]
				        <% if ojumun.FOneItem.FCancelYn<>"N" then %>
				        <font color="<%= ojumun.FOneItem.CancelYnColor %>"><%= ojumun.FOneItem.CancelYnName %></font>
				        <% end if %>
				        <% if ojumun.FOneItem.FokcashbagSpend<>0 then %>
				        <br>(ĳ�ù��� : <strong><%= formatNumber(ojumun.FOneItem.FokcashbagSpend,0) %></strong>)
				        <% end if %>
				    </td>
				</tr>
				<% if ojumun.FOneItem.FAccountDiv="7" then %>
				<tr height="22">
				    <td bgcolor="<%= adminColor("topbar") %>">�Աݰ���</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 then %>
				    	<%= ojumun.FOneItem.FAccountNo %>
				    	&nbsp;
						<% end if %>
				    	<% if ojumun.FOneItem.IsDacomCyberAccountPay then %>
					    <a href="javascript:popDacomCyberPayEdit('<%= orderserial %>')"><font color="red">[����]</font></a>
					    <% else %>
					    <a href="javascript:popDacomCyberPayEdit('<%= orderserial %>')">[�Ϲ�]</a>
					    <% end if %>
				    </td>
				</tr>
				<% end if %>
				<tr height="22">
				    <td bgcolor="<%= adminColor("topbar") %>">�ֹ��Ͻ�</td>
				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FRegDate %></td>
				</tr>
				<tr height="22">
				    <td bgcolor="<%= adminColor("topbar") %>">�Ա�Ȯ��</td>
				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FIpkumDate %></td>
				</tr>
				<tr height="22">
				    <td bgcolor="<%= adminColor("topbar") %>">�ֹ��뺸</td>
				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.Fbaljudate %></td>
				</tr>
				<!--
				<tr height="22">
				    <td bgcolor="<%= adminColor("topbar") %>">����Ͻ�</td>
				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.Fbeadaldate %></td>
				</tr>
				-->
				<tr height="24">
				    <td bgcolor="<%= adminColor("topbar") %>">���ι�ȣ</td>
				    <td bgcolor="#FFFFFF">
				        <input type="text" class="text_ro" value="<%= ojumun.FOneItem.FAuthcode %>" readonly size="25">
				    </td>
				</tr>
				<tr height="24">
				    <td bgcolor="<%= adminColor("topbar") %>">PG��</td>
				    <td bgcolor="#FFFFFF"><%= fnGetPggubunName(ojumun.FOneItem.Fpggubun) %></td>
				</tr>
				<tr height="24">
				    <td bgcolor="<%= adminColor("topbar") %>">PG�� TID</td>
				    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" value="<%= ojumun.FOneItem.FPaygatetID %>" readonly size="25"></td>
				</tr>
			</table>
			<!-- �ֹ����� -->
		</td>
		<td width="5"></td>
		<td width="250" align="left" valign="top">
		    <!-- �������� -->
		    <table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="24" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="3">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td width="100">
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���� ����</b>
                                    <% if C_ADMIN_AUTH or C_CSPowerUser then %>
                                    <% if (ojumun.FOneItem.Fsitename = "10x10") or (ojumun.FOneItem.Fsitename = "10x10cs") then %>
                                    <input type="button" class="button" value="������ȯ" style="width:80px;" onClick="jsSetExtOrder('<%= orderserial %>');">
                                    <% else %>
                                    <input type="button" class="button" value="������ȯ" style="width:80px;" onClick="jsSetTenOrder('<%= orderserial %>');">
                                    <% end if %>
                                    (��Ʈ�����)
                                    <% end if %>
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr height="22">
				    <td width="100" bgcolor="<%= adminColor("topbar") %>">�ְ���(����)</td>
				    <td bgcolor="#FFFFFF">
						<% for ix = 0 to oetcpayment.FResultCount - 1 %>
						<% if (oetcpayment.FItemList(ix).Facctdiv = ojumun.FOneItem.FAccountDiv) then %>
						<%= FormatNumber(oetcpayment.FItemList(ix).FrealPayedsum, 0) %> ��
						<% if (oetcpayment.FItemList(ix).FrealPayedsum <> oetcpayment.FItemList(ix).Facctamount) then %>
						(<%= FormatNumber(oetcpayment.FItemList(ix).Facctamount, 0) %> ��)
						<% end if %>
						<% end if %>
						<% next %>
				    </td>
				</tr>
				<tr height="22">
				    <td bgcolor="<%= adminColor("topbar") %>">���������հ�</td>
				    <td bgcolor="#FFFFFF">
						<%= FormatNumber(ojumun.FOneItem.FsumPaymentEtc, 0) %> ��
				    </td>
				</tr>
				<% for ix = 0 to oetcpayment.FResultCount - 1 %>
				<% if (oetcpayment.FItemList(ix).Facctdiv <> ojumun.FOneItem.FAccountDiv) then %>
							<tr height="22">
							    <td bgcolor="<%= adminColor("topbar") %>"> - <%= oetcpayment.FItemList(ix).FacctdivName %></td>
							    <td bgcolor="#FFFFFF">
							    	<%= FormatNumber(oetcpayment.FItemList(ix).FrealPayedsum, 0) %> ��<br>
							    	<% if (oetcpayment.FItemList(ix).FrealPayedsum <> oetcpayment.FItemList(ix).Facctamount) then %>
							    	(<%= FormatNumber(oetcpayment.FItemList(ix).Facctamount, 0) %> ��)
							    	<% end if %>
							    </td>
							</tr>
				<% end if %>
				<% next %>
			</table>
			<!-- �������� -->
		</td>
	</tr>
	</table>
<% else %>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	    <tr height="50">
	        <td align="center"> [ �󼼳����� ���÷��� �ֹ���ȣ�� ���� �ϼ��� ]</td>
	    </tr>
	</table>
<% end if %>

<% if (orderserial <> "") then %>
	<script type="text/javascript">
	    GotoHistoryCS('','<%= orderserial %>');
	</script>
<% end if %>

<%
set ojumun = Nothing
set oaslist = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
