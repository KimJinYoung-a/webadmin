<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_tenGiftCardCls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%
dim giftorderserial, totalascount, ix, yyyy, mm, dd, hh, bookingDate
	giftorderserial = RequestCheckVar(request("giftorderserial"),11)

totalascount = 0

dim oGiftOrder
set oGiftOrder = new cGiftCardOrder
	if (giftorderserial <> "") then
		oGiftOrder.FRectGiftOrderSerial = giftorderserial
		oGiftOrder.getCSGiftcardOrderDetail
	end if

dim oaslist
set oaslist = new CCSASList
	if (giftorderserial <> "") then
	    oaslist.FRectOrderSerial = giftorderserial
	    oaslist.GetCSASTotalCount
	    totalascount = oaslist.FResultCount
	end if

if (giftorderserial <> "") then
	if (oGiftOrder.FOneItem.FbookingYn = "Y") and (oGiftOrder.FOneItem.FbookingDate <> "") then
		yyyy = Year(oGiftOrder.FOneItem.FbookingDate)
		mm = Right("0" & (Month(oGiftOrder.FOneItem.FbookingDate) + 1), 2)
		dd = Right("0" & (Day(oGiftOrder.FOneItem.FbookingDate)), 2)
		hh = Right("0" & (Hour(oGiftOrder.FOneItem.FbookingDate)), 2)

		bookingDate = yyyy & "-" & mm & "-" & dd & " " & hh
	end if
end if

%>

<script type="text/javascript">

function misendmaster(v){
	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + v,"misendmaster","width=1200 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cs_mileage(v){
	var popwin = window.open("/cscenter/mileage/cs_mileage.asp?userid=" + v,"cs_mileage","width=1000 height=700 scrollbars=yes resizable=yes");
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
	var popwin = window.open("orderdetail_editorder.asp?orderserial=" + orderserial,"PopOpenModifyOrder","width=1200 height=800 scrollbars=yes resizable=yes");
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

    if (menuname == "memo") {
            document.frmhistory.writebutton.value = "MEMO�ۼ�";
    } else if (menuname == "cs") {
            document.frmhistory.writebutton.value = "CS����Ʈ";
    } else if (menuname == "mileage") {
            document.frmhistory.writebutton.value = "���ϸ�������";
    } else if (menuname == "coupon") {
            document.frmhistory.writebutton.value = "��������";
    } else if (menuname == "qna") {
            document.frmhistory.writebutton.value = "1:1������";
    }
}

function OpenHistoryWindow(userid, orderserial) {
    if (selected_history_menu == "memo") {
		GotoHistoryMemoWrite(userid, orderserial);
    } else if (selected_history_menu == "cs") {
		Cscenter_Action_List(orderserial,'','')
    } else if (selected_history_menu == "mileage") {
		cs_mileage(userid)
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
	var popwin = window.open(receiptUrl,"rdINIreceipt","width=415,height=600,scrollbars=yes,resizable=yes");
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

    var popwin = window.open('/cscenter/mileage/pop_mileage_request.asp?userid=' + userid + '&orderserial=' + orderserial + '&mileage=' + mileage + '&jukyo=' + jukyo,'popMileageRequest','width=660,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//������� �����.
function popDacomCyberPayEdit(iorderserial){
    var popUrl = "/cscenter/cyberAcct/popCyberAcctChange.asp?orderserial=" + iorderserial;
	var popwin = window.open(popUrl,"DcCyberAcct","width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function PopBuyerInfo(giftorderserial) {
	if (giftorderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 300;
    var window_height = 250;
	var popwin = window.open("cs_giftcard_order_buyer_info.asp?giftorderserial=" + giftorderserial,"PopBuyerInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopReceiverInfo(giftorderserial) {
	if (giftorderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 250;
    var window_height = 100;
	var popwin = window.open("cs_giftcard_order_receiver_info.asp?giftorderserial=" + giftorderserial,"PopReceiverInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopModiPGKey(giftorderserial) {
	if (giftorderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 500;
    var window_height = 100;
	var popwin = window.open("cs_giftcard_order_pgkey.asp?giftorderserial=" + giftorderserial,"PopModiPGKey","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopNextIpkumDiv(giftorderserial){
    if (giftorderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 400;
    var window_height = 200;
	var popwin = window.open("cs_giftcard_order_nextstep.asp?giftorderserial=" + giftorderserial,"PopNextIpkumDiv","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopPrevIpkumDiv(giftorderserial){
    if (giftorderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 300;
    var window_height = 160;
	var popwin = window.open("cs_giftcard_order_prevstep.asp?giftorderserial=" + giftorderserial,"PopPrevIpkumDiv","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopOpenCancelOrder(giftorderserial){
    if (giftorderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 1000;
    var window_height = 800;
	var popwin = window.open("pop_cs_giftcard_action_new.asp?giftorderserial=" + giftorderserial + "&mode=regcsas&divcd=A008&ckAll=on","PopOpenCancelOrder","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopMMSInfo(giftorderserial){
    if (giftorderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 500;
    var window_height = 300;
	var popwin = window.open("cs_giftcard_order_mms_info.asp?giftorderserial=" + giftorderserial,"PopMMSInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopResendMMS(giftorderserial, iscreatenewcode){
    if (giftorderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 500;
    var window_height = 300;
	var popwin = window.open("cs_giftcard_order_resendmms.asp?giftorderserial=" + giftorderserial + '&iscreatenewcode=' + iscreatenewcode,"PopResendMMS","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

function PopEmailInfo(giftorderserial){
    if (giftorderserial == "") {
        alert("���� �ֹ��� �����ϼ���.");
        return;
    }

    var window_width = 500;
    var window_height = 300;
	var popwin = window.open("cs_giftcard_order_email_info.asp?giftorderserial=" + giftorderserial,"PopEmailInfo","width=" + window_width + " height=" + window_height + " left=" + GetCenterX(window_width) + " top=" + GetCenterY(window_height) + " scrollbars=no resizable=yes status=no");
	popwin.focus();
}

</script>

<% if (giftorderserial<>"") then %>
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="FFFFFF">
	<tr height="25">
		<td align="left">
		    &nbsp;&nbsp;
			<input type="button" class="button" value="��ü���" class="csbutton" style="width:60px;" onclick="javascript:PopOpenCancelOrder('<%= giftorderserial %>');">
	    </td>
	</tr>
	</table>

	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
	<tr>
		<td width="400" align="left">
			<!-- ���������� -->
			<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<form name="frmbuyerinfo" onsubmit="return false;">
				<tr height="24" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="5">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td>
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ ����</b>
							    	[<b><%= giftorderserial %></b>]
		    				    </td>
		    				    <td align="right">
		    				    	<input type="button" class="button" value="��������������" class="csbutton" onclick="javascript:PopBuyerInfo('<%= giftorderserial %>');">
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr height="24">
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">������ID</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 then %>
							<%= oGiftOrder.FOneItem.FUserID %>
						<% else %>
							<%= printUserId(oGiftOrder.FOneItem.FUserID, 2, "*") %>
						<% end if %>
				    </td>
					<td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
							<%= oGiftOrder.FOneItem.FBuyPhone %>
						<% else %>
							----
						<% end if %>
				    </td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">�����ڸ�</td>
				    <td bgcolor="#FFFFFF"><%= oGiftOrder.FOneItem.FBuyName %></td>
				    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
							<%= oGiftOrder.FOneItem.FBuyHp %>
						<% else %>
							<%= printUserId(oGiftOrder.FOneItem.FBuyHp, 2, "*") %>
						<% end if %>
				    </td>
				    <td bgcolor="#FFFFFF">
				    	<a href="javascript:PopCSSMSSend('<%= oGiftOrder.FOneItem.FBuyHp %>','<%= oGiftOrder.FOneItem.FgiftOrderSerial %>','<%= oGiftOrder.FOneItem.Fuserid %>','');"><font color="blue">[SMS]</font></a>
				    </td>
				</tr>
				<tr height="24">
					<td bgcolor="<%= adminColor("topbar") %>">ȸ�����</td>
				    <td bgcolor="#FFFFFF">
				    	<font color="<%= getUserLevelColorByDate(oGiftOrder.FOneItem.fUserLevel, left(oGiftOrder.FOneItem.FRegDate,10)) %>">
						<%= getUserLevelStrByDate(oGiftOrder.FOneItem.fUserLevel, left(oGiftOrder.FOneItem.FRegDate,10)) %></font>
				    </td>
					<td bgcolor="<%= adminColor("topbar") %>">�̸���</td>
				    <td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
							<%= oGiftOrder.FOneItem.FBuyEmail %>
						<% else %>
							<%= printUserId(oGiftOrder.FOneItem.FBuyEmail, 2, "*") %>
						<% end if %>
				    </td>
				    <td bgcolor="#FFFFFF">
				    	<a href="javascript:PopCSMailSend('<%= oGiftOrder.FOneItem.FBuyEmail %>','<%= oGiftOrder.FOneItem.FgiftOrderSerial %>','<%= oGiftOrder.FOneItem.Fuserid %>');"><font color="blue">[MAIL]</font></a>
					</td>
				</tr>
				<tr height="24">
				    <td bgcolor="<%= adminColor("topbar") %>">��������</td>
				    <td colspan="4" bgcolor="#FFFFFF" height="47">

				        <!-- �������� : All@ üũī�� -->
	                  	<% if (trim(oGiftOrder.FOneItem.Faccountdiv)="80") and (oGiftOrder.FOneItem.FIpkumDiv >= 4) then %>

	                  	    Giftī��� ��ǰ�� �ƴϹǷ� ��꼭/������ ����Ұ�<br>(���� ��ǰ ���Ž� ���డ��)

	                  	<% end if %>

				        <!-- �������� : �ſ�ī�� -->
	                  	<% if (oGiftOrder.FOneItem.FAccountDiv="100") and (oGiftOrder.FOneItem.FIpkumDiv >= 4) then %>
	                  		<% if oGiftOrder.FOneItem.Fpaydateid<>"" then %>
		                  		<% if (Left(oGiftOrder.FOneItem.Fpaydateid,9)="IniTechPG") then %>
		                  		    <!-- �̴Ͻý� ��ǥ -->
		                  		    <input type="button" class="button" value="INICIS��ǥ" onclick="javascript:receiptCardRedirect('<%= giftorderserial %>','<%= oGiftOrder.FOneItem.Fpaydateid %>');">
		                  		<% else %>
		                  			<!-- KCP ��ǥ -->
		                  		    <input type="button" class="button" value="KCP��ǥ" onclick="javascript:receiptkcp('<%= oGiftOrder.FOneItem.Fpaydateid %>')">
		                  		<% end if %>
		                  	<% end if %>
		                <% end if %>

				        <!-- �������� : OK+�ſ� -->
	                  	<% if (oGiftOrder.FOneItem.FAccountDiv="110") and (oGiftOrder.FOneItem.FIpkumDiv >= 4) then %>

	                  		Giftī��� ��ǰ�� �ƴϹǷ� ��꼭/������ ����Ұ�<br>(���� ��ǰ ���Ž� ���డ��)

		                <% end if %>

		                <!-- �������� : �ǽð���ü-->
						<% if (oGiftOrder.FOneItem.FAccountDiv="20") and (oGiftOrder.FOneItem.FIpkumDiv >= 4) then %>

						    Giftī��� ��ǰ�� �ƴϹǷ� ��꼭/������ ����Ұ�<br>(���� ��ǰ ���Ž� ���డ��)

						<% end if %>

						<!-- �������� : ������ -->
						<% if (oGiftOrder.FOneItem.FAccountDiv="7") then %>

					        Giftī��� ��ǰ�� �ƴϹǷ� ��꼭/������ ����Ұ�<br>(���� ��ǰ ���Ž� ���డ��)

	                    <% end if %>

						<!-- �������� : �ڵ������� -->
						<% if (oGiftOrder.FOneItem.FAccountDiv="400") then %>

						Giftī��� ��ǰ�� �ƴϹǷ� ��꼭/������ ����Ұ�<br>(���� ��ǰ ���Ž� ���డ��)

						<% end if %>

						<!-- �������� : ���������� -->
						<% if (oGiftOrder.FOneItem.FAccountDiv="50") then %>

						Giftī��� ��ǰ�� �ƴϹǷ� ��꼭/������ ����Ұ�<br>(���� ��ǰ ���Ž� ���డ��)

						<% end if %>

						<!-- ���������κ� ��꼭/���������� - �ǽð���ü/������ ���������ϳ��� �ջ�  -->
						<% if (oGiftOrder.FOneItem.FAccountDiv <> "7") and (oGiftOrder.FOneItem.FAccountDiv <> "20") and (oGiftOrder.FOneItem.FIpkumDiv >= 4) and (oGiftOrder.FOneItem.FsumPaymentEtc > 0) then %>

					        Giftī��� ��ǰ�� �ƴϹǷ� ��꼭/������ ����Ұ�<br>(���� ��ǰ ���Ž� ���డ��)

						<% end if %>

						<!-- ���ں������� -->
	            		<% if (oGiftOrder.FOneItem.FInsureCd = "0") then %>
	            		    <input type="button" class="button" value="����" onclick="javascript:insurePrint('<%= oGiftOrder.FOneItem.FgiftOrderSerial %>','ZZcube1010')">
						<% end if %>

				    </td>
				</tr>
				</form>
			</table>
			<!-- ���������� -->
		</td>
	    <td width="5"></td>
		<td width="300" align="left">
			<!-- ������� -->
			<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<form name="frmreqinfo" onsubmit="return false;">
				<tr height="24" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="5">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td width="100">
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ ����</b>
		    				    </td>
		    				    <td align="right">
		    				    	<input type="button" class="button" value="������ ��������" class="csbutton" onclick="javascript:PopReceiverInfo('<%= giftorderserial %>');">
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr height="24">
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
					<td bgcolor="#FFFFFF" colspan="3">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
							<%= oGiftOrder.FOneItem.FReqHp %>
						<% else %>
							<%= printtel(oGiftOrder.FOneItem.FReqHp) %>
						<% end if %>
					</td>
				    <td bgcolor="#FFFFFF">
				    	<a href="javascript:PopCSSMSSend('<%= oGiftOrder.FOneItem.FReqHp %>','<%= oGiftOrder.FOneItem.FgiftOrderSerial %>','<%= oGiftOrder.FOneItem.Fuserid %>','');"><font color="blue">[SMS]</font></a>
				    </td>
				</tr>
				<tr height="24">
					<td bgcolor="<%= adminColor("topbar") %>">�̸���</td>
					<td bgcolor="#FFFFFF" colspan="3">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
							<%= oGiftOrder.FOneItem.FReqEmail %>
						<% else %>
							<%= printUserId(oGiftOrder.FOneItem.FReqEmail, 2, "*") %>
						<% end if %>
					</td>
				    <td bgcolor="#FFFFFF">
				   		<a href="javascript:PopCSMailSend('<%= oGiftOrder.FOneItem.FReqEmail %>','<%= oGiftOrder.FOneItem.FgiftOrderSerial %>','<%= oGiftOrder.FOneItem.Fuserid %>');"><font color="blue">[MAIL]</font></a>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>" height="72">��Ÿ����</td>
				    <td colspan="4" bgcolor="#FFFFFF">

				    </td>
				</tr>
				</form>
			</table>
			<!-- ������� -->
		</td>
	    <td width="5"></td>
		<td align="left">
		    <!-- �ֹ����� -->
		    <table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="24" bgcolor="<%= adminColor("topbar") %>">
			    <td colspan="3">
			    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			    		<tr>
			    			<td width="80">
			    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ� ����</b>
	    				    </td>
	    				    <td align="right">
	    				    	<input type="button" class="button" value="����������ȯ" class="csbutton" onclick="javascript:PopPrevIpkumDiv('<%= giftorderserial %>');">
	    				    	<input type="button" class="button" value="������������" class="csbutton" onclick="javascript:PopNextIpkumDiv('<%= giftorderserial %>');">
	    				    </td>
	    				</tr>
	    			</table>
	    		</td>
			</tr>
			<tr height="22">
			    <td width="50" bgcolor="<%= adminColor("topbar") %>">�������</td>
			    <td bgcolor="#FFFFFF">
			        <%= oGiftOrder.FOneItem.GetAccountdivName %>
			        [<font color="<%= oGiftOrder.FOneItem.IpkumDivColor %>"><%= oGiftOrder.FOneItem.GetIpkumDivName %></font>]
			        <% if oGiftOrder.FOneItem.FCancelYn<>"N" then %>
			        <font color="<%= oGiftOrder.FOneItem.CancelYnColor %>"><%= oGiftOrder.FOneItem.CancelYnName %></font>
			        <% end if %>
			    </td>
			</tr>
			<% if oGiftOrder.FOneItem.FAccountDiv="7" then %>
			<tr height="22">
			    <td bgcolor="<%= adminColor("topbar") %>">�Աݰ���</td>
			    <td bgcolor="#FFFFFF">
			    	<%= oGiftOrder.FOneItem.FAccountNo %>
			    	&nbsp;
			    	<% if oGiftOrder.FOneItem.IsDacomCyberAccountPay then %>
				    <a href="javascript:popDacomCyberPayEdit('<%= giftorderserial %>')"><font color="red">[����]</font></a>
				    <% else %>
				    <a href="javascript:popDacomCyberPayEdit('<%= giftorderserial %>')">[�Ϲ�]</a>
				    <% end if %>
			    </td>
			</tr>
			<% end if %>
			<tr height="22">
			    <td bgcolor="<%= adminColor("topbar") %>">�ֹ��Ͻ�</td>
			    <td bgcolor="#FFFFFF"><%= oGiftOrder.FOneItem.FRegDate %></td>
			</tr>
			<tr height="22">
			    <td bgcolor="<%= adminColor("topbar") %>">�Ա�Ȯ��</td>
			    <td bgcolor="#FFFFFF"><%= oGiftOrder.FOneItem.FIpkumDate %></td>
			</tr>
			<tr height="24">
			    <td bgcolor="<%= adminColor("topbar") %>">���ι�ȣ</td>
			    <td bgcolor="#FFFFFF">
			        <input type="text" class="text_ro" value="<%= oGiftOrder.FOneItem.FAuthcode %>" readonly size="25">
			    </td>
			</tr>
			<tr height="24">
			    <td bgcolor="<%= adminColor("topbar") %>">PG�� ID</td>
			    <td bgcolor="#FFFFFF">
					<input type="text" class="text_ro" value="<%= oGiftOrder.FOneItem.Fpaydateid %>" readonly size="35">
					<input type="button" class="button" value="����" class="csbutton" onclick="javascript:PopModiPGKey('<%= giftorderserial %>');">
				</td>
			</tr>
			</table>
			<!-- �ֹ����� -->
		</td>
	</tr>
	</table>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr valign="top">
		<td>
			<!-- ���Ż�ǰ���� -->
			<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
				    <td colspan="10">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			    		<tr>
			    			<td width="500">
			    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���Ż�ǰ����</b>
						    	&nbsp;
						    	[<b><%= giftorderserial %></b>]
						    	&nbsp;
						    	<input type="button" class="button" value="����CS <%= totalascount %>��" class="csbutton" style="width:90px;" onclick="javascript:Cscenter_Action_List('<%= giftorderserial %>','','');">
	    				    </td>
	    				    <td align="right">

	    				    </td>
	    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr height="100" bgcolor="#FFFFFF">
				    <td valign="top">
						<table height="25" width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#BABABA">
						<tr align="center" bgcolor="<%= adminColor("topbar") %>">
							<td width="30" height="40">����</td>
							<td width="50">�������</td>
							<td width="40">CODE</td>
						  	<td width="50">�̹���</td>
							<td>Giftī���<br><font color="blue">[�ɼ�]</font></td>
							<td width="60">�ǸŰ�</td>

							<td width="100">������</td>
							<td width="100">�����</td>
							<td width="100">�����</td>
						</tr>
						<tr>
						    <td height="1" colspan="10" bgcolor="#BABABA"></td>
						</tr>
						<tr align="center" bgcolor="<%= adminColor("topbar") %>">
							<td height="60"></td>
							<td><%= oGiftOrder.FOneItem.GetCardStatusName %></td>
							<td><%= oGiftOrder.FOneItem.FcardItemid %></td>
							<td><img src="<%= oGiftOrder.FOneItem.FSmallimage %>"></td>
							<td>
								<%= oGiftOrder.FOneItem.FCarditemname %><br><font color="blue">[<%= oGiftOrder.FOneItem.FcardOptionName %>]</font>
							</td>
							<td><%= FormatNumber(oGiftOrder.FOneItem.Fsubtotalprice, 0) %></td>
							<td><%= Left(oGiftOrder.FOneItem.FsendDate, 10) %></td>
							<td><%= oGiftOrder.FOneItem.FcardregDate %></td>
							<td><%= oGiftOrder.FOneItem.Fcanceldate %></td>
						</tr>
						 </table>
				    </td>
				</tr>
			</table>
			<!-- ���Ż�ǰ���� -->
		</td>
	</tr>
	</table>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
	<tr>
		<td width="600" align="left">
			<!-- ������� -->
			<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<form name="frmreqinfo" onsubmit="return false;">
				<tr height="24" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="2">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td width="80">
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���� ����</b>
		    				    </td>
		    				    <td align="right">
		    				    	<% if (oGiftOrder.FOneItem.Fjumundiv = "5") then %>
			    				    	<input type="button" class="button" value="���������ڵ� ������" class="csbutton" onclick="javascript:PopResendMMS('<%= giftorderserial %>', 'N');">
			    				    	<input type="button" class="button" value="�ű������ڵ� ����" class="csbutton" onclick="javascript:PopResendMMS('<%= giftorderserial %>', 'Y');">
		    				    	<% end if %>
		    				    	<input type="button" class="button" value="��������" class="csbutton" onclick="javascript:PopMMSInfo('<%= giftorderserial %>');">
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr height="24">
				    <td width="100" bgcolor="<%= adminColor("topbar") %>">��������</td>
					<td bgcolor="#FFFFFF">
				    	<% if (oGiftOrder.FOneItem.FbookingYn = "Y") then %>
				    		��������
				    	<% else %>
				    		�������
				    	<% end if %>
					</td>
				</tr>
				<tr height="24">
					<td bgcolor="<%= adminColor("topbar") %>">�����Ͻ�</td>
					<td bgcolor="#FFFFFF">
						<% if (oGiftOrder.FOneItem.FbookingYn = "Y") then %>
						<%= bookingDate %>
						<% end if %>
					</td>
				</tr>
				<tr height="24">
				    <td bgcolor="<%= adminColor("topbar") %>">�����º�HP</td>
					<td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
							<%= oGiftOrder.FOneItem.Fsendhp %>
						<% else %>
							<%= printtel(oGiftOrder.FOneItem.Fsendhp) %>
						<% end if %>
					</td>
				</tr>
				<tr height="24">
				    <td bgcolor="<%= adminColor("topbar") %>">�޴º�HP</td>
					<td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
							<%= oGiftOrder.FOneItem.Freqhp %>
						<% else %>
							<%= printtel(oGiftOrder.FOneItem.Freqhp) %>
						<% end if %>
					</td>
				</tr>

				<tr height="24">
					<td bgcolor="<%= adminColor("topbar") %>">MMS ����</td>
					<td bgcolor="#FFFFFF"><%= oGiftOrder.FOneItem.FMMSTitle %></td>
				</tr>
				<tr height="150">
					<td bgcolor="<%= adminColor("topbar") %>">MMS ����</td>
					<td bgcolor="#FFFFFF"><%= nl2br(oGiftOrder.FOneItem.FMMSContent) %></td>
				</tr>
				</form>
			</table>
			<!-- ������� -->
		</td>
	    <td width="5"></td>
		<td width="600" align="left">
			<!-- ������� -->
			<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="24" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="3">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td width="150">
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�����̸��� ����</b>
		    				    </td>
		    				    <td align="right">
		    				    	<input type="button" class="button" value="�����̸��� ��������" class="csbutton" onclick="javascript:PopEmailInfo('<%= giftorderserial %>');">
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr height="24">
				    <td width="100" bgcolor="<%= adminColor("topbar") %>">���ۿ���</td>
					<td bgcolor="#FFFFFF">
				    	<% if (oGiftOrder.FOneItem.FsendDiv = "E") then %>
				    		��������
				    	<% else %>
				    		�߼۾���
				    	<% end if %>
					</td>
				</tr>
				<tr height="24">
				    <td bgcolor="<%= adminColor("topbar") %>">�����º� Email</td>
					<td bgcolor="#FFFFFF">
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
							<%= oGiftOrder.FOneItem.Fsendemail %>
						<% else %>
							<%= printUserId(oGiftOrder.FOneItem.Fsendemail, 2, "*") %>
						<% end if %>
					</td>
				</tr>
				<tr height="24">
				    <td bgcolor="<%= adminColor("topbar") %>">�޴º� Email</td>
					<td bgcolor="#FFFFFF"><%= oGiftOrder.FOneItem.FreqEmail %></td>
				</tr>
				<tr height="24">
					<td bgcolor="<%= adminColor("topbar") %>">Email ����</td>
					<td bgcolor="#FFFFFF"><%= oGiftOrder.FOneItem.FemailTitle %></td>
				</tr>
				<tr height="175">
					<td bgcolor="<%= adminColor("topbar") %>">Email ����</td>
					<td bgcolor="#FFFFFF"><%= nl2br(oGiftOrder.FOneItem.FemailContent) %></td>
				</tr>
			</table>
			<!-- ������� -->
		</td>
	</tr>
	</table>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr height="100">
		<td valign="top">
		    <!-- �ֹ��� History -->
		    <table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<form name="frmhistory" onsubmit="return false;">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
				    <td>
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td>
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryMemo('','<%= giftorderserial %>')"><b>MEMO</b></a>
		    				    	[<b><%= giftorderserial %></b>]
		    				    	|
		    				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryCS('','<%= giftorderserial %>')"><b>CS ó����</b></a>
		                            |
		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryMileage('<%= oGiftOrder.FOneItem.FUserID %>','')"><b>���ϸ���</b></a>
		                            [<b>
										<% if C_CriticInfoUserLV1 then %>
											<%= oGiftOrder.FOneItem.FUserID %>
										<% else %>
											<%= printUserId(oGiftOrder.FOneItem.FUserID, 2, "*") %>
										<% end if %>
		                            </b>]
		                            |
		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryCoupon('<%= oGiftOrder.FOneItem.FUserID %>','')"><b>����</b></a>
		                            |
		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryQna('<%=oGiftOrder.FOneItem.FUserID%>','<%If oGiftOrder.FOneItem.FUserID = "" Then response.write orderserial End If %>')"><b>1:1���</b></a>
		    				    </td>
		    				    <td width="100" align="right">
		    				    	<input type="button" class="button" name="writebutton" value="History���" class="csbutton" onclick="OpenHistoryWindow('<%= oGiftOrder.FOneItem.FUserID %>','<%= giftorderserial %>');">
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				</form>
				<tr>
				    <td>
				        <iframe name="history" src="blank.asp" border=0 frameSpacing=0 frameborder="no" width="100%" height="100%" leftmargin="0"></iframe>
	`			    </td>
				</tr>
			</table>
			<!-- �ֹ��� History-->
		</td>
	</tr>
	</table>
	<br>
<% else %>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr height="50">
        <td align="center"> [ �󼼳����� ���÷��� �ֹ���ȣ�� ���� �ϼ��� ]</td>
    </tr>
	</table>
<% end if %>

<% if (giftorderserial <> "") then %>
	<script type="text/javascript">
	    GotoHistoryCS('','<%= giftorderserial %>');
	</script>
<% end if %>

<%
set oGiftOrder = Nothing
set oaslist = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
