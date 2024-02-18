
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
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%
dim orderserial, oaslist, totalascount, ix
	orderserial = RequestCheckVar(request("orderserial"),32)

totalascount = 0

dim ojumun
set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster_3PL
end if

if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster_3PL
end if

dim ojumunitemsummary
set ojumunitemsummary = new COrderMaster
	ojumunitemsummary.FRectOldOrder = ojumun.FRectOldOrder
	ojumunitemsummary.FRectOrderSerial = orderserial
	ojumunitemsummary.getOrderItemSummary

set oaslist = new CCSASList
	if (orderserial <> "") then
	    oaslist.FRectOrderSerial = orderserial
	    oaslist.GetCSASTotalCount_3PL

	    totalascount = oaslist.FResultCount
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
	// var popwin = window.open("orderdetail_editorder.asp?orderserial=" + orderserial,"PopOpenModifyOrder","width=1200 height=800 scrollbars=yes resizable=yes");
	var popwin = window.open("orderdetail_simple_editorder.asp?orderserial=" + orderserial,"PopOpenModifyOrder","width=1200 height=800 scrollbars=yes resizable=yes");
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
	return;
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("memo");

                document.history.location.href = "/cscenter/history/history_memo.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}

function GotoHistoryCS(userid, orderserial) {
	return;
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

</script>

<% if (orderserial<>"") then %>
	<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="FFFFFF">
		<tr height="25">
			<td align="left">
			    &nbsp;
				<input type="button" class="button" value="��ü���" class="csbutton" style="width:60px;" onclick="javascript:PopOpenCancelOrder('<%= orderserial %>');" disabled>
				&nbsp;
				<input type="button" class="button" value="�κ����" class="csbutton" style="width:60px;" onclick="javascript:PopOpenCancelItem('<%= orderserial %>');" disabled>
				&nbsp;|&nbsp;
				<input type="button" class="button" value="��ǰ����" class="csbutton" style="width:70px;" onclick="javascript:PopOpenReceiveItemByUpche('<%= orderserial %>');" disabled>
				&nbsp;
				<input type="button" class="button" value="��ȯ���" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemChange('<%= orderserial %>');" disabled>
				&nbsp;
				<input type="button" class="button" value="������߼�" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemOmit('<%= orderserial %>');" disabled>
				&nbsp;
				<input type="button" class="button" value="���񽺹߼�" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemMore('<%= orderserial %>');" disabled>
		        &nbsp;|&nbsp;
				<input type="button" class="button" value="��Ÿȸ��" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceRecvItemMore('<%= orderserial %>');" disabled>
		        &nbsp;|&nbsp;
				<input type="button" class="button" value="�������ǻ���" class="csbutton" style="width:90px;" onclick="javascript:PopOpenReadMe('<%= orderserial %>');" disabled>
		    </td>
		    <td align="right">

			</td>
		</tr>
		</table>
	<% end if %>

	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
	<tr>
		<td width="500" align="left" valign="top">
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
		    				    </td>
		    				    <td align="right">
									<% if C_CriticInfoUserLV1 then %>
		    				    	<input type="button" class="button" value="��������������" class="csbutton" onclick="javascript:PopBuyerInfo('<%= orderserial %>');" style="width:120px;" disabled>
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
		    				    </td>
		    				    <td align="right">
									<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
		    				    	<input type="button" class="button" value="�������������" class="csbutton" onclick="javascript:PopReceiverInfo('<%= orderserial %>');" style="width:120px;" disabled>
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
							    	<input type="button" class="button" value="����CS <%= totalascount %>��" class="csbutton" style="width:90px;" onclick="javascript:Cscenter_Action_List_3PL('<%= orderserial %>','','');">
		    				    </td>
		    				    <td align="right"  width="200">
		    				    	<input type="button" class="button" value="������ǰ����" class="csbutton" style="width:90px;" onclick="misendmaster('<%= orderserial %>');" disabled>
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
	                    	        <iframe name="orderdetail" src="orderitemmaster_3PL.asp?orderserial=<%= orderserial %>" border="0" frameborder="no" frameSpacing=0  width="100%" height="100%" leftmargin="0"></iframe>
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
									<% end if %>
		    				    </td>
		    				    <td width="100" align="right">
									<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
		    				    	<input type="button" class="button" name="writebutton" value="History���" class="csbutton" onclick="OpenHistoryWindow('<%= ojumun.FOneItem.FUserID %>','<%= orderserial %>');" disabled>
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
				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.GetPggubunName %></td>
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
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->