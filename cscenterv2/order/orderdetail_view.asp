<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ΰŽ� ������ ��ǰ �ֹ�����
' Hieditor : 2015.05.27 �̻� ����
'			 2017.07.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog_ACA.asp" -->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->
<%
dim orderserial, totalascount, ix
	orderserial = RequestCheckVar(request("orderserial"),11)

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

dim oaslist
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
	var popwin = window.showModalDialog("order_buyer_info.asp?orderserial=" + v,"order_buyer_info","resizable:no; scroll:no; dialogWidth:250px; dialogHeight:270px");
	popwin.focus();
}

// ============================================================================
// CS��ϰ���

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

                document.history.location.href = "/cscenterv2/history/history_memo.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}

function GotoHistoryCS(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("cs");

                document.history.location.href = "/cscenterv2/history/history_cs.asp?userid=" + userid + "&orderserial=" + orderserial;
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
	var popwin = window.open("/cscenterv2/history/history_memo_write.asp?userid=" + userid + "&orderserial=" + orderserial + "&sitename=diyitem&backwindow=" + "opener.document.history","GotoHistoryMemoWrite","width=600 height=600 scrollbars=yes resizable=no");
	popwin.focus();
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
function receiptinicis(tid){
	var receiptUrl = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?" +
		"noTid=" + tid + "&noMethod=1";
	var popwin = window.open(receiptUrl,"INIreceipt","width=415,height=600");
	popwin.focus();
}

// �ſ�ī�� ������ǥ �˾�_KCP
function receiptkcp(tid){
	var receiptUrl = "https://<%=chkIIF((application("Svr_Info")="Dev"),"dev","")%>admin.kcp.co.kr/Modules/Sale/CARD/ADSA_CARD_BILL_Receipt.jsp?" +"c_trade_no=" + tid + "&mnu_no=AA000001";
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
    alert('������. - ������ ���� ���');
	//var receiptUrl = "http://www.bankpay.or.kr/pgmember/customcashreceipt.jsp?bill_key1=" + tid;
	//var popwin = window.open(receiptUrl,"BankPayreceipt","width=400,height=560");
	//popwin.focus();
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
	var receiptUrl = "/cscenterv2/taxsheet/popFnCashReceipt.asp?orderserial=" + iorderserial;
	var popwin = window.open(receiptUrl,"FnCashReceipt","width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function popMileageRequest(userid, orderserial, mileage, jukyo) {
	// �ʼ� : ���̵�
	// �ɼ� : �ֹ���ȣ, ���ϸ���, ���䳻��

	if (userid == "") {
		alert("���̵� �����ϴ�.");
		return;
	}

    var popwin = window.open('/cscenterv2/mileage/pop_mileage_request.asp?userid=' + userid + '&orderserial=' + orderserial + '&mileage=' + mileage + '&jukyo=' + jukyo,'popMileageRequest','width=660,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//������� �����.
function popDacomCyberPayEdit(iorderserial){
    var popUrl = "/cscenter/cyberAcct/popCyberAcctChange.asp?orderserial=" + iorderserial;
	var popwin = window.open(popUrl,"DcCyberAcct","width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>

<% if (orderserial<>"") then %>
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="FFFFFF">
	<tr height="25">
		<td align="left">
		    &nbsp;&nbsp;
			<input type="button" class="button" value="��ü���" class="csbutton" style="width:60px;" onclick="javascript:PopOpenCancelOrder('<%= orderserial %>');">
			&nbsp;
			<input type="button" class="button" value="�κ����" class="csbutton" style="width:60px;" onclick="javascript:PopOpenCancelItem('<%= orderserial %>');">
			&nbsp;&nbsp;|&nbsp;&nbsp;
			<input type="button" class="button" value="ȸ����û" class="csbutton" style="width:70px;" onclick="javascript:PopOpenReceiveItemByTenTen('<%= orderserial %>');">
			&nbsp;
			<input type="button" class="button" value="��ǰ����" class="csbutton" style="width:70px;" onclick="javascript:PopOpenReceiveItemByUpche('<%= orderserial %>');">
			&nbsp;
			<input type="button" class="button" value="�±�ȯ" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemChange('<%= orderserial %>');">
			&nbsp;
			<input type="button" class="button" value="������߼�" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemOmit('<%= orderserial %>');">
			&nbsp;
			<input type="button" class="button" value="���񽺹߼�" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemMore('<%= orderserial %>');">
	        &nbsp;&nbsp;|&nbsp;&nbsp;
			<input type="button" class="button" value="�������ǻ���" class="csbutton" style="width:90px;" onclick="javascript:PopOpenReadMe('<%= orderserial %>');">
			<!--
			&nbsp;|&nbsp;
			<input type="button" class="button" value="�ſ�ī�����" class="csbutton" style="width:90px;" onclick="javascript:PopOpenCancelCard('<%= orderserial %>');">
			-->
			<input type="button" class="button" value="ȯ������" class="csbutton" style="width:90px;" onclick="javascript:PopCSActionCom('','<%= orderserial %>','regcsas','A003','');">
			&nbsp;
			<input type="button" class="button" value="���ϸ�������" class="csbutton" style="width:90px;" onclick="javascript:popMileageRequest('<%= ojumun.FOneItem.FUserID %>','<%= orderserial %>',0,'');">
			<!--
			<input type="button" class="button" value="�ܺθ�ȯ�ҿ�û" class="csbutton" style="width:90px;" onclick="javascript:PopOpenCancelOtherSite('<%= orderserial %>');">
			-->
	    </td>
	    <td align="right">
			<!--<input type="button" class="button" value="��Ÿ���׵��" class="csbutton" style="width:90px;" onclick="javascript:PopOpenEtcNote('<%= orderserial %>');">-->
			<!--
			<input type="button" class="button" value="�ֹ�������߼�" class="csbutton" style="width:90px;" onclick="javascript:PopCSMailSendOrder('<%= orderserial %>');">
			&nbsp;
			-->
			<input type="button" class="button" value="�����������" style="width:90px;" onclick="javascript:popOrderReceipt('<%= orderserial %>');">

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
						    	[<b><%= orderserial %></b>]
						    	&nbsp;
						    	<input type="button" class="button" value="����CS <%= totalascount %>��" class="csbutton" style="width:90px;" onclick="javascript:Cscenter_Action_List('<%= orderserial %>','','');">
	    				    </td>
	    				    <td align="right">
	    				    	<input type="button" class="button" value="������ǰ����" class="csbutton" style="width:90px;" onclick="misendmaster('<%= orderserial %>');">
	    				    </td>
	    				</tr>
	    			</table>
	    		</td>
			</tr>
			<tr height="345" bgcolor="#FFFFFF">
			    <td valign="top">
					<table height="320" width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
					<tr height="100%">
						<td colspan="12">
						    <iframe name="orderdetail" src="orderdetail_item_list.asp?orderserial=<%= orderserial %>" border=0 frameSpacing=0 frameborder="no" width="100%" height="100%" leftmargin="0"></iframe>
						</td>
					<tr>
					</table>
			    </td>
			</tr>
			</table>
			<!-- ���Ż�ǰ���� -->
		    <p>
			<!-- �ϴܺκ� -->
	        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        	<tr valign="top" height="210">
        		<td colspan="3">
        		    <!-- �ֹ��� History -->
        		    <table width="100%" height="210" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    				<form name="frmhistory" onsubmit="return false;">
    				<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
    				    <td colspan="10">
    				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    				    		<tr>
    				    			<td>
    				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryMemo('','<%= orderserial %>')"><b>MEMO</b></a>
    		    				    	[<b><%= orderserial %></b>]
    		    				    	|
    		    				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryCS('','<%= orderserial %>')"><b>CS ó����</b></a>
    		                            |
    		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryMileage('<%= ojumun.FOneItem.FUserID %>','')"><b>���ϸ���</b></a>
    		                            [<b>
											<% if (session("ssAdminCLsn") >= 500) then %>
												(<%= ojumun.FOneItem.FUserID %>)
											<% else %>
												(<%= printUserId(ojumun.FOneItem.FUserID, 2, "*") %>)
											<% end if %>
    		                            </b>]
    		                            |
    		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryCoupon('<%= ojumun.FOneItem.FUserID %>','')"><b>����</b></a>
    		                            |
    		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryQna('<%=ojumun.FOneItem.FUserID%>','<%If ojumun.FOneItem.FUserID = "" Then response.write orderserial End If %>')"><b>1:1���</b></a>
    		    				    </td>
    		    				    <td width="100" align="right">
    		    				    	<input type="button" class="button" name="writebutton" value="History���" class="csbutton" onclick="OpenHistoryWindow('<%= ojumun.FOneItem.FUserID %>','<%= orderserial %>');">
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
            	<td width="5"></td>
        		<td width="225">
        		    <!-- �ֹ����� -->
        		    <table width="225" height="210" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    				    <td colspan="10">
    				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    				    		<tr>
    				    			<td width="100">
    				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ� ����</b>
    		    				    </td>
    		    				    <td align="right">
    		    				    	<input type="button" class="button" value="������������" class="csbutton" onclick="javascript:PopNextIpkumDiv('<%= orderserial %>');">
    		    				    </td>
    		    				</tr>
    		    			</table>
    		    		</td>
    				</tr>
    				<!--
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">�ֹ���ȣ</td>
    				    <td bgcolor="#FFFFFF"><%= orderserial %></td>
    				</tr>
    				-->
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">�������</td>
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
	    				<tr>
	    				    <td bgcolor="<%= adminColor("topbar") %>">
	    				    <% if ojumun.FOneItem.IsDacomCyberAccountPay then %>
								<a href="javascript:popDacomCyberPayEdit('<%= orderserial %>')"><font color="red">����</font></a>
	    				    <% else %>
							    <a href="javascript:popDacomCyberPayEdit('<%= orderserial %>')">�Ϲ�</a>
	    				    <% end if %>
	    				    </td>
	    				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FAccountNo %></td>
	    				</tr>
    				<% end if %>

    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">�ֹ��Ͻ�</td>
    				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FRegDate %></td>
    				</tr>
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">�Ա�Ȯ��</td>
    				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FIpkumDate %></td>
    				</tr>
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">�ֹ��뺸</td>
    				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.Fbaljudate %></td>
    				</tr>
    				<!--
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">����Ͻ�</td>
    				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.Fbeadaldate %></td>
    				</tr>
    				-->
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">���ι�ȣ</td>
    				    <td bgcolor="#FFFFFF">
    				        <input type="text" class="text_ro" value="<%= ojumun.FOneItem.FAuthcode %>" readonly size="20">
    				    </td>
    				</tr>
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">PG�� ID</td>
    				    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" value="<%= ojumun.FOneItem.FPaygatetID %>" readonly></td>
    				</tr>
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">��������</td>
    				    <td bgcolor="#FFFFFF">
    				    	<!-- All@ ������ ��� -->
    	                  	<% if (trim(ojumun.FOneItem.Faccountdiv)="80") and (ojumun.FOneItem.FIpkumDiv >= 4) then %>
    	                  	    <input type="button" class="button" value="�ſ�" onclick="javascript:receiptallat('<%= ojumun.FOneItem.Fpaygatetid %>');">
    	                  	<% end if %>

    				        <!-- �ſ�ī�� ������ǥ -->
    	                  	<% if (ojumun.FOneItem.FAccountDiv="100") and (ojumun.FOneItem.FIpkumDiv >= 4) then %>
    	                  		<% if ojumun.FOneItem.Fpaygatetid<>"" then %>
    		                  		<% if (Left(ojumun.FOneItem.Fpaygatetid,9)="IniTechPG") then %>
    		                  		    <input type="button" class="button" value="�ſ�" onclick="javascript:receiptinicis('<%= ojumun.FOneItem.FPaygatetID %>');">
    		                  		<% else %>
    		                  		    <input type="button" class="button" value="�ſ�" onclick="javascript:receiptkcp('<%= ojumun.FOneItem.FPaygatetID %>')">
    		                  		<% end if %>
    		                  	<% end if %>
    		                <% end if %>

    		                <!-- ���ݿ����� ����Ȯ�� �ǽð���ü-->
    						<% if (ojumun.FOneItem.FAccountDiv="20") and (ojumun.FOneItem.FIpkumDiv >= 4) then %>
    						    <% if (Left(ojumun.FOneItem.Fpaygatetid,9)="IniTechPG") then %>
    						        <% if ojumun.FOneItem.FAuthCode<>"" then %> <!-- ������ ���ݿ����� ��û�� ��� -->

    						            <input type="button" class="button" value="����" onclick="javascript:receiptinicis('<%= ojumun.FOneItem.Fpaygatetid %>')">

    						            <!-- input type="button" class="button" value="����2" onclick="javascript:cashreceipt('<%= ojumun.FOneItem.ForderSerial %>')" -->
    						        <% elseif (ojumun.FOneItem.FcashreceiptReq="T") then %>
    						        <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">����</a>
    						        <% else %>
                                    <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">���º���</a>
    						        <% end if %>
    						    <% else %>
    						        <% if (ojumun.FOneItem.Fcashreceiptreq<>"") then %>
    						        <input type="button" class="button" value="����" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
    						        <% else %>
    						        <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">���º���</a>
    						        <% end if %>
    						        <!-- input type="button" class="button" value="����" onclick="javascript:receiptbankpay('<%= ojumun.FOneItem.Fpaygatetid %>')"-->
    						    <% end if %>
    						<% end if %>

    						<!-- ���ݿ����� ����Ȯ�� ������ ��� -->
    						<% if (ojumun.FOneItem.FAccountDiv="7") then %>
								<% if (ojumun.FOneItem.Fauthcode<>"") then %>
                                    <input type="button" class="button" value="����" onclick="javascript:cashreceipt('<%= ojumun.FOneItem.ForderSerial %>')">
                                <% elseif (ojumun.FOneItem.FcashreceiptReq="R") then %>
                                <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">���ݿ����� �����û ����</a>
                                <% elseif (ojumun.FOneItem.FcashreceiptReq="T") then %>
    						        <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">����</a>
                                <% else %>
                                    <% if (ojumun.FOneItem.FIpkumdiv>3) then %>
                                    <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">���º���</a>
                                    <% else %>
                                    <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">���º���</a>
                                    <% end if %>
                                <% end if %>
                            <% end if %>
                            <!-- ���ݰ�꼭 �����û ���� -->

    						<!-- ���ں������� -->
                    		<% if (ojumun.FOneItem.FInsureCd = "0") then %>
                    		    <input type="button" class="button" value="����" onclick="javascript:insurePrint('<%= ojumun.FOneItem.ForderSerial %>','ZZcube1010')">
    						<% end if %>
    				    </td>
    				</tr>
        			</table>
        			<!-- �ֹ����� -->
        		</td>
        	</tr>
	        </table>
		</td>
		<td width="5"></td>
		<td width="250" align="right">
			<!-- ���������� -->
			<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frmbuyerinfo" onsubmit="return false;">
			<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			    <td colspan="2">
			    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			    		<tr>
			    			<td width="100">
			    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ ����</b>
	    				    </td>
	    				    <td align="right">
	    				    	<input type="button" class="button" value="��������������" class="csbutton" onclick="javascript:PopBuyerInfo('<%= orderserial %>');">
	    				    </td>
	    				</tr>
	    			</table>
	    		</td>
			</tr>
			<tr height="24">
			    <td bgcolor="<%= adminColor("topbar") %>">������ID</td>
			    <td bgcolor="#FFFFFF">
					<% if (session("ssAdminCLsn") >= 500) then %>
						(<%= ojumun.FOneItem.FUserID %>)
					<% else %>
						(<%= printUserId(ojumun.FOneItem.FUserID, 2, "*") %>)
					<% end if %>
			    	&nbsp;[<font color="<%= ojumun.FOneItem.GetUserLevelColor %>"><%= ojumun.FOneItem.GetUserLevelName %></font>]
			    </td>
			</tr>
			<tr height="23">
			    <td bgcolor="<%= adminColor("topbar") %>">�ֹ���ȣ</td>
			    <td bgcolor="#FFFFFF"><%= orderserial %></td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">�����ڸ�</td>
			    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="buyname" value="<%= ojumun.FOneItem.FBuyName %>" size="8" readonly></td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
			    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="buyphone" value="<%= ojumun.FOneItem.FBuyPhone %>" readonly></td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
			    <td bgcolor="#FFFFFF">
			        <input type="text" class="text_ro" name="buyhp" value="<%= ojumun.FOneItem.FBuyHp %>" readonly>
			        <input type="button" name="buyhp" class="button" value="SMS" onclick="PopCSSMSSendNew({reqhp:'<%= ojumun.FOneItem.FBuyHp %>', orderserial:'<%= ojumun.FOneItem.Forderserial %>', userid:'<%= ojumun.FOneItem.Fuserid %>'});">
			    </td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">�̸���</td>
			    <td bgcolor="#FFFFFF">
			        <input type="text" class="text_ro" name="buyemail" value="<%= ojumun.FOneItem.FBuyEmail %>" size="20" readonly>
			        <input type="button" name="email" class="button" value="mail" onclick="javascript:PopCSMailSend('<%= ojumun.FOneItem.FBuyEmail %>','<%= ojumun.FOneItem.Forderserial %>','<%= ojumun.FOneItem.Fuserid %>');">
			    </td>
			</tr>

			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">�Ա��ڸ�</td>
			    <td bgcolor="#FFFFFF">
			        <input type="text" class="text_ro" name="accountname" value="<%= ojumun.FOneItem.FAccountName %>" size="14" readonly>
			        <input type="button" class="button" value="�˻�" class="csbutton" onclick="FindByIpkumname()">
			        <acronym title="<%= ojumun.FOneItem.Faccountno %>"><%= left(ojumun.FOneItem.Faccountno,2) %></acronym>
			    </td>
			</tr>
			</form>
			</table>
			<!-- ���������� -->
	        <br>
			<!-- ������� -->
			<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frmreqinfo" onsubmit="return false;">
			<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			    <td colspan="2">
			    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			    		<tr>
			    			<td width="100">
			    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��� ����</b>
	    				    </td>
	    				    <td align="right">
	    				    	<input type="button" class="button" value="�������������" class="csbutton" onclick="javascript:PopReceiverInfo('<%= orderserial %>');">
	    				    </td>
	    				</tr>
	    			</table>
	    		</td>
			</tr>
			<tr>
			    <td width="100" bgcolor="<%= adminColor("topbar") %>">�����θ�</td>
			    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqName %>" readonly></td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
			    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqPhone %>" readonly></td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
			    <td bgcolor="#FFFFFF">
			        <input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqHp %>" readonly>
			        <input type="button" name="reqhp" class="button" value="SMS" onclick="PopCSSMSSendNew({reqhp:'<%= ojumun.FOneItem.FReqHp %>', orderserial:'<%= ojumun.FOneItem.Forderserial %>', userid:'<%= ojumun.FOneItem.Fuserid %>'});">
			    </td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">����ּ�</td>
			    <td bgcolor="#FFFFFF">
			        <input type="text" class="text_ro" name="txzip1" value="<%= ojumun.FOneItem.FReqZipCode %>" size="7" readonly>
			        <input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqZipAddr %>" size="18" readonly><br>
			        <textarea class="textarea_ro" rows="2" cols="28" readonly><%= ojumun.FOneItem.FReqAddress %></textarea>
                </td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">��Ÿ����</td>
			    <td bgcolor="#FFFFFF">
			        <textarea class="textarea_ro" rows="2" cols="28" readonly><%= ojumun.FOneItem.FComment %></textarea>
			    </td>
			</tr>
			</form>
			</table>
			<!-- ������� -->
			<br>
			<!-- �ؿܹ���� ��� �ؿܹ�� ���� �ƴҰ��, �ö���ֹ����� -->

			<% if ojumun.FOneItem.IsForeignDeliver=true then %>
				<!-- �ؿܹ�� ���� -->
				<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="2">
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
				<tr>
				    <td width="55" bgcolor="<%= adminColor("topbar") %>">��ǰ�߷�</td>
				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FitemWeigth %>(g)</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">�ڽ��߷�</td>
				    <td bgcolor="#FFFFFF">200(g)</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">��۱�����</td>
				    <td bgcolor="#FFFFFF">
				    	<%= ojumun.FOneItem.FcountryNameEn %>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">��۱���</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" value="<%= ojumun.FOneItem.FDlvcountryCode %>" size="2" readonly>
				    	<input type="text" class="text_ro" value="<%= ojumun.FOneItem.FemsAreaCode %>" size="2" readonly>
						<input type="button" class="button" value="���ǥ����" class="csbutton" style="width:100px;" onclick="popForeignDeliverPay('<%= ojumun.FOneItem.FemsAreaCode %>');">
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">EMS���</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" style="text-align:right;" value="<%= FormatNumber(ojumun.FOneItem.FemsDlvCost,0) %>" size="8" maxlength="10" readonly>��
				    </td>
				</tr>
				<%'If ojumun.FOneItem.FemsInsureYn = "Y" Then %>
					<tr>
					    <td bgcolor="<%= adminColor("topbar") %>">���谡��(<%=ojumun.FOneItem.FemsInsureYn%>)</td>
					    <td bgcolor="#FFFFFF">
					    	<input type="text" class="text_ro" style="text-align:right;" value="<%=FormatNumber(ojumun.FOneItem.FemsInsurePrice,0)%>" size="8" maxlength="10" readonly>��
					    </td>
					</tr>
				<%'End If %>
				</table>
			<% else %>
				<!-- �ö�� �ֹ�  -->
				<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="2">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td width="100">
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ö������</b>
		    				    </td>
		    				    <td align="right">
		    				    	<input type="button" class="button" value="�ö���޼�������" class="csbutton" onclick="javascript:PopFlowerDeliverInfo('<%= orderserial %>');">
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">����</td>
				    <td bgcolor="#FFFFFF">
				        <input type="radio" name="cardribbon" value="1" <% if ojumun.FOneItem.Fcardribbon="1" then response.write "checked" %> >ī��
				        <input type="radio" name="cardribbon" value="2" <% if ojumun.FOneItem.Fcardribbon="2" then response.write "checked" %> >����
				        <input type="radio" name="cardribbon" value="3" <% if ojumun.FOneItem.Fcardribbon="3" then response.write "checked" %> >����
				    </td>
				</tr>
				<tr>
				    <td colspan="2" bgcolor="#FFFFFF">
				        <textarea class="textarea_ro" name="message" rows="3" cols="37" readonly><%= ojumun.FOneItem.Fmessage %></textarea><br>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">FROM</td>
				    <td bgcolor="#FFFFFF">
				        <input type="text" class="text_ro" name="fromname" value="<%= ojumun.FOneItem.Ffromname %>" size="20" maxlength="20" readonly>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">�����</td>
				    <td bgcolor="#FFFFFF">
				        <input type="text" class="text_ro" name="reqdate" value="<%= ojumun.FOneItem.Freqdate %>" size="10" readonly>��
				        <input type="text" class="text_ro" name="reqtime" value="<%= ojumun.FOneItem.GetReqTimeText %>" size="10" readonly>
				    </td>
				</tr>
				</table>
			<% end if %>
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
	<script language='javascript'>
	    GotoHistoryCS('','<%= orderserial %>');
	</script>
<% end if %>

<%
set ojumun = Nothing
set oaslist = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
