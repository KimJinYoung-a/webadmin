<%
function getCardDEtailInfo(pggubun, cardcodeALL)
    dim ret : ret = ""
    getCardDEtailInfo = ""

    if isNULL(cardcodeALL) then Exit Function

    Dim acardCode , bcardcode, cinstallment

    acardCode       = Left(cardcodeALL,2)
    ''cinstallment    = Right(cardcodeALL,2) ''14|26|00 ==> 14|26|00|1 ''������ �ڵ� �κ���� ���ɿ��� (2011-08-25)-----------
    if (LEN(cardcodeALL)=10) then
        cinstallment = Mid(cardcodeALL,7,2)
    else
        cinstallment = Right(cardcodeALL,2)
    end if
    ''------------------------------------------------------------------------------------------------------------------------

    SELECT CASE acardCode
        CASE "11" : ret = "BC"
        CASE "06" : ret = "����"
        CASE "12" : ret = "�Ｚ"
        CASE "14" : ret = "����"
        CASE "01" : ret = "��ȯ"
        CASE "04" : ret = "����"
        CASE "03" : ret = "�Ե�"
        CASE "16" : ret = "NH"
        CASE "17" : ret = "�ϳ�SK"

        CASE ELSE :
    END SELECT

	if (pggubun = "KA") then
		'// īī��PAY
		SELECT CASE acardCode
			CASE "01" :
				ret = "īī��-��"
			CASE "02" :
				ret = "īī��-����"
			CASE "03" :
				ret = "īī��-��ȯ"
			CASE "04" :
				ret = "īī��-�Ｚ"
			CASE "05" :
				ret = "īī��-����"
			CASE "06" :
				ret = "īī��-����"
			CASE "07" :
				ret = "īī��-����"
			CASE "08" :
				ret = "īī��-�Ե�"
			CASE "11" :
				ret = "īī��-��Ƽ"
			CASE "12" :
				ret = "īī��-NHä��"
			CASE "13" :
				ret = "īī��-����"
			CASE "15" :
				ret = "īī��-�츮"
			CASE "16" :
				ret = "īī��-�ϳ�SK"
			CASE "18" :
				ret = "īī��-����"
			CASE "19" :
				ret = "īī��-����"
			CASE "21" :
				ret = "īī��-����"
			CASE "22" :
				ret = "īī��-����"
			CASE "23" :
				ret = "īī��-����"
			CASE "25" :
				ret = "īī��-����"
			CASE "26" :
				ret = "īī��-������"
			CASE "27" :
				ret = "īī��-���̳ʽ�"
			CASE "28" :
				ret = "īī��-AMX"
			CASE "29" :
				ret = "īī��-JCB"
			CASE "30" :
				ret = "īī��-��Ŀ��"
			CASE "34" :
				ret = "īī��-����"
			CASE ELSE :
				ret = "īī��-???"
		END SELECT
	end if

    if (cinstallment="00") then ret = ret + " �Ͻú�"
    if (cinstallment<>"00") and (cinstallment<>"") then ret = ret + " " + cinstallment + "����"
    if (ret<>"") then ret = "(" + ret + ")"
    getCardDEtailInfo = ret
end function

function getPhoneDetailInfo(payMethod, ipkumdate)
    dim ipkumMonth, currMonth
	dim result : result = ""

	ipkumMonth = Left(ipkumdate, 7)
	currMonth = Left(now(), 7)

	if (payMethod = "400") then
		'�ڵ��� ����
		if (ipkumMonth = currMonth) then
			result = "<font color=blue>�ڵ��� ���� ��Ұ���</font>"
		else
			result = "<font color=red>������� �Ұ�(�������� ��Ұ���)</font>"
		end if
	end if

	getPhoneDetailInfo = result
end function

%>

<% if (IsDisplayRefundInfo) then %>

	<%
	' ��ġ�� ������ ȯ�Ұ��� �ֹ���ȣ�� ���� ������ �־ ����ó��		'(not(IsNumeric(orderserial)) and orefund.FOneItem.Freturnmethod="R007")	' 2018.12.04 �ѿ��
	if (IsCSRefundNeeded(divcd, OrderMasterState) or (IsChangeOrder and IsCSReturnProcess(divcd))) or (orderserial = exceptOrderserial) or (not(IsNumeric(orderserial)) and orefund.FOneItem.Freturnmethod="R007") then
	%>

        <tr bgcolor="#FFFFFF">
            <td width="100" height="30">��������</td>
            <td width="600">
            	<b>
            	<% if (mainpaymentorg<>oordermaster.FOneItem.Fsubtotalprice-realSubPaymentSum) then %>
            	    ���ʰ����ݾ� : <%= mainpaymentorg %>
            	    <br>
            	<% end if %>

            	<% if oordermaster.FOneItem.IsErrSubtotalPrice then %>
            		<font color="red"><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice-realSubPaymentSum,0) %>��</font>
            	<% else %>
            		<%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice-realSubPaymentSum,0) %>��
				<% end if %>
				<% if (prevrefundsum > 0) then %>
				    <% if (not IsTicketOrder) then %>
    					<% if (oordermaster.FOneItem.FCancelyn = "Y") and ((prevrefundsum - oordermaster.FOneItem.Fsubtotalprice - csbeasongpaysum) <> 0) then %>
    						(ȯ�� <%= FormatNumber((prevrefundsum - oordermaster.FOneItem.Fsubtotalprice - csbeasongpaysum), 0) %>�� ���� )
    					<% elseif (oordermaster.FOneItem.FCancelyn <> "Y") then %>
    						(ȯ�� <%= FormatNumber(prevrefundsum - csbeasongpaysum, 0) %>�� ����)
    					<% end if %>
    				<% end if %>
				<% end if %>
				<% if (csbeasongpaysum > 0) then %>
					��ۺ�ȯ�� : <%= FormatNumber(csbeasongpaysum, 0) %>��
				<% end if %>
            	&nbsp;
                [<%= oordermaster.FOneItem.JumunMethodName %> <%= getCardDEtailInfo(oordermaster.FOneItem.Fpggubun, cardcodeall) %>]
                [<font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font>]
                [<font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %></font>]
                <% if iPgGubun="NP" then %><font color=red>���̹�����</font><% end if %>
                <% if (realdepositsum>0) then %>
                   /&nbsp; <%= FormatNumber(realdepositsum,0) %>��&nbsp; [��ġ��]
                <% end if %>
                <% if (realgiftcardsum>0) then %>
                   /&nbsp; <%= FormatNumber(realgiftcardsum,0) %>��&nbsp; [Giftī��]
                <% end if %>

				<% if (oordermaster.FOneItem.Faccountdiv="110") then %>
                	(OK Cashbag��� : <strong><%= FormatNumber(oordermaster.FOneItem.FokcashbagSpend,0) %></strong> ��)
                <% end if %>
                </b>
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td width="100" height="30">ȯ�ҹ��</td>
            <td width="600">
            	<%
            	'// drawSelectBoxCancelTypeBox �� /lib/classes/cscenter/cs_aslistcls.asp ����
            	%>
                <% call drawSelectBoxCancelTypeBox("returnmethod",orefund.FOneItem.Freturnmethod,oordermaster.FOneItem.Faccountdiv,divcd,"onChange='ChangeReturnMethod(this);'") %>
                <% if (Not IsStatusRegister) then %>
                (<%= orefund.FOneItem.FreturnmethodName %>)
                <% end if %>
                <input name="RefundRecalcuButton" class="csbutton" type="button" value="����" onClick="CalculateAndApplyItemCostSum(frmaction);">
                <% if (oordermaster.FOneItem.Faccountdiv = "100") or (oordermaster.FOneItem.Faccountdiv = "110") then %>
                	<% if (cardPartialCancelok = "Y") then %>
                		<font color="blue">�ſ�ī�� �κ���� ����ī��</font>
                	<% else %>
						<%= cardcancelerrormsg %>
                	<% end if %>
				<% elseif (oordermaster.FOneItem.Faccountdiv = "400") then %>
					<%= getPhoneDetailInfo(oordermaster.FOneItem.Faccountdiv, oordermaster.FOneItem.Fipkumdate) %>
				<% end if %>

				<input type="hidden" name="paygateTid" value="<%= oordermaster.FOneItem.Fpaygatetid %>">
            </td>
        </tr>
        <tr  bgcolor="FFFFFF" id="refundinfo_R007" <% if orefund.FOneItem.Freturnmethod="R007" then response.write "" else response.write "style=""display:none""" %> >
            <td width="100" height="30">��������</td>
            <td align="left">
                <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="BABABA">
	            	<tr bgcolor="FFFFFF">
	            		<td width="80">���¹�ȣ</td>
	            		<td>
	            		    <input class="text" type="text" size="20" name="rebankaccount" value="<%= orefund.FOneItem.Frebankaccount %>" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %> >
	            		    <input class="csbutton" type="button" value="��������(<%= prevrefundhistorycnt %>)" onClick="popPreReturnAcct('<%= oordermaster.FOneItem.Fuserid %>','frmaction','rebankaccount','rebankownername','rebankname');" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>>
                            &nbsp;
                            <input class="csbutton" type="button" value="ȯ�������Է¿�û" onClick="popRequestReturnAcctLMS('<%= id %>','<%= oordermaster.FOneItem.Forderserial %>', '<%= oordermaster.FOneItem.Fbuyhp %>');" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>>
	            		</td>
	            	</tr>
	            	<tr bgcolor="FFFFFF">
	            		<td>�����ָ�</td>
	            		<td><input class="text" type="text" size="20" name="rebankownername" value="<%= orefund.FOneItem.Frebankownername %>" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>></td>
	            	</tr>
	                <tr bgcolor="FFFFFF">
	            		<td>�ŷ�����</td>
	            		<td><% DrawBankCombo "rebankname", orefund.FOneItem.Frebankname %></td>
	            	</tr>
            	</table>
            </td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R100" <% if orefund.FOneItem.Freturnmethod="R100" then response.write "" else response.write "style=""display:none""" %>>
    		<td width="100" height="30">PG�� ID</td>
    		<td><input class="text_ro" type="text" name="paygateTid1" size="30" value="<%= oordermaster.FOneItem.Fpaygatetid %>" readonly></td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R550" <% if orefund.FOneItem.Freturnmethod="R550" then response.write "" else response.write "style=""display:none""" %>>
    		<td width="100" height="30">������ȣ</td>
    		<td><input class="text_ro" type="text" name="paygateTid2" size="30" value="<%= oordermaster.FOneItem.Fpaygatetid %>" readonly></td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R560" <% if orefund.FOneItem.Freturnmethod="R560" then response.write "" else response.write "style=""display:none""" %>>
    		<td width="100" height="30">������ȣ</td>
    		<td><input class="text_ro" type="text" name="paygateTid3" size="30" value="<%= oordermaster.FOneItem.Fpaygatetid %>" readonly></td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R050" style="display:none">
            <td colspan="2" align="left" height="30">�ܺθ� ȯ�ҿ�û</td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R900" style="display:none">
    		<td width="100" height="30">���̵�</td>
    		<td><input class="text_ro" type="text" name="refund_userid" value="<%= oordermaster.FOneItem.Fuserid %>" readonly></td>
        </tr>
        <tr bgcolor="FFFFFF">
    		<td width="100" height="30">���ʰ����ݾ�</td>
    		<td>
    		    <input class="text_ro" type="text" size="10" name="mainpaymentorg" value="<%= mainpaymentorg %>" maxlength=10 readonly>

    		    <input type=hidden name=cardcode value="<%= cardcode %>">
    		</td>
    	</tr>
        <tr bgcolor="FFFFFF">
    		<td width="100" height="30">��ȯ�ұݾ�</td>
    		<td>
    		    <input class="text_ro" type="text" size="10" name="prevrefundsum" value="<%= CHKIIF(IsStatusRegister=True, prevrefundsum, prevrefundsum - orefund.FOneItem.Frefundrequire) %>" maxlength=10 readonly>
				* ȯ�� ��������
    		</td>
    	</tr>

        <tr bgcolor="FFFFFF">
    		<td width="100" height="30">ȯ�� ������</td>
    		<% if (orefund.FResultCount>0) then %>
    		<td>
    		    <input class="text_ro" type="text" size="10" name="refundrequire" value="<%= orefund.FOneItem.Frefundrequire %>" maxlength=7 readonly>
    		    (<%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %>)
				<input type="hidden" name="refundrequire_org" value = "<%= orefund.FOneItem.Frefundrequire %>">
    		</td>
    		<% else %>
    		<td>
    			<input class="text_ro" type="text" size="10" name="refundrequire" value="<%= orefund.FOneItem.Frefundrequire %>" <% if (divcd <> "A003") then %>readonly<% end if %>>
	            <% if (divcd = "A003") and (RefundAllowLimit <> -1) then %>
	          	* <font color=red><%= FormatNumber(RefundAllowLimit,0) %> ��</font> �ʰ� ȯ�ҺҰ�
	            <% end if %>
    		</td>
    		<% end if %>
    	</tr>
    	<% IF (Not (IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""))) then %>
        <tr bgcolor="FFFFFF">
    	    <td colspan="2" height="30"><b>ȯ�� ���� �ۼ����̹Ƿ� ���� �� �� �����ϴ�.</b> [<%= orefund.FOneItem.Fupfiledate %>]</td>
    	</tr>
        <% end if %>

		<!-- ���� ȯ�������� ����, ȯ�ҿ�û�� ��� ȯ�ҿ����� �������� -->
		<% if (divcd <> "A003") then %>
    	<tr bgcolor="FFFFFF">
    	    <td colspan="2" height="30">
    	    	* ȯ�ҿ������� ������ �� �����ϴ�.<br>
    	    	* ȯ�Ҿ��� ȯ��CS�������¸� ������ �ݾ��Դϴ�.<br>
    	    	* ��ۺ�ȯ���� ��ۺ���Ҿ��� �̷���� ȯ���� �ǹ��մϴ�.
    	    </td>
    	</tr>
    	<% end if %>

		<% if (IsStatusFinishing or IsStatusFinished) then %>
	    <script language='javascript'>
	    frmaction.returnmethod.disabled=true;
	    frmaction.RefundRecalcuButton.disabled=true;
	    frmaction.rebankaccount.disabled=true;
	    frmaction.rebankname.disabled=true;
	    frmaction.rebankownername.disabled=true;
	    frmaction.refundrequire.disabled=true;
	    frmaction.paygateTid.disabled=true;
	    frmaction.refund_userid.disabled=true;

		<% if (IsStatusFinishing) then %>
	    if ((divcd=="A003")&&(frmaction.returnmethod.value=="R900")){
	        alert('���ϸ��� ȯ���� �Ϸ�ó���� �ڵ� ȯ�� �˴ϴ�.');
	    }
	    <% end if %>
	    </script>
		<% end if %>

	<% else %>
		<tr bgcolor="FFFFFF" ><td align="center" height="30">ȯ�� ���� ���°� �ƴմϴ�.</td></tr>
	<% end if %>

<% end if %>
