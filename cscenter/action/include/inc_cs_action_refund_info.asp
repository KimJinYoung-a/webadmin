<%
function getCardDEtailInfo(pggubun, cardcodeALL)
    dim ret : ret = ""
    getCardDEtailInfo = ""

    if isNULL(cardcodeALL) then Exit Function

    Dim acardCode , bcardcode, cinstallment

    acardCode       = Left(cardcodeALL,2)
    ''cinstallment    = Right(cardcodeALL,2) ''14|26|00 ==> 14|26|00|1 ''마지막 코드 부분취소 가능여부 (2011-08-25)-----------
    if (LEN(cardcodeALL)=10) then
        cinstallment = Mid(cardcodeALL,7,2)
    else
        cinstallment = Right(cardcodeALL,2)
    end if
    ''------------------------------------------------------------------------------------------------------------------------

    SELECT CASE acardCode
        CASE "11" : ret = "BC"
        CASE "06" : ret = "국민"
        CASE "12" : ret = "삼성"
        CASE "14" : ret = "신한"
        CASE "01" : ret = "외환"
        CASE "04" : ret = "현대"
        CASE "03" : ret = "롯데"
        CASE "16" : ret = "NH"
        CASE "17" : ret = "하나SK"

        CASE ELSE :
    END SELECT

	if (pggubun = "KA") then
		'// 카카오PAY
		SELECT CASE acardCode
			CASE "01" :
				ret = "카카오-비씨"
			CASE "02" :
				ret = "카카오-국민"
			CASE "03" :
				ret = "카카오-외환"
			CASE "04" :
				ret = "카카오-삼성"
			CASE "05" :
				ret = "카카오-신한"
			CASE "06" :
				ret = "카카오-신한"
			CASE "07" :
				ret = "카카오-현대"
			CASE "08" :
				ret = "카카오-롯데"
			CASE "11" :
				ret = "카카오-씨티"
			CASE "12" :
				ret = "카카오-NH채움"
			CASE "13" :
				ret = "카카오-수협"
			CASE "15" :
				ret = "카카오-우리"
			CASE "16" :
				ret = "카카오-하나SK"
			CASE "18" :
				ret = "카카오-주택"
			CASE "19" :
				ret = "카카오-조흥"
			CASE "21" :
				ret = "카카오-광주"
			CASE "22" :
				ret = "카카오-전북"
			CASE "23" :
				ret = "카카오-제주"
			CASE "25" :
				ret = "카카오-비자"
			CASE "26" :
				ret = "카카오-마스터"
			CASE "27" :
				ret = "카카오-다이너스"
			CASE "28" :
				ret = "카카오-AMX"
			CASE "29" :
				ret = "카카오-JCB"
			CASE "30" :
				ret = "카카오-디스커버"
			CASE "34" :
				ret = "카카오-은런"
			CASE ELSE :
				ret = "카카오-???"
		END SELECT
	end if

    if (cinstallment="00") then ret = ret + " 일시불"
    if (cinstallment<>"00") and (cinstallment<>"") then ret = ret + " " + cinstallment + "개월"
    if (ret<>"") then ret = "(" + ret + ")"
    getCardDEtailInfo = ret
end function

function getPhoneDetailInfo(payMethod, ipkumdate)
    dim ipkumMonth, currMonth
	dim result : result = ""

	ipkumMonth = Left(ipkumdate, 7)
	currMonth = Left(now(), 7)

	if (payMethod = "400") then
		'핸드폰 결제
		if (ipkumMonth = currMonth) then
			result = "<font color=blue>핸드폰 결제 취소가능</font>"
		else
			result = "<font color=red>결제취소 불가(결제월만 취소가능)</font>"
		end if
	end if

	getPhoneDetailInfo = result
end function

%>

<% if (IsDisplayRefundInfo) then %>

	<%
	' 예치금 무통장 환불건이 주문번호가 없이 꽂힌게 있어서 예외처리		'(not(IsNumeric(orderserial)) and orefund.FOneItem.Freturnmethod="R007")	' 2018.12.04 한용민
	if (IsCSRefundNeeded(divcd, OrderMasterState) or (IsChangeOrder and IsCSReturnProcess(divcd))) or (orderserial = exceptOrderserial) or (not(IsNumeric(orderserial)) and orefund.FOneItem.Freturnmethod="R007") then
	%>

        <tr bgcolor="#FFFFFF">
            <td width="100" height="30">결제정보</td>
            <td width="600">
            	<b>
            	<% if (mainpaymentorg<>oordermaster.FOneItem.Fsubtotalprice-realSubPaymentSum) then %>
            	    최초결제금액 : <%= mainpaymentorg %>
            	    <br>
            	<% end if %>

            	<% if oordermaster.FOneItem.IsErrSubtotalPrice then %>
            		<font color="red"><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice-realSubPaymentSum,0) %>원</font>
            	<% else %>
            		<%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice-realSubPaymentSum,0) %>원
				<% end if %>
				<% if (prevrefundsum > 0) then %>
				    <% if (not IsTicketOrder) then %>
    					<% if (oordermaster.FOneItem.FCancelyn = "Y") and ((prevrefundsum - oordermaster.FOneItem.Fsubtotalprice - csbeasongpaysum) <> 0) then %>
    						(환불 <%= FormatNumber((prevrefundsum - oordermaster.FOneItem.Fsubtotalprice - csbeasongpaysum), 0) %>원 제외 )
    					<% elseif (oordermaster.FOneItem.FCancelyn <> "Y") then %>
    						(환불 <%= FormatNumber(prevrefundsum - csbeasongpaysum, 0) %>원 제외)
    					<% end if %>
    				<% end if %>
				<% end if %>
				<% if (csbeasongpaysum > 0) then %>
					배송비환급 : <%= FormatNumber(csbeasongpaysum, 0) %>원
				<% end if %>
            	&nbsp;
                [<%= oordermaster.FOneItem.JumunMethodName %> <%= getCardDEtailInfo(oordermaster.FOneItem.Fpggubun, cardcodeall) %>]
                [<font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font>]
                [<font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %></font>]
                <% if iPgGubun="NP" then %><font color=red>네이버페이</font><% end if %>
                <% if (realdepositsum>0) then %>
                   /&nbsp; <%= FormatNumber(realdepositsum,0) %>원&nbsp; [예치금]
                <% end if %>
                <% if (realgiftcardsum>0) then %>
                   /&nbsp; <%= FormatNumber(realgiftcardsum,0) %>원&nbsp; [Gift카드]
                <% end if %>

				<% if (oordermaster.FOneItem.Faccountdiv="110") then %>
                	(OK Cashbag사용 : <strong><%= FormatNumber(oordermaster.FOneItem.FokcashbagSpend,0) %></strong> 원)
                <% end if %>
                </b>
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td width="100" height="30">환불방식</td>
            <td width="600">
            	<%
            	'// drawSelectBoxCancelTypeBox 는 /lib/classes/cscenter/cs_aslistcls.asp 참조
            	%>
                <% call drawSelectBoxCancelTypeBox("returnmethod",orefund.FOneItem.Freturnmethod,oordermaster.FOneItem.Faccountdiv,divcd,"onChange='ChangeReturnMethod(this);'") %>
                <% if (Not IsStatusRegister) then %>
                (<%= orefund.FOneItem.FreturnmethodName %>)
                <% end if %>
                <input name="RefundRecalcuButton" class="csbutton" type="button" value="재계산" onClick="CalculateAndApplyItemCostSum(frmaction);">
                <% if (oordermaster.FOneItem.Faccountdiv = "100") or (oordermaster.FOneItem.Faccountdiv = "110") then %>
                	<% if (cardPartialCancelok = "Y") then %>
                		<font color="blue">신용카드 부분취소 가능카드</font>
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
            <td width="100" height="30">은행정보</td>
            <td align="left">
                <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="BABABA">
	            	<tr bgcolor="FFFFFF">
	            		<td width="80">계좌번호</td>
	            		<td>
	            		    <input class="text" type="text" size="20" name="rebankaccount" value="<%= orefund.FOneItem.Frebankaccount %>" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %> >
	            		    <input class="csbutton" type="button" value="이전내역(<%= prevrefundhistorycnt %>)" onClick="popPreReturnAcct('<%= oordermaster.FOneItem.Fuserid %>','frmaction','rebankaccount','rebankownername','rebankname');" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>>
                            &nbsp;
                            <input class="csbutton" type="button" value="환불정보입력요청" onClick="popRequestReturnAcctLMS('<%= id %>','<%= oordermaster.FOneItem.Forderserial %>', '<%= oordermaster.FOneItem.Fbuyhp %>');" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>>
	            		</td>
	            	</tr>
	            	<tr bgcolor="FFFFFF">
	            		<td>예금주명</td>
	            		<td><input class="text" type="text" size="20" name="rebankownername" value="<%= orefund.FOneItem.Frebankownername %>" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>></td>
	            	</tr>
	                <tr bgcolor="FFFFFF">
	            		<td>거래은행</td>
	            		<td><% DrawBankCombo "rebankname", orefund.FOneItem.Frebankname %></td>
	            	</tr>
            	</table>
            </td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R100" <% if orefund.FOneItem.Freturnmethod="R100" then response.write "" else response.write "style=""display:none""" %>>
    		<td width="100" height="30">PG사 ID</td>
    		<td><input class="text_ro" type="text" name="paygateTid1" size="30" value="<%= oordermaster.FOneItem.Fpaygatetid %>" readonly></td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R550" <% if orefund.FOneItem.Freturnmethod="R550" then response.write "" else response.write "style=""display:none""" %>>
    		<td width="100" height="30">쿠폰번호</td>
    		<td><input class="text_ro" type="text" name="paygateTid2" size="30" value="<%= oordermaster.FOneItem.Fpaygatetid %>" readonly></td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R560" <% if orefund.FOneItem.Freturnmethod="R560" then response.write "" else response.write "style=""display:none""" %>>
    		<td width="100" height="30">쿠폰번호</td>
    		<td><input class="text_ro" type="text" name="paygateTid3" size="30" value="<%= oordermaster.FOneItem.Fpaygatetid %>" readonly></td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R050" style="display:none">
            <td colspan="2" align="left" height="30">외부몰 환불요청</td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R900" style="display:none">
    		<td width="100" height="30">아이디</td>
    		<td><input class="text_ro" type="text" name="refund_userid" value="<%= oordermaster.FOneItem.Fuserid %>" readonly></td>
        </tr>
        <tr bgcolor="FFFFFF">
    		<td width="100" height="30">최초결제금액</td>
    		<td>
    		    <input class="text_ro" type="text" size="10" name="mainpaymentorg" value="<%= mainpaymentorg %>" maxlength=10 readonly>

    		    <input type=hidden name=cardcode value="<%= cardcode %>">
    		</td>
    	</tr>
        <tr bgcolor="FFFFFF">
    		<td width="100" height="30">기환불금액</td>
    		<td>
    		    <input class="text_ro" type="text" size="10" name="prevrefundsum" value="<%= CHKIIF(IsStatusRegister=True, prevrefundsum, prevrefundsum - orefund.FOneItem.Frefundrequire) %>" maxlength=10 readonly>
				* 환불 접수포함
    		</td>
    	</tr>

        <tr bgcolor="FFFFFF">
    		<td width="100" height="30">환불 예정액</td>
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
	          	* <font color=red><%= FormatNumber(RefundAllowLimit,0) %> 원</font> 초과 환불불가
	            <% end if %>
    		</td>
    		<% end if %>
    	</tr>
    	<% IF (Not (IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""))) then %>
        <tr bgcolor="FFFFFF">
    	    <td colspan="2" height="30"><b>환불 파일 작성중이므로 수정 할 수 없습니다.</b> [<%= orefund.FOneItem.Fupfiledate %>]</td>
    	</tr>
        <% end if %>

		<!-- 기존 환불정보가 없고, 환불요청인 경우 환불예정액 수정가능 -->
		<% if (divcd <> "A003") then %>
    	<tr bgcolor="FFFFFF">
    	    <td colspan="2" height="30">
    	    	* 환불예정액은 수정할 수 없습니다.<br>
    	    	* 환불액은 환불CS접수상태를 포함한 금액입니다.<br>
    	    	* 배송비환급은 배송비취소없이 이루어진 환급을 의미합니다.
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
	        alert('마일리지 환급은 완료처리시 자동 환급 됩니다.');
	    }
	    <% end if %>
	    </script>
		<% end if %>

	<% else %>
		<tr bgcolor="FFFFFF" ><td align="center" height="30">환불 가능 상태가 아닙니다.</td></tr>
	<% end if %>

<% end if %>
