<% if (IsDisplayRefundInfo) then %>

	<% if (IsCSRefundNeeded(divcd, OrderMasterState)) then %>

        <tr bgcolor="#FFFFFF">
            <td width="100" height="30">결제정보</td>
            <td width="600">
            	<b>
           	    결제금액 : <%= FormatNumber(ogiftcardordermaster.FOneItem.Fsubtotalprice, 0) %>

            	&nbsp;
                [<%= ogiftcardordermaster.FOneItem.GetAccountdivName %>]
                [<font color="<%= ogiftcardordermaster.FOneItem.CancelYnColor %>"><%= ogiftcardordermaster.FOneItem.CancelYnName %></font>]
                [<font color="<%= ogiftcardordermaster.FOneItem.IpkumDivColor %>"><%= ogiftcardordermaster.FOneItem.GetJumunDivName %></font>]
                </b>
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td width="100" height="30">환불방식</td>
            <td width="600">
                <% call drawSelectBoxCancelTypeBox("returnmethod",orefund.FOneItem.Freturnmethod,ogiftcardordermaster.FOneItem.Faccountdiv,divcd,"onChange='ChangeReturnMethod(this);'") %>
                <% if (Not IsStatusRegister) then %>
                (<%= orefund.FOneItem.FreturnmethodName %>)
                <% end if %>
            </td>
        </tr>
        <tr  bgcolor="FFFFFF" id="refundinfo_R007" <% if orefund.FOneItem.Freturnmethod="R007" then response.write "" else response.write "style='display:none'" %>>
            <td width="100" height="30">은행정보</td>
            <td align="left">
                <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="BABABA">
	            	<tr bgcolor="FFFFFF">
	            		<td width="80">계좌번호</td>
	            		<td>
	            		    <input class="text" type="text" size="20" name="rebankaccount" value="<%= orefund.FOneItem.Frebankaccount %>" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %> >
	            		    <input class="csbutton" type="button" value="이전내역" onClick="popPreReturnAcct('<%= ogiftcardordermaster.FOneItem.Fuserid %>','frmaction','rebankaccount','rebankownername','rebankname');" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>>
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
        <tr bgcolor="FFFFFF" id="refundinfo_R100" <% if orefund.FOneItem.Freturnmethod="R100" then response.write "" else response.write "style='display:none'" %>>
    		<td width="100" height="30">PG사 ID</td>
    		<td><input class="text_ro" type="text" name="paygateTid" size="30" value="<%= ogiftcardordermaster.FOneItem.Fpaydateid %>" readonly></td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R050" style="display:none">
            <td colspan="2" align="left" height="30">외부몰 환불요청</td>
        </tr>
        <tr bgcolor="FFFFFF" id="refundinfo_R900" style="display:none">
    		<td width="100" height="30">아이디</td>
    		<td><input class="text_ro" type="text" name="refund_userid" value="<%= ogiftcardordermaster.FOneItem.Fuserid %>" readonly></td>
        </tr>
        <tr bgcolor="FFFFFF">
    		<td width="100" height="30">환불 예정액</td>
    		<% if (orefund.FResultCount>0) then %>
    		<td>
    		    <input class="text_ro" type="text" size="10" name="refundrequire" value="<%= orefund.FOneItem.Frefundrequire %>" maxlength=7 readonly>
    		    (<%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %>)
    		</td>
    		<% else %>
    		<td><input class="text_ro" type="text" size="10" name="refundrequire" value="<%= ogiftcardordermaster.FOneItem.Fsubtotalprice %>"></td>
    		<% end if %>
    	</tr>
    	<% IF (Not (IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""))) then %>
        <tr bgcolor="FFFFFF">
    	    <td colspan="2" height="30"><b>환불 파일 작성중이므로 수정 할 수 없습니다.</b> [<%= orefund.FOneItem.Fupfiledate %>]</td>
    	</tr>
        <% end if %>

		<% if (IsStatusFinishing or IsStatusFinished) then %>
	    <script language='javascript'>
	    frmaction.returnmethod.disabled=true;
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
