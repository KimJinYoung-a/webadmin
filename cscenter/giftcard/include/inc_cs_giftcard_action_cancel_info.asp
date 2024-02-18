<% if (IsDisplayRefundInfo) then %>

	<% if (IsCSCancelInfoNeeded(divcd)) then %>

        <tr bgcolor="FFFFFF" align="center" height="25">
            <td></td>
            <td>선택</td>
            <td>원 내역</td>
            <td>선택상품</td>
            <td>남는상품</td>
        </tr>
			<% if (IsDisplayItemList) and (IsStatusEdit) and (orefund.FOneItem.Frefunditemcostsum<>regitemcostsum) and (regitemcostsum<>0) then %>
            <script language='javascript'>alert('접수 금액 불일치-관리자 문의 요망');</script>
            <% end if %>
    	<tr bgcolor="FFFFFF" height="25">
            <td>총액/취소액</td>
            <td></td>
            <td align="right">
                <input class="text_ro" type="text" name="orgsubtotalprice" value="<%= ogiftcardordermaster.FOneItem.Fsubtotalprice %>" size="8" readonly style="text-align:right;background-color:#FFFFFF;border-style:none" ><!-- subtotalprice -->
            </td>
            <td align="right">
            	<input class="text_ro" type="text" name="refundsubtotalprice" value="<%= ogiftcardordermaster.FOneItem.Fsubtotalprice %>" size="8" readonly style="text-align:right;background-color:#DDFFDD" ><!-- canceltotal -->
            </td>
            <td align="right">
            	<input class="text_ro" type="text" name="remainsubtotalprice" value="0" size="8" readonly style="text-align:right" ><!-- nextsubtotal -->
            </td>
            <input type="hidden" name="subtotalprice" value="<%= ogiftcardordermaster.FOneItem.Fsubtotalprice %>">
            <input type="hidden" name="canceltotal" value="<%= ogiftcardordermaster.FOneItem.Fsubtotalprice %>">
            <input type="hidden" name="nextsubtotal" value="0">
            <input type="hidden" name="refunditemcostsum" value="<%= ogiftcardordermaster.FOneItem.Fsubtotalprice %>">
            <input type="hidden" name="remainitemcostsum" value="0">
        </tr>

	<% else %>

	<tr bgcolor="FFFFFF" ><td align="center" height="30">취소 가능 상태가 아닙니다.</td></tr>

	<% end if %>

<% end if %>