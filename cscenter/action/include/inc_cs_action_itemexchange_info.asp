<!--
<% if IsCSItemExchangeReceiveInfoNeeded(divcd) then %>
	    	<tr bgcolor="FFFFFF" height="25">
	    	    <td width="100">상품회수상태</td>
	    	    <input type="hidden" name="orgreceiveyn" value="<%= ocsaslist.FOneItem.Freceiveyn %>">
	    	    <td>
					<input type="radio" class="radio" name="receiveyn" value="Y" <% if (ocsaslist.FOneItem.Freceiveyn = "Y") then %>checked<% end if %>> 상품회수완료
					<input type="radio" class="radio" name="receiveyn" value="N" <% if (ocsaslist.FOneItem.Freceiveyn = "N") then %>checked<% end if %>> 상품회수이전(맞교환출고를 먼저 할 경우)
	    	    </td>
	    	</tr>
<% end if %>
-->
