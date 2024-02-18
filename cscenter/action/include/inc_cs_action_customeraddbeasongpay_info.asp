
<% if IsCSItemExchangeCustomerBeasongPayNeeded(divcd) then %>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">추가배송비(예정)</td>
	    	    <td>
	    	    	<input type="text" class="text_ro" name="add_customeraddbeasongpay" value="<%= ocsaslist.FOneItem.Fcustomeraddbeasongpay %>" size="20" ReadOnly >
	    	    	&nbsp;
		    	    <select class="select" name="add_customeraddmethod" class="text">
			    	    <option value="">선택
			    	    <option value="1" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "1") then %>selected<% end if %>>박스동봉
			    	    <option value="2" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "2") then %>selected<% end if %>>택배비 고객부담
			    	    <option value="5" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "5") then %>selected<% end if %>>기타
		    	    </select>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">추가배송비(확인)</td>
	    	    <td>
					<% if IsStatusEdit and (divcd = "A111") then %>
						<input type="text" class="text" name="customerrealbeasongpay" value="<%= ocsaslist.FOneItem.Fcustomerrealbeasongpay %>" size="20">
					<% else %>
						<input type="text" class="text_ro" name="customerrealbeasongpay" value="<%= ocsaslist.FOneItem.Fcustomerrealbeasongpay %>" size="20" ReadOnly >
					<% end if %>
	    	    </td>
	    	</tr>
	    	<!--
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">결제상태</td>
				<input type="hidden" name="orgcustomerreceiveyn" value="<%= ocsaslist.FOneItem.Fcustomerreceiveyn %>">
	    	    <td>
					<input type="radio" class="radio" name="add_customerreceiveyn" value="Y" <% if (ocsaslist.FOneItem.Fcustomerreceiveyn = "Y") then %>checked<% end if %>> 입금확인
					<input type="radio" class="radio" name="add_customerreceiveyn" value="N" <% if (ocsaslist.FOneItem.Fcustomerreceiveyn = "N") then %>checked<% end if %>> 확인이전
	    	    </td>
	    	</tr>
	    	-->
<% end if %>
