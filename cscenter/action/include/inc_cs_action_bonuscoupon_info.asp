<% if (IsCSCancelProcess(divcd) or IsCSReturnProcess(divcd)) then %>

	<% if IsBonusCouponExist then %>

	    	<tr bgcolor="FFFFFF">
	    	    <td width="100" height="27">쿠폰명</td>
	    	    <td ><%= ocscoupon.FOneItem.Fcouponname %></td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100" height="27">할인값</td>
	    	    <td >
					<%= FormatNumber(ocscoupon.FOneItem.Fcouponvalue, 0) %> <%= ocscoupon.FOneItem.GetCouponTypeUnit %>
					(최소구매 : <%= FormatNumber(ocscoupon.FOneItem.Fminbuyprice, 0) %> 원)
				</td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100" height="27">유효기간</td>
	    	    <td >
					<%= ocscoupon.FOneItem.Fstartdate %> ~ <%= ocscoupon.FOneItem.Fexpiredate %>
				</td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100" height="27">상태</td>
	    	    <td >
					<% if IsStatusFinished then %>
						<input type="checkbox" name="tmpcopycouponinfo" value="Y" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> onClick="jsCheckCopyCoupon(frmaction)" <% if orefund.FOneItem.Fcopycouponinfo = "Y" then %>checked<% end if %> > <font color="red">재발급</font>
					<% else %>
						<% if ocscoupon.FOneItem.IsCouponCopyValid then %>
							<input type="checkbox" name="tmpcopycouponinfo" value="Y" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> onClick="jsCheckCopyCoupon(frmaction)" <% if orefund.FOneItem.Fcopycouponinfo = "Y" then %>checked<% end if %> > <font color="red">재발급</font>
							&nbsp;
							&nbsp;
						<% end if %>
						<%= ocscoupon.FOneItem.GetCouponStatus %>
					<% end if %>
				</td>
	    	</tr>

	<% end if %>

<% end if %>
