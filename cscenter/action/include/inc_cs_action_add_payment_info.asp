
<% if divcd = "A999" then %>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">상품대금</td>
	    	    <td>
	    	    	<input type="text" class="text_ro" id="add_customeradditempay" name="add_customeradditempay" value="<%= ocsaslist.FOneItem.Fcustomeradditempay %>" size="20" ReadOnly>
					&nbsp;
					<% if IsStatusRegister then %>
					<input type="radio" name="customerpayordertype" value="B" <%= CHKIIF(ocsaslist.FOneItem.Fcustomerpayordertype="B" or ocsaslist.FOneItem.Fcustomerpayordertype="", "checked", "") %> onClick="CheckSetAddItemPayment(frmaction)"> 결제안함
					&nbsp;
					<input type="radio" name="customerpayordertype" value="A" <%= CHKIIF(ocsaslist.FOneItem.Fcustomerpayordertype="A", "checked", "") %> onClick="CheckSetAddItemPayment(frmaction)"> 기출고결제
                    &nbsp;
					<input type="radio" name="customerpayordertype" value="N" <%= CHKIIF(ocsaslist.FOneItem.Fcustomerpayordertype="N", "checked", "") %> onClick="CheckSetAddItemPayment(frmaction)" <%= CHKIIF((session("ssBctId") = "hasora") or (session("ssBctId") = "boyishP") or (session("ssBctId") = "oesesang52") or (session("ssBctId") = "rabbit1693"), "", "disabled") %>> 신규주문
					<% else %>
					* <%= ocsaslist.FOneItem.GetCustomerPayOrderTypeName() %>
                    <input type="hidden" name="customerpayordertype" value="<%= ocsaslist.FOneItem.Fcustomerpayordertype %>">
					<% end if %>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">상품매입가</td>
	    	    <td>
	    	    	<input type="text" class="text_ro" id="add_customeradditembuypay" name="add_customeradditembuypay" value="<%= ocsaslist.FOneItem.Fcustomeradditembuypay %>" size="20" ReadOnly>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">배송비</td>
	    	    <td>
	    	    	<input type="text" class="text_ro" id="add_customeraddbeasongpay" name="add_customeraddbeasongpay" value="<%= ocsaslist.FOneItem.Fcustomeraddbeasongpay %>" size="20" ReadOnly onKeyUp="CalculateUpcheAddPayment(frmaction);">
					&nbsp;
					<% if IsStatusRegister then %>
					<input type="button" class="button" value="배송비" style="width: 80px;" onClick="CheckSetAddBeasongPayment(frmaction, '1')">
					&nbsp;
					<input type="button" class="button" value="왕복배송비" style="width: 80px;" onClick="CheckSetAddBeasongPayment(frmaction, '2')">
					&nbsp;
					<input type="button" class="button" value="기타" style="width: 80px;" onClick="CheckSetAddBeasongPayment(frmaction, 'E')">
					<% else %>
					* <font color="red">접수 이후</font>에는 수정불가
					<% end if %>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">추가결제(예정)</td>
	    	    <td>
	    	    	<input type="text" class="text_ro" id="add_customeraddpay_sum" name="add_customeraddpay_sum" value="<%= ocsaslist.FOneItem.Fcustomeradditempay + ocsaslist.FOneItem.Fcustomeraddbeasongpay %>" size="20" ReadOnly onKeyUp="CalculateUpcheAddPayment(frmaction);">
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100" height="25">결제수단</td>
	    	    <td>
					<input type="radio" name="accountdiv" value="7" checked> 무통장 입금(가상계좌)
	    	    </td>
	    	</tr>
			<% if IsStatusRegister then %>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">입금하실 은행</td>
	    	    <td>
					<select name='acctno' class="select" title="입금 은행 선택">
						<option value="">입금하실 은행을 선택하세요.</option>
						<option value="11">농    협</option>
						<option value="06">국민은행</option>
						<option value="20">우리은행</option>
						<option value="26">신한은행</option>
						<option value="81">하나은행</option>
						<option value="03">기업은행</option>
						<!-- option value="05">외환은행 : 사용불가</option -->
						<option value="39">경남은행</option>
						<option value="32">부산은행</option>
						<!-- option value="31">대구은행</option  -->
						<option value="71">우체국</option>
						<option value="07">수협</option>
					</select>
					예금주 : (주)텐바이텐
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">입금자명</td>
	    	    <td>
					<input name="acctname" type="text" maxlength="12" class="txtInp" style="width:200px;" value="<%= oordermaster.FOneItem.FBuyname %>" id="depositName" />
	    	    </td>
	    	</tr>
			<% else %>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100" height="25">관련주문번호</td>
	    	    <td>
					<% if (payorderserial <> "") then %>
						<%= payorderserial %>
	                	[<font color="<%= opayordermaster.FOneItem.CancelYnColor %>"><%= opayordermaster.FOneItem.CancelYnName %></font>]
	                	[<font color="<%= opayordermaster.FOneItem.IpkumDivColor %>"><%= opayordermaster.FOneItem.IpkumDivName %></font>]
					<% end if %>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100" height="25">가상계좌</td>
	    	    <td>
					<% if (payorderserial <> "") then %>
					<% if opayordermaster.FOneItem.FAccountDiv="7" then %>
						<% if C_CriticInfoUserLV1 then %>
				    	<%= opayordermaster.FOneItem.FAccountNo %>
				    	&nbsp;
						<% end if %>
				    	<% if opayordermaster.FOneItem.IsDacomCyberAccountPay then %>
					    <font color="red">[가상]</font>
					    <% else %>
					    [일반]
					    <% end if %>
					<% end if %>
					<% end if %>
	    	    </td>
	    	</tr>
			<% end if %>
<% end if %>
