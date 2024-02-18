<% if (IsDisplayRefundInfo) then %>

	<% if (IsCSCancelInfoNeeded(divcd) or (IsChangeOrder and IsCSReturnProcess(divcd))) then %>
		<%

		if (Not IsStatusRegister) then

			'// 정액 쿠폰 or 비율 쿠폰 중 한가지만 사용가능
			if (orgpercentcouponpricesum <> 0) or (regpercentcouponpricesum <> 0) or (curr_percentcouponsum <> 0) then

				'비율 쿠폰
				orefund.FOneItem.Forgpercentcouponsum		= orefund.FOneItem.Forgcouponsum
				orefund.FOneItem.Frefundpercentcouponsum 	= orefund.FOneItem.Frefundcouponsum

				orefund.FOneItem.Forgfixedcouponsum 		= 0
				orefund.FOneItem.Frefundfixedcouponsum 		= 0

			else

				if (curr_tencardspend <> 0) then

					'정액 쿠폰
					orefund.FOneItem.Forgpercentcouponsum		= 0
					orefund.FOneItem.Frefundpercentcouponsum	= 0

					orefund.FOneItem.Forgfixedcouponsum 		= orefund.FOneItem.Forgcouponsum
					orefund.FOneItem.Frefundfixedcouponsum 		= orefund.FOneItem.Frefundcouponsum

				else

					'비율 쿠폰
					orefund.FOneItem.Forgpercentcouponsum		= orefund.FOneItem.Forgcouponsum
					orefund.FOneItem.Frefundpercentcouponsum 	= orefund.FOneItem.Frefundcouponsum

					orefund.FOneItem.Forgfixedcouponsum 		= 0
					orefund.FOneItem.Frefundfixedcouponsum 		= 0

				end if

			end if

		end if

		%>

        <tr bgcolor="FFFFFF" align="center" height="25">
            <td width="120"></td>
            <td>선택</td>
            <td width="80">원 내역</td>
            <td width="80">선택상품</td>
            <td width="80">남는상품</td>
        </tr>
			<% if (IsDisplayItemList) and (IsStatusEdit) and (orefund.FOneItem.Frefunditemcostsum<>regitemcostsum) and (regitemcostsum<>0) then %>
            <script language='javascript'>alert('접수 금액 불일치-관리자 문의 요망');</script>
            <% end if %>
        <tr bgcolor="FFFFFF" height="25">
    		<td>
    			상품쿠폰적용액
    			<% if IsChangeOrder then %>
    				(브랜드)
    			<% end if %>
    		</td>
    		<td></td>
    		<td align="right">
    			<input class="text_ro" type="text" name="orgitemcostsum" value="<%= orefund.FOneItem.Forgitemcostsum %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    			<input class="text_ro" type="text" name="refunditemcostsum" value="0" size="8" style="text-align:right" readonly>
    		</td>
    	    <td align="right">
    	    	<input class="text_ro" type="text" name="remainitemcostsum" value="0" size="8" style="text-align:right" readonly>
    	    </td>
    	</tr>

<% if (IsDisplayItemList and (ocsOrderDetail.FResultCount > 0)) then %>

		<tr bgcolor="FFFFFF" height="25">
			<td>
				배송비
    			<% if IsChangeOrder then %>
    				(브랜드)
    			<% end if %>
			</td>
			<td>
			</td>
			<td align="right">
				<input class="text_ro" type="text" name="orgbeasongpay" value="<%= orefund.FOneItem.Forgbeasongpay %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
			</td>
			<td align="right">
				<input class="text_ro" type="text" name="refundbeasongpay" value="0" size="8" style="text-align:right" readonly>
			</td>
			<td align="right">
				<input class="text_ro" type="text" name="remainbeasongpay" value="0" size="8" style="text-align:right" readonly>
			</td>
		</tr>

<% end if %>

    	<tr bgcolor="FFFFFF" height="25">
    		<td><b>구매총액</b></td>
    		<td align="center"></td>
    		<td align="right">
    			<input class="text_ro" type="text" name="orgtotalbuypaysum" value="<%= orefund.FOneItem.Forgitemcostsum + orefund.FOneItem.Forgbeasongpay %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundtotalbuypaysum" value="0" size="8" style="text-align:right;background-color:#DDFFDD" readonly>
    		</td>
    		<td align="right">
    			<input class="text_ro" type="text" name="remaintotalbuypaysum" value="0" size="8" style="text-align:right;background-color:#FFFFFF;border-style: none" readonly>
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="10">
    		<td colspan="5"></td>
    	</tr>
        <tr bgcolor="FFFFFF" height="25">
    		<td>보너스쿠폰(비율)</td>
    		<td></td>
    		<td align="right">
    			<input class="text_ro" type="text" name="orgpercentcouponsum" value="<%= orefund.FOneItem.Forgpercentcouponsum*-1 %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    			<input class="text_ro" type="text" name="refundpercentcouponsum" value="<%= orefund.FOneItem.Frefundpercentcouponsum %>" size="8" style="text-align:right" readonly>
    		</td>
    	    <td align="right">
    	    	<input class="text_ro" type="text" name="remainpercentcouponsum" value="0" size="8" style="text-align:right" readonly>
    	    </td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25">
    		<td>사용 기타할인</td>
    		<td></td>
    		<td align="right">
    			<input class="text_ro" type="text" name="orgallatsubtractsum" value="<%= orefund.FOneItem.Forgallatdiscountsum*-1 %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundallatsubtractsum" value="0" size="8" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remainallatsubtractsum" value="0" size="8" style="text-align:right" readonly>

	            <!-- orgallatsubtractsum = allatdiscountprice * -1 -->
	            <input type="hidden" name="allatsubtractsum" value="0"><!-- refundallatsubtractsum -->
	            <input type="hidden" name="remainallatdiscount" value="0"><!-- remainallatsubtractsum -->
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25">
    		<td>보너스쿠폰(정액)</td>
    		<td><input type="checkbox" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> name="forcecouponreturn" <% if (orefund.FOneItem.Frefundfixedcouponsum <> 0) then %>checked<% end if %> onClick="CalculateAndApplyItemCostSum(frmaction)">환원</td>
    		<td align="right">
    			<input class="text_ro" type="text" name="orgfixedcouponsum" value="<%= orefund.FOneItem.Forgfixedcouponsum*-1 %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundfixedcouponsum" value="<%= orefund.FOneItem.Frefundfixedcouponsum %>" size="8" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remainfixedcouponsum" value="0" size="8" style="text-align:right" readonly>

		    	<input type="hidden" name="orgcouponsum" value="<%= orefund.FOneItem.Forgcouponsum * -1 %>"><!-- tencardspend * -1 -->
		    	<input type="hidden" name="refundcouponsum" value="<%= orefund.FOneItem.Frefundcouponsum %>">
		    	<input type="hidden" name="remaincouponsum" value="0">
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25">
    		<td>사용 마일리지</td>
    		<td><input type="checkbox" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> name="forcemileagereturn" <% if (orefund.FOneItem.Frefundmileagesum <> 0) then %>checked<% end if %> onClick="CalculateAndApplyItemCostSum(frmaction)">환원</td>
    		<td align="right">
    			<input class="text_ro" type="text" name="orgmileagesum" value="<%= orefund.FOneItem.Forgmileagesum * -1 %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundmileagesum" value="<%= orefund.FOneItem.Frefundmileagesum %>" size="8" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remainmileagesum" value="0" size="8" style="text-align:right" readonly>
    		</td>
    	</tr>
    	<!-- orgmileagesum = miletotalprice * -1 -->
    	<tr bgcolor="FFFFFF" height="25">
    		<td>사용 Gift카드</td>
    		<td><input type="checkbox" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> name="forcegiftcardreturn" <% if (orefund.FOneItem.Frefundgiftcardsum <> 0) then %>checked<% end if %> onClick="CalculateAndApplyItemCostSum(frmaction)">환원</td>
    		<td align="right">
   			<input class="text_ro" type="text" name="orggiftcardsum" value="<%= orefund.FOneItem.Forggiftcardsum * -1 %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundgiftcardsum" value="<%= orefund.FOneItem.Frefundgiftcardsum %>" size="8" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remaingiftcardsum" value="0" size="8" style="text-align:right" readonly>
    		</td>
    	</tr>
    	<!-- orggiftcardsum = giftcardsum * -1 -->
    	<tr bgcolor="FFFFFF" height="25">
    		<td>사용 예치금</td>
    		<td><input type="checkbox" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> name="forcedepositreturn" <% if (orefund.FOneItem.Frefunddepositsum <> 0) then %>checked<% end if %> onClick="CalculateAndApplyItemCostSum(frmaction)">환원</td>
    		<td align="right">
    			<input class="text_ro" type="text" name="orgdepositsum" value="<%= orefund.FOneItem.Forgdepositsum * -1 %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refunddepositsum" value="<%= orefund.FOneItem.Frefunddepositsum %>" size="8" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remaindepositsum" value="0" size="8" style="text-align:right" readonly>
    		</td>
    	</tr>
    	<!-- realdepositsum = depositsum * -1 -->
    	<tr bgcolor="FFFFFF" height="25">
    		<td>실결제금액</td>
    		<td></td>
    		<td align="right">
    			<input class="text_ro" type="text" name="orgtotalrealbuypaysum" value="0" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundtotalrealbuypaysum" value="0" size="8" style="text-align:right;background-color:#DDFFDD" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remaintotalrealbuypaysum" value="0" size="8" style="text-align:right" readonly>
    		</td>
    	</tr>

    	<tr bgcolor="FFFFFF" height="10">
    		<td colspan="5"></td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25" <% if (Not IsCSCancelProcess(divcd)) and (Not IsCSReturnProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td rowspan="3">
				<% if (IsTravelOrder) then %>
				<font color="red">취소 수수료</font>
				<% elseif IsCSCancelProcess(divcd) then %>
				<font color="red">추가 배송비</font>
				<% else %>
				<font color="red">반품 배송비</font>
				<% end if %>
			</td>
    		<td>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpay" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)"  >
				<% if (IsTravelOrder) then %>
				취소수수료 차감
				<% elseif IsCSCancelProcess(divcd) then %>
				----
				<% else %>
				업체왕복배송비
				<% end if %>
        		<!-- 추후 출고 배송비 차감으로 변경 -->
        		<br>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpayHalf" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)"  >
				<% if (IsTravelOrder) then %>
				일부차감
				<% elseif IsCSCancelProcess(divcd) then %>
				추가배송비
				<% else %>
				회수배송비
				<% end if %>
        		<br>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpayZero" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)">
        		차감안함
    		</td>
    		<td align="center" style="line-height: 20px;">
				차감예정<br><font color="red">차감액</font>
			</td>
    		<td align="right">
				<input class="text" type="text_ro" name="addbeasongpay" value="<%= ocsaslist.FOneItem.Fcustomeraddbeasongpay %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly><br>
				<input class="text" type="text" name="refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay %>" size="8" style="text-align:right" style="text-align:right" onChange="CalculateUpcheReturnBeasongPay(frmaction);  CalculateAndApplyItemCostSum(frmaction);">
			</td>
    	    <td></td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25" <% if (Not IsCSCancelProcess(divcd)) and (Not IsCSReturnProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td height="25">
        		&nbsp;
				<% if (IsTravelOrder) then %>
				수수료 추가결제
				<% else %>
				배송비 추가결제
				<% end if %>
    		</td>
    		<td colspan="3" height="25">
        		<input type="radio" name="addmethod" value="0" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "0") then %>checked<% end if %> >없음
				<input type="radio" name="addmethod" value="3" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "3") then %>checked<% end if %> >무통장
				<input type="radio" name="addmethod" value="1" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "1") then %>checked<% end if %> >박스동봉
				<input type="radio" name="addmethod" value="5" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "5") then %>checked<% end if %> >기타
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25" <% if (Not IsCSCancelProcess(divcd)) and (Not IsCSReturnProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td colspan="4" height="70">
				<% if (IsTravelOrder) then %>
        		&nbsp; * 항공권발행 다음날부터 취소수수료 차감<br>
        		&nbsp; * 출발 6일전부터 <font color="red">취소환불 불가</font><br>
				&nbsp; * 수수료는 티켓당 부과
				<% elseif IsCSCancelProcess(divcd) then %>
				&nbsp; * 단순변심<br>
				&nbsp; * 브랜드 잔여상품이 업체조건배송비 미만 + 무료배송 쿠폰/상품 없음<br>
				&nbsp; * 배송비는 <font color="red">브랜드별 배송비</font>로 차감
				<% else %>
        		&nbsp; * 단순변심 + 브랜드전체반품 : 업체왕복배송비<br>
        		&nbsp; * 기타 단순변심반품 : 회수배송비<br>
        		&nbsp; * 상품불량반품 등 : 차감없음<br>
        		&nbsp; * 배송비는 <font color="red">브랜드별 배송비</font>로 차감
				<% end if %>
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25" <% if (TRUE) or (Not IsCSCancelProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td rowspan="2"><font color="red">추가 배송비</font></td>
    		<td></td>
    		<td></td>
    		<td align="right"><input class="text_ro" type="text" name="adddeliverypay" value="0" size="8" style="text-align:right" style="text-align:right" ></td>
    	    <td></td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25" <% if (TRUE) or (Not IsCSCancelProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td colspan="4" height="40">
        		&nbsp; 단순변심 + 업체조건배송 + 무료배송 + 무료(착불)배송상품X<br>
        		&nbsp; + 무료배송조건 이하로 상품취소
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25">
    		<td>기타보정금액
    		<% if (IsTicketOrder) then %>
    		<br><strong>(취소 수수료)</strong>
    		<% end if %>
    		</td>
    		<% if (IsTicketOrder) then %>
    		<td colspan="2">
        		<input type="radio" name="tRefundPro" value="0" onClick="calcuTicketCancelCharge(this);" > 없음
				<input type="radio" name="tRefundPro" value="2000" <%= chkIIF(mayTicketCancelChargePro=2000,"checked","") %> onClick="calcuTicketCancelCharge(this);" > 2,000원(티켓금액의 10%한도)<br>
        		<input type="radio" name="tRefundPro" value="10" <%= chkIIF(mayTicketCancelChargePro=10,"checked","") %> onClick="calcuTicketCancelCharge(this);" > 10%
        		<input type="radio" name="tRefundPro" value="20" <%= chkIIF(mayTicketCancelChargePro=20,"checked","") %> onClick="calcuTicketCancelCharge(this);" > 20%
        		<input type="radio" name="tRefundPro" value="30" <%= chkIIF(mayTicketCancelChargePro=30,"checked","") %> onClick="calcuTicketCancelCharge(this);" > 30%
    		</td>
    		<% else %>
    		<td></td>
    		<td align="right"></td>
    		<% end if %>
			<td align="right"><input class="text_ro" type="text" name="refundadjustpay" value="<%= orefund.FoneItem.Frefundadjustpay %>" size="8" style="text-align:right" onBlur="CalculateAndApplyItemCostSum(frmaction);" readonly></td>
            <td align="right"></td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25">
            <td>총액/취소액</td>
            <td></td>
            <td align="right">
                <input class="text_ro" type="text" name="orgsubtotalprice" value="0" size="8" readonly style="text-align:right;background-color:#FFFFFF;border-style:none" ><!-- subtotalprice -->
            </td>
            <td align="right">
            	<input class="text_ro" type="text" name="refundsubtotalprice" value="0" size="8" readonly style="text-align:right;background-color:#DDFFDD" ><!-- canceltotal -->
            </td>
            <td align="right">
            	<input class="text_ro" type="text" name="remainsubtotalprice" value="0" size="8" readonly style="text-align:right" ><!-- nextsubtotal -->

	            <input type="hidden" name="subtotalprice" value="0">
	            <input type="hidden" name="canceltotal" value="0">
	            <input type="hidden" name="nextsubtotal" value="0">
            </td>
        </tr>

	<% else %>

	<tr bgcolor="FFFFFF" ><td align="center" height="30">취소 가능 상태가 아닙니다.</td></tr>

	<% end if %>

<% end if %>
