<% if (IsDisplayRefundInfo) then %>

	<% if (IsCSCancelInfoNeeded(divcd)) then %>
		<%

'		if (IsStatusRegister) then
'			'접수시 초기값 세팅
'			orefund.FOneItem.Forgitemcostsum 	= orgitemcostsum
'			orefund.FOneItem.Forgbeasongpay 	= SumBeasongPayNotCanceled
'
'			orefund.FOneItem.Forgmileagesum 	= oordermaster.FOneItem.Fmiletotalprice
'			'orefund.FOneItem.Forgcouponsum 		= oordermaster.FOneItem.Ftencardspend
'			orefund.FOneItem.Fallatsubtractsum 	= oordermaster.FOneItem.Fallatdiscountprice*-1
'			orefund.FOneItem.Forgdepositsum		= realdepositsum
'			orefund.FOneItem.Forggiftcardsum	= realgiftcardsum
'
'			orefund.FoneItem.Frefundadjustpay 	= 0
'			orefund.FOneItem.Frefunddeliverypay = 0
'
'			orefund.FOneItem.Frefundcouponsum	= 0
'			orefund.FOneItem.Frefundmileagesum	= 0
'
'            orefund.FOneItem.Frefundgiftcardsum = 0
'            orefund.FOneItem.Frefunddepositsum  = 0
'
'		end if
'
'		'반품접수후 주문취소가 발생하면 값이 틀려진다.
'		orefund.FOneItem.Forgcouponsum 		= oordermaster.FOneItem.Ftencardspend
'
'		orefund.FOneItem.Forgpercentcouponsum = orgpercentcouponpricesum
'		orefund.FOneItem.Frefundpercentcouponsum = regpercentcouponpricesum * -1
'
'		orefund.FOneItem.Forgfixedcouponsum = orefund.FOneItem.Forgcouponsum - orefund.FOneItem.Forgpercentcouponsum
'		orefund.FOneItem.Frefundfixedcouponsum = orefund.FOneItem.Frefundcouponsum - orefund.FOneItem.Frefundpercentcouponsum

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
    		<td>상품쿠폰적용액</td>
    		<td width="80"></td>
    		<td align="right" width="80">
    			<input class="text_ro" type="text" name="orgitemcostsum" value="<%= orefund.FOneItem.Forgitemcostsum %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right" width="80">
    			<input class="text_ro" type="text" name="refunditemcostsum" value="0" size="8" style="text-align:right" readonly>
    		</td>
    	    <td align="right" width="80">
    	    	<input class="text_ro" type="text" name="remainitemcostsum" value="0" size="8" style="text-align:right" readonly>
    	    </td>
    	</tr>

<% if (IsDisplayItemList and (ocsOrderDetail.FResultCount > 0)) then %>

		<tr bgcolor="FFFFFF" height="25">
			<td>배송비</td>
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

		<% '스크립트를 단순화하기 위해 아래와 같이 더미를 더 만들어 둔다.(orderdetailidx 가 한개인 경우와 2개이상인 경우를 분리해서 작성하지 않아도 된다.) %>
		<input type="hidden" name="CancelDeliverMakerid">
		<input type="hidden" name="ckbeasongpayAssign">
		<input type="hidden" name="orgbrandbeasongpay">
		<input type="hidden" name="refundbrandbeasongpay">
		<input type="hidden" name="remainbrandbeasongpay">

		<input type="hidden" name="CancelDeliverMakerid">
		<input type="hidden" name="ckbeasongpayAssign">
		<input type="hidden" name="orgbrandbeasongpay">
		<input type="hidden" name="refundbrandbeasongpay">
		<input type="hidden" name="remainbrandbeasongpay">

	<% for i = 0 to ocsOrderDetail.FResultCount - 1 %>

		<%
		IsItemCanceled = (ocsOrderDetail.FItemList(i).FCancelyn = "Y")
		OrderDetailState = ocsOrderDetail.FItemList(i).ForderDetailcurrstate
		IsBeasongPay = (ocsOrderDetail.FItemList(i).Fitemid = 0)

		if (IsBeasongPay) then
			IsUpcheBeasong = (ocsOrderDetail.FItemList(i).Fmakerid <> "")
		else
			IsUpcheBeasong = (ocsOrderDetail.FItemList(i).Fisupchebeasong = "Y")
		end if
		%>

		<% if (IsBeasongPay and ((Not IsItemCanceled) or (IsStatusFinished and IsCSCancelProcess(divcd) and (ocsOrderDetail.FItemList(i).Fgubun01name <> "")))) then %>

    	<tr bgcolor="FFFFFF" height="25">
			<input type="hidden" name="CancelDeliverMakerid" value="<%= ocsOrderDetail.FItemList(i).FMakerId %>">
    		<td>
    			&nbsp; -
    			<% if Not IsUpcheBeasong then %>
    				텐바이텐
    			<% else %>
    				<%= ocsOrderDetail.FItemList(i).FMakerId %>
    			<% end if %>
    		</td>
    		<td>
    			<input type="checkbox" name="ckbeasongpayAssign" value="" <% if Not (IsStatusEdit or IsStatusRegister) or (Not IsCSCancelProcess(divcd) and Not IsCSReturnProcess(divcd)) then %>disabled<% end if %> onClick="FouceCheckDeliverPay(frmaction, '<%= ocsOrderDetail.FItemList(i).FMakerId %>', this.checked)" <% if (Not IsStatusRegister) and (ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) > 0) then %>checked<% end if %>><font color="red">환급</font>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="orgbrandbeasongpay" value="<%= ocsOrderDetail.FItemList(i).GetItemCouponPrice %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundbrandbeasongpay" value="<% if (Not IsStatusRegister) and (ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) > 0) then %><%= ocsOrderDetail.FItemList(i).GetItemCouponPrice %><% else %>0<% end if %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remainbrandbeasongpay" value="<%= ocsOrderDetail.FItemList(i).GetItemCouponPrice %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    	</tr>

		<% end if %>

	<% next %>

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
    		<td>사용 보너스쿠폰(비율)</td>
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
    			<input class="text_ro" type="text" name="orgallatsubtractsum" value="<%= orefund.FOneItem.Fallatsubtractsum %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly> <!-- 서동석 수정 *-1 뺌 13번 라인 *-1-->
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundallatsubtractsum" value="0" size="8" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remainallatsubtractsum" value="0" size="8" style="text-align:right" readonly>
    		</td>
            <!-- orgallatsubtractsum = allatdiscountprice * -1 -->
            <input type="hidden" name="allatsubtractsum" value="0"><!-- refundallatsubtractsum -->
            <input type="hidden" name="remainallatdiscount" value="0"><!-- remainallatsubtractsum -->
    	</tr>
    	<tr bgcolor="FFFFFF" height="25">
    		<td>사용 보너스쿠폰(정액)</td>
    		<td><input type="checkbox" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> name="forcecouponreturn" <% if (orefund.FOneItem.Frefundfixedcouponsum <> 0) then %>checked<% end if %> onClick="CalculateAndApplyItemCostSum(frmaction)">환원</td>
    		<td align="right">
    			<input class="text_ro" type="text" name="orgfixedcouponsum" value="<%= orefund.FOneItem.Forgfixedcouponsum*-1 %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="refundfixedcouponsum" value="<%= orefund.FOneItem.Frefundfixedcouponsum %>" size="8" style="text-align:right" readonly>
    		</td>
    		<td align="right">
    		    <input class="text_ro" type="text" name="remainfixedcouponsum" value="0" size="8" style="text-align:right" readonly>
    		</td>
    	</tr>
    	<input type="hidden" name="orgcouponsum" value="<%= orefund.FOneItem.Forgcouponsum * -1 %>"><!-- tencardspend * -1 -->
    	<input type="hidden" name="refundcouponsum" value="<%= orefund.FOneItem.Frefundcouponsum %>">
    	<input type="hidden" name="remaincouponsum" value="0">
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
    	<tr bgcolor="FFFFFF" height="25" <% if (Not IsCSReturnProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td rowspan="2"><font color="red">반품 배송비</font></td>
    		<td>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpay" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)" <% if (orefund.FOneItem.Frefunddeliverypay <= -4000) then response.write "checked" %> >
        		업체왕복배송비
        		<!-- 추후 출고 배송비 차감으로 변경 -->
        		<br>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpayHalf" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)"  <% if (orefund.FOneItem.Frefunddeliverypay < 0) and (orefund.FOneItem.Frefunddeliverypay > -4000) then response.write "checked" %> >
        		회수배송비
        		<br>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpayZero" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)">
        		차감안함
    		</td>
    		<td align="center"></td>
    		<td align="right"><input class="text" type="text" name="refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay %>" size="8" style="text-align:right" style="text-align:right" onBlur="CalculateAndApplyItemCostSum(frmaction);"></td>
    	    <td></td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25" <% if (Not IsCSReturnProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td colspan="4" height="70">
        		&nbsp; * 단순변심 + 브랜드전체반품 : 업체왕복배송비<br>
        		&nbsp; * 기타 단순변심반품 : 회수배송비<br>
        		&nbsp; * 상품불량반품 등 : 차감없음<br>
        		&nbsp; * 배송비는 <font color="red">브랜드별 배송비</font>로 차감
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
        		<input type="radio" name="tRefundPro" value="10" <%= chkIIF(mayTicketCancelChargePro=10,"checked","") %> onClick="calcuTicketCancelCharge(this);" > 10%
        		<input type="radio" name="tRefundPro" value="20" <%= chkIIF(mayTicketCancelChargePro=20,"checked","") %> onClick="calcuTicketCancelCharge(this);" > 20%
        		<input type="radio" name="tRefundPro" value="30" <%= chkIIF(mayTicketCancelChargePro=30,"checked","") %> onClick="calcuTicketCancelCharge(this);" > 30%
    		</td>
    		<% else %>
    		<td></td>
    		<td align="right"></td>
    		<% end if %>
    		<td align="right"><input class="text" type="text" name="refundadjustpay" value="<%= orefund.FoneItem.Frefundadjustpay %>" size="8" style="text-align:right" onBlur="CalculateAndApplyItemCostSum(frmaction);"></td>
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
            </td>
            <input type="hidden" name="subtotalprice" value="0">
            <input type="hidden" name="canceltotal" value="0">
            <input type="hidden" name="nextsubtotal" value="0">
        </tr>

	<% else %>

	<tr bgcolor="FFFFFF" ><td align="center" height="30">취소 가능 상태가 아닙니다.</td></tr>

	<% end if %>

<% end if %>