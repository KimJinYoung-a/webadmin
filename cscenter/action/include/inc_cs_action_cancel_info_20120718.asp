<% if (IsDisplayRefundInfo) then %>

	<% if (IsCSCancelInfoNeeded(divcd)) then %>
		<%

'		if (IsStatusRegister) then
'			'������ �ʱⰪ ����
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
'		'��ǰ������ �ֹ���Ұ� �߻��ϸ� ���� Ʋ������.
'		orefund.FOneItem.Forgcouponsum 		= oordermaster.FOneItem.Ftencardspend
'
'		orefund.FOneItem.Forgpercentcouponsum = orgpercentcouponpricesum
'		orefund.FOneItem.Frefundpercentcouponsum = regpercentcouponpricesum * -1
'
'		orefund.FOneItem.Forgfixedcouponsum = orefund.FOneItem.Forgcouponsum - orefund.FOneItem.Forgpercentcouponsum
'		orefund.FOneItem.Frefundfixedcouponsum = orefund.FOneItem.Frefundcouponsum - orefund.FOneItem.Frefundpercentcouponsum

		if (Not IsStatusRegister) then

			'// ���� ���� or ���� ���� �� �Ѱ����� ��밡��
			if (orgpercentcouponpricesum <> 0) or (regpercentcouponpricesum <> 0) or (curr_percentcouponsum <> 0) then

				'���� ����
				orefund.FOneItem.Forgpercentcouponsum		= orefund.FOneItem.Forgcouponsum
				orefund.FOneItem.Frefundpercentcouponsum 	= orefund.FOneItem.Frefundcouponsum

				orefund.FOneItem.Forgfixedcouponsum 		= 0
				orefund.FOneItem.Frefundfixedcouponsum 		= 0

			else

				if (curr_tencardspend <> 0) then

					'���� ����
					orefund.FOneItem.Forgpercentcouponsum		= 0
					orefund.FOneItem.Frefundpercentcouponsum	= 0

					orefund.FOneItem.Forgfixedcouponsum 		= orefund.FOneItem.Forgcouponsum
					orefund.FOneItem.Frefundfixedcouponsum 		= orefund.FOneItem.Frefundcouponsum

				else

					'���� ����
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
            <td>����</td>
            <td>�� ����</td>
            <td>���û�ǰ</td>
            <td>���»�ǰ</td>
        </tr>
			<% if (IsDisplayItemList) and (IsStatusEdit) and (orefund.FOneItem.Frefunditemcostsum<>regitemcostsum) and (regitemcostsum<>0) then %>
            <script language='javascript'>alert('���� �ݾ� ����ġ-������ ���� ���');</script>
            <% end if %>
        <tr bgcolor="FFFFFF" height="25">
    		<td>��ǰ���������</td>
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
			<td>��ۺ�</td>
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

		<% '��ũ��Ʈ�� �ܼ�ȭ�ϱ� ���� �Ʒ��� ���� ���̸� �� ����� �д�.(orderdetailidx �� �Ѱ��� ���� 2���̻��� ��츦 �и��ؼ� �ۼ����� �ʾƵ� �ȴ�.) %>
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
    				�ٹ�����
    			<% else %>
    				<%= ocsOrderDetail.FItemList(i).FMakerId %>
    			<% end if %>
    		</td>
    		<td>
    			<input type="checkbox" name="ckbeasongpayAssign" value="" <% if Not (IsStatusEdit or IsStatusRegister) or (Not IsCSCancelProcess(divcd) and Not IsCSReturnProcess(divcd)) then %>disabled<% end if %> onClick="FouceCheckDeliverPay(frmaction, '<%= ocsOrderDetail.FItemList(i).FMakerId %>', this.checked)" <% if (Not IsStatusRegister) and (ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) > 0) then %>checked<% end if %>><font color="red">ȯ��</font>
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
    		<td><b>�����Ѿ�</b></td>
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
    		<td>��� ���ʽ�����(����)</td>
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
    		<td>��� ��Ÿ����</td>
    		<td></td>
    		<td align="right">
    			<input class="text_ro" type="text" name="orgallatsubtractsum" value="<%= orefund.FOneItem.Fallatsubtractsum %>" size="8" style="text-align:right;background-color:#FFFFFF;border-style:none" readonly> <!-- ������ ���� *-1 �� 13�� ���� *-1-->
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
    		<td>��� ���ʽ�����(����)</td>
    		<td><input type="checkbox" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> name="forcecouponreturn" <% if (orefund.FOneItem.Frefundfixedcouponsum <> 0) then %>checked<% end if %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
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
    		<td>��� ���ϸ���</td>
    		<td><input type="checkbox" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> name="forcemileagereturn" <% if (orefund.FOneItem.Frefundmileagesum <> 0) then %>checked<% end if %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
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
    		<td>��� Giftī��</td>
    		<td><input type="checkbox" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> name="forcegiftcardreturn" <% if (orefund.FOneItem.Frefundgiftcardsum <> 0) then %>checked<% end if %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
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
    		<td>��� ��ġ��</td>
    		<td><input type="checkbox" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> name="forcedepositreturn" <% if (orefund.FOneItem.Frefunddepositsum <> 0) then %>checked<% end if %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
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
    		<td>�ǰ����ݾ�</td>
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
    		<td rowspan="2"><font color="red">��ǰ ��ۺ�</font></td>
    		<td>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpay" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)" <% if (orefund.FOneItem.Frefunddeliverypay <= -4000) then response.write "checked" %> >
        		��ü�պ���ۺ�
        		<!-- ���� ��� ��ۺ� �������� ���� -->
        		<br>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpayHalf" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)"  <% if (orefund.FOneItem.Frefunddeliverypay < 0) and (orefund.FOneItem.Frefunddeliverypay > -4000) then response.write "checked" %> >
        		ȸ����ۺ�
        		<br>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpayZero" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)">
        		��������
    		</td>
    		<td align="center"></td>
    		<td align="right"><input class="text" type="text" name="refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay %>" size="8" style="text-align:right" style="text-align:right" onBlur="CalculateAndApplyItemCostSum(frmaction);"></td>
    	    <td></td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25" <% if (Not IsCSReturnProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td colspan="4" height="70">
        		&nbsp; * �ܼ����� + �귣����ü��ǰ : ��ü�պ���ۺ�<br>
        		&nbsp; * ��Ÿ �ܼ����ɹ�ǰ : ȸ����ۺ�<br>
        		&nbsp; * ��ǰ�ҷ���ǰ �� : ��������<br>
        		&nbsp; * ��ۺ�� <font color="red">�귣�庰 ��ۺ�</font>�� ����
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25" <% if (TRUE) or (Not IsCSCancelProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td rowspan="2"><font color="red">�߰� ��ۺ�</font></td>
    		<td></td>
    		<td></td>
    		<td align="right"><input class="text_ro" type="text" name="adddeliverypay" value="0" size="8" style="text-align:right" style="text-align:right" ></td>
    	    <td></td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25" <% if (TRUE) or (Not IsCSCancelProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td colspan="4" height="40">
        		&nbsp; �ܼ����� + ��ü���ǹ�� + ������ + ����(����)��ۻ�ǰX<br>
        		&nbsp; + ���������� ���Ϸ� ��ǰ���
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25">
    		<td>��Ÿ�����ݾ�
    		<% if (IsTicketOrder) then %>
    		<br><strong>(��� ������)</strong>
    		<% end if %>
    		</td>
    		<% if (IsTicketOrder) then %>
    		<td colspan="2">
        		<input type="radio" name="tRefundPro" value="0" onClick="calcuTicketCancelCharge(this);" > ����
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
            <td>�Ѿ�/��Ҿ�</td>
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

	<tr bgcolor="FFFFFF" ><td align="center" height="30">��� ���� ���°� �ƴմϴ�.</td></tr>

	<% end if %>

<% end if %>