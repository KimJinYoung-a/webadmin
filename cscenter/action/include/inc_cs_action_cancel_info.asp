<% if (IsDisplayRefundInfo) then %>

	<% if (IsCSCancelInfoNeeded(divcd) or (IsChangeOrder and IsCSReturnProcess(divcd))) then %>
		<%

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
            <td width="120"></td>
            <td>����</td>
            <td width="80">�� ����</td>
            <td width="80">���û�ǰ</td>
            <td width="80">���»�ǰ</td>
        </tr>
			<% if (IsDisplayItemList) and (IsStatusEdit) and (orefund.FOneItem.Frefunditemcostsum<>regitemcostsum) and (regitemcostsum<>0) then %>
            <script language='javascript'>alert('���� �ݾ� ����ġ-������ ���� ���');</script>
            <% end if %>
        <tr bgcolor="FFFFFF" height="25">
    		<td>
    			��ǰ���������
    			<% if IsChangeOrder then %>
    				(�귣��)
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
				��ۺ�
    			<% if IsChangeOrder then %>
    				(�귣��)
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
    		<td>���ʽ�����(����)</td>
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
    		<td>���ʽ�����(����)</td>
    		<td><input type="checkbox" <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %> name="forcecouponreturn" <% if (orefund.FOneItem.Frefundfixedcouponsum <> 0) then %>checked<% end if %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
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
    	<tr bgcolor="FFFFFF" height="25" <% if (Not IsCSCancelProcess(divcd)) and (Not IsCSReturnProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td rowspan="3">
				<% if (IsTravelOrder) then %>
				<font color="red">��� ������</font>
				<% elseif IsCSCancelProcess(divcd) then %>
				<font color="red">�߰� ��ۺ�</font>
				<% else %>
				<font color="red">��ǰ ��ۺ�</font>
				<% end if %>
			</td>
    		<td>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpay" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)"  >
				<% if (IsTravelOrder) then %>
				��Ҽ����� ����
				<% elseif IsCSCancelProcess(divcd) then %>
				----
				<% else %>
				��ü�պ���ۺ�
				<% end if %>
        		<!-- ���� ��� ��ۺ� �������� ���� -->
        		<br>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpayHalf" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)"  >
				<% if (IsTravelOrder) then %>
				�Ϻ�����
				<% elseif IsCSCancelProcess(divcd) then %>
				�߰���ۺ�
				<% else %>
				ȸ����ۺ�
				<% end if %>
        		<br>
        		<input <% if (Not IsPossibleModifyRefundInfo) or (IsStatusRegister and Not IsJupsuProcessAvail) then response.write "disabled" %>  type="checkbox" name="ckreturnpayZero" onClick="CheckDoubleCheck(frmaction,this); CalculateReturnBeasongPay(frmaction); CalculateAndApplyItemCostSum(frmaction)">
        		��������
    		</td>
    		<td align="center" style="line-height: 20px;">
				��������<br><font color="red">������</font>
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
				������ �߰�����
				<% else %>
				��ۺ� �߰�����
				<% end if %>
    		</td>
    		<td colspan="3" height="25">
        		<input type="radio" name="addmethod" value="0" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "0") then %>checked<% end if %> >����
				<input type="radio" name="addmethod" value="3" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "3") then %>checked<% end if %> >������
				<input type="radio" name="addmethod" value="1" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "1") then %>checked<% end if %> >�ڽ�����
				<input type="radio" name="addmethod" value="5" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "5") then %>checked<% end if %> >��Ÿ
    		</td>
    	</tr>
    	<tr bgcolor="FFFFFF" height="25" <% if (Not IsCSCancelProcess(divcd)) and (Not IsCSReturnProcess(divcd)) then %>style='display:none'<% end if %>>
    		<td colspan="4" height="70">
				<% if (IsTravelOrder) then %>
        		&nbsp; * �װ��ǹ��� ���������� ��Ҽ����� ����<br>
        		&nbsp; * ��� 6�������� <font color="red">���ȯ�� �Ұ�</font><br>
				&nbsp; * ������� Ƽ�ϴ� �ΰ�
				<% elseif IsCSCancelProcess(divcd) then %>
				&nbsp; * �ܼ�����<br>
				&nbsp; * �귣�� �ܿ���ǰ�� ��ü���ǹ�ۺ� �̸� + ������ ����/��ǰ ����<br>
				&nbsp; * ��ۺ�� <font color="red">�귣�庰 ��ۺ�</font>�� ����
				<% else %>
        		&nbsp; * �ܼ����� + �귣����ü��ǰ : ��ü�պ���ۺ�<br>
        		&nbsp; * ��Ÿ �ܼ����ɹ�ǰ : ȸ����ۺ�<br>
        		&nbsp; * ��ǰ�ҷ���ǰ �� : ��������<br>
        		&nbsp; * ��ۺ�� <font color="red">�귣�庰 ��ۺ�</font>�� ����
				<% end if %>
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
				<input type="radio" name="tRefundPro" value="2000" <%= chkIIF(mayTicketCancelChargePro=2000,"checked","") %> onClick="calcuTicketCancelCharge(this);" > 2,000��(Ƽ�ϱݾ��� 10%�ѵ�)<br>
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

	            <input type="hidden" name="subtotalprice" value="0">
	            <input type="hidden" name="canceltotal" value="0">
	            <input type="hidden" name="nextsubtotal" value="0">
            </td>
        </tr>

	<% else %>

	<tr bgcolor="FFFFFF" ><td align="center" height="30">��� ���� ���°� �ƴմϴ�.</td></tr>

	<% end if %>

<% end if %>
