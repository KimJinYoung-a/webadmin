<%

dim IsItemCanceled, IsPossibleModifyItem, IsItemDisabled, IsItemChecked, IsBeasongPay, IsUpcheBeasong
dim OrderDetailState

'��ۿɼ�, ��Ҿʵȹ�ۺ��հ�
dim BaesongMethod, SumBeasongPayNotCanceled, SumItemCostSumNotCanceled

dim strhtmldisabled, strhtmlcancel, strhtmlmodify

%>
<% if (IsDisplayItemList and (ocsOrderDetail.FResultCount > 0)) then %>
<tr >
    <td >
		<% '��ũ��Ʈ�� �ܼ�ȭ�ϱ� ���� �Ʒ��� ���� ���̸� �� ����� �д�.(orderdetailidx �� �Ѱ��� ���� 2���̻��� ��츦 �и��ؼ� �ۼ����� �ʾƵ� �ȴ�.) %>
		<input type="hidden" name="dummystarter" value="">
		<input type="hidden" name="orderdetailidx">
		<input type="hidden" name="odlvtype">
		<input type="hidden" name="itemno">
		<input type="hidden" name="itemcost">
		<input type="hidden" name="allatitemdiscount">
		<input type="hidden" name="percentBonusCouponDiscount">
		<input type="hidden" name="etcDiscountDiscount">
		<input type="hidden" name="regitemno">
		<input type="hidden" name="itemid">
		<input type="hidden" name="makerid">
		<input type="hidden" name="isupchebeasong">
		<input type="hidden" name="orderdetailcurrstate">
		<input type="hidden" name="cancelyn">
		<input type="hidden" name="prevcsreturnfinishno">
		<input type="hidden" name="dummystopper" value="">

		<input type="hidden" name="dummystarter" value="">
		<input type="hidden" name="orderdetailidx">
		<input type="hidden" name="odlvtype">
		<input type="hidden" name="itemno">
		<input type="hidden" name="itemcost">
		<input type="hidden" name="allatitemdiscount">
		<input type="hidden" name="percentBonusCouponDiscount">
		<input type="hidden" name="etcDiscountDiscount">
		<input type="hidden" name="regitemno">
		<input type="hidden" name="itemid">
		<input type="hidden" name="makerid">
		<input type="hidden" name="isupchebeasong">
		<input type="hidden" name="orderdetailcurrstate">
		<input type="hidden" name="cancelyn">
		<input type="hidden" name="prevcsreturnfinishno">
		<input type="hidden" name="dummystopper" value="">
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" width="80">������ǰ</td>
            <td colspan="3" bgcolor="#FFFFFF">
                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
				<tr height="20" align="center" bgcolor="#F4F4F4">
					<td width="30">����</td>
					<td width="50">�̹���</td>
					<td width="30">����</td>
					<td width="50">������</td>
					<td width="50">��ǰ�ڵ�</td>
					<td width="90">�귣��ID</td>
					<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
					<td width="80">����/���ֹ�<br>(������ǰ)</td>
					<td width="60">�ǸŰ�<br>(���ΰ�)</td>
					<td width="60">��ǰ<br>������</td>
					<td width="60">���ʽ�<br>������</td>
					<td width="60">��Ÿ<br />���ΰ�</td>
					<td width="130">��������</td>
				</tr>
	<%

	SumBeasongPayNotCanceled = 0
	SumItemCostSumNotCanceled = 0

	orgitemcostsum = 0
	regitemcostsum = 0

	orgpercentcouponpricesum = 0
	regpercentcouponpricesum = 0

	%>
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



		'----------------------------------------------------------------------
		if (Not IsItemCanceled) then
			if (IsBeasongPay) then
				SumBeasongPayNotCanceled = SumBeasongPayNotCanceled + ocsOrderDetail.FItemList(i).Fitemcost
			else
				SumItemCostSumNotCanceled = SumItemCostSumNotCanceled + ocsOrderDetail.FItemList(i).FItemNo*ocsOrderDetail.FItemList(i).Fitemcost
			end if

			if (Not IsBeasongPay) then
				orgitemcostsum = orgitemcostsum + ocsOrderDetail.FItemList(i).FItemNo*ocsOrderDetail.FItemList(i).Fitemcost
				if (Not IsStatusRegister) then
					regitemcostsum = regitemcostsum + ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister)*ocsOrderDetail.FItemList(i).Fitemcost
				end if
			end if

			'��������
			if Not IsNULL(ocsOrderDetail.FItemList(i).Fbonuscouponidx ) then    '''������ �߰�
				orgpercentcouponpricesum = orgpercentcouponpricesum + ocsOrderDetail.FItemList(i).FItemNo * (ocsOrderDetail.FItemList(i).Fitemcost - ocsOrderDetail.FItemList(i).FdiscountAssingedCost)
			end if

			if (Not IsStatusRegister) and (Not IsNull(ocsOrderDetail.FItemList(i).Fitemcouponidx) or Not IsNull(ocsOrderDetail.FItemList(i).Fbonuscouponidx)) then
				regpercentcouponpricesum = regpercentcouponpricesum + ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) * (ocsOrderDetail.FItemList(i).Fitemcost - ocsOrderDetail.FItemList(i).FdiscountAssingedCost)
			end if
		else
			if (IsStatusFinished and IsCSCancelProcess(divcd) and (ocsOrderDetail.FItemList(i).Fgubun01name <> "")) then
				'CS�Ϸ� ���¿����� ��ҿϷ� ������ ������ �ջ��Ѵ�.
				orgpercentcouponpricesum = orgpercentcouponpricesum + ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) * (ocsOrderDetail.FItemList(i).Fitemcost - ocsOrderDetail.FItemList(i).FdiscountAssingedCost)
				regpercentcouponpricesum = regpercentcouponpricesum + ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) * (ocsOrderDetail.FItemList(i).Fitemcost - ocsOrderDetail.FItemList(i).FdiscountAssingedCost)
			end if
		end if

		'----------------------------------------------------------------------
		strhtmlcancel = ""
		strhtmldisabled = ""
		strhtmlmodify = ""

		if (IsItemCanceled) and (ocsOrderDetail.FItemList(i).Fgubun01name = "") then
			'���
			strhtmlcancel 		= "bgcolor='#CCCCCC' class='gray'"
			strhtmldisabled 	= "disabled"
			strhtmlmodify 		= "style='text-align:center;background-color:#DDDDFF;' readonly"
		elseif (ocsOrderDetail.FItemList(i).Forderitemno < 0) then
			'// ���̳ʽ� ���� ���� �Ұ�(���̳ʽ� �ֹ�, ��ȯ�ֹ�)
			strhtmlcancel 		= "bgcolor='#DDDDFF'"
			strhtmldisabled 	= "disabled"
			strhtmlmodify 		= "style='text-align:center;background-color:#DDDDFF;' readonly"
		else
			if (Not IsStatusRegister) and (ocsOrderDetail.FItemList(i).Fgubun01name <> "") then
				'��������
				strhtmldisabled = "checked disabled"

				if ((Not IsStatusRegister) and (Not IsStatusEdit)) then
					strhtmlmodify = "style='text-align:center;background-color:#DDDDFF;' readonly"

					strhtmldisabled = strhtmldisabled + " disabled"
				elseif (IsStatusEdit) then
					'�����̿ܿ��� ��ǰ�� ������ �� ����.
					strhtmldisabled = strhtmldisabled + " disabled"

					if (InStr(strhtmldisabled, "checked") <= 0) then
						strhtmlmodify = "style='text-align:center;background-color:#DDDDFF;' readonly"
					elseif IsBeasongPay then
						'// 2016-05-25, skyer9
						strhtmlmodify = "style='text-align:center;background-color:#DDDDFF;' readonly"
					end if
				end if
			elseif (Not IsBeasongPay) and (IsCSCancelProcess(divcd) or IsCSReturnProcess(divcd) or (Not IsBeasongPay)) and ((IsPossibleCheckItem(divcd, IsOrderCanceled, IsItemCanceled, OrderMasterState, OrderDetailState, IsUpcheBeasong) = true) or (IsChangeOrder and IsCSReturnProcess(divcd))) then
				'üũ���ɻ�ǰ
				strhtmlcancel = "bgcolor='#FFFFFF'"
				if _
					((IsStatusRegister) and (IsCSCancelProcess(divcd)) and (ckAll = "on")) _
					or _
					((Not IsStatusRegister) and (ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) > 0)) _
				then
					strhtmldisabled = "checked"
				end if

				if ((Not IsStatusRegister) and (Not IsStatusEdit)) then
					strhtmlmodify = "style='text-align:center;background-color:#DDDDFF;' readonly"

					strhtmldisabled = strhtmldisabled + " disabled"
				elseif (IsStatusEdit) then
					'�����̿ܿ��� ��ǰ�� ������ �� ����.
					strhtmldisabled = strhtmldisabled + " disabled"

					if (InStr(strhtmldisabled, "checked") <= 0) then
						strhtmlmodify = "style='text-align:center;background-color:#DDDDFF;' readonly"
					end if
				end if
			else
				'üũ�Ұ���ǰ
				strhtmlcancel 		= "bgcolor='#EEEEEE' class='gray'"
				strhtmldisabled 	= "disabled"
				strhtmlmodify 		= "style='text-align:center;background-color:#DDDDFF;' readonly"
			end if
		end if

		if (IsBeasongPay) then
			if (IsStatusEdit) and (ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) <> 0) then
				BaesongMethod = oordermaster.BeasongCD2Name(ocsOrderDetail.FItemList(i).Fitemoption)
				strhtmldisabled = "checked disabled"
			end if
		end if

		%>
			<%
			distinctid = ocsOrderDetail.FItemList(i).Forderdetailidx
			%>
				<tr align="center" <%= strhtmlcancel %>>
					<td height="25">
						<input type="hidden" name="dummystarter" value="">
						<input type="checkbox" id="orderdetailidx_<%= i %>" name="orderdetailidx" value="<%= ocsOrderDetail.FItemList(i).Forderdetailidx %>" onClick="AnCheckClick(this); CheckSelect(this, <%= LCase(IsBeasongPay) %>);" <%= strhtmldisabled %>>
					</td>
					<td width="50">
						<% if (Not IsBeasongPay) then %>
						<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= ocsOrderDetail.FItemList(i).Fitemid %>" target="_blank"><img src="<%= ocsOrderDetail.FItemList(i).FSmallImage %>" width="50" border="0"></a>
						<% else %>
						��ۺ�
						<% end if %>
					</td>
					<td><font color="<%= ocsOrderDetail.FItemList(i).CancelStateColor %>"><%= ocsOrderDetail.FItemList(i).CancelStateStr %></font></td>
					<td>
						<font color="<%= ocsOrderDetail.FItemList(i).GetStateColor %>"><%= ocsOrderDetail.FItemList(i).GetStateName %></font>
					</td>
					<td>
			<% if ocsOrderDetail.FItemList(i).Fisupchebeasong="Y" then %>
						<font color="red"><%= ocsOrderDetail.FItemList(i).Fitemid %><br>(��ü)</font>
			<% else %>
						<%= ocsOrderDetail.FItemList(i).Fitemid %>
			<% end if %>
					</td>
					<td width="90">
						<acronym title="<%= ocsOrderDetail.FItemList(i).Fmakerid %>">
							<% if ocsOrderDetail.FItemList(i).Fisupchebeasong="Y" then %>
								<a href="javascript:popSimpleBrandInfo('<%= ocsOrderDetail.FItemList(i).Fmakerid %>');"><%= Left(ocsOrderDetail.FItemList(i).Fmakerid,32) %></a>
							<% else %>
								<%= Left(ocsOrderDetail.FItemList(i).Fmakerid,32) %>
							<% end if %>
						</acronym>
					</td>
					<td align="left">
			<% if (Not IsBeasongPay) then %>
						<acronym title="<%= ocsOrderDetail.FItemList(i).FItemName %>"><%= DDotFormat(ocsOrderDetail.FItemList(i).FItemName,64) %></acronym>
				<% if (ocsOrderDetail.FItemList(i).FItemoptionName <> "") then %>
						<br>
						<font color="blue">[<%= ocsOrderDetail.FItemList(i).FItemoptionName %>]</font><br>
				<% end if %>
			<% else %>
						(<%= BaesongMethod %>)
						<% if (IsStatusRegister and Not IsItemCanceled and IsCSCancelProcess(divcd)) then %>
						&nbsp;&nbsp;
						<input type="button" class="button" value="��ۺ����" onClick="CsRegCancelBeasongPayProc(frmaction, <%= ocsOrderDetail.FItemList(i).Forderdetailidx %>);" <% if (ocsOrderDetail.FItemList(i).GetBonusCouponPrice = 0) or (ocsOrderDetail.FItemList(i).Fprevcsreturnfinishno <> 0) or (OrderDetailState = "7") then %>disabled<% end if %> >
						<% end if %>
			<% end if %>
						<div id="causepop_<%= distinctid %>" style="position:absolute;"></div>
					</td>
					<td>
						<input type="text" id="regitemno_<%= i %>" name="regitemno" value="<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) %>" size="2" style="text-align:center" onKeyUp="CheckMaxItemNo(this, <%= ocsOrderDetail.FItemList(i).FItemNo %>);" <%= strhtmlmodify %>>
						/
						<input type="text" id="itemno_<%= i %>" name="itemno" value="<%= ocsOrderDetail.FItemList(i).Forderitemno %>" size="2" style="text-align:center;background-color:#DDDDFF;" readonly>

			<% if ((IsCSReturnProcess(divcd) or IsCSCancelProcess(divcd)) and ocsOrderDetail.FItemList(i).Fprevcsreturnfinishno <> 0) then %>
						<br><b>(<%= ocsOrderDetail.FItemList(i).Fprevcsreturnfinishno %>)</b>
			<% end if %>
					</td>
					<td align="right">
						<input type="hidden" name="itemcost" value="<%= ocsOrderDetail.FItemList(i).Fitemcost %>">
						<% if (Not ocsOrderDetail.FItemList(i).IsOldJumun) then %>
	                    	<span title="<%= ocsOrderDetail.FItemList(i).GetSaleText %>" style="cursor:hand">
	                    	<font color="<%= ocsOrderDetail.FItemList(i).GetSaleColor %>">
	                    		<%= FormatNumber(ocsOrderDetail.FItemList(i).GetSalePrice,0) %>
	                    	</font>
	                    	</span>
                    	<% else %>
                    		----
                    	<% end if %>
					</td>
					<td align="right">
						<!-- ����ī�� ������������ ������ -->
						<% if (oordermaster.FOneItem.FAccountDiv="80") or (ocsOrderDetail.FItemList(i).getAllAtDiscountedPrice<>0) then %>
								<input type="hidden" name="allatitemdiscount" value="<%= ocsOrderDetail.FItemList(i).getAllAtDiscountedPrice %>">
						<% else %>
								<input type="hidden" name="allatitemdiscount" value="0">
						<% end if %>

						<input type="hidden" name="percentBonusCouponDiscount" value="<%= ocsOrderDetail.FItemList(i).GetBonusCouponDiscountPrice %>">
						<input type="hidden" name="etcDiscountDiscount" value="<%= ocsOrderDetail.FItemList(i).GetEtcDiscountDiscountPrice %>">

                    	<% if (IsItemCanceled) then %>
                    		<font color="gray"><%= FormatNumber(ocsOrderDetail.FItemList(i).Fitemcost,0) %></font>
                    	<% elseif ocsOrderDetail.FItemList(i).FItemNo < 1 then %>
                    		<br><font color="red">(<%= FormatNumber(ocsOrderDetail.FItemList(i).GetItemCouponPrice,0) %>)</font>
                    	<% else %>
	                    	<span title="<%= ocsOrderDetail.FItemList(i).GetItemCouponText %>" style="cursor:hand">
	                    	<font color="<%= ocsOrderDetail.FItemList(i).GetItemCouponColor %>">
	                    		<%= FormatNumber(ocsOrderDetail.FItemList(i).GetItemCouponPrice,0) %>
	                    	</font>
	                    	</span>
                    	<% end if %>
					</td>

					<td align="right">
						<!-- %���� or All@���� : ��ǰ�� ��밪. -->
                    	<% if ocsOrderDetail.FItemList(i).FItemNo < 1 then %>
                    		<br><font color="red">(<%= FormatNumber(ocsOrderDetail.FItemList(i).GetBonusCouponPrice,0) %>)</font>
                    	<% else %>
	                    	<span title="<%= ocsOrderDetail.FItemList(i).GetBonusCouponText %>" style="cursor:hand">
	                    	<font color="<%= ocsOrderDetail.FItemList(i).GetBonusCouponColor %>">
	                    		<%= FormatNumber(ocsOrderDetail.FItemList(i).GetBonusCouponPrice,0) %>
	                    	</font>
	                    	</span>
	                    <% end if %>
					</td>

					<td align="right">
	                    <span title="<%= ocsOrderDetail.FItemList(i).GetEtcDiscountText %>" style="cursor:hand">
	                    	<font color="<%= ocsOrderDetail.FItemList(i).GetEtcDiscountColor %>">
	                    		<%= FormatNumber(ocsOrderDetail.FItemList(i).GetEtcDiscountPrice,0) %>
	                    	</font>
	                    </span>
					</td>

					<td align="center">
						<input class="input_01" type="text" name="gubun01name_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun01name %>" size="7" Readonly >
						&gt;
						<input class="input_01" type="text" name="gubun02name_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun02name %>" size="7" Readonly >

			<% if (IsStatusFinished) and ((divcd="A010") or (divcd="A011")) and ((ocsOrderDetail.FItemList(i).Fgubun02="CE01") or (ocsOrderDetail.FItemList(i).Fgubun02="CF02") or (ocsOrderDetail.FItemList(i).Fgubun02="CG02")) then %>
						<!-- �Ϸ�ó�� ���Ŀ� ���������� ��ǰ�ҷ��̸� ǥ�õȴ�.[0] inc_cs_action_change_item_list.asp���� ������ -->
						<br><input type="button" class="button" value="�ҷ����" onClick="popBadItemReg('10<%= CHKIIF(ocsOrderDetail.FItemList(i).FItemid>=1000000,Format00(8,ocsOrderDetail.FItemList(i).FItemid),Format00(6,ocsOrderDetail.FItemList(i).FItemid)) %><%= ocsOrderDetail.FItemList(i).FItemOption %>','<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) %>');">
			<% elseif (IsStatusRegister) or (Not IsNULL(ocsOrderDetail.FItemList(i).Fid)) then %>
						<a href="javascript:divCsAsGubunSelect(frmaction.gubun01_<%= distinctid %>.value, frmaction.gubun02_<%= distinctid %>.value, frmaction.gubun01_<%= distinctid %>.name, frmaction.gubun02_<%= distinctid %>.name, frmaction.gubun01name_<%= distinctid %>.name,frmaction.gubun02name_<%= distinctid %>.name,'frmaction','causepop_<%= distinctid %>')"><div id='causestring_<%= distinctid %>' >����ϱ�</div></a>
			<% end if %>

						<input type="hidden" id="cancelyn_<%= i %>" name="cancelyn" value="<%= ocsOrderDetail.FItemList(i).FCancelyn %>">
						<input type="hidden" name="orderdetailcurrstate" value="<%= ocsOrderDetail.FItemList(i).Forderdetailcurrstate %>">
						<input type="hidden" id="isupchebeasong_<%= i %>" name="isupchebeasong" value="<%= ocsOrderDetail.FItemList(i).Fisupchebeasong %>">
						<input type="hidden" id="itemid_<%= i %>" name="itemid" value="<%= ocsOrderDetail.FItemList(i).Fitemid %>">
						<input type="hidden" id="makerid_<%= i%>" name="makerid" value="<%= ocsOrderDetail.FItemList(i).Fmakerid %>">
						<input type="hidden" name="odlvtype" value="<%= ocsOrderDetail.FItemList(i).Fodlvtype %>">
						<input type="hidden" id="prevcsreturnfinishno_<%= i %>" name="prevcsreturnfinishno" value="<%= ocsOrderDetail.FItemList(i).Fprevcsreturnfinishno %>">
						<input type="hidden" name="gubun01_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun01 %>">
						<input type="hidden" name="gubun02_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun02 %>">
						<input type="hidden" name="dummystopper" value="">
					</td>
				</tr>
	<% next %>
            	<tr bgcolor="FFFFFF" height="25">
            	    <td colspan="7">
						<input type="radio" name="showitemtype" value="" onClick="ShowOnlySelectedItem(frmaction)" <% if (Not IsStatusRegister) then  %>checked<% end if %>> ���� ��ǰ�� ǥ��
            	    	<input type="radio" name="showitemtype" value="" onClick="ShowAllItem(frmaction)" <% if (IsStatusRegister) then  %>checked<% end if %>> ��ü ��ǰ ǥ��
            	    </td>
            	    <td colspan="2">�����ǰ�հ�</td>
            	    <td align="right"><%= FormatNumber(SumItemCostSumNotCanceled, 0) %>��</td>
            	    <td colspan="3"></td>
            	</tr>
            	<tr bgcolor="FFFFFF" height="20">
            	    <td colspan="7">
            	        &nbsp;
            	    </td>
            	    <td align="right" colspan="3">
            	        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
            	        <tr>
            	            <td>���û�ǰ�հ�</td>
            	            <td align="right"><input type="text" name="itemcanceltotal" size="7" readonly style="text-align:right;border: 1px solid #333333;" ></td>
            	        </tr>
            	        </table>
            	    </td>
            	    <td colspan="3"></td>
            	</tr>
            	</table>
            </td>
		</tr>
		</table>
	</td>
</tr>
<% end if %>
