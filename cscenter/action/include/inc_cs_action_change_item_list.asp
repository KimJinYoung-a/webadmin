<% if (IsDisplayChangeItemList) then %>
<tr >
    <td >

		<% if (divcd = "A100") then %><!-- 상품변경 맞교환출고 -->
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" width="80">출고상품</td>
            <td colspan="3" bgcolor="#FFFFFF">
                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
				<tr height="20" align="center" bgcolor="#F4F4F4">
					<td width="30">선택</td>
					<td width="50">이미지</td>
					<td width="30">구분</td>
					<td width="50">현상태</td>
					<td width="50">상품코드</td>
					<td width="90">브랜드ID</td>
					<td>상품명<font color="blue">[옵션명]</font></td>
					<td width="80">접수/원주문</td>
					<td width="60">판매가<br>(할인가)</td>
					<td width="60">쿠폰가</td>
					<td width="130">사유구분</td>
				</tr>
	<% for i = 0 to ocsChangeOrderDetail.FResultCount - 1 %>
		<%
		OrderDetailState = ocsChangeOrderDetail.FItemList(i).ForderDetailcurrstate
		IsUpcheBeasong = (ocsChangeOrderDetail.FItemList(i).Fisupchebeasong = "Y")
		distinctid = ocsChangeOrderDetail.FItemList(i).Fid
		%>
		<% if (IsNull(ocsChangeOrderDetail.FItemList(i).Forderdetailidx) = True) then %>
				<tr align="center" bgcolor='#FFFFFF'>
					<td height="25">
					<input type="hidden" name="dummystarter" value="">
					<!-- cs detail -->
					<input type="checkbox" name="changecsdetailidx" value="<%= ocsChangeOrderDetail.FItemList(i).Fid %>" onClick="AnCheckClick(this); CheckSelect(this);" checked disabled>

					<input type="hidden" name="reforderdetailidx_<%= distinctid %>" value="<%= ocsChangeOrderDetail.FItemList(i).Freforderdetailidx %>">

					<input type="hidden" name="itemid_<%= distinctid %>" value="<%= ocsChangeOrderDetail.FItemList(i).Fitemid %>">
					<input type="hidden" name="itemoption_<%= distinctid %>" value="<%= ocsChangeOrderDetail.FItemList(i).Fitemoption %>">

					</td>
					<td width="50"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= ocsChangeOrderDetail.FItemList(i).Fitemid %>" target="_blank"><img src="<%= ocsChangeOrderDetail.FItemList(i).FSmallImage %>" width="50" border="0"></a></td>
						<input type="hidden" name="gubun01_<%= distinctid %>" value="<%= ocsChangeOrderDetail.FItemList(i).Fgubun01 %>">
						<input type="hidden" name="gubun02_<%= distinctid %>" value="<%= ocsChangeOrderDetail.FItemList(i).Fgubun02 %>">
					<td><font color="<%= ocsChangeOrderDetail.FItemList(i).CancelStateColor %>"><%= ocsChangeOrderDetail.FItemList(i).CancelStateStr %></font></td>
					<td>
						<font color="<%= ocsChangeOrderDetail.FItemList(i).GetStateColor %>"><%= ocsChangeOrderDetail.FItemList(i).GetStateName %></font>
					</td>
					<td>
			<% if ocsChangeOrderDetail.FItemList(i).Fisupchebeasong="Y" then %>
						<font color="red"><%= ocsChangeOrderDetail.FItemList(i).Fitemid %><br>(업체)</font>
			<% else %>
					<%= ocsChangeOrderDetail.FItemList(i).Fitemid %>
			<% end if %>
					</td>
					<td width="90">
						<acronym title="<%= ocsChangeOrderDetail.FItemList(i).Fmakerid %>">
							<% if ocsChangeOrderDetail.FItemList(i).Fisupchebeasong="Y" then %>
								<a href="javascript:popSimpleBrandInfo('<%= ocsChangeOrderDetail.FItemList(i).Fmakerid %>');"><%= Left(ocsChangeOrderDetail.FItemList(i).Fmakerid,32) %></a>
							<% else %>
								<%= Left(ocsChangeOrderDetail.FItemList(i).Fmakerid,32) %>
							<% end if %>
						</acronym>
					</td>
					<td align="left">
						<acronym title="<%= ocsChangeOrderDetail.FItemList(i).FItemName %>"><%= DDotFormat(ocsChangeOrderDetail.FItemList(i).FItemName,64) %></acronym>
			<% if (ocsChangeOrderDetail.FItemList(i).FItemoptionName <> "") then %>
						<br>
						<font color="blue">[<%= ocsChangeOrderDetail.FItemList(i).FItemoptionName %>]</font><br>
			<% end if %>
						<div id="causepop_<%= distinctid %>" style="position:absolute;"></div>
					</td>
					<td>
						<input type="text" name="regitemno" value="<%= ocsChangeOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) %>" size="2" style="text-align:center" onKeyUp="CheckMaxItemNo(this, <%= ocsChangeOrderDetail.FItemList(i).FItemNo %>);" style='text-align:center;background-color:#DDDDFF;' readonly>
						/
						<input type="text" name="itemno" value="<%= ocsChangeOrderDetail.FItemList(i).Forderitemno %>" size="2" style="text-align:center;background-color:#DDDDFF;" readonly>
			<% if IsCSReturnProcess(divcd) and ocsChangeOrderDetail.FItemList(i).Fprevcsreturnfinishno <> 0 then %>
						<br><b>(<%= ocsChangeOrderDetail.FItemList(i).Fprevcsreturnfinishno %>)</b>
			<% end if %>
					</td>
					<input type="hidden" name="itemcost" value="<%= ocsChangeOrderDetail.FItemList(i).Fitemcost %>">
					<td align="right">

					</td>

					<td align="right">

					</td>
					<td align="center">
						<input class="input_01" type="text" name="gubun01name_<%= distinctid %>" value="<%= ocsChangeOrderDetail.FItemList(i).Fgubun01name %>" size="7" Readonly >
						&gt;
						<input class="input_01" type="text" name="gubun02name_<%= distinctid %>" value="<%= ocsChangeOrderDetail.FItemList(i).Fgubun02name %>" size="7" Readonly >

			<% if (IsStatusRegister) or (Not IsNULL(ocsChangeOrderDetail.FItemList(i).Fid)) then %>
						<a href="javascript:divCsAsGubunSelect(frmaction.gubun01_<%= distinctid %>.value, frmaction.gubun02_<%= distinctid %>.value, frmaction.gubun01_<%= distinctid %>.name, frmaction.gubun02_<%= distinctid %>.name, frmaction.gubun01name_<%= distinctid %>.name,frmaction.gubun02name_<%= distinctid %>.name,'frmaction','causepop_<%= distinctid %>')"><div id='causestring_<%= distinctid %>' >등록하기</div></a>
			<% end if %>
					</td>
					<input type="hidden" name="cancelyn" value="<%= ocsChangeOrderDetail.FItemList(i).FCancelyn %>">
					<input type="hidden" name="isupchebeasong" value="<%= ocsChangeOrderDetail.FItemList(i).Fisupchebeasong %>">
					<input type="hidden" name="makerid" value="<%= ocsChangeOrderDetail.FItemList(i).Fmakerid %>">
					<input type="hidden" name="odlvtype" value="<%= ocsChangeOrderDetail.FItemList(i).Fodlvtype %>">
					<input type="hidden" name="prevcsreturnfinishno" value="<%= ocsChangeOrderDetail.FItemList(i).Fprevcsreturnfinishno %>">
					<input type="hidden" name="dummystopper" value="">
				</tr>
		<% end if %>
	<% next %>
            	</table>
            </td>
		</tr>
		</table>
		<% end if %>

		<p>

		<% if (divcd = "A111") or (divcd = "A112") then %><!-- 상품변경 맞교환회수(텐배), 상품변경 맞교환회수(업배) -->
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" width="80">회수상품</td>
            <td colspan="3" bgcolor="#FFFFFF">
                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
				<tr height="20" align="center" bgcolor="#F4F4F4">
					<td width="30">선택</td>
					<td width="50">이미지</td>
					<td width="30">구분</td>
					<td width="50">현상태</td>
					<td width="50">상품코드</td>
					<td width="90">브랜드ID</td>
					<td>상품명<font color="blue">[옵션명]</font></td>
					<td width="80">접수/원주문</td>
					<td width="60">판매가<br>(할인가)</td>
					<td width="60">쿠폰가</td>
					<td width="130">사유구분</td>
				</tr>
	<% for i = 0 to ocsChangeOrderDetail.FResultCount - 1 %>
		<%
		OrderDetailState = ocsChangeOrderDetail.FItemList(i).ForderDetailcurrstate
		IsUpcheBeasong = (ocsChangeOrderDetail.FItemList(i).Fisupchebeasong = "Y")
		distinctid = ocsChangeOrderDetail.FItemList(i).Forderdetailidx
		%>
		<% if (ocsChangeOrderDetail.FItemList(i).Forderdetailidx <> "") then %>
				<tr align="center" bgcolor='#FFFFFF'>
					<td height="25">
					<input type="hidden" name="dummystarter" value="">
					<!-- order detail -->
					<input type="checkbox" name="changecsdetailidx" value="<%= ocsChangeOrderDetail.FItemList(i).Forderdetailidx %>" onClick="AnCheckClick(this); CheckSelect(this);" checked disabled>
					</td>
					<td width="50"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= ocsChangeOrderDetail.FItemList(i).Fitemid %>" target="_blank"><img src="<%= ocsChangeOrderDetail.FItemList(i).FSmallImage %>" width="50" border="0"></a></td>
						<input type="hidden" name="gubun01_<%= distinctid %>" value="<%= ocsChangeOrderDetail.FItemList(i).Fgubun01 %>">
						<input type="hidden" name="gubun02_<%= distinctid %>" value="<%= ocsChangeOrderDetail.FItemList(i).Fgubun02 %>">
					<td><font color="<%= ocsChangeOrderDetail.FItemList(i).CancelStateColor %>"><%= ocsChangeOrderDetail.FItemList(i).CancelStateStr %></font></td>
					<td>
						<font color="<%= ocsChangeOrderDetail.FItemList(i).GetStateColor %>"><%= ocsChangeOrderDetail.FItemList(i).GetStateName %></font>
					</td>
					<td>
			<% if ocsChangeOrderDetail.FItemList(i).Fisupchebeasong="Y" then %>
						<font color="red"><%= ocsChangeOrderDetail.FItemList(i).Fitemid %><br>(업체)</font>
			<% else %>
					<%= ocsChangeOrderDetail.FItemList(i).Fitemid %>
			<% end if %>
					</td>
					<td width="90">
						<acronym title="<%= ocsChangeOrderDetail.FItemList(i).Fmakerid %>">
							<% if ocsChangeOrderDetail.FItemList(i).Fisupchebeasong="Y" then %>
								<a href="javascript:popSimpleBrandInfo('<%= ocsChangeOrderDetail.FItemList(i).Fmakerid %>');"><%= Left(ocsChangeOrderDetail.FItemList(i).Fmakerid,32) %></a>
							<% else %>
								<%= Left(ocsChangeOrderDetail.FItemList(i).Fmakerid,32) %>
							<% end if %>
						</acronym>
					</td>
					<td align="left">
						<acronym title="<%= ocsChangeOrderDetail.FItemList(i).FItemName %>"><%= DDotFormat(ocsChangeOrderDetail.FItemList(i).FItemName,64) %></acronym>
			<% if (ocsChangeOrderDetail.FItemList(i).FItemoptionName <> "") then %>
						<br>
						<font color="blue">[<%= ocsChangeOrderDetail.FItemList(i).FItemoptionName %>]</font><br>
			<% end if %>
						<div id="causepop_<%= distinctid %>" style="position:absolute;"></div>
					</td>
					<td>
						<input type="text" name="regitemno" value="<%= ocsChangeOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) %>" size="2" style="text-align:center" onKeyUp="CheckMaxItemNo(this, <%= ocsChangeOrderDetail.FItemList(i).FItemNo %>);" style='text-align:center;background-color:#DDDDFF;' readonly>
						/
						<input type="text" name="itemno" value="<%= ocsChangeOrderDetail.FItemList(i).Forderitemno %>" size="2" style="text-align:center;background-color:#DDDDFF;" readonly>
			<% if IsCSReturnProcess(divcd) and ocsChangeOrderDetail.FItemList(i).Fprevcsreturnfinishno <> 0 then %>
						<br><b>(<%= ocsChangeOrderDetail.FItemList(i).Fprevcsreturnfinishno %>)</b>
			<% end if %>
					</td>
					<input type="hidden" name="itemcost" value="<%= ocsChangeOrderDetail.FItemList(i).Fitemcost %>">
					<td align="right">
						<% if (Not ocsChangeOrderDetail.FItemList(i).IsOldJumun) then %>
	                    	<span title="<%= ocsChangeOrderDetail.FItemList(i).GetSaleText %>" style="cursor:hand">
	                    	<font color="<%= ocsChangeOrderDetail.FItemList(i).GetSaleColor %>">
	                    		<%= FormatNumber(ocsChangeOrderDetail.FItemList(i).GetSalePrice,0) %>
	                    	</font>
	                    	</span>
                    	<% else %>
                    		----
                    	<% end if %>
					</td>

			<!-- 국민카드 할인으로인해 변경함 -->
			<% if (oordermaster.FOneItem.FAccountDiv="80") or (ocsChangeOrderDetail.FItemList(i).getAllAtDiscountedPrice<>0) then %>
					<input type="hidden" name="allatitemdiscount" value="<%= ocsChangeOrderDetail.FItemList(i).getAllAtDiscountedPrice %>">
			<% else %>
					<input type="hidden" name="allatitemdiscount" value="0">
			<% end if %>

					<input type="hidden" name="percentBonusCouponDiscount" value="<%= ocsChangeOrderDetail.FItemList(i).getPercentBonusCouponDiscountedPrice %>">

					<td align="right">
                    	<% if (IsItemCanceled) then %>
                    		<font color="gray"><%= FormatNumber(ocsChangeOrderDetail.FItemList(i).Fitemcost,0) %></font>
                    	<% elseif ocsChangeOrderDetail.FItemList(i).FItemNo < 1 then %>
                    		<br><font color="red">(<%= FormatNumber(ocsChangeOrderDetail.FItemList(i).GetItemCouponPrice,0) %>)</font>
                    	<% else %>
	                    	<span title="<%= ocsChangeOrderDetail.FItemList(i).GetItemCouponText %>" style="cursor:hand">
	                    	<font color="<%= ocsChangeOrderDetail.FItemList(i).GetItemCouponColor %>">
	                    		<%= FormatNumber(ocsChangeOrderDetail.FItemList(i).GetItemCouponPrice,0) %>
	                    	</font>
	                    	</span>
                    	<% end if %>
				<% if ocsChangeOrderDetail.FItemList(i).FdiscountAssingedCost<>0 and ocsChangeOrderDetail.FItemList(i).FdiscountAssingedCost<>ocsChangeOrderDetail.FItemList(i).Fitemcost then %>
						<!-- %할인 or All@할인 : 반품시 사용값. -->
                    	<% if ocsChangeOrderDetail.FItemList(i).FItemNo < 1 then %>
                    		<br><font color="red">(<%= FormatNumber(ocsChangeOrderDetail.FItemList(i).GetBonusCouponPrice,0) %>)</font>
                    	<% else %>
	                    	<span title="<%= ocsChangeOrderDetail.FItemList(i).GetBonusCouponText %>" style="cursor:hand">
	                    	<font color="<%= ocsChangeOrderDetail.FItemList(i).GetBonusCouponColor %>">
	                    		<br>(<%= FormatNumber(ocsChangeOrderDetail.FItemList(i).GetBonusCouponPrice,0) %>)
	                    	</font>
	                    	</span>
	                    <% end if %>
				<% end if %>
					</td>
					<td align="center">
						<input class="input_01" type="text" name="gubun01name_<%= distinctid %>" value="<%= ocsChangeOrderDetail.FItemList(i).Fgubun01name %>" size="7" Readonly >
						&gt;
						<input class="input_01" type="text" name="gubun02name_<%= distinctid %>" value="<%= ocsChangeOrderDetail.FItemList(i).Fgubun02name %>" size="7" Readonly >

			<% if (IsStatusFinished) and (divcd="A111") and ((ocsChangeOrderDetail.FItemList(i).Fgubun02="CE01") or (ocsChangeOrderDetail.FItemList(i).Fgubun02="CF02") or (ocsOrderDetail.FItemList(i).Fgubun02="CG02")) then %>
						<!-- 완료처리 이후에 사유구분이 상품불량이면 표시된다.[1] inc_cs_action_item_list.asp에도 존재함 -->
						<br><input type="button" class="button" value="불량등록" onClick="popBadItemReg('10<%= CHKIIF(ocsChangeOrderDetail.FItemList(i).FItemid>=1000000,Format00(8,ocsChangeOrderDetail.FItemList(i).FItemid),Format00(6,ocsChangeOrderDetail.FItemList(i).FItemid)) %><%= ocsChangeOrderDetail.FItemList(i).FItemOption %>','<%= (ocsChangeOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister)) %>');">
			<% elseif (IsStatusRegister) or (Not IsNULL(ocsChangeOrderDetail.FItemList(i).Fid)) then %>
						<a href="javascript:divCsAsGubunSelect(frmaction.gubun01_<%= distinctid %>.value, frmaction.gubun02_<%= distinctid %>.value, frmaction.gubun01_<%= distinctid %>.name, frmaction.gubun02_<%= distinctid %>.name, frmaction.gubun01name_<%= distinctid %>.name,frmaction.gubun02name_<%= distinctid %>.name,'frmaction','causepop_<%= distinctid %>')"><div id='causestring_<%= distinctid %>' >등록하기</div></a>
			<% end if %>
					</td>
					<input type="hidden" name="cancelyn" value="<%= ocsChangeOrderDetail.FItemList(i).FCancelyn %>">
					<input type="hidden" name="isupchebeasong" value="<%= ocsChangeOrderDetail.FItemList(i).Fisupchebeasong %>">
					<input type="hidden" name="makerid" value="<%= ocsChangeOrderDetail.FItemList(i).Fmakerid %>">
					<input type="hidden" name="odlvtype" value="<%= ocsChangeOrderDetail.FItemList(i).Fodlvtype %>">
					<input type="hidden" name="prevcsreturnfinishno" value="<%= ocsChangeOrderDetail.FItemList(i).Fprevcsreturnfinishno %>">
					<input type="hidden" name="dummystopper" value="">
				</tr>
		<% end if %>
	<% next %>
            	</table>
            </td>
		</tr>
		</table>
		<% end if %>

	</td>
</tr>
<% end if %>
