<% if (IsDisplayRefundInfo) then %>

	<% if (IsCSUpcheJungsanNeeded(divcd)) then %>

	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">브랜드ID</td>
	    	    <td ><input type="text" class="text_ro" name="buf_requiremakerid" value="<%= ocsaslist.FOneItem.Fmakerid %>" size="20" ReadOnly >
	    	    <% if (divcd="A700") or (divcd="A999") then %>
		    	    <!-- 업체기타정산 -->
		    	    <input type="button" class="button" value="브랜드ID검색" onclick="jsSearchBrandID(this.form.name,'buf_requiremakerid');" >
	    	    <% end if %>
	    	    </td>
	    	</tr>

	    	<tr bgcolor="FFFFFF">
	    	    <td width="100"><%= CHKIIF(divcd="A999", "업체정산", "회수배송비") %></td>
	    	    <td ><input type="text" class="text_ro" name="buf_refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >원</td>
	    	</tr>
            <% if ExistsCustomerAddPayRegedCSCount > 0 then %>
	    	<tr bgcolor="FFFFFF">
                <td width="100">
                    <font color="red">고객추가결제</font>
                </td>
	    	    <td>
                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
                        <tr bgcolor="FFFFFF" align="center">
                            <td>접수일</td>
                            <td>브랜드</td>
                            <td>상태</td>
                            <td>주문번호</td>
                            <td>결제금액</td>
                        </tr>
                        <%
                        for i = 0 to (oOldcsaslist.FResultCount - 1)
                            if (oOldcsaslist.FItemList(i).Fdeleteyn <> "Y") and (oOldcsaslist.FItemList(i).Fdivcd = "A999") then
                        %>
                        <tr bgcolor="#FFFFFF" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= oOldcsaslist.FItemList(i).Fid %>');" style="cursor:hand">
                            <td nowrap width="100"><acronym title="<%= oOldcsaslist.FItemList(i).Fregdate %>"><%= Left(oOldcsaslist.FItemList(i).Fregdate,10) %></acronym></td>
                            <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).Fmakerid %>"><%= Left(oOldcsaslist.FItemList(i).Fmakerid,32) %></acronym></td>
                            <td nowrap width="100"><font color="<%= oOldcsaslist.FItemList(i).GetCurrstateColor %>"><%= oOldcsaslist.FItemList(i).GetCurrstateName %></font></td>
                            <td nowrap><%= oOldcsaslist.FItemList(i).Fpayorderserial %> <%= CHKIIF(oOldcsaslist.FItemList(i).Fpaycancelyn<>"N", "&nbsp;<font color=red>[취소]</font>", "") %></td>
                            <td nowrap><%= FormatNumber((oOldcsaslist.FItemList(i).Fcustomeraddbeasongpay + oOldcsaslist.FItemList(i).Fcustomeradditempay), 0) %></td>
                        </tr>
                        <%
                        	end if
                        next
                        %>
                    </table>
                </td>
	    	</tr>
            <% end if %>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">추가정산</td>
	    	    <td >
					<input type="text" class="text" name="add_upchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay %>" size="10" onKeyUp="Change_add_upchejungsandeliverypay(this);">원
					&nbsp;
					<select class="select" name="add_upchejungsancause" class="text" onChange='Change_add_upchejungsancause(this);'>
						<option value="" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="","selected","") %>>사유선택</option>
						<option value="배송비" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="배송비","selected","") %> >배송비</option>
						<option value="상품대금" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="상품대금","selected","") %>>상품대금</option>
						<option value="도선료" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="도선료","selected","") %>>도선료</option>
						<option value="직접입력" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"배송비" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"상품대금" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"도선료" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","selected","") %>>직접입력</option>
					</select>

					<span name="span_add_upchejungsancauseText" id="span_add_upchejungsancauseText" style='display:<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"배송비" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"상품대금" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"도선료" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","inline","none") %>'>
						<input type="text" name="add_upchejungsancauseText" value="<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"배송비" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"상품대금" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"도선료" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"",ocsaslist.FOneItem.Fadd_upchejungsancause,"") %>" size="10" maxlength="16" >
					</span>
					<a href="javascript:clearAddUpchejungsan(frmaction);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">총추가정산금액</td>
	    	    <td ><input type="text" class="text_ro" name="buf_totupchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay + orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >원</td>
	    	</tr>
	<% else %>

	<tr bgcolor="FFFFFF" ><td align="center" height="30">업체 추가정산 가능 상태가 아닙니다.</td></tr>

	<% end if %>

<% end if %>
