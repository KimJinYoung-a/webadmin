<% if (IsDisplayRefundInfo) then %>

	<% if (IsCSUpcheJungsanNeeded(divcd)) then %>

	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�귣��ID</td>
	    	    <td ><input type="text" class="text_ro" name="buf_requiremakerid" value="<%= ocsaslist.FOneItem.Fmakerid %>" size="20" ReadOnly >
	    	    <% if (divcd="A700") or (divcd="A999") then %>
		    	    <!-- ��ü��Ÿ���� -->
		    	    <input type="button" class="button" value="�귣��ID�˻�" onclick="jsSearchBrandID(this.form.name,'buf_requiremakerid');" >
	    	    <% end if %>
	    	    </td>
	    	</tr>

	    	<tr bgcolor="FFFFFF">
	    	    <td width="100"><%= CHKIIF(divcd="A999", "��ü����", "ȸ����ۺ�") %></td>
	    	    <td ><input type="text" class="text_ro" name="buf_refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >��</td>
	    	</tr>
            <% if ExistsCustomerAddPayRegedCSCount > 0 then %>
	    	<tr bgcolor="FFFFFF">
                <td width="100">
                    <font color="red">���߰�����</font>
                </td>
	    	    <td>
                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
                        <tr bgcolor="FFFFFF" align="center">
                            <td>������</td>
                            <td>�귣��</td>
                            <td>����</td>
                            <td>�ֹ���ȣ</td>
                            <td>�����ݾ�</td>
                        </tr>
                        <%
                        for i = 0 to (oOldcsaslist.FResultCount - 1)
                            if (oOldcsaslist.FItemList(i).Fdeleteyn <> "Y") and (oOldcsaslist.FItemList(i).Fdivcd = "A999") then
                        %>
                        <tr bgcolor="#FFFFFF" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= oOldcsaslist.FItemList(i).Fid %>');" style="cursor:hand">
                            <td nowrap width="100"><acronym title="<%= oOldcsaslist.FItemList(i).Fregdate %>"><%= Left(oOldcsaslist.FItemList(i).Fregdate,10) %></acronym></td>
                            <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).Fmakerid %>"><%= Left(oOldcsaslist.FItemList(i).Fmakerid,32) %></acronym></td>
                            <td nowrap width="100"><font color="<%= oOldcsaslist.FItemList(i).GetCurrstateColor %>"><%= oOldcsaslist.FItemList(i).GetCurrstateName %></font></td>
                            <td nowrap><%= oOldcsaslist.FItemList(i).Fpayorderserial %> <%= CHKIIF(oOldcsaslist.FItemList(i).Fpaycancelyn<>"N", "&nbsp;<font color=red>[���]</font>", "") %></td>
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
	    	    <td width="100">�߰�����</td>
	    	    <td >
					<input type="text" class="text" name="add_upchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay %>" size="10" onKeyUp="Change_add_upchejungsandeliverypay(this);">��
					&nbsp;
					<select class="select" name="add_upchejungsancause" class="text" onChange='Change_add_upchejungsancause(this);'>
						<option value="" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="","selected","") %>>��������</option>
						<option value="��ۺ�" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="��ۺ�","selected","") %> >��ۺ�</option>
						<option value="��ǰ���" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="��ǰ���","selected","") %>>��ǰ���</option>
						<option value="������" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="������","selected","") %>>������</option>
						<option value="�����Է�" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"��ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"��ǰ���" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"������" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","selected","") %>>�����Է�</option>
					</select>

					<span name="span_add_upchejungsancauseText" id="span_add_upchejungsancauseText" style='display:<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"��ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"��ǰ���" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"������" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","inline","none") %>'>
						<input type="text" name="add_upchejungsancauseText" value="<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"��ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"��ǰ���" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"������" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"",ocsaslist.FOneItem.Fadd_upchejungsancause,"") %>" size="10" maxlength="16" >
					</span>
					<a href="javascript:clearAddUpchejungsan(frmaction);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">���߰�����ݾ�</td>
	    	    <td ><input type="text" class="text_ro" name="buf_totupchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay + orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >��</td>
	    	</tr>
	<% else %>

	<tr bgcolor="FFFFFF" ><td align="center" height="30">��ü �߰����� ���� ���°� �ƴմϴ�.</td></tr>

	<% end if %>

<% end if %>
