
<% if IsCSItemExchangeCustomerBeasongPayNeeded(divcd) then %>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�߰���ۺ�(����)</td>
	    	    <td>
	    	    	<input type="text" class="text_ro" name="add_customeraddbeasongpay" value="<%= ocsaslist.FOneItem.Fcustomeraddbeasongpay %>" size="20" ReadOnly >
	    	    	&nbsp;
		    	    <select class="select" name="add_customeraddmethod" class="text">
			    	    <option value="">����
			    	    <option value="1" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "1") then %>selected<% end if %>>�ڽ�����
			    	    <option value="2" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "2") then %>selected<% end if %>>�ù�� ���δ�
			    	    <option value="5" <% if (ocsaslist.FOneItem.Fcustomeraddmethod = "5") then %>selected<% end if %>>��Ÿ
		    	    </select>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�߰���ۺ�(Ȯ��)</td>
	    	    <td>
					<% if IsStatusEdit and (divcd = "A111") then %>
						<input type="text" class="text" name="customerrealbeasongpay" value="<%= ocsaslist.FOneItem.Fcustomerrealbeasongpay %>" size="20">
					<% else %>
						<input type="text" class="text_ro" name="customerrealbeasongpay" value="<%= ocsaslist.FOneItem.Fcustomerrealbeasongpay %>" size="20" ReadOnly >
					<% end if %>
	    	    </td>
	    	</tr>
	    	<!--
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">��������</td>
				<input type="hidden" name="orgcustomerreceiveyn" value="<%= ocsaslist.FOneItem.Fcustomerreceiveyn %>">
	    	    <td>
					<input type="radio" class="radio" name="add_customerreceiveyn" value="Y" <% if (ocsaslist.FOneItem.Fcustomerreceiveyn = "Y") then %>checked<% end if %>> �Ա�Ȯ��
					<input type="radio" class="radio" name="add_customerreceiveyn" value="N" <% if (ocsaslist.FOneItem.Fcustomerreceiveyn = "N") then %>checked<% end if %>> Ȯ������
	    	    </td>
	    	</tr>
	    	-->
<% end if %>
