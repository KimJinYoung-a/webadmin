
<% if divcd = "A999" then %>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">��ǰ���</td>
	    	    <td>
	    	    	<input type="text" class="text_ro" id="add_customeradditempay" name="add_customeradditempay" value="<%= ocsaslist.FOneItem.Fcustomeradditempay %>" size="20" ReadOnly>
					&nbsp;
					<% if IsStatusRegister then %>
					<input type="radio" name="customerpayordertype" value="B" <%= CHKIIF(ocsaslist.FOneItem.Fcustomerpayordertype="B" or ocsaslist.FOneItem.Fcustomerpayordertype="", "checked", "") %> onClick="CheckSetAddItemPayment(frmaction)"> ��������
					&nbsp;
					<input type="radio" name="customerpayordertype" value="A" <%= CHKIIF(ocsaslist.FOneItem.Fcustomerpayordertype="A", "checked", "") %> onClick="CheckSetAddItemPayment(frmaction)"> ��������
                    &nbsp;
					<input type="radio" name="customerpayordertype" value="N" <%= CHKIIF(ocsaslist.FOneItem.Fcustomerpayordertype="N", "checked", "") %> onClick="CheckSetAddItemPayment(frmaction)" <%= CHKIIF((session("ssBctId") = "hasora") or (session("ssBctId") = "boyishP") or (session("ssBctId") = "oesesang52") or (session("ssBctId") = "rabbit1693"), "", "disabled") %>> �ű��ֹ�
					<% else %>
					* <%= ocsaslist.FOneItem.GetCustomerPayOrderTypeName() %>
                    <input type="hidden" name="customerpayordertype" value="<%= ocsaslist.FOneItem.Fcustomerpayordertype %>">
					<% end if %>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">��ǰ���԰�</td>
	    	    <td>
	    	    	<input type="text" class="text_ro" id="add_customeradditembuypay" name="add_customeradditembuypay" value="<%= ocsaslist.FOneItem.Fcustomeradditembuypay %>" size="20" ReadOnly>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">��ۺ�</td>
	    	    <td>
	    	    	<input type="text" class="text_ro" id="add_customeraddbeasongpay" name="add_customeraddbeasongpay" value="<%= ocsaslist.FOneItem.Fcustomeraddbeasongpay %>" size="20" ReadOnly onKeyUp="CalculateUpcheAddPayment(frmaction);">
					&nbsp;
					<% if IsStatusRegister then %>
					<input type="button" class="button" value="��ۺ�" style="width: 80px;" onClick="CheckSetAddBeasongPayment(frmaction, '1')">
					&nbsp;
					<input type="button" class="button" value="�պ���ۺ�" style="width: 80px;" onClick="CheckSetAddBeasongPayment(frmaction, '2')">
					&nbsp;
					<input type="button" class="button" value="��Ÿ" style="width: 80px;" onClick="CheckSetAddBeasongPayment(frmaction, 'E')">
					<% else %>
					* <font color="red">���� ����</font>���� �����Ұ�
					<% end if %>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�߰�����(����)</td>
	    	    <td>
	    	    	<input type="text" class="text_ro" id="add_customeraddpay_sum" name="add_customeraddpay_sum" value="<%= ocsaslist.FOneItem.Fcustomeradditempay + ocsaslist.FOneItem.Fcustomeraddbeasongpay %>" size="20" ReadOnly onKeyUp="CalculateUpcheAddPayment(frmaction);">
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100" height="25">��������</td>
	    	    <td>
					<input type="radio" name="accountdiv" value="7" checked> ������ �Ա�(�������)
	    	    </td>
	    	</tr>
			<% if IsStatusRegister then %>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�Ա��Ͻ� ����</td>
	    	    <td>
					<select name='acctno' class="select" title="�Ա� ���� ����">
						<option value="">�Ա��Ͻ� ������ �����ϼ���.</option>
						<option value="11">��    ��</option>
						<option value="06">��������</option>
						<option value="20">�츮����</option>
						<option value="26">��������</option>
						<option value="81">�ϳ�����</option>
						<option value="03">�������</option>
						<!-- option value="05">��ȯ���� : ���Ұ�</option -->
						<option value="39">�泲����</option>
						<option value="32">�λ�����</option>
						<!-- option value="31">�뱸����</option  -->
						<option value="71">��ü��</option>
						<option value="07">����</option>
					</select>
					������ : (��)�ٹ�����
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�Ա��ڸ�</td>
	    	    <td>
					<input name="acctname" type="text" maxlength="12" class="txtInp" style="width:200px;" value="<%= oordermaster.FOneItem.FBuyname %>" id="depositName" />
	    	    </td>
	    	</tr>
			<% else %>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100" height="25">�����ֹ���ȣ</td>
	    	    <td>
					<% if (payorderserial <> "") then %>
						<%= payorderserial %>
	                	[<font color="<%= opayordermaster.FOneItem.CancelYnColor %>"><%= opayordermaster.FOneItem.CancelYnName %></font>]
	                	[<font color="<%= opayordermaster.FOneItem.IpkumDivColor %>"><%= opayordermaster.FOneItem.IpkumDivName %></font>]
					<% end if %>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100" height="25">�������</td>
	    	    <td>
					<% if (payorderserial <> "") then %>
					<% if opayordermaster.FOneItem.FAccountDiv="7" then %>
						<% if C_CriticInfoUserLV1 then %>
				    	<%= opayordermaster.FOneItem.FAccountNo %>
				    	&nbsp;
						<% end if %>
				    	<% if opayordermaster.FOneItem.IsDacomCyberAccountPay then %>
					    <font color="red">[����]</font>
					    <% else %>
					    [�Ϲ�]
					    <% end if %>
					<% end if %>
					<% end if %>
	    	    </td>
	    	</tr>
			<% end if %>
<% end if %>
