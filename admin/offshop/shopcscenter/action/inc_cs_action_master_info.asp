<%
'###########################################################
' Description : ���� ������
' Hieditor : 2012.03.20 �ѿ�� ����
'###########################################################
%>

<% if (IsDisplayCSMaster = true) then %>
<tr >
    <td >
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
		<tr>
			<td>
				<% getcurrstate_table currstate,divcd %>
			</td>
		</tr>
		</table>
		
		<br>    
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr>
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">��������</td>
            <td bgcolor="#FFFFFF">
		    	<font style='line-height:100%; font-size:15px; color:blue; font-family:����; font-weight:bold'><%= GetCSCommName_off("Z001", divcd) %></font>
		    	&nbsp;
                <% if (Not IsStatusRegister) then %>
			    	<font style='line-height:100%; font-size:15px; color:#CC3333; font-family:����; font-weight:bold'>[<%= ocsaslist.FOneItem.shopGetCurrstateName %>]</font>
			    	<% if ocsaslist.FOneITem.FDeleteyn<>"N" then %>
						<font style='line-height:100%; font-size:15px; color:#FF0000; font-family:����; font-weight:bold'>- ������ ����</font>
			    	<% end if %>
		    	<% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">�ֹ���ȣ<br>(�ϷĹ�ȣ)</td>
            <td bgcolor="#FFFFFF" width="200" >
                <%= oordermaster.FOneItem.forderno %><Br>(<%= oordermaster.FOneItem.fmasteridx %>)
                [<font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">������</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsStatusRegister) then %>
                    <%= session("ssbctid") %>
                <% else %>
                    <%= ocsaslist.FOneItem.Fwriteuser %>
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">�����Ͻ�</td>
            <td bgcolor="#FFFFFF">
                <% if (IsStatusRegister) then %>
                	<%= now() %>
                <% else %>
                	<%= ocsaslist.FOneItem.Fregdate %>
                <% end if %>
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsStatusRegister) then %>
                	<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= GetDefaultTitle_off(divcd,"", masteridx, orderno) %>" size="56" maxlength="56">
                <% else %>
                	<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= ocsaslist.FOneItem.Ftitle %>" size="56" maxlength="56">
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
            <td bgcolor="#FFFFFF">
            	<textarea <% if IsStatusFinishing then response.write "class='textarea_ro' ReadOnly" else response.write "class='textarea'" end if %> name="contents_jupsu" cols="68" rows="6"><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea>
            </td>
        </tr>
        <% if divcd = "A030" or divcd = "A031" then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" rowspan="2">�������</td>
            <td bgcolor="#FFFFFF" colspan=3>
				<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr bgcolor="#FFFFFF">
				    <td width=100 bgcolor="<%= adminColor("topbar") %>">
				    	<font color="red"><%=deliverydivname%></font>
				    </td>
				    <td>
				    	<% if divcd = "A030" then %>
				    		<% DrawdeliveryCombo "deliverybox" ,deliverybox ," onchange='deliverych(this.value,searchfrm)'" ,"'1'" ,"Y" ,"frmaction" %>
				    	<% elseif divcd = "A031" then %>
				    		<%= ocsaslist.FOneItem.Fmakerid %>
				    	<% end if %>
				    </td>
				    <td width=100 bgcolor="<%= adminColor("topbar") %>">*�����θ�</td>
				    <td><input type="text" class="text" name="reqname" value="<%= ocsaslist.FOneItem.FReqName %>"></td>				    
				</tr>
				<tr bgcolor="#FFFFFF">
				    <td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
				    <td><input type="text" class="text" name="reqphone" value="<%= ocsaslist.FOneItem.FReqPhone %>"></td>
				    <td bgcolor="<%= adminColor("topbar") %>">*�ڵ���</td>
				    <td>
				    	<input type="text" class="text" name="reqhp" value="<%= ocsaslist.FOneItem.FReqHp %>">
				    	<input type="button" name="buyhp" class="button" value="SMS" onclick="javascript:PopCSSMSSend_off('<%= ocsaslist.FOneItem.Freqhp %>','<%= ocsaslist.FOneItem.Fmasteridx %>','<%= oordermaster.FOneItem.forderno %>','');">
				    </td>				    
				</tr>
				<tr bgcolor="#FFFFFF">
				    <td valign="top" bgcolor="<%= adminColor("topbar") %>">*�����ּ�</td>
				    <td>
				        <input type="text" class="text" name="reqzipcode" value="<%= ocsaslist.FOneItem.FReqZipCode %>" size="7"  readonly><!-- id="[on,off,7,7][�����ȣ]" -->
				        <input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frmaction','A')">
						<input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('frmaction','A')">
				        <% '<input type="button" class="button" value="�˻�(��)" onClick="PopSearchZipcode('frmaction')"> %>
				        <Br><input type="text" class="text" name="reqzipaddr" id="[on,off,1,64][�ּ�]" size="35" value="<%= ocsaslist.FOneItem.FReqZipAddr %>">
				        <Br><input type="text" class="text" name="reqaddress" id="[on,off,1,200][�ּ�]" size="35" value="<%= ocsaslist.FOneItem.FReqAddress %>">
				    </td>
				    <td bgcolor="<%= adminColor("topbar") %>">��Ÿ����</td>
				    <td>
				        <textarea class="textarea" rows="3" cols="35" name="comment" id="[off,off,off,off][��Ÿ����]"><%= ocsaslist.FOneItem.FComment %></textarea>
					</td>				    
				</tr>
				<tr bgcolor="#FFFFFF">
				    <td bgcolor="<%= adminColor("topbar") %>">�̸���</td>
				    <td colspan=3>
				    	<input type="text" class="text" name="reqemail" value="<%= ocsaslist.FOneItem.freqemail %>">
				    </td>
				</tr>
				</table>
            </td>
        </tr>
		<% end if %>
        <% if (IsStatusFinishing) or (IsUpcheConfirmState) or (IsStatusFinished) then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">ó������</td>
            <td bgcolor="#FFFFFF" colspan=3>
	            <% if (IsUpcheConfirmState) then %>
	            	<textarea class='textarea_ro' readOnly name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
	            <% else %>
	            	<textarea class='textarea' name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
	            <% end if %>
            </td>
            <!--<td bgcolor="<%= adminColor("pink") %>" align="center">ó������<br>������<br>�����Է�</td>
            <td bgcolor="#FFFFFF">
            	<table border="0" cellspacing="0" cellpadding="0" class="a" valign="top">
            	<tr>
				    <td>
				    	<input class="text" type="text" name="opentitle" value="<%'= ocsaslist.FOneItem.Fopentitle %>" size="48" maxlength="60" readonly>
				    </td>
				</tr>
				<tr>
				    <td>
				    	<textarea class="textarea" name="opencontents" cols="48" rows="5" readonly><%'= ocsaslist.FOneItem.Fopencontents %></textarea>
				    </td>
				</tr>
				</table>
			</td>-->
        </tr>
        <% end if %>

        </table>
	</td>
</tr>
<% end if %>
