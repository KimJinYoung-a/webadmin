<%
'###########################################################
' Description : ���ڵ� ��� ����Ʈ ���� �Ŵ�
' Hieditor : 2016.12.15 �ѿ�� ����
'###########################################################
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�� ������ ����
		<br>
		* �����԰� :
		<select name="printername" onchange="reg('');">
			<option value="FORMTEC_A4" <% if printername = "FORMTEC_A4" then response.write " selected" %>>���� (�԰�A4)</option>
			<option value="TEC_B-FV4_45x22" <% if printername = "TEC_B-FV4_45x22" then response.write " selected" %>>TEC B-FV4 (�԰�45x22)</option>
			<option value="TEC_B-FV4_35x15" <% if printername = "TEC_B-FV4_35x15" then response.write " selected" %>>TEC B-FV4 (�԰�35x15)</option>

			<% if onoffgubun="OFFLINE" then %>
				<option value="TEC_B-FV4_45x45" <% if printername = "TEC_B-FV4_45x45" then response.write " selected" %>>TEC B-FV4 (�԰�45x45)</option>
			<% end if %>

			<option value="TEC_B-FV4_80x50" <% if printername = "TEC_B-FV4_80x50" then response.write " selected" %>>TEC B-FV4 (�԰�80x50)</option>
			<option value="TTP-243_45x22" <% if printername = "TTP-243_45x22" then response.write " selected" %>>TTP-243 (�԰�45x22)</option>
			<option value="TTP-243_35x15" <% if printername = "TTP-243_35x15" then response.write " selected" %>>TTP-243 (�԰�35x15)</option>

			<% if onoffgubun="OFFLINE" then %>
				<option value="TTP-243_45x45" <% if printername = "TTP-243_45x45" then response.write " selected" %>>TTP-243 (�԰�45x45)</option>
			<% end if %>

			<option value="TTP-243_80x50" <% if printername = "TTP-243_80x50" then response.write " selected" %>>TTP-243 (�԰�80x50)</option>
		</select>
		&nbsp;
		* ǥ�û�ǰ�� :
		<select name="isforeignprint" onchange="reg('');">
			<option value="N" <% if (isforeignprint = "N") then %>selected<% end if %>>������ǰ��</option>
			<option value="Y" <% if (isforeignprint = "Y") then %>selected<% end if %>>�ؿܻ�ǰ��</option>
		</select>
		&nbsp;
		* �ݾ�ǥ�ù�� :
		<select name="printpriceyn" onchange="reg('');">
			<option value="Y" <% if (printpriceyn = "Y") then %>selected<% end if %>>�Һ��ڰ�ǥ��</option>
			<option value="C" <% if (printpriceyn = "C") then %>selected<% end if %>>���ΰ�ǥ��</option>
			<option value="R" <% if (printpriceyn = "R") then %>selected<% end if %>>�ǸŰ�ǥ��</option>
			<option value="S" <% if (printpriceyn = "S") then %>selected<% end if %>>���ñݾ�ǥ��</option>
			<option value="N" <% if (printpriceyn = "N") then %>selected<% end if %>>�ݾ�ǥ�þ���</option>
		</select>
		&nbsp;
		* �귣��ǥ�� :
		<select name="makeriddispyn" onchange="reg('');">
			<option value="Y" <% if (makeriddispyn = "Y") then %>selected<% end if %>>�귣��ǥ��</option>
			<option value="N" <% if (makeriddispyn = "N") then %>selected<% end if %>>�귣��ǥ�þ���</option>
		</select>
		&nbsp;
		* ����ǥ�� :
		<select name="titledispyn" onchange="reg('');">
			<option value="Y" <% if (titledispyn = "Y") then %>selected<% end if %>>����ǥ��</option>
			<option value="N" <% if (titledispyn = "N") then %>selected<% end if %>>����ǥ�þ���</option>
		</select>

        <!--* ���ڵ� ���� �԰� --->
		<% if printername = "TTP-243_45x22" then %>
			<input type="hidden" name="paperwidth" value="45" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="22" size="4" maxlength=9>
		<% elseif printername = "TTP-243_35x15" then %>
			<input type="hidden" name="paperwidth" value="35" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="15" size="4" maxlength=9>
		<% elseif printername = "TTP-243_45x45" then %>
			<input type="hidden" name="paperwidth" value="45" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="45" size="4" maxlength=9>
		<% elseif printername = "TTP-243_80x50" then %>
			<input type="hidden" name="paperwidth" value="80" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="50" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4_45x22" then %>
			<input type="hidden" name="paperwidth" value="450" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="220" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4_35x15" then %>
			<input type="hidden" name="paperwidth" value="350" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="150" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4_45x45" then %>
			<input type="hidden" name="paperwidth" value="450" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="450" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4_80x50" then %>
			<input type="hidden" name="paperwidth" value="800" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="500" size="4" maxlength=9>
		<% end if %>

		<input type="hidden" name="papermargin" value="3" size="4" maxlength=9>

		<script type='text/javascript'>
			jsSetSelectBoxColor();
		</script>
	</td>
</tr>
</table>