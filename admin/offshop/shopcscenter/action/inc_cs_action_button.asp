<%
'###########################################################
' Description : ���� ������
' Hieditor : 2012.03.20 �ѿ�� ����
'###########################################################
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<% if (Not IsStatusFinished) then %>
<tr>
    <td colspan="4" align="center">
	    <% if (IsStatusRegister) then %>
	
	        <% if (IsJupsuProcessAvail) then %>
	        	<input class="csbutton" type="button" value=" �� �� " onClick="CsRegProc(frmaction)">
	        <% else %>
	            <% if JupsuInValidMsg<>"" then %>
	            	<font color="red"><%= JupsuInValidMsg %></font>
	            	<script language='javascript'>alert('<%= JupsuInValidMsg %>');</script>
	            <% end if %>
	        <% end if %>
	
	    <% elseif (Not IsStatusFinished) and (ocsaslist.FOneITem.FDeleteyn="N") then %>
	
	        <% if (mode="finishreginfo") then %>
	                <input class="csbutton" type="button" value="�����Ϸ�ó��" onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
	        <% else %>
	            <input class="csbutton" type="button" value="�������" onClick="CsRegCancelProc(frmaction)" onFocus="blur()">
	            <input class="csbutton" type="button" value="�����������" onClick="CsRegEditProc(frmaction)" onFocus="blur()">
	            
	            <%
	            '/��üó���Ϸ� ���³� ����ó���Ϸ� ���� �϶� ���º��� ����
	            if (IsUpcheConfirmState or IsmaejangConfirmState) then
	            %>
	                <input class="csbutton" type="button" value="�������·κ���" onClick="CsUpcheConfirm2RegProc(frmaction)" onFocus="blur()">
	            <% end if %>
	        <% end if %>
	    <% end if %>
    </td>
</tr>
<% end if %>
</table>