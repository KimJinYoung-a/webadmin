<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<% if (Not IsStatusFinished) then %>
<tr>
    <td colspan="4" align="center">
    <%
    'CS �̸��� �߼ۿ���(�����ϰ� ó������ ���̰� 3�� �ʰ��ϴ� ��� üũ�� �����صд�.)
	if (IsStatusRegister or IsStatusFinishing) and _
    		( _
    			(divcd="A000") or (divcd="A001") or _
    			(divcd="A002") or (divcd="A003") or _
    			(divcd="A004") or (divcd="A007") or _
    			(divcd="A008") or (divcd="A010") or _
    			(divcd="A011") _
    		) then
	%>

        <% if ((not (IsStatusRegister)) and (datediff("d", ocsaslist.FOneItem.Fregdate, now()) > 21)) then %>
	        <input type="checkbox" name="csmailsend" value="on" > CS ����/ó�� �̸��� �߼�
	        <font color=red>(�ʿ��Ѱ�� üũ�ϼ���. �����ϰ� ó������ ���̰� 3�� �ʰ�)</font>
        <% else %>
        	<input type="checkbox" name="csmailsend" value="on" > CS ����/ó�� �̸��� �߼�
        <% end if %>
    <% end if %>
    </td>
</tr>
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
	                <input class="csbutton" type="button" value=" �Ϸ� ó�� " onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
	        <% else %>
	            <input class="csbutton" type="button" value=" ���� ��� " onClick="CsRegCancelProc(frmaction)" onFocus="blur()">
	            <input class="csbutton" type="button" value=" �������� ���� " onClick="CsRegEditProc(frmaction)" onFocus="blur()">
	            <% if (IsUpcheConfirmState) then %>	                
	                <input class="csbutton" type="button" value=" �������·� ���� " onClick="CsUpcheConfirm2RegProc(frmaction)" onFocus="blur()">
	            <% end if %>
	        <% end if %>
	    <% end if %>
    </td>
</tr>
<% end if %>
</table>