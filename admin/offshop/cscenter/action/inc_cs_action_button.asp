<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<% if (Not IsStatusFinished) then %>
<tr>
    <td colspan="4" align="center">
    <%
    'CS 이메일 발송여부(접수일과 처리일의 차이가 3주 초과하는 경우 체크를 해제해둔다.)
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
	        <input type="checkbox" name="csmailsend" value="on" > CS 접수/처리 이메일 발송
	        <font color=red>(필요한경우 체크하세요. 접수일과 처리일의 차이가 3주 초과)</font>
        <% else %>
        	<input type="checkbox" name="csmailsend" value="on" > CS 접수/처리 이메일 발송
        <% end if %>
    <% end if %>
    </td>
</tr>
<tr>
    <td colspan="4" align="center">
	    <% if (IsStatusRegister) then %>
	
	        <% if (IsJupsuProcessAvail) then %>
	        	<input class="csbutton" type="button" value=" 접 수 " onClick="CsRegProc(frmaction)">
	        <% else %>
	            <% if JupsuInValidMsg<>"" then %>
	            	<font color="red"><%= JupsuInValidMsg %></font>
	            	<script language='javascript'>alert('<%= JupsuInValidMsg %>');</script>
	            <% end if %>
	        <% end if %>
	
	    <% elseif (Not IsStatusFinished) and (ocsaslist.FOneITem.FDeleteyn="N") then %>
	
	        <% if (mode="finishreginfo") then %>
	                <input class="csbutton" type="button" value=" 완료 처리 " onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
	        <% else %>
	            <input class="csbutton" type="button" value=" 접수 취소 " onClick="CsRegCancelProc(frmaction)" onFocus="blur()">
	            <input class="csbutton" type="button" value=" 접수내용 수정 " onClick="CsRegEditProc(frmaction)" onFocus="blur()">
	            <% if (IsUpcheConfirmState) then %>	                
	                <input class="csbutton" type="button" value=" 접수상태로 변경 " onClick="CsUpcheConfirm2RegProc(frmaction)" onFocus="blur()">
	            <% end if %>
	        <% end if %>
	    <% end if %>
    </td>
</tr>
<% end if %>
</table>