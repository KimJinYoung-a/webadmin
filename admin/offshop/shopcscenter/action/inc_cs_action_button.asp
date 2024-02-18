<%
'###########################################################
' Description : 매장 고객센터
' Hieditor : 2012.03.20 한용민 생성
'###########################################################
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<% if (Not IsStatusFinished) then %>
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
	                <input class="csbutton" type="button" value="최종완료처리" onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
	        <% else %>
	            <input class="csbutton" type="button" value="접수취소" onClick="CsRegCancelProc(frmaction)" onFocus="blur()">
	            <input class="csbutton" type="button" value="접수내용수정" onClick="CsRegEditProc(frmaction)" onFocus="blur()">
	            
	            <%
	            '/업체처리완료 상태나 매장처리완료 상태 일때 상태변경 가능
	            if (IsUpcheConfirmState or IsmaejangConfirmState) then
	            %>
	                <input class="csbutton" type="button" value="접수상태로변경" onClick="CsUpcheConfirm2RegProc(frmaction)" onFocus="blur()">
	            <% end if %>
	        <% end if %>
	    <% end if %>
    </td>
</tr>
<% end if %>
</table>