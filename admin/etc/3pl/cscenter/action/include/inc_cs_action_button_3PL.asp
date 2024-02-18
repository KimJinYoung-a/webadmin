<% if (Not IsStatusFinished) then %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td colspan="4" align="center">

    <% if (IsStatusRegister) then %>

        <% if (IsJupsuProcessAvail) then %>
        	<input class="csbutton" type="button" value=" 접 수 " onClick="CsRegProc(frmaction)">
        <% else %>
            <% if JupsuInValidMsg<>"" then %>
            	<font color="red"><%= JupsuInValidMsg %></font>
            	<script language='javascript'>alert('<%= JupsuInValidMsg %>');</script>
				<% if (C_CSPowerUser = True) and (divcd="A008") and (Left(now, 10) = "2013-12-06") and (JupsuInValidMsg = "출고완료 이후에는 회수요청/반품접수 만 가능합니다. - 취소 불가능 ") then %>
				<br><br><input class="csbutton" type="button" value=" [관리자권한] 접 수 " onClick="CsRegProc(frmaction)">(2013-12-06 일까지)
				<% end if %>
            <% end if %>
        <% end if %>

    <% elseif (Not IsStatusFinished) and (ocsaslist.FOneITem.FDeleteyn="N") then %>

        <% if (mode="finishreginfo") then %>
            <% if (divcd="A004") or (divcd="A010") then %>
                <% if (IsFinishProcessAvail) then %>
	                <input id="btnFinishReturn" class="csbutton" type="button" value=" 완료 처리 (마이너스/환불요청 등록)" onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
	                <input class="csbutton" type="button" value=" [마이너스/환불요청 없는] 완료 처리 " onClick="CsRegFinishProcNoRefund(frmaction)" onFocus="blur()" disabled>
		        <% else %>
		            <% if FinishInValidMsg<>"" then %>
		            	<font color="red"><%= FinishInValidMsg %></font>
		            	<script language='javascript'>alert('<%= FinishInValidMsg %>');</script>
		            <% end if %>
		        <% end if %>
            <% else %>
		        <% if (IsFinishProcessAvail) then %>
		        	<input class="csbutton" type="button" value=" 완료 처리 " onClick="CsRegFinishProc(frmaction)" onFocus="blur()" name="finishbutton">
		        <% else %>
		            <% if FinishInValidMsg<>"" then %>
		            	<font color="red"><%= FinishInValidMsg %></font>
		            	<script language='javascript'>alert('<%= FinishInValidMsg %>');</script>
		            <% end if %>
		        <% end if %>
            <% end if %>
        <% else %>
            <% IF (Not (IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""))) then %>
            	환불파일 작성중이므로 수정 불가 합니다.
            <% else %>
                <input class="csbutton" type="button" value=" 접수 취소 " onClick="CsRegCancelProc(frmaction)" onFocus="blur()">
                <input class="csbutton" type="button" value=" 접수내용 수정 " onClick="CsRegEditProc(frmaction)" onFocus="blur()">
                <% if (IsUpcheConfirmState) then %>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input class="csbutton" type="button" value=" 접수상태로 변경 " onClick="CsUpcheConfirm2RegProc(frmaction)" onFocus="blur()">
                <% end if %>
            <% end if %>
        <% end if %>

        <% if ((divcd="A111") or (divcd="A112")) then %>
        	<input class="csbutton" type="button" value=" 교환주문 수기생성 " onClick="CsChangeOrderRegProc(frmaction)" onFocus="blur()">
        <% end if %>

    <% end if %>
    </td>
</tr>
</table>

<% elseif IsStatusFinished and (ocsaslist.FOneITem.FDeleteyn="N") then %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td colspan="4" align="center" height="50">
		<% if (divcd="A004") or (divcd="A010") or (divcd="A008") then %>
			<% if (HasAuthTodayDelCancelReturn) then %>
				<% if IsDelFinishedCSAvail = True then %>
					당일취소반품 : <input class="csbutton" type="button" value=" 완료CS(취소,반품) 삭제 " onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
				<% else %>
					당일취소반품 : <%= DelFinishedCSInValidMsg %>
				<% end if %>
			<% elseif (C_CSPowerUser) then %>
				<% if IsDelFinishedCSAvail = True then %>
					관리자뷰 : <input class="csbutton" type="button" value=" 완료CS(취소,반품) 삭제 " onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
				<% else %>
					관리자뷰 : <%= DelFinishedCSInValidMsg %>
				<% end if %>
			<% else %>
				관리자문의 : 당일완료건 아닌 반품취소 완료CS삭제는 관리자만 가능합니다.
			<% end if %>
		<% elseif (divcd = "A005") and (C_CSPowerUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			관리자뷰 : <input class="csbutton" type="button" value="완료CS삭제" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()"> *제휴몰 고객환불 이전상태인지 확인하세요.
		<% elseif (divcd = "A700") and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			당일완료건 삭제 : <input class="csbutton" type="button" value="완료CS삭제" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% elseif C_ADMIN_AUTH then %>
			<input class="csbutton" type="button" value="완료CS삭제[관리자]" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% end if %>
    </td>
</tr>
</table>

<% end if %>
