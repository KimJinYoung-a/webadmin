<% if (Not IsStatusFinished) then %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
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
        	<input type="checkbox" name="csmailsend" value="on" <%= chkIIF(oordermaster.FOneItem.FSiteName="10x10","checked","") %> > CS 접수/처리 이메일 발송
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
				<% if (C_CSPowerUser = True) and (divcd="A008") and (Left(now, 10) = "2013-12-06") and (JupsuInValidMsg = "출고완료 이후에는 회수요청/반품접수 만 가능합니다. - 취소 불가능 ") then %>
				<br><br><input class="csbutton" type="button" value=" [CS관리자] 접 수 " onClick="CsRegProc(frmaction)">(2013-12-06 일까지)
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
				<% if IsDelNotFinishedCSAvail then %>
				<input class="csbutton" type="button" value=" 접수 취소 " onClick="CsRegCancelProc(frmaction)" onFocus="blur()">
				<% else %>
				<input class="csbutton" type="button" value=" 접수 취소 " onClick="alert('<%= DelNotFinishedCSInValidMsg %>')" onFocus="blur()">
				<% end if %>
                <input class="csbutton" type="button" value=" 접수내용 수정 " onClick="CsRegEditProc(frmaction)" onFocus="blur()">
                <% if (IsUpcheConfirmState) then %>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                	<input class="csbutton" type="button" value=" 접수상태로 변경 " onClick="CsUpcheConfirm2RegProc(frmaction)" onFocus="blur()">
					<!--
					<input class="csbutton" type="button" value=" 업체 재확인요청 " onClick="CsUpcheConfirm2ReConfirmProc(frmaction)" onFocus="blur()">
					-->
				<% end if %>
            <% end if %>
        <% end if %>

        <% if ((divcd="A111") or (divcd="A112")) then %>
        	<input class="csbutton" type="button" value=" 교환주문 수기생성 " onClick="CsChangeOrderRegProc(frmaction)" onFocus="blur()">
        <% end if %>

	<% elseif (Not IsStatusFinished) and (ocsaslist.FOneITem.FDeleteyn="Y") then %>
		<%
		if (_
			(divcd="A001") or (divcd="A002") or _
			(divcd="A200") or (divcd="A009") or _
			(divcd="A006") or (divcd="A060") or _
			(divcd="A005") or (divcd="A700") _
			) then _
			'A001			누락재발송
			'A002			서비스발송
			'A200			기타회수
			'A009			기타사항
			'A006			출고시유의사항
			'A060			업체긴급문의
			'A700			업체기타정산
			'A005			외부몰환불요청
		%>
			<input class="csbutton" type="button" value="삭제내역 복구" onClick="CsRestoreDelProc(frmaction)" onFocus="blur()">
		<% elseif C_CSPowerUser then %>
			<input class="csbutton" type="button" value="[CS관리자]삭제내역 복구" onClick="CsRestoreDelProc(frmaction)" onFocus="blur()">
		<% elseif C_ADMIN_AUTH then %>
			<input class="csbutton" type="button" value="[관리자]삭제내역 복구" onClick="CsRestoreDelProc(frmaction)" onFocus="blur()">
		<% end if %>
    <% end if %>
    </td>
</tr>
</table>

<% elseif IsStatusFinished and (ocsaslist.FOneITem.FDeleteyn="N") then %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td colspan="4" align="center" height="50">
        <% if C_ADMIN_AUTH then %>
        	<input class="csbutton" type="button" value="[관리자]완료CS 완료이전 전환" onClick="CsFinishToJupsu(frmaction)" onFocus="blur()">
        <% end if %>
		<% if (divcd="A004") or (divcd="A010") or (divcd="A008") then %>
        	<% if (divcd="A010") and (ocsaslist.FOneItem.Fextsitename = "10x10_cs") then %>
        		<% if C_ADMIN_AUTH then %>
					<input class="csbutton" type="button" value="[관리자]완료CS 삭제" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
        		<% else %>
        		삭제불가 : <%= DelFinishedCSInValidMsg %>
        		<% end if %>
			<% elseif (HasAuthTodayDelCancelReturn) then %>
				<% if IsDelFinishedCSAvail = True and not(C_CSOutsourcingPowerUser) then %>
					당일취소반품 : <input class="csbutton" type="button" value="완료CS(취소,반품)삭제" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
				<% else %>
					당일취소반품 : <%= DelFinishedCSInValidMsg %>
				<% end if %>
			<% elseif C_ADMIN_AUTH or (C_CSPowerUser) or C_CSpermanentUser then %>
        		<% if C_ADMIN_AUTH then %>
        			<input class="csbutton" type="button" value="[관리자]완료CS(취소,반품)삭제" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
				<% elseif IsDelFinishedCSAvail = True then %>
					CS관리자뷰 : <input class="csbutton" type="button" value="완료CS(취소,반품)삭제" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
				<% else %>
					CS관리자뷰 : <%= DelFinishedCSInValidMsg %>
				<% end if %>
			<% else %>
				관리자문의 : 당일완료건 아닌 반품취소 완료CS 삭제는 관리자만 가능합니다.
			<% end if %>
		<% elseif (divcd = "A005") and (C_CSpermanentUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			<input class="csbutton" type="button" value="완료CS 삭제" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()"> *제휴몰 고객환불 이전상태인지 확인하세요.
		<% elseif (divcd = "A700" or divcd = "A100") and not(C_CSOutsourcingPowerUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			당월완료건 삭제 : <input class="csbutton" type="button" value="완료CS 삭제" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% elseif (divcd = "A011" or divcd = "A012" or divcd = "A111" or divcd = "A112") and not(C_CSOutsourcingPowerUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			맞교환회수 당월완료건 삭제 : <input class="csbutton" type="button" value="맞교환회수 완료CS 삭제 " onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% elseif (divcd = "A000") and not(C_CSOutsourcingPowerUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			교환출고 당월완료건 삭제 : <input class="csbutton" type="button" value="교환출고 완료CS 삭제 " onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% elseif (divcd = "A003") and not(C_CSOutsourcingPowerUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			환불 당월완료건 삭제 : <input class="csbutton" type="button" value="환불 완료CS 삭제 " onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()" disabled> 임시로 비활성화!!
		<%
		' cs취소 복구 버튼. 카드,이체,휴대폰취소요청(관리자가 pg사 어드민 확인필요. 함부로 권한 열어주지 말것. 사고남.)
		elseif (divcd = "A007") and C_CSPowerUser and (Left(ocsaslist.FOneItem.Ffinishdate,10) = Left(Now(),10)) then
		%>
			<input class="csbutton" type="button" value="[관리자]환불 완료CS 삭제(당일완료건)" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% elseif C_ADMIN_AUTH then %>
			<input class="csbutton" type="button" value="[관리자]완료CS 삭제" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% end if %>
		<% if C_ADMIN_AUTH then %>
			<input class="csbutton" type="button" value="[개발자]DELETE" onClick="CsRealDelProc(frmaction)" onFocus="blur()">
		<% end if %>
    </td>
</tr>
</table>

<% elseif IsStatusFinished and (ocsaslist.FOneITem.FDeleteyn="Y") then %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td colspan="4" align="center" height="50">
        <% if C_ADMIN_AUTH then %>
        	<input class="csbutton" type="button" value="[관리자] 삭제된 완료CS 복구" onClick="CsRestoreDelProc(frmaction)" onFocus="blur()">
        <% end if %>
        * 시스템팀 문의
    </td>
</tr>
</table>

<% end if %>
