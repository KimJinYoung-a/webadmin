<% if (Not IsStatusFinished) then %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td colspan="4" align="center">

    <% if (IsStatusRegister) then %>

        <% if (IsJupsuProcessAvail) then %>
        	<input class="csbutton" type="button" value=" �� �� " onClick="CsRegProc(frmaction)">
        <% else %>
            <% if JupsuInValidMsg<>"" then %>
            	<font color="red"><%= JupsuInValidMsg %></font>
            	<script language='javascript'>alert('<%= JupsuInValidMsg %>');</script>
				<% if (C_CSPowerUser = True) and (divcd="A008") and (Left(now, 10) = "2013-12-06") and (JupsuInValidMsg = "���Ϸ� ���Ŀ��� ȸ����û/��ǰ���� �� �����մϴ�. - ��� �Ұ��� ") then %>
				<br><br><input class="csbutton" type="button" value=" [�����ڱ���] �� �� " onClick="CsRegProc(frmaction)">(2013-12-06 �ϱ���)
				<% end if %>
            <% end if %>
        <% end if %>

    <% elseif (Not IsStatusFinished) and (ocsaslist.FOneITem.FDeleteyn="N") then %>

        <% if (mode="finishreginfo") then %>
            <% if (divcd="A004") or (divcd="A010") then %>
                <% if (IsFinishProcessAvail) then %>
	                <input id="btnFinishReturn" class="csbutton" type="button" value=" �Ϸ� ó�� (���̳ʽ�/ȯ�ҿ�û ���)" onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
	                <input class="csbutton" type="button" value=" [���̳ʽ�/ȯ�ҿ�û ����] �Ϸ� ó�� " onClick="CsRegFinishProcNoRefund(frmaction)" onFocus="blur()" disabled>
		        <% else %>
		            <% if FinishInValidMsg<>"" then %>
		            	<font color="red"><%= FinishInValidMsg %></font>
		            	<script language='javascript'>alert('<%= FinishInValidMsg %>');</script>
		            <% end if %>
		        <% end if %>
            <% else %>
		        <% if (IsFinishProcessAvail) then %>
		        	<input class="csbutton" type="button" value=" �Ϸ� ó�� " onClick="CsRegFinishProc(frmaction)" onFocus="blur()" name="finishbutton">
		        <% else %>
		            <% if FinishInValidMsg<>"" then %>
		            	<font color="red"><%= FinishInValidMsg %></font>
		            	<script language='javascript'>alert('<%= FinishInValidMsg %>');</script>
		            <% end if %>
		        <% end if %>
            <% end if %>
        <% else %>
            <% IF (Not (IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""))) then %>
            	ȯ������ �ۼ����̹Ƿ� ���� �Ұ� �մϴ�.
            <% else %>
                <input class="csbutton" type="button" value=" ���� ��� " onClick="CsRegCancelProc(frmaction)" onFocus="blur()">
                <input class="csbutton" type="button" value=" �������� ���� " onClick="CsRegEditProc(frmaction)" onFocus="blur()">
                <% if (IsUpcheConfirmState) then %>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input class="csbutton" type="button" value=" �������·� ���� " onClick="CsUpcheConfirm2RegProc(frmaction)" onFocus="blur()">
                <% end if %>
            <% end if %>
        <% end if %>

        <% if ((divcd="A111") or (divcd="A112")) then %>
        	<input class="csbutton" type="button" value=" ��ȯ�ֹ� ������� " onClick="CsChangeOrderRegProc(frmaction)" onFocus="blur()">
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
					������ҹ�ǰ : <input class="csbutton" type="button" value=" �Ϸ�CS(���,��ǰ) ���� " onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
				<% else %>
					������ҹ�ǰ : <%= DelFinishedCSInValidMsg %>
				<% end if %>
			<% elseif (C_CSPowerUser) then %>
				<% if IsDelFinishedCSAvail = True then %>
					�����ں� : <input class="csbutton" type="button" value=" �Ϸ�CS(���,��ǰ) ���� " onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
				<% else %>
					�����ں� : <%= DelFinishedCSInValidMsg %>
				<% end if %>
			<% else %>
				�����ڹ��� : ���ϿϷ�� �ƴ� ��ǰ��� �Ϸ�CS������ �����ڸ� �����մϴ�.
			<% end if %>
		<% elseif (divcd = "A005") and (C_CSPowerUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			�����ں� : <input class="csbutton" type="button" value="�Ϸ�CS����" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()"> *���޸� ��ȯ�� ������������ Ȯ���ϼ���.
		<% elseif (divcd = "A700") and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			���ϿϷ�� ���� : <input class="csbutton" type="button" value="�Ϸ�CS����" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% elseif C_ADMIN_AUTH then %>
			<input class="csbutton" type="button" value="�Ϸ�CS����[������]" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% end if %>
    </td>
</tr>
</table>

<% end if %>
