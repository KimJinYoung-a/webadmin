<% if (Not IsStatusFinished) then %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
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
        	<input type="checkbox" name="csmailsend" value="on" <%= chkIIF(oordermaster.FOneItem.FSiteName="10x10","checked","") %> > CS ����/ó�� �̸��� �߼�
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
				<% if (C_CSPowerUser = True) and (divcd="A008") and (Left(now, 10) = "2013-12-06") and (JupsuInValidMsg = "���Ϸ� ���Ŀ��� ȸ����û/��ǰ���� �� �����մϴ�. - ��� �Ұ��� ") then %>
				<br><br><input class="csbutton" type="button" value=" [CS������] �� �� " onClick="CsRegProc(frmaction)">(2013-12-06 �ϱ���)
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
				<% if IsDelNotFinishedCSAvail then %>
				<input class="csbutton" type="button" value=" ���� ��� " onClick="CsRegCancelProc(frmaction)" onFocus="blur()">
				<% else %>
				<input class="csbutton" type="button" value=" ���� ��� " onClick="alert('<%= DelNotFinishedCSInValidMsg %>')" onFocus="blur()">
				<% end if %>
                <input class="csbutton" type="button" value=" �������� ���� " onClick="CsRegEditProc(frmaction)" onFocus="blur()">
                <% if (IsUpcheConfirmState) then %>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                	<input class="csbutton" type="button" value=" �������·� ���� " onClick="CsUpcheConfirm2RegProc(frmaction)" onFocus="blur()">
					<!--
					<input class="csbutton" type="button" value=" ��ü ��Ȯ�ο�û " onClick="CsUpcheConfirm2ReConfirmProc(frmaction)" onFocus="blur()">
					-->
				<% end if %>
            <% end if %>
        <% end if %>

        <% if ((divcd="A111") or (divcd="A112")) then %>
        	<input class="csbutton" type="button" value=" ��ȯ�ֹ� ������� " onClick="CsChangeOrderRegProc(frmaction)" onFocus="blur()">
        <% end if %>

	<% elseif (Not IsStatusFinished) and (ocsaslist.FOneITem.FDeleteyn="Y") then %>
		<%
		if (_
			(divcd="A001") or (divcd="A002") or _
			(divcd="A200") or (divcd="A009") or _
			(divcd="A006") or (divcd="A060") or _
			(divcd="A005") or (divcd="A700") _
			) then _
			'A001			������߼�
			'A002			���񽺹߼�
			'A200			��Ÿȸ��
			'A009			��Ÿ����
			'A006			�������ǻ���
			'A060			��ü��޹���
			'A700			��ü��Ÿ����
			'A005			�ܺθ�ȯ�ҿ�û
		%>
			<input class="csbutton" type="button" value="�������� ����" onClick="CsRestoreDelProc(frmaction)" onFocus="blur()">
		<% elseif C_CSPowerUser then %>
			<input class="csbutton" type="button" value="[CS������]�������� ����" onClick="CsRestoreDelProc(frmaction)" onFocus="blur()">
		<% elseif C_ADMIN_AUTH then %>
			<input class="csbutton" type="button" value="[������]�������� ����" onClick="CsRestoreDelProc(frmaction)" onFocus="blur()">
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
        	<input class="csbutton" type="button" value="[������]�Ϸ�CS �Ϸ����� ��ȯ" onClick="CsFinishToJupsu(frmaction)" onFocus="blur()">
        <% end if %>
		<% if (divcd="A004") or (divcd="A010") or (divcd="A008") then %>
        	<% if (divcd="A010") and (ocsaslist.FOneItem.Fextsitename = "10x10_cs") then %>
        		<% if C_ADMIN_AUTH then %>
					<input class="csbutton" type="button" value="[������]�Ϸ�CS ����" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
        		<% else %>
        		�����Ұ� : <%= DelFinishedCSInValidMsg %>
        		<% end if %>
			<% elseif (HasAuthTodayDelCancelReturn) then %>
				<% if IsDelFinishedCSAvail = True and not(C_CSOutsourcingPowerUser) then %>
					������ҹ�ǰ : <input class="csbutton" type="button" value="�Ϸ�CS(���,��ǰ)����" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
				<% else %>
					������ҹ�ǰ : <%= DelFinishedCSInValidMsg %>
				<% end if %>
			<% elseif C_ADMIN_AUTH or (C_CSPowerUser) or C_CSpermanentUser then %>
        		<% if C_ADMIN_AUTH then %>
        			<input class="csbutton" type="button" value="[������]�Ϸ�CS(���,��ǰ)����" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
				<% elseif IsDelFinishedCSAvail = True then %>
					CS�����ں� : <input class="csbutton" type="button" value="�Ϸ�CS(���,��ǰ)����" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
				<% else %>
					CS�����ں� : <%= DelFinishedCSInValidMsg %>
				<% end if %>
			<% else %>
				�����ڹ��� : ���ϿϷ�� �ƴ� ��ǰ��� �Ϸ�CS ������ �����ڸ� �����մϴ�.
			<% end if %>
		<% elseif (divcd = "A005") and (C_CSpermanentUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			<input class="csbutton" type="button" value="�Ϸ�CS ����" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()"> *���޸� ��ȯ�� ������������ Ȯ���ϼ���.
		<% elseif (divcd = "A700" or divcd = "A100") and not(C_CSOutsourcingPowerUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			����Ϸ�� ���� : <input class="csbutton" type="button" value="�Ϸ�CS ����" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% elseif (divcd = "A011" or divcd = "A012" or divcd = "A111" or divcd = "A112") and not(C_CSOutsourcingPowerUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			�±�ȯȸ�� ����Ϸ�� ���� : <input class="csbutton" type="button" value="�±�ȯȸ�� �Ϸ�CS ���� " onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% elseif (divcd = "A000") and not(C_CSOutsourcingPowerUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			��ȯ��� ����Ϸ�� ���� : <input class="csbutton" type="button" value="��ȯ��� �Ϸ�CS ���� " onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% elseif (divcd = "A003") and not(C_CSOutsourcingPowerUser) and (Left(ocsaslist.FOneItem.Ffinishdate,7) = Left(Now(),7)) then %>
			ȯ�� ����Ϸ�� ���� : <input class="csbutton" type="button" value="ȯ�� �Ϸ�CS ���� " onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()" disabled> �ӽ÷� ��Ȱ��ȭ!!
		<%
		' cs��� ���� ��ư. ī��,��ü,�޴�����ҿ�û(�����ڰ� pg�� ���� Ȯ���ʿ�. �Ժη� ���� �������� ����. ���.)
		elseif (divcd = "A007") and C_CSPowerUser and (Left(ocsaslist.FOneItem.Ffinishdate,10) = Left(Now(),10)) then
		%>
			<input class="csbutton" type="button" value="[������]ȯ�� �Ϸ�CS ����(���ϿϷ��)" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% elseif C_ADMIN_AUTH then %>
			<input class="csbutton" type="button" value="[������]�Ϸ�CS ����" onClick="CsRegCancelFinishedProc(frmaction)" onFocus="blur()">
		<% end if %>
		<% if C_ADMIN_AUTH then %>
			<input class="csbutton" type="button" value="[������]DELETE" onClick="CsRealDelProc(frmaction)" onFocus="blur()">
		<% end if %>
    </td>
</tr>
</table>

<% elseif IsStatusFinished and (ocsaslist.FOneITem.FDeleteyn="Y") then %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td colspan="4" align="center" height="50">
        <% if C_ADMIN_AUTH then %>
        	<input class="csbutton" type="button" value="[������] ������ �Ϸ�CS ����" onClick="CsRestoreDelProc(frmaction)" onFocus="blur()">
        <% end if %>
        * �ý����� ����
    </td>
</tr>
</table>

<% end if %>
