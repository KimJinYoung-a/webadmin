<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Event
' History : 2008.04.02 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->

<%
Dim eCode,egKindCode
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate,stype,estatedesc, ekinddesc, prizeyn
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim strparm
Dim sStateDesc, sEKindDesc
Dim arrEvtStatus, arrEvtKind

	egKindCode = request("evt_code")


dim oip, i

	'--�̺�Ʈ ����
	set oip = new ClsEvent
		oip.FECode = egKindCode	'�̺�Ʈ �ڵ�

		oip.fnGetEventCont	 '�̺�Ʈ ���� ��������
		ekind 		=	oip.FEKind
		ekinddesc	=	oip.FEKindDesc
		eman 		=	oip.FEManager
		escope 		=	oip.FEScope
		ename 		=	db2html(oip.FEName)
		esday 		=	oip.FESDay
		eeday 		=	oip.FEEDay
		epday 		=	oip.FEPDay
		elevel 		=	oip.FELevel
		estate 		=	oip.FEState
		estatedesc 	= oip.FEStateDesc
		eregdate 	=	oip.FERegdate
		prizeyn		=	oip.FPrizeYN
	set oip = nothing
%>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<form name="frm" action="/common/event_prize_process.asp" method="get">
<input type="hidden" name="egKindCode" value="egKindCode">
<input type="hidden" name="eCode" value="<%=egKindCode%>">
<tr>
	<td><!-- ��÷�� ���-->
	<span style="height:25px;padding:10 0 5 0"><img src="/images/icon_arrow_link.gif" align="absmiddle"> ��÷���� : �ѹ� ��ϵ� ��÷�ڴ� ����� �� �����ϴ�. �Է½� ������ �ּ���</span><br>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�ڵ�</td>
			<td width="30%" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=egKindCode%></td>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ��</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ename%></td>
		</tr>
		<tr height="25">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ekinddesc%></td>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=estatedesc%></td>
		</tr>
		<tr height="25">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�Ⱓ</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=esday%>~ <%=eeday%></td>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��÷ ��ǥ��</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=epday%></td>
		</tr>
		<% If prizeyn = "N" Then %>
		<tr height="25">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��÷�ڵ��</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5" colspan="3">
				<input type="button" class="button" value="��÷�ھ���" onclick="frm.submit();">
				* ��÷�ڰ� ���� ��쿡�� �Է��ϼ���.
			</td>
		</tr>
		<% End If %>
		</table>
	</td>
</tr>
</form>

<!-- /��÷�� ���-->
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
