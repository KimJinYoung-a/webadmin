<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Event  
' History : 2008.04.02 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->

<%
Dim eCode,egKindCode
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate,stype
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim strparm
Dim sStateDesc, sEKindDesc
Dim arrEvtStatus, arrEvtKind
	eCode = 4
	egKindCode = request("evt_code")

	
dim oip, i 
	set oip = new cevent_list
	oip.frectevt_code = request("evt_code")
	oip.fevent_oneitem()
%>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<form name="frm" action="/common/event_prize_process.asp" method="get">
<input type="hidden" name="egKindCode" value="<%= oip.foneitem.fevt_code %>">
<input type="hidden" name="eCode" value=4>
<tr>
	<td><!-- ��÷�� ���-->				
	<span style="height:25px;padding:10 0 5 0"><img src="/images/icon_arrow_link.gif" align="absmiddle"> ��÷���� : �ѹ� ��ϵ� ��÷�ڴ� ����� �� �����ϴ�. �Է½� ������ �ּ���</span><br>		
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�ڵ�</td>
				<td width="200" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= oip.foneitem.fevt_code %></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ��</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= oip.foneitem.fevt_name %></td>
			</tr>	
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">culturestation Event</td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�Ⱓ</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=oip.foneitem.fstartdate%> ~ <%=oip.foneitem.fenddate%></td>
			</tr>
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��÷ ��ǥ��</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=oip.foneitem.feventdate%></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��÷�ڵ��</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%=oip.foneitem.fprizeyn%>
					<input type="button" class="button" value="����Y����ȯ�ϱ�" onclick="frm.submit();">
				</td>
			</tr>			
		</table>
	</td>
</tr>
</form>
<tr>
		<td>
		<!-- #include virtual="/admin/eventmanage/common/inc_eventprize.asp"-->	
	</td>
</tr>	
<!-- /��÷�� ���-->
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

