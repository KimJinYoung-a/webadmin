<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �̺�Ʈ ��÷�� ���
' History : 2010.03.22 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->

<%
dim evt_code , chkdisp , evt_using , evt_kind , evt_name , evt_startdate ,evt_enddate
dim evt_state , evt_prizedate , opendate ,closedate , brand , isgift ,partMDid ,evt_forward
dim evt_comment , regdate , shopid , sEKindDesc ,sStateDesc
	evt_code= requestCheckVar(Request("evt_code"),10)	'�̺�Ʈ�ڵ�	

	IF evt_code = "" THEN	'�̺�Ʈ �ڵ尪�� ���� ��� back
%>
		<script language="javascript">
		<!--
			alert("���ް��� ������ �߻��Ͽ����ϴ�. �����ڿ��� �������ֽʽÿ�");
			history.back();
		//-->
		</script>
<%	dbget.close()	:	response.End
	END IF	
	
dim cEvtCont
set cEvtCont = new cevent_list
	cEvtCont.frectevt_code = evt_code	'�̺�Ʈ �ڵ�
	
	'//�����ϰ�쿡�� ����
	if evt_code <> "" then
		
	'�̺�Ʈ ���� ��������	
	cEvtCont.fnGetEventCont_off
	evt_kind = cEvtCont.FOneItem.fevt_kind
	evt_name = cEvtCont.FOneItem.fevt_name
	evt_startdate = cEvtCont.FOneItem.Fevt_startdate
	evt_enddate = cEvtCont.FOneItem.Fevt_enddate
	evt_prizedate =	cEvtCont.FOneItem.Fevt_prizedate
	evt_state =	cEvtCont.FOneItem.Fevt_state
	IF datediff("d",now,evt_enddate) <0 THEN evt_state = 9 '�Ⱓ �ʰ��� ����ǥ��
	regdate	= cEvtCont.FOneItem.fevt_regdate
	evt_using = cEvtCont.FOneItem.Fevt_using
	shopid = cEvtCont.FOneItem.fshopid
	sEKindDesc = cEvtCont.FOneItem.fevt_kinddesc
	sStateDesc = cEvtCont.FOneItem.fevt_statedesc
	end if	
%>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
	<tR>
		<td><!-- ��÷�� ���-->				
		<span style="height:25px;padding:10 0 5 0"><img src="/images/icon_arrow_link.gif" align="absmiddle"> ��÷���� : �ѹ� ��ϵ� ��÷�ڴ� ����� �� �����ϴ�. �Է½� ������ �ּ���</span><br>		
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�ڵ�</td>
					<td width="200" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_code%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ��</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_name%></td>
				</tr>	
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sEKindDesc%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�Ⱓ</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_startdate%> ~ <%=evt_enddate%></td>
				</tr>
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sStateDesc%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��÷ ��ǥ��</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_prizedate%></td>
				</tr>			
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<!-- #include virtual="/admin/offshop/event_off/inc_eventprize.asp"-->	
	</td>
</tr>	
<!-- /��÷�� ���-->
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->