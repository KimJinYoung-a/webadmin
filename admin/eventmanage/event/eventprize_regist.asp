<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/eventprize_regist.asp
' Description :  �̺�Ʈ ��÷�� ���
' History : 2007.02.13 ������ ����
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
'--------------------------------------------------------
' ��������
'--------------------------------------------------------
Dim eCode,egKindCode
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate,stype
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim strparm
Dim sStateDesc, sEKindDesc
Dim arrEvtStatus, arrEvtKind
eCode = Request("eC")
	
	IF eCode = "" THEN	'�̺�Ʈ �ڵ尪�� ���� ��� back
%>
		<script language="javascript">
		<!--
			alert("���ް��� ������ �߻��Ͽ����ϴ�. �����ڿ��� �������ֽʽÿ�");
			history.back();
		//-->
		</script>
<%	dbget.close()	:	response.End
	END IF	
'--------------------------------------------------------
' �̺�Ʈ ������ ��������
'--------------------------------------------------------
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
	
	cEvtCont.fnGetEventCont	 '�̺�Ʈ ���� ��������
	ekind		= cEvtCont.FEKind 
	sEKindDesc 	= cEvtCont.FEKindDesc 
	eman 		= cEvtCont.FEManager 
	escope 		= cEvtCont.FEScope 
	ename 		= db2html(cEvtCont.FEName)
	esday 		= cEvtCont.FESDay
	eeday 		= cEvtCont.FEEDay
	epday 		= cEvtCont.FEPDay
	elevel 		= cEvtCont.FELevel
	estate 		= cEvtCont.FEState
	sStateDesc  = cEvtCont.FEStateDesc
	eregdate 	= cEvtCont.FERegdate 
	set cEvtCont = nothing
%>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
	<tR>
		<td><!-- ��÷�� ���-->				
		<span style="height:25px;padding:10 0 5 0"><img src="/images/icon_arrow_link.gif" align="absmiddle"> ��÷���� : �ѹ� ��ϵ� ��÷�ڴ� ����� �� �����ϴ�. �Է½� ������ �ּ���</span><br>		
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�ڵ�</td>
					<td width="200" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=eCode%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ��</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ename%></td>
				</tr>	
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sEKindDesc%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�Ⱓ</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=esday%> ~ <%=eeday%></td>
				</tr>
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sStateDesc%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��÷ ��ǥ��</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=epday%></td>
				</tr>			
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<!-- #include virtual="/admin/eventmanage/common/inc_eventprize.asp"-->	
	</td>
</tr>	
<!-- /��÷�� ���-->
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->