<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/only_sys_cls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%
	Dim cEvent, vEvtCode, vEvtSubj, vEvtStatus, vEvtRealStatus, vEvtSDate, vEvtEDate, vEvtMDName, vEvtIsSale, vEvtIsGift, vEvtSaleCnt, vEvtGiftCnt, vQuery
	vEvtCode = requestCheckVar(Request("evt_code"),6)
	
	IF vEvtCode <> "" Then
		If IsNumeric(vEvtCode) = True Then
			Set cEvent = new cOnlySys
			cEvent.FEvtcode = vEvtCode
			cEvent.fnEventCont
			
			vEvtSubj = cEvent.FEvtSubject
			vEvtStatus = cEvent.FEvtStatus
			vEvtRealStatus = cEvent.FEvtRealStatus
			vEvtSDate = cEvent.FEvtSDate
			vEvtEDate = cEvent.FEvtEDate
			vEvtMDName = cEvent.FEvtMDname
			vEvtIsSale = cEvent.FEvtIsSale
			vEvtIsGift = cEvent.FEvtIsGift
			vEvtSaleCnt = cEvent.FEvtSaleCnt
			vEvtGiftCnt = cEvent.FEvtGiftCnt
			Set cEvent = Nothing
		End If
	End IF

	vQuery = "select * from" & vbCrLf
	vQuery = vQuery & "db_event.dbo.tbl_event" & vbCrLf
	vQuery = vQuery & "where evt_code = '" & vEvtCode & "'" & vbCrLf & vbCrLf
	vQuery = vQuery & "--update db_event.dbo.tbl_event set" & vbCrLf
	vQuery = vQuery & "evt_state = '', evt_enddate = ''" & vbCrLf
	vQuery = vQuery & "where evt_code = '" & vEvtCode & "'" & vbCrLf
%>

<script language="javascript">
function jsEventSearch()
{
	if(frm1.evt_code.value == "")
	{
		alert("�̺�Ʈ �ڵ� �Է��ϼ���.");
		frm1.evt_code.focus();
		return;
	}
	if(isNaN(frm1.evt_code.value))
	{
		alert("�̺�Ʈ �ڵ带 ���ڷθ� �Է��ϼ���.");
		frm1.evt_code.value = "";
		frm1.evt_code.focus();
		return;
	}
	frm1.submit();
}
function jsEventUpdate()
{
	if(frm1.evt_code.value == "")
	{
		alert("�̺�Ʈ �ڵ� �Է��ϼ���.");
		frm1.evt_code.focus();
		return;
	}
	if(isNaN(frm1.evt_code.value))
	{
		alert("�̺�Ʈ �ڵ带 ���ڷθ� �Է��ϼ���.");
		frm1.evt_code.value = "";
		frm1.evt_code.focus();
		return;
	}
	if(frm1.evt_status.value != "" && isNaN(frm1.evt_status.value))
	{
		alert("�̺�Ʈ ������°��� ���ڷθ� �Է��ϼ���.");
		frm1.evt_status.value = "";
		frm1.evt_status.focus();
		return;
	}
	if(frm1.evt_edate.value == "")
	{
		alert("�̺�Ʈ �������� �����ϼ���.");
		return;
	}
	
	if(confirm("�̴�� �����Ͻðڽ��ϱ�?") == true) {
		frm1.method = "post";
		frm1.action = "event_date_update_proc.asp";
		frm1.submit();
	} else {
		return;
	}
}
</script>

<table class="a">
<tr>
	<td>
		<form name="frm1" action="<%=CurrURL%>" method="get">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td colspan="2">
				�̺�Ʈ �ڵ� : <input type="text" name="evt_code" value="<%=vEvtCode%>" maxlength="5">
				<input type="button" class="button" value="�� ��" onClick="jsEventSearch()">
			</td>
		</tr>
		<% If vEvtCode <> "" Then %>
		<tr>
			<td style="padding:15 10 0 0;">
				<b>������°�</b> : <input type="text" name="evt_status" value="" maxlength="1" size="3">(�����Ҷ��� �Է�. <u>�ʿ������ ����.</u> ����Ȱ��� �ٽ� �����Ϸ��� 7)&nbsp;
				<b>������</b> : <input type="text" name="evt_edate" size="10" maxlength=10 readonly value="">
				<a href="javascript:calendarOpen(frm1.evt_edate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
			</td>
			<td style="padding:15 0 0 0;">
				<div id="btn" style="display:block;"><input type="button" value="�ٷκ����ϱ�" onClick="jsEventUpdate()"></div>
			</td>
		</tr>
		<% End If %>
		</table>
		</form>
	</td>
</tr>
</table>

<% If vEvtSubj <> "" Then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap>�̺�Ʈ�ڵ�</td>
  	<td nowrap>�������</td>
  	<td nowrap>�̺�Ʈ��</td>
  	<td width="60">������</td>
  	<td width="60">������</td>
  	<td nowrap>���MD</td>
  	<td nowrap></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td><%=vEvtCode%></td>
	<td><%=vEvtRealStatus%> (<%=fnGetCommCodeArrDesc(fnSetCommonCodeArr("eventstate",False),vEvtStatus)%>)</td>
	<td><%=vEvtSubj%></td>
	<td><%=vEvtSDate%></td>
	<td><%=vEvtEDate%></td>
	<td><%=vEvtMDName%></td>
	<td>
		<%if vEvtIsSale then%>����(<%=vEvtSaleCnt%>)<br><font color="red">������ �ɷ��ֽ��ϴ�. ���� ó���ϼ���.</font><script language="javascript">document.getElementById("btn").style.display="none";</script><%end if%>&nbsp;
		<%if vEvtIsGift then%>����ǰ(<%=vEvtGiftCnt%>)<br><font color="red">������ �ɷ��ֽ��ϴ�. ���� ó���ϼ���.</font><script language="javascript">document.getElementById("btn").style.display="none";</script><%end if%>
	</td>
</tr>
</table>
<br><br>* ��������<br>
<textarea name="" cols="100" rows="15"><%=vQuery%></textarea>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->