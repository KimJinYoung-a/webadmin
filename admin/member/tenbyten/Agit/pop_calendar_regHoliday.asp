<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� �޷� ���� ���
' History :2017.03.30 ���������
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/tenAgitCalendarCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	dim  sDt ,sHolidayname
	dim idx
	dim cCal
	
	sDt = requestCheckvar(request("ChkStart"),10)
	response.write sDt
	if sDt="" then sDt=date 
	set cCal = new CAgitCalendar
	cCal.FRectDate = sDt
	sHolidayname = cCal.fnGetHolidayname
	set cCal = nothing 
%>
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--
	// ��� Ȯ�� �� ó��
	function jsSubmit()	{
				
		if(document.frm.ChkStart.value == "") {
			alert("��¥�� �������ּ���.");
			document.frm.ChkStart.focus();
			return ;
		}
		
		if(document.frm.sHnm.value == "") {
			alert("���ϸ��� �Է����ּ���.");
			document.frm.sHolidayname.focus();
			return  ;
		}
 
 if (confirm("������ ����Ͻðڽ��ϱ�? ������ ������ ��ϵǾ� �־����� ���� �������� ����˴ϴ�.")){
 	document.frm.action="tenbyten_agit_Process.asp";
 	document.frm.submit();
	}
 }
 
	 function jsSetHnm(){ 
	 	 document.frm.action="pop_calendar_regHoliday.asp";
	 	 document.frm.submit();
	 }

	//��üó��
	function delBook() {
		if(confirm("�� ���೻���� �����Ͻðڽ��ϱ�?"))	{
			frm.mode.value = "del";
			frm.submit();
		}
	}
//-->
</script>
<form name="frm" method="POST" action="">
<input type="hidden" name="mode" value="cal">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><b>�ٹ����� �޷� ���ϵ��</b><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090"> 
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>��¥</b></td>
			<td style="line-height:18px;">
				<input id="ChkStart" name="ChkStart" value="<%=sDt%>" class="text" size="10" maxlength="10" onChange="jsSetHnm();"/><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="ChkStart_trigger" border="0" style="cursor:pointer" align="absmiddle" />										    	
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "ChkStart", trigger    : "ChkStart_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							//CAL_End.args.min = date;
							//CAL_End.redraw();
							jsSetHnm();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});				
				</script>
				
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>���ϸ�</b></td>
			<td><input type="text" name="sHnm" size="30" class="text" value="<%=sHolidayname%>" ></td>
		</tr>		
		<tr bgcolor="#FFFFFF">
			<td colspan="2" align="center">
				<input type="submit" value="�� ��" class="button" style="width:60px;text-align:center;" onClick="jsSubmit()">
				<% if idx<>"" then %>
				&nbsp;&nbsp;&nbsp;<input type="button" value="�� ��" class="button" style="width:60px;text-align:center;" onclick="delBook()">
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form> 
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->