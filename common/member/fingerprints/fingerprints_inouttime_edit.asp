<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �����ν� ���°���
' Hieditor : 2011.03.22 �ѿ�� ����
'            2012.02.15 ������ - �̴ϴ޷� ��ü
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/member/fingerprints/fingerprints_cls.asp" -->

<%
dim idx , placeid , empno ,YYYYMMDD ,inoutType ,inoutTime ,posIdx ,posDate ,regdate
dim lasteditupdate, isusing ,ofingerprints ,inoutTimes, dispadmin, job_sn, loginjob_sn
	idx = requestCheckVar(request("idx"),10)
	dispadmin = false

set ofingerprints = new cfingerprints_list
	ofingerprints.frectidx = idx
	ofingerprints.ffingerprints_item()
	
	if ofingerprints.ftotalcount >0 then
		idx = ofingerprints.FOneItem.fidx
		placeid = ofingerprints.FOneItem.fplaceid
		empno = ofingerprints.FOneItem.fempno
		YYYYMMDD = ofingerprints.FOneItem.fYYYYMMDD
		inoutType = ofingerprints.FOneItem.finoutType
		inoutTime = ofingerprints.FOneItem.finoutTime
		inoutTime = left(inoutTime,10)
		inoutTimes = Num2Str(Hour(ofingerprints.FOneItem.finoutTime),2,"0","R") & ":" & Num2Str(Minute(ofingerprints.FOneItem.finoutTime),2,"0","R")& ":" & Num2Str(second(ofingerprints.FOneItem.finoutTime),2,"0","R")	
		posIdx = ofingerprints.FOneItem.fposIdx
		posDate = ofingerprints.FOneItem.fposDate
		regdate = ofingerprints.FOneItem.fregdate
		lasteditupdate = ofingerprints.FOneItem.flasteditupdate
		isusing	 = ofingerprints.FOneItem.fisusing
	end if

if inoutTime="" then inoutTime=date
if inoutTimes="" then inoutTimes="00:00:00"	

' �۾����� ��å.
job_sn=replace(replace(getjob_sn(empno, ""),"","0"),"0","9999")		' ���þ����� 9999
job_sn=replace(job_sn,"13","1")		' CFO -> CEO
' �α����ѻ���� ��å.
loginjob_sn=replace(replace(getjob_sn(session("ssBctSn"), ""),"","0"),"0","9999")		' ���þ����� 9999
loginjob_sn=replace(loginjob_sn,"13","1")		' CFO -> CEO

' ������
if C_ADMIN_AUTH then
	dispadmin = true

' ��å : CEO, CFO, ������
elseif loginjob_sn="1" or loginjob_sn="2" or loginjob_sn="3" then
	dispadmin = true
elseif job_sn>loginjob_sn then
	dispadmin = true
end if

%>

<script type='text/javascript'>

	function editfingerprints(){
	
		if (frmfingerprints.empno.value==''){
			alert('�����ȣ�� �Է��� �ּ���');
			frmfingerprints.empno.focus();
			return;
		}
		
		if (frmfingerprints.inoutType.value==''){
			alert('���¸� ������ �ּ���');
			frmfingerprints.inoutType.focus();
			return;
		}
		
		if (frmfingerprints.inoutTime.value==''){
			alert('��¥�� �Է��� �ּ���');
			frmfingerprints.inoutTime.focus();
			return;
		}

		if (frmfingerprints.inoutTimes.value==''){
			alert('��¥�� �ð� ������ �Է��� �ּ���');
			frmfingerprints.inoutTimes.focus();
			return;
		}
				
		if (frmfingerprints.isusing.value==''){
			alert('��뿩�θ� ������ �ּ���');
			frmfingerprints.isusing.focus();
			return;
		}
		
		frmfingerprints.action='/common/member/fingerprints/fingerprints_inouttime_process.asp';
		frmfingerprints.mode.value='fingerprintsedit';
		frmfingerprints.submit();
	}
	
</script>

<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form method="post" name="frmfingerprints">
<input type="hidden" name="mode">
<tr bgcolor="FFFFFF" align="center">
	<td>
		��ȣ
	</td>
	<td align="left">
		<%= idx %>
		<input type="hidden" name="idx" value="<%=idx%>">
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td>
		�����ȣ
	</td>
	<td align="left">
		<input type="text" name="empno" value="<%=empno%>" size=32 maxlength=32>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td>
		����
	</td>
	<td align="left">
		<% DrawinoutType "inoutType",inoutType,"" %>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">	
	<td>
		�����
	</td>
	<td align="left">
		<input id="inoutTime" name="inoutTime" value="<%=inoutTime%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="inoutTime_trigger" border="0" style="cursor:pointer" align="absmiddle" />
    	<input type="text" name="inoutTimes" size="8" maxlength="8" class="text" value="<%=inoutTimes%>">
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "inoutTime", trigger    : "inoutTime_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>	
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td>
		��뿩��
	</td>
	<td align="left">
		<% Drawisusing "isusing",isusing,"" %>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td colspan="2">
		<% if dispadmin then %>
			<input type="button" onclick="editfingerprints();" class="button" value="����">
		<% else %>
			������ ���´� ������ �����ϽǼ� ������, ���κ��� ����ڸ� ������ ���� �մϴ�.
		<% end if %>
	</td>
</tr>
</form>
</table>

<%
set ofingerprints = nothing
%>	
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->