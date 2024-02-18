<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����� �޿� �⺻���� ���
' History : 2010.12.23 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenPayCls.asp" -->
<%
Dim intLoop
Dim clsPayForm, spatternname, ipatternseq,part_sn
Dim sempno,susername, susermail, sdirect070, djoinday, blnstatediv, spart_name, sposit_name, sjob_name
Dim startdate, enddate,defaultpay ,foodpay,jobpay ,inBreakTime  ,holidaywdtime	,regdate    ,lastupdate ,adminid,iposit_sn
Dim StartHour(8), StartMinute(8), EndHour(8), EndMinute(8), BreakSHour(8), BreakSMinute(8),  BreakEHour(8), BreakEMinute(8),DutyTime(8) ,NightTime(8),iworktype(8)
Dim totDutyTime,iOverTime, totNightTime, iHolidayTime, totPaySum
Dim sMode
Dim avgWeek,iDefaultPaySeq
iDefaultPaySeq =requestCheckvar(request("iDPS"),10)
ipatternseq 	=  requestCheckvar(request("iPS"),10)
sempno		= requestCheckvar(request("sEN"),14)
sMode ="I"
avgWeek = 4.345238095

IF ipatternseq <> "" THEN
	Set clsPayForm = new CPayForm
	clsPayForm.Fpatternseq= ipatternseq
	clsPayForm.fnGetPayPatternData

	part_sn		= clsPayForm.Fpart_sn
	spatternname	= clsPayForm.Fpatternname
	defaultpay  = clsPayForm.Fdefaultpay
	foodpay	    = clsPayForm.Ffoodpay
	jobpay		= clsPayForm.Fjobpay
	inBreakTime	= clsPayForm.FinBreakTime
	iOverTime	= clsPayForm.FOverTime

	For intLoop = 1 To 7
	StartHour(intLoop) 		= clsPayForm.FStartHour(intLoop)
	StartMinute(intLoop)  	= clsPayForm.FStartMinute(intLoop)
	EndHour(intLoop)       	= clsPayForm.FEndHour(intLoop)
	EndMinute(intLoop)     	= clsPayForm.FEndMinute(intLoop)
	BreakSHour(intLoop)     = clsPayForm.FBreakSHour(intLoop)
	BreakSMinute(intLoop)   = clsPayForm.FBreakSMinute(intLoop)
	BreakEHour(intLoop)     = clsPayForm.FBreakEHour(intLoop)
	BreakEMinute(intLoop)   = clsPayForm.FBreakEMinute(intLoop)
	DutyTime(intLoop)		= clsPayForm.FDutyTime(intLoop)
	NightTime(intLoop)		= clsPayForm.FNightTime(intLoop)
	iworktype(intLoop)		= clsPayForm.Fworktype(intLoop)
	Next

	totDutyTime  	= clsPayForm.FTotDutyTime
	totNightTime	= clsPayForm.FtotNightTime
	totPaySum		= clsPayForm.FTotPaySum

	holidaywdtime	= clsPayForm.Fholidaywdtime
	regdate        	= clsPayForm.Fregdate
	lastupdate     	= clsPayForm.Flastupdate
	adminid        	= clsPayForm.Fadminid
	sMode ="U"
	Set clsPayForm = nothing
 END IF

if defaultpay ="" THEN defaultpay =0
if foodpay ="" THEN foodpay =0
if jobpay ="" THEN jobpay =0
if inBreakTime ="" then inBreakTime = 0
if iOverTime = "" or isNull(iOverTime) THEN iOverTime = 0
IF totDutyTime = "" THEN totDutyTime = 0
IF totNightTime = "" THEN totNightTime = 0
IF totPaySum = "" THEN totPaySum = 0
if part_sn ="" then part_sn = 1

%>
<html>
<head>
<title>������� ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript" src="/js/jsPayCal.js"></script>
<script language="javascript">
<!--
//�� üũ �� submit ó��
	function jsChkform(frm){

		if(frm.part_sn.value ==""){
			 frm.part_sn.value = 1;
		}

		if(frm.sPN.value ==""){
			alert("���ϸ��� �Է����ּ���");
			frm.sPN.focus();
			return false;
		}

		if(!IsDigit(frm.iHP.value)){
			alert("�ñ��� ���ڸ� �Է°����մϴ�.");
			frm.iHP.focus();
			return false;
		}

		if(!IsDigit(frm.iEP.value)){
			alert("�Ĵ�� ���ڸ� �Է°����մϴ�.");
			frm.iEP.focus();
			return false;
		}

		var selWH = 0;
		if(frm.selWH1.value == "3") { selWH = selWH + 1; }
		if(frm.selWH2.value == "3") { selWH = selWH + 1; }
		if(frm.selWH3.value == "3") { selWH = selWH + 1; }
		if(frm.selWH4.value == "3") { selWH = selWH + 1; }
		if(frm.selWH5.value == "3") { selWH = selWH + 1; }
		if(frm.selWH6.value == "3") { selWH = selWH + 1; }
		if(frm.selWH7.value == "3") { selWH = selWH + 1; }


		var totDuty =document.all.totDuty.innerHTML;
		 totDuty = jsFormToTime(totDuty);

		 if(totDuty < 900 && selWH > 0){
		 alert("�ѱٹ� �ð��� 15�ð������� ��� ������ ������ �Ұ����մϴ�.  ");
		 return false;
		 }

		 if(totDuty >= 900 && selWH == 0){
		 alert("�������� �������ּ���");
		 return false;
		 }

		if( selWH > 1){
		alert("������ ������ �Ϸ縸 �����մϴ�.");
		return false;
		}

		return true;

	}

	//����
	function jsDel(){
	 if(confirm("������ �����Ͻðڽ��ϱ�?")){
	 document.frmDel.submit();
	 }
	}

	// ������ �̵�
	function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}
//-->
</script>
</head>
<body leftmargin="10" topmargin="10">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<form name="frmDel" method="post" action="procPayformPattern.asp">
<input type="hidden" name="hidPS" value="<%=ipatternseq%>">
<input type="hidden" name="hidEN" value="<%=sempno%>">
<input type="hidden" name="iDPS" value="<%=idefaultPaySeq%>">
<input type="hidden" name="hidM" value="D">
</form>
<form name="frmPay" method="post" action="procPayformPattern.asp" onsubmit="return jsChkform(this)">
<input type="hidden" name="hidPS" value="<%=ipatternseq%>">
<input type="hidden" name="hidEN" value="<%=sempno%>">
<input type="hidden" name="hidM" value="<%=sMode%>">
<input type="hidden" name="iDPS" value="<%=idefaultPaySeq%>">
<tr>
	<td><strong>�������� ��� ���� ���</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�μ�</td>
			<td bgcolor="#FFFFFF">
			<%=printPartOption("part_sn", part_sn)%>
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">���ϸ�</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sPN" value="<%=spatternname%>">
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ñ�</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iHP" size="10" style="text-align:right;" value="<%=defaultpay%>" onKeyUp="jsSetMonthlypay();"> ��</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�Ĵ�</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iEP" size="10" style="text-align:right;" value="<%=foodpay%>"> ��</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ް�ð�</td>
			<td bgcolor="#FFFFFF"><input type="checkbox" name="blnBT" value="1" onClick="jsSetInBreakTime();" <%IF inBreakTime THEN%>checked<%END IF%>>�ٹ��ð� ���� </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ð��� ����</td>
			<td bgcolor="#FFFFFF"><input type="checkbox" name="blnOT" value="1"  <%IF iOverTime > 0  THEN%>checked<%END IF%> onClick="jsSetOverTime();">����
				<span id="spanOT" style="display:<%IF  iOverTime = 0  THEN%>none<%END IF%>;"><input type="text" size="5" maxlength="10" style="text-align:right;" name="iot" value="<%=iOverTime%>" onKeyUp="jsSetOverTimePay();"> �ð�</span> </td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td><!-- ���Ϻ� �ٹ��ð� ���� -->
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr align="center">
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">����</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">����</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" colspan="2">�ٹ��ð�</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" colspan="2">�ް�ð�</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">�ѱٹ��ð�</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">���޽ð�</td>
		</tr>
		<tr align="center">
			<td  bgcolor="<%= adminColor("tabletop") %>" >����</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" >����</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" >����</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" >����</td>
		</tr>
		<%
		For intLoop = 1 To 7%>
		<tr align="center">
			<td  bgcolor="<%= adminColor("tabletop") %>"><%=fnGetStringWD(intLoop)%></td>
			<td  bgcolor="#FFFFFF">
				<select name="selWH<%=intLoop%>"  onChange="jsSetWH(<%=intLoop%>);">
				<option value="1" <%IF iworktype(intLoop) ="1"  THEN%>selected<%END IF%>>�ٹ���</option>
				<option value="2" <%IF iworktype(intLoop) ="2" THEN%>selected<%END IF%> style="color:blue">��������</option>
				<option value="3" <%IF iworktype(intLoop) ="3" THEN%>selected<%END IF%> style="color:red">������</option>
				<option value="4" <%IF iworktype(intLoop) ="4" THEN%>selected<%END IF%>>��������</option>
				</select>
			</td>
			<td  bgcolor="#FFFFFF">
				<input type="text" name="iSH<%=intLoop%>" value="<%=StartHour(intLoop)%>" size="2" maxlength="2" style="text-align:right" <%IF iworktype(intLoop) =  "3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iSH<%=intLoop%>','iSM<%=intLoop%>',2);">
				:
			 	<input type="text" name="iSM<%=intLoop%>" value="<%=StartMinute(intLoop)%>" size="2"  maxlength="2" style="text-align:right" <%IF iworktype(intLoop) =  "3"  THEN%>disabled<%END IF%>  onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iSM<%=intLoop%>','iEH<%=intLoop%>',2);">
			</td>
			<td  bgcolor="#FFFFFF">
				<input type="text" name="iEH<%=intLoop%>" value="<%=EndHour(intLoop)%>" size="2"  maxlength="2" style="text-align:right" <%IF iworktype(intLoop) =  "3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iEH<%=intLoop%>','iEM<%=intLoop%>',2);">
				:
			 	<input type="text" name="iEM<%=intLoop%>" value="<%=EndMinute(intLoop)%>" size="2"  maxlength="2" style="text-align:right"  <%IF iworktype(intLoop) =  "3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iEM<%=intLoop%>','iBSH<%=intLoop%>',2);">
			</td>
			<td  bgcolor="#FFFFFF">
				<input type="text" name="iBSH<%=intLoop%>" value="<%=BreakSHour(intLoop)%>" size="2"  maxlength="2" style="text-align:right"  <%IF iworktype(intLoop) =  "3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iBSH<%=intLoop%>','iBSM<%=intLoop%>',2);">
				:
			 	<input type="text" name="iBSM<%=intLoop%>" value="<%=BreakSMinute(intLoop)%>" size="2"  maxlength="2" style="text-align:right"  <%IF iworktype(intLoop) =  "3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iBSM<%=intLoop%>','iBEH<%=intLoop%>',2);">
			</td>
			<td  bgcolor="#FFFFFF">
				<input type="text" name="iBEH<%=intLoop%>" value="<%=BreakEHour(intLoop)%>" size="2"  maxlength="2" style="text-align:right"  <%IF iworktype(intLoop) =  "3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);TnTabNumber('iBEH<%=intLoop%>','iBEM<%=intLoop%>',2);">
				:
			 	<input type="text" name="iBEM<%=intLoop%>" value="<%=BreakEMinute(intLoop)%>" size="2"  maxlength="2" style="text-align:right"  <%IF iworktype(intLoop) =  "3"  THEN%>disabled<%END IF%> onKeyUp="jsCalDutyTime(<%=intLoop%>);<%IF (intLoop+1)<8 THEN%>TnTabNumber('iBEM<%=intLoop%>','iSH<%=intLoop+1%>',2);<%END IF%>">
			</td>
			<td  bgcolor="#FFFFFF"><input type="text" name="iD<%=intLoop%>" size="5" value="<%=DutyTime(intLoop)%>" readonly style="border:0;" <%IF iworktype(intLoop) =  "3" THEN%>disabled<%END IF%>></td>
			<td  bgcolor="#FFFFFF"><input type="text" name="iWHT<%=intLoop%>" size="5" value="<%IF iworktype(intLoop) =  "3"  THEN%><%=format00(2,Fix(holidaywdtime/60))&":"&format00(2,holidaywdtime mod 60)%><%END IF%>"  style="border:0;" ></td>
				<input type="hidden" name="intd<%=intLoop%>" size="5" value="<%=NightTime(intLoop)%>">
		</tr>
		<%
		Next %>
		<tr  align="center">
			<td colspan="6" bgcolor="<%= adminColor("tabletop") %>">�ְ� �� �ٹ��ð�</td>
			<td bgcolor="<%=adminColor("sky")%>"><span id="totDuty"><%=format00(2,Fix(totDutyTime/60))&":"&format00(2,totDutyTime mod 60)%></span></td>
			<td bgcolor="<%=adminColor("sky")%>"><span id="totWHT"><%=format00(2,Fix(holidaywdtime/60))&":"&format00(2,holidaywdtime mod 60)%></span></td>
		</tr>
 		</table>
 	</td>
 </tr>
 <tr>
 	<td>
 		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
 		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" colspan="4" align="center">�� �հ�</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�⺻��</td>
			<td bgcolor="#FFFFFF"><input type="text" name="idp"  size="10" style="text-align:right;" value="<%=defaultpay*ceilValue(totDutyTime/60*avgWeek)%>"> ��</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�⺻�ٹ��ð�</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totdt" value="<%=ceilValue(totDutyTime/60*avgWeek)%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">���޼���</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iwhdp"  size="10" style="text-align:right;" value="<%=defaultpay*ceilValue(holidaywdtime/60*avgWeek)%>"> ��</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">���޽ð�</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totwhdt" value="<%=ceilValue(holidaywdtime/60*avgWeek)%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ð��ܼ���</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iotp"  size="10" style="text-align:right;" value="<%=defaultpay*iOverTime*1.5%>"> ��</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ð��ܱٹ��ð�</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totot" value="<%=iOverTime%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�߰��ٹ�����</td>
			<td bgcolor="#FFFFFF"><input type="text" name="inp"  size="10" style="text-align:right;" value="<%=defaultpay*ceilValue(totNightTime/60*avgWeek)*0.5%>"> ��</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�߰��ٹ��ð�</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totnt" value="<%=ceilValue(totNightTime/60*avgWeek)%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">���ϱٹ�����</td>
			<td bgcolor="#FFFFFF"><input type="text" name="ihdp"  size="10" style="text-align:right;" value="0"> ��</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">���ϱٹ��ð�</td>
			<td bgcolor="#FFFFFF"><input type="text" name="tothdt" value="0" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">���޿��հ�</td>
			<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="itotp"  size="10" style="text-align:right;"value="<%=totPaySum%>"> ��</td>
		</tr>

		</table>
	</td>
</tr>
<tr>
	<td align="center"><%IF sMode="U" THEN%><input type="button" class="button" value="����" onClick="jsDel();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	<%END IF%>
	<input type="submit" class="button" value="���">
	<input type="button" class="button" value="���" onClick="history.back(-1);"></td>
</tr>
</form>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->