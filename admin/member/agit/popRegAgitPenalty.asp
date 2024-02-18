<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ����Ʈ ������ �г�Ƽ ���
' History : 2018.05.08 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/member/tenAgitCalendarCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAgitCls.asp" -->
<%
	dim empno, idx
	idx		= requestCheckvar(request("idx"),10)

	'//���� ���� Ȯ��
	dim areaDiv, checkIn, checkOut, usePoint, useMoney, isUsing
 	if idx<>"" then
		dim oAgit
		Set oAgit = new CAgitCalendarDetail
		oAgit.read(idx)

		areaDiv			= oAgit.FAreaDiv
		empno			= oAgit.Fempno
		checkIn			= oAgit.FChkStart
		checkOut		= oAgit.FChkEnd
	    isUsing		= oAgit.FUsing
		checkIn = left(checkIn,10) & " " & Num2Str(Hour(checkIn),2,"0","R") & ":" & Num2Str(Minute(checkIn),2,"0","R")
		checkOut = left(checkOut,10) & " " & Num2Str(Hour(checkOut),2,"0","R") & ":" & Num2Str(Minute(checkOut),2,"0","R")

		Set oAgit = Nothing
	 
		if empno ="00000000000000" and userid ="admin" then
			Call Alert_Close("������ ��ϰ��Դϴ�.\n�г�Ƽ�� �ο��� �� �����ϴ�.")
			Response.End
		end if
	end if

	'// �������� Ȯ��
	dim userId, sUserName, sPartName, sDepartmentNameFull
	if empno <> "" then
		dim cMember
		Set cMember = new CTenByTenMember
		cMember.Fempno = empno
		cMember.fnGetMemberData

		userId			= cMember.Fuserid 
		sUserName      	= cMember.Fusername 
		sPartName     	= cMember.Fpart_name
		sDepartmentNameFull	= cMember.FdepartmentNameFull
		Set cMember = nothing
	end if

	dim penaltyStartDate, penaltyEndDate
	penaltyStartDate = Date
	penaltyEndDate = Left(DateAdd("d",-1,DateAdd("yyyy",1,penaltyStartDate)),10)
%>
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js?v=2"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
function fnPresetPenaltyTerm(tm) {
	var nowDt = new Date();

	switch(tm){
		case "1m":
			var toDt = toDateString(addDate(addMonth(nowDt,1),-1));
			break;
		case "3m":
			var toDt = toDateString(addDate(addMonth(nowDt,3),-1));
			break;
		case "6m":
			var toDt = toDateString(addDate(addMonth(nowDt,6),-1));
			break;
	}
	nowDt = toDateString(nowDt);
	
	document.getElementById("ChkStart").value=nowDt;
	document.getElementById("ChkEnd").value=toDt;
}

function fnChangeTakePoint(val) {
	if(val==1) {
		document.getElementById("lyrInputPoint").style.display = "inline";
	} else {
		document.getElementById("lyrInputPoint").style.display = "none";
	}
}

function fnSubmit() {
	var frm = document.frm;

	if(getDayInterval(toDate(frm.psdate.value), toDate(frm.pedate.value))<0) {
		alert("�Ⱓ�� �߸��Ǿ� �ֽ��ϴ�. ��¥�� Ȯ�����ּ���.");
		return;
	}

	if(frm.isTakePoint[1].checked&&frm.penaltyPoint.value==0) {
		alert("���� ����Ʈ�� �Է����ּ���.");
		return;
	}

	if(confirm("�Է��Ͻ� �������� �г�Ƽ�� ����Ͻðڽ��ϱ�?")) {
		frm.submit();
	}
}
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><b>�ٹ����� ����Ʈ ������ �г�Ƽ ���</b><br /><hr /></td>
</tr>
<tr>
	<td>
		<!-- ���� ���� -->
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%=adminColor("sky")%>" align="center"><b>���</b></td>
			<td colspan="3"><%=empno%></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%=adminColor("sky")%>" align="center"><b>����ID</b></td>
			<td><%=userId%></td>
			<td width="100" bgcolor="<%=adminColor("sky")%>" align="center"><b>�̸�</b></td>
			<td><%=sUserName%></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%=adminColor("sky")%>" align="center"><b>�ҼӺμ�</b></td>
			<td colspan="3"><%=sdepartmentNameFull %></td>
		</tr> 
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%=adminColor("sky")%>" align="center"><b>����Ʈ</b></td>
			<td colspan="3"><%=AgitName(areaDiv)%></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%=adminColor("sky")%>" align="center"><b>�̿�Ⱓ</b></td>
			<td colspan="3"><%=checkIn & " ~ " & checkOut%></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<!-- �г�Ƽ ��� -->
		<form name="frm" method="post" action="/admin/member/tenbyten/agit/tenbyten_agit_Process.asp" >
		<input type="hidden" name="mode" value="ptAdd">
		<input type="hidden" name="idx" value="<%=idx%>">
		<input type="hidden" name="sEn" value="<%=empno%>">
		<input type="hidden" name="userid" value="<%=userid%>">
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%=adminColor("sky")%>" align="center"><b>�г�Ƽ �Ⱓ</b></td>
			<td>
				<input id="ChkStart" name="psdate" value="<%=penaltyStartDate%>" class="text" size="10" maxlength="10" /> ~
				<input id="ChkEnd" name="pedate" value="<%=penaltyEndDate%>" class="text" size="10" maxlength="10" />
				<script type="text/javascript">
					var CAL_Start = new Calendar({
						inputField : "ChkStart", trigger    : "ChkStart",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "ChkEnd", trigger    : "ChkEnd",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
				<input type="button" class="button" value="1����" onclick="fnPresetPenaltyTerm('1m');" />
				<input type="button" class="button" value="3����" onclick="fnPresetPenaltyTerm('3m');" />
				<input type="button" class="button" value="6����" onclick="fnPresetPenaltyTerm('6m');" />
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%=adminColor("sky")%>" align="center"><b>����Ʈ ����</b></td>
			<td>
				<label><input type="radio" name="isTakePoint" value="0" checked onclick="fnChangeTakePoint(this.value);"> ����</label>
				<label><input type="radio" name="isTakePoint" value="1" onclick="fnChangeTakePoint(this.value);"> ����</label>
				<span id="lyrInputPoint" style="display:none;">
					/ ���� ����Ʈ: <input type="text" name="penaltyPoint" class="text" size="3" value="0" style="text-align:center;" />
				</span>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%=adminColor("sky")%>" align="center"><b>�г�Ƽ ����</b></td>
			<td><textarea name="penaltyCause" class="textarea" style="width:90%"></textarea></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td colspan="2" align="center">	 
				<input type="botton" value="���" class="button" style="width:60px;text-align:center;" onClick="fnSubmit();">
			</td>
		</tr>
		</table>
		</form>
	</td>
</tr>
</table>

<iframe id="ifmProc" src="" width="0" height="0" frameborder="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->