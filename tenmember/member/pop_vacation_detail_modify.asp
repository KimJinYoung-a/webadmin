<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%

dim empno
empno = session("ssBctSn")

dim login_level
login_level = session("ssAdminLsn")

dim masteridx
dim part_sn

dim i

masteridx = Request("masteridx")


'==============================================================================
dim oVacation
Set oVacation = new CTenByTenVacation

oVacation.FRectMasterIdx = masteridx
oVacation.FRectpart_sn = part_sn

oVacation.GetMasterOne

if (oVacation.FItemOne.Fempno <> empno) and (login_level >= "9") then
	'// �������� �α����� ��츸 �ٸ������ �ް���û ����(������)
	response.write "�߸��� �����Դϴ�."
	response.end
end if


'==============================================================================
dim cMember, posit_sn

Set cMember = new CTenByTenMember
	cMember.Fempno = empno
	cMember.fnGetMemberData

	posit_sn = cMember.Fposit_sn			'// 13 : �ñް����

Set cMember = nothing


'==============================================================================
dim yearvacation_startday, yearvacation_endday

yearvacation_startday = Cstr(Year(now())) & "-01-01"
yearvacation_endday = Cstr(Year(now()) + 1) & "-03-01"

%>
<html>
<head>
<title>����(�ް�) ��û</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script type="text/javascript">

function RequestVacation() {
	var frm = document.frm;

	if ((frm.startday.value == "") || (frm.endday.value == "")) {
		alert("�Ⱓ�� �Է����ֽʽÿ�.");
		return false;
	}

	if (frm.totalday.value.length < 1) {
		alert("�ް��ϼ��� �Է����ֽʽÿ�.\n\n�ް��Ⱓ �Է� �� �ڵ���� ��ư�� ��������.");
		frm.totalday.focus();
		return false;
	}

	if (frm.totalday.value*0 != 0) {
		alert("�ް��ϼ��� ���ڸ� �Է°����մϴ�.");
		frm.totalday.focus();
		return false;
	}

	if (frm.totalday.value <= 0) {
		alert("�ް��ϼ��� 1 �̻��̾�� �մϴ�.");
		frm.totalday.focus();
		return false;
	}

	if ((frm.ishalfvacation[0].checked == true) && (frm.totalday.value != 0.25)) {
		alert("����ϼ��� 0.25 ���̾�� �ݹ�������� �����մϴ�.");
		return false;
	}

	if ((frm.ishalfvacation[1].checked == true) && (frm.totalday.value != 0.5)) {
		alert("����ϼ��� 0.5 ���̾�� ��������� �����մϴ�.");
		return false;
	}

	<% if (posit_sn = "13") then %>
		if (frm.totalhour.value.length < 1) {
			alert("�ް��ð��� ���� �Է����ֽʽÿ�.");
			frm.totalhour.focus();
			return false;
		}
	
		if (frm.totalhour.value*0 != 0) {
			alert("�ް��ð��� ���ڸ� �Է°����մϴ�.");
			frm.totalhour.focus();
			return false;
		}
	
		if (frm.totalhour.value <= 0) {
			alert("�ް��ð��� 1�ð� �̻��̾�� �մϴ�.");
			frm.totalhour.focus();
			return false;
		}
	<% end if%>
	
	if (checkDate() == false) { return false; }

	if(confirm("��� �Ͻðڽ��ϱ�?"))
	{
		frm.submit();
	}
}

function checkDate() {
	var frm = document.frm;

	var startday = frm.startday.value;
	var endday = frm.endday.value;
	var totalday = frm.totalday.value;

	var startdate = toDate(startday);
	var enddate = toDate(endday);

	var tmp;
	var i;

	if (startdate > enddate) {
		alert("�������� �����Ϻ��� ���ų�¥�Դϴ�.");
		return false;
	}

	// ���������� �ָ����� �ٹ��Ѵ�.
	/*
	for (i = 0; i <= getDayInterval(startdate, enddate); i++) {
		tmp = addDate(startdate, i);
		tmp = getDayOfWeek(tmp);

		if ((tmp == "��") || (tmp == "��")) {
			alert("�ް��Ⱓ���� �ָ��� �־�� �ȵ˴ϴ�.");
			return false;
		}
	}
	*/

	// �ָ�,�������� �����Ͽ� ����ϴ� ��찡 ����!
	/*
	if(frm.divcd.value=="5"){
		if(document.frm.totvd.value > totalday){
			if (confirm("����ް� ��û�� �־��� �ް��ϼ�("+document.frm.totvd.value+"��) ��ŭ ����ϼ��� �����ؾ��մϴ�. �Ⱓ�� �ٽ� �Է��Ͻðڽ��ϱ�?")){
				return false;
			}
		}
	}
	*/
	var accTotDay = 0 ;
	<% if (posit_sn = "13") then %> 
	alert(document.frm.totalhour.value);
	alert(document.frm.totvd.value);
		accTotDay =  document.frm.totalhour.value - document.frm.totvd.value ; 
		if (accTotDay >= 1 ) {
			alert("�ް� �ܿ��ð����� �ް���û �ð��� �� �����ϴ�.");
			return false;
		}
	<% else %> 
		accTotDay =  totalday - document.frm.totvd.value ;
		if (accTotDay >= 1 || (accTotDay==0.5  && frm.ishalfvacation[1].checked==true) || (accTotDay==0.25  && frm.ishalfvacation[0].checked==true)) {
			alert("�ް� �ܿ��Ⱓ���� �ް���û �ϼ��� �� �����ϴ�.");
			return false;
		} 
 
		if (frm.ishalfvacation[2].checked) {
			// ��������
			if ((totalday*1 - 1) != getDayInterval(startdate, enddate)) {
				alert("�ް��Ⱓ�� �ް� �ϼ��� ��ġ���� �ʽ��ϴ�.");
				return false;
			}
		}
	<% end if %>

	return true;
}

function doInsertDayInterval() {
	var frm = document.frm;

	var startday = frm.startday;
	var endday = frm.endday;
	var totalday = frm.totalday;

	var startdate = toDate(startday.value);
	var enddate = toDate(endday.value);

	if ((startday.value == "") || (endday.value == "")) {
		//alert("�Ⱓ�� �Է����ֽʽÿ�.");
		return;
	}

	if (getDayInterval(startdate, enddate) < 0) {
		alert("�߸��� �Ⱓ�Դϴ�.");
		return;
	}
 	 
	<% if (posit_sn = "13") then %>
		// �ñް����
		var totday =  getDayInterval(startdate, enddate) + 1;
		 totalday.value = totday/0.125;
		frm.btday.value = totday;
		//document.ifrchk.location.href = "ifr_check_vacation.asp?mode=checkparthour&empno=<%= empno %>&startday=" + startday.value + "&endday=" + endday.value;
	<% else %>
		// ��Ÿ �����, ������
		totalday.value = getDayInterval(startdate, enddate) + 1;
	<% end if %>

	// ���� ���� Ȯ��
	if(frm.ishalfvacation[0].checked||frm.ishalfvacation[1].checked) {
		frm.ishalfvacation[2].checked = true;
		halfgubun_tr();
	}
}

function jsReActFromIframe(totalDay) {
	var frm = document.frm;

	frm.totalday.value = totalDay;
	if (frm.totalhour) {
		// �Ϸ�� 8�ð�, �ѽð��� 0.125(= 1/8)
		frm.totalhour.value = totalDay / 0.125
	}
}

function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function halfgubun_tr() {
	var frm = document.frm;

	if(frm.ishalfvacation[2].checked == true) {
		// ����
		frm.halfgubun.value = "no";
		document.getElementById("halfgubuntr").style.display = "none";

		frm.totalday.value = "";
		if (frm.totalhour) {
			frm.totalhour.value = "";
		}
		doInsertDayInterval();
	} else if(frm.ishalfvacation[0].checked == true) {
		// �ݹ���
		document.getElementById("halfgubuntr").style.display = "none";

		frm.halfgubun.value = "qt";

		frm.totalday.value = "0.25";
		if (frm.totalhour) {
			// �Ϸ�� 8�ð�, �ѽð��� 0.125(= 1/8)
			frm.totalhour.value = 0.25 / 0.125
		}
	} else {
		// ����
		document.getElementById("halfgubuntr").style.display = "";

		var ret;
		for (var i=0; i< frm.halfgubun_tmp.length; i++)
		{
			if (frm.halfgubun_tmp[i].checked == true)
			{
				ret = frm.halfgubun_tmp[i].value;
			}
		}
		halfgubunchk(ret)

		frm.totalday.value = "0.5";
		if (frm.totalhour) {
			// �Ϸ�� 8�ð�, �ѽð��� 0.125(= 1/8)
			frm.totalhour.value = 0.5 / 0.125
		}
	}
}

function halfgubunchk(v)
{
	if(v == "no")
	{
		document.frm.halfgubun.value = "no";
	}
	else
	{
		document.frm.halfgubun.value = v;
	}
}


function jsChkPartTime(){
	document.frm.totalday.value = (document.frm.totalhour.value)*0.125;
}
</script>
</head>
<body leftmargin="5" topmargin="5">
<form name="frm" method="post" action="modifyvacation_process.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="adddetail">
	<input type="hidden" name="masteridx" value="<%= masteridx %>">
	<input type="hidden" name="halfgubun" value="no">
<table width="470" border="0" cellpadding="2" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
	<tr height="25">
		<td valign="bottom" colspan=2  bgcolor="F4F4F4">
			<font color="red"><strong>����(�ް�) ��û</strong></font>
		</td>
	</tr>
	<tr align="left" height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">���</td>
		<td bgcolor="#FFFFFF">
			<%= empno %>
		</td>
	</tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF">
			<input type="hidden" name="divcd" value="<%=oVacation.FItemOne.Fdivcd%>">
			<%= oVacation.FItemOne.GetDivCDStr %>
		</td>
	</tr>
	<tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�ܿ��ϼ�</td>
    	<td bgcolor="#FFFFFF">	<input type="hidden" name="totvd" value="<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, (oVacation.FItemOne.Ftotalvacationday - (oVacation.FItemOne.Fusedvacationday + oVacation.FItemOne.Frequestedday))) %>">
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, (oVacation.FItemOne.Ftotalvacationday - (oVacation.FItemOne.Fusedvacationday + oVacation.FItemOne.Frequestedday))) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
    	</td>
    </tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��밡��</td>
		<td bgcolor="#FFFFFF">
			<%= oVacation.FItemOne.IsAvailableVacation %>
		</td>
	</tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��밡�ɱⰣ</td>
		<td bgcolor="#FFFFFF">
			<%= Left(oVacation.FItemOne.Fstartday,10) %> - <%= Left(oVacation.FItemOne.Fendday,10) %>
		</td>
	</tr>
	<tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�Ⱓ</td>
    	<td bgcolor="#FFFFFF">
   			<input type="text" id="termSdt" name="startday" readonly size="11" maxlength="10" value="" style="cursor:pointer; text-align:center;" /> ~
   			<input type="text" id="termEdt" name="endday" readonly size="11" maxlength="10" value="" style="cursor:pointer; text-align:center;" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "termSdt", trigger    : "termSdt",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();

						if(frm.endday.value==""||getDayInterval(frm.startday.value, frm.endday.value) < 0) frm.endday.value=frm.startday.value;
						doInsertDayInterval();	// ��¥ �ڵ����
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "termEdt", trigger    : "termEdt",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();

						if(frm.startday.value==""||getDayInterval(frm.startday.value, frm.endday.value) < 0) frm.startday.value=frm.endday.value;
						doInsertDayInterval();	// ��¥ �ڵ����
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
    	</td>
    </tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����ϼ�</td>
		<td bgcolor="#FFFFFF">
			
			<% if (posit_sn = "13") then %> 
			<input type="hidden" name="totalday" class="text_ro" size="4" maxlength="6" value="" readonly> ��
			  �ñް����: 
			<input type="text" name="btday" class="text_ro" size="4" maxlength="6" value="" readonly > �� ���� ��
			<input type="text" name="totalhour" class="text" size="4" maxlength="6" value="" onKeyUp="jsChkPartTime();"> �ð� 
			<div style="padding:3px;font-size:11px;color:blue;"> �ð��� �����Է����ּ���</div>
			<% else %>
			<input type="text" name="totalday" class="text_ro" size="4" maxlength="6" value="" readonly> ��
			<% end if %>
		</td>
	</tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td bgcolor="#FFFFFF">
			<label id="halfgubun0"><input type="radio"  name="ishalfvacation" value="Q" onClick="halfgubun_tr();">�ݹ���(2�ð�)</label>&nbsp;
			<label id="halfgubun1"><input type="radio"  name="ishalfvacation" value="Y" onClick="halfgubun_tr();">����(4�ð�)</label>&nbsp;
			<label id="halfgubun2"><input type="radio"  name="ishalfvacation" value="N" onClick="halfgubun_tr();" checked>�ƴϿ�</label>&nbsp;
		</td>
	</tr>
	<tr align="left" height="25" id="halfgubuntr" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td bgcolor="#FFFFFF">
			<label id="halfgubun5"><input type="radio"   name="halfgubun_tmp" value="am" onClick="halfgubunchk('am');" checked>��������</label>&nbsp;
			<label id="halfgubun6"><input type="radio"   name="halfgubun_tmp" value="pm" onClick="halfgubunchk('pm');">���Ĺ���</label>
		</td>
	</tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����������</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="workAgent" value="" size="20" maxlength="20" class="text" />
		</td>
	</tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�ް���/��󿬶�ó</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="callNum" value="" size="30" maxlength="30" class="text" />
		</td>
	</tr>
	<tr align="center" height="25">
		<td colspan="2" bgcolor="#FFFFFF">
			<input type="button" class="button" value="���" onclick="RequestVacation()">
			<input type="button" class="button" value="���" onClick="self.close()">
		</td>
	</tr>
</table><br>
</form>

<iframe src="" width="0" height="0" name="ifrchk"></iframe>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
