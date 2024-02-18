<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%

dim userid, empno, username
dim masteridx
dim part_sn, posit_sn

dim i

masteridx = Request("masteridx")



dim oVacation
Set oVacation = new CTenByTenVacation

oVacation.FRectMasterIdx = masteridx
oVacation.FRectpart_sn = part_sn

oVacation.GetMasterOne

userid = oVacation.FItemOne.Fuserid
empno = oVacation.FItemOne.Fempno
username = oVacation.FItemOne.Fusername
posit_sn = oVacation.FItemOne.Fposit_sn

%>
<html>
<head>
<title>연차(휴가) 신청</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script type="text/javascript">

function RequestVacation() {
	var frm = document.frm;

	if ((frm.startday.value == "") || (frm.endday.value == "")) {
		alert("기간을 입력해주십시요.");
		return false;
	}

	if (frm.totalday.value.length < 1) {
		alert("휴가일수를 입력해주십시요.\n\n휴가기간 입력 후 자동계산 버튼을 누르세요.");
		frm.totalday.focus();
		return false;
	}

	if (frm.totalday.value*0 != 0) {
		alert("휴가일수는 숫자만 입력가능합니다.");
		frm.totalday.focus();
		return false;
	}

	if (frm.totalday.value <= 0) {
		alert("휴가일수는 1 이상이어야 합니다.");
		frm.totalday.focus();
		return false;
	}

	if ((frm.ishalfvacation[0].checked == true) && (frm.totalday.value != 0.25)) {
		alert("사용일수가 0.25 일이어야 반반차등록이 가능합니다.");
		return false;
	}

	if ((frm.ishalfvacation[1].checked == true) && (frm.totalday.value != 0.5)) {
		alert("사용일수가 0.5 일이어야 반차등록이 가능합니다.");
		return false;
	}

	<% if (posit_sn = "13") then %>
		if (frm.totalhour.value.length < 1) {
			alert("휴가시간을 직접 입력해주십시요.");
			frm.totalhour.focus();
			return false;
		}
	
		if (frm.totalhour.value*0 != 0) {
			alert("휴가시간은 숫자만 입력가능합니다.");
			frm.totalhour.focus();
			return false;
		}
	
		if (frm.totalhour.value <= 0) {
			alert("휴가시간은 1시간 이상이어야 합니다.");
			frm.totalhour.focus();
			return false;
		}
	<% end if%>

	if (checkDate() == false) { return false; }

	if(confirm("등록 하시겠습니까?"))
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
		alert("종료일이 시작일보다 과거날짜입니다.");
		return false;
	}

	// 오프라인은 주말에도 근무한다.
	/*
	for (i = 0; i <= getDayInterval(startdate, enddate); i++) {
		tmp = addDate(startdate, i);
		tmp = getDayOfWeek(tmp);

		if ((tmp == "토") || (tmp == "일")) {
			alert("휴가기간에는 주말이 있어서는 안됩니다.");
			return false;
		}
	}
	*/

	// 주말,공휴일을 포함하여 사용하는 경우가 있음!
	/*
	if(frm.divcd.value=="5"){
		if(document.frm.totvd.value > totalday){
			if (confirm("장기휴가 신청시 주어진 휴가일수("+document.frm.totvd.value+"일) 만큼 사용일수를 지정해야합니다. 기간을 다시 입력하시겠습니까?")){
				return false;
			}
		}
	}
	*/

	var accTotDay = 0 ;
	<% if (posit_sn = "13") then %> 
		accTotDay =   document.frm.totalhour.value - document.frm.totvd.value ; 
		if (accTotDay >= 1 ) {
			alert("휴가 잔여시간보다 휴가신청 시간이 더 많습니다.");
			return false;
		}
	<%else%>
		accTotDay =  totalday - document.frm.totvd.value ;
		if (accTotDay >= 1 || (accTotDay==0.5  && frm.ishalfvacation[1].checked==true) || (accTotDay==0.25  && frm.ishalfvacation[0].checked==true)) {
			alert("휴가 잔여기간보다 휴가신청 일수가 더 많습니다.");
			return false;
		} 
	
		if (frm.ishalfvacation[2].checked) {
			// 반차제외
			if ((totalday*1 - 1) != getDayInterval(startdate, enddate)) {
				alert("휴가기간과 휴가 일수가 일치하지 않습니다.");
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
		alert("기간을 입력해주십시요.");
		return;
	}

	if (getDayInterval(startdate, enddate) < 0) {
		alert("잘못된 기간입니다.");
		return;
	}

	<% if (posit_sn = "13") then %>
		// 시급계약직
		var totday =  getDayInterval(startdate, enddate) + 1;
		 totalday.value = totday/0.125;
		frm.btday.value = totday;
		//document.ifrchk.location.href = "ifr_check_vacation.asp?mode=checkparthour&empno=<%= empno %>&startday=" + startday.value + "&endday=" + endday.value;
	<% else %>
		// 기타 계약직, 정규직
		totalday.value = getDayInterval(startdate, enddate) + 1;
	<% end if %>

	// 반차 여부 확인
	if(frm.ishalfvacation[0].checked||frm.ishalfvacation[1].checked) {
		frm.ishalfvacation[2].checked = true;
		halfgubun_tr();
	}
}

function jsReActFromIframe(totalDay) {
	var frm = document.frm;

	frm.totalday.value = totalDay;
	if (frm.totalhour) {
		// 하루는 8시간, 한시간은 0.125(= 1/8)
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
		// 연차
		frm.halfgubun.value = "no";
		document.getElementById("halfgubuntr").style.display = "none";

		frm.totalday.value = "";
		if (frm.totalhour) {
			frm.totalhour.value = "";
		}
		doInsertDayInterval();
	} else if(frm.ishalfvacation[0].checked == true) {
		// 반반차
		document.getElementById("halfgubuntr").style.display = "none";

		frm.halfgubun.value = "qt";

		frm.totalday.value = "0.25";
		if (frm.totalhour) {
			// 하루는 8시간, 한시간은 0.125(= 1/8)
			frm.totalhour.value = 0.25 / 0.125
		}
	} else {
		// 반차
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
			// 하루는 8시간, 한시간은 0.125(= 1/8)
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
<form name="frm" method="post" action="domodifyvacation.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="adddetail">
	<input type="hidden" name="masteridx" value="<%= masteridx %>">
	<input type="hidden" name="halfgubun" value="no">
<table width="470" border="0" cellpadding="2" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
	<tr height="25">
		<td valign="bottom" colspan=2  bgcolor="F4F4F4">
			<font color="red"><strong>연차(휴가) 신청</strong></font>
		</td>
	</tr>
	<tr align="left" height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">이름</td>
		<td bgcolor="#FFFFFF">
			<%= username %>
		</td>
	</tr>
	<tr align="left" height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">어드민 아이디</td>
		<td bgcolor="#FFFFFF">
			<%= userid %>
		</td>
	</tr>
	<tr align="left" height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">사번</td>
		<td bgcolor="#FFFFFF">
			<%= empno %>
		</td>
	</tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">구분</td>
		<td bgcolor="#FFFFFF">
			<input type="hidden" name="divcd" value="<%=oVacation.FItemOne.Fdivcd%>">
			<%= oVacation.FItemOne.GetDivCDStr %>
		</td>
	</tr>
	<tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">잔여일수</td>
    	<td bgcolor="#FFFFFF">	<input type="hidden" name="totvd" value="<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, (oVacation.FItemOne.GetRemainVacationDay)) %>">
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, (oVacation.FItemOne.GetRemainVacationDay)) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
    	</td>
    </tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">사용가능</td>
		<td bgcolor="#FFFFFF">
			<%= oVacation.FItemOne.IsAvailableVacation %>
		</td>
	</tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">사용가능기간</td>
		<td bgcolor="#FFFFFF">
			<%= Left(oVacation.FItemOne.Fstartday,10) %> - <%= Left(oVacation.FItemOne.Fendday,10) %>
		</td>
	</tr>
	<tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">기간</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="startday" class="text" size="11" maxlength="10" value="" onClick="jsPopCal('frm','startday');" style="cursor:hand;">
    		-
    		<input type="text" name="endday" class="text" size="11" maxlength="10" value="" onClick="jsPopCal('frm','endday');" style="cursor:hand;"> 
    		<input type="button" class="button" value="자동계산" onclick="doInsertDayInterval()"> 
    	</td>
    </tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">사용일수</td>
		<td bgcolor="#FFFFFF"> 
			<% if (posit_sn = "13") then %>
			<input type="hidden" name="totalday" class="text_ro" size="4" maxlength="6" value="" readonly>
			 시급계약직:   
			<input type="text" name="btday" class="text_ro" size="4" maxlength="6" value="" readonly>일 동안 총
			<input type="text" name="totalhour" class="text" size="4" maxlength="6" value="" onKeyUp="jsChkPartTime();"> 시간 
			<div style="padding:3px;font-size:11px;color:blue;"> 시간을 직접입력해주세요</div>
			<%else%>
			<input type="text" name="totalday" class="text_ro" size="4" maxlength="6" value="" readonly>
			 <% end if %>
		</td>
	</tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">반차여부</td>
		<td bgcolor="#FFFFFF">
			<label id="halfgubun0"><input type="radio"  name="ishalfvacation" value="Q" onClick="halfgubun_tr();">반반차(2시간)</label>&nbsp;
			<label id="halfgubun1"><input type="radio"  name="ishalfvacation" value="Y" onClick="halfgubun_tr();">반차(4시간)</label>&nbsp;
			<label id="halfgubun2"><input type="radio"  name="ishalfvacation" value="N" onClick="halfgubun_tr();" checked>아니오</label>&nbsp;
		</td>
	</tr>
	<tr align="left" height="25" id="halfgubuntr" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">반차구분</td>
		<td bgcolor="#FFFFFF">
			<label id='halfgubun5'><input type="radio"   name="halfgubun_tmp" value="am" onClick="halfgubunchk('am');" checked>오전반차</label>&nbsp;
			<label id='halfgubun6'><input type="radio"   name="halfgubun_tmp" value="pm" onClick="halfgubunchk('pm');">오후반차</label>
		</td>
	</tr>
	<tr align="center" height="25">
		<td colspan="2" bgcolor="#FFFFFF">
			<input type="button" class="button" value="등록" onclick="RequestVacation()">
			<input type="button" class="button" value="취소" onClick="self.close()">
		</td>
	</tr>
</table><br>
</form>

<iframe src="" width="0" height="0" name="ifrchk"></iframe>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
