<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사원별 급여 기본정보 등록
' History : 2010.12.23 정윤정  생성
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
<title>계약패턴 등록</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript" src="/js/jsPayCal.js"></script>
<script language="javascript">
<!--
//폼 체크 후 submit 처리
	function jsChkform(frm){

		if(frm.part_sn.value ==""){
			 frm.part_sn.value = 1;
		}

		if(frm.sPN.value ==""){
			alert("패턴명을 입력해주세요");
			frm.sPN.focus();
			return false;
		}

		if(!IsDigit(frm.iHP.value)){
			alert("시급은 숫자만 입력가능합니다.");
			frm.iHP.focus();
			return false;
		}

		if(!IsDigit(frm.iEP.value)){
			alert("식대는 숫자만 입력가능합니다.");
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
		 alert("총근무 시간이 15시간이하일 경우 주휴일 설정은 불가능합니다.  ");
		 return false;
		 }

		 if(totDuty >= 900 && selWH == 0){
		 alert("주휴일을 설정해주세요");
		 return false;
		 }

		if( selWH > 1){
		alert("주휴일 설정은 하루만 가능합니다.");
		return false;
		}

		return true;

	}

	//삭제
	function jsDel(){
	 if(confirm("패턴을 삭제하시겠습니까?")){
	 document.frmDel.submit();
	 }
	}

	// 페이지 이동
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
	<td><strong>계약직사원 계약 패턴 등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">부서</td>
			<td bgcolor="#FFFFFF">
			<%=printPartOption("part_sn", part_sn)%>
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">패턴명</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sPN" value="<%=spatternname%>">
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">시급</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iHP" size="10" style="text-align:right;" value="<%=defaultpay%>" onKeyUp="jsSetMonthlypay();"> 원</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">식대</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iEP" size="10" style="text-align:right;" value="<%=foodpay%>"> 원</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">휴계시간</td>
			<td bgcolor="#FFFFFF"><input type="checkbox" name="blnBT" value="1" onClick="jsSetInBreakTime();" <%IF inBreakTime THEN%>checked<%END IF%>>근무시간 포함 </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">시간외 수당</td>
			<td bgcolor="#FFFFFF"><input type="checkbox" name="blnOT" value="1"  <%IF iOverTime > 0  THEN%>checked<%END IF%> onClick="jsSetOverTime();">지급
				<span id="spanOT" style="display:<%IF  iOverTime = 0  THEN%>none<%END IF%>;"><input type="text" size="5" maxlength="10" style="text-align:right;" name="iot" value="<%=iOverTime%>" onKeyUp="jsSetOverTimePay();"> 시간</span> </td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td><!-- 요일별 근무시간 지정 -->
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr align="center">
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">요일</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">구분</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" colspan="2">근무시간</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" colspan="2">휴계시간</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">총근무시간</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">주휴시간</td>
		</tr>
		<tr align="center">
			<td  bgcolor="<%= adminColor("tabletop") %>" >시작</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" >종료</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" >시작</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" >종료</td>
		</tr>
		<%
		For intLoop = 1 To 7%>
		<tr align="center">
			<td  bgcolor="<%= adminColor("tabletop") %>"><%=fnGetStringWD(intLoop)%></td>
			<td  bgcolor="#FFFFFF">
				<select name="selWH<%=intLoop%>"  onChange="jsSetWH(<%=intLoop%>);">
				<option value="1" <%IF iworktype(intLoop) ="1"  THEN%>selected<%END IF%>>근무일</option>
				<option value="2" <%IF iworktype(intLoop) ="2" THEN%>selected<%END IF%> style="color:blue">무급휴일</option>
				<option value="3" <%IF iworktype(intLoop) ="3" THEN%>selected<%END IF%> style="color:red">주휴일</option>
				<option value="4" <%IF iworktype(intLoop) ="4" THEN%>selected<%END IF%>>유급휴일</option>
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
			<td colspan="6" bgcolor="<%= adminColor("tabletop") %>">주간 총 근무시간</td>
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
			<td  bgcolor="<%= adminColor("tabletop") %>" colspan="4" align="center">월 합계</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">기본급</td>
			<td bgcolor="#FFFFFF"><input type="text" name="idp"  size="10" style="text-align:right;" value="<%=defaultpay*ceilValue(totDutyTime/60*avgWeek)%>"> 원</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">기본근무시간</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totdt" value="<%=ceilValue(totDutyTime/60*avgWeek)%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">주휴수당</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iwhdp"  size="10" style="text-align:right;" value="<%=defaultpay*ceilValue(holidaywdtime/60*avgWeek)%>"> 원</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">주휴시간</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totwhdt" value="<%=ceilValue(holidaywdtime/60*avgWeek)%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">시간외수당</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iotp"  size="10" style="text-align:right;" value="<%=defaultpay*iOverTime*1.5%>"> 원</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">시간외근무시간</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totot" value="<%=iOverTime%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">야간근무수당</td>
			<td bgcolor="#FFFFFF"><input type="text" name="inp"  size="10" style="text-align:right;" value="<%=defaultpay*ceilValue(totNightTime/60*avgWeek)*0.5%>"> 원</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">야간근무시간</td>
			<td bgcolor="#FFFFFF"><input type="text" name="totnt" value="<%=ceilValue(totNightTime/60*avgWeek)%>" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">휴일근무수당</td>
			<td bgcolor="#FFFFFF"><input type="text" name="ihdp"  size="10" style="text-align:right;" value="0"> 원</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">휴일근무시간</td>
			<td bgcolor="#FFFFFF"><input type="text" name="tothdt" value="0" size="5" style="text-align:right;border:0;" > </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">월급여합계</td>
			<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="itotp"  size="10" style="text-align:right;"value="<%=totPaySum%>"> 원</td>
		</tr>

		</table>
	</td>
</tr>
<tr>
	<td align="center"><%IF sMode="U" THEN%><input type="button" class="button" value="삭제" onClick="jsDel();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	<%END IF%>
	<input type="submit" class="button" value="등록">
	<input type="button" class="button" value="취소" onClick="history.back(-1);"></td>
</tr>
</form>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->