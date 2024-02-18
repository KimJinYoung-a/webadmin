<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사원별 급여 기본정보 등록
' History : 2010.12.23 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenPayCls.asp" -->
<!-- #include virtual="/lib/classes/member/fingerprints/fingerprints_cls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%

'// 현재 페이지에서는 시간과 관련된 내용만 계산한다.
'// 금액 관련된 내용은 처리하지 않는다.

'//2016.12 이후 급여계산 정책변경
'// - 1일~말일까지 급여계산 
'// - but! 인사팀 급여 확정일 26일 
'// - so! 26~말일까지 미리 계약서 내용으로 급여계산(실근무시간 아님)
'// - so! 다음달에 전달의 26~말일까지 실근무시간으로 급여재계산 처리
 

'변수선언
Dim sempno, ino
Dim djoinday, susername, iposit_sn,sposit_name,blnstatediv,dretireday
Dim startdate,enddate ,defaultpay,foodpay,jobpay,inbreaktime,holidaywdtime,predefaultpay
Dim intY, intM, intD, dYear, dMonth,dWeekday
Dim dEndDay,dNextDate
Dim clsPay, arrList
Dim dyyyymmdd, dstartHour, dstartMinute, dendHour, dendMinute, dbreakSHour,dbreakSMinute, dbreakEHour, dbreakEMinute,doutHour, dOutMinute,  iworktype, dstate, dStart, dEnd, dBreakS, dBreakE
Dim iWorkTime,iBreak, iextendWT ,inightWT,iholidayWT,iweekholidayWT,dNStart, dNEnd, dNBreakS, dNBreakE, iVacationTime
Dim totWorkTime, totextendWT ,totnightWT,totholidayWT,totweekholidayWT,totVacationTime
Dim  preStartDay, preEndDay,arrPre
Dim dSWD, dEWD,iWD, totWD,iWT,totWH, chkWHD
Dim dcStartHour(8),dcStartMinute(8),dcEndHour(8),dcEndMinute(8),dcBreakSHour(8),dcBreakSMinute(8),dcBreakEHour(8),dcBreakEMinute(8) ,defaulttime(8), dcWorkType(8), intLoop
Dim arrWorkTime(31),arrWorkType(31)
Dim sFingerYN, sVacationYN, ipart_sn
Dim ofingerprints,i,j,ircount,mstate
dim currDate
dim dDay, dSPayDay, dEPayDay, dFullPayDate, dSPayDate, dEPayDate
dim iLoopCnt
dim dPreYear, dPreMonth,dPreEPayDay
Dim iVer
Dim chkDate ,dSGetPayDay
dim arrRe
dim totWorkTime_pre,totextendWT_pre,totnightWT_pre,totholidayWT_pre,totweekholidayWT_pre,totVacationTime_pre
dim totWorkTime_re,totextendWT_re,totnightWT_re,totholidayWT_re,totweekholidayWT_re,totVacationTime_re 
dim totWorkTime_sum,totextendWT_sum,totnightWT_sum,totholidayWT_sum,totweekholidayWT_sum,totVacationTime_sum
dim stDCnt

'값 받아오기
sempno= requestCheckvar(Request("sEN"),14)
ino= requestCheckvar(Request("ino"),10)
dYear = requestCheckvar(Request("selY"),4)
dMonth = requestCheckvar(Request("selM"),2)
chkDate = dYear&"-"&format00(2,dMonth)
sFingerYN = requestCheckvar(Request("sFYN"),1) '지문인식근태관리 가져왔는지 여부
sVacationYN = requestCheckvar(Request("sVYN"),1)

if session("ssBctSn")<>sEmpno then 
	response.write "<script>self.close();</script>"
	response.end
end if

'기본값 설정 (현재 년월)
IF dYear = "" THEN dYear  = Year(Date())
IF dMonth = "" THEN dMonth  = Month(Date())

dNextDate = dateadd("m",1, dateserial(dYear,dMonth,1))	'검색다음달 1일
dEndDay = day(dateadd("d",-1,dNextDate))	'검색달 마지막 날짜 (다음달 1일 - 하루)  
preEndDay = dateadd("d", -1, dateserial(dYear,dMonth,1)) '이전달  마지막 일
'preStartDay = day(dateadd("d", -(weekday(preEndDay)-1),preEndDay)) 
dPreYear = year(preEndDay) '이전달 년
dPreMonth = month(preEndDay) '이전달 월
preEndDay  =day(preEndDay)  '이전달 마지막 날짜
 
''급여계산 기준일 변경 2014/01(1~31 => 26~25)
'------------------------------------------------------------------ 
IF  dYear&"-"&format00(2,dMonth)  = "2014-01" THEN '2014.01부터 급여종료일 25일로 변경됨
	dSPayDay = 1 '급여시작일
	dEPayDay	= 25 '급여종료일   
	dSPayDate = dateserial(dYear,dMonth,dSPayDay) '급여시작일: 해당월 1일부터
	dEPayDate = dateserial(dYear,dMonth,dEPayDay) '급여종료일: 해당월 25일까지 
	iLoopCnt = dEPayDay	- 1'총 급여일수 
ELSEIF dYear&"-"&format00(2,dMonth) > "2014-01"  and chkDate <"2016-12" then
	dSPayDay = 26 '급여시작일
	dEPayDay	= 25 '급여종료일   
	dSPayDate = dateserial(dPreYear,dPreMonth,dSPayDay) '급여시작일: 이전월 26일부터
	dEPayDate = dateserial(dYear,dMonth,dEPayDay) '급여종료일: 해당월 25일까지 
	iLoopCnt = (preEndDay-dSPayDay)+dEPayDay	'총 급여일수 
ELSEIF chkDate >= "2016-12"   then
	dSPayDay = 26 '급여시작일 
	dEPayDay	= dEndDay '급여종료일   
	dSPayDate = dateserial(dPreYear,dPreMonth,dSPayDay) '급여시작일: 이전월 26일부터	 
	dEPayDate = dateserial(dYear,dMonth,dEPayDay) '급여종료일: 해당월 25일까지 
	iLoopCnt = (preEndDay-dSPayDay)+dEPayDay	'총 급여일수 	
ELSE
	dSPayDay = 1
	dEPayDay	= dEndDay
	dSPayDate = dateserial(dYear,dMonth,dSPayDay) '급여시작일: 해당월 1일부터
	dEPayDate = dateserial(dYear,dMonth,dEPayDay)  '급여종료일: 해당월 말일까지
	iLoopCnt	= dEndDay - 1 '총 급여일수 
END IF  
	dPreEPayDay = day(dateadd("d", -1, dSPayDate))'이전달 급여 종료일
'------------------------------------------------------------------
 
'데이터 가져오기
set clsPay = new CPay
	'// ========================================================================
	'--사원 기본계약정보 가져오기
	'// ========================================================================
	clsPay.Fempno = sempno
	clsPay.Fyyyymm = dYear&"-"&format00(2,dMonth)
	clsPay.Fino	= ino
	clsPay.fnGetUserPayData

	sempno			= clsPay.Fempno
	susername		= clsPay.Fusername
	djoinday	  	= clsPay.Fjoinday
	blnstatediv 	= clsPay.Fstatediv
	iposit_sn		= clsPay.Fposit_sn
	sposit_name 	= clsPay.Fposit_name
	dretireday		= clsPay.Fretireday
	ipart_sn		= clsPay.Fpart_sn

	holidaywdtime 	= clsPay.Fholidaywdtime
	ino				= clsPay.Fino
	startdate		= clsPay.Fstartdate
	enddate			= clsPay.Fenddate
	defaultpay    	= clsPay.Fdefaultpay
	foodpay	    	= clsPay.Ffoodpay
	jobpay			= clsPay.Fjobpay
	inbreaktime		= clsPay.FinBreakTime
	predefaultpay    	= clsPay.FpreDefaultpay
	if IsNull(holidaywdtime) or (holidaywdtime = "") then
		holidaywdtime = 0
	end if

	For intLoop = 1 To 7
		dcStartHour(intLoop) 		= format00(2,Fix(clsPay.FStartTime(intLoop)/60))
		dcStartMinute(intLoop)  	= format00(2,clsPay.FStartTime(intLoop) mod 60)
		dcEndHour(intLoop)       	= format00(2,Fix(clsPay.FEndTime(intLoop)/60))
		dcEndMinute(intLoop)       	= format00(2,clsPay.FEndTime(intLoop)  mod 60)
		dcBreakSHour(intLoop)     	= format00(2,Fix(clsPay.FBreakSTime(intLoop)/60))
		dcBreakSMinute(intLoop)     = format00(2,clsPay.FBreakSTime(intLoop) mod 60)
		dcBreakEHour(intLoop)     	= format00(2,Fix(clsPay.FBreakETime(intLoop)/60))
		dcBreakEMinute(intLoop)     = format00(2,clsPay.FBreakETime(intLoop) mod 60)
		defaulttime(intLoop)		= clsPay.FdefaultTime(intLoop)
		dcWorkType(intLoop)			= clsPay.Fworktype(intLoop)
	Next

	'// ========================================================================
	'// 기작성된 월간 근무시간이 있는 경우 가져오기
	'// ========================================================================
	clsPay.fnGetmonthlypayData
	mstate = clsPay.Fstate

	'// ========================================================================
	'// 현재달의 첫 일요일에 해당하는 지난달 일수 + 그 이전 일주일 목록가져오기
	'// 2013-02-01 금인경우 2013-01-27 까지의 목록 + 그 이전 1주일목록
	'// 주휴일 산정에 사용한다.
	'// ========================================================================
	clsPay.FPreyyyymmdd = dSPayDate
	arrPre =clsPay.fnGetPreDailypayData
  arrRe  =clsPay.fnGetPreReDailypayData
  
	'--검색달 근무시간 내역
	IF sFingerYN = "Y" THEN    '지문인식내역 가져올 경우
		set clsPay = nothing
		set ofingerprints = new cfingerprints_list
		ofingerprints.frectpart_sn = ipart_sn
		ofingerprints.frectempno = sempno
		ofingerprints.FrectSDate = dSPayDate
		ofingerprints.FrectEDate = dateadd("d",1,dEPayDate)
		ofingerprints.ffingerprints_sum()
		 ircount = ofingerprints.FresultCount
		if ircount<=0 then
			set ofingerprints =nothing
			 Alert_return("지문인식근태내역이 존재하지 않습니다. 확인 후 다시 시도해주세요")
			response.end
		END IF
	ELSE 
		clsPay.FSyyyymm = dSPayDate
		clsPay.FEyyyymm = dEPayDate
		arrList = clsPay.fnGetDailypayData
		set clsPay = nothing
	END IF

	'// ========================================================================
	'// 휴가목록 가져오기
	dim vacationRequestCount
	vacationRequestCount = 0

	dim oVacation
	Set oVacation = new CTenByTenVacation

	if (sVacationYN = "Y") then
		oVacation.FRectEmpNO = sempno
		oVacation.FRectIsDelete = "N"
		oVacation.FRectStartDate =  dSPayDate 
		oVacation.FRectEndDate = dEPayDate
		oVacation.FPageSize = 50
		oVacation.FCurrPage = 1

		oVacation.GetDetailList

		for i = 0 to oVacation.FResultCount - 1
			if (oVacation.FItemList(i).Fstatedivcd = "R") then
				'// 승인대기
				vacationRequestCount = vacationRequestCount + 1
			end if
		next

		if (vacationRequestCount > 0) then
			response.write "<script>alert('승인대기 상태의 휴가가 있습니다.\n\n먼저 승인해야 휴가내역을 가져올 수 있습니다.');</script>"
		end if
	end if

IF dYear >= 2011 and susername ="" or isnull(susername) THEN
	IF Request("selY") = "" THEN
%>
	<script language="javascript">
	alert("계약정보가 존재하지 않거나 해당 회차에 해당하는 날짜가 안됐습니다.  확인후 다시 시도해주세요");
	self.close();
	</script>
<%
	ELSE
	Alert_return("계약정보가 존재하지 않거나 해당 회차에 해당하는 날짜가 안됐습니다.  확인후 다시 시도해주세요")
	END IF

END IF 

IF datediff("m",startdate,dEPayDate)  < 0 or datediff("m",enddate,dateadd("m",-1,dEPayDate)) > 0 THEN
	dstate = 9
END IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<html>
<head>
<title>근무시간 등록</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jsPayCal.js"></script> 
<script type="text/javascript">
<!--
	function jsSearch(){
		var dNowYear, dNowMonth;
		var date = new Date();
		dNowYear = date.getFullYear();
		dNowMonth = date.getMonth() + 1;

        document.frmSearch.sFYN.value  ="";
	    document.frmSearch.sVYN.value  = "";
	 	document.frmSearch.submit();
	}

	function jsComplete(){
		if(confirm("근무시간등록을 작성완료하시겠습니까? 작성완료시 월급여가 생성됩니다.")){
			document.frmPay.hidS.value="1";
			document.frmPay.submit();
		}
		return;
	}

//-->
</script>
</head>
<body leftmargin="10" topmargin="10">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>계약직사원 근무시간 등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">사번</td>
			<td bgcolor="#FFFFFF" width="180"><%=sempno%> <%IF blnstatediv ="N" THEN%><font color="red">[퇴사]</font><%END IF%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">입사일</td>
			<td bgcolor="#FFFFFF"><%IF djoinday <> "" THEN%><%=formatdate(djoinday,"0000-00-00")%><%END IF%></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">이름</td>
			<td bgcolor="#FFFFFF"><%=susername%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">퇴사일</td>
			<td bgcolor="#FFFFFF"><%IF blnstatediv = "N" THEN%><%=formatdate(dretireday,"0000-00-00")%><%END IF%></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">계약구분</td>
			<td bgcolor="#FFFFFF"><%=sposit_name%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">시간급</td>
			<td bgcolor="#FFFFFF"><%if predefaultpay>0 then%>(전월: <%=formatnumber(predefaultpay,0)%> 원) <%end if%><%=formatnumber(defaultpay,0)%> 원</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">계약일</td>
			<td bgcolor="#FFFFFF"><%IF startdate <> "" THEN%><%=ino%>. <%=formatdate(startdate,"0000-00-00")%> ~ <%=formatdate(enddate,"0000-00-00")%><%END IF%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">휴계시간</td>
			<td bgcolor="#FFFFFF"><%IF inbreaktime THEN%>근무시간 포함<%ELSE%>근무시간 포함안함<%END IF%></td>
		</tr>
		</table>
	</td>
</tr>
<form name="frmSearch" method="get" action="">
<input type="hidden" name="sEN" value="<%=sEmpno%>">
<input type="hidden" name="ino" value="<%=ino%>">
<input type="hidden" name="sFYN" value="<%= sFingerYN %>">
<input type="hidden" name="sVYN" value="<%= sVacationYN %>">

<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a">
		<tr>
			<td>
				근무날짜:
				<select name="selY">
				<%For intY = Year(date()) To 2010 Step -1%>
				<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dYear) THEN%>selected<%END IF%>><%=intY%></option>
				<%Next%>
				</select>
				년
				<select name="selM">
				<%For intM = 1 To 12%>
				<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dMonth) THEN%>selected<%END IF%>><%=intM%></option>
				<%Next%>
				</select>
				월
				<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr bgcolor="<%= adminColor("gray") %>"  align="center">
			<td rowspan="2">일</td>
			<td rowspan="2">요일</td>
			<td rowspan="2">구분</td>
			<td colspan="2">근무시간</td>
			<td colspan="2">휴계시간</td>
			<td rowspan="2">외출<br>시간</td>
			<td rowspan="2">기준<br>시간</td>
			<td rowspan="2">연차<br>시간</td> 
			<td rowspan="2">기본근무<br>시간</td>
			<td rowspan="2">연장근무<br>시간</td>
			<td rowspan="2">야간근무<br>시간</td>
			<td rowspan="2">휴일근무<br>시간</td>
			<td rowspan="2">주휴일<br>시간</td>
		</tr>
		<tr  bgcolor="<%= adminColor("gray") %>"  align="center">
			<td>시작</td>
			<td>종료</td>
			<td>시작</td>
			<td>종료</td>
		</tr>
		<%
		'// ========================================================================
		'// 현재달의 첫 일요일에 해당하는 지난달 일수 + 그 이전 일주일 목록가져오기
		'// 2013-02 인경우 2013-01-27 까지의 목록 + 그 이전 1주일목록
		'// 주휴일 산정에 사용한다.
		'// ========================================================================
		dim totPWD, chkPWHD, blnPWH, imaxday, iminday, imaxD
		dim totCWD, chkCWHD, blnCWH
		dim sundayCnt
		sundayCnt = 0
		chkPWHD = 0
		blnPWH = 0
		totPWD = 0
		totCWD = 0
		chkCWHD = 0
		blnCWH = 0
		imaxday = 0
		iminday = 0
		imaxD = 0
		IF isArray(arrPre) THEN
			'// ================================================================
			'// 전월 데이타가 있는 경우 표시
			'// ================================================================
			imaxD = UBound(arrPre,2)
			iminday = day(arrPre(0,0))
			imaxday = right(arrPre(0,UBound(arrPre,2)),2)
			if imaxday = 32 then imaxD = UBound(arrPre,2)-1

			For intD = 0 To imaxD
				iWorkTime 		= 0
				iextendWT  		= 0
				inightWT		= 0
				iholidayWT		= 0
				iweekholidayWT	= 0
				iVacationTime	= 0
				'// vbSunday = 1
				if weekday(arrPre(0,intD)) = 1 then
					sundayCnt = sundayCnt + 1
				end if

				iWorkTime 		= arrPre(7,intD)
				iextendWT 		= arrPre(8,intD)
				inightWT		= arrPre(9,intD)
				iholidayWT		= arrPre(10,intD)
				iweekholidayWT	= arrPre(11,intD)
				iVacationTime	= arrPre(13,intD)
				
				if (sundayCnt = 1) then
					'전주
					totPWD  		= totPWD + iWorkTime	'전체 근무시간

					if (arrPre(5,intD) = "3") then
						blnPWH = blnPWH + 1
					end if

					IF arrPre(5,intD) = "1"  and iWorkTime = 0 THEN
						'근무일에 근무를 안했을경우 수당 지급안됨
						chkPWHD  =  chkPWHD  + 1
					END IF
				else
					'이번주 전달부분
					totCWD  		= totCWD + iWorkTime	'전체 근무시간

					if (arrPre(5,intD) = "3") then
						blnCWH = blnCWH + 1
					end if

					IF arrPre(5,intD) = "1"  and iWorkTime = 0 THEN
						'근무일에 근무를 안했을경우 수당 지급안됨
						chkCWHD  =  chkCWHD  + 1
					END IF
				end if
			%>
			<% if (weekday(arrPre(0,intD)) = 1) then %>
			<tr   bgcolor="#CCCCCC" align="center"><td colspan="14" height="2"></td></tr>
			<% end if %>
			<tr   bgcolor="#DFDFDF" align="center">
				<td><div  id="<%=day(arrPre(0,intD))%>"><%=day(arrPre(0,intD))%></div></td>
				<td><%=fnGetStringWD(weekday(arrPre(0,intD)))%><input type="hidden" name="hidPWeD<%=day(arrPre(0,intD))%>" value="<%=weekday(arrPre(0,intD))%>"></td>
				<td>
					<%IF arrPre(5,intD)  = "1" THEN%>
						근무일
					<%ELSEIF arrPre(5,intD)  = "2" THEN%>
						<font color="blue">무급휴일<font>
					<%ELSEIF arrPre(5,intD)  = "3" THEN%>
						<font color="red">주휴일</font>
					<%ELSEIF arrPre(5,intD)  = "6" THEN%>
						<font color="red">주휴일(무)<font>
					<%ELSEIF arrPre(5,intD)  = "7" THEN%>
						<font color="red">주휴일(유)<font>
					<%ELSEIF arrPre(5,intD)  = "4" THEN%>
						<font color="red">유급휴일<font>
					<%ELSEIF arrPre(5,intD)  = "5" THEN%>
						<font color="red">공휴일<font>
					<%END IF%>
					<input type="hidden" name="iPWH<%=day(arrPre(0,intD))%>" value="<%=arrPre(5,intD)%>">
				</td>
				<td><input type="text" class="text" name="iPWS<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(1,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF " readonly></td>
				<td><input type="text" class="text"  name="iPWE<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(2,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
				<td><input type="text" class="text"  name="iPBS<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(3,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
				<td><input type="text" class="text"  name="iPBE<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(4,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
				<td><input type="text" class="text"  name="iPO<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(12,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
				<td></td>
				<td><input type="text" class="text"  name="iPVT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iVacationTime)%>"></td>
				<td><input type="text" class="text"  name="iPWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iWorkTime)%>"></td>
				<td><input type="text" class="text"  name="iPeWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iextendWT)%>"></td>
				<td><input type="text" class="text"  name="iPnWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(inightWT)%>"></td>
				<td><input type="text" class="text"  name="iPhWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iholidayWT)%>"></td>
				<td><input type="text" class="text"  name="iPwhWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iweekholidayWT)%>"></td>
			</tr>
			<%	Next
		END IF
		%>
		<input type="hidden" name="hidPWD" value="<%=totPWD%>"><!--  지난 달 마지막 주 근무시간(2013-02 의 경우 2013-01-27 이후) -->
		<input type="hidden" name="hidPWHD" value="<%=chkPWHD%>"><!--  - 결근횟수(2013-02 의 경우 2013-01-27 이후) -->
		<input type="hidden" name="blnPWH" value="<%=blnPWH%>"><!--    - 주휴일수(2013-02 의 경우 2013-01-27 이후) -->
		<input type="hidden" name="hidCWD" value="<%=totCWD%>"><!--  이번주 전달부분 근무시간(2013-02 의 경우 2013-01-27 이후) -->
		<input type="hidden" name="hidCWHD" value="<%=chkCWHD%>"><!--  - 결근횟수(2013-02 의 경우 2013-01-27 이후) -->
		<input type="hidden" name="blnCWH" value="<%=blnCWH%>"><!--    - 주휴일수(2013-02 의 경우 2013-01-27 이후) -->
		<input type="hidden" name="hidPSD" value="<%=iminday%>"> 
		<input type="hidden" name="hidPED" value="<%=imaxday%>">
		<tr>
			<td colspan="15" bgcolor="#FFFFFF"></td>
		</tr>
		<%
 
			if isArray(arrRe) THEN
				For intD = 0 To UBound(arrRe,2)  
				iWorkTime 		= arrRe(7,intD)
				iextendWT 		= arrRe(8,intD)
				inightWT		= arrRe(9,intD)
				iholidayWT		= arrRe(10,intD)
				iweekholidayWT	= arrRe(11,intD)
				iVacationTime	= arrRe(13,intD) 
					'전주
				totWorkTime_pre 	=    totWorkTime_pre +  iWorkTime	  
				totextendWT_pre  	=    totextendWT_pre  + 	iextendWT
				totnightWT_pre		=    totnightWT_pre		 + inightWT
				totholidayWT_pre	=    totholidayWT_pre	+  iholidayWT
				totweekholidayWT_pre = totweekholidayWT_pre + iweekholidayWT
				totVacationTime_pre = totVacationTime_pre + iVacationTime
			%>
			<% if (weekday(arrRe(0,intD)) = 1) then %>
			<tr   bgcolor="#CCCCCC" align="center"><td colspan="14" height="2"></td></tr>
			<% end if %>
			<tr   bgcolor="#e3f1fb" align="center">
				<td><div  id="<%=day(arrRe(0,intD))%>"><%=day(arrRe(0,intD))%></div></td>
				<td><%=fnGetStringWD(weekday(arrRe(0,intD)))%><input type="hidden" name="hidPWeD<%=day(arrRe(0,intD))%>" value="<%=weekday(arrPre(0,intD))%>"></td>
				<td>
					<%IF arrRe(5,intD)  = "1" THEN%>
						근무일
					<%ELSEIF arrRe(5,intD)  = "2" THEN%>
						<font color="blue">무급휴일<font>
					<%ELSEIF arrRe(5,intD)  = "3" THEN%>
						<font color="red">주휴일</font>
					<%ELSEIF arrRe(5,intD)  = "6" THEN%>
						<font color="red">주휴일(무)</font>
					<%ELSEIF arrRe(5,intD)  = "7" THEN%>
						<font color="red">주휴일(유)</font>
					<%ELSEIF arrRe(5,intD)  = "4" THEN%>
						<font color="red">유급휴일<font>
					<%ELSEIF arrRe(5,intD)  = "5" THEN%>
						<font color="red">공휴일<font>
					<%END IF%>
					<input type="hidden" name="iPWH<%=day(arrRe(0,intD))%>" value="<%=arrRe(5,intD)%>">
				</td>
				<td><input type="text" class="text"  name="iPWS<%=day(arrRe(0,intD))%>" value="<%=fnSetTimeFormat(arrRe(1,intD))%>" size="5" maxlength="5" style="border:0;background:#e3f1fb " readonly></td>
				<td><input type="text" class="text"  name="iPWE<%=day(arrRe(0,intD))%>" value="<%=fnSetTimeFormat(arrRe(2,intD))%>" size="5" maxlength="5" style="border:0;background:#e3f1fb" readonly></td>
				<td><input type="text" class="text"  name="iPBS<%=day(arrRe(0,intD))%>" value="<%=fnSetTimeFormat(arrRe(3,intD))%>" size="5" maxlength="5" style="border:0;background:#e3f1fb" readonly></td>
				<td><input type="text" class="text"  name="iPBE<%=day(arrRe(0,intD))%>" value="<%=fnSetTimeFormat(arrRe(4,intD))%>" size="5" maxlength="5" style="border:0;background:#e3f1fb" readonly></td>
				<td><input type="text" class="text"  name="iPO<%=day(arrRe(0,intD))%>" value="<%=fnSetTimeFormat(arrRe(12,intD))%>" size="5" maxlength="5" style="border:0;background:#e3f1fb" readonly></td>
				<td></td>
				<td><input type="text" class="text"  name="iPVT<%=day(arrRe(0,intD))%>" style="border:0;background:#e3f1fb" readonly size="5" value="<%=fnSetTimeFormat(iVacationTime)%>"></td>
				<td><input type="text" class="text"  name="iPWT<%=day(arrRe(0,intD))%>" style="border:0;background:#e3f1fb" readonly size="5" value="<%=fnSetTimeFormat(iWorkTime)%>"></td>
				<td><input type="text" class="text"  name="iPeWT<%=day(arrRe(0,intD))%>" style="border:0;background:#e3f1fb" readonly size="5" value="<%=fnSetTimeFormat(iextendWT)%>"></td>
				<td><input type="text" class="text"  name="iPnWT<%=day(arrRe(0,intD))%>" style="border:0;background:#e3f1fb" readonly size="5" value="<%=fnSetTimeFormat(inightWT)%>"></td>
				<td><input type="text" class="text"  name="iPhWT<%=day(arrRe(0,intD))%>" style="border:0;background:#e3f1fb" readonly size="5" value="<%=fnSetTimeFormat(iholidayWT)%>"></td>
				<td><input type="text" class="text"  name="iPwhWT<%=day(arrRe(0,intD))%>" style="border:0;background:#e3f1fb" readonly size="5" value="<%=fnSetTimeFormat(iweekholidayWT)%>"></td>
			</tr>
			<%	Next
	
		%>
	 
 <tr   bgcolor="<%=adminColor("sky")%>" align="center"> 
				<td colspan="9"><b>A.</b> [<%=dPreMonth%>/26 ~ <%=dPreMonth%>/<%=preEndDay%>] <b>합계</b></td> 
				<td><input type="text" class="text" name="totPVT" style="border:0;background:#DDDDFF;color:<%IF totVacationTime_pre  = 0  THEN %>gray<%ELSEIF  totVacationTime_pre < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totVacationTime_pre)%>"></td>
				<td><input type="text" class="text"  name="totPWT" style="border:0;background:#DDDDFF;color:<%IF totWorkTime_pre  = 0  THEN %>gray<%ELSEIF  totWorkTime_pre < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totWorkTime_pre)%>"></td> 
				<td><input type="text" class="text"  name="totPeWT" style="border:0;background:#DDDDFF;color:<%IF totextendWT_pre  = 0  THEN %>gray<%ELSEIF  totextendWT_pre < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totextendWT_pre)%>"></td>
				<td><input type="text" class="text"  name="totPnWT" style="border:0;background:#DDDDFF;color:<%IF totnightWT_pre  = 0  THEN %>gray<%ELSEIF  totnightWT_pre < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totnightWT_pre)%>"></td>
				<td><input type="text" class="text"  name="totPhWT" style="border:0;background:#DDDDFF;color:<%IF totholidayWT_pre  = 0  THEN %>gray<%ELSEIF  totholidayWT_pre < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totholidayWT_pre)%>"></td>
				<td><input type="text" class="text"  name="totPwhWT" style="border:0;background:#DDDDFF;color:<%IF totweekholidayWT_pre  = 0  THEN %>gray<%ELSEIF  totweekholidayWT_pre < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totweekholidayWT_pre)%>"></td>
		</tr> 
		<tr>
			<td colspan="15" bgcolor="#ffffff"></td>
		</tr>
		<%	END IF
		 %>
	<%'--------------- 현재 달 급여 설정 start ----------------
			totWorkTime = 0
			totextendWT  = 0
			totnightWT	=0
			totholidayWT=0
			totweekholidayWT=0
			totVacationTime = 0
		 i = 0  
	 
		 dFullPayDate = dSPaydate'급여시작일
		For intD = 0 To iLoopCnt
			iworktype = ""
		 
			dDay = day(dFullPayDate)
			dWeekday = weekday(dFullPayDate) 
		   
			dStart		= 0
			dEnd		= 0
			dBreakS		= 0
			dBreakE		= 0
			dstartHour	= 0
			dstartMinute= 0
			dendHour	= 0
			dendMinute	= 0 
			doutHour	= 0
			doutMinute	= 0 
			dstate		= 0 
			iWorkTime	= 0
			iextendWT	= 0
			inightWT	= 0
			iholidayWT	= 0
			iweekholidayWT= 0
			ivacationTime = 0
				
 			'계약정보 가져오기-------------------- 
			IF djoinday  > dFullPayDate  or enddate < dFullPayDate or dretireday < dFullPayDate THEN 
				dbreakSHour = "00"
				dbreakSMinute ="00"
				dbreakEHour = "00"
				dbreakEMinute ="00"
	 			iworktype = 0 
	 			dstartHour = "00"
	 			dstartMinute= "00"
	 			dendHour= "00"
	 			dendMinute= "00"
	 		ELSEIF 	dFullPayDate <  dateserial(dYear,dMonth,"1") and dFullPayDate>"2016-12-25" THEN
	 			IF isArray(arrRe) THEN 
				 	IF intD <= UBound(arrRe,2) THEN 
						dyyyymmdd	= arrRe(0,intD) 
						dStart		= arrRe(1,intD)
						dEnd		= arrRe(2,intD)
						dBreakS		= arrRe(3,intD)
						dBreakE		= arrRe(4,intD)
						dstartHour	= format00(2,Fix(dStart/60))
						dstartMinute= format00(2,dStart mod 60)
						dendHour	= format00(2,Fix(dEnd/60))
						dendMinute	= format00(2,dEnd mod 60)
						dbreakSHour	= format00(2,Fix(dBreakS/60))
						dbreakSMinute= format00(2,dBreakS mod 60)
						dbreakEHour	= format00(2,Fix(dBreakE/60))
						dbreakEMinute= format00(2,dBreakE mod 60)
						doutHour	= format00(2,Fix(arrRe(12,intD)/60))
						doutMinute	= format00(2, arrRe(12,intD) mod 60)
						iworktype	= arrRe(5,intD) 

						iWorkTime	= arrRe(7,intD)
						iextendWT	= arrRe(8,intD)
						inightWT	= arrRe(9,intD)
						iholidayWT	= arrRe(10,intD)
						iweekholidayWT= arrRe(11,intD) 
						ivacationTime = arrRe(13,intD) 
			 		END IF
				END IF
			ELSE 
				dbreakSHour 	= dcbreakSHour(dWeekday)
				dbreakSMinute = dcbreakSMinute(dWeekday)
				dbreakEHour 	= dcbreakEHour(dWeekday)
				dbreakEMinute = dcbreakEMinute(dWeekday)
	 			iworktype 		= dcWorkType(dWeekday)
	 			dstartHour 		= dcStartHour(dWeekday)
	 			dstartMinute 	= dcStartMinute(dWeekday)
	 			dendHour			= dcEndHour(dWeekday)
	 			dendMinute		= dcEndMinute(dWeekday)  
	 			
	 			iextendWT	= 0
	 			iWorkTime	=  (dendHour*60+dendMinute)-(dstartHour*60+dstartMinute)
	 			if iWorkTime > 480 THEN
	 				iextendWT = iWorkTime -480
	 				iWorkTime = 480
	 			end if	
				
				inightWT	= 0
				
				
				iholidayWT	= 0
				iweekholidayWT= 0
				ivacationTime = 0
		 
	 	  	
 			END IF  
 			'--------------------------------------- 
  
			IF sFingerYN = "Y" THEN     '지문인식근태내역 가져오기
				if dFullPayDate < dateserial(dYear,dMonth,"26") then
			 	dstartHour 	= "00"
				dstartMinute= "00"
				dendHour 	= "00"
				dendMinute 	= "00"
				doutHour 	= "00"
				doutMinute 	= "00"
				arrWorkTime(dDay) = 0
				arrWorkType(dDay) = iworktype
				end if
				 if i < ircount then
					dyyyymmdd	= ofingerprints.FItemList(i).fyyyymmdd  
					if   dyyyymmdd  =  Cstr(dFullPayDate) then 
						dstate = 0
						dstartHour	= format00(2,hour(ofingerprints.FItemList(i).fInTime))
						dstartMinute= format00(2,minute(ofingerprints.FItemList(i).fInTime))
						if ofingerprints.FItemList(i).fOutTime <> "1900-01-01" then
						dendHour	= format00(2,hour(ofingerprints.FItemList(i).fOutTime))
						dendMinute	= format00(2,minute(ofingerprints.FItemList(i).fOutTime))
						end if

						if (dstartHour*1 > dendHour*1) then
							'// 야간근무
							dendHour = dendHour*1 + 24
						end if

						doutHour	= format00(2,Fix(ofingerprints.FItemList(i).fexmin/60))
						doutMinute	= format00(2, ofingerprints.FItemList(i).fexmin mod 60)

						iWorkTime	= ofingerprints.FItemList(i).fworkmin
						ibreak = (dbreakEHour*60+dbreakEMinute)-(dbreakSHour*60+ dbreakSMinute)

					 i = i + 1
					end if

 				end if
			ELSE
				
				IF isArray(arrList) THEN 
				 	IF intD <= UBound(arrList,2) THEN 
						dyyyymmdd	= arrList(0,intD) 
						dStart		= arrList(1,intD)
						dEnd		= arrList(2,intD)
						dBreakS		= arrList(3,intD)
						dBreakE		= arrList(4,intD)
						dstartHour	= format00(2,Fix(dStart/60))
						dstartMinute= format00(2,dStart mod 60)
						dendHour	= format00(2,Fix(dEnd/60))
						dendMinute	= format00(2,dEnd mod 60)
						dbreakSHour	= format00(2,Fix(dBreakS/60))
						dbreakSMinute= format00(2,dBreakS mod 60)
						dbreakEHour	= format00(2,Fix(dBreakE/60))
						dbreakEMinute= format00(2,dBreakE mod 60)
						doutHour	= format00(2,Fix(arrList(12,intD)/60))
						doutMinute	= format00(2, arrList(12,intD) mod 60)
						iworktype	= arrList(5,intD)
						dstate		= arrList(6,intD)

						iWorkTime	= arrList(7,intD)
						iextendWT	= arrList(8,intD)
						inightWT	= arrList(9,intD)
						iholidayWT	= arrList(10,intD)
						iweekholidayWT= arrList(11,intD) 
						ivacationTime = arrList(13,intD) 
			 		END IF
				END IF
			END IF 
		 
		  if   chkDate >= "2017-01" then
		 if   Cstr(dFullPayDate) =  Cstr(dateserial(dYear,dMonth,1)) then 
							
				totWorkTime_re 	=    totWorkTime 	  
				totextendWT_re  	=    totextendWT  	
				totnightWT_re		=    totnightWT		  
				totholidayWT_re	=    totholidayWT	  
				totweekholidayWT_re = totweekholidayWT
				totVacationTime_re = totVacationTime
				
				totWorkTime_sum 	=    totWorkTime_re 	       - totWorkTime_pre 	    
				totextendWT_sum  	=    totextendWT_re  	     - totextendWT_pre  	    
				totnightWT_sum		=    totnightWT_re		       - totnightWT_pre		    
				totholidayWT_sum	=    totholidayWT_re	       - totholidayWT_pre	    
				totweekholidayWT_sum = totweekholidayWT_re    - totweekholidayWT_pre  
				totVacationTime_sum = totVacationTime_re      - totVacationTime_pre    
				 
				totWorkTime = 0
				totextendWT  = 0
				totnightWT	=0
				totholidayWT=0
				totweekholidayWT=0
				totVacationTime = 0
				
				stDCnt = intD
		%>
		
		<tr bgcolor="<%=adminColor("sky")%>" align="center"> 
				<td colspan="9"><b>B.</b> [<%=dPreMonth%>/26 ~ <%=dPreMonth%>/<%=preEndDay%>] <b>재계산</b></td>
				<td><input type="text" class="text" name="totRVT" style="border:0;background:#DFDFDF;color:<%IF totVacationTime_re  = 0  THEN %>gray<%ELSEIF  totVacationTime_re < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totVacationTime_re)%>"></td>
				<td><input type="text" class="text"  name="totRWT" style="border:0;background:#DFDFDF;color:<%IF totWorkTime_re  = 0  THEN %>gray<%ELSEIF  totWorkTime_re < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totWorkTime_re)%>"></td> 
				<td><input type="text" class="text"  name="totReWT" style="border:0;background:#DFDFDF;color:<%IF totextendWT_re  = 0  THEN %>gray<%ELSEIF  totextendWT_re < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totextendWT_re)%>"></td>
				<td><input type="text" class="text"  name="totRnWT" style="border:0;background:#DFDFDF;color:<%IF totnightWT_re  = 0  THEN %>gray<%ELSEIF  totnightWT_re < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totnightWT_re)%>"></td>
				<td><input type="text" class="text"  name="totRhWT" style="border:0;background:#DFDFDF;color:<%IF totholidayWT_re  = 0  THEN %>gray<%ELSEIF  totholidayWT_re < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totholidayWT_re)%>"></td>
				<td><input type="text" class="text"  name="totRwhWT" style="border:0;background:#DFDFDF;color:<%IF totweekholidayWT_re  = 0  THEN %>gray<%ELSEIF  totweekholidayWT_re < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totweekholidayWT_re)%>"></td>
		</tr> 
		<tr bgcolor="<%=adminColor("sky")%>" align="center"> 
				<td colspan="9"> <b> B - A = 차액</b></td>
				<td><input type="text" class="text" name="totSumVT" style="border:0;background:#DFDFDF;color:<%IF totVacationTime_sum  = 0  THEN %>gray<%ELSEIF  totVacationTime_sum < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totVacationTime_sum)%>"></td>
				<td><input type="text" class="text"  name="totSumWT" style="border:0;background:#DFDFDF;color:<%IF totWorkTime_sum  = 0  THEN %>gray<%ELSEIF  totWorkTime_sum < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totWorkTime_sum)%>"></td> 
				<td><input type="text" class="text"  name="totSumeWT" style="border:0;background:#DFDFDF;color:<%IF totextendWT_sum  = 0  THEN %>gray<%ELSEIF  totextendWT_sum < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totextendWT_sum)%>"></td>
				<td><input type="text" class="text"  name="totSumnWT" style="border:0;background:#DFDFDF;color:<%IF totnightWT_sum  = 0  THEN %>gray<%ELSEIF  totnightWT_sum < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totnightWT_sum)%>"></td>
				<td><input type="text" class="text"  name="totSumhWT" style="border:0;background:#DFDFDF;color:<%IF totholidayWT_sum  = 0  THEN %>gray<%ELSEIF  totholidayWT_sum < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totholidayWT_sum)%>"></td>
				<td><input type="text" class="text"  name="totSumwhWT" style="border:0;background:#DFDFDF;color:<%IF totweekholidayWT_sum  = 0  THEN %>gray<%ELSEIF  totweekholidayWT_sum < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totweekholidayWT_sum)%>"></td>
		</tr>
		<tr>
			<td colspan="15" bgcolor="#FFFFFF"></td>
		</tr>
		<%end if
		end if
		  totWorkTime 	= totWorkTime + iWorkTime
			totextendWT  	= totextendWT + iextendWT
			totnightWT		= totnightWT + inightWT
			totholidayWT	= totholidayWT + iholidayWT
			totweekholidayWT= totweekholidayWT  + iweekholidayWT

		'	currDate = Format00(4, dYear) + "-" + Format00(2, dMonth) + "-" + Format00(2, intD)
			if (sVacationYN = "Y") and (oVacation.FResultCount > 0) and (vacationRequestCount = 0) and (dstate = 0) then
				for j = 0 to oVacation.FResultCount - 1  
					if ((Cstr(dFullPayDate) >= Left(oVacation.FItemList(j).Fstartday, 10)) and (Cstr(dFullPayDate) <= Left(oVacation.FItemList(j).Fendday, 10))) then 
						' if (oVacation.FItemList(j).FmasterDivCD = "1") then
							'// 연차 = 유급휴가, 나머지 무급휴가
							iworktype = "4"  '//휴가종류 상관없이 유급휴가설정(2014-08-08 정윤정 수정)
						 
							iVacationTime = (oVacation.FItemList(j).Ftotalday/0.125)/(datediff("d",oVacation.FItemList(j).Fstartday,oVacation.FItemList(j).Fendday)+1) *60
						 
						 if dFullPayDate >= dateserial(dYear,dMonth,"26") then
							 	dstartHour 	= "00"
								dstartMinute= "00"
								dendHour 	= "00"
								dendMinute 	= "00"
								doutHour 	= "00"
								doutMinute 	= "00"
								dbreakSHour ="00"
								dbreakEHour ="00"
						 end if	 
						'else
						'	iworktype = "2"
						'end if
					end if
				next				
			end if
			totVacationTime = totVacationTime + iVacationTime
		%> 
		
		<% if (dWeekday = 1) then %>
		<tr   bgcolor="#CCCCCC" align="center"><td colspan="14" height="2"></td></tr>
		<% end if %>
		<tr   bgcolor="#FFFFFF"  align="center" >
			<td><div style="display:none;" id="i<%=dDay%>"><%=intD%></div><%=dDay%></td>
			<td><%=fnGetStringWD(dWeekday)%><input type="hidden" name="hidWD<%=intD%>" value="<%=dWeekday%>"></td>
			<td>
					<%IF iworktype  = "1" THEN%>
						근무일
					<%ELSEIF iworktype  = "2" THEN%>
						<font color="blue">무급휴일<font>
					<%ELSEIF iworktype  = "3" THEN%>
						<font color="red">주휴일</font>
					<%ELSEIF iworktype  = "6" THEN%>
						<font color="red">주휴일(무)</font>
					<%ELSEIF iworktype  = "7" THEN%>
						<font color="red">주휴일(유)</font>
					<%ELSEIF iworktype  = "4" THEN%>
						 		유급휴일
					<%ELSEIF iworktype  = "5" THEN%>
						 		공휴일
					<%ELSEIF iworktype  = "0" THEN%>
						 	<font color="Gray">입사전/퇴사후</font>
					<%END IF%>
			</td>
			<td> 
				<input type="text"  class="text" name="iSH<%=intD%>" value="<%=dstartHour%>" size="2" maxlength="2" style="text-align:right;border:0;" readonly>
				:
			 	<input type="text" class="text"  name="iSM<%=intD%>" value="<%=dstartMinute%>" size="2"  maxlength="2" style="text-align:right;border:0;" readonly>
			</td>
			<td>
				<input type="text" class="text"  name="iEH<%=intD%>" value="<%=dendHour%>" size="2"  maxlength="2" style="text-align:right;border:0;" readonly>
				:
			 	<input type="text" class="text"  name="iEM<%=intD%>" value="<%=dendMinute%>" size="2"  maxlength="2" style="text-align:right;border:0;" readonly>
			</td>
			<td>
				<input type="text" class="text"  name="iBSH<%=intD%>" value="<%=dbreakSHour%>"  size="2"  maxlength="2" style="text-align:right;border:0;" readonly>
				:
			 	<input type="text" class="text"  name="iBSM<%=intD%>" value="<%=dbreakSMinute%>" size="2"  maxlength="2" style="text-align:right;border:0;" readonly>
			</td>
			<td>
				<input type="text" class="text"  name="iBEH<%=intD%>"  value="<%=dbreakEHour%>" size="2"  maxlength="2" style="text-align:right;border:0;" readonly>
				:
			 	<input type="text" class="text"  name="iBEM<%=intD%>" value="<%=dbreakEMinute%>"  size="2"  maxlength="2" style="text-align:right;border:0;" readonly>
			</td>
			<td><input type="text"  class="text" name="iOH<%=intD%>"  value="<%=doutHour%>" size="2"  maxlength="2" style="text-align:right;border:0;" readonly>
				:
			 	<input type="text" class="text"  name="iOM<%=intD%>" value="<%=doutMinute%>"  size="2"  maxlength="2" style="text-align:right;border:0;" readonly></td>
			
			<td><input type="text" class="text"  name="dfT<%=intD%>" value="<%=defaulttime(dWeekday)%>" style="border:0;" readonly size="5"></td>
			<td><input type="text" class="text"  name="iVT<%=intD%>" value="<%=fnSetTimeFormat(iVacationTime)%>" style="border:0;" readonly size="5"></td> 	
			<td><b>(</b>&nbsp;<input type="text"  class="text" name="iWT<%=intD%>" style="border:0;color:<%IF iWorkTime  = 0  THEN %>gray<%ELSEIF  iWorkTime < 0 THEN%>red<%ELSE%>blue<%END IF%>;" readonly size="5" value="<%=fnSetTimeFormat(iWorkTime)%>"></td>
			<td><input type="text"  class="text" name="ieWT<%=intD%>" style="border:0;color:<%IF iextendWT  = 0  THEN %>gray<%ELSEIF  iextendWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(iextendWT)%>"><b>)</b></td>
			<td><input type="text"  class="text" name="inWT<%=intD%>" style="border:0;color:<%IF inightWT  = 0  THEN %>gray<%ELSEIF  inightWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(inightWT)%>"></td>
			<td><input type="text" class="text"  name="ihWT<%=intD%>" style="border:0;color:<%IF iholidayWT  = 0  THEN %>gray<%ELSEIF  iholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(iholidayWT)%>"></td>
			<td><input type="text"  class="text" name="iwhWT<%=intD%>" style="border:0;color:<%IF iweekholidayWT  = 0  THEN %>gray<%ELSEIF  iweekholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(iweekholidayWT)%>"></td>
		</tr> 
		<%	 
			  dFullPayDate =  dateadd("d",1,dFullPayDate)
			    
			Next
		set ofingerprints = nothing
		%>
		
		<%IF isArray(arrList) THEN
			if  arrList(0,ubound(arrList,2))  = dYear&"-"&format00(2,dMonth)&"-32" THEN
				totweekholidayWT = totweekholidayWT + arrList(11,ubound(arrList,2))
			%>
		<tr   bgcolor="#FFFFFF"  align="center">
			<td colspan="13"> 추가주휴수당 </td>
			<td><div id="dNMWT" style="display:;"><input type="text"  class="text" name="iwhWT40"  id="iwhWT40" style="border:0;color:blue"  size="5" value="<%=fnSetTimeFormat(arrList(11,ubound(arrList,2)))%>"></div></td>
	 	</tr>
	 	<%else%>
	 	<tr>
			<td><div id="dNMWT" style="display:none;"><input type="text"  class="text"  name="iwhWT40" id="iwhWT40" value="0"></div></td>
		</tr>
	 	<%end if%>
		<%ELSE%>
		<tr>
			<td><div id="dNMWT" style="display:none;"><input type="text"  class="text" name="iwhWT40"  id="iwhWT40" value="0"></div></td>
		</tr>
		<% 
		END IF%>
		<input type="hidden" name="hidSday" value="<%=stDCnt%>"><!-- 급여일수--> 
		<input type="hidden" name="hidEday" value="<%=iLoopCnt%>"><!-- 급여일수--> 
		<%if   chkDate >= "2017-01" then%> 
		<tr   bgcolor="<%=adminColor("sky")%>" align="center">
			<td colspan="9">  [<%=dMonth%>/1 ~ <%=dMonth%>/<%=dDay%>]  <b>합계</b></td> 
			<td><input type="text" class="text"  name="totVT" style="border:0;background:#DDDDFF;color:<%IF totVacationTime  = 0  THEN %>gray<%ELSEIF  totVacationTime < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totVacationTime)%>"></td>
			<td><input type="text" class="text"  name="totWT" style="border:0;background:#DDDDFF;color:<%IF totWorkTime  = 0  THEN %>gray<%ELSEIF  totWorkTime < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totWorkTime)%>"></td> 
			<td><input type="text" class="text"  name="toteWT" style="border:0;background:#DDDDFF;color:<%IF totextendWT  = 0  THEN %>gray<%ELSEIF  totextendWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totextendWT)%>"></td>
			<td><input type="text" class="text"  name="totnWT" style="border:0;background:#DDDDFF;color:<%IF totnightWT  = 0  THEN %>gray<%ELSEIF  totnightWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totnightWT)%>"></td>
			<td><input type="text" class="text"  name="tothWT" style="border:0;background:#DDDDFF;color:<%IF totholidayWT  = 0  THEN %>gray<%ELSEIF  totholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totholidayWT)%>"></td>
			<td><input type="text" class="text"  name="totwhWT" style="border:0;background:#DDDDFF;color:<%IF totweekholidayWT  = 0  THEN %>gray<%ELSEIF  totweekholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totweekholidayWT)%>"></td>
		</tr> 
		<%
		totVacationTime = totVacationTime + totVacationTime_sum
		totWorkTime 	=   totWorkTime     + totWorkTime_sum      
		totextendWT		=   totextendWT  	  + totextendWT_sum  	  
		totnightWT		=   totnightWT			+ totnightWT_sum		    
		totholidayWT	=   totholidayWT    + totholidayWT_sum	    
		totweekholidayWT =totweekholidayWT+ totweekholidayWT_sum 

		 end if%>
		<tr   bgcolor="<%=adminColor("sky")%>" align="center">
			<td colspan="9">  <%if   chkDate >= "2017-01" then%> [<%=dpreMonth%>/26 ~ <%=dMonth%>/<%=dDay%>]<%end if%>   <b>총 합계</b></td> 
			<td><input type="text" class="text"  name="totSVT" style="font-weight:bold;border:0;background:#DDDDFF;color:<%IF totVacationTime  = 0  THEN %>gray<%ELSEIF  totVacationTime < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totVacationTime)%>"></td>
			<td><input type="text" class="text"  name="totSWT" style="font-weight:bold;border:0;background:#DDDDFF;color:<%IF totWorkTime  = 0  THEN %>gray<%ELSEIF  totWorkTime < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totWorkTime)%>"></td> 
			<td><input type="text" class="text"  name="totSeWT" style="font-weight:bold;border:0;background:#DDDDFF;color:<%IF totextendWT  = 0  THEN %>gray<%ELSEIF  totextendWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totextendWT)%>"></td>
			<td><input type="text" class="text"  name="totSnWT" style="font-weight:bold;border:0;background:#DDDDFF;color:<%IF totnightWT  = 0  THEN %>gray<%ELSEIF  totnightWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totnightWT)%>"></td>
			<td><input type="text" class="text"  name="totShWT" style="font-weight:bold;border:0;background:#DDDDFF;color:<%IF totholidayWT  = 0  THEN %>gray<%ELSEIF  totholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totholidayWT)%>"></td>
			<td><input type="text" class="text"  name="totSwhWT" style="font-weight:bold;border:0;background:#DDDDFF;color:<%IF totweekholidayWT  = 0  THEN %>gray<%ELSEIF  totweekholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totweekholidayWT)%>"></td>
		</tr> 
		
		</table>
	</td>
</tr>
</table>
</body>
</html>

	<script type="text/javascript">
	var chk = 0;
	window.onload = function() {
		jsSetHolidayWD(<%= holidaywdtime %>);
	 
		if(chk==0){
			jsSetTotTimeALL(<%= intD%>);
			chk = 1;
		}
		 
	}
</script>
		