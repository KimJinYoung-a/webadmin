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
<!-- #include virtual="/lib/classes/member/fingerprints/fingerprints_cls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%

'// 현재 페이지에서는 시간과 관련된 내용만 계산한다.
'// 금액 관련된 내용은 처리하지 않는다.

'변수선언
Dim sempno, ino
Dim djoinday, susername, iposit_sn,sposit_name,blnstatediv,dretireday
Dim startdate,enddate ,defaultpay,foodpay,jobpay,inbreaktime,holidaywdtime
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
'값 받아오기
sempno= requestCheckvar(Request("sEN"),14)
ino= requestCheckvar(Request("ino"),10)
dYear = requestCheckvar(Request("selY"),4)
dMonth = requestCheckvar(Request("selM"),2)
chkDate = dYear&"-"&format00(2,dMonth)
sFingerYN = requestCheckvar(Request("sFYN"),1) '지문인식근태관리 가져왔는지 여부
sVacationYN = requestCheckvar(Request("sVYN"),1)

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
ELSEIF chkDate = "2016-12"   then
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
<link rel="stylesheet" href="/scm.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jsPayCal.js"></script> 
<script type="text/javascript">
<!--
	function jsSearch(){
		var dNowYear, dNowMonth;
		var date = new Date();
		dNowYear = date.getFullYear();
		dNowMonth = date.getMonth() + 1;

//	 	if (document.frmSearch.selY.value > dNowYear){
//	 		alert("현재 달 이전까지만  검색 가능합니다.");
//	 		return;
//	 	}else if (document.frmSearch.selY.value == dNowYear && document.frmSearch.selM.value > dNowMonth){
//	 		alert("현재 달 이전까지만  검색 가능합니다.");
//	 		return;
//	 	}
 document.frmSearch.sFYN.value  ="";
	  document.frmSearch.sVYN.value  = "";
	 	document.frmSearch.submit();
	}

	function jsSubmitPay(){

		var swday = 0;
		var blnPWH, hidPWHD, itwt, arrValue, iValue;
		var startj,endj;
		var LastDayOfThisMonth;
		var startMinFromMidnight, endMinFromMidnight;
		var hidCWD, hidCWHD, blnCWH;
		var startDayOfCurrWeek;

		// * 주휴시간
		//  - 일주일은 일요일부터 토요일까지이다.
		//  - 무단 결근시 주휴시간 0 시간
		//  - 주간 근무시간 15시간 미만일 경우 주휴시간 0 시간
		//  - 15시간 이상 근무시 주휴시간 = (시간/40)*8 시간
		//  - 근무시간 40시간 초과시 주휴시간 8 시간

		// * 주휴일 갯수만 체크한다.
		//  - 주휴시간 입력은 jsSetTotTimeALL() 에서 한다.

		LastDayOfThisMonth = document.frmPay.hidEday.value*1; 
		for(var i = 0; i <= LastDayOfThisMonth; i++) {
			
			// =================================================================
			// 근무시간 체크
			startMinFromMidnight = parseInt(eval("document.frmPay.iSH"+i).value,10)*60+parseInt(eval("document.frmPay.iSM"+i).value,10);
			endMinFromMidnight = parseInt(eval("document.frmPay.iEH"+i).value,10)*60+parseInt(eval("document.frmPay.iEM"+i).value,10);
 
			if( startMinFromMidnight > endMinFromMidnight) {
				alert("근무시작시간은 근무종료시간보다 빨라야합니다. 다시 설정해주세요");
				eval("document.frmPay.iSH"+i).focus();
				return false;
			}
 
			// =================================================================
			// 휴게시간 체크
			startMinFromMidnight = parseInt(eval("document.frmPay.iBSH"+i).value,10)*60+parseInt(eval("document.frmPay.iBSM"+i).value,10);
			endMinFromMidnight = parseInt(eval("document.frmPay.iBEH"+i).value,10)*60+parseInt(eval("document.frmPay.iBEM"+i).value,10);

			if( startMinFromMidnight >  endMinFromMidnight) {
				alert("휴게시작시간은 휴게종료시간보다 빨라야합니다. 다시 설정해주세요");
				eval("document.frmPay.iBSH"+i).focus();
				return false;
			}

			// =================================================================
			// 그 달에  첫번째 일욜일에 해당하는 날짜
			if(eval("document.frmPay.hidWD"+i).value  == 1 && swday == 0) {
				swday =  i;
			}
		
		}
 
		itwt = 0;
		blnPWH = 0;

		itwt = document.frmPay.hidPWD.value*1; 			// 전달 마지막주 총근무시간
		blnPWH = document.frmPay.blnPWH.value*1; 			// 전달 마지막주 주휴일수
		hidPWHD = document.frmPay.hidPWHD.value*1; 		// 전달 마지막주 결근횟수

		hidCWD = document.frmPay.hidCWD.value*1;		// 이번주 전달부분
		blnCWH = document.frmPay.blnCWH.value*1;
		hidCWHD = document.frmPay.hidCWHD.value*1;

		if (blnPWH > 1) {
			// 전달 주휴일자 2일 이상 입력된것 무시
			blnPWH = 1;
		}

		if (blnCWH > 1) {
			// 이번주 전달부분 주휴일자 2일 이상 입력된것 무시
			blnCWH = 1;
		}

		// =====================================================================
		// 01. 첫째 주 체크
		// =====================================================================
		if (swday == 1) {
			// 현재 달의 1일이 일요일인 경우
		} else {
			// 달의 첫번째 날이 일요일인 아닌 경우
			// * 전달 마지막 일요일 이후의 날짜와 현재달의 첫번째 토요일까지의 날짜를 합쳐서 주휴시간 체크한다.
	 		for (var i = 0; i < swday; i++) {
				// 주휴일수
		 		if (eval("document.frmPay.selWH"+i).value == "3") {
		 	  		blnCWH = blnCWH + 1;
		 		}

				// 근무시간
				arrValue = eval("document.frmPay.iWT"+i).value.split(":");
				iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
				hidCWD = hidCWD + iValue;

				if ((eval("document.frmPay.selWH"+i).value == "1") && (iValue == 0)) {
					hidCWHD = hidCWHD + 1;
				}
			}

			if (blnCWH > 1) {
				alert("주휴일(첫째주)은 일주일에 한번만 설정 가능합니다.");
				return false;
			} else if ((itwt >= 900) && (blnCWH < 1)) {
				alert("주휴일(첫째주)을 설정해주세요.");
				return false;
			}
		}

		// =====================================================================
		// 02. 첫째 주 이후(또는 1일이 일요일인 경우) 체크
		// =====================================================================

		if (swday != 1) {
			itwt = hidCWD;
			blnPWH = blnCWH;
			hidPWHD = hidCWHD;
		}

		hidCWD = 0;
		blnCWH = 0;
		hidCWHD = 0;
		for (var i = swday; i <= LastDayOfThisMonth; i++) {
			if (eval("document.frmPay.selWH"+i).value=="3") {
			  	blnCWH = blnCWH + 1;
			}

			// 근무시간
			arrValue = eval("document.frmPay.iWT"+i).value.split(":");
			iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
			hidCWD = hidCWD + iValue;

			if ((eval("document.frmPay.selWH"+i).value == "1") && (iValue == 0)) {
				hidCWHD = hidCWHD + 1;
			}

			if ((eval("document.frmPay.hidWD"+i).value*1 == 7) || (i == LastDayOfThisMonth)) {
				if (blnCWH > 1) {
					alert("주휴일은 일주일에 한번만 설정 가능합니다.");
					return false;
				} else if ((itwt >= 900) && (blnCWH < 1)) {
				    if (i*1!=LastDayOfThisMonth*1){ //추가
					    alert("주휴일을 설정해주세요...." + itwt + " " + blnCWH + " " +LastDayOfThisMonth+ " " +i);
					    return false;
					}
				}

				itwt = hidCWD;
				blnPWH = blnCWH;
				hidPWHD = hidCWHD;

				hidCWD = 0;
				blnCWH = 0;
				hidCWHD = 0;
			}
		}
 
		jsAddWH(); 
		jsSetTotTimeALL(LastDayOfThisMonth + 1); //등록전 재계산처리
 
	 	return true;
	 
	}

	function jsAddWH(){
	 //퇴사자 주휴일 설정
		<%IF blnstatediv = "N" THEN
			IF   dretireday >= dSPayDate and dretireday<= dEPayDate THEN	'퇴사자이고 퇴사일이 이번달 근무일에 포함된  경우%>

			if ((<%=day(dretireday)%>+(7-<%=weekday(dretireday)%>)) >= document.frmPay.hidEPday.value){ //퇴사주 마지막일이 퇴사일보다 크거나 같은 경우
			 
				var iLwt = 0;
				var	iValue = 0;
				var	iEValue = 0;
				var	arrValue = "";
				var chkWHD  = 0;
				var iLWHD = 0;
				var iLDutyH = 0;
				var iLWHTime = 0;
				var iLSday = $("#i<%=day(dretireday)-weekday(dretireday)+1%>").text();
				var iLEday = $("#i<%=day(dretireday)%>").text(); 
					for(i=iLSday;i<=iLEday;i++){ //퇴사주 주 근무시간 계산
						if(eval("document.frmPay.iWT"+i).value!=""){
							arrValue = eval("document.frmPay.iWT"+i).value.split(":");
							iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);

							arrValue = eval("document.frmPay.ieWT"+i).value.split(":");
							iEValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
						}

						if( eval("document.frmPay.selWH"+i).value=="1" && ( iValue+iEValue) == 0 ){//근무일에 근무했는지 여부 확인
							chkWHD = chkWHD + 1 ;
						}
						if( eval("document.frmPay.selWH"+i).value=="3"){
							iLWHD=i;
						}

						iLwt = iLwt + iValue+iEValue; 
					}
 
					if(chkWHD==0){
						iLDutyH = parseInt(iLwt/60,10); 
						if (iLDutyH > 40){
							iLDutyH=40
						};
						if(iLDutyH>=15){
							iLWHTime= (8*(iLDutyH/40))*60;
						}
					}

					chkWHD = 0
					iLwt = 0
					if (iLWHD == 0){ //근무 마지막주에 주휴일이 없을 경우 전주 근무 주휴수당 포함에서 준다.
						iLSday = parseInt(iLSday,10)-7;
						iLEday = parseInt(iLSday,10)+6;

						for(i=iLSday;i<=iLEday;i++){
							if(eval("document.frmPay.iWT"+i).value!=""){
								arrValue = eval("document.frmPay.iWT"+i).value.split(":");
								iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);

								arrValue = eval("document.frmPay.ieWT"+i).value.split(":");
								iEValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
							}

							if( eval("document.frmPay.selWH"+i).value=="1" && ( iValue+iEValue) == 0 ){//근무일에 근무했는지 여부 확인
								chkWHD  = chkWHD  + 1 ;
							}
							iLwt  = iLwt  + iValue+iEValue;
						}
					}

					if(chkWHD==0){
						iLDutyH = parseInt(iLwt/60,10);
						if (iLDutyH > 40){
							iLDutyH=40;
						}
						if(iLDutyH>=15){
							iLWHTime= iLWHTime+ (8*(iLDutyH/40))*60;
						}
					}

					document.frmPay.iwhWT40.value = jsTimeForm(iLWHTime);
					document.all.dNMWT.style.display = "";

					var totwhWT = document.frmPay.totwhWT.value.split(":");
					totwhWT = parseInt(totwhWT[0],10)*60+parseInt(totwhWT[1],10)+iLWHTime;
					document.frmPay.totwhWT.value =  jsTimeForm(totwhWT);
				}

		<%END IF%>
	<%END IF%>
	}

	function jsComplete(){
		if(confirm("근무시간등록을 작성완료하시겠습니까? 작성완료시 월급여가 생성됩니다.")){
			document.frmPay.hidS.value="1";
			document.frmPay.submit();
		}
		return;
	}

	//지문인식근태 내역 가져오기
	function jsGetFinger(){
		document.frmSearch.sFYN.value  = "Y";
		document.frmSearch.submit();
	}

	//휴가신청내역 가져오기
	function jsGetVacation(){
		document.frmSearch.sVYN.value  = "Y";
		document.frmSearch.submit();
	}

	function jsSetTotTimeALL(ilen){   
	    for(var j=0;j<=ilen-1;j++){  
	        jsSetTotTime(j);
	       // alert(j+"-"+eval("document.frmPay.ieWT"+j).value);
	    }
	    //alert('Fin');
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
			<td bgcolor="#FFFFFF"><%=formatnumber(defaultpay,0)%> 원</td>
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
			<td align="right">
				<%IF mstate = 0 THEN%>
					<input type="button" class="button" value="휴가신청내역 가져오기" onClick="jsGetVacation();">
					<input type="button" class="button" value="지문인식근태내역 가져오기" onClick="jsGetFinger();">
				<%END IF%>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<form name="frmPay" method="post" action="tenbyten_pay_process.asp" onSubmit="return jsSubmitPay();">
		<input type="hidden" name="hidEN" value="<%=sempno%>">
		<input type="hidden" name="ino" value="<%=ino%>">
		<input type="hidden" name="hidM" value="D">
		<input type="hidden" name="hidS" value="0">
		<input type="hidden" name="hidPSN" value="<%=iposit_sn%>">
		<input type="hidden" name="hidInB" value="<%=inbreaktime%>">
		<input type="hidden" name="hidSPday" value="<%= dSPayDay %>"><!-- 급여시작일-->
		<input type="hidden" name="hidEPday" value="<%= dEPayDay %>"><!-- 급여종료일--> 
		<input type="hidden" name="hidSPdate" value="<%= dSPayDate %>"><!-- 급여시작일자 년월일-->
		<input type="hidden" name="hidEPdate" value="<%= dEPayDate %>"><!-- 급여종료일자 년월일-->
		<input type="hidden" name="hidEday" value="<%=iLoopCnt%>"><!-- 급여일수--> 
		<input type="hidden" name="hidVer" value="<%=iVer%>">
		<input type="hidden" name="hidYear" value="<%=dYear%>"> 
		<input type="hidden" name="hidMonth" value="<%=dMonth%>"> 
		<input type="hidden" name="hidDP" value="<%=defaultpay%>"> 
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
					<%ELSEIF arrPre(5,intD)  = "4" THEN%>
						<font color="red">유급휴일<font>
					<%ELSEIF arrPre(5,intD)  = "5" THEN%>
						<font color="red">공휴일<font>
					<%END IF%>
					<input type="hidden" name="iPWH<%=day(arrPre(0,intD))%>" value="<%=arrPre(5,intD)%>">
				</td>
				<td><input type="text" name="iPWS<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(1,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF " readonly></td>
				<td><input type="text" name="iPWE<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(2,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
				<td><input type="text" name="iPBS<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(3,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
				<td><input type="text" name="iPBE<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(4,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
				<td><input type="text" name="iPO<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(12,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
				<td></td>
				<td><input type="text" name="iPVT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iVacationTime)%>"></td>
				<td><input type="text" name="iPWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iWorkTime)%>"></td>
				<td><input type="text" name="iPeWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iextendWT)%>"></td>
				<td><input type="text" name="iPnWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(inightWT)%>"></td>
				<td><input type="text" name="iPhWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iholidayWT)%>"></td>
				<td><input type="text" name="iPwhWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iweekholidayWT)%>"></td>
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
			iWorkTime = 0
			iextendWT  = 0
			inightWT	=0
			iholidayWT=0
			iweekholidayWT=0   
			iVacationTime = 0 
			dDay = day(dFullPayDate)
			dWeekday = weekday(dFullPayDate) 
		   
		
 			'계약정보 가져오기-------------------- 
			IF djoinday  > dFullPayDate  or enddate < dFullPayDate or dretireday < dFullPayDate THEN 
				dbreakSHour = "00"
				dbreakSMinute ="00"
				dbreakEHour = "00"
				dbreakEMinute ="00"
	 			iworktype = 0 
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
 			END IF  
 			'--------------------------------------- 
 
			IF sFingerYN = "Y" THEN     '지문인식근태내역 가져오기
			 	dstartHour 	= "00"
				dstartMinute= "00"
				dendHour 	= "00"
				dendMinute 	= "00"
				doutHour 	= "00"
				doutMinute 	= "00"
				arrWorkTime(dDay) = 0
				arrWorkType(dDay) = iworktype
				
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
				ELSE
				  
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
			 	END IF
				END IF
			END IF

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
				<%IF dstate > 0 THEN%>
					<%IF iworktype  = "1" THEN%>
						근무일
					<%ELSEIF iworktype  = "2" THEN%>
						<font color="blue">무급휴일<font>
					<%ELSEIF iworktype  = "3" THEN%>
						<font color="red">주휴일</font>
					<%ELSEIF iworktype  = "4" THEN%>
						 		유급휴일
					<%ELSEIF iworktype  = "5" THEN%>
						 		공휴일
					<%ELSEIF iworktype  = "0" THEN%>
						 	<font color="Gray">입사전/퇴사후</font>
					<%END IF%>
				<%ELSE%> 
				<select name="selWH<%=intD%>"  onChange="jsChangeWeekHoliday_Pre(<%=dWeekday%>,<%=intD%>);jsSetTotTime(<%=intD%>);">
				<option value="1" <%IF iworktype ="1"  THEN%>selected<%END IF%>>근무일</option>
				<option value="2" <%IF iworktype ="2" THEN%>selected<%END IF%> style="color:blue">무급휴일</option>
				<option value="3" <%IF iworktype ="3" THEN%>selected<%END IF%> style="color:red">주휴일</option>
				<option value="4" <%IF iworktype ="4" THEN%>selected<%END IF%> style="color:red">유급휴일</option>
				<option value="5" <%IF iworktype ="5" THEN%>selected<%END IF%> style="color:red">공휴일</option>
				<option value="0" <%IF iworktype ="0" THEN%>selected<%END IF%>  style="color:gray">입사전/퇴사후</option>
				</select>
				<%END IF%>
			</td>
			<td> 
				<input type="text" name="iSH<%=intD%>" value="<%=dstartHour%>" size="2" maxlength="2" style="text-align:right;<%IF dstate > 0 THEN%> border:0;" readonly<%ELSE%>"<%END IF%>  onKeyUp="jsSetTotTime(<%=dDay%>);TnTabNumber('iSH<%=intD%>','iSM<%=intD%>',2);">
				:
			 	<input type="text" name="iSM<%=intD%>" value="<%=dstartMinute%>" size="2"  maxlength="2" style="text-align:right;<%IF dstate > 0 THEN%> border:0;" readonly<%ELSE%>"<%END IF%>    onKeyUp="jsSetTotTime(<%=intD%>);TnTabNumber('iSM<%=intD%>','iEH<%=intD%>',2);">
			</td>
			<td>
				<input type="text" name="iEH<%=intD%>" value="<%=dendHour%>" size="2"  maxlength="2" style="text-align:right;<%IF dstate > 0 THEN%> border:0;" readonly<%ELSE%>"<%END IF%>   onKeyUp="jsSetTotTime(<%=intD%>);TnTabNumber('iEH<%=intD%>','iEM<%=intD%>',2);">
				:
			 	<input type="text" name="iEM<%=intD%>" value="<%=dendMinute%>" size="2"  maxlength="2" style="text-align:right;<%IF dstate > 0 THEN%> border:0;" readonly<%ELSE%>"<%END IF%>    onKeyUp="jsSetTotTime(<%=intD%>);TnTabNumber('iEM<%=intD%>','iBSH<%=intD%>',2);">
			</td>
			<td>
				<input type="text" name="iBSH<%=intD%>" value="<%=dbreakSHour%>"  size="2"  maxlength="2" style="text-align:right;<%IF dstate > 0 THEN%> border:0;" readonly<%ELSE%>"<%END IF%>   onKeyUp="jsSetTotTime(<%=intD%>);TnTabNumber('iBSH<%=intD%>','iBSM<%=intD%>',2);">
				:
			 	<input type="text" name="iBSM<%=intD%>" value="<%=dbreakSMinute%>" size="2"  maxlength="2" style="text-align:right;<%IF dstate > 0 THEN%> border:0;" readonly<%ELSE%>"<%END IF%>    onKeyUp="jsSetTotTime(<%=intD%>);TnTabNumber('iBSM<%=intD%>','iBEH<%=intD%>',2);">
			</td>
			<td>
				<input type="text" name="iBEH<%=intD%>"  value="<%=dbreakEHour%>" size="2"  maxlength="2" style="text-align:right;<%IF dstate > 0 THEN%> border:0;" readonly<%ELSE%>"<%END IF%>   onKeyUp="jsSetTotTime(<%=intD%>);TnTabNumber('iBEH<%=intD%>','iBEM<%=intD%>',2);">
				:
			 	<input type="text" name="iBEM<%=intD%>" value="<%=dbreakEMinute%>"  size="2"  maxlength="2" style="text-align:right;<%IF dstate > 0 THEN%> border:0;" readonly<%ELSE%>"<%END IF%>   onKeyUp="jsSetTotTime(<%=intD%>);TnTabNumber('iBEM<%=intD%>','iOH<%=intD%>',2);">
			</td>
			<td><input type="text" name="iOH<%=intD%>"  value="<%=doutHour%>" size="2"  maxlength="2" style="text-align:right;<%IF dstate > 0 THEN%> border:0;" readonly<%ELSE%>"<%END IF%>   onKeyUp="jsSetTotTime(<%=intD%>);TnTabNumber('iOH<%=intD%>','iOM<%=intD%>',2);">
				:
			 	<input type="text" name="iOM<%=intD%>" value="<%=doutMinute%>"  size="2"  maxlength="2" style="text-align:right;<%IF dstate > 0 THEN%> border:0;" readonly<%ELSE%>"<%END IF%>   onKeyUp="jsSetTotTime(<%=intD%>);<%IF  intD  < dEndDay THEN%>TnTabNumber('iOM<%=intD%>','iSH<%=intD+1%>',2);<%END IF%>"></td>
			
			<td><input type="text" name="dfT<%=intD%>" value="<%=defaulttime(dWeekday)%>" style="border:0;" readonly size="5"></td>
			<td><input type="text" name="iVT<%=intD%>" value="<%=fnSetTimeFormat(iVacationTime)%>" style="border:0;" readonly size="5"></td> 	
			<td><b>(</b>&nbsp;<input type="text" name="iWT<%=intD%>" style="border:0;color:<%IF iWorkTime  = 0  THEN %>gray<%ELSEIF  iWorkTime < 0 THEN%>red<%ELSE%>blue<%END IF%>;" readonly size="5" value="<%=fnSetTimeFormat(iWorkTime)%>"></td>
			<td><input type="text" name="ieWT<%=intD%>" style="border:0;color:<%IF iextendWT  = 0  THEN %>gray<%ELSEIF  iextendWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(iextendWT)%>"><b>)</b></td>
			<td><input type="text" name="inWT<%=intD%>" style="border:0;color:<%IF inightWT  = 0  THEN %>gray<%ELSEIF  inightWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(inightWT)%>"></td>
			<td><input type="text" name="ihWT<%=intD%>" style="border:0;color:<%IF iholidayWT  = 0  THEN %>gray<%ELSEIF  iholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(iholidayWT)%>"></td>
			<td><input type="text" name="iwhWT<%=intD%>" style="border:0;color:<%IF iweekholidayWT  = 0  THEN %>gray<%ELSEIF  iweekholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(iweekholidayWT)%>"></td>
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
			<td><div id="dNMWT" style="display:;"><input type="text" name="iwhWT40"  id="iwhWT40" style="border:0;color:blue"  size="5" value="<%=fnSetTimeFormat(arrList(11,ubound(arrList,2)))%>"></div></td>
	 	</tr>
	 	<%else%>
	 	<tr>
			<td><div id="dNMWT" style="display:none;"><input type="text" name="iwhWT40" id="iwhWT40" value="0"></div></td>
		</tr>
	 	<%end if%>
		<%ELSE%>
		<tr>
			<td><div id="dNMWT" style="display:none;"><input type="text" name="iwhWT40"  id="iwhWT40" value="0"></div></td>
		</tr>
		<% 
		END IF%>
		<tr   bgcolor="<%=adminColor("sky")%>" align="center">
			<td colspan="9">합계</td> 
			<td><input type="text" name="totVT" style="border:0;background:#DFDFDF;color:<%IF totVacationTime  = 0  THEN %>gray<%ELSEIF  totVacationTime < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totVacationTime)%>"></td>
			<td><input type="text" name="totWT" style="border:0;background:#DFDFDF;color:<%IF totWorkTime  = 0  THEN %>gray<%ELSEIF  totWorkTime < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totWorkTime)%>"></td> 
			<td><input type="text" name="toteWT" style="border:0;background:#DFDFDF;color:<%IF totextendWT  = 0  THEN %>gray<%ELSEIF  totextendWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totextendWT)%>"></td>
			<td><input type="text" name="totnWT" style="border:0;background:#DFDFDF;color:<%IF totnightWT  = 0  THEN %>gray<%ELSEIF  totnightWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totnightWT)%>"></td>
			<td><input type="text" name="tothWT" style="border:0;background:#DFDFDF;color:<%IF totholidayWT  = 0  THEN %>gray<%ELSEIF  totholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totholidayWT)%>"></td>
			<td><input type="text" name="totwhWT" style="border:0;background:#DFDFDF;color:<%IF totweekholidayWT  = 0  THEN %>gray<%ELSEIF  totweekholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totweekholidayWT)%>"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
	<%IF dstate =  "" THEN%>
		<input type="submit" class="button" value="등록">
	<%ELSEIF dstate =  "0" THEN%>
		<input type="submit" class="button" value="수정">
	<%END IF%>
    </td>
</tr>
<tr>
    <td align="right"> <input type="button" value="재계산" onClick="jsSetTotTimeALL(<%=intD%>)"></td>
</tr>
</form>
</table>
</body>
</html>

	<script type="text/javascript">
	var chk = 0;
	window.onload = function() {
	//	jsSetHolidayWD(<%= holidaywdtime %>);
	 
		if(chk==0){
			jsSetTotTimeALL(<%= intD%>);
			chk = 1;
		}
		 
	}
</script>
		