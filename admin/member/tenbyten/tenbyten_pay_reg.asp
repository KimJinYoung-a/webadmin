<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사원별 급여 정보 등록
' History : 2010.12.27 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenPayCls.asp" -->
<%
Dim intY, intM, dYear, dMonth
Dim sEmpno, sUsername, dJoinday, blnstatediv,iposit_sn,sposit_name,dretireday,holidaywdtime,ino,startdate,enddate,defaultpay,foodpay,jobpay,inbreaktime,iDefaultPaySeq,predefaultpay,prefoodpay
Dim iworktime,iextendtime,inighttime,iholidaytime ,mtimepay,mextendpay ,mnightpay ,mholidaypay, mwholidaypay,mfoodpay,mjobpay ,mlongtimepay       ,maddpay, myearpay, mbonuspay
dim mworkday
Dim moutstandingpay,mtotpay,mnpensionpay,mhealthinspay,mrecuinspay,munempinspay,mtaxtotpay,mrealtotpay,dregdate,sadminid,istate
Dim  clsPay
Dim arrList , arrPre, intLoop
dim totDutyTime, totNightTime,totPaySum,avgWeek,iOverTime
dim arrdtime, idtime,dSWD,dEWD,totWD,iWD,dNStart,dNEnd,dNBreakS,dNBreakE, iweekholidaytime,totPWD, dweekday,totWH,iWT
dim dNextDate ,dEndDay,dREday,chkWHD,dEndDate, blnReset
Dim totWorkDay, totWorkDayReal
dim monthlyPayDataExist
dim iReworktime,iReextendtime,iRenighttime,iReholidaytime,iReweekholidaytime
dim iRefoodtime,mReExtimepay,mReNTtimepay,mReHDtimepay,mReFtimepay
dim mretimepay,mreextendpay,mrenightpay,mreholidaypay, mrefoodpay,mretotpay
dim intP, iP,totReWorkDay ,totReWorkDayReal
dim iTReworktime,iTReextendtime,iTRenighttime,iTReholidaytime,iTReweekholidaytime ,ireworkday

sEmpno	= requestCheckvar(request("sEN"),14)	'사번
dYear	= requestCheckvar(request("selY"),4)	'년
dMonth	= requestCheckvar(request("selM"),2)	'월
ino		= requestCheckvar(request("ino"),10)	'회차
blnReset = requestCheckvar(request("blnR"),1) '리셋여부
'기본값 설정 (현재 년월)
IF dYear = "" THEN dYear  = Year(Date())
IF dMonth = "" THEN dMonth  = Month(Date())

'// 4.345238095 == 월 평균 WEEK 수 = (365일 / 7일 / 12개월)
avgWeek = 4.345238095

totWorkDay = 0
totWorkDayReal = 0
totReWorkDay = 0
totReWorkDayReal =0

monthlyPayDataExist = True
dim dSPayDate,dEPayDate,dPreYear,dPreMonth ,preEndDay
dim chkDate 

preEndDay = dateadd("d", -1, dateserial(dYear,dMonth,1)) '이전달  마지막 일 
dPreYear = year(preEndDay) '이전달 년
dPreMonth = month(preEndDay) '이전달 월
dNextDate = dateadd("m",1, dateserial(dYear,dMonth,1))	'검색다음달 1일
dEndDate = dateadd("d",-1,dNextDate)
dEndDay = day(dEndDate)
chkDate =dYear&"-"&format00(2,dMonth)

'------------------------------------------------------------------ 
IF chkDate  = "2014-01" THEN '2014.01부터 급여종료일 25일로 변경됨 
	dSPayDate = dateserial(dYear,dMonth,1) '급여시작일: 해당월 1일부터
	dEPayDate = dateserial(dYear,dMonth,25) '급여종료일: 해당월 25일까지  
ELSEIF chkDate > "2014-01" and chkDate < "2016-12" THEN 
	dSPayDate = dateserial(dPreYear,dPreMonth,26) '급여시작일: 이전월 26일부터
	dEPayDate = dateserial(dYear,dMonth,25) '급여종료일: 해당월 25일까지   
ELSEIF chkDate >= "2016-12"	then '전달 26일부터~ 해당월 말일까지
	dSPayDate = dateserial(dPreYear,dPreMonth,26) 
	dEPayDate = dateserial(dYear,dMonth,dEndDay)  '급여종료일: 해당월 말일까지 
ELSE    
	dSPayDate = dateserial(dYear,dMonth,1) '급여시작일: 해당월 1일부터
	dEPayDate = dateserial(dYear,dMonth,dEndDay)  '급여종료일: 해당월 말일까지 
END IF  
'------------------------------------------------------------------ 
set clsPay = new CPay
	'// ========================================================================
	'// 사원 기본계약정보
	'// ========================================================================
	clsPay.Fempno = sempno
	clsPay.Fyyyymm = dYear&"-"&format00(2,dMonth)
	clsPay.Fino	= ino
	clsPay.fnGetUserPayData

	sempno		= clsPay.Fempno
	susername	= clsPay.Fusername
	djoinday	= clsPay.Fjoinday
	blnstatediv = clsPay.Fstatediv
	iposit_sn	= clsPay.Fposit_sn
	sposit_name = clsPay.Fposit_name
	dretireday	= clsPay.Fretireday

	holidaywdtime = clsPay.Fholidaywdtime
	ino						= clsPay.Fino
	startdate			= clsPay.Fstartdate
	enddate				= clsPay.Fenddate
	defaultpay  	= clsPay.Fdefaultpay
	foodpay	    	= clsPay.Ffoodpay
	jobpay				= clsPay.Fjobpay
	inbreaktime		= clsPay.FinBreakTime
	iDefaultPaySeq= clsPay.Fdefaultpayseq
	iOverTime			= clsPay.Fovertime
	predefaultpay = clsPay.FpreDefaultpay
	prefoodpay		= clsPay.FpreFoodpay

	totDutyTime 	= clsPay.FTotDutyTime
	totNightTime	= clsPay.FtotNightTime
	totPaySum		= clsPay.FTotPaySum 
	totWorkDay		= ceilValue(clsPay.FWeekWorkDay * avgWeek)		'// 기본 근무일수
 
	'// ========================================================================
	'// 저장된 계약정보
	'// ========================================================================
	clsPay.fnGetmonthlypayData
	iworktime      	= clsPay.Fworktime
	iextendtime    	= clsPay.Fextendtime
	inighttime     	= clsPay.Fnight
	iholidaytime   	= clsPay.Fholidaytime
	mtimepay       	= clsPay.Ftimepay
	mextendpay     	= clsPay.Fextendpay
	mnightpay      	= clsPay.Fnightpay
	mholidaypay		= clsPay.Fholidaypay
	mwholidaypay 	= clsPay.Fwholidaypay
	mfoodpay       	= clsPay.Ffoodpay
	mjobpay        	= clsPay.Fjobpay
	moutstandingpay = clsPay.Foutstandingpay
	mlongtimepay		= clsPay.Flongtimepay
	maddpay					= clsPay.Faddpay
	mtotpay        	= clsPay.Ftotpay
	mnpensionpay 		= clsPay.Fnpensionpay
	mhealthinspay 	= clsPay.Fhealthinspay
	mrecuinspay   	= clsPay.Frecuinspay
	munempinspay		= clsPay.Funempinspay
	mtaxtotpay     	= clsPay.Ftaxtotpay
	mrealtotpay    	= clsPay.Frealtotpay
	dregdate       	= clsPay.Fregdate
	sadminid       	= clsPay.Fadminid
	istate         	= clsPay.Fstate
	myearpay				= clsPay.Fyearpay
	mbonuspay		= clsPay.Fbonuspay
	mworkday		= clsPay.Fworkday 
	
	iReworktime    	= clsPay.FReworktime   
	iReextendtime  	= clsPay.FReextendtime 
	iRenighttime   	= clsPay.FRenighttime      
	iReholidaytime 	= clsPay.FReholidaytime
	iRefoodtime 		= clsPay.FReFoodtime
	mretimepay     	= clsPay.FRetimepay 
	mreextendpay    = clsPay.FReExtimepay 
	mrenightpay    = clsPay.FReNTtimepay 
	mreholidaypay    = clsPay.FReHDtimepay 
	mrefoodpay      = clsPay.FReFtimepay 
	mretotpay				= clsPay.FReTotpay  
	ireworkday 			= clsPay.FReWorkday
	
 if isNull(mretimepay) or mretimepay ="" then mretimepay = 0
 if isNull(mreextendpay) or mreextendpay ="" then mreextendpay = 0
 if isNull(mrenightpay) or mrenightpay ="" then mrenightpay = 0
 if isNull(mreholidaypay) or mreholidaypay ="" then mreholidaypay = 0
 if isNull(mrefoodpay) or mrefoodpay ="" then mrefoodpay = 0 
 if isNull(ireworkday) or ireworkday="" then ireworkday = 0	
  if isNull(mretotpay) or mretotpay ="" then mretotpay =0 					
if Not isNull(iworktime) and iworktime <> "" then
	totWorkDay = mworkday
	totReWorkday = ireworkday
end if
 
 
if Not isNull(iworktime) and iworktime <> "" and iposit_sn<>12 and iposit_sn<>14 and iposit_sn<>15 then 
	'// 시급직일 경우 저장된 근무일수 가져오기(일일 데이타 기준)
	clsPay.FSyyyymm = dSPayDate
	clsPay.FEyyyymm = dEPayDate 
	clsPay.FPreyyyymmdd = dSPayDate
	arrList = clsPay.fnGetDailypayData
  arrPre  = clsPay.fnGetPreReDailypayData
	totWorkDayReal = 0
	totReWorkDayReal = 0
	if isArray(arrList) then
		For intLoop = 0 To UBOund(arrList,2) 
	 
			IF arrList(0,intLoop) < chkDate&"-01"  THEN  
				IF isArray(arrPre) THEN
					iP = 0  
					For intP = iP To UBound(arrPre,2) 
						 if arrList(0,intLoop) = arrPre(0,intP) THEN  
							if arrList(7,intLoop) < 60 and arrPre(7,intP) >=60 THEN
									totReWorkDayReal = totReWorkDayReal - 1 
							ELSEif arrList(7,intLoop) >= 60 and arrPre(7,intP) <60 THEN
							 		totReWorkDayReal = totReWorkDayReal + 1  
							end if	
						iP= iP+1
						end if
					Next
				END IF
			elseif arrList(7,intLoop) >= 240  then  
				'// 4시간 이상 근무시 근무일수 추가
				totWorkDayReal = totWorkDayReal + 1
			end if
		Next
	end if

end if

IF  iworktime ="" or isNull(iworktime) or blnReset = "1" THEN
	'// ========================================================================
	'// 월계약정보가 없으면(또는 월 급여 재계산시) 기본 계약에서 데이타 가져온다.
	'// ========================================================================

	monthlyPayDataExist = False
 
	IF iposit_sn=12 or iposit_sn=14 or iposit_sn=15 THEN	'월급제(월급/프리/인턴)
		iworktime    	= (ceilValue(totDutyTime/60*avgWeek)+ceilValue(holidaywdtime/60*avgWeek))*60
		iextendtime 	= iOverTime
		inighttime    	= ceilValue(totNightTime/60*avgWeek)*60
		iholidaytime   	=  0
		mtimepay       	= defaultpay*ceilValue(totDutyTime/60*avgWeek)+ defaultpay*ceilValue(holidaywdtime/60*avgWeek)
		if (foodpay=0) then
		    mfoodpay		= 0
		else
		    mfoodpay		= ceilValue(totWorkDay * foodpay)   '' totWorkDay 가 널이라 일단 foodpay 0인지체크 '' 상구엉아 작업내용인듯
	    end if
		mextendpay     	= defaultpay*iOverTime*1.5
		mnightpay      	= defaultpay*ceilValue(totNightTime/60*avgWeek)*0.5
		mholidaypay		= 0
		mtotpay        	= totPaySum 

		IF blnstatediv ="N"  and  left(dretireday,7) =  dYear&"-"&format00(2,dMonth) and dretireday < dEndDate  and dretireday <= enddate  THEN
			'퇴사한 경우 퇴사일이 검색달 마지막 날짜보다 빠르면  퇴사일까지 총 금액에서 날짜로 나눈다.

			IF left(startdate,7) =  dYear&"-"&format00(2,dMonth) and startdate  >  dateserial(dYear,dMonth,1)  THEN
				dREday =  day(dretireday)-day(startdate) + 1
			ELSE
				dREday =  day(dretireday)
			END IF

			iworktime		= (iworktime/dEndDay)*dREday
			iextendtime		= round((iextendtime/dEndDay)*dREday,0)
			inighttime		= round((inighttime/dEndDay)*dREday,0)
			mtimepay       	= round((mtimepay/dEndDay)*dREday,0)
			mextendpay     	= round((mextendpay/dEndDay)*dREday,0)
			mnightpay      	= round((mnightpay/dEndDay)*dREday,0)
			mfoodpay      	= round((mfoodpay/dEndDay)*dREday,0)
			mtotpay        	= mtimepay +  mextendpay + mnightpay
		ELSEIF 	left(enddate,7) =  dYear&"-"&format00(2,dMonth) and enddate <  dEndDate THEN
		   '마지막 날보다 계약 종료일이 빠른 경우 계약종료일까지 총 금액에서 날짜로 나눈다.
		   IF left(startdate,7) =  dYear&"-"&format00(2,dMonth) and startdate  >  dateserial(dYear,dMonth,1)  THEN
				dREday =  day(enddate) -day(startdate) + 1
			ELSE
				dREday =  day(enddate)
			END IF

			iworktime		= round((iworktime/dEndDay)*dREday,0)
			iextendtime		= round((iextendtime/dEndDay)*dREday,0)
			inighttime		= round((inighttime/dEndDay)*dREday,0)
			mtimepay       	= round((mtimepay/dEndDay)*dREday,0)
			mextendpay     	= round((mextendpay/dEndDay)*dREday,0)
			mnightpay      	= round((mnightpay/dEndDay)*dREday,0)
			mfoodpay      	= round((mfoodpay/dEndDay)*dREday,0)
			mtotpay        	= mtimepay +  mextendpay + mnightpay
		ELSEIF left(startdate,7) =  dYear&"-"&format00(2,dMonth) and startdate  >  dateserial(dYear,dMonth,1)  THEN

			' 월 중간 입사자일 경우 ..(총임급/해당 월 일수)*(해당 월 마지막 날 - 입사일 + 1)
			dREday =  dEndDay-day(startdate) + 1

			iworktime		= round((iworktime/dEndDay)*dREday,0)
			iextendtime		= round((iextendtime/dEndDay)*dREday,0)
			inighttime		= round((inighttime/dEndDay)*dREday,0)
			mtimepay       	= round((mtimepay/dEndDay)*dREday,0)
			mextendpay     	= round((mextendpay/dEndDay)*dREday,0)
			mnightpay      	= round((mnightpay/dEndDay)*dREday,0)
			mfoodpay      	= round((mfoodpay/dEndDay)*dREday,0)
			mtotpay        	= mtimepay +  mextendpay + mnightpay
		END IF
		mrealtotpay = mtotpay
	 ELSE	'시급제
	 	 
			clsPay.FSyyyymm = dSPayDate
			clsPay.FEyyyymm = dEPayDate 
			arrPre 	= clsPay.fnGetPreReDailypayData
			arrList = clsPay.fnGetDailypayData
			iworktime = 0
			iextendtime  = 0
			inighttime	=0
			iholidaytime=0
			iweekholidaytime=0

			 
			totWorkDay = 0
			totreWorkDay = 0
			IF isArray(arrList) THEN
				 
				For intLoop = 0 To UBOund(arrList,2)
				 IF arrList(0,intLoop) <  dateserial(dYear,dMonth,1) THEN
				 		IF isArray(arrPre) THEN
				 			iP = 0
							For intP = iP To UBound(arrPre,2)
								 if arrList(0,intLoop) = arrPre(0,intP) THEN
									 	iTReworktime =arrList(7,intLoop) - arrPre(7,intP) 
									 	iTReextendtime =arrList(8,intLoop) - arrPre(8,intP) 
									 	iTRenoghttime =arrList(9,intLoop) - arrPre(9,intP) 
									 	iTReholidayime =arrList(10,intLoop) - arrPre(10,intP) 
									 	iTReweekholidayime =arrList(11,intLoop) - arrPre(11,intP) 
									 	
										 	if arrList(7,intLoop) < 60 and arrPre(7,intP) >=60 THEN
										 		totReWorkDay = totReWorkDay - 1
										 	ELSEif arrList(7,intLoop) >= 60 and arrPre(7,intP) <60 THEN
										 		totReWorkDay = totReWorkDay + 1 
											end if	
									 	iP= iP+1
									end if
							Next
						END IF
				 
				 		iReworktime		= iReworktime + iTReworktime
						iReextendtime  	= iReextendtime + iTReextendtime
						iRenighttime		= iRenighttime + iTRenoghttime
						iReholidaytime	= iReholidaytime + iTReholidayime
						iReweekholidaytime= iReweekholidaytime  + iTReweekholidayime 
				 
				 else
						iworktime 		= iworktime +  arrList(7,intLoop)
						iextendtime  	= iextendtime + arrList(8,intLoop)
						inighttime		= inighttime +  arrList(9,intLoop)
						iholidaytime	= iholidaytime + arrList(10,intLoop)
						iweekholidaytime= iweekholidaytime  + arrList(11,intLoop)

						if (arrList(7,intLoop) >= 240)   then
							'// 한시간 이상 근무시 근무일수 추가
							totWorkDay = totWorkDay + 1
						end if
					end if
				Next

				iworktime 	= iworktime+iweekholidaytime
				''mtimepay    = defaultpay*(iworktime/60)+ defaultpay*(iweekholidaytime/60)
				mtimepay    = round(defaultpay*(iworktime/60),0)
				mextendpay  = round(defaultpay*(iextendtime/60)*1.5,0)
				mnightpay   = round(defaultpay*(inighttime/60)*0.5,0)
				mholidaypay	= round(defaultpay*(iholidaytime/60)*0.5 ,0)

				iReworktime 	= iReworktime+iReweekholidaytime 
				mretimepay    = round(predefaultpay*(iReworktime/60),0)
				mreextendpay  = round(predefaultpay*(iReextendtime/60)*1.5,0)
				mrenightpay   = round(predefaultpay*(iRenighttime/60)*0.5,0)
				mreholidaypay	= round(predefaultpay*(iReholidaytime/60)*0.5 ,0)
			END IF

			mfoodpay		= ceilValue(totWorkDay * foodpay)
			mrefoodpay  = ceilValue(totReWorkDay * prefoodpay)
		END IF
 
		mtotpay     = mtimepay+mextendpay+mnightpay+mholidaypay+mfoodpay+mjobpay+moutstandingpay+mlongtimepay+maddpay+myearpay+mbonuspay
		mretotpay   = mretimepay+mreextendpay+mreextendpay+mreholidaypay+mrefoodpay 
		mrealtotpay = mtotpay + mretotpay
END IF
set clsPay = nothing

%>
  <script language="javascript">
  <!--
  	function jsSearch(){
  		var dNowYear, dNowMonth;
		var date = new Date();
		dNowYear = date.getFullYear();
		dNowMonth = date.getMonth() + 1;

	 	if (document.frmSearch.selY.value > dNowYear){
	 		alert("현재 달 이전까지만  검색 가능합니다.");
	 		return;
	 	}else if (document.frmSearch.selY.value == dNowYear && document.frmSearch.selM.value > dNowMonth){
	 		alert("현재 달 이전까지만  검색 가능합니다.");
	 		return;
	 	}

	 	//계약기간 내에서 검색 가능하도록 변경
	 	document.frmSearch.submit();
  	}

  	//계약정보 등록
	function jsViewPay(empno,ino){
		var wpay = window.open("pop_payform.asp?sEN="+empno+"&ino="+ino,"popPay","width=700,height=600,scrollbars=yes,resizable=yes");
		wpay.focus();
	}

  	//근무시간 등록
 	function jsWorkTime(empno,ino){
 		var wwt =window.open("pop_worktime.asp?sEN="+empno+"&ino="+ino+"&selY=<%=dYear%>&selM=<%=dMonth%>","popWT","width=1200,height=800,scrollbars=yes,resizable=yes");
		wwt.focus();
 	}

 	//총 합계금액 변경
 	function jsSetTotPay(iVal){
 		 <%	IF iposit_sn = 13 THEN %>
 		if (iVal =="iFP"){
 			 document.frmPay.iFPS.value =  parseInt(document.frmPay.iRFP.value.replace(/,/g,""),10) +  parseInt(document.frmPay.iFP.value.replace(/,/g,""),10);
 		}else{
 			eval("document.frmPay."+iVal+"S").value = eval("document.frmPay."+iVal).value ;
 		}
 		
 		document.frmPay.itotP.value  = parseInt(document.frmPay.iTP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iETP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iNTP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iHDP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iFP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iJP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iOP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iLP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iAP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iYP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iBP.value.replace(/,/g,""),10);
		
		document.frmPay.iRtotP.value  = parseInt(document.frmPay.iRTP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iRETP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iRNTP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iRHDP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iRFP.value.replace(/,/g,""),10);
		
		document.frmPay.itotPS.value  = parseInt(	document.frmPay.itotP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iRtotP.value.replace(/,/g,""),10) 
		<%else%>
			document.frmPay.itotP.value  = parseInt(document.frmPay.iTP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iETP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iNTP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iHDP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iFP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iJP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iOP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iLP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iAP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iYP.value.replace(/,/g,""),10)
		+ parseInt(document.frmPay.iBP.value.replace(/,/g,""),10);
		  
		document.frmPay.itotPS.value  = parseInt(	document.frmPay.itotP.value.replace(/,/g,""),10)
		<%end if%>
 	}
 	 

 	function jsSetRealWorkDayToSaved(iVal) {
		var frm = document.frmPay;

	if (iVal=="N"){ 
		frm.totWorkDay.value = frm.totWorkDayReal.value;
		frm.iFP.value = frm.foodpay.value*1 * frm.totWorkDay.value;
	}else{
		frm.totReWorkDay.value = frm.totReWorkDayReal.value;
		frm.iRFP.value = frm.prefoodpay.value*1 * frm.totReWorkDay.value;
	}
		jsSetTotPay('iFP');
 	}

 	//급여등록
 	function jsSubmit(){
		var strMsg,istate;
		for(i=0;i<document.frmPay.hidS.length;i++){
			if(document.frmPay.hidS[i].checked){
				istate = document.frmPay.hidS[i].value;
			}
		}
		jsSetTotPay('iFP');

 		if(istate == 1){
 			strMsg = "작성완료상태로 등록하시겠습니까?" ;
 		}else if(istate == 5){
 			strMsg = "확인완료상태로 등록하시겠습니까?" ;
 		}else if(istate == 7){
 			strMsg = "입금완료상태로 등록하시겠습니까?" ;
 		}else if(istate == 0){
 			strMsg = "급여작성중상태로 등록하시겠습니까?" ;
 		}
 		if(confirm(strMsg)){
 			return true;
 		} else {
			return false;
		}

 	}

 	//프린트
 	function jsPrint(){
 	 var winPrint = window.open("print_worktime.asp?sEN=<%=sempno%>&ino=<%=ino%>&selY=<%=dYear%>&selM=<%=dMonth%>","prtWT","width=1020,height=600,scrollbars=yes,resizable=yes");
 	 winPrint.focus();
 	}

 	//월 급여 재계산
 	function jsRestWorkTime(){
 		document.frmSearch.blnR.value = 1;
 		document.frmSearch.submit();
 	}
  //-->
  </script>
<table width="100%"  cellpadding="3" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">사번</td>
			<td bgcolor="#FFFFFF" width="180"><a href="javascript:jsViewPay('<%=sempno%>','<%=ino%>')"><%=sempno%></a> <%IF blnstatediv ="N" THEN%><font color="red">[퇴사]</font><%END IF%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">입사일</td>
			<td bgcolor="#FFFFFF"><%=formatdate(djoinday,"0000-00-00")%></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">이름</td>
			<td bgcolor="#FFFFFF"><%=susername%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">퇴사일</td>
			<td bgcolor="#FFFFFF">
				<%IF blnstatediv = "N" THEN%>
					<% if Not IsNull(dretireday) then %>
						<%=formatdate(dretireday,"0000-00-00")%>
					<% else %>
						<font color="red">에러 : 시스템팀 문의</font>
					<% end if %>
				<%END IF%>
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">계약구분</td>
			<td bgcolor="#FFFFFF"><%=sposit_name%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">시간급</td>
			<td bgcolor="#FFFFFF"><%if predefaultpay>0 then%>(전월: <%=formatnumber(predefaultpay,0)%> 원) <%end if%><%=formatnumber(defaultpay,0)%> 원</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">근무일수</td>
			<td bgcolor="#FFFFFF">
				<% if (iposit_sn=12 or iposit_sn=14 or iposit_sn=15) then %>
					<% if (monthlyPayDataExist = True) then %>
						<%= mworkday %>
					<% else %>
						<%= totWorkDay %>
					<% end if %>
				<% else %>
					--
				<% end if %>
			</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">식대</td>
			<td bgcolor="#FFFFFF"><%if prefoodpay>0 then%>(전월: <%=formatnumber(prefoodpay,0)%> 원) <%end if%><%=formatnumber(foodpay,0)%> 원</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">계약일</td>
			<td bgcolor="#FFFFFF">[<%=ino%>] <%=formatdate(startdate,"0000-00-00")%> ~ <%=formatdate(enddate,"0000-00-00")%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">휴계시간</td>
			<td bgcolor="#FFFFFF"><%IF inbreaktime THEN%>근무시간 포함<%ELSE%>근무시간 포함안함<%END IF%></td>
		</tr>

		</table>
	</td>
</tr>
<form name="frmSearch" method="get" action="">
<input type="hidden" name="sEN" value="<%=sEmpno%>">
<input type="hidden" name="ino" value="<%=ino%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="blnR" value="0">
<tr>
	<td>근무날짜:
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
		&nbsp;&nbsp;
<%
	If Request.ServerVariables("REMOTE_ADDR") = "61.252.133.15" Then
%>
<input type="button" class="button" value="일별 근무시간" onClick="jsWorkTime('<%=sempno%>','<%=ino%>');">
<%
	End If
%>

		<%IF iposit_sn = 13 THEN%><input type="button" class="button" value="일별 근무시간" onClick="jsWorkTime('<%=sempno%>','<%=ino%>');"><%END IF%>
		 <input type="button" value="프린트" class="button" onClick="jsPrint();">
		 <%IF (iposit_sn=12 or iposit_sn=14 or iposit_sn=15) and istate  = 0 THEN%><input type="button" class="button" value="월 급여 재계산" onClick="jsRestWorkTime();"> <br><div style="padding-top:5px"><font color="Red">* [월 급여 재계산]시 데이터 확인 후 [등록]버튼을 꼭 눌러주세요. 등록 버튼 미처리시 재계산 이전 데이터로 처리됩니다. </font></div><%END IF%>
	</td>
</tr>
</form>
<tr>
	<td><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<form name="frmPay" method="post" action="tenbyten_pay_process.asp" onSubmit="return jsSubmit();">
		<input type="hidden" name="hidSPdate" value="<%= dSPayDate %>"><!-- 급여시작일자 년월일-->
		<input type="hidden" name="hidEPdate" value="<%= dEPayDate %>"><!-- 급여종료일자 년월일-->
		<tr>
		<td  bgcolor="<%= adminColor("gray") %>" width="120" align="center">급여등록상태</td>
		<td bgcolor="#FFFFFF">
				<input type="radio" name="hidS" value="0" <%IF istate  = 0  THEN%>checked<%ELSEIF istate >1 and not(C_ADMIN_AUTH or C_PSMngPart)  THEN%>disabled<%END IF%>><%IF istate  = 0  THEN%><font color="red"><%END IF%>급여작성중 ></font>
				<input type="radio" name="hidS" value="1" <%IF istate  = 1  THEN%>checked<%ELSEIF istate >1 THEN%>disabled<%END IF%>><%IF istate  = 1  THEN%><font color="red"><%END IF%>작성완료 ></font>
				<input type="radio" name="hidS" value="5" <%IF istate  = 5  THEN%>checked<%END IF%>><%IF istate  = 5  THEN%><font color="red"><%END IF%>경영지원확인완료 ></font>
				<input type="radio" name="hidS" value="7" <%IF istate  = 7  THEN%>checked<%END IF%>><%IF istate  = 7  THEN%><font color="red"><%END IF%>입금완료 </font>
		</td>

	</tr>
	</table>
	</td>
</tr>
<tr>
	<td>
		<table border="0" width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<input type="hidden" name="hidM" value="M">
		<input type="hidden" name="hidEN" value="<%=sempno%>">
		<input type="hidden" name="ino" value="<%=ino%>">
		<input type="hidden" name="hidPSN" value="<%=iposit_sn%>">
		<input type="hidden" name="hidYear" value="<%=dYear%>">
		<input type="hidden" name="hidMonth" value="<%=dMonth%>">
		<input type="hidden" name="iWT" value="<%=iworktime%>">
		<input type="hidden" name="iEWT" value="<%=iextendtime%>">
		<input type="hidden" name="iNWT" value="<%=inighttime%>">
		<input type="hidden" name="iHDT" value="<%=iholidaytime%>">
		<input type="hidden" name="iRWT" value="<%=ireworktime%>">
		<input type="hidden" name="iREWT" value="<%=ireextendtime%>">
		<input type="hidden" name="iRNWT" value="<%=irenighttime%>">
		<input type="hidden" name="iRHDT" value="<%=ireholidaytime%>">
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<input type="hidden" name="foodpay" value="<%=foodpay%>">
		<input type="hidden" name="prefoodpay" value="<%=prefoodpay%>">
		<input type="hidden" name="totWorkDay" value="<%=totWorkDay%>"><!-- 시급=실근무일수 / 월급직=기본근무일수(중간입사퇴사인 경우도 동일) -->
		<input type="hidden" name="totWorkDayReal" value="<%=totWorkDayReal%>">
		<input type="hidden" name="totReWorkDay" value="<%=totReWorkDay%>"><!-- 시급=실근무일수 / 월급직=기본근무일수(중간입사퇴사인 경우도 동일) -->
		<input type="hidden" name="totReWorkDayReal" value="<%=totReWorkDayReal%>">
		<tr  bgcolor="<%= adminColor("gray") %>" align="center">
			<td>구분</td>
			<td>기본급</td>
			<td>시간외수당</td>
			<td>야간근무수당</td>
			<td>휴일근무수당</td>
			<td>식대지원</td>
			<td>직책수당</td>
			<td>우수사원</td>
			<td>장기근속수당</td>
			<td>추가수당</td>
			<td>연차수당</td>
			<td>상여금</td> 
			<td>총액</td>
		</tr>
		<%IF iposit_sn = 13 THEN %>
						<%if sempno= "90201501120013" or sempno="90201610010124" or sempno="90201611130141" or sempno="90201611140136" or sempno="90201611200158" or sempno="90201611260172" or sempno="90201612060169" or sempno="90201612100180" or sempno="90201612120174" or sempno="90201612210190" then%>
							<tr  bgcolor="#FFFFFF" align="center">
							<td bgcolor="<%= adminColor("gray") %>">당월금액</td>
							<td><input type="text" name="iTP" value="<%=formatnumber(mtimepay,0)%>" class="text" style="text-align:right;border:0;" readonly size="9"></td>
							<td><input type="text" name="iETP" value="<%=formatnumber(mextendpay,0)%>" class="text"  style="text-align:right;border:0;" readonly  size="8"> </td>
							<td><input type="text" name="iNTP" value="<%=formatnumber(mnightpay,0)%>" class="text" style="text-align:right;border:0;" readonly  size="8"></td>
							<td><input type="text" name="iHDP" value="<%=formatnumber(mholidaypay,0)%>" class="text" style="text-align:right;border:0;" readonly  size="8"></td>
							<td><input type="text" name="iFP" value="<%=formatnumber(mfoodpay,0)%>" class="text" style="text-align:right;border:0;" readonly  size="8"></td>
							<td><input type="text" name="iJP" value="<%=formatnumber(mjobpay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="8" onKeyUp="jsSetTotPay('iJP');"></td>
							<td><input type="text" name="iOP" value="<%=formatnumber(moutstandingpay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="8" onKeyUp="jsSetTotPay('iOP');"></td>
							<td><input type="text" name="iLP" value="<%=formatnumber(mlongtimepay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="8" onKeyUp="jsSetTotPay('iLP');"></td>
							<td><input type="text" name="iAP" value="<%=formatnumber(maddpay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="10" onKeyUp="jsSetTotPay('iAP');"></td>
							<td><input type="text" name="iYP" value="<%=formatnumber(myearpay,0)%>"  class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="10" onKeyUp="jsSetTotPay('iYP');"></td>
							<td><input type="text" name="iBP" value="<%=formatnumber(mbonuspay,0)%>"  class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="10" onKeyUp="jsSetTotPay('iBP');"></td>			
							<td><input type="text" name="itotP" value="<%=formatnumber(mtotpay,0)%>"  class="text" style="text-align:right;border:0;" readonly size="10"></td>
						</tr>
						<tr  bgcolor="#FFFFFF" align="center" height="30">
							<td bgcolor="<%= adminColor("gray") %>">당월시간</td>
							<td><%=fnSetTimeFormat(iWorkTime)%></td>
							<td><%=fnSetTimeFormat(iextendtime)%></td>
							<td><%=fnSetTimeFormat(inighttime)%></td>
							<td><%=fnSetTimeFormat(iholidaytime)%></td>
							<td colspan="8" align="left">
								* 근무일수 : 
								<% if (monthlyPayDataExist = True) then %>
									<%= mworkday %>일
									<% if (mworkday <> totWorkDayReal) or (foodpay <> 0 and totWorkDay <> 0 and mfoodpay = 0) then %>
										<font color="red">(실근무일수 : <%= totWorkDayReal %>일)</font>
										<input type="button" class="button" value="실근무일수 적용" onClick="jsSetRealWorkDayToSaved('N')">
									<% end if %>
								<% else %>
									<%= totWorkDay %>
								<% end if %>
							</td>
						</tr> 
						<%ELSE%>
						<tr  bgcolor="<%=adminColor("sky")%>" align="center">
							<td bgcolor="<%= adminColor("sky") %>"><b>총금액</b></td>
							<td><input type="text" name="iTPS" value="<%=formatnumber(mtimepay+mretimepay,0)%>" class="text" style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly size="9"></td>
							<td><input type="text" name="iETPS" value="<%=formatnumber(mextendpay+mreextendpay,0)%>" class="text" style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly  size="8"> </td>
							<td><input type="text" name="iNTPS" value="<%=formatnumber(mnightpay+mrenightpay,0)%>" class="text"  style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly  size="8"></td>
							<td><input type="text" name="iSHDPS" value="<%=formatnumber(mholidaypay+mreholidaypay,0)%>" class="text" style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly  size="8"></td>
							<td><input type="text" name="iFPS" value="<%=formatnumber(mfoodpay+mrefoodpay,0)%>" class="text" style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly  size="8"></td>
							<td><input type="text" name="iJPS" value="<%=formatnumber(mjobpay,0)%>" class="text"  style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly size="8"  ></td>
							<td><input type="text" name="iOPS" value="<%=formatnumber(moutstandingpay,0)%>"  class="text" style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly    size="8" ></td>
							<td><input type="text" name="iLPS" value="<%=formatnumber(mlongtimepay,0)%>" class="text"  style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly    size="8"  ></td>
							<td><input type="text" name="iAPS" value="<%=formatnumber(maddpay,0)%>" class="text"  style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly    size="10"  ></td>
							<td><input type="text" name="iYPS" value="<%=formatnumber(myearpay,0)%>" class="text"  style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly    size="10" ></td>
							<td><input type="text" name="iBPS" value="<%=formatnumber(mbonuspay,0)%>" class="text" style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly    size="10" ></td>			
							<td><input type="text" name="itotPS" value="<%=formatnumber(mrealtotpay,0)%>"  class="text" style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly size="10"></td>
						</tr>
						
						<tr  bgcolor="<%=adminColor("sky")%>" align="center" height="30">
							<td bgcolor="<%=adminColor("sky")%>">총시간</td>
							<td><%=fnSetTimeFormat(iWorkTime+iReWorktime)%></td>
							<td><%=fnSetTimeFormat(iextendtime+ireextendtime)%></td>
							<td><%=fnSetTimeFormat(inighttime+irenighttime)%></td>
							<td><%=fnSetTimeFormat(iholidaytime+ireholidaytime)%></td>
							<td colspan="8" align="left">
								* 근무일수 : 
								<% if (monthlyPayDataExist = True) then %>
									<%= mworkday+ireworkday %>일 
								<% else %>
									<%= totWorkDay+totReworkday %>
								<% end if %>
							</td>
						</tr>
						<tr  bgcolor="#FFFFFF" align="center">
							<td bgcolor="<%= adminColor("gray") %>">당월금액</td>
							<td><input type="text" name="iTP" value="<%=formatnumber(mtimepay,0)%>" class="text" style="text-align:right;border:0;" readonly size="9"></td>
							<td><input type="text" name="iETP" value="<%=formatnumber(mextendpay,0)%>" class="text"  style="text-align:right;border:0;" readonly  size="8"> </td>
							<td><input type="text" name="iNTP" value="<%=formatnumber(mnightpay,0)%>" class="text" style="text-align:right;border:0;" readonly  size="8"></td>
							<td><input type="text" name="iHDP" value="<%=formatnumber(mholidaypay,0)%>" class="text" style="text-align:right;border:0;" readonly  size="8"></td>
							<td><input type="text" name="iFP" value="<%=formatnumber(mfoodpay,0)%>" class="text" style="text-align:right;border:0;" readonly  size="8"></td>
							<td><input type="text" name="iJP" value="<%=formatnumber(mjobpay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="8" onKeyUp="jsSetTotPay('iJP');"></td>
							<td><input type="text" name="iOP" value="<%=formatnumber(moutstandingpay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="8" onKeyUp="jsSetTotPay('iOP');"></td>
							<td><input type="text" name="iLP" value="<%=formatnumber(mlongtimepay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="8" onKeyUp="jsSetTotPay('iLP');"></td>
							<td><input type="text" name="iAP" value="<%=formatnumber(maddpay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="10" onKeyUp="jsSetTotPay('iAP');"></td>
							<td><input type="text" name="iYP" value="<%=formatnumber(myearpay,0)%>"  class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="10" onKeyUp="jsSetTotPay('iYP');"></td>
							<td><input type="text" name="iBP" value="<%=formatnumber(mbonuspay,0)%>"  class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="10" onKeyUp="jsSetTotPay('iBP');"></td>			
							<td><input type="text" name="itotP" value="<%=formatnumber(mtotpay,0)%>"  class="text" style="text-align:right;border:0;" readonly size="10"></td>
						</tr>
						<tr  bgcolor="#FFFFFF" align="center" height="30">
							<td bgcolor="<%= adminColor("gray") %>">당월시간</td>
							<td><%=fnSetTimeFormat(iWorkTime)%></td>
							<td><%=fnSetTimeFormat(iextendtime)%></td>
							<td><%=fnSetTimeFormat(inighttime)%></td>
							<td><%=fnSetTimeFormat(iholidaytime)%></td>
							<td colspan="8" align="left">
								* 근무일수 : 
								<% if (monthlyPayDataExist = True) then %>
									<%= mworkday %>일 
									<% if (mworkday <> totWorkDayReal) or (foodpay <> 0 and totWorkDay <> 0 and mfoodpay = 0) then %>
										<font color="red">(실근무일수 : <%= totWorkDayReal %>일)</font>
										<input type="button" class="button" value="실근무일수 적용" onClick="jsSetRealWorkDayToSaved('N')">
									<% end if %>
								<% else %>
									<%= totWorkDay %>
								<% end if %>
							</td>
						</tr>
						<tr  bgcolor="#e3f1fb" align="center">
							<td bgcolor="#e3f1fb" nowrap>전월차액금</td>
							<td><input type="text" name="iRTP" value="<%=formatnumber(mretimepay,0)%>" class="text" style="text-align:right;border:0;background:#e3f1fb;" readonly size="9"></td>
							<td><input type="text" name="iRETP" value="<%=formatnumber(mreextendpay,0)%>" class="text" style="text-align:right;border:0;background:#e3f1fb;" readonly  size="8"> </td>
							<td><input type="text" name="iRNTP" value="<%=formatnumber(mrenightpay,0)%>" class="text" style="text-align:right;border:0;background:#e3f1fb;" readonly  size="8"></td>
							<td><input type="text" name="iRHDP" value="<%=formatnumber(mreholidaypay,0)%>" class="text" style="text-align:right;border:0;background:#e3f1fb;" readonly  size="8"></td>
							<td><input type="text" name="iRFP" value="<%=formatnumber(mrefoodpay,0)%>" class="text" style="text-align:right;border:0;background:#e3f1fb;" readonly  size="8"></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>			
							<td><input type="text" name="iRtotP" value="<%=formatnumber(mretotpay,0)%>" class="text"  style="text-align:right;border:0;background:#e3f1fb;" readonly size="10"></td>
						</tr>
						<tr  bgcolor="#e3f1fb" align="center" height="30">
							<td bgcolor="#e3f1fb" nowrap>전월차액시간</td>
							<td><%=fnSetTimeFormat(ireWorkTime)%></td>
							<td><%=fnSetTimeFormat(ireextendtime)%></td>
							<td><%=fnSetTimeFormat(irenighttime)%></td>
							<td><%=fnSetTimeFormat(ireholidaytime)%></td>
							<td colspan="8" align="left">
								* 근무일수 :
								<% if (monthlyPayDataExist = True) then %>
									<%= ireworkday %>일
									<% if (ireworkday <> totReWorkDayReal) or (prefoodpay <> 0 and totReWorkDay <> 0 and mrefoodpay = 0) then %>
										<font color="red">(실근무일수 : <%= totReWorkDayReal %>일)</font>
										<input type="button" class="button" value="실근무일수 적용" onClick="jsSetRealWorkDayToSaved('P')">
									<% end if %>
								<% else %>
									<%= totreWorkDay %>
								<% end if %>
							</td>
						</tr>
						<%	END IF%>
		<%ELSE%>			
			<tr  bgcolor="#FFFFFF" align="center">
							<td bgcolor="<%= adminColor("gray") %>">총금액</td>
							<td><input type="text" name="iTP" value="<%=formatnumber(mtimepay,0)%>" class="text" style="text-align:right;border:0;" readonly size="9"></td>
							<td><input type="text" name="iETP" value="<%=formatnumber(mextendpay,0)%>" class="text"  style="text-align:right;border:0;" readonly  size="8"> </td>
							<td><input type="text" name="iNTP" value="<%=formatnumber(mnightpay,0)%>" class="text" style="text-align:right;border:0;" readonly  size="8"></td>
							<td><input type="text" name="iHDP" value="<%=formatnumber(mholidaypay,0)%>" class="text" style="text-align:right;border:0;" readonly  size="8"></td>
							<td><input type="text" name="iFP" value="<%=formatnumber(mfoodpay,0)%>" class="text" style="text-align:right;border:0;" readonly  size="8"></td>
							<td><input type="text" name="iJP" value="<%=formatnumber(mjobpay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="8" onKeyUp="jsSetTotPay('iJP');"></td>
							<td><input type="text" name="iOP" value="<%=formatnumber(moutstandingpay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="8" onKeyUp="jsSetTotPay('iOP');"></td>
							<td><input type="text" name="iLP" value="<%=formatnumber(mlongtimepay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="8" onKeyUp="jsSetTotPay('iLP');"></td>
							<td><input type="text" name="iAP" value="<%=formatnumber(maddpay,0)%>" class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="10" onKeyUp="jsSetTotPay('iAP');"></td>
							<td><input type="text" name="iYP" value="<%=formatnumber(myearpay,0)%>"  class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="10" onKeyUp="jsSetTotPay('iYP');"></td>
							<td><input type="text" name="iBP" value="<%=formatnumber(mbonuspay,0)%>"  class="text" style="text-align:right;<%IF istate > 5 THEN%>border:0;" readonly<%ELSE%>"<%END IF%>   size="10" onKeyUp="jsSetTotPay('iBP');"></td>			
							<td><input type="text" name="itotP" value="<%=formatnumber(mtotpay,0)%>"  class="text" style="text-align:right;border:0;" readonly size="10">
								<input type="hidden" name="itotPS" value="<%=formatnumber(mrealtotpay,0)%>"  class="text" style="text-align:right;border:0;font-weight:bold;background:#DDDDFF;" readonly size="10">
								</td>
						</tr>	
		<%END IF%>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
	<%
		If istate < 5 Then
			Response.Write "<input type=""submit"" value=""등록"" class=""button"">"
		Else
			If istate >= 5 Then
				If C_ADMIN_AUTH or C_PSMngPart Then
					Response.Write "<input type=""submit"" value=""등록"" class=""button"">"
				Else
					Response.Write "※ 급여등록상태가 <font color=blue><b>[작성완료]인 경우 부터</b></font>는 <font color=red><b>경영지원부 - 인사교육파트 에서만 변경 가능</b></font>합니다."
					Response.Write "<br>이에 대한 문의는 경영지원부 - 인사교육파트(070-7515-5440)로 문의하시기 바랍니다."
				End If
			End If
		End If
	%>
	</td>
</tr>
</form>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
