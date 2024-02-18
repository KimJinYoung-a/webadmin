<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사원별 급여 기본정보 프린트
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
'변수선언
Dim sempno, ino
Dim djoinday, susername, iposit_sn,sposit_name,blnstatediv,dretireday
Dim startdate,enddate ,defaultpay,foodpay,jobpay,inbreaktime,holidaywdtime
Dim intY, intM, intD, dYear, dMonth,dWeekday
Dim dEndDay,dNextDate
Dim clsPay, arrList,arrRe
Dim dyyyymmdd, dstartHour, dstartMinute, dendHour, dendMinute, dbreakSHour,dbreakSMinute, dbreakEHour, dbreakEMinute,  doutHour, doutMinute, iworktype, dstate, dStart, dEnd, dBreakS, dBreakE
Dim iWorkTime,iextendWT ,inightWT,iholidayWT,iweekholidayWT,dNStart, dNEnd, dNBreakS, dNBreakE,ivacationTime
Dim iextendtime,inighttime,iholidaytime ,mtimepay,mextendpay ,mnightpay ,mholidaypay, mwholidaypay,mfoodpay,mjobpay ,mlongtimepay, maddpay, myearpay, mbonuspay 
Dim moutstandingpay,mtotpay,mnpensionpay,mhealthinspay,mrecuinspay,munempinspay,mtaxtotpay,mrealtotpay,dregdate,sadminid,istate 
Dim totWorkTime, totextendWT ,totnightWT,totholidayWT,totweekholidayWT,totVacationTime
Dim  preStartDay, preEndDay,arrPre
Dim dSWD, dEWD,iWD, totWD,iWT,totWH, chkWHD
dim totPWD
 Dim dcStartHour(8),dcStartMinute(8),dcEndHour(8),dcEndMinute(8),dcBreakSHour(8),dcBreakSMinute(8),dcBreakEHour(8),dcBreakEMinute(8) ,defaulttime(8), intLoop
 dim iReworktime,iReextendtime,iRenighttime,iReholidaytime,iRefoodtime,mretimepay,mreextendpay,mrenightpay,mreholidaypay,mrefoodpay,mretotpay,ireworkday
'값 받아오기
sempno= requestCheckvar(Request("sEN"),14)
ino= requestCheckvar(Request("ino"),10)
dYear = requestCheckvar(Request("selY"),4)
dMonth = requestCheckvar(Request("selM"),2)


dim dSPayDate,dEPayDate,dPreYear,dPreMonth ,dEndDate  
'전달 일주일 시작일, 종료일
preEndDay = dateadd("d", -1, dateserial(dYear,dMonth,1)) '이전달  마지막 일 
dPreYear = year(preEndDay) '이전달 년
dPreMonth = month(preEndDay) '이전달 월
dNextDate = dateadd("m",1, dateserial(dYear,dMonth,1))	'검색다음달 1일
dEndDate = dateadd("d",-1,dNextDate)
dEndDay = day(dEndDate)
'------------------------------------------------------------------ 
IF  dYear&"-"&format00(2,dMonth)  = "2014-01" THEN '2014.01부터 급여종료일 25일로 변경됨 
	dSPayDate = dateserial(dYear,dMonth,1) '급여시작일: 해당월 1일부터
	dEPayDate = dateserial(dYear,dMonth,25) '급여종료일: 해당월 25일까지  
ELSEIF dYear&"-"&format00(2,dMonth) > "2014-01" and dYear&"-"&format00(2,dMonth)  < "2016-12" THEN  
	dSPayDate = dateserial(dPreYear,dPreMonth,26) '급여시작일: 이전월 26일부터 
	dEPayDate = dateserial(dYear,dMonth,25) '급여종료일: 해당월 25일까지  
ELSEIF dYear&"-"&format00(2,dMonth) >= "2016-12" THEN  
	dSPayDate = dateserial(dPreYear,dPreMonth,26) '급여시작일: 이전월 26일부터 
	dEPayDate = dateserial(dYear,dMonth,dEndDay) '급여종료일: 해당월 25일까지  
ELSE   
	dSPayDate = dateserial(dYear,dMonth,1) '급여시작일: 해당월 1일부터
	dEPayDate = dateserial(dYear,dMonth,dEndDay)  '급여종료일: 해당월 말일까지 
END IF  
'------------------------------------------------------------------ 
  
'데이터 가져오기
set clsPay = new CPay
	'--사원 기본계약정보
	clsPay.Fempno = sempno
	clsPay.Fyyyymm = dYear&"-"&format00(2,dMonth)
	clsPay.Fino	= ino
	clsPay.fnGetUserPayData		
	sempno	= clsPay.Fempno
	susername	= clsPay.Fusername	     
	djoinday	  	= clsPay.Fjoinday	     
	blnstatediv 	= clsPay.Fstatediv 	   
	iposit_sn		= clsPay.Fposit_sn	
	sposit_name 	= clsPay.Fposit_name	    
	dretireday		= clsPay.Fretireday 
	 
	holidaywdtime = clsPay.Fholidaywdtime
	ino			= clsPay.Fino
	startdate		= clsPay.Fstartdate
	enddate		= clsPay.Fenddate		
	defaultpay    	= clsPay.Fdefaultpay 	
	foodpay	    	= clsPay.Ffoodpay		
	jobpay		= clsPay.Fjobpay					
	inbreaktime	= clsPay.FinBreakTime
	
	For intLoop = 1 To 7	
		dcStartHour(intLoop) 		= format00(2,Fix(clsPay.FStartTime(intLoop)/60) )
		dcStartMinute(intLoop)  	= format00(2,clsPay.FStartTime(intLoop) mod 60)  
		dcEndHour(intLoop)       	= format00(2,Fix(clsPay.FEndTime(intLoop)/60) )
		dcEndMinute(intLoop)       	= format00(2,clsPay.FEndTime(intLoop) )
		dcBreakSHour(intLoop)     	= format00(2,Fix(clsPay.FBreakSTime(intLoop)/60))	
		dcBreakSMinute(intLoop)     = format00(2,clsPay.FBreakSTime(intLoop) mod 60)
		dcBreakEHour(intLoop)     	= format00(2,Fix(clsPay.FBreakETime(intLoop)/60))	
		dcBreakEMinute(intLoop)     = format00(2,clsPay.FBreakETime(intLoop) mod 60)
		defaulttime(intLoop)		= clsPay.FdefaultTime(intLoop)
	 
	Next
	
	clsPay.FSyyyymm = dSPayDate
	clsPay.FEyyyymm = dEPayDate
	clsPay.FPreyyyymmdd =dSPayDate
	'--지난달 일주일 근무시간(주휴일 계산을 위해)
	arrPre =clsPay.fnGetPreDailypayData
	'--검색달 근무시간 내역
	arrList = clsPay.fnGetDailypayData
	 arrRe  =clsPay.fnGetPreReDailypayData
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
	mlongtimepay	= clsPay.Flongtimepay
	maddpay			= clsPay.Faddpay
	myearpay		= clsPay.Fyearpay
	mbonuspay		= clsPay.Fbonuspay
	mtotpay        	= clsPay.Ftotpay        
	mnpensionpay 	= clsPay.Fnpensionpay   
	mhealthinspay 	= clsPay.Fhealthinspay  
	mrecuinspay   	= clsPay.Frecuinspay    
	munempinspay	= clsPay.Funempinspay   
	mtaxtotpay     	= clsPay.Ftaxtotpay     
	mrealtotpay    	= clsPay.Frealtotpay    
	dregdate       	= clsPay.Fregdate       
	sadminid       	= clsPay.Fadminid       
	istate         	= clsPay.Fstate   
	
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
set clsPay = nothing
 
 
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
	response.end
END IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/scm.css" type="text/css"> 
<script language="javascript">
<!--
	document.body.onload=function(){window.print();} 
//-->
</script>
</head>
<body leftmargin="10" topmargin="10">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a"> 
<tr>
	<td>근무월:  <%=dYear%>년 <%=dMonth%>월 </td>
</tr>  
<tr>
	<td>
		<table width="100%" border="1" cellpadding="5" cellspacing="0" align="center" class="a" bgcolor=#BABABA> 
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">사번</td>
			<td bgcolor="#FFFFFF"><%=sempno%> <%IF blnstatediv ="N" THEN%><font color="red">[퇴사]</font><%END IF%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">입사일</td>
			<td bgcolor="#FFFFFF"><%IF djoinday <> "" THEN%><%=formatdate(djoinday,"0000-00-00")%><%END IF%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="80" align="center" rowspan="3">근로자<br>확인서명</td>
			<td  bgcolor="#FFFFFF" width="80" align="center" rowspan="3">&nbsp;</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" align="center">이름</td>
			<td bgcolor="#FFFFFF"><%=susername%></td>			
			 <td  bgcolor="<%= adminColor("tabletop") %>" align="center">시간급</td>
			<td bgcolor="#FFFFFF"><%=formatnumber(defaultpay,0)%> 원</td>			
		</tr> 
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" align="center">계약기간</td>
			<td bgcolor="#FFFFFF" colspan="3"><%IF startdate <> "" THEN%> <%=formatdate(startdate,"0000-00-00")%> ~ <%=formatdate(enddate,"0000-00-00")%><%END IF%></td> 
		</tr>
		</table>
	</td>
</tr> 

<tr>
	<td>
		<table border="1" width="100%" cellpadding="2" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr  bgcolor="<%= adminColor("gray") %>" align="center">
			<td>구분</td>
			<td>기본급</td>
			<td>연장근무<bR>수당</td>
			<td>야간근무<bR>수당</td>
			<td>휴일근무<bR>수당</td>			
			<td>식대지원</td>
			<td>직책수당</td>	
			<td>우수사원</td>
			<td>장기근속<bR>수당</td>
			<td>추가수당</td>
			<td>연차수당</td>
			<td>상여금</td>
			<td>총액</td>
		</tr>
		<tr  bgcolor="#FFFFFF" align="center">  
			<td bgcolor="<%= adminColor("gray") %>">금액</td>
			<td><%=formatnumber(mtimepay+mretimepay,0)%></td>
			<td><%=formatnumber(mextendpay+mreextendpay,0)%></td>
			<td><%=formatnumber(mnightpay+mrenightpay,0)%></td>
			<td><%=formatnumber(mholidaypay+mreholidaypay,0)%></td>
			<td><%=formatnumber(mfoodpay+mrefoodpay,0)%></td>
			<td><%=formatnumber(mjobpay,0)%></td>
			<td><%=formatnumber(moutstandingpay,0)%></td>
			<td><%=formatnumber(mlongtimepay,0)%></td>
			<td><%=formatnumber(maddpay,0)%></td>
			<td><%=formatnumber(myearpay,0)%></td>
			<td><%=formatnumber(mbonuspay,0)%></td>
			<td><%=formatnumber(mrealtotpay,0)%></td>
		</tr> 
		</table>
	</td>
</tr> 
<tr>
	<td>
		<table width="100%" border="1" cellpadding="2" cellspacing="0" align="center" class="a" bgcolor=#BABABA>  
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
		<% '== 검색 지난 달 일주일 데이터 정보 가져오기(주휴일 계산을 위함)
		dim chkPWHD
		totPWD = 0
		chkPWHD = 0
		IF isArray(arrPre) THEN
			For intD = 0 To UBound(arrPre,2)
			iWorkTime = 0
			iextendWT  = 0
			inightWT	=0
			iholidayWT=0
			iweekholidayWT=0
			
			iWorkTime 	= arrPre(7,intD) 
			totPWD  	= totPWD  + iWorkTime	'전체 근무시간
			iextendWT 	= arrPre(8,intD) 
			inightWT	= arrPre(9,intD) 
			iholidayWT	= arrPre(10,intD)	
			iweekholidayWT=arrPre(11,intD)
					
			IF arrPre(5,intD) = "1"  and iWorkTime = 0 THEN '근무일에 근무를 안했을경우 수당 지급안됨
			 chkPWHD  =  chkPWHD  + 1
			END IF			
			
			'휴일근무
			IF iworktype =  "3" THEN
				 IF iWorkTime > 0 THEN
					iholidayWT = iWorkTime		
					iWorkTime = 0								
				END IF
			END IF		
			
		%>
		<!--<tr   bgcolor="#DFDFDF" align="center">
			<td><%=day(arrPre(0,intD))%></td>
			<td><%=fnGetStringWD(weekday(arrPre(0,intD)))%></td>
			<td> 
				<%IF arrPre(5,intD)  = "1" THEN%>
						근무일
					<%ELSEIF arrPre(5,intD)  = "2" THEN%>	
						<font color="blue">무급휴일<font>
					<%ELSEIF arrPre(5,intD)  = "3" THEN%>		
						<font color="red">주휴일</font>
					<%ELSEIF arrPre(5,intD)  = "4" THEN%>		
						 유급휴일 
					<%END IF%>	
			</td>		 
			<td><%=fnSetTimeFormat(arrPre(1,intD))%></td>		                                                                     
			<td><%=fnSetTimeFormat(arrPre(2,intD))%></td>		                                                                     
			<td><%=fnSetTimeFormat(arrPre(3,intD))%></td>		 
			<td><%=fnSetTimeFormat(arrPre(4,intD))%></td>	
			<td><%=fnSetTimeFormat(arrPre(12,intD))%></td>	
			<td>&nbsp;</td>
			<td><%=fnSetTimeFormat(iWorkTime)%></td>
			<td><%=fnSetTimeFormat(iweekholidayWT)%></td>	
			<td><%=fnSetTimeFormat(iextendWT)%></td>
			<td><%=fnSetTimeFormat(inightWT)%></td>
			<td><%=fnSetTimeFormat(iholidayWT)%></td> 
		</tr>	-->
		<%	Next		
		END IF
		%>
	<% dim totWorkTime_pre,totextendWT_pre,totnightWT_pre,totholidayWT_pre,totweekholidayWT_pre,totVacationTime_pre
dim totWorkTime_re,totextendWT_re,totnightWT_re,totholidayWT_re,totweekholidayWT_re,totVacationTime_re 
dim totWorkTime_sum,totextendWT_sum,totnightWT_sum,totholidayWT_sum,totweekholidayWT_sum,totVacationTime_sum
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
					<%ELSEIF arrRe(5,intD)  = "4" THEN%>
						<font color="red">유급휴일<font>
					<%ELSEIF arrRe(5,intD)  = "5" THEN%>
						<font color="red">공휴일<font>
					<%END IF%>
					<input type="hidden" name="iPWH<%=day(arrRe(0,intD))%>" value="<%=arrRe(5,intD)%>">
				</td>
				<td><%=fnSetTimeFormat(arrRe(1,intD))%></td>
				<td><%=fnSetTimeFormat(arrRe(2,intD))%></td>
				<td><%=fnSetTimeFormat(arrRe(3,intD))%></td>
				<td><%=fnSetTimeFormat(arrRe(4,intD))%></td>
				<td><%=fnSetTimeFormat(arrRe(12,intD))%></td>
				<td></td>
				<td><%=fnSetTimeFormat(iVacationTime)%></td>
				<td><%=fnSetTimeFormat(iWorkTime)%></td>
				<td><%=fnSetTimeFormat(iextendWT)%></td>
				<td><%=fnSetTimeFormat(inightWT)%></td>
				<td><%=fnSetTimeFormat(iholidayWT)%></td>
				<td><%=fnSetTimeFormat(iweekholidayWT)%></td>
			</tr>
			<%	Next
	
		%>
	 
 <tr   bgcolor="<%=adminColor("sky")%>" align="center"> 
				<td colspan="9"><b>A.</b> [<%=dPreMonth%>/26 ~ <%=dPreMonth%>/<%=preEndDay%>] <b>합계</b></td> 
				<td><%=fnSetTimeFormat(totVacationTime_pre)%></td>
				<td><%=fnSetTimeFormat(totWorkTime_pre)%></td> 
				<td><%=fnSetTimeFormat(totextendWT_pre)%></td>
				<td><%=fnSetTimeFormat(totnightWT_pre)%></td>
				<td><%=fnSetTimeFormat(totholidayWT_pre)%></td>
				<td><%=fnSetTimeFormat(totweekholidayWT_pre)%></td>
		</tr>  
		<%	END IF
		 %>
		<%
			totWorkTime = 0
			totextendWT  = 0
			totnightWT	=0
			totholidayWT=0
			totweekholidayWT=0
			totVacationTime = 0
IF isArray(arrList) THEN			
		For intD = 0 To ubound(arrList,2)
			iworktype = ""
			iWorkTime = 0
			iextendWT  = 0
			inightWT	=0
			iholidayWT=0
			iweekholidayWT=0
			ivacationTime = 0
			
			dbreakSHour  =""
			dbreakSMinute  =""
			dbreakEHour  =""
			dbreakEMinute  =""			
				dyyyymmdd	= arrList(0,intD)
				if right(dyyyymmdd,2) <> 32 then
				dWeekday = weekday(dyyyymmdd)
				end if
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
				
					 
			if dbreakSHour = "" THEN dbreakSHour = dcbreakSHour(dWeekday)
			if dbreakSMinute = "" THEN dbreakSMinute = dcbreakSMinute(dWeekday)
			if dbreakEHour = "" THEN dbreakEHour = dcbreakEHour(dWeekday)
			if dbreakEMinute = "" THEN dbreakEMinute = dcbreakEMinute(dWeekday)
		 
		 
		if    dYear&"-"&format00(2,dMonth) >= "2017-01" then
		 if   dyyyymmdd =  Cstr(dateserial(dYear,dMonth,1)) then 
							
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
		%>
		
		<tr   bgcolor="<%=adminColor("sky")%>" align="center"> 
				<td colspan="9"><b>B.</b> [<%=dPreMonth%>/26 ~ <%=dPreMonth%>/<%=preEndDay%>] <b>재계산</b></td>
				<td><%=fnSetTimeFormat(totVacationTime_re)%></td>
				<td><%=fnSetTimeFormat(totWorkTime_re)%></td> 
				<td><%=fnSetTimeFormat(totextendWT_re)%></td>
				<td><%=fnSetTimeFormat(totnightWT_re)%></td>
				<td><%=fnSetTimeFormat(totholidayWT_re)%></td>
				<td><%=fnSetTimeFormat(totweekholidayWT_re)%></td>
		</tr> 
		<tr   bgcolor="<%=adminColor("sky")%>" align="center"> 
				<td colspan="9"> <b> B - A = 차액</b></td>
				<td><%=fnSetTimeFormat(totVacationTime_sum)%></td>
				<td><%=fnSetTimeFormat(totWorkTime_sum)%></td> 
				<td><%=fnSetTimeFormat(totextendWT_sum)%></td>
				<td><%=fnSetTimeFormat(totnightWT_sum)%></td>
				<td><%=fnSetTimeFormat(totholidayWT_sum)%></td>
				<td><%=fnSetTimeFormat(totweekholidayWT_sum)%></td>
		</tr>
		 
		<%end if
		end if
		 
				totWorkTime 	= totWorkTime + iWorkTime
				totextendWT  	= totextendWT + iextendWT
				totnightWT	= totnightWT + inightWT
				totholidayWT	= totholidayWT + iholidayWT
				totweekholidayWT= totweekholidayWT  + iweekholidayWT 
	 			totVacationTime = totVacationTime + ivacationTime
		
			 
		%>
		<tr   bgcolor="#FFFFFF" align="center">
			<%if right(dyyyymmdd,2) = 32 then '추가 주휴수당%>
			<td colspan="9">추가주휴수당</td>
			<%else%>
			<td><%=right(dyyyymmdd,2)%></td>
			<td><%=fnGetStringWD(dWeekday)%><input type="hidden" name="hidWD<%=intD%>" value="<%=dWeekday%>"></td>
			<td>
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
			</td>		 
			<td><%IF dstartHour<>"00" or dstartMinute<>"00" THEN%><b><%END IF%><%=dstartHour%>:<%=dstartMinute%></b></td>		                                                                     
			<td><%IF dendHour<>"00" or dendMinute<>"00" THEN%><b><%END IF%><%=dendHour%>:<%=dendMinute%></b></td>		                                                                     
			<td><%IF dbreakSHour<>"00" or dbreakSMinute<>"00" THEN%><b><%END IF%><%=dbreakSHour%>:<%=dbreakSMinute%></b></td>		 
			<td><%IF dbreakEHour<>"00" or dbreakEMinute<>"00" THEN%><b><%END IF%><%=dbreakEHour%>:<%=dbreakEMinute%></b></td>
			<td><%IF doutHour<>"00" or doutMinute<>"00" THEN%><b><%END IF%><%=doutHour%>:<%=doutMinute%></b></td>
			<td><%=defaulttime(dWeekday)%></td>
			<%end if%>
			<td><%IF iVacationTime <>"0" THEN%><b><%END IF%><%=fnSetTimeFormat(iVacationTime)%></b></td>
			<td><%IF iWorkTime <>"0" THEN%><b><%END IF%><%=fnSetTimeFormat(iWorkTime)%></b></td>  
			<td><%IF iextendWT <>"0" THEN%><b><%END IF%><%=fnSetTimeFormat(iextendWT)%></b></td>
			<td><%IF inightWT <>"0" THEN%><b><%END IF%><%=fnSetTimeFormat(inightWT)%></b></td>
			<td><%IF iholidayWT <>"0" THEN%><b><%END IF%><%=fnSetTimeFormat(iholidayWT)%></b></td> 
			<td><%IF iweekholidayWT <>"0" THEN%><b><%END IF%><%=fnSetTimeFormat(iweekholidayWT)%></b></td>	
		</tr>	
		<%Next%>
		<%	END IF	%>
		<%if  dYear&"-"&format00(2,dMonth) >= "2017-01" then%> 
		<tr   bgcolor="<%=adminColor("sky")%>" align="center">			
			<td colspan="9">[<%=dMonth%>/1 ~ <%=dMonth%>/<%=right(dyyyymmdd,2)%>] 합계</td>
			<td><B><%=fnSetTimeFormat(totVacationTime)%></b></td>
			<td><B><%=fnSetTimeFormat(totWorkTime)%></b></td> 
			<td><B><%=fnSetTimeFormat(totextendWT)%></b></td>
			<td><B><%=fnSetTimeFormat(totnightWT)%></b></td>
			<td><B><%=fnSetTimeFormat(totholidayWT)%></b></td> 
			<td><B><%=fnSetTimeFormat(totweekholidayWT)%></b></td>		
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
			<td colspan="9">  <%if   dYear&"-"&format00(2,dMonth) >= "2017-01" then%> [<%=dpreMonth%>/26 ~ <%=dMonth%>/<%=right(dyyyymmdd,2)%>]<%end if%>   <b>총 합계</b></td> 
			<td><B><%=fnSetTimeFormat(totVacationTime)%></b></td>
			<td><B><%=fnSetTimeFormat(totWorkTime)%></b></td> 
			<td><B><%=fnSetTimeFormat(totextendWT)%></b></td>
			<td><B><%=fnSetTimeFormat(totnightWT)%></b></td>
			<td><B><%=fnSetTimeFormat(totholidayWT)%></b></td> 
			<td><B><%=fnSetTimeFormat(totweekholidayWT)%></b></td>		
		</tr> 
		 
		</table>
	</td>
</tr> 
</table>
</body>
</html>
