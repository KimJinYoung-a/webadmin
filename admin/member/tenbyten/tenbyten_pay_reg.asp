<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����� �޿� ���� ���
' History : 2010.12.27 ������  ����
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

sEmpno	= requestCheckvar(request("sEN"),14)	'���
dYear	= requestCheckvar(request("selY"),4)	'��
dMonth	= requestCheckvar(request("selM"),2)	'��
ino		= requestCheckvar(request("ino"),10)	'ȸ��
blnReset = requestCheckvar(request("blnR"),1) '���¿���
'�⺻�� ���� (���� ���)
IF dYear = "" THEN dYear  = Year(Date())
IF dMonth = "" THEN dMonth  = Month(Date())

'// 4.345238095 == �� ��� WEEK �� = (365�� / 7�� / 12����)
avgWeek = 4.345238095

totWorkDay = 0
totWorkDayReal = 0
totReWorkDay = 0
totReWorkDayReal =0

monthlyPayDataExist = True
dim dSPayDate,dEPayDate,dPreYear,dPreMonth ,preEndDay
dim chkDate 

preEndDay = dateadd("d", -1, dateserial(dYear,dMonth,1)) '������  ������ �� 
dPreYear = year(preEndDay) '������ ��
dPreMonth = month(preEndDay) '������ ��
dNextDate = dateadd("m",1, dateserial(dYear,dMonth,1))	'�˻������� 1��
dEndDate = dateadd("d",-1,dNextDate)
dEndDay = day(dEndDate)
chkDate =dYear&"-"&format00(2,dMonth)

'------------------------------------------------------------------ 
IF chkDate  = "2014-01" THEN '2014.01���� �޿������� 25�Ϸ� ����� 
	dSPayDate = dateserial(dYear,dMonth,1) '�޿�������: �ش�� 1�Ϻ���
	dEPayDate = dateserial(dYear,dMonth,25) '�޿�������: �ش�� 25�ϱ���  
ELSEIF chkDate > "2014-01" and chkDate < "2016-12" THEN 
	dSPayDate = dateserial(dPreYear,dPreMonth,26) '�޿�������: ������ 26�Ϻ���
	dEPayDate = dateserial(dYear,dMonth,25) '�޿�������: �ش�� 25�ϱ���   
ELSEIF chkDate >= "2016-12"	then '���� 26�Ϻ���~ �ش�� ���ϱ���
	dSPayDate = dateserial(dPreYear,dPreMonth,26) 
	dEPayDate = dateserial(dYear,dMonth,dEndDay)  '�޿�������: �ش�� ���ϱ��� 
ELSE    
	dSPayDate = dateserial(dYear,dMonth,1) '�޿�������: �ش�� 1�Ϻ���
	dEPayDate = dateserial(dYear,dMonth,dEndDay)  '�޿�������: �ش�� ���ϱ��� 
END IF  
'------------------------------------------------------------------ 
set clsPay = new CPay
	'// ========================================================================
	'// ��� �⺻�������
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
	totWorkDay		= ceilValue(clsPay.FWeekWorkDay * avgWeek)		'// �⺻ �ٹ��ϼ�
 
	'// ========================================================================
	'// ����� �������
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
	'// �ñ����� ��� ����� �ٹ��ϼ� ��������(���� ����Ÿ ����)
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
				'// 4�ð� �̻� �ٹ��� �ٹ��ϼ� �߰�
				totWorkDayReal = totWorkDayReal + 1
			end if
		Next
	end if

end if

IF  iworktime ="" or isNull(iworktime) or blnReset = "1" THEN
	'// ========================================================================
	'// ����������� ������(�Ǵ� �� �޿� �����) �⺻ ��࿡�� ����Ÿ �����´�.
	'// ========================================================================

	monthlyPayDataExist = False
 
	IF iposit_sn=12 or iposit_sn=14 or iposit_sn=15 THEN	'������(����/����/����)
		iworktime    	= (ceilValue(totDutyTime/60*avgWeek)+ceilValue(holidaywdtime/60*avgWeek))*60
		iextendtime 	= iOverTime
		inighttime    	= ceilValue(totNightTime/60*avgWeek)*60
		iholidaytime   	=  0
		mtimepay       	= defaultpay*ceilValue(totDutyTime/60*avgWeek)+ defaultpay*ceilValue(holidaywdtime/60*avgWeek)
		if (foodpay=0) then
		    mfoodpay		= 0
		else
		    mfoodpay		= ceilValue(totWorkDay * foodpay)   '' totWorkDay �� ���̶� �ϴ� foodpay 0����üũ '' �󱸾��� �۾������ε�
	    end if
		mextendpay     	= defaultpay*iOverTime*1.5
		mnightpay      	= defaultpay*ceilValue(totNightTime/60*avgWeek)*0.5
		mholidaypay		= 0
		mtotpay        	= totPaySum 

		IF blnstatediv ="N"  and  left(dretireday,7) =  dYear&"-"&format00(2,dMonth) and dretireday < dEndDate  and dretireday <= enddate  THEN
			'����� ��� ������� �˻��� ������ ��¥���� ������  ����ϱ��� �� �ݾ׿��� ��¥�� ������.

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
		   '������ ������ ��� �������� ���� ��� ��������ϱ��� �� �ݾ׿��� ��¥�� ������.
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

			' �� �߰� �Ի����� ��� ..(���ӱ�/�ش� �� �ϼ�)*(�ش� �� ������ �� - �Ի��� + 1)
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
	 ELSE	'�ñ���
	 	 
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
							'// �ѽð� �̻� �ٹ��� �ٹ��ϼ� �߰�
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
	 		alert("���� �� ����������  �˻� �����մϴ�.");
	 		return;
	 	}else if (document.frmSearch.selY.value == dNowYear && document.frmSearch.selM.value > dNowMonth){
	 		alert("���� �� ����������  �˻� �����մϴ�.");
	 		return;
	 	}

	 	//���Ⱓ ������ �˻� �����ϵ��� ����
	 	document.frmSearch.submit();
  	}

  	//������� ���
	function jsViewPay(empno,ino){
		var wpay = window.open("pop_payform.asp?sEN="+empno+"&ino="+ino,"popPay","width=700,height=600,scrollbars=yes,resizable=yes");
		wpay.focus();
	}

  	//�ٹ��ð� ���
 	function jsWorkTime(empno,ino){
 		var wwt =window.open("pop_worktime.asp?sEN="+empno+"&ino="+ino+"&selY=<%=dYear%>&selM=<%=dMonth%>","popWT","width=1200,height=800,scrollbars=yes,resizable=yes");
		wwt.focus();
 	}

 	//�� �հ�ݾ� ����
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

 	//�޿����
 	function jsSubmit(){
		var strMsg,istate;
		for(i=0;i<document.frmPay.hidS.length;i++){
			if(document.frmPay.hidS[i].checked){
				istate = document.frmPay.hidS[i].value;
			}
		}
		jsSetTotPay('iFP');

 		if(istate == 1){
 			strMsg = "�ۼ��Ϸ���·� ����Ͻðڽ��ϱ�?" ;
 		}else if(istate == 5){
 			strMsg = "Ȯ�οϷ���·� ����Ͻðڽ��ϱ�?" ;
 		}else if(istate == 7){
 			strMsg = "�ԱݿϷ���·� ����Ͻðڽ��ϱ�?" ;
 		}else if(istate == 0){
 			strMsg = "�޿��ۼ��߻��·� ����Ͻðڽ��ϱ�?" ;
 		}
 		if(confirm(strMsg)){
 			return true;
 		} else {
			return false;
		}

 	}

 	//����Ʈ
 	function jsPrint(){
 	 var winPrint = window.open("print_worktime.asp?sEN=<%=sempno%>&ino=<%=ino%>&selY=<%=dYear%>&selM=<%=dMonth%>","prtWT","width=1020,height=600,scrollbars=yes,resizable=yes");
 	 winPrint.focus();
 	}

 	//�� �޿� ����
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
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">���</td>
			<td bgcolor="#FFFFFF" width="180"><a href="javascript:jsViewPay('<%=sempno%>','<%=ino%>')"><%=sempno%></a> <%IF blnstatediv ="N" THEN%><font color="red">[���]</font><%END IF%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�Ի���</td>
			<td bgcolor="#FFFFFF"><%=formatdate(djoinday,"0000-00-00")%></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�̸�</td>
			<td bgcolor="#FFFFFF"><%=susername%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�����</td>
			<td bgcolor="#FFFFFF">
				<%IF blnstatediv = "N" THEN%>
					<% if Not IsNull(dretireday) then %>
						<%=formatdate(dretireday,"0000-00-00")%>
					<% else %>
						<font color="red">���� : �ý����� ����</font>
					<% end if %>
				<%END IF%>
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">��౸��</td>
			<td bgcolor="#FFFFFF"><%=sposit_name%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ð���</td>
			<td bgcolor="#FFFFFF"><%if predefaultpay>0 then%>(����: <%=formatnumber(predefaultpay,0)%> ��) <%end if%><%=formatnumber(defaultpay,0)%> ��</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ٹ��ϼ�</td>
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
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�Ĵ�</td>
			<td bgcolor="#FFFFFF"><%if prefoodpay>0 then%>(����: <%=formatnumber(prefoodpay,0)%> ��) <%end if%><%=formatnumber(foodpay,0)%> ��</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�����</td>
			<td bgcolor="#FFFFFF">[<%=ino%>] <%=formatdate(startdate,"0000-00-00")%> ~ <%=formatdate(enddate,"0000-00-00")%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ް�ð�</td>
			<td bgcolor="#FFFFFF"><%IF inbreaktime THEN%>�ٹ��ð� ����<%ELSE%>�ٹ��ð� ���Ծ���<%END IF%></td>
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
	<td>�ٹ���¥:
		<select name="selY">
		<%For intY = Year(date()) To 2010 Step -1%>
		<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dYear) THEN%>selected<%END IF%>><%=intY%></option>
		<%Next%>
		</select>
		��
		<select name="selM">
		<%For intM = 1 To 12%>
		<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dMonth) THEN%>selected<%END IF%>><%=intM%></option>
		<%Next%>
		</select>
		��
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch();">
		&nbsp;&nbsp;
<%
	If Request.ServerVariables("REMOTE_ADDR") = "61.252.133.15" Then
%>
<input type="button" class="button" value="�Ϻ� �ٹ��ð�" onClick="jsWorkTime('<%=sempno%>','<%=ino%>');">
<%
	End If
%>

		<%IF iposit_sn = 13 THEN%><input type="button" class="button" value="�Ϻ� �ٹ��ð�" onClick="jsWorkTime('<%=sempno%>','<%=ino%>');"><%END IF%>
		 <input type="button" value="����Ʈ" class="button" onClick="jsPrint();">
		 <%IF (iposit_sn=12 or iposit_sn=14 or iposit_sn=15) and istate  = 0 THEN%><input type="button" class="button" value="�� �޿� ����" onClick="jsRestWorkTime();"> <br><div style="padding-top:5px"><font color="Red">* [�� �޿� ����]�� ������ Ȯ�� �� [���]��ư�� �� �����ּ���. ��� ��ư ��ó���� ���� ���� �����ͷ� ó���˴ϴ�. </font></div><%END IF%>
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
		<input type="hidden" name="hidSPdate" value="<%= dSPayDate %>"><!-- �޿��������� �����-->
		<input type="hidden" name="hidEPdate" value="<%= dEPayDate %>"><!-- �޿��������� �����-->
		<tr>
		<td  bgcolor="<%= adminColor("gray") %>" width="120" align="center">�޿���ϻ���</td>
		<td bgcolor="#FFFFFF">
				<input type="radio" name="hidS" value="0" <%IF istate  = 0  THEN%>checked<%ELSEIF istate >1 and not(C_ADMIN_AUTH or C_PSMngPart)  THEN%>disabled<%END IF%>><%IF istate  = 0  THEN%><font color="red"><%END IF%>�޿��ۼ��� ></font>
				<input type="radio" name="hidS" value="1" <%IF istate  = 1  THEN%>checked<%ELSEIF istate >1 THEN%>disabled<%END IF%>><%IF istate  = 1  THEN%><font color="red"><%END IF%>�ۼ��Ϸ� ></font>
				<input type="radio" name="hidS" value="5" <%IF istate  = 5  THEN%>checked<%END IF%>><%IF istate  = 5  THEN%><font color="red"><%END IF%>�濵����Ȯ�οϷ� ></font>
				<input type="radio" name="hidS" value="7" <%IF istate  = 7  THEN%>checked<%END IF%>><%IF istate  = 7  THEN%><font color="red"><%END IF%>�ԱݿϷ� </font>
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
		<input type="hidden" name="totWorkDay" value="<%=totWorkDay%>"><!-- �ñ�=�Ǳٹ��ϼ� / ������=�⺻�ٹ��ϼ�(�߰��Ի������ ��쵵 ����) -->
		<input type="hidden" name="totWorkDayReal" value="<%=totWorkDayReal%>">
		<input type="hidden" name="totReWorkDay" value="<%=totReWorkDay%>"><!-- �ñ�=�Ǳٹ��ϼ� / ������=�⺻�ٹ��ϼ�(�߰��Ի������ ��쵵 ����) -->
		<input type="hidden" name="totReWorkDayReal" value="<%=totReWorkDayReal%>">
		<tr  bgcolor="<%= adminColor("gray") %>" align="center">
			<td>����</td>
			<td>�⺻��</td>
			<td>�ð��ܼ���</td>
			<td>�߰��ٹ�����</td>
			<td>���ϱٹ�����</td>
			<td>�Ĵ�����</td>
			<td>��å����</td>
			<td>������</td>
			<td>���ټӼ���</td>
			<td>�߰�����</td>
			<td>��������</td>
			<td>�󿩱�</td> 
			<td>�Ѿ�</td>
		</tr>
		<%IF iposit_sn = 13 THEN %>
						<%if sempno= "90201501120013" or sempno="90201610010124" or sempno="90201611130141" or sempno="90201611140136" or sempno="90201611200158" or sempno="90201611260172" or sempno="90201612060169" or sempno="90201612100180" or sempno="90201612120174" or sempno="90201612210190" then%>
							<tr  bgcolor="#FFFFFF" align="center">
							<td bgcolor="<%= adminColor("gray") %>">����ݾ�</td>
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
							<td bgcolor="<%= adminColor("gray") %>">����ð�</td>
							<td><%=fnSetTimeFormat(iWorkTime)%></td>
							<td><%=fnSetTimeFormat(iextendtime)%></td>
							<td><%=fnSetTimeFormat(inighttime)%></td>
							<td><%=fnSetTimeFormat(iholidaytime)%></td>
							<td colspan="8" align="left">
								* �ٹ��ϼ� : 
								<% if (monthlyPayDataExist = True) then %>
									<%= mworkday %>��
									<% if (mworkday <> totWorkDayReal) or (foodpay <> 0 and totWorkDay <> 0 and mfoodpay = 0) then %>
										<font color="red">(�Ǳٹ��ϼ� : <%= totWorkDayReal %>��)</font>
										<input type="button" class="button" value="�Ǳٹ��ϼ� ����" onClick="jsSetRealWorkDayToSaved('N')">
									<% end if %>
								<% else %>
									<%= totWorkDay %>
								<% end if %>
							</td>
						</tr> 
						<%ELSE%>
						<tr  bgcolor="<%=adminColor("sky")%>" align="center">
							<td bgcolor="<%= adminColor("sky") %>"><b>�ѱݾ�</b></td>
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
							<td bgcolor="<%=adminColor("sky")%>">�ѽð�</td>
							<td><%=fnSetTimeFormat(iWorkTime+iReWorktime)%></td>
							<td><%=fnSetTimeFormat(iextendtime+ireextendtime)%></td>
							<td><%=fnSetTimeFormat(inighttime+irenighttime)%></td>
							<td><%=fnSetTimeFormat(iholidaytime+ireholidaytime)%></td>
							<td colspan="8" align="left">
								* �ٹ��ϼ� : 
								<% if (monthlyPayDataExist = True) then %>
									<%= mworkday+ireworkday %>�� 
								<% else %>
									<%= totWorkDay+totReworkday %>
								<% end if %>
							</td>
						</tr>
						<tr  bgcolor="#FFFFFF" align="center">
							<td bgcolor="<%= adminColor("gray") %>">����ݾ�</td>
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
							<td bgcolor="<%= adminColor("gray") %>">����ð�</td>
							<td><%=fnSetTimeFormat(iWorkTime)%></td>
							<td><%=fnSetTimeFormat(iextendtime)%></td>
							<td><%=fnSetTimeFormat(inighttime)%></td>
							<td><%=fnSetTimeFormat(iholidaytime)%></td>
							<td colspan="8" align="left">
								* �ٹ��ϼ� : 
								<% if (monthlyPayDataExist = True) then %>
									<%= mworkday %>�� 
									<% if (mworkday <> totWorkDayReal) or (foodpay <> 0 and totWorkDay <> 0 and mfoodpay = 0) then %>
										<font color="red">(�Ǳٹ��ϼ� : <%= totWorkDayReal %>��)</font>
										<input type="button" class="button" value="�Ǳٹ��ϼ� ����" onClick="jsSetRealWorkDayToSaved('N')">
									<% end if %>
								<% else %>
									<%= totWorkDay %>
								<% end if %>
							</td>
						</tr>
						<tr  bgcolor="#e3f1fb" align="center">
							<td bgcolor="#e3f1fb" nowrap>�������ױ�</td>
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
							<td bgcolor="#e3f1fb" nowrap>�������׽ð�</td>
							<td><%=fnSetTimeFormat(ireWorkTime)%></td>
							<td><%=fnSetTimeFormat(ireextendtime)%></td>
							<td><%=fnSetTimeFormat(irenighttime)%></td>
							<td><%=fnSetTimeFormat(ireholidaytime)%></td>
							<td colspan="8" align="left">
								* �ٹ��ϼ� :
								<% if (monthlyPayDataExist = True) then %>
									<%= ireworkday %>��
									<% if (ireworkday <> totReWorkDayReal) or (prefoodpay <> 0 and totReWorkDay <> 0 and mrefoodpay = 0) then %>
										<font color="red">(�Ǳٹ��ϼ� : <%= totReWorkDayReal %>��)</font>
										<input type="button" class="button" value="�Ǳٹ��ϼ� ����" onClick="jsSetRealWorkDayToSaved('P')">
									<% end if %>
								<% else %>
									<%= totreWorkDay %>
								<% end if %>
							</td>
						</tr>
						<%	END IF%>
		<%ELSE%>			
			<tr  bgcolor="#FFFFFF" align="center">
							<td bgcolor="<%= adminColor("gray") %>">�ѱݾ�</td>
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
			Response.Write "<input type=""submit"" value=""���"" class=""button"">"
		Else
			If istate >= 5 Then
				If C_ADMIN_AUTH or C_PSMngPart Then
					Response.Write "<input type=""submit"" value=""���"" class=""button"">"
				Else
					Response.Write "�� �޿���ϻ��°� <font color=blue><b>[�ۼ��Ϸ�]�� ��� ����</b></font>�� <font color=red><b>�濵������ - �λ米����Ʈ ������ ���� ����</b></font>�մϴ�."
					Response.Write "<br>�̿� ���� ���Ǵ� �濵������ - �λ米����Ʈ(070-7515-5440)�� �����Ͻñ� �ٶ��ϴ�."
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
