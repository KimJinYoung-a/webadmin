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
<!-- #include virtual="/lib/classes/member/fingerprints/fingerprints_cls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%

'// ���� ������������ �ð��� ���õ� ���븸 ����Ѵ�.
'// �ݾ� ���õ� ������ ó������ �ʴ´�.

'��������
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
'�� �޾ƿ���
sempno= requestCheckvar(Request("sEN"),14)
ino= requestCheckvar(Request("ino"),10)
dYear = requestCheckvar(Request("selY"),4)
dMonth = requestCheckvar(Request("selM"),2)
chkDate = dYear&"-"&format00(2,dMonth)
sFingerYN = requestCheckvar(Request("sFYN"),1) '�����νı��°��� �����Դ��� ����
sVacationYN = requestCheckvar(Request("sVYN"),1)

'�⺻�� ���� (���� ���)
IF dYear = "" THEN dYear  = Year(Date())
IF dMonth = "" THEN dMonth  = Month(Date())

dNextDate = dateadd("m",1, dateserial(dYear,dMonth,1))	'�˻������� 1��
dEndDay = day(dateadd("d",-1,dNextDate))	'�˻��� ������ ��¥ (������ 1�� - �Ϸ�)  
preEndDay = dateadd("d", -1, dateserial(dYear,dMonth,1)) '������  ������ ��
'preStartDay = day(dateadd("d", -(weekday(preEndDay)-1),preEndDay)) 
dPreYear = year(preEndDay) '������ ��
dPreMonth = month(preEndDay) '������ ��
preEndDay  =day(preEndDay)  '������ ������ ��¥
 
''�޿���� ������ ���� 2014/01(1~31 => 26~25)
'------------------------------------------------------------------ 
IF  dYear&"-"&format00(2,dMonth)  = "2014-01" THEN '2014.01���� �޿������� 25�Ϸ� �����
	dSPayDay = 1 '�޿�������
	dEPayDay	= 25 '�޿�������   
	dSPayDate = dateserial(dYear,dMonth,dSPayDay) '�޿�������: �ش�� 1�Ϻ���
	dEPayDate = dateserial(dYear,dMonth,dEPayDay) '�޿�������: �ش�� 25�ϱ��� 
	iLoopCnt = dEPayDay	- 1'�� �޿��ϼ� 
ELSEIF dYear&"-"&format00(2,dMonth) > "2014-01"  and chkDate <"2016-12" then
	dSPayDay = 26 '�޿�������
	dEPayDay	= 25 '�޿�������   
	dSPayDate = dateserial(dPreYear,dPreMonth,dSPayDay) '�޿�������: ������ 26�Ϻ���
	dEPayDate = dateserial(dYear,dMonth,dEPayDay) '�޿�������: �ش�� 25�ϱ��� 
	iLoopCnt = (preEndDay-dSPayDay)+dEPayDay	'�� �޿��ϼ� 
ELSEIF chkDate = "2016-12"   then
	dSPayDay = 26 '�޿������� 
	dEPayDay	= dEndDay '�޿�������   
	dSPayDate = dateserial(dPreYear,dPreMonth,dSPayDay) '�޿�������: ������ 26�Ϻ���	 
	dEPayDate = dateserial(dYear,dMonth,dEPayDay) '�޿�������: �ش�� 25�ϱ��� 
	iLoopCnt = (preEndDay-dSPayDay)+dEPayDay	'�� �޿��ϼ� 	
ELSE
	dSPayDay = 1
	dEPayDay	= dEndDay
	dSPayDate = dateserial(dYear,dMonth,dSPayDay) '�޿�������: �ش�� 1�Ϻ���
	dEPayDate = dateserial(dYear,dMonth,dEPayDay)  '�޿�������: �ش�� ���ϱ���
	iLoopCnt	= dEndDay - 1 '�� �޿��ϼ� 
END IF  
	dPreEPayDay = day(dateadd("d", -1, dSPayDate))'������ �޿� ������
'------------------------------------------------------------------
 
'������ ��������
set clsPay = new CPay
	'// ========================================================================
	'--��� �⺻������� ��������
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
	'// ���ۼ��� ���� �ٹ��ð��� �ִ� ��� ��������
	'// ========================================================================
	clsPay.fnGetmonthlypayData
	mstate = clsPay.Fstate

	'// ========================================================================
	'// ������� ù �Ͽ��Ͽ� �ش��ϴ� ������ �ϼ� + �� ���� ������ ��ϰ�������
	'// 2013-02-01 ���ΰ�� 2013-01-27 ������ ��� + �� ���� 1���ϸ��
	'// ������ ������ ����Ѵ�.
	'// ========================================================================
	clsPay.FPreyyyymmdd = dSPayDate
	arrPre =clsPay.fnGetPreDailypayData
 
	'--�˻��� �ٹ��ð� ����
	IF sFingerYN = "Y" THEN    '�����νĳ��� ������ ���
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
			 Alert_return("�����νı��³����� �������� �ʽ��ϴ�. Ȯ�� �� �ٽ� �õ����ּ���")
			response.end
		END IF
	ELSE 
		clsPay.FSyyyymm = dSPayDate
		clsPay.FEyyyymm = dEPayDate
		arrList = clsPay.fnGetDailypayData
		set clsPay = nothing
	END IF

	'// ========================================================================
	'// �ް���� ��������
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
				'// ���δ��
				vacationRequestCount = vacationRequestCount + 1
			end if
		next

		if (vacationRequestCount > 0) then
			response.write "<script>alert('���δ�� ������ �ް��� �ֽ��ϴ�.\n\n���� �����ؾ� �ް������� ������ �� �ֽ��ϴ�.');</script>"
		end if
	end if

IF dYear >= 2011 and susername ="" or isnull(susername) THEN
	IF Request("selY") = "" THEN
%>
	<script language="javascript">
	alert("��������� �������� �ʰų� �ش� ȸ���� �ش��ϴ� ��¥�� �ȵƽ��ϴ�.  Ȯ���� �ٽ� �õ����ּ���");
	self.close();
	</script>
<%
	ELSE
	Alert_return("��������� �������� �ʰų� �ش� ȸ���� �ش��ϴ� ��¥�� �ȵƽ��ϴ�.  Ȯ���� �ٽ� �õ����ּ���")
	END IF

END IF 

IF datediff("m",startdate,dEPayDate)  < 0 or datediff("m",enddate,dateadd("m",-1,dEPayDate)) > 0 THEN
	dstate = 9
END IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<html>
<head>
<title>�ٹ��ð� ���</title>
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
//	 		alert("���� �� ����������  �˻� �����մϴ�.");
//	 		return;
//	 	}else if (document.frmSearch.selY.value == dNowYear && document.frmSearch.selM.value > dNowMonth){
//	 		alert("���� �� ����������  �˻� �����մϴ�.");
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

		// * ���޽ð�
		//  - �������� �Ͽ��Ϻ��� ����ϱ����̴�.
		//  - ���� ��ٽ� ���޽ð� 0 �ð�
		//  - �ְ� �ٹ��ð� 15�ð� �̸��� ��� ���޽ð� 0 �ð�
		//  - 15�ð� �̻� �ٹ��� ���޽ð� = (�ð�/40)*8 �ð�
		//  - �ٹ��ð� 40�ð� �ʰ��� ���޽ð� 8 �ð�

		// * ������ ������ üũ�Ѵ�.
		//  - ���޽ð� �Է��� jsSetTotTimeALL() ���� �Ѵ�.

		LastDayOfThisMonth = document.frmPay.hidEday.value*1; 
		for(var i = 0; i <= LastDayOfThisMonth; i++) {
			
			// =================================================================
			// �ٹ��ð� üũ
			startMinFromMidnight = parseInt(eval("document.frmPay.iSH"+i).value,10)*60+parseInt(eval("document.frmPay.iSM"+i).value,10);
			endMinFromMidnight = parseInt(eval("document.frmPay.iEH"+i).value,10)*60+parseInt(eval("document.frmPay.iEM"+i).value,10);
 
			if( startMinFromMidnight > endMinFromMidnight) {
				alert("�ٹ����۽ð��� �ٹ�����ð����� ������մϴ�. �ٽ� �������ּ���");
				eval("document.frmPay.iSH"+i).focus();
				return false;
			}
 
			// =================================================================
			// �ްԽð� üũ
			startMinFromMidnight = parseInt(eval("document.frmPay.iBSH"+i).value,10)*60+parseInt(eval("document.frmPay.iBSM"+i).value,10);
			endMinFromMidnight = parseInt(eval("document.frmPay.iBEH"+i).value,10)*60+parseInt(eval("document.frmPay.iBEM"+i).value,10);

			if( startMinFromMidnight >  endMinFromMidnight) {
				alert("�ްԽ��۽ð��� �ް�����ð����� ������մϴ�. �ٽ� �������ּ���");
				eval("document.frmPay.iBSH"+i).focus();
				return false;
			}

			// =================================================================
			// �� �޿�  ù��° �Ͽ��Ͽ� �ش��ϴ� ��¥
			if(eval("document.frmPay.hidWD"+i).value  == 1 && swday == 0) {
				swday =  i;
			}
		
		}
 
		itwt = 0;
		blnPWH = 0;

		itwt = document.frmPay.hidPWD.value*1; 			// ���� �������� �ѱٹ��ð�
		blnPWH = document.frmPay.blnPWH.value*1; 			// ���� �������� �����ϼ�
		hidPWHD = document.frmPay.hidPWHD.value*1; 		// ���� �������� ���Ƚ��

		hidCWD = document.frmPay.hidCWD.value*1;		// �̹��� ���޺κ�
		blnCWH = document.frmPay.blnCWH.value*1;
		hidCWHD = document.frmPay.hidCWHD.value*1;

		if (blnPWH > 1) {
			// ���� �������� 2�� �̻� �ԷµȰ� ����
			blnPWH = 1;
		}

		if (blnCWH > 1) {
			// �̹��� ���޺κ� �������� 2�� �̻� �ԷµȰ� ����
			blnCWH = 1;
		}

		// =====================================================================
		// 01. ù° �� üũ
		// =====================================================================
		if (swday == 1) {
			// ���� ���� 1���� �Ͽ����� ���
		} else {
			// ���� ù��° ���� �Ͽ����� �ƴ� ���
			// * ���� ������ �Ͽ��� ������ ��¥�� ������� ù��° ����ϱ����� ��¥�� ���ļ� ���޽ð� üũ�Ѵ�.
	 		for (var i = 0; i < swday; i++) {
				// �����ϼ�
		 		if (eval("document.frmPay.selWH"+i).value == "3") {
		 	  		blnCWH = blnCWH + 1;
		 		}

				// �ٹ��ð�
				arrValue = eval("document.frmPay.iWT"+i).value.split(":");
				iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
				hidCWD = hidCWD + iValue;

				if ((eval("document.frmPay.selWH"+i).value == "1") && (iValue == 0)) {
					hidCWHD = hidCWHD + 1;
				}
			}

			if (blnCWH > 1) {
				alert("������(ù°��)�� �����Ͽ� �ѹ��� ���� �����մϴ�.");
				return false;
			} else if ((itwt >= 900) && (blnCWH < 1)) {
				alert("������(ù°��)�� �������ּ���.");
				return false;
			}
		}

		// =====================================================================
		// 02. ù° �� ����(�Ǵ� 1���� �Ͽ����� ���) üũ
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

			// �ٹ��ð�
			arrValue = eval("document.frmPay.iWT"+i).value.split(":");
			iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
			hidCWD = hidCWD + iValue;

			if ((eval("document.frmPay.selWH"+i).value == "1") && (iValue == 0)) {
				hidCWHD = hidCWHD + 1;
			}

			if ((eval("document.frmPay.hidWD"+i).value*1 == 7) || (i == LastDayOfThisMonth)) {
				if (blnCWH > 1) {
					alert("�������� �����Ͽ� �ѹ��� ���� �����մϴ�.");
					return false;
				} else if ((itwt >= 900) && (blnCWH < 1)) {
				    if (i*1!=LastDayOfThisMonth*1){ //�߰�
					    alert("�������� �������ּ���...." + itwt + " " + blnCWH + " " +LastDayOfThisMonth+ " " +i);
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
		jsSetTotTimeALL(LastDayOfThisMonth + 1); //����� ����ó��
 
	 	return true;
	 
	}

	function jsAddWH(){
	 //����� ������ ����
		<%IF blnstatediv = "N" THEN
			IF   dretireday >= dSPayDate and dretireday<= dEPayDate THEN	'������̰� ������� �̹��� �ٹ��Ͽ� ���Ե�  ���%>

			if ((<%=day(dretireday)%>+(7-<%=weekday(dretireday)%>)) >= document.frmPay.hidEPday.value){ //����� ���������� ����Ϻ��� ũ�ų� ���� ���
			 
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
					for(i=iLSday;i<=iLEday;i++){ //����� �� �ٹ��ð� ���
						if(eval("document.frmPay.iWT"+i).value!=""){
							arrValue = eval("document.frmPay.iWT"+i).value.split(":");
							iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);

							arrValue = eval("document.frmPay.ieWT"+i).value.split(":");
							iEValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
						}

						if( eval("document.frmPay.selWH"+i).value=="1" && ( iValue+iEValue) == 0 ){//�ٹ��Ͽ� �ٹ��ߴ��� ���� Ȯ��
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
					if (iLWHD == 0){ //�ٹ� �������ֿ� �������� ���� ��� ���� �ٹ� ���޼��� ���Կ��� �ش�.
						iLSday = parseInt(iLSday,10)-7;
						iLEday = parseInt(iLSday,10)+6;

						for(i=iLSday;i<=iLEday;i++){
							if(eval("document.frmPay.iWT"+i).value!=""){
								arrValue = eval("document.frmPay.iWT"+i).value.split(":");
								iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);

								arrValue = eval("document.frmPay.ieWT"+i).value.split(":");
								iEValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
							}

							if( eval("document.frmPay.selWH"+i).value=="1" && ( iValue+iEValue) == 0 ){//�ٹ��Ͽ� �ٹ��ߴ��� ���� Ȯ��
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
		if(confirm("�ٹ��ð������ �ۼ��Ϸ��Ͻðڽ��ϱ�? �ۼ��Ϸ�� ���޿��� �����˴ϴ�.")){
			document.frmPay.hidS.value="1";
			document.frmPay.submit();
		}
		return;
	}

	//�����νı��� ���� ��������
	function jsGetFinger(){
		document.frmSearch.sFYN.value  = "Y";
		document.frmSearch.submit();
	}

	//�ް���û���� ��������
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
	<td><strong>�������� �ٹ��ð� ���</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">���</td>
			<td bgcolor="#FFFFFF" width="180"><%=sempno%> <%IF blnstatediv ="N" THEN%><font color="red">[���]</font><%END IF%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�Ի���</td>
			<td bgcolor="#FFFFFF"><%IF djoinday <> "" THEN%><%=formatdate(djoinday,"0000-00-00")%><%END IF%></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�̸�</td>
			<td bgcolor="#FFFFFF"><%=susername%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�����</td>
			<td bgcolor="#FFFFFF"><%IF blnstatediv = "N" THEN%><%=formatdate(dretireday,"0000-00-00")%><%END IF%></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">��౸��</td>
			<td bgcolor="#FFFFFF"><%=sposit_name%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ð���</td>
			<td bgcolor="#FFFFFF"><%=formatnumber(defaultpay,0)%> ��</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�����</td>
			<td bgcolor="#FFFFFF"><%IF startdate <> "" THEN%><%=ino%>. <%=formatdate(startdate,"0000-00-00")%> ~ <%=formatdate(enddate,"0000-00-00")%><%END IF%></td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ް�ð�</td>
			<td bgcolor="#FFFFFF"><%IF inbreaktime THEN%>�ٹ��ð� ����<%ELSE%>�ٹ��ð� ���Ծ���<%END IF%></td>
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
				�ٹ���¥:
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
			</td>
			<td align="right">
				<%IF mstate = 0 THEN%>
					<input type="button" class="button" value="�ް���û���� ��������" onClick="jsGetVacation();">
					<input type="button" class="button" value="�����νı��³��� ��������" onClick="jsGetFinger();">
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
		<input type="hidden" name="hidSPday" value="<%= dSPayDay %>"><!-- �޿�������-->
		<input type="hidden" name="hidEPday" value="<%= dEPayDay %>"><!-- �޿�������--> 
		<input type="hidden" name="hidSPdate" value="<%= dSPayDate %>"><!-- �޿��������� �����-->
		<input type="hidden" name="hidEPdate" value="<%= dEPayDate %>"><!-- �޿��������� �����-->
		<input type="hidden" name="hidEday" value="<%=iLoopCnt%>"><!-- �޿��ϼ�--> 
		<input type="hidden" name="hidVer" value="<%=iVer%>">
		<input type="hidden" name="hidYear" value="<%=dYear%>"> 
		<input type="hidden" name="hidMonth" value="<%=dMonth%>"> 
		<input type="hidden" name="hidDP" value="<%=defaultpay%>"> 
		<tr bgcolor="<%= adminColor("gray") %>"  align="center">
			<td rowspan="2">��</td>
			<td rowspan="2">����</td>
			<td rowspan="2">����</td>
			<td colspan="2">�ٹ��ð�</td>
			<td colspan="2">�ް�ð�</td>
			<td rowspan="2">����<br>�ð�</td>
			<td rowspan="2">����<br>�ð�</td>
			<td rowspan="2">����<br>�ð�</td> 
			<td rowspan="2">�⺻�ٹ�<br>�ð�</td>
			<td rowspan="2">����ٹ�<br>�ð�</td>
			<td rowspan="2">�߰��ٹ�<br>�ð�</td>
			<td rowspan="2">���ϱٹ�<br>�ð�</td>
			<td rowspan="2">������<br>�ð�</td>
		</tr>
		<tr  bgcolor="<%= adminColor("gray") %>"  align="center">
			<td>����</td>
			<td>����</td>
			<td>����</td>
			<td>����</td>
		</tr>
		<%
		'// ========================================================================
		'// ������� ù �Ͽ��Ͽ� �ش��ϴ� ������ �ϼ� + �� ���� ������ ��ϰ�������
		'// 2013-02 �ΰ�� 2013-01-27 ������ ��� + �� ���� 1���ϸ��
		'// ������ ������ ����Ѵ�.
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
			'// ���� ����Ÿ�� �ִ� ��� ǥ��
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
					'����
					totPWD  		= totPWD + iWorkTime	'��ü �ٹ��ð�

					if (arrPre(5,intD) = "3") then
						blnPWH = blnPWH + 1
					end if

					IF arrPre(5,intD) = "1"  and iWorkTime = 0 THEN
						'�ٹ��Ͽ� �ٹ��� ��������� ���� ���޾ȵ�
						chkPWHD  =  chkPWHD  + 1
					END IF
				else
					'�̹��� ���޺κ�
					totCWD  		= totCWD + iWorkTime	'��ü �ٹ��ð�

					if (arrPre(5,intD) = "3") then
						blnCWH = blnCWH + 1
					end if

					IF arrPre(5,intD) = "1"  and iWorkTime = 0 THEN
						'�ٹ��Ͽ� �ٹ��� ��������� ���� ���޾ȵ�
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
						�ٹ���
					<%ELSEIF arrPre(5,intD)  = "2" THEN%>
						<font color="blue">��������<font>
					<%ELSEIF arrPre(5,intD)  = "3" THEN%>
						<font color="red">������</font>
					<%ELSEIF arrPre(5,intD)  = "4" THEN%>
						<font color="red">��������<font>
					<%ELSEIF arrPre(5,intD)  = "5" THEN%>
						<font color="red">������<font>
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
		<input type="hidden" name="hidPWD" value="<%=totPWD%>"><!--  ���� �� ������ �� �ٹ��ð�(2013-02 �� ��� 2013-01-27 ����) -->
		<input type="hidden" name="hidPWHD" value="<%=chkPWHD%>"><!--  - ���Ƚ��(2013-02 �� ��� 2013-01-27 ����) -->
		<input type="hidden" name="blnPWH" value="<%=blnPWH%>"><!--    - �����ϼ�(2013-02 �� ��� 2013-01-27 ����) -->
		<input type="hidden" name="hidCWD" value="<%=totCWD%>"><!--  �̹��� ���޺κ� �ٹ��ð�(2013-02 �� ��� 2013-01-27 ����) -->
		<input type="hidden" name="hidCWHD" value="<%=chkCWHD%>"><!--  - ���Ƚ��(2013-02 �� ��� 2013-01-27 ����) -->
		<input type="hidden" name="blnCWH" value="<%=blnCWH%>"><!--    - �����ϼ�(2013-02 �� ��� 2013-01-27 ����) -->
		<input type="hidden" name="hidPSD" value="<%=iminday%>"> 
		<input type="hidden" name="hidPED" value="<%=imaxday%>">
	<%'--------------- ���� �� �޿� ���� start ----------------
			totWorkTime = 0
			totextendWT  = 0
			totnightWT	=0
			totholidayWT=0
			totweekholidayWT=0
			totVacationTime = 0
		 i = 0  
	 
		 dFullPayDate = dSPaydate'�޿�������
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
		   
		
 			'������� ��������-------------------- 
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
 
			IF sFingerYN = "Y" THEN     '�����νı��³��� ��������
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
							'// �߰��ٹ�
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
							'// ���� = �����ް�, ������ �����ް�
							iworktype = "4"  '//�ް����� ������� �����ް�����(2014-08-08 ������ ����)
						 
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
						�ٹ���
					<%ELSEIF iworktype  = "2" THEN%>
						<font color="blue">��������<font>
					<%ELSEIF iworktype  = "3" THEN%>
						<font color="red">������</font>
					<%ELSEIF iworktype  = "4" THEN%>
						 		��������
					<%ELSEIF iworktype  = "5" THEN%>
						 		������
					<%ELSEIF iworktype  = "0" THEN%>
						 	<font color="Gray">�Ի���/�����</font>
					<%END IF%>
				<%ELSE%> 
				<select name="selWH<%=intD%>"  onChange="jsChangeWeekHoliday_Pre(<%=dWeekday%>,<%=intD%>);jsSetTotTime(<%=intD%>);">
				<option value="1" <%IF iworktype ="1"  THEN%>selected<%END IF%>>�ٹ���</option>
				<option value="2" <%IF iworktype ="2" THEN%>selected<%END IF%> style="color:blue">��������</option>
				<option value="3" <%IF iworktype ="3" THEN%>selected<%END IF%> style="color:red">������</option>
				<option value="4" <%IF iworktype ="4" THEN%>selected<%END IF%> style="color:red">��������</option>
				<option value="5" <%IF iworktype ="5" THEN%>selected<%END IF%> style="color:red">������</option>
				<option value="0" <%IF iworktype ="0" THEN%>selected<%END IF%>  style="color:gray">�Ի���/�����</option>
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
			<td colspan="13"> �߰����޼��� </td>
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
			<td colspan="9">�հ�</td> 
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
		<input type="submit" class="button" value="���">
	<%ELSEIF dstate =  "0" THEN%>
		<input type="submit" class="button" value="����">
	<%END IF%>
    </td>
</tr>
<tr>
    <td align="right"> <input type="button" value="����" onClick="jsSetTotTimeALL(<%=intD%>)"></td>
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
		