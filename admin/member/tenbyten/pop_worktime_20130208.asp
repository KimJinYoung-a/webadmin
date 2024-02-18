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
<%
'��������
Dim sempno, ino
Dim djoinday, susername, iposit_sn,sposit_name,blnstatediv,dretireday
Dim startdate,enddate ,defaultpay,foodpay,jobpay,inbreaktime,holidaywdtime
Dim intY, intM, intD, dYear, dMonth,dWeekday
Dim dEndDay,dNextDate
Dim clsPay, arrList
Dim dyyyymmdd, dstartHour, dstartMinute, dendHour, dendMinute, dbreakSHour,dbreakSMinute, dbreakEHour, dbreakEMinute,doutHour, dOutMinute,  iworktype, dstate, dStart, dEnd, dBreakS, dBreakE
Dim iWorkTime,iBreak, iextendWT ,inightWT,iholidayWT,iweekholidayWT,dNStart, dNEnd, dNBreakS, dNBreakE
Dim totWorkTime, totextendWT ,totnightWT,totholidayWT,totweekholidayWT
Dim  preStartDay, preEndDay,arrPre
Dim dSWD, dEWD,iWD, totWD,iWT,totWH, chkWHD
dim totPWD
Dim dcStartHour(8),dcStartMinute(8),dcEndHour(8),dcEndMinute(8),dcBreakSHour(8),dcBreakSMinute(8),dcBreakEHour(8),dcBreakEMinute(8) ,defaulttime(8), dcWorkType(8), intLoop
Dim arrWorkTime(31),arrWorkType(31)
Dim sFingerYN, ipart_sn
Dim ofingerprints,i,ircount,mstate

'�� �޾ƿ���
sempno= requestCheckvar(Request("sEN"),14)
ino= requestCheckvar(Request("ino"),10)
dYear = requestCheckvar(Request("selY"),4)
dMonth = requestCheckvar(Request("selM"),2)

sFingerYN = requestCheckvar(Request("sFYN"),1) '�����νı��°��� �����Դ��� ����

'�⺻�� ���� (���� ���)
IF dYear = "" THEN dYear  = Year(Date())
IF dMonth = "" THEN dMonth  = Month(Date())

dNextDate = dateadd("m",1, dateserial(dYear,dMonth,1))	'�˻������� 1��
dEndDay = day(dateadd("d",-1,dNextDate))	'�˻��� ������ ��¥ (������ 1�� - �Ϸ�)

'���� ������ ������, ������
preEndDay = dateadd("d", -1, dateserial(dYear,dMonth,1))
preStartDay = day(dateadd("d", -(weekday(preEndDay)-1),preEndDay))
preEndDay  =day(preEndDay)

'������ ��������
set clsPay = new CPay
	'--��� �⺻�������
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

	For intLoop = 1 To 7
		dcStartHour(intLoop) 		= format00(2,Fix(clsPay.FStartTime(intLoop)/60))
		dcStartMinute(intLoop)  	= format00(2,clsPay.FStartTime(intLoop) mod 60)
		dcEndHour(intLoop)       	= format00(2,Fix(clsPay.FEndTime(intLoop)/60))
		dcEndMinute(intLoop)       	= format00(2,clsPay.FEndTime(intLoop) )
		dcBreakSHour(intLoop)     	= format00(2,Fix(clsPay.FBreakSTime(intLoop)/60))
		dcBreakSMinute(intLoop)     = format00(2,clsPay.FBreakSTime(intLoop) mod 60)
		dcBreakEHour(intLoop)     	= format00(2,Fix(clsPay.FBreakETime(intLoop)/60))
		dcBreakEMinute(intLoop)     = format00(2,clsPay.FBreakETime(intLoop) mod 60)
		defaulttime(intLoop)		= clsPay.FdefaultTime(intLoop)
		dcWorkType(intLoop)			= clsPay.Fworktype(intLoop)
	Next

	clsPay.fnGetmonthlypayData
	mstate = clsPay.Fstate

	'--������ ������ �ٹ��ð�(������ ����� ����)
	arrPre =clsPay.fnGetPreDailypayData

	'--�˻��� �ٹ��ð� ����
	IF sFingerYN = "Y" THEN    '�����νĳ��� ������ ���
		set clsPay = nothing
		set ofingerprints = new cfingerprints_list
		ofingerprints.frectpart_sn = ipart_sn
		ofingerprints.frectempno = sempno
		ofingerprints.FrectSDate = dateserial(dYear,dMonth,1)
		ofingerprints.FrectEDate = dateadd("m",1,dateserial(dYear,dMonth,1))
		ofingerprints.ffingerprints_sum()
		 ircount = ofingerprints.FresultCount
		if ircount<=0 then
			set ofingerprints =nothing
			 Alert_return("�����νı��³����� �������� �ʽ��ϴ�. Ȯ�� �� �ٽ� �õ����ּ���")
			response.end
		END IF
	ELSE
		arrList = clsPay.fnGetDailypayData
		set clsPay = nothing
	END IF

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
IF datediff("m",startdate,dateserial(dYear,dMonth,1)) < 0 or datediff("m",enddate,dateserial(dYear,dMonth,1)) > 0 THEN
	dstate = 9
END IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<html>
<head>
<title>�ٹ��ð� ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/scm.css" type="text/css">
<script language="javascript" src="/js/jsPayCal.js"></script>
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
	 	document.frmSearch.submit();
	}

	function jsSubmitPay(){

		var swday = 0;
		var blnWH, itwt,arrValue, iValue;
		var startj,endj;

		//���۽ð��� ����ð����� �۾ƾ�. (�ٹ��ð�,�ް�ð�)  . �����Է� ����üũ
		for(i=1;i<=document.frmPay.hidEday.value;i++){
			if( parseInt(eval("document.frmPay.iSH"+i).value,10)*60+parseInt(eval("document.frmPay.iSM"+i).value,10) >  parseInt(eval("document.frmPay.iEH"+i).value,10)*60+parseInt(eval("document.frmPay.iEM"+i).value,10) ){
				alert("�ٹ����۽ð��� �ٹ�����ð����� ������մϴ�. �ٽ� �������ּ���");
				eval("document.frmPay.iSH"+i).focus();
				return false;
			}

			if( parseInt(eval("document.frmPay.iBSH"+i).value,10)*60+parseInt(eval("document.frmPay.iBSM"+i).value,10) >  parseInt(eval("document.frmPay.iBEH"+i).value,10)*60+parseInt(eval("document.frmPay.iBEM"+i).value,10) ){
				alert("�ٹ����۽ð��� �ٹ�����ð����� ������մϴ�. �ٽ� �������ּ���");
				eval("document.frmPay.iBSH"+i).focus();
				return false;
			}

			if(eval("document.frmPay.hidWD"+i).value  == 1 && swday == 0 ){
				swday =  i ;   //�� �޿�  ó�� �Ͽ��Ͽ� �ش��ϴ� ��¥
			}
		}

		//**1.������ ���ָ� �� �ٹ��� �� �� ù�� ������ ���� Ȯ��**//
		itwt = 0;
		blnWH = 0;
		var ipday1 =document.frmPay.hidPSD.value;
		if (ipday1 > 0) {
		 for(i=ipday1;i<(parseInt(ipday1,10)+7);i++){
		 	 arrValue = eval("document.frmPay.iPWT"+i).value.split(":");
			 iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
			  itwt = itwt + iValue;
		 }
		var preendday = document.frmPay.hidPED.value; //������ ������ ��
		var preendweekday = eval("document.frmPay.hidPWeD"+preendday).value;   //������ �������� ����
		 startj = document.frmPay.hidPED.value - (preendweekday-1); //������ ������ ���� ���� �� �Ͽ��Ͽ� �ش��ϴ� ��¥ ���ϱ�(��������-(������������-1))
		 for(i=startj;i<=document.frmPay.hidPED.value;i++){ //������ ������ ���� ���� �� �Ͽ��Ϻ��� ������ ������ ������ üũ
		 	 if (eval("document.frmPay.iPWH"+i).value=="3"){
		 	  	blnWH = blnWH + 1;
		 	  }
	     }
	 	}

	 	for(i=1;i<swday;i++){
		 	 if (eval("document.frmPay.selWH"+i).value=="3"){
		 	  	blnWH = blnWH + 1;
		 	  }
	     }

	     if(blnWH >1){
			  	alert("�������� �����Ͽ� �ѹ�  ���� �����մϴ�.");
			  	return false;
			}else  if(itwt >= 900 && blnWH < 1 && swday != 1){
				alert("�������� �������ּ���1");
			  	return false;
			}


	     //**2.ù��°�� �ٹ��� �� �� ������ ������ ���� Ȯ��**//
	     itwt = 0;
		 blnWH = 0;
		 for(i=1;i<swday;i++){
		 	 arrValue = eval("document.frmPay.iWT"+i).value.split(":");
			 iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
			  itwt = itwt + iValue;
		 }

		 itwt = itwt + document.frmPay.hidPWD.value;

		  for(i=swday;i<=swday+6;i++){
				 if (eval("document.frmPay.selWH"+i).value=="3"){
			  	blnWH = blnWH + 1;
			  }
			 }
		//if(itwt < 900 && blnWH > 0){
			  //	alert("�� �ٹ��ð��� 15�ð� �����ϋ� ������ ������ �Ұ����մϴ�.");
			  	//return false;
			//}else

			if(blnWH >1){
			  	alert("�������� �����Ͽ� �ѹ�  ���� �����մϴ�.");
			  	return false;
			}else  if(itwt >= 900 && blnWH < 1){
				alert("�������� �������ּ���2");
			  	return false;
			}


		 //**3.ù��°�� ����**//
		 swday = swday + 7;
		 //���� 15�ð� �̻� �ٹ��ϋ� �������� 1���Ͽ� �ѹ� �ִ°�?15�ð� �����ϋ� ������ ����ߵȴ�.

		 for(j=swday;j<=document.frmPay.hidEday.value;j++){
		 blnWH = 0;
		 endj = j+6;
		 itwt = 0;

		 //������ �ٹ����� ��
		 for(i=j-7;i<=endj-7;i++){
			 arrValue = eval("document.frmPay.iWT"+i).value.split(":");
				iValue = parseInt(arrValue[0],10)*60+parseInt(arrValue[1],10);
			  itwt = itwt + iValue;
			}
		//���� �� ������ �������� Ȯ��
		 if(endj > document.frmPay.hidEday.value){endj = document.frmPay.hidEday.value}
			for(i=j;i<=endj;i++){
				 if (eval("document.frmPay.selWH"+i).value=="3"){
			  	blnWH = blnWH + 1;
			  }
			 }

	 		//if(itwt < 900 && blnWH > 0){
			//  	alert("�� �ٹ��ð��� 15�ð� �����ϋ� ������ ������ �Ұ����մϴ�.");
			  //	return false;
			//}else
			if( blnWH >1){
			  	alert("�������� �����Ͽ� �ѹ�  ���� �����մϴ�.");
			  	return false;
			}else  if(itwt >= 900 && blnWH < 1 && eval("document.frmPay.hidWD"+endj).value == 7){
				alert("�������� �������ּ���3");
			  	return false;
			}
		  	j = i-1;

		}

		jsAddWH();
		jsSetTotTimeALL(parseInt(document.frmPay.hidEday.value)+1); //����� ����ó��
		  return true;
	}

	function jsAddWH(){
	 //����� ������ ����
		<%IF blnstatediv = "N" THEN
			IF  Cstr(month(dretireday))  = Cstr(dMonth) THEN	'������̰� ������ ���%>
			if ((<%=day(dretireday)%>+(7-<%=weekday(dretireday)%>)) >= document.frmPay.hidEday.value){ //����� ���������� ����Ϻ��� ũ�ų� ���� ���
		var iLwt = 0;
		var	iValue = 0;
		var	iEValue = 0;
		var	arrValue = "";
		var chkWHD  = 0;
		var iLWHD = 0;
		var iLDutyH,iLWHTime;
		var iLSday = "<%=day(dretireday)-weekday(dretireday)+1%>";
		var iLEday = "<%=day(dretireday)%>";
			iLWHTime = 0;

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
					if( eval("document.frmPay.selWH"+i).value=="3"){iLWHD=i};

					iLwt = iLwt + iValue+iEValue;
				 }

				 if(chkWHD==0){
					iLDutyH = parseInt(iLwt/60,10);
					if (iLDutyH > 40){iLDutyH=40};
					if(iLDutyH>=15){iLWHTime= (8*(iLDutyH/40))*60;}
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
					if (iLDutyH > 40){iLDutyH=40};
					if(iLDutyH>=15){iLWHTime= iLWHTime+ (8*(iLDutyH/40))*60;}
				 }

		 		 document.frmPay.iwhWT32.value = jsTimeForm(iLWHTime);
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

	function jsSetTotTimeALL(ilen){
	    for(var j=1;j<=ilen-1;j++){
	        jsSetTotTime(j);
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
<input type="hidden" name="sFYN" value="N">
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
			<td align="right"><%IF mstate = 0 THEN%><input type="button" class="button" value="�����νı��³��� ��������" onClick="jsGetFinger();"><%END IF%>
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
		<input type="hidden" name="hidEday" value="<%=dEndDay%>">
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
			<td rowspan="2">�⺻�ٹ�<br>�ð�</td>
			<td rowspan="2">����ٹ�<br>�ð�</td>
			<td rowspan="2">������<br>�ð�</td>
			<td rowspan="2">�߰��ٹ�<br>�ð�</td>
			<td rowspan="2">���ϱٹ�<br>�ð�</td>

		</tr>
		<tr  bgcolor="<%= adminColor("gray") %>"  align="center">
			<td>����</td>
			<td>����</td>
			<td>����</td>
			<td>����</td>
		</tr>
		<% '== �˻� ���� �� ������ ������ ���� ��������(������ ����� ����)
		dim chkPWHD, imaxday, iminday, imaxD
		totPWD = 0
		chkPWHD = 0
		imaxday = 0
		iminday = 0
		imaxD = 0
		IF isArray(arrPre) THEN
			imaxD = UBound(arrPre,2)
			iminday = day(arrPre(0,0))
			imaxday = right(arrPre(0,UBound(arrPre,2)),2)
			if imaxday = 32 then imaxD = UBound(arrPre,2)-1
			For intD = 0 To imaxD
			iWorkTime = 0
			iextendWT  = 0
			inightWT	=0
			iholidayWT=0
			iweekholidayWT=0
			IF weekday(arrPre(0,intD)) = 1 then  totPWD = 0
			iWorkTime = arrPre(7,intD)
			totPWD  	= totPWD  + iWorkTime	'��ü �ٹ��ð�
			iextendWT = arrPre(8,intD)
			inightWT	= arrPre(9,intD)
			iholidayWT= arrPre(10,intD)
			iweekholidayWT=arrPre(11,intD)

			IF arrPre(5,intD) = "1"  and iWorkTime = 0 THEN '�ٹ��Ͽ� �ٹ��� ��������� ���� ���޾ȵ�
			 chkPWHD  =  chkPWHD  + 1
			END IF

			'���ϱٹ�
			IF iworktype =  "3" THEN
				 IF iWorkTime > 0 THEN
					iholidayWT = iWorkTime
					iWorkTime = 0
				END IF
			END IF

		%>
		<tr   bgcolor="#DFDFDF" align="center">
			<td><%=day(arrPre(0,intD))%></td>
			<td><%=fnGetStringWD(weekday(arrPre(0,intD)))%><input type="hidden" name="hidPWeD<%=day(arrPre(0,intD))%>" value="<%=weekday(arrPre(0,intD))%>"></td>
			<td>
				<%IF arrPre(5,intD)  = "1" THEN%>
						�ٹ���
					<%ELSEIF arrPre(5,intD)  = "2" THEN%>
						<font color="blue">��������<font>
					<%ELSEIF arrPre(5,intD)  = "3" THEN%>
						<font color="red">������</font>
					<%ELSEIF arrPre(5,intD)  = "4" THEN%>
						 ��������
					<%ELSEIF arrPre(5,intD)  = "5" THEN%>
						 ������
					<%END IF%>
					<input type="hidden" name="iPWH<%=day(arrPre(0,intD))%>" value="<%=arrPre(5,intD)%>">
			</td>
			<td><input type="text" name="iPWS<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(1,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF " readonly></td>
			<td><input type="text" name="iPWE<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(2,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
			<td><input type="text" name="iPBS<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(3,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
			<td><input type="text" name="iPBE<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(4,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
			<td><input type="text" name="iPO<%=day(arrPre(0,intD))%>" value="<%=fnSetTimeFormat(arrPre(12,intD))%>" size="5" maxlength="5" style="border:0;background:#DFDFDF" readonly></td>
			<td></td>
			<td><input type="text" name="iPWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iWorkTime)%>"></td>
			<td><input type="text" name="iPeWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iextendWT)%>"></td>
			<td><input type="text" name="iPwhWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iweekholidayWT)%>"></td>
			<td><input type="text" name="iPnWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(inightWT)%>"></td>
			<td><input type="text" name="iPhWT<%=day(arrPre(0,intD))%>" style="border:0;background:#DFDFDF" readonly size="5" value="<%=fnSetTimeFormat(iholidayWT)%>"></td>
		</tr>
		<%	Next
		END IF
		%>
		<input type="hidden" name="hidPWD" value="<%=totPWD%>">
		<input type="hidden" name="hidPSD" value="<%=iminday%>">
		<input type="hidden" name="hidPED" value="<%=preEndDay%>">
		<%
			totWorkTime = 0
			totextendWT  = 0
			totnightWT	=0
			totholidayWT=0
			totweekholidayWT=0

		 i = 0

		For intD = 1 To dEndDay
			iworktype = ""
			iWorkTime = 0
			iextendWT  = 0
			inightWT	=0
			iholidayWT=0
			iweekholidayWT=0

			dWeekday = weekday(dateserial(dYear,dMonth,intD))

			'������� ��������--------------------
			IF djoinday  > dateserial(dYear,dMonth,intD)  THEN
				dbreakSHour = "00"
				dbreakSMinute ="00"
				dbreakEHour = "00"
				dbreakEMinute ="00"
	 			iworktype = 0
			ELSE
				dbreakSHour = dcbreakSHour(dWeekday)
				dbreakSMinute = dcbreakSMinute(dWeekday)
				dbreakEHour = dcbreakEHour(dWeekday)
				dbreakEMinute = dcbreakEMinute(dWeekday)
	 			iworktype = dcWorkType(dWeekday)
 			END IF
 			'---------------------------------------

			IF sFingerYN = "Y" THEN     '�����νı��³��� ��������
			 	dstartHour 	= "00"
				dstartMinute= "00"
				dendHour 	= "00"
				dendMinute 	= "00"
				doutHour 	= "00"
				doutMinute 	= "00"
				arrWorkTime(intD) = 0
				arrWorkType(intD) = iworktype

				 if i < ircount then
					dyyyymmdd	= ofingerprints.FItemList(i).fyyyymmdd
					if day(dyyyymmdd) =  intD then
						dstartHour	= format00(2,hour(ofingerprints.FItemList(i).fInTime))
						dstartMinute= format00(2,minute(ofingerprints.FItemList(i).fInTime))
						if ofingerprints.FItemList(i).fOutTime <> "1900-01-01" then
						dendHour	= format00(2,hour(ofingerprints.FItemList(i).fOutTime))
						dendMinute	= format00(2,minute(ofingerprints.FItemList(i).fOutTime))
						end if

						doutHour	= format00(2,Fix(ofingerprints.FItemList(i).fexmin/60))
						doutMinute	= format00(2, ofingerprints.FItemList(i).fexmin mod 60)

						iWorkTime	= ofingerprints.FItemList(i).fworkmin
						ibreak = (dbreakEHour*60+dbreakEMinute)-(dbreakSHour*60+ dbreakSMinute)

'						 //�ް�ð� ���Կ���
'						 if not inBreakTime THEN iWorkTime = iWorkTime - ibreak
'						arrWorkTime(intD) =  iWorkTime '�ش糯¥�� �ش��ϴ� �迭��ġ�� �ٹ��ð� �ֱ�
'
'						if iWorkTime > 480 THEN
'							iextendWT = iWorkTime - 480
'							iWorkTime = 480
'						END IF
'
'						Dim nightS, nightE,nightBS,nightBE
'
'						//�߰��ٹ�����
'						if (dendHour*60+ dendMinute)>22*60  then
'							if  (dstartHour*60+ dstartMinute) < 22*60  then
'								nightS = 22*60
'							else
'								nightS = dstartHour*60+ dstartMinute
'							end if
'
'							if (dendHour*60+ dendMinute) > 30*60  then
'								nightE = 30*60
'							else
'								nightE = dendHour*60+ dendMinute
'							end if
'
'						  	if  (dbreakSHour*60+ dbreakSMinute) < 22*60  then
'								nightBS = 22*60
'							elseif (dbreakSHour*60+ dbreakSMinute) >=30*60 then
'								nightBS = 0
'							else
'								nightBS = dbreakSHour*60+ dbreakSMinute
'						 	end if
'
'							if  (dbreakEHour*60+dbreakEMinute) < 22*60  then
'								nightBE = 22*60
'							elseif (dbreakEHour*60+ dbreakEMinute) >30*60 then
'								nightBE = 0
'							else
'								nightBE = dbreakEHour*60+dbreakEMinute
'							end if
'
'							if inBreakTime="0"  then
'								inightWT = nightE- nightS- (nightBE-nightBS)
'						    else
'								inightWT = nightE- nightS
'							end if
'						 end if
'
'
'						 //���ϱٹ�����
'						 IF iworktype =  "3" or iworktype ="5" THEN
'							 IF iWorkTime > 0 THEN
'								iholidayWT = iWorkTime
'								iWorkTime = 0
'							END IF
'						END IF
'
					 i = i + 1
					end if
'
'					 //���޼���
'					Dim iSwday, iEwday, iD , totDuty, chkwd, totDutyH
'					 totDuty = 0
'					 totDutyH = 0
'					 chkwd = 0
'					 IF iworktype  = 3 Then
'					 	iSwday = intD -dWeekday-6
'					 	iEwday	= iSwday + 6
'
'					 	IF iSWday< 0 THEN iSWday=1
'	 	 				IF iEWday > dEndDay THEN iEWday = dEndDay
'
'	 	 				FOR iD = iSWday To iEWday
'	 						totDuty = totDuty + arrWorkTime(iD)
'	 						IF iD = 1 THEN totDuty = totDuty + totPWD
'
'	 						IF ((arrWorkType(iD) = 1 and   arrWorkTime(iD)<= 0 ) or arrWorkTime(iD) ="" ) THEN
'	 							chkwd = 1
'	 							Exit For
'	 						END IF
'	 					NEXT
'
'	 					totDutyH =  Cint(totDuty/60)
'	 					IF totDutyH > 40 THEN totDutyH = 40
'	 					IF totDutyH > 15 and chkwd = 0 THEN
'	 						iweekholidayWT = 8*(totDutyH/40)*60
'	 					END IF
'
' 					END IF
 				end if
			ELSE
				IF isArray(arrList) THEN
				dyyyymmdd	= arrList(0,intD-1)
				dStart		= arrList(1,intD-1)
				dEnd		= arrList(2,intD-1)
				dBreakS		= arrList(3,intD-1)
				dBreakE		= arrList(4,intD-1)
				dstartHour	= format00(2,Fix(dStart/60))
				dstartMinute= format00(2,dStart mod 60)
				dendHour	= format00(2,Fix(dEnd/60))
				dendMinute	= format00(2,dEnd mod 60)
				dbreakSHour	= format00(2,Fix(dBreakS/60))
				dbreakSMinute= format00(2,dBreakS mod 60)
				dbreakEHour	= format00(2,Fix(dBreakE/60))
				dbreakEMinute= format00(2,dBreakE mod 60)
				doutHour	= format00(2,Fix(arrList(12,intD-1)/60))
				doutMinute	= format00(2, arrList(12,intD-1) mod 60)
				iworktype	= arrList(5,intD-1)
				dstate		= arrList(6,intD-1)

				iWorkTime	= arrList(7,intD-1)
				iextendWT	= arrList(8,intD-1)
				inightWT	= arrList(9,intD-1)
				iholidayWT	= arrList(10,intD-1)
				iweekholidayWT= arrList(11,intD-1)
				END IF
			END IF

		  totWorkTime 	= totWorkTime + iWorkTime
			totextendWT  	= totextendWT + iextendWT
			totnightWT		= totnightWT + inightWT
			totholidayWT	= totholidayWT + iholidayWT
			totweekholidayWT= totweekholidayWT  + iweekholidayWT

		%>
		<tr   bgcolor="#FFFFFF"  align="center" >
			<td><%=intD%></td>
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
						 	<font color="Gray">�Ի���</font>
					<%END IF%>
				<%ELSE%>
				<select name="selWH<%=intD%>"  onChange="jsChangeWeekHoliday_Pre(<%=dWeekday%>,<%=intD%>);jsSetTotTime(<%=intD%>);">
				<option value="1" <%IF iworktype ="1"  THEN%>selected<%END IF%>>�ٹ���</option>
				<option value="2" <%IF iworktype ="2" THEN%>selected<%END IF%> style="color:blue">��������</option>
				<option value="3" <%IF iworktype ="3" THEN%>selected<%END IF%> style="color:red">������</option>
				<option value="4" <%IF iworktype ="4" THEN%>selected<%END IF%>>��������</option>
				<option value="5" <%IF iworktype ="5" THEN%>selected<%END IF%>>������</option>
				<option value="0" <%IF iworktype ="0" THEN%>selected<%END IF%>  style="color:gray">�Ի���</option>
				</select>
				<%END IF%>
			</td>
			<td>
				<input type="text" name="iSH<%=intD%>" value="<%=dstartHour%>" size="2" maxlength="2" style="text-align:right;<%IF dstate > 0 THEN%> border:0;" readonly<%ELSE%>"<%END IF%>  onKeyUp="jsSetTotTime(<%=intD%>);TnTabNumber('iSH<%=intD%>','iSM<%=intD%>',2);">
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
			<td><input type="text" name="iWT<%=intD%>" style="border:0;color:<%IF iWorkTime  = 0  THEN %>gray<%ELSEIF  iWorkTime < 0 THEN%>red<%ELSE%>blue<%END IF%>;" readonly size="5" value="<%=fnSetTimeFormat(iWorkTime)%>"></td>
			<td><input type="text" name="ieWT<%=intD%>" style="border:0;color:<%IF iextendWT  = 0  THEN %>gray<%ELSEIF  iextendWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(iextendWT)%>"></td>
			<td><input type="text" name="iwhWT<%=intD%>" style="border:0;color:<%IF iweekholidayWT  = 0  THEN %>gray<%ELSEIF  iweekholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(iweekholidayWT)%>"></td>
			<td><input type="text" name="inWT<%=intD%>" style="border:0;color:<%IF inightWT  = 0  THEN %>gray<%ELSEIF  inightWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(inightWT)%>"></td>
			<td><input type="text" name="ihWT<%=intD%>" style="border:0;color:<%IF iholidayWT  = 0  THEN %>gray<%ELSEIF  iholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(iholidayWT)%>"></td>

		</tr>
		<%Next
		set ofingerprints = nothing
		%>
		<%IF isArray(arrList) THEN
			if right(arrList(0,ubound(arrList,2)),2) = "32" THEN
				totweekholidayWT = totweekholidayWT + arrList(11,ubound(arrList,2))
			%>
		<tr   bgcolor="#FFFFFF"  align="center">
			<td colspan="10"> �߰����޼��� </td>
			<td><div id="dNMWT" style="display:;"><input type="text" name="iwhWT32" style="border:0;color:blue"  size="5" value="<%=fnSetTimeFormat(arrList(11,ubound(arrList,2)))%>"></div></td>
			<td colspan="3"></td>
		</tr>
		<%ELSE%>
		<tr>
			<td><div id="dNMWT" style="display:none;"><input type="text" name="iwhWT32" value="0"></div></td>
		</tr>
		<% end if
		END IF%>
		<tr   bgcolor="<%=adminColor("sky")%>" align="center">
			<td colspan="9">�հ�</td>
			<td><input type="text" name="totWT" style="border:0;background:#DFDFDF;color:<%IF totWorkTime  = 0  THEN %>gray<%ELSEIF  totWorkTime < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totWorkTime)%>"></td>
			<td><input type="text" name="toteWT" style="border:0;background:#DFDFDF;color:<%IF totextendWT  = 0  THEN %>gray<%ELSEIF  totextendWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totextendWT)%>"></td>
			<td><input type="text" name="totwhWT" style="border:0;background:#DFDFDF;color:<%IF totweekholidayWT  = 0  THEN %>gray<%ELSEIF  totweekholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totweekholidayWT)%>"></td>
			<td><input type="text" name="totnWT" style="border:0;background:#DFDFDF;color:<%IF totnightWT  = 0  THEN %>gray<%ELSEIF  totnightWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totnightWT)%>"></td>
			<td><input type="text" name="tothWT" style="border:0;background:#DFDFDF;color:<%IF totholidayWT  = 0  THEN %>gray<%ELSEIF  totholidayWT < 0 THEN%>red<%ELSE%>blue<%END IF%>" readonly size="5" value="<%=fnSetTimeFormat(totholidayWT)%>"></td>
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
    <td align="right"> <input type="button" value="����" onClick="jsSetTotTimeALL(<%= intD%>)"></td>
</tr>
</form>
</table>
</body>
</html>
<%IF sFingerYN = "Y" THEN %>
<script language="javascript">
	var chk = 0;
	window.onload = function()
	{
		if(chk==0){
		jsSetTotTimeALL(<%= intD%>);
		chk = 1;
	}
	}
</script>
<%END IF%>