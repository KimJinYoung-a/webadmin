<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ñް���� ��� �޿� ����
' History : 2020.06.19 ������  ����
'###########################################################
%>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenPayCls.asp" -->
<%
Dim intY, intM, dYear, dMonth
Dim sEmpno, sUsername, dJoinday, blnstatediv,iposit_sn,sposit_name,dretireday
dim holidaywdtime,ino,startdate,enddate,defaultpay,foodpay,jobpay,inbreaktime,iDefaultPaySeq,predefaultpay,prefoodpay
Dim iworktime,iextendtime,inighttime,iholidaytime ,mtimepay,mextendpay ,mnightpay ,mholidaypay
dim mwholidaypay,mfoodpay,mjobpay ,mlongtimepay ,maddpay, myearpay, mbonuspay
dim mworkday
Dim moutstandingpay,mtotpay,mnpensionpay,mhealthinspay,mrecuinspay,munempinspay,mtaxtotpay,mrealtotpay,dregdate,sadminid,istate
Dim clsPay
Dim arrList , arrPre, intLoop
dim totDutyTime, totNightTime,totPaySum,avgWeek,iOverTime
dim arrdtime, idtime,dSWD,dEWD,totWD,iWD,dNStart,dNEnd,dNBreakS,dNBreakE, iweekholidaytime,totPWD, dweekday,totWH,iWT
dim dNextDate ,dEndDay,dREday,chkWHD,dEndDate, blnReset
Dim totWorkDay, totWorkDayReal
dim monthlyPayDataExist
dim iReworktime,iReextendtime,iRenighttime,iReholidaytime,iReweekholidaytime
dim iRefoodtime,mReExtimepay,mReNTtimepay,mReHDtimepay,mReFtimepay
dim mretimepay,mreextendpay,mrenightpay,mreholidaypay, mrefoodpay,mretotpay
dim intP, iP,totReWorkDay ,totReWorkDayReal, strSql
dim iTReworktime,iTReextendtime,iTRenighttime,iTReholidaytime,iTReweekholidaytime ,ireworkday

sEmpno	= requestCheckvar(request("sEN"),14)	'���
dYear	= requestCheckvar(request("selY"),4)	'��
dMonth	= requestCheckvar(request("selM"),2)	'��
ino		= requestCheckvar(request("ino"),10)	'ȸ��
blnReset = requestCheckvar(request("blnR"),1) '���¿���
'�⺻�� ���� (���� ���)
IF dYear = "" THEN dYear  = Year(Date())
IF dMonth = "" THEN dMonth  = Month(Date())

if session("ssBctSn")<>sEmpno then 
	response.write "<script>self.close();</script>"
	response.end
end if

'// 4.345238095 == �� ��� WEEK �� = (365�� / 7�� / 12����)
avgWeek = 4.345238095

totWorkDay = 0
totWorkDayReal = 0
totReWorkDay = 0
totReWorkDayReal =0

if ino="" then'ȸ�� ���� ��������
	strSql ="select max(ino) as ino from [db_partner].[dbo].tbl_user_monthlypay where empno='"&sEmpno&"'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly
	IF Not (rsget.EOF OR rsget.BOF) THEN
		ino = rsget(0)
	END IF
	rsget.close
end if

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

	//����Ʈ
	function jsPrint(){
		var winPrint = window.open("print_worktime.asp?sEN=<%=sempno%>&ino=<%=ino%>&selY=<%=dYear%>&selM=<%=dMonth%>","prtWT","width=1020,height=600,scrollbars=yes,resizable=yes");
		winPrint.focus();
	}

 	function jsWorkTime(empno,ino){
 		var wwt =window.open("pop_worktimeview.asp?sEN="+empno+"&ino="+ino+"&selY=<%=dYear%>&selM=<%=dMonth%>","popWT","width=1200,height=800,scrollbars=yes,resizable=yes");
		wwt.focus();
 	}
//-->
</script>
<table width="100%"  cellpadding="3" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">���</td>
			<td bgcolor="#FFFFFF" width="180"><%=sempno%></td>
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
		<input type="button" class="button" value="�Ϻ� �ٹ��ð�" onClick="jsWorkTime('<%=sempno%>','<%=ino%>');">
		<input type="button" value="����Ʈ" class="button" onClick="jsPrint();">
	</td>
</tr>
</form>
<tr>
	<td><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table border="0" width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
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
				<td><%=formatnumber(mtimepay,0)%></td>
				<td><%=formatnumber(mextendpay,0)%></td>
				<td><%=formatnumber(mnightpay,0)%></td>
				<td><%=formatnumber(mholidaypay,0)%></td>
				<td><%=formatnumber(mfoodpay,0)%></td>
				<td><%=formatnumber(mjobpay,0)%></td>
				<td><%=formatnumber(moutstandingpay,0)%></td>
				<td><%=formatnumber(mlongtimepay,0)%></td>
				<td><%=formatnumber(maddpay,0)%></td>
				<td><%=formatnumber(myearpay,0)%></td>
				<td><%=formatnumber(mbonuspay,0)%></td>			
				<td><%=formatnumber(mtotpay,0)%></td>
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
						<% end if %>
					<% else %>
						<%= totWorkDay %>
					<% end if %>
				</td>
			</tr> 
			<%ELSE%>
			<tr  bgcolor="<%=adminColor("sky")%>" align="center">
				<td bgcolor="<%= adminColor("sky") %>"><b>�ѱݾ�</b></td>
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
				<td><%=formatnumber(mtimepay,0)%></td>
				<td><%=formatnumber(mextendpay,0)%></td>
				<td><%=formatnumber(mnightpay,0)%></td>
				<td><%=formatnumber(mholidaypay,0)%></td>
				<td><%=formatnumber(mfoodpay,0)%></td>
				<td><%=formatnumber(mjobpay,0)%></td>
				<td><%=formatnumber(moutstandingpay,0)%></td>
				<td><%=formatnumber(mlongtimepay,0)%></td>
				<td><%=formatnumber(maddpay,0)%></td>
				<td><%=formatnumber(myearpay,0)%></td>
				<td><%=formatnumber(mbonuspay,0)%></td>			
				<td><%=formatnumber(mtotpay,0)%></td>
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
						<% end if %>
					<% else %>
						<%= totWorkDay %>
					<% end if %>
				</td>
			</tr>
			<tr  bgcolor="#e3f1fb" align="center">
				<td bgcolor="#e3f1fb" nowrap>�������ױ�</td>
				<td><%=formatnumber(mretimepay,0)%></td>
				<td><%=formatnumber(mreextendpay,0)%></td>
				<td><%=formatnumber(mrenightpay,0)%></td>
				<td><%=formatnumber(mreholidaypay,0)%></td>
				<td><%=formatnumber(mrefoodpay,0)%></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>
				<td></td>			
				<td><%=formatnumber(mretotpay,0)%></td>
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
				<td><%=formatnumber(mtimepay,0)%></td>
				<td><%=formatnumber(mextendpay,0)%></td>
				<td><%=formatnumber(mnightpay,0)%></td>
				<td><%=formatnumber(mholidaypay,0)%></td>
				<td><%=formatnumber(mfoodpay,0)%></td>
				<td><%=formatnumber(mjobpay,0)%></td>
				<td><%=formatnumber(moutstandingpay,0)%></td>
				<td><%=formatnumber(mlongtimepay,0)%></td>
				<td><%=formatnumber(maddpay,0)%></td>
				<td><%=formatnumber(myearpay,0)%></td>
				<td><%=formatnumber(mbonuspay,0)%></td>			
				<td><%=formatnumber(mtotpay,0)%></td>
			</tr>	
		<%END IF%>
		</table>
	</td>
</tr>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
