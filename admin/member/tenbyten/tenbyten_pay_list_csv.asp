<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����������޿�
' History : 2011.09.07 ������ ����
'			2011.12.16 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenPayCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
if (isNull(session("ssAdminPOSITsn")) or session("ssAdminPOSITsn")="" or isNull(session("ssAdminPOsn")) or session("ssAdminPOsn")="" or isNull(session("ssAdminLsn")) or session("ssAdminLsn")="" or isNull(session("ssAdminPsn")) or session("ssAdminPsn")="") then
	response.write  "������ �����ϴ�. - �α׾ƿ� �� �ٽ� �α������ּ���. "
	dbget.close() : response.end
end if

if (session("ssAdminLsn")=0) then
	dim bufbuf : bufbuf = 1/0  ''raize Error
	dbget.close() : response.end
end if

if (Not (C_ADMIN_AUTH or C_MngPart or C_ManagerUpJob or C_PSMngPart)) then
    response.write  "������ �����ϴ�. - �ý����� ���� " ''eastone
    dbget.close() : response.end
end if

'// 2015-06-22, skyer9
''if Not C_ManagerPartTimeMember then
''	response.write  "������ �����ϴ�. - �ý����� ���� "
''	dbget.close() : response.end
''end if

'if (session("ssAdminPsn") = "10") and (session("ssBctId") <> "bseo") and (session("ssBctId") <> "boyishP") then
'	'// CS����� ��û����, 2015-04-08 2017-10-23 ������ ����
'	response.write  "������ �����ϴ�. - �ý����� ���� " ''eastone
'	dbget.close() : response.end
'end if

Dim page, SearchKey, SearchString, isUsing, part_sn, research, orderby, selState
Dim job_sn, posit_sn, intY, intM, sYear, sMonth ,shopid
Dim iTotCnt,iPageSize, iTotalPage, totsum,totReCalSum
dim totBasePay, totOverTimePay, totNightTimePay, totHolidayPay, totFoodPay, totPositionPay, totBestPay, totLongWorkPay, totAddPay, totYearPay, totBonusPay, totWorkTime
dim department_id, inc_subdepartment
dim tmpM
	iPageSize	  = request("pagesize")
	if (iPageSize = "") then
		iPageSize = 50
	end if

	page = requestCheckvar(Request("page"),10)
	isUsing = requestCheckvar(Request("isUsing"),1)
	SearchKey = requestCheckvar(Request("SearchKey"),1)
	SearchString = requestCheckvar(Request("SearchString"),32)
	part_sn = requestCheckvar(Request("part_sn"),10)
	job_sn = requestCheckvar(Request("job_sn"),10)
	posit_sn = requestCheckvar(Request("posit_sn"),10)
	sYear = requestCheckvar(Request("sel_DY"),4)
	sMonth = requestCheckvar(Request("sel_DM"),2)
	research = requestCheckvar(Request("research"),2)

	orderby = requestCheckvar(Request("orderby"),1)
	selState = requestCheckvar(Request("selState"),4)

	department_id = requestCheckvar(Request("department_id"),10)
	inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)

if isUsing="" and research="" then isUsing="Y"
IF orderby = "" then orderby  = 1
if page="" then page=1
IF sYear = "" and research="" THEN sYear = year(date())
IF sMonth = "" and research="" THEN sMonth = month(date())

'// �α�������(���)�� ���� �⺻ �μ� ����(������ �̻�:2 �� �ý�����:7 ����)
'SCM �޴����� �������� �����Ѵ�.
if Not (session("ssAdminLsn")<=2  or session("ssAdminPsn")=7  or session("ssAdminPsn")= 8  or session("ssAdminPsn")= 20 )  then
    if (part_sn="") then
	    ''part_sn = session("ssAdminPsn")
	else
	    ''part_sn = checkValidPart(session("ssBctId"),part_sn)   '' if inValid return -999
    end if

	if (department_id = "") then
		department_id = GetUserDepartmentID("",session("ssBctID"))
	end if
end if

'����/������
if (C_IS_SHOP) then

	'/�������� �����ڰ� �ƴ�
	if not(C_OFF_AUTH) then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			shopid = C_STREETSHOPID
		'end if
	end if
end if
'-------------------------------------------------------------------------------
dim dSPayDate,dEPayDate,dPreEndDate,sPreYear,sPreMonth,dNextDate,dEndDay
dim searchdate
	searchdate ="N"
IF sYear <> ""   then
	if sMonth <> "" then
		searchdate ="Y"
		dPreEndDate = dateadd("d", -1, dateserial(sYear,sMonth,1)) '������ ������ ��¥
		sPreYear = year(dPreEndDate) '������
		sPreMonth = month(dPreEndDate) '������
		dNextDate = dateadd("m",1, dateserial(sYear,sMonth,1)) '������
		dEndDay = day(dateadd("d",-1,dNextDate)) '�̹��� ��������¥

		IF  sYear&"-"&format00(2,sMonth)  = "2014-01" THEN
			dSPayDate = dateserial(sYear,sMonth,1) '�޿�������: �ش�� 1�Ϻ���
			dEPayDate =dateserial(sYear,sMonth,25)
		ELSEIF sYear&"-"&format00(2,sMonth) > "2014-01" and sYear&"-"&format00(2,sMonth) < "2017-01" THEN
			dSPayDate = dateserial(sPreYear,sPreMonth,26) '�޿�������: ������ 26�Ϻ���
			dEPayDate = dateserial(sYear,sMonth,25) '�޿�������: �ش�� 25�ϱ���
		ELSE
			dSPayDate = dateserial(sYear,sMonth,1) '�޿�������: �ش�� 1�Ϻ���
			dEPayDate = dateserial(sYear,sMonth,dEndDay)  '�޿�������: �ش�� ���ϱ���
		END IF
	 else

	 	 IF  sYear   = "2014" THEN
			dSPayDate = dateserial(sYear,1,1) '�޿�������: �ش�� 1�Ϻ���
			dEPayDate =dateserial(sYear,12,25)
		ELSEIF sYear > "2014" and sYear < "2017" THEN
			dSPayDate = dateserial(sYear,1,1) '�޿�������: ������ 26�Ϻ���
			dEPayDate = dateserial(sYear,12,25) '�޿�������: �ش�� 25�ϱ���
		ELSE
			dSPayDate = dateserial(sYear,1,1) '�޿�������: �ش�� 1�Ϻ���
			dEPayDate = dateserial(sYear,12,31)  '�޿�������: �ش�� ���ϱ���
		END IF
	end if
ELSE
	dSPayDate = ""
	dEPayDate =""
END IF
'-------------------------------------------------------------------------------
'// ����Ʈ
dim clsPay, arrList,intLoop
Set clsPay = new CPay
	clsPay.FPageSize 	= iPageSize
	clsPay.FCurrPage 	= page
	clsPay.FSYYYYMM		= dSPayDate
	clsPay.FEYYYYMM		= dEPayDate
	clsPay.FSearchType 	= searchKey
	clsPay.FSearchText 	= searchString
	clsPay.Fstatediv 	= isUsing
	clsPay.Fpart_sn 	= part_sn
	clsPay.Fjob_sn 		= job_sn
	clsPay.Fposit_sn 	= posit_sn
	clsPay.Forderby 	= orderby
	clsPay.Fstate 		= selState
	clsPay.FIsMonth		=  1
	clsPay.fshopid		= shopid
	clsPay.FSearchDate  = searchdate
	clsPay.Fdepartment_id 		= department_id
	clsPay.Finc_subdepartment 	= inc_subdepartment

	arrList = clsPay.fnGetMonthlypayListCSV

set clsPay = nothing



totsum = 0
totBasePay = 0
totOverTimePay = 0
totNightTimePay = 0
totHolidayPay = 0
totFoodPay = 0
totPositionPay = 0
totBestPay = 0
totLongWorkPay = 0
totAddPay = 0
totYearPay = 0
totBonusPay = 0
totWorkTime = 0
totReCalSum = 0
dim sTitle,bufStr
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=����������޿�����.csv"
Response.CacheControl = "public"


sTitle = "��¥,���,�̸�,�μ�,��ǥ����(���������),��౸��,�ð���,�⺻��,�ð��ܼ���,�߰��ٹ�����,���ϱٹ�����,�Ĵ�����,��å����,����������,���ټӼ���,�߰�����,��������,�󿩱�,�ѱٹ��ð�,�հ�,����"
if C_MngPart or C_ADMIN_AUTH or C_PSMngPart then
sTitle = sTitle&",�������ױ� "
end if
response.write sTitle& VbCrlf
bufStr = ""
	if isArray(arrList) then
	  for intLoop=0 to ubound(arrList,2)
      totsum = totsum + arrList(16,intLoop)

	if arrList(0,intLoop)= "90201501120013" or arrList(0,intLoop)="90201610010124" or arrList(0,intLoop)="90201611130141" or arrList(0,intLoop)="90201611140136" or arrList(0,intLoop)="90201611200158" or arrList(0,intLoop)="90201611260172" or arrList(0,intLoop)="90201612060169" or arrList(0,intLoop)="90201612100180" or arrList(0,intLoop)="90201612120174" or arrList(0,intLoop)="90201612210190" then
			totBasePay 			= totBasePay + arrList(7,intLoop)
			totOverTimePay 		= totOverTimePay + arrList(8,intLoop)
			totNightTimePay 	= totNightTimePay + arrList(9,intLoop)
			totHolidayPay 		= totHolidayPay + arrList(10,intLoop)
			totFoodPay 			= totFoodPay + arrList(11,intLoop)
	else
			totBasePay 			= totBasePay + arrList(7,intLoop)+arrList(39,intLoop)
			totOverTimePay 		= totOverTimePay + arrList(8,intLoop)+arrList(40,intLoop)
			totNightTimePay 	= totNightTimePay + arrList(9,intLoop)+arrList(41,intLoop)
			totHolidayPay 		= totHolidayPay + arrList(10,intLoop)+arrList(42,intLoop)
			totFoodPay 			= totFoodPay + arrList(11,intLoop)+arrList(43,intLoop)
	end if
			totPositionPay 		= totPositionPay + arrList(12,intLoop)
			totBestPay 			= totBestPay + arrList(13,intLoop)
			totLongWorkPay 		= totLongWorkPay + arrList(14,intLoop)
			totAddPay 			= totAddPay + arrList(31,intLoop)
			totYearPay 			= totYearPay + arrList(36,intLoop)
			totBonusPay 		= totBonusPay + arrList(37,intLoop)

			if arrList(26,intLoop)=13 then
				totWorkTime 		= totWorkTime + arrList(35,intLoop)
			end if

			totReCalSum = totReCalSum + arrList(44,intLoop)

            bufStr = bufStr & arrList(45,intLoop)
			bufStr = bufStr & "," &arrList(0,intLoop)
			bufStr = bufStr & "," &arrList(22,intLoop)
			bufStr = bufStr & "," &arrList(38,intLoop)
            if arrList(33,intLoop) <> "" then
			bufStr = bufStr & "," &arrList(32,intLoop)& "/" &arrList(33,intLoop)& "(" &arrList(34,intLoop)&"��)"
            else
			bufStr = bufStr & ","
			end if
            IF arrList(26,intLoop)  =   13 THEN
            bufStr = bufStr & ","&arrList(28,intLoop)
            else
            bufStr = bufStr & ","
            end if
			bufStr = bufStr & ","&arrList(20,intLoop)
			 if arrList(0,intLoop)= "90201501120013" or arrList(0,intLoop)="90201610010124" or arrList(0,intLoop)="90201611130141" or arrList(0,intLoop)="90201611140136" or arrList(0,intLoop)="90201611200158" or arrList(0,intLoop)="90201611260172" or arrList(0,intLoop)="90201612060169" or arrList(0,intLoop)="90201612100180" or arrList(0,intLoop)="90201612120174" or arrList(0,intLoop)="90201612210190" then
			bufStr = bufStr & ","&arrList(7,intLoop)
			bufStr = bufStr & ","&arrList(8,intLoop)
			bufStr = bufStr & ","&arrList(9,intLoop)
			bufStr = bufStr & ","&arrList(10,intLoop)
			bufStr = bufStr & ","&arrList(11,intLoop)
			 else
			bufStr = bufStr & ","&arrList(7,intLoop)+arrList(39,intLoop)
			bufStr = bufStr & ","&arrList(8,intLoop)+arrList(40,intLoop)
			bufStr = bufStr & ","&arrList(9,intLoop)+arrList(41,intLoop)
			bufStr = bufStr & ","&arrList(10,intLoop)+arrList(42,intLoop)
			bufStr = bufStr & ","&arrList(11,intLoop)+arrList(43,intLoop)
			 end if
			bufStr = bufStr & ","&arrList(12,intLoop)
			bufStr = bufStr & ","&arrList(13,intLoop)
			bufStr = bufStr & ","&arrList(14,intLoop)
			bufStr = bufStr & ","&arrList(31,intLoop)
			bufStr = bufStr & ","&arrList(36,intLoop)
			bufStr = bufStr & ","&arrList(37,intLoop)
			 IF arrList(26,intLoop)=13 then
			bufStr = bufStr & ","&fnSetTimeFormat(arrList(35,intLoop))
            else
            bufStr = bufStr & ","
			end if
			bufStr = bufStr & ","&arrList(16,intLoop)
			bufStr = bufStr & "," &fnGetStateDesc(arrList(17,intLoop))

			if C_MngPart or C_ADMIN_AUTH or C_PSMngPart then
			bufStr = bufStr & "," &arrList(44,intLoop)
			end if

            bufStr =bufStr & VbCrlf
		next

			response.write bufStr
	 end if
%>
 <!-- #include virtual="/lib/db/dbclose.asp" -->