<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����α���
' Hieditor : ������ ����
'			 2023.09.07 �ѿ��(�����й�ȣ �����α��γ�¥ üũ ����)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/login/partner_loginCheck_function.asp"-->
<%
'// ������ ���� �α� ���� �Լ�
''Sub AddLoginLog(param1,param2,param3)
''    dim sqlStr, reFAddr
''    reFAddr = request.ServerVariables("REMOTE_ADDR")
''
''    sqlStr = " insert into [db_log].[dbo].tbl_partner_login_log" + VbCrlf
''	sqlStr = sqlStr + " (userid,refip,loginSuccess,USBTokenSn)" + VbCrlf
''	sqlStr = sqlStr + " values(" + VbCrlf
''	sqlStr = sqlStr + " '" + param1 + "'," + VbCrlf
''	sqlStr = sqlStr + " '" + Left(reFAddr,16) + "'," + VbCrlf
''	sqlStr = sqlStr + " '" + param2 + "'," + VbCrlf
''	sqlStr = sqlStr + " '" + param3 + "'" + VbCrlf
''	sqlStr = sqlStr + " )" + VbCrlf
''
''    dbget.Execute sqlStr
''end Sub 
 
'// ���� ���� �� ���۰� ����
dim empno, userpass, backurl
dim saved_eno
dim IsLoginSuccess
empno  = requestCheckVar(trim(request.Form("usn")),32)			'// ���
userpass = requestCheckVar(trim(request.Form("unpwd")),32) 
saved_eno= requestCheckVar(trim(request.Form("saved_eno")),1)
	
dim dbpassword
dim sql
dim errMsg

dim lockTerm, failNo
failNo = 5			'// �α��� ���� ����
lockTerm = 15		'// ���� ��� �ð� ����(��)

'// ============================================================================
'### ���۰� Ȯ��
if (empno = "" or userpass = "") then
    response.write("<script>window.alert('��� �Ǵ� ��й�ȣ�� �Էµ��� �ʾҽ��ϴ�.');</script>")
    response.write("<script>window.location.href ='/index.asp?lgnMethod=N'</script>")
    dbget.close()	:	response.End
end if

''// 2017/06/19 �߰�============================================================
dim GeoIpCCD : GeoIpCCD = getGeoIpCountryCode()
if (GeoIpCCD="--") and (application("Svr_Info")="Dev") then GeoIpCCD="KR" 
    
dim iref : iref = Request.ServerVariables("HTTP_REFERER")
dim irefIP : irefIP = request.ServerVariables("REMOTE_ADDR")

'���� �����϶� ó��
if (GeoIpCCD="--") and (left(irefIP,8)="192.168.") then GeoIpCCD="KR" 

''SCM ���ؼ��� ����.
' if (Instr(iref,"webadmin.10x10.co.kr")<1) then 
'     Call fn_plogin_AddIISLOG("addlog=plogin&sub=noref&empno="&empno) 
'     response.write("<script>window.alert('���̻� ����� �� ���� �������Դϴ�.');</script>")
'     response.write("<script>history.go(-1);</script>")
'     dbget.close()	:	response.End
' end if

''�ؿ�IP �Ұ�.
if (GeoIpCCD<>"KR") then
    Call fn_plogin_AddIISLOG("addlog=plogin&sub=geoipfail&empno="&empno&"&geoipccd="&GeoIpCCD) 
    response.write("<script>window.alert('������ �Ұ� �մϴ�.');</script>")
    response.write("<script>history.go(-1);</script>")
    dbget.close()	:	response.End
end if

'' ����α����� Ư��IP�� ���� > �α��� ��� IP DB �˻�
if (NOT (application("Svr_Info")="Dev")) then
if NOT(fncheckAllowIPWithByDB("Y", "", "")) then
    Call fn_plogin_AddIISLOG("addlog=plogin&sub=invalidip&empno="&empno&"&refip="&irefIP) 
    response.write("<script>window.alert('������ �Ұ� �մϴ�. ������ ���� ���');</script>")
    response.write("<script>history.go(-1);</script>")
    dbget.close()	:	response.End
end if
end if
'// ============================================================================
'### ���� �α� Ȯ��
dim isFirstOrreqChgPwd : isFirstOrreqChgPwd = false ''���� ���� ������ �ʿ��Ұ�� ���� ���.
dim isRequirePwdUp : isRequirePwdUp = false '' �н����� ���� �ʿ�

sql = "select  isNull(max(regdate),getdate()) as regdate " &VbCRLF
sql = sql + "	,isNull(sum(Case loginSuccess When 'N' Then 1 end),0) as FailCnt " &VbCRLF
sql = sql + "	,(select top 1 regdate from [db_log].[dbo].tbl_partner_login_log " &VbCRLF
sql = sql + "		where userid='"&empno&"' " &VbCRLF
sql = sql + "		and loginSuccess in ('Y','R')" &VbCRLF  '' R �н����� ����.
sql = sql + "		order by idx desc) as lastloginSuccDt " &VbCRLF  ''���� �α��� �߰�
sql = sql + " from (select top " & failNo & " regdate, loginSuccess " &VbCRLF
sql = sql + "	from [db_log].[dbo].tbl_partner_login_log " &VbCRLF
sql = sql + "	where userid='" & empno & "' " &VbCRLF
sql = sql + "	order by idx desc) as pLog " &VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
    isFirstOrreqChgPwd = IsNull(rsget("lastloginSuccDt")) ''2014/05/19
	'// ���� �α��� ���� �� �����ð� ���� ���� ���
	if (datediff("n",rsget("regdate"),now)<lockTerm) and (rsget("FailCnt")>=failNo) then
	    response.write("<script>window.alert('��й�ȣ�� �������� " & failNo & "�� Ʋ�� ���̵� �����ϴ�.\n" & (lockTerm-datediff("n",rsget("regdate"),now)) & "�� �� �ٽ� �α����� ���ּ���.');</script>")
	     response.write("<script>window.location.href ='/?lgnMethod=N'</script>")
	    dbget.close()	:	response.End
	end if
rsget.Close

'// ============================================================================
'### �������� ����
Dim i_part_sn, i_username, i_level_sn, i_posit_sn
Dim i_LastEmpnoPassWordChangeDate , i_LastLoginOrRegDiff

sql = "SELECT TOP 1 "
sql = sql + "	B.Enc_emppass64 "
sql = sql + "	,B.part_sn "
sql = sql + "	,B.job_sn "
sql = sql + "	,B.username "
sql = sql + "	,B.direct070 "
sql = sql + "	,B.usermail "
sql = sql + "	,B.posit_sn "

'// TODO : ���� ���� �����ϰ� ��� �α��ν� �������� ��ȸ�������� ����
if (application("Svr_Info")="Dev") then
	'// sql = sql + "	,IsNull(A.level_sn, 10) as level_sn "
	sql = sql + "	,10 as level_sn "
else
	'// sql = sql + "	,IsNull(A.level_sn, 9) as level_sn "
	sql = sql + "	,9 as level_sn "
end if

sql = sql + "	,isNULL(b.lastEmpnoPwChgDT,'2001-01-01') as lastEmpnoPwChgDT "
sql = sql + "	,datediff(d,isnull((CASE WHEN isNULL(A.lastlogindt,'2001-01-01')>isNULL(b.lastEmpnoPwChgDT,'2001-01-01') THEN A.lastlogindt ELSE b.lastEmpnoPwChgDT END),A.regdate),getdate()) as lastloginOrRegDiff " 
sql = sql & " FROM db_partner.dbo.tbl_user_tenbyten B"				'// ��� �α���
sql = sql & " LEFT JOIN [db_partner].[dbo].tbl_partner AS A"
sql = sql & "	ON A.id = B.userid"
sql = sql + " WHERE "
sql = sql + "	1 = 1 "
sql = sql + "	AND B.isUsing = 1 "
sql = sql + "	AND B.empno = '" + CStr(empno) + "' "
'sql = sql + "	AND B.statediv = 'Y' "
'sql = sql + "	AND IsNull(A.isusing, 'Y') = 'Y' "

'��翹���� ������� �α����� �ȵǼ�, ��翹���� ���� ����ؼ� ó��. '/2017.02.23 �ѿ��
sql = sql & " and (B.statediv='Y' or (" + vbCrlf
sql = sql & " 		(B.statediv='N' or IsNull(A.isusing, 'N') = 'N') and dateDiff(d,convert(varchar(10),getdate(),121),B.retireday)>=0 and B.retireday is not null" + vbCrlf
sql = sql & " 	)" + vbCrlf
sql = sql & " )" + vbCrlf

'response.write sql
rsget.CursorLocation = adUseClient
rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

IsLoginSuccess = False
if not rsget.EOF then
	if rtrim(LCase(rsget("Enc_emppass64"))) = trim(LCase(sha256(md5(userpass)))) then
	    i_part_sn = rsget("part_sn")
	    i_username = rsget("username")
	    i_level_sn = rsget("level_sn")
		i_posit_sn = rsget("posit_sn")
		i_LastEmpnoPassWordChangeDate = rsget("lastEmpnoPwChgDT")
		i_LastLoginOrRegDiff = rsget("lastloginOrRegDiff")  ''���� �α��� OR �����
		IsLoginSuccess = True
	end if
end if
rsget.close

if (IsLoginSuccess = True) then
    '// �Ʒ� ���Ǹ� �����Ѵ�.
	session("ssBctSn") 		= empno
	session("ssBctDiv")		= "5000"		'// ���� ���� : ����������ȸ
	session("ssAdminPsn") 	= i_part_sn
	session("ssAdminPOSITsn") = i_posit_sn		'���� ��ȣ
	session("ssBctCname") 	= i_username

	'��� ��ȣ
	session("ssAdminLsn") 	= i_level_sn

	if (i_LastLoginOrRegDiff>91) then
		Call fn_plogin_AddIISLOG("addlog=plogin&sub=logntimenosee&empno="&empno)
		response.write("<script>window.alert('��Ⱓ ������� �ʾ� ������ �����ϴ�.');</script>")
		response.write("<script>history.go(-1);</script>")
		dbget.close()	:	response.End
	END IF
    
    ''// ���� �α��� ������ �н����� ����(2014/05/19 ������)
    if (isFirstOrreqChgPwd) then
        response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('���� �α��ν� ��й�ȣ�� �����ϼž� �մϴ�. \n��й�ȣ ������������ �̵��մϴ�.');" &vbCrLf &_
						"	self.location='/login/modifyPassword_empno.asp';" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End

		isRequirePwdUp = true
	else
		isRequirePwdUp = (datediff("d",i_LastEmpnoPassWordChangeDate,now())>91)  ''2017/04/10
    end if

	''// N(3)������ ��й�ȣ ������Ѱ�� ������������ �̵�
	if (isRequirePwdUp) then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('3���� �̻� ��й�ȣ�� �������� �����̽��ϴ�. \n��й�ȣ ������������ �̵��մϴ�.');" &vbCrLf &_
						"	self.location.href='/login/modifyPassword_empno.asp';" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
    
    '// ��й�ȣ ��ȭ ��å ����(2008.12.12; ������)
	if chkPasswordComplex(empno,userpass)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(empno,userpass) & "\n��й�ȣ ������������ �̵��մϴ�.');" &vbCrLf &_
						"	self.location='/login/modifyPassword_empno.asp';" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	
	'���̵�����
	response.Cookies("ScmSave").domain = "10x10.co.kr"
	response.cookies("ScmSave").Expires = Date + 30	'1������ ��Ű ����
    If saved_eno = "o" Then
    	response.cookies("ScmSave")("SAVED_Eno") = tenEnc(CStr(empno))
    Else
    	response.cookies("ScmSave")("SAVED_Eno") = ""
    End If	
end if

if (IsLoginSuccess) then
	''Call AddLoginLog (empno,"Y","")
	Call AddPartnerLoginLogWithGeoIpCode (empno,"Y","",GeoIpCCD)
else
	''Call AddLoginLog (empno,"N","")
	Call AddPartnerLoginLogWithGeoIpCode (empno,"N","",GeoIpCCD)

	response.write("<script>window.alert('���̵� �Ǵ� ��й�ȣ�� Ʋ�Ƚ��ϴ�.');</script>")
	response.write("<script>window.location.href ='/?lgnMethod=N'</script>")
	dbget.close()	:	response.End
end if


'// ============================================================================
response.write "<script language='javascript'>location.replace('/tenmember/index.asp')</script>"
dbget.close()	:	response.End

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
