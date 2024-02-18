<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/NoUSBAllowIpList.asp"-->
<!-- #include virtual="/lib/checkUSBAllowIpList.asp"-->
<!-- #include virtual="/login/partner_loginCheck_function.asp"-->
<!-- #include virtual="/admin/incTenRedisSession.asp"-->
<%
dim manageUrl
IF (application("Svr_Info")	= "Dev") Then
	manageUrl = "http://"&request.ServerVariables("HTTP_HOST")
Else
	manageUrl = "https://"&request.ServerVariables("HTTP_HOST")
End If

'// ���� ���� �� ���۰� ����
dim userid, userpass, Enc_userpass, Enc_userpass64, backurl, tokenSn, lgnMethod, AuthNo
dim saved_id

lgnMethod = requestCheckVar(trim(request.Form("lgnMethod")),1)
if lgnMethod="S" THEN
	userid  = requestCheckVar(trim(request.Form("usid")),32)
	userpass = requestCheckVar(trim(request.Form("uspwd")),32)
	saved_id= requestCheckVar(trim(request.Form("saved_sid")),1)
else
	userid  = requestCheckVar(trim(request.Form("uid")),32)
	userpass = requestCheckVar(trim(request.Form("upwd")),32)
	saved_id= requestCheckVar(trim(request.Form("saved_id")),1)
end if

''USB �������� �α��� üũ
'Dim NoUsbValidIP : NoUsbValidIP = fnIsNoUsbAllowIp
Dim NoUsbValidIP
'if (NoUsbValidIP = False) then
	NoUsbValidIP = fncheckAllowIPWithByDB("Y", "", "")
'end if

''2017/04/20 REFERER ���°� �߰��� ����.----------------------, ��ü�α����� ���̻� �̰��� ��Ž..
dim iref : iref = Request.ServerVariables("HTTP_REFERER")
dim irefIP : irefIP = request.ServerVariables("REMOTE_ADDR")

if (Instr(iref,"webadmin.10x10.co.kr")<1) then
	if NOT G_IsLocalDev then
		Call fn_plogin_AddIISLOG("addlog=plogin&sub=noref&uid="&userid)
		response.write("<script>window.alert('���̻� ����� �� ���� �������Դϴ�.');</script>")
		response.write("<script>history.go(-1);</script>")
		dbget.close()	:	response.End
	end if
end if
''-----------------------------------------------------------


Enc_userpass = md5(userpass)
Enc_userpass64 = SHA256(md5(userpass))
tokenSn = requestCheckVar(trim(request.Form("tokenSn")),26)

AuthNo = requestCheckVar(trim(request.Form("sAuthNo")),6)

dim dbpassword
dim sql
dim errMsg
dim frontId

dim lockTerm, failNo
failNo = 5			'// �α��� ���� ����
lockTerm = 15		'// ���� ��� �ð� ����(��)

'### ���۰� Ȯ��
if ( userid = "" or userpass = "") then
    response.write("<script>window.alert('���̵� �Ǵ� ��й�ȣ�� �Էµ��� �ʾҽ��ϴ�.');</script>")
    response.write("<script>history.go(-1);</script>")
    dbget.close()	:	response.End
end if

'### ���� �α� Ȯ��
dim lastlogindt,lastpwchgdt,lastInfoChgDT, lastloginOrRegDiff, partnerUserdiv, Enc_2password64, regdiffDt
dim isFirstConnect : isFirstConnect = false '' ���̵� �߱� ���� ��������
dim isRequirePwdUp : isRequirePwdUp = false '' �н����� ���� �ʿ�
dim isRequireInfoUp : isRequireInfoUp = false '' ��������� ���� �ʿ�
dim isChangedIp : isChangedIp = false 			'' ����IP ��������

sql = "select  isNull(max(regdate),getdate()) as regdate " &VbCRLF
sql = sql + "	,isNull(sum(Case loginSuccess When 'N' Then 1 end),0) as FailCnt " &VbCRLF
sql = sql + "	,max(Case When rowNum=1 and loginSuccess='Y' Then refip end) as lastIP " &VbCRLF
sql = sql + "from (select top " & failNo & " regdate, loginSuccess, refip, ROW_NUMBER() over(partition by loginSuccess order by idx desc) as rowNum " &VbCRLF
sql = sql + "	from [db_log].[dbo].tbl_partner_login_log with (nolock)" &VbCRLF
sql = sql + "	where userid='" & userid & "' " &VbCRLF
sql = sql + "		and loginSuccess in ('Y','S') " &VbCRLF
sql = sql + "	order by idx desc) as pLog " &VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
	'// ���� �α��� ���� �� �����ð� ���� ���� ���
	if (datediff("n",rsget("regdate"),now)<lockTerm) and (rsget("FailCnt")>=failNo) then
	    Call fn_plogin_AddIISLOG("addlog=plogin&sub=lock&uid="&userid)
	    response.write("<script>window.alert('��й�ȣ�� �������� " & failNo & "�� Ʋ�� ���̵� �����ϴ�.\n" & (lockTerm-datediff("n",rsget("regdate"),now)) & "�� �� �ٽ� �α����� ���ּ���.');</script>")
	    response.write("<script>history.go(-1);</script>")
	    dbget.close()	:	response.End
	end if
	if rsget("lastIP")<>irefIP and left(irefIP,8)<>"192.168." and left(irefIP,9)<>"172.16.1." and irefIP <> "::1" then isChangedIp = true
rsget.Close


''## IP  Ȯ�� 2017/04/11
if (IspartnerLoginRejectIP()) then
    Call fn_plogin_AddIISLOG("addlog=plogin&sub=rjtip&uid="&userid)
    response.write("<script>window.alert('������ �������ϴ�. ');</script>")
    response.write("<script>history.go(-1);</script>")
	dbget.close()	:	response.End
end if

dim GeoIpCCD : GeoIpCCD = getGeoIpCountryCode()
dim RefCode : RefCode = getConSVCByUagentOrRefer()
''dim AuthReqIP : AuthReqIP = IsPartnerAuthRequireIP(userid,GeoIpCCD,FALSE)

if (GeoIpCCD="--") and (application("Svr_Info")="Dev") then GeoIpCCD="KR"
RefCode = "P:"&RefCode  ''(��)�α��� �����ϱ����Ѱ�


'### �������� ���� '//2011-03-9 �ѿ��(������) ���� - ��Ǫ�� ���̵� �߰�
sql = "select top 1 A.id, A.company_name, A.tel, A.fax, A.url, A.email, A.userdiv, A.Enc_password, A.Enc_password64, A.groupid " + vbCrlf
sql = sql + "	, B.part_sn, A.level_sn, B.job_sn, B.username,  B.direct070, B.usermail, B.posit_sn, IsNull(B.empno, '') as empno, B.frontid " + vbCrlf
sql = sql + "	, A.lastlogindt, isNULL(A.lastpwchgdt,'2001-01-01') as lastpwchgdt " + vbCrlf
sql = sql + "	, isNULL(A.lastInfoChgDT,A.regdate) as lastInfoChgDT, IsNull(B.criticinfouser,0) as criticinfouser " + vbCrlf
sql = sql + " , b.lv1customerYN, b.lv2partnerYN, b.lv3InternalYN" + vbCrlf
sql = sql + " ,datediff(d,isnull((CASE WHEN isNULL(A.lastlogindt,'2001-01-01')>isNULL(A.lastPwChgDT,'2001-01-01') THEN A.lastlogindt ELSE A.lastPwChgDT END),A.regdate),getdate()) as lastloginOrRegDiff " + vbCrlf ''�����α���or����� �����Ⱓmonth.  2017/04/10 �߰�
sql = sql + " ,datediff(d,A.regdate,getdate()) as regdiffDt " + vbCrlf
sql = sql + " , isNULL(A.Enc_2password64,'') as Enc_2password64" + vbCrlf
sql = sql + " ,(select top 1 shopid" + vbCrlf
sql = sql + " 	from db_partner.dbo.tbl_partner_shopuser with (nolock)" + vbCrlf
sql = sql + " 	where b.empno=empno and firstisusing='Y') as firstshopid" + vbCrlf
sql = sql + " from [db_partner].[dbo].tbl_partner as A with (nolock)" + vbCrlf
sql = sql + " 	left join db_partner.dbo.tbl_user_tenbyten as B with (nolock) ON A.id = B.userid AND B.isUsing = 1" + vbCrlf		'AND B.statediv = 'Y'
sql = sql + " where A.id = '" + userid + "'" + vbCrlf
'sql = sql + " and A.isusing='Y'"

'��翹���� ������� �α����� �ȵǼ�, ��翹���� ���� ����ؼ� ó��. '/2017.02.23 �ѿ��
sql = sql & " and (A.isusing='Y' or (" + vbCrlf
sql = sql & " 		(A.isusing='N' or B.statediv = 'N') and dateDiff(d,convert(varchar(10),getdate(),121),B.retireday)>=0 and B.retireday is not null" + vbCrlf
sql = sql & " 	)" + vbCrlf
sql = sql & " )" + vbCrlf

rsget.CursorLocation = adUseClient
rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

if  not rsget.EOF  then
	'// �α��� ���� Ȯ��
	partnerUserdiv  = rsget("userdiv")
	RefCode=replace(RefCode,"P:",partnerUserdiv&":")

	if rtrim(UCase(rsget("Enc_password64")))=trim(UCase(Enc_userpass64)) then


    	dbpassword  = rsget("Enc_password64")
    	lastlogindt = rsget("lastlogindt")  ''���� ���� ������
    	lastpwchgdt = rsget("lastpwchgdt")  ''���� �н����� ������
    	lastInfoChgDT = rsget("lastInfoChgDT")  ''���� ���������  ������
		frontId = rsget("frontid") '' ����Ʈ ID

    	isFirstConnect = isNULL(lastlogindt)
        lastloginOrRegDiff  = rsget("lastloginOrRegDiff")  ''���� �α��� OR �����
        Enc_2password64 = rsget("Enc_2password64")
        regdiffDt = rsget("regdiffDt")  ''����� ���� day

    	if (isFirstConnect) then
    	    isRequirePwdUp = true
    	    isRequireInfoUp = true
    	else
    	    ''  isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)
    	    ''''isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91) and (datediff("d",lastlogindt,now())>0) '' �н����� ������������ 2014/07/15 ���� �־����Ƿ�.. �켱 lastlogindt ���ǳ���.
    	    ''''if (CLNG(rsget("userdiv"))<10) then ''�ϴ� ����
    	    ''''    isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)
    	    ''''end if

    	    isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)  ''2017/04/10
    	    isRequireInfoUp=  (datediff("d",lastInfoChgDT,now())>91)
        end if

        ''�ϴ� ����.. // �ƿ� ����.(2017/04/20) ,CLNG(partnerUserdiv)>=10)(2017/04/21)
        if (CLNG(partnerUserdiv)>=10) then
            Call fn_plogin_AddIISLOG("addlog=plogin&sub=noscmpartner&uid="&userid)
            response.write("<script>window.alert('���̻� ����� �� ���� �������Դϴ�.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
        end if


        ''2017/04/10 �߰� 3���� �̻� ���� ���°�� / 2���н����� ���°��
        ''/login/partner_loginCheck_function.asp �� ������.. �߰��� ���� �Ҽ� ����.. �� ��� ���ľ�..
        if (lastloginOrRegDiff>91) then
            Call fn_plogin_AddIISLOG("addlog=plogin&sub=logntimenosee&uid="&userid)
            response.write("<script>window.alert('��Ⱓ ������� �ʾ� ������ �����ϴ�.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
        elseif ((partnerUserdiv="9999") and (isNULL(Enc_2password64) or Enc_2password64="")) then
            Call fn_plogin_AddIISLOG("addlog=plogin&sub=2ndpassnull&uid="&userid)
            response.write("<script>window.alert('2�� �н����� ����� �ʿ��մϴ�.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
        end if

        ''2017/04/12 �ؿ�IP => SMS ������ ��� �ؾ���.

        if (GeoIpCCD<>"KR") and Not(NoUsbValidIP) then
            if (AuthNo="") then
                Call fn_plogin_AddIISLOG("addlog=plogin&sub=geoipfail&uid="&userid&"&geoipccd="&GeoIpCCD)
                response.write("<script>window.alert('SMS ������ �ʿ��մϴ�..');</script>")
                response.write("<script>history.go(-1);</script>")
                dbget.close()	:	response.End
            end if
        end if

		'2022/08/01 ������ ����IP�� �ٸ� IP�ΰ�� SMS��������
		if isChangedIp and lgnMethod="U" then
			response.write("<script>window.alert('���ο� ȯ�濡�� �����ϼ̽��ϴ�.\nSMS ������ �������ּ���.');</script>")
			response.write("<script>location.replace(""/index.asp?lgnMethod=S"");</script>")
			dbget.close()	:	response.End
		end if



        ''------------------------------------------------------------------------------------

        session("ssBctId") = rsget("id")
        session("ssBctDiv") = rsget("userdiv")
        session("ssBctBigo") = rsget("firstshopid")
		session("ssBctSn") = rsget("empno")
        IF session("ssBctDiv") <= 9 THEN
        	 session("ssBctCname") = rsget("username")
        	 session("ssBctEmail") = db2html(rsget("usermail"))
        ELSE
        	if isnull(rsget("company_name")) then
        		session("ssBctCname") = rsget("username")
        	else
        		session("ssBctCname") = db2html(rsget("company_name"))
        	end if

        	session("ssBctEmail") = db2html(rsget("email"))
    	END IF

		session("ssGroupid") = rsget("groupid")
		session("ssAdminPsn") = rsget("part_sn")		'�μ� ��ȣ
		session("ssAdminLsn") = rsget("level_sn")		'��� ��ȣ
		session("ssAdminPOsn") = rsget("job_sn")		'��å ��ȣ
		session("ssAdminPOSITsn") = rsget("posit_sn")		'���� ��ȣ
		session("ssAdminCLsn") = rsget("criticinfouser")	'�������� ��ޱ���
		session("ssAdminlv1customerYN") = rsget("lv1customerYN")
		session("ssAdminlv2partnerYN") = rsget("lv2partnerYN")
		session("ssAdminlv3InternalYN") = rsget("lv3InternalYN")

		'3PL SSO �� ��Ű����(�������̵� + ���Ӿ����� + ��������) �� ��ȣȭ
		'�α��� �� �����ǰ� ����Ǹ�(����Ʈ�� ���� ��) �α����� �����Ѵ�.
		'�ڵ� ����ȭ�� ���� ��й�ȣ�� ��Ű�� �������� �ʴ´�. ���� �����ʿ�.(��� �ܹ��� ��ȣȭ �� ��Ű����)
		Response.Cookies("ThreePL").Domain				= "10x10.co.kr"
		Response.Cookies("ThreePL")("UserID")			= TBTEncrypt(CStr(rsget("id")) & "," & Request.ServerVariables("REMOTE_HOST") & "," & Left(now(), 10))

        '2014-12-17 ������ // API������ ��Ű����
		Response.Cookies("wapi").Domain				= "10x10.co.kr"
		Response.Cookies("wapi")("UserID")			= TBTEncrypt(CStr(rsget("id")) & "," & Request.ServerVariables("REMOTE_HOST") & "," & Left(now(), 10))
		If isnull(rsget("part_sn")) OR rsget("part_sn") = "" Then
		Else
			Response.Cookies("wapi")("PartSN") 		= TBTEncrypt(rsget("part_sn") & "," & Request.ServerVariables("REMOTE_HOST") & "," & Left(now(), 10))
		End If

		'// FrontAPI�� ��Ű����
		If frontId <> "" Then
			Dim ssnlogindt, retSsnHash, cookieDomain
			ssnlogindt = fnDateTimeToLongTime(now())
			session("ssnlogindt") = ssnlogindt
			retSsnHash = fnDBSessionCreateV2(frontId)

			If Application("Svr_Info") = "Dev" And InStr(Request.ServerVariables("HTTP_REFERER"), "localhost") > 0 Then
				cookieDomain = "localhost"
			Else
				cookieDomain = "10x10.co.kr"
			End If

			Response.Cookies("pinfo").domain = cookieDomain
			Response.Cookies("pinfo")("ssndt") = ssnlogindt
			Response.Cookies("pinfo")("ssnhash") = retSsnHash
		End If

		'���̵�����
		response.Cookies("ScmSave").domain = "10x10.co.kr"
    	response.cookies("ScmSave").Expires = Date + 30	'1������ ��Ű ����
	    If saved_id = "o" Then
	    	response.cookies("ScmSave")("SAVED_ID") = tenEnc(CStr(rsget("id")))
	    Else
	    	response.cookies("ScmSave")("SAVED_ID") = ""
	    End If

		'���� ��Ű �������� �ɱ�
		Dim cookieData, scResult, lp
		cookieData = Request.ServerVariables("HTTP_COOKIE")

		if instr(cookieData,"ASPSESSIONID")>0 then
			cookieData = Split(cookieData,";")
			for lp=0 to ubound(cookieData)
				if instr(Split(cookieData(lp),"=")(0),"ASPSESSIONID")>0 then
					scResult = scResult & cookieData(lp) & ";"
				end if
			next
		end if

		response.Cookies("TENSSID").domain = "10x10.co.kr"
		response.Cookies("TENSSID") = Base64EncodeUnicode(scResult)

		''�α�����(����)
	    rsget.close

	    if (isFirstConnect) then
	        ''���������ΰ�� �������� �Ⱥ� ��� �������� // ��� ������ ���ʷα��� ���� Update , 3�������� ������ �Ұ�� isRequirePwdUp �����߰�.
	    else
    	    if AuthNo<>"" then
    	    	Call AddPartnerLoginLogWithGeoIpCode (userid,"Y",AuthNo,GeoIpCCD)
    	    elseif tokenSn<>"" then
    	    	Call AddPartnerLoginLogWithGeoIpCode (userid,"Y",tokenSn,GeoIpCCD)
    	    else
    	        Call AddPartnerLoginLogWithGeoIpCode (userid,"Y",RefCode,GeoIpCCD)
    	    end if
    	end if
    	Call fn_plogin_AddIISLOG("addlog=plogin&sub=loginsucc&uid="&userid&"&userdiv="&partnerUserdiv)
	else
	    ''�α�����(����)
	    rsget.close
	    if AuthNo<>"" then
	    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",AuthNo,GeoIpCCD)
	    elseif tokenSn<>"" then
	        Call AddPartnerLoginLogWithGeoIpCode (userid,"N",tokenSn,GeoIpCCD)
	    else
	    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",RefCode,GeoIpCCD)
	    end if

        Call fn_plogin_AddIISLOG("addlog=plogin&sub=faillogin&uid="&userid&"&userdiv="&partnerUserdiv)
        response.write("<script>window.alert('���̵� �Ǵ� ��й�ȣ�� Ʋ�Ƚ��ϴ�.\n��й�ȣ ��ҹ��ڸ� Ȯ�����ּ���. ');</script>")
        response.write("<script>history.go(-1);</script>")
        dbget.close()	:	response.End
	end if
else
	'' �α��߰�.2017/04/12 F
	if AuthNo<>"" then
    	Call AddPartnerLoginLogWithGeoIpCode (userid,"F",AuthNo,GeoIpCCD)
    elseif tokenSn<>"" then
        Call AddPartnerLoginLogWithGeoIpCode (userid,"F",tokenSn,GeoIpCCD)
    else
    	Call AddPartnerLoginLogWithGeoIpCode (userid,"F",RefCode,GeoIpCCD)
    end if

    '// �������� , �������
	Call fn_plogin_AddIISLOG("addlog=plogin&sub=nouserid&uid="&userid)
    response.write("<script>window.alert('������ Ȱ��ȭ ���� �ʾҰų�, ���̵� �Ǵ� ��й�ȣ�� Ʋ�Ƚ��ϴ�.');</script>")
    response.write("<script>history.go(-1);</script>")
    dbget.close()	:	response.End
end if


'### �α��� ���� ���к� ó�� ------------------------------------------------------

''�����ӽ� => ���� �̰����� �α��� �Ұ�. 2017/04/21


dim ssnTmpUID
dim isNoComplexPwTxt : isNoComplexPwTxt = chkPasswordComplex_Len6Ver(userid,userpass)

if trim(UCase(dbpassword))=trim(UCase(Enc_userpass64)) then

    ''// ���� �α��� ������ �н����� ����(2014/05/19 ������)
    if ((isFirstConnect) or (isRequirePwdUp) or (isNoComplexPwTxt<>"")) then

        ssnTmpUID = session("ssBctId")

        '' session.Abandon is Async ?
        '' http://stackoverflow.com/questions/1470445/what-is-the-difference-between-session-abandon-and-session-clear
        Session.Contents.RemoveAll()

        CAll fnCookieExpire()

        session("ssnTmpUID")= ssnTmpUID   ''��� ����� ��� ���ǰ�

        if (isFirstConnect) then
            response.write "<script language='javascript'>" &vbCrLf &_
    						"	alert('���� �α��ν� ��й�ȣ�� �����ϼž� �մϴ�. \n��й�ȣ ������������ �̵��մϴ�.');" &vbCrLf &_
    						"	self.location='/login/modifyPassword.asp';" &vbCrLf &_
    						"</script>"
    		dbget.close()	:	response.End
        end if

        ''// N(3)������ ��й�ȣ ������Ѱ�� ������������ �̵�
        if (isRequirePwdUp) then
            response.write "<script language='javascript'>" &vbCrLf &_
    						"	alert('3���� �̻� ��й�ȣ�� �������� �����̽��ϴ�. \n��й�ȣ ������������ �̵��մϴ�.');" &vbCrLf &_
    						"	self.location.href='/login/modifyPassword.asp';" &vbCrLf &_
    						"</script>"
    		dbget.close()	:	response.End
        end if

        '// ��й�ȣ ��ȭ ��å ����(2008.12.12; ������)  'chkPasswordComplex => chkPasswordComplex_Len6Ver 2016/09/20
    	if (isNoComplexPwTxt)<>"" then
    		response.write "<script language='javascript'>" &vbCrLf &_
    						"	alert('" & chkPasswordComplex_Len6Ver(userid,userpass) & "\n��й�ȣ ������������ �̵��մϴ�.');" &vbCrLf &_
    						"	self.location='/login/modifyPassword.asp';" &vbCrLf &_
    						"</script>"
    		dbget.close()	:	response.End
    	end if

    end if



    response.Cookies("partner").domain = "10x10.co.kr"
    response.Cookies("partner")("userid") = session("ssBctId")
    response.Cookies("partner")("userdiv") = session("ssBctDiv")



    ''�����ΰ�� ����ƮPw�� ����Pw�� ������� errMsg �� USB��ū Ȯ��
    if (session("ssBctDiv")<=9) then

        '''20120621 �߰�//������ - ���Ѽ����� ���Դ°�찡 ����...
        if (session("ssAdminLsn")<1) then
            session.Abandon
            CAll fnCookieExpire()
            response.write("<script>window.alert('������ ������ ���� �ʽ��ϴ�. ������ ���ǿ��.');top.location = '/';</script>")
			dbget.close()	:	response.End
        end if

        sql = "select top 1 * from "
        sql = sql + " [db_user].[dbo].tbl_logindata u with (nolock)"
        sql = sql + " where u.userid='" & session("ssBctId") & "'" &VbCRLF
        sql = sql + " and u.Enc_userpass64='" & Enc_userpass64 & "'" &VbCRLF		'2014.06.25 SHA256����
        ''sql = sql + " and u.Enc_userpass='" & Enc_userpass & "'" &VbCRLF

        rsget.CursorLocation = adUseClient
        rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
    	if  not rsget.EOF  then
    		errMsg = "����Ʈ �� ���� ��й�ȣ�� �����ϰ� ����ϰ� �ֽ��ϴ�. \n\nMyInfo���� ���� ��й�ȣ�� �����Ͽ� ����ϼ���."
    	end if
    	rsget.close


		'// ���� �α��ι�� �߰�(2011.06.14; ������)
		if lgnMethod="" then
		    session.Abandon
		    CAll fnCookieExpire()
		    Call fn_plogin_AddIISLOG("addlog=plogin&sub=lologinmtd&uid="&userid)
		    response.write("<script>window.alert('�ٹ����� ������ �α����������� �ƴմϴ�.\n������ �������� �̵��մϴ�.\n\n���������� �α����������� ����Ǿ����ϴ�.\n������������ ������ ���� ���ã�⸦ �������ּ���.');top.location = '"&getSCMURL&"/';</script>")
		    dbget.close()	:	response.End
		else

			if lgnMethod="U" then
				'// USB��ū Ȯ��(2008.06.19; ������) //
				if (tokenSn="") then
				    if (NoUsbValidIP) then ''2014/10/29 �߰�, 2018-05-31 ����, skyer9
				        session("sslgnMethod") = "S"
				    else
				        session.Abandon
				        CAll fnCookieExpire()
				        Call fn_plogin_AddIISLOG("addlog=plogin&sub=notokensn&uid="&userid)
    				    response.write("<script>window.alert('USBŰ�� �����ϴ�.\n\n�ٹ����� USBŰ�� ����� ��ġ�Ǿ��ִ��� Ȯ�� �� �ٽ� �α������ּ���.');top.location = '/';</script>")
    				    dbget.close()	:	response.End
    				end if
				else
					'### ��ȿ��ȣ ó��(db_partner.dbo.tbl_admin_key���� ��ȿ��ȣ Ȯ��) ###
					'Token �Ϸù�ȣ Ȯ��(DB)
					sql = "select count(key_idx) " & vbCRLF
					sql = sql & " from db_partner.dbo.tbl_admin_key with (nolock)" & vbCRLF
					sql = sql & " where key_idx='" & tokenSn & "' and del_isusing='Y'"
					rsget.CursorLocation = adUseClient
                    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

					if rsget(0)<=0 then
					    session.Abandon
					    CAll fnCookieExpire()
					    Call fn_plogin_AddIISLOG("addlog=plogin&sub=invalidkensn&uid="&userid)
						response.write("<script>window.alert('��ȿ�� USBŰ�� �ƴմϴ�.\n�����ڿ��� �������ּ���.');top.location = '/';</script>")
						dbget.close()	:	response.End
					end if
					rsget.Close

                    ''2017/06/20 �߰�. USB���� ���IP ����.
                    if (Not IsUsbLoginAlowIp) then
                        session.Abandon
                        CAll fnCookieExpire()
					    Call fn_plogin_AddIISLOG("addlog=plogin&sub=usbnoip&uid="&userid)
						response.write("<script>window.alert('���� ������ ��ΰ� �ƴմϴ�. SMS ������ ����ϼ���.');top.location = '/';</script>")
						dbget.close()	:	response.End
                    end if
				end if
				'// USB��ūȮ�� �� //

			elseif lgnMethod="S" then

				'// SMS���� �α���
				if AuthNo="" then
				    session.Abandon
				    CAll fnCookieExpire()
				    Call fn_plogin_AddIISLOG("addlog=plogin&sub=noauthno&uid="&userid)
				    response.write("<script>window.alert('������ȣ�� �����ϴ�.\n�޴������� ���۵� ������ȣ�� ��Ȯ�� �Է����ּ���.');top.location = '/?lgnMethod="&lgnMethod&"';</script>")
				    dbget.close()	:	response.End
				else
					'��ȿ�� �ð����� ������ȣ Ȯ��
					sql = "select USBTokenSn " & vbCRLF
					sql = sql & " from db_log.dbo.tbl_partner_login_log with (nolock)" & vbCRLF
					sql = sql & " where userid='" & userid & "' " & vbCRLF
					sql = sql & " 	and loginSuccess='S' " & vbCRLF
					sql = sql & " 	and datediff(ss,regdate,getdate()) between 0 and 180"
					rsget.CursorLocation = adUseClient
                    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

					if rsget.EOF or rsget.BOF  then
					    session.Abandon
					    CAll fnCookieExpire()
					    Call fn_plogin_AddIISLOG("addlog=plogin&sub=expireauthno&uid="&userid)
						response.write("<script>window.alert('�Է� ���ѽð��� �ʰ��Ǿ����ϴ�.\n�ٽ� ������ȣ�� �߱޹޾� �Է����ּ���.');top.location = '/?lgnMethod="&lgnMethod&"';</script>")
						dbget.close()	:	response.End
					else
						if trim(rsget("USBTokenSn"))<>trim(AuthNo) then
						    session.Abandon
						    CAll fnCookieExpire()
						    Call fn_plogin_AddIISLOG("addlog=plogin&sub=nomatchauthno&uid="&userid)
							response.write("<script>window.alert('�޴������� �߼۵� ������ȣ���� �ƴմϴ�.\n��Ȯ�� �Է����ּ���.');top.location = '/?lgnMethod="&lgnMethod&"';</script>")
							dbget.close()	:	response.End
						else
							'// adminbodyhead.asp�� USBüũ�� ���Ϸ��� ���ǿ� SMS�������� ����
							session("sslgnMethod") = "S"
						end if
					end if
					rsget.Close
				end if
            else
                ''2017/06/20 �߰�
                session.Abandon
    		    CAll fnCookieExpire()
    		    Call fn_plogin_AddIISLOG("addlog=plogin&sub=xloginmtd&uid="&userid)
    		    response.write("<script>window.alert('�ٹ����� ������ �α����������� �ƴմϴ�.\n������ �������� �̵��մϴ�.\n\n���������� �α����������� ����Ǿ����ϴ�.\n������������ ������ ���� ���ã�⸦ �������ּ���.');top.location = '"&getSCMURL&"/';</script>")
    		    dbget.close()	:	response.End
			end if

		end if
    end if


    if (session("ssBctId")="10x10") then
        ''������.
        session.Abandon
        CAll fnCookieExpire()
        Call fn_plogin_AddIISLOG("addlog=plogin&sub=notenid&uid="&userid)
        dbget.close()	:	response.End

    ''����Level
    elseif (session("ssBctDiv")<=9) then

    	if (errMsg<>"") then
            response.write "<script language='javascript'>alert('" & errMsg & "');</script>"
        end if

		''2018/12/18 incTenRedisSession
		Call fn_RDS_SSN_SET()

    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/admin/index.asp')</script>"
        dbget.close()	:	response.End

    else
        session.Abandon
        CAll fnCookieExpire()
        Call fn_plogin_AddIISLOG("addlog=plogin&sub=notauth&uid="&userid)
        response.write "<script language='javascript'>alert('�����̾����ϴ�.');location.replace('/')</script>"
        dbget.close()	:	response.End
    end if
end if

function fnCookieExpire()
    Response.Cookies("partner").domain = "10x10.co.kr"
    Response.Cookies("partner") = ""
    Response.Cookies("partner").Expires = Date - 1

    Response.Cookies("ThreePL").Domain	= "10x10.co.kr"
    Response.Cookies("ThreePL") = ""
    Response.Cookies("ThreePL").Expires = Date - 1

    Response.Cookies("wapi").Domain	= "10x10.co.kr"
    Response.Cookies("wapi") = ""
    Response.Cookies("wapi").Expires = Date - 1

	response.Cookies("TENSSID").domain = "10x10.co.kr"
	response.Cookies("TENSSID") = ""
    Response.Cookies("TENSSID").Expires = Date - 1

	'' require /admin/incTenRedisSession.asp
    response.Cookies(GG_RDS_COOKIE_KEYNAME).domain = fn_RDS_getCookieDomain()
    response.Cookies(GG_RDS_COOKIE_KEYNAME) = ""
	Response.Cookies(GG_RDS_COOKIE_KEYNAME).Expires = Date - 1
end function

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
