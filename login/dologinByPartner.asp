<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0
 Response.AddHeader "Pragma","no-cache"
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/NoUSBAllowIpList.asp"-->
<!-- #include virtual="/login/partner_loginCheck_function.asp"-->
<%

function FnAddIISLOG(iAddLogs)
    ''addLog �߰� �α� //2016/12/29
    if (request.ServerVariables("QUERY_STRING")<>"") then iAddLogs="&"&iAddLogs
    response.AppendToLog iAddLogs
end function

dim manageUrl
IF application("Svr_Info")="Dev" THEN
	'manageUrl 	 = "http://testwebadmin.10x10.co.kr"
	manageUrl 	 = getSCMURL
ELSE
	manageUrl 	 = "http://webadmin.10x10.co.kr"
END IF

'// ���� ���� �� ���۰� ����
dim userid, userpass, Enc_userpass, Enc_userpass64, backurl, tokenSn, lgnMethod, AuthNo
dim loginNo,vIsSec
dim userpassSec1,userpassSec2,userpassSec, Enc_2userpass64
dim saved_id

loginNo = requestCheckVar(trim(request.Form("loginNo")),1)
vIsSec	= requestCheckVar(trim(request.Form("hidSec")),1)
saved_id= requestCheckVar(trim(request.Form("saved_id")),1)
if loginNo ="" then loginNo ="1"
if loginNo = "2" then
	userid  =  session("tmpUID")
	userpass = session("tmpUPWD")
else
	userid  = requestCheckVar(trim(request.Form("uid")),32)
	userpass = requestCheckVar(trim(request.Form("upwd")),32)
end if

Enc_userpass = md5(userpass)
Enc_userpass64 = SHA256(md5(userpass))
userpassSec1 = requestCheckVar(trim(request.Form("upwdS1")),32)
userpassSec2 = requestCheckVar(trim(request.Form("upwdS2")),32)
userpassSec = requestCheckVar(trim(request.Form("upwdS")),32)

tokenSn = requestCheckVar(trim(request.Form("tokenSn")),26)
lgnMethod = requestCheckVar(trim(request.Form("lgnMethod")),1)
AuthNo = requestCheckVar(trim(request.Form("sAuthNo")),6)



dim dbpassword,isdbpassword_sec,dbpassword2
dim db_id, db_userdiv,db_company_name,db_email,db_groupid

dim sql
dim errMsg

dim lockTerm, failNo
failNo = 5			'// �α��� ���� ����
lockTerm = 15		'// ���� ��� �ð� ����(��)

'### ���۰� Ȯ��
if ( userid = "" or userpass = "") then
    Call FnAddIISLOG("addlog=plogin&sub=no1step&loginNo="&loginNo&"&uid="&userid) ''2016/12/29
    response.write("<script>window.alert('���̵� �Ǵ� ��й�ȣ�� �Էµ��� �ʾҽ��ϴ�.');</script>")
    response.write("<script>history.go(-1);</script>")
    dbget.close()	:	response.End
end if

'### ���� �α� Ȯ��
dim lastlogindt,lastpwchgdt,lastInfoChgDT
dim isFirstConnect : isFirstConnect = false '' ���̵� �߱� ���� ��������
dim isRequirePwdUp : isRequirePwdUp = false '' �н����� ���� �ʿ�
dim isRequireInfoUp : isRequireInfoUp = false '' ��������� ���� �ʿ�

sql = "select  isNull(max(regdate),getdate()) as regdate " &VbCRLF
sql = sql & "	,isNull(sum(Case loginSuccess When 'N' Then 1 end),0) as FailCnt " &VbCRLF
sql = sql & "from (select top " & failNo & " regdate, loginSuccess " &VbCRLF
sql = sql & "	from [db_log].[dbo].tbl_partner_login_log " &VbCRLF
sql = sql & "	where userid='" & userid & "' " &VbCRLF
sql = sql & "	order by idx desc) as pLog " &VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
	'// ���� �α��� ���� �� �����ð� ���� ���� ���
	if (datediff("n",rsget("regdate"),now)<lockTerm) and (rsget("FailCnt")>=failNo) then
	    Call FnAddIISLOG("addlog=plogin&sub=lock&uid="&userid) ''2016/12/29
	    response.write("<script>window.alert('��й�ȣ�� �������� " & failNo & "�� Ʋ�� ���̵� �����ϴ�.\n" & (lockTerm-datediff("n",rsget("regdate"),now)) & "�� �� �ٽ� �α����� ���ּ���.');</script>")
	    response.write("<script>history.go(-1);</script>")
	    dbget.close()	:	response.End
	end if

rsget.Close

''## IP  Ȯ�� 2017/04/11
if (IspartnerLoginRejectIP()) then
    Call FnAddIISLOG("addlog=plogin&sub=rjtip&uid="&userid) ''2016/12/29
    response.write("<script>window.alert('������ �������ϴ�. ');</script>")
    response.write("<script>history.go(-1);</script>")
	dbget.close()	:	response.End
end if

dim GeoIpCCD : GeoIpCCD = getGeoIpCountryCode()
dim RefCode : RefCode = getConSVCByUagentOrRefer()
dim AuthReqIP
dim db_id_Exists, db_Enc_password64, db_Enc_2password64
db_id_Exists = FALSE

if (GeoIpCCD="--") and (application("Svr_Info")="Dev") then GeoIpCCD="KR"

if loginNo = "2" then
	if vIsSec = "N" and userpassSec1 ="" then  ''�����.
	    Call FnAddIISLOG("addlog=plogin&sub=no2step&uid="&userid) ''2016/12/29
		response.write("<script>window.alert('��ϵ� 2�� ��й�ȣ�� �����ϴ�.���� �������ּ���');</script>")
        response.write("<script>history.go(-1);</script>")
        dbget.close()	:	response.End
	end if

	AuthReqIP = IsPartnerAuthRequireIP(userid,GeoIpCCD,TRUE)

	'### �������� ���� ��ü �������� ����. 2017/04/24
	sql = "select top 1 A.id, A.company_name, A.tel, A.fax, A.url, A.email, A.userdiv, A.Enc_password, A.Enc_password64, A.groupid " & vbCrlf
	sql = sql & "	, A.level_sn " & vbCrlf
	sql = sql & "	, A.lastlogindt, isNULL(A.lastpwchgdt,'2001-01-01') as lastpwchgdt " & vbCrlf
	sql = sql & "	, isNULL(A.lastInfoChgDT,A.regdate) as lastInfoChgDT" & vbCrlf
	sql = sql & " 	, isNULL(A.Enc_2password64,'') as Enc_2password64"   ''���� 2017/04/10
	sql = sql & " from [db_partner].[dbo].tbl_partner as A " & vbCrlf
	sql = sql & " where A.id = '" & userid & "'" & vbCrlf
	sql = sql & " and A.isusing='Y'"
	rsget.CursorLocation = adUseClient
    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

	if  not rsget.EOF  then
	    db_id_Exists       = TRUE
	    dbpassword  = rsget("Enc_password64")
    	dbpassword2  = rsget("Enc_2password64")
    	lastlogindt = rsget("lastlogindt")  ''���� ���� ������
    	lastpwchgdt = rsget("lastpwchgdt")  ''���� �н����� ������
    	lastInfoChgDT = rsget("lastInfoChgDT")  ''���� ���������  ������

    	db_id = rsget("id")
    	db_userdiv = rsget("userdiv")
    	db_company_name = db2html(rsget("company_name"))
	    db_email = db2html(rsget("email"))
		db_groupid = rsget("groupid")

    end if
    rsget.close

    if (db_id_Exists) then
		'// �α��� ���� Ȯ��
		if trim(UCase(dbpassword))=trim(UCase(Enc_userpass64)) then
			if vIsSec = "Y" then
				if userpassSec ="" then
					Call FnAddIISLOG("addlog=plogin&sub=no2nd&uid="&userid&"&ruid="&request("uid")) ''2016/12/29
					response.write("<script>window.alert('2�� ��й�ȣ ���� �����ϴ�.Ȯ�����ּ���');</script>")
			        response.write("<script>history.go(-1);</script>")
			        dbget.close()	:	response.End
				end if


				Enc_2userpass64 = SHA256(md5(userpassSec))


				if trim(UCase(dbpassword2))<>trim(UCase(Enc_2userpass64)) then

    			    if AuthNo<>"" then
    			    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",AuthNo,GeoIpCCD)
    			    elseif tokenSn<>"" then
    			        Call AddPartnerLoginLogWithGeoIpCode (userid,"N",tokenSn,GeoIpCCD)
    			    else
    			    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",RefCode,GeoIpCCD)
    			    end if
    		        Call FnAddIISLOG("addlog=plogin&sub=fail2nd&uid="&userid) ''2016/12/29
    		        response.write("<script>window.alert('2�� ��й�ȣ�� Ʋ�Ƚ��ϴ�.Ȯ���� �ٽ� �õ����ּ���.');</script>")
    		        response.write("<script>history.go(-1);</script>")
    		        dbget.close()	:	response.End
				end if

				if (AuthReqIP) then
                    Call fn_plogin_AddIISLOG("addlog=plogin&sub=reqauth&uid="&userid)
                    Session.Contents.Remove("tmpUID")
  			        Session.Contents.Remove("tmpUPWD")

                    session("reauthUID") =  userid
                    response.write("<script>window.alert('���� ���� ȯ��� �ٸ� ȯ�濡�� �α��� �ϼ̽��ϴ�. ���� �������� �̵��մϴ�.');</script>")
                    response.write("<script>location.href='/login/reconfirmip.asp'</script>")
                    dbget.close() : response.End
              end if
			else
				''2��������� ���̻� �̰����� ����. 2017/04/24
				Call FnAddIISLOG("addlog=plogin&sub=2ndfail2&uid="&userid) ''2016/12/29
				response.write("<script>window.alert('���԰�ο� ������ �ֽ��ϴ�.');</script>")
		        response.write("<script>history.go(-1);</script>")
		        dbget.close()	:	response.End

''				Enc_2userpass64 = SHA256(md5(userpassSec1))
''
''				''2����� ������ 2017/04/10 1,2�� ������ �� ����. by eastone
''				if (trim(UCase(Enc_userpass64))=trim(UCase(Enc_2userpass64))) then
''				    Call FnAddIISLOG("addlog=plogin&sub=dupp2nd&uid="&userid) ''2016/12/29
''    		        response.write("<script>window.alert('1,2�� ��й�ȣ�� �����ϰ� ������ �� �����ϴ�.Ȯ���� �ٽ� �õ����ּ���.');</script>")
''    		        response.write("<script>history.go(-1);</script>")
''    		        dbget.close()	:	response.End
''				end if
''
''				dim objCmd,returnValue
''					Set objCmd = Server.CreateObject("ADODB.COMMAND")
''						With objCmd
''							.ActiveConnection = dbget
''							.CommandType = adCmdText
''							.CommandText = "{?= call db_partner.[dbo].[sp_Ten_partner_SetSecondPassWord]('"&userid&"',  '"&Enc_2userpass64&"' )}"
''							.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
''							.Execute, , adExecuteNoRecords
''							End With
''						    returnValue = objCmd(0).Value
''					Set objCmd = nothing
''
''				 if returnValue =-1 then
''				 	Call FnAddIISLOG("addlog=plogin&sub=exists2nd&uid="&userid) ''2016/12/29
''					response.write("<script>window.alert('2�� ��й�ȣ�� �̹� ��ϵǾ��ֽ��ϴ�.');</script>")
''			        response.write("<script>history.go(-1);</script>")
''			        dbget.close()	:	response.End
''				 elseif returnValue =0 then
''				 	Call FnAddIISLOG("addlog=plogin&sub=2ndfail&uid="&userid) ''2016/12/29
''					response.write("<script>window.alert('2�� ��й�ȣ ��Ͽ� �����߽��ϴ�. Ȯ�� �� �ٽ� ������ּ���');</script>")
''			        response.write("<script>history.go(-1);</script>")
''			        dbget.close()	:	response.End
''			     else
''			        Call FnAddIISLOG("addlog=plogin&sub=2ndok&uid="&userid) ''2016/12/29
''			   	    response.write("<script>window.alert('2�� ��й�ȣ�� ��ϵǾ����ϴ�. ');</script>")
''				 end if
			end if

	    	isFirstConnect = isNULL(lastlogindt)

	    	if (isFirstConnect) then
	    	    isRequirePwdUp = true
	    	    isRequireInfoUp = true
	    	else
	    	    '' isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)
	    	    ''isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91) and (datediff("d",lastlogindt,now())>0) '' �н����� ������������ 2014/07/15 ���� �־����Ƿ�.. �켱 lastlogindt ���ǳ���.
	    	    ''if (CLNG(db_userdiv)<10) then ''�ϴ� ����
	    	    ''    isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)
	    	    ''end if

	    	    isRequirePwdUp = (datediff("d",lastpwchgdt,now())>91)  ''2017/04/13
	    	    isRequireInfoUp=  (datediff("d",lastInfoChgDT,now())>91)
	        end if

			Session.Contents.Remove("tmpUID")
  			Session.Contents.Remove("tmpUPWD")

	        session("ssBctId") = db_id
	        session("ssBctDiv") = db_userdiv

        	if isnull(db_company_name) then
        		session("ssBctCname") = "..."
        	else
        		session("ssBctCname") = db_company_name
        	end if

        	session("ssBctEmail") = db_email
			session("ssGroupid") = db_groupid

			'���̵�����
    		response.Cookies("PASave").domain = "10x10.co.kr"
        	response.cookies("PASave").Expires = Date + 30	'1������ ��Ű ����
    	    If saved_id = "o" Then
    	    	response.cookies("PASave")("SAVED_ID") = tenEnc(CStr(db_id))
    	    Else
    	    	response.cookies("PASave")("SAVED_ID") = ""
    	    End If


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

	    	Call FnAddIISLOG("addlog=plogin&sub=pass2nd&uid="&userid) ''2016/12/29
		else
		    ''�α�����(����)
		    if AuthNo<>"" then
		    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",AuthNo,GeoIpCCD)
		    elseif tokenSn<>"" then
		        Call AddPartnerLoginLogWithGeoIpCode (userid,"N",tokenSn,GeoIpCCD)
		    else
		    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",RefCode,GeoIpCCD)
		    end if

	        Call FnAddIISLOG("addlog=plogin&sub=faillogin&uid="&userid) ''2016/12/29
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

		'// ��������
		Call FnAddIISLOG("addlog=plogin&sub=nouserid&uid="&userid) ''2016/12/29
	    response.write("<script>window.alert('������ Ȱ��ȭ ���� �ʾҰų�, ���̵� �Ǵ� ��й�ȣ�� Ʋ�Ƚ��ϴ�.');</script>")
	    response.write("<script>history.go(-1);</script>")
	    dbget.close()	:	response.End
	end if


	'### �α��� ���� ���к� ó�� ------------------------------------------------------

	dim ssnTmpUIDPartner
    dim isNoComplexPwTxt : isNoComplexPwTxt = chkPasswordComplex_Len6Ver(userid,userpass)

	''�����ӽ�
	dim cuseridv
	sql = "select top 1 * "
	sql = sql + " from [db_user].[dbo].tbl_user_c"
	sql = sql + " where userid = '" + userid + "'" + vbCrlf

	rsget.CursorLocation = adUseClient
    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		cuseridv = rsget("userdiv")
	end if
	rsget.close

	if (trim(UCase(dbpassword))=trim(UCase(Enc_userpass64))) AND (trim(UCase(dbpassword2))=trim(UCase(Enc_2userpass64)) ) then
	    if ((isFirstConnect) or (isRequirePwdUp) or (isNoComplexPwTxt<>"")) then

	        ssnTmpUIDPartner = session("ssBctId")

	        Session.Contents.RemoveAll()

	        Response.Cookies("partner").domain = "10x10.co.kr"
            Response.Cookies("partner") = ""
            Response.Cookies("partner").Expires = Date - 1

            session("ssnTmpUIDPartner")= ssnTmpUIDPartner   ''��� ����� ��� ���ǰ� (partner)

    	    ''// ���� �α��� ������ �н����� ����(2014/05/19 ������)
    	    if (isFirstConnect) then
    	        response.write "<script language='javascript'>" &vbCrLf &_
    							"	alert('���� �α��ν� ��й�ȣ�� �����ϼž� �մϴ�. \n��й�ȣ ������������ �̵��մϴ�.');" &vbCrLf &_
    							"	self.location='/login/modifyPassword_partner.asp';" &vbCrLf &_
    							"</script>"
    			dbget.close()	:	response.End
    	    end if

    	    ''// N(3)������ ��й�ȣ ������Ѱ�� ������������ �̵�
    	    if   (isRequirePwdUp) then
    	        response.write "<script language='javascript'>" &vbCrLf &_
    							"	alert('3���� �̻� ��й�ȣ�� �������� �����̽��ϴ�. \n��й�ȣ ������������ �̵��մϴ�.');" &vbCrLf &_
    							"	self.location='/login/modifyPassword_partner.asp';" &vbCrLf &_
    							"</script>"
    			dbget.close()	:	response.End
    	    end if


    		'// ��й�ȣ ��ȭ ��å ����(2008.12.12; ������)  'chkPasswordComplex => chkPasswordComplex_Len6Ver 2016/09/20
    		if chkPasswordComplex_Len6Ver(userid,userpass)<>"" then
    			response.write "<script language='javascript'>" &vbCrLf &_
    							"	alert('" & chkPasswordComplex_Len6Ver(userid,userpass) & "\n��й�ȣ ������������ �̵��մϴ�.');" &vbCrLf &_
    							"	self.location='/login/modifyPassword_partner.asp';" &vbCrLf &_
    							"</script>"
    			dbget.close()	:	response.End
    		end if
	    end if

	    response.Cookies("partner").domain = "10x10.co.kr"
	    response.Cookies("partner")("userid") = session("ssBctId")
	    response.Cookies("partner")("userdiv") = session("ssBctDiv")

		''���� �ӽ� (cuseridv="15") �� �̵� 2016/06/23
		if (cuseridv="14") then
		        session("ssUserCDiv")=cuseridv ''2016/08/11
				response.write "<script language='javascript'>location.replace('"&manageUrl&"/lectureadmin/index.asp')</script>"
	        	dbget.close()	:	response.End
		end if


	    if (session("ssBctId")="10x10") then
	        ''������.
	        session.Abandon
	        dbget.close()	:	response.End

	    ''����Level
	    elseif (session("ssBctDiv")<=9) then
	        ''������.
	        session.Abandon
	        response.write "<script language='javascript'>alert('����Ҽ� �����������Դϴ�.');location.replace('/')</script>"
	        dbget.close()	:	response.End

	    	if (errMsg<>"") then
	            response.write "<script language='javascript'>alert('" & errMsg & "');</script>"
	        end if

	    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/admin/index.asp')</script>"
	        dbget.close()	:	response.End
	    elseif (session("ssBctDiv")=999) then
	    	''���� ��ü (yahoo, empas..)
	        response.write "<script language='javascript'>location.replace('" & manageUrl & "/company/index.asp')</script>"
	        dbget.close()	:	response.End
	    elseif (session("ssBctDiv")=9999) then
	    	''�귣�� ��ü

		    ''// N(3)����������� ���� ������Ѱ�� ������������ �̵�
		    if (isRequireInfoUp) then
		   %>
	 		<script language='javascript'>
			 	alert('<%if datediff("d","2014-12-04",date())>90 then%>3���� �̻� ����������� �������� �����̽��ϴ�.<%else%>2015�� ���ظ� �����Ͽ� ����� ���� ������Ʈ�� ��û �帳�ϴ�.<%end if%> \n����� ���� ������������ �̵��մϴ�.');
			 	self.location='/login/modifyManagerInfo.asp'
			 </script>
			 <%
				dbget.close()	:	response.End
		    end if
	    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/partner/index.asp')</script>"
	        dbget.close()	:	response.End
	    elseif (session("ssBctDiv")=9000) then
	    	''���� ��ü
	    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/lectureradmin/index.asp')</script>"
	        dbget.close()	:	response.End
	    elseif (session("ssBctDiv")=501) or (session("ssBctDiv")=502) or (session("ssBctDiv")=503) or (session("ssBctDiv")=509) then

	    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/offshop/index.asp')</script>"
	        dbget.close()	:	response.End
	    elseif (session("ssBctDiv")=101) or (session("ssBctDiv")=111) or (session("ssBctDiv")=112) or (session("ssBctDiv")=201) or (session("ssBctDiv")=301) then

	    	response.write "<script language='javascript'>location.replace('" & manageUrl & "/admin/index.asp')</script>"
	        dbget.close()	:	response.End
	    else
	        session.Abandon
	        response.write "<script language='javascript'>alert('�����̾����ϴ�.');location.replace('/')</script>"
	        dbget.close()	:	response.End
	    end if
	else
	    session.Abandon
        response.write "<script language='javascript'>alert('�����̾����ϴ�.-2');location.replace('/')</script>"
        dbget.close()	:	response.End
	end if
else


	'### 1�� �α��� Ȯ��
	sql = "select top 1 A.id,   A.Enc_password, A.Enc_password64, A.groupid,  A.Enc_2password64 " & vbCrlf
	sql = sql & " from [db_partner].[dbo].tbl_partner as A " & vbCrlf
	sql = sql & " where A.id = '" & userid & "'" & vbCrlf
	sql = sql & " and A.isusing='Y'"
	sql = sql & " and A.userdiv>10"  ''2017/04/21 �߰� ������ �������� �Ұ�..

	rsget.CursorLocation = adUseClient
    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
	if  (not rsget.EOF)  then
	    db_id_Exists       = TRUE
	    db_Enc_password64  = rsget("Enc_password64")
	    db_Enc_2password64 = rsget("Enc_2password64")
	end if
	rsget.Close

	if (db_id_Exists) then
		'// �α��� ���� Ȯ��
        if (rtrim(UCase(db_Enc_password64))=trim(UCase(Enc_userpass64))) then
            ''dbpassword  = db_Enc_password64  '''?

            if isNull(db_Enc_2password64) or (db_Enc_2password64="") then
            	isdbpassword_sec= "N"
            else
            	isdbpassword_sec= "Y"
            end if

		  ''2���н����尡 ���ΰ�� �α��� ����. 2017/04/11
		    if (isdbpassword_sec="N") then
    		    if (Is2ndPwdNotExistsReject(userid)) then
                    Call FnAddIISLOG("addlog=plogin&sub=2ndpassnull&uid="&userid)
                    response.write("<script>window.alert('2����й�ȣ ������ �α������ֽñ� �ٶ��ϴ�.');</script>")
    		        response.write("<script>history.go(-1);</script>")
    		        dbget.close() : response.End
    		    end if
		    end if

		  ''��Ⱓ �α��� ���� ���
		    if (IsLongTimeNotLoginUserid(userid)) then
		        Call FnAddIISLOG("addlog=plogin&sub=logntimenosee&uid="&userid)
                response.write("<script>window.alert('��Ⱓ ������� �ʾ� ������ �����ϴ�.\n��й�ȣ ã�⸦ ���� ������ȣ ������ ������ Ȱ��ȭ ���� �ֽñ� �ٶ��ϴ�.');</script>")
		        response.write("<script>history.go(-1);</script>")
		        dbget.close() : response.End
		    end if


		    '�ӽ� ���� ����
		    session("tmpUID") =  userid
		    session("tmpUPWD") = userpass

		%>
			<form name="frmLogin" method="post" action="<%=getSCMSSLURL%>/login/loginS.asp">
				<input type="hidden" name="chkAuth" value="Y">
				<input type="hidden" name="hidSec" value="<%=isdbpassword_sec%>">
				<input type="hidden" name="saved_id" value="<%=saved_id%>">
			</form>
		<%
		      Call FnAddIISLOG("addlog=plogin&sub=pass1st&uid="&userid) ''2016/12/29
		      response.write("<script>document.frmLogin.submit();</script>")
		      dbget.close()	:	response.End
        else
            ''�α�����(����)
		    if AuthNo<>"" then
		    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",AuthNo,GeoIpCCD)
		    elseif tokenSn<>"" then
		        Call AddPartnerLoginLogWithGeoIpCode (userid,"N",tokenSn,GeoIpCCD)
		    else
		    	Call AddPartnerLoginLogWithGeoIpCode (userid,"N",RefCode,GeoIpCCD)
		    end if

	        Call FnAddIISLOG("addlog=plogin&sub=faillogin1st&uid="&userid) ''2016/12/29
	        response.write("<script>window.alert('������ Ȱ��ȭ ���� �ʾҰų�, ���̵� �Ǵ� ��й�ȣ�� Ʋ�Ƚ��ϴ�.');</script>")
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

		'// ��������
		Call FnAddIISLOG("addlog=plogin&sub=nouserid1st&uid="&userid) ''2016/12/29
	    response.write("<script>window.alert('������ Ȱ��ȭ ���� �ʾҰų�, ���̵� �Ǵ� ��й�ȣ�� Ʋ�Ƚ��ϴ�.');</script>")
	    response.write("<script>history.go(-1);</script>")
	    dbget.close()	:	response.End
	end if
	response.end
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
