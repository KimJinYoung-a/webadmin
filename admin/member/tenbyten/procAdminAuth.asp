<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사원권한등록
' History : 정윤정 생성
'			2022.05.09 한용민 수정(ISMS개인정보취급권한 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
dim empno, userid, frontid, username, userpass, part_sn, posit_sn, level_sn, job_sn, isusing, userdiv, criticinfouser, userpass64
dim objCmd,returnValue, isdispmember
dim adminid, lv1customerYN, lv2partnerYN, lv3InternalYN
lv1customerYN 	= requestCheckvar(request("lv1customerYN"),1)
lv2partnerYN 	= requestCheckvar(request("lv2partnerYN"),1)
lv3InternalYN 	= requestCheckvar(request("lv3InternalYN"),1)
empno = requestCheckvar(request("sEN"),14)
userid = trim(requestCheckvar(request("sUI"),32))
frontid = requestCheckvar(request("sFUI"),32)
userpass = requestCheckvar(request("sP"),32)

userdiv = requestCheckvar(request("selUD"),10)
level_sn = requestCheckvar(request("selLN"),10)
part_sn = requestCheckvar(request("selPN"),10)
posit_sn = requestCheckvar(request("selPoN"),10)
job_sn = requestCheckvar(request("selJN"),10)
username = requestCheckvar(request("sUN"),32)
userid = Replace(userid, " ", "")
frontid = Replace(frontid, " ", "")
username = Replace(username, " ", "")
adminid = session("ssBctId")
 
''response.write "aaa" & requestCheckvar(request("selPN"),10)
''dbget.close
''response.end

''2014/07/14
criticinfouser = requestCheckvar(request("criticinfouser"),10)
criticinfouser = CHKIIF(criticinfouser="","0",criticinfouser)
if lv1customerYN="" or isnull(lv1customerYN) then lv1customerYN="N"
if lv2partnerYN="" or isnull(lv2partnerYN) then lv2partnerYN="N"
if lv3InternalYN="" or isnull(lv3InternalYN) then lv3InternalYN="N"

IF application("Svr_Info")="Dev" THEN
	isdispmember = true
else
	' ISMS 심사로 인해 개인정보 접근권한 생성/수정/변경 특정사람만 보이게(한용민,허진원,이문재)	' 2020.10.12 한용민
	if C_privacyadminuser or C_PSMngPart then
		isdispmember = true
	else
		isdispmember = false
	end if
end if

if not(isdispmember) then
	response.write "인사팀이거나 개인정보 접근권한 변경자만 접속 가능한 매뉴 입니다."
	response.end
end if

if (posit_sn = "") then
	posit_sn = "0"
end if
	' 부서가 권한이 있는 경우에만
	if part_sn<>"35" and part_sn<>"" then
		IF frontid <> "" THEN
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_partner].[dbo].sp_Ten_user_tenbyten_chkFrontId('"&frontid&"','"&username&"')}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
		Set objCmd = Nothing

		IF  returnValue = 0 THEN
			dbget.Close
		%>
			<script type="text/javascript">
				alert("텐바이텐아이디의 이름과 WEBADMIN 어드민 이름이 동일하지 않습니다.\n텐바이텐 아이디를 확인 후 다시 입력해주세요.");
				history.back();
			</script>
	<%		response.end
		END IF
		END IF
	END IF

		'//패스워드 정책 검사
	if userid <> "" and userpass <> "" then
		if chkPasswordComplex(userid,userpass)<>"" then
	    	response.write "<script language='javascript'>" &vbCrLf &_
	    					"	alert('" & chkPasswordComplex(userid,userpass) & "\nSCM 패스워드를 확인후 다시 시도해주세요.');" &vbCrLf &_
	    					" 	history.back();" &vbCrLf &_
	    					"</script>"
	    	dbget.close()	:	response.End
	    end if
	     userpass = md5(trim(userpass))
	     userpass64 = sha256((trim(userpass)))
    end if

    Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_partner].[dbo].sp_Ten_user_tenbyten_AdminAuth('"&empno&"','"&userid&"','"&frontid&"','"&username&"','"&userpass&"','"&userpass64&"','"&level_sn&"','"&part_sn&"','"&posit_sn&"','"&job_sn&"','"&userdiv&"',"&criticinfouser&",'"&adminid&"','"&lv1customerYN&"','"&lv2partnerYN&"','"&lv3InternalYN&"')}"

		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
	Set objCmd = Nothing

''' 추가 권한 관련..

    dim StrSQL, lp1

    '기존 자료 정리
    strSQL = "Delete From db_partner.dbo.tbl_partner_AddLevel Where userID='" & userid &"'"
    dbget.Execute(strSQL)

    strSQL = "Insert into db_partner.dbo.tbl_partner_AddLevel"
    strSQL = strSQL & " (userid,part_sn,level_sn,isDefault)"
    strSQL = strSQL & " select p.id, t.part_sn, p.level_sn,'Y'"
    strSQL = strSQL & " from db_partner.dbo.tbl_partner p"
    strSQL = strSQL & " 	Join db_partner.dbo.tbl_user_tenbyten t"
    strSQL = strSQL & " 	on p.id=t.userid"
    strSQL = strSQL & " where p.id='" & userid &"'"
    dbget.Execute(strSQL)


    dim ARRpart_sn : ARRpart_sn		= request.Form("part_sn")
    dim ARRlevel_sn : ARRlevel_sn	= request.Form("level_sn")
    dim splPsn, splLsn

    '수정자료 저장
    if ARRpart_sn<>"" then splPsn = Split(ARRpart_sn, ",")
    if ARRlevel_sn<>"" then splLsn = Split(ARRlevel_sn, ",")

    If IsArray(splPsn) Then
        For lp1=0 to Ubound(splPsn)
        	IF Trim(splPsn(lp1))<>"" and Trim(splLsn(lp1))<>"" THEN
        	strSQL =	"Insert into db_partner.dbo.tbl_partner_AddLevel (userid,part_sn,level_sn,isDefault)"
        	strSQL = strSQL & "Values ('" & userid & "', " & splPsn(lp1) & ", " & splLsn(lp1) & ",'N')"

        	dbget.Execute(strSQL)
        	END IF
        Next
    end if


if (returnValue = "1") then
			response.write	"<script  type='text/javascript'>" &_
							"	alert('처리완료되었습니다.');" &_
							"	opener.location.reload();" &_
							"	opener.focus();" &_
							"	self.close();" &_
							"</script>"
	Else
		response.write	"<script  type='text/javascript'>" &_
						"	alert('처리중 에러가 발생했습니다.');" &_
						"	history.back();" &_
						"</script>"

	End If
%>
