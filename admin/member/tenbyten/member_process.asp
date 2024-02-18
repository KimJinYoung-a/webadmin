<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사원등록
' History : 2010.12.15 정윤정 생성
'           2018.04.10 한용민 수정(직원계정정보 체크후, 계정정보 매칭 변경 로직 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
dim mode, part_sn, posit_sn,  job_sn, isusing, department_id, rank_sn, olduserpass, newuserpass, frontid, bigo, emppass, personalmail
dim userid, empno, username,  birthday, issolar, sexflag, zipcode, zipaddr, useraddr, userphone, usercell, usermail, msnmail,messenger,juminno
dim interphoneno, extension, direct070, jobdetail, statediv, joinday, realjoinday, retireday,  userimage,  retirereason
dim i, sql, AssignedRow, mywork, adminid, changedate, logidx, gsshopuserid, cMember, frontusername, adminusername,existsUseridYN
dim strSql, chkname, userNameEN, isdispmember, frontUserHp, adminUserHp
adminid = session("ssBctId")
gsshopuserid = requestCheckvar(request("gsshopuserid"),32)
mode = requestCheckvar(request("mode"),32)
userid = trim(requestCheckvar(request("sUI"),32))
frontid = trim(requestCheckvar(request("sFUI"),32))
frontid = Replace(frontid, " ", "")
part_sn = requestCheckvar(request("selPN"),10)
posit_sn = requestCheckvar(request("selPoN"),10)
job_sn = requestCheckvar(request("selJN"),10)
empno = requestCheckvar(request("sEN"),14)
userNameEN = requestCheckvar(request("userNameEN"),32)
username = requestCheckvar(request("sUN"),32)
juminno	=requestCheckvar(request("sJN1"),6)&"-"&requestCheckvar(request("sJN2"),7)
birthday = Num2Str(getNumeric(requestCheckvar(request("selBD_y"),4)),4,"0","R")&"-"&_
		Num2Str(getNumeric(requestCheckvar(request("selBD_m"),2)),2,"0","R")&"-"&_
		Num2Str(getNumeric(requestCheckvar(request("selBD_d"),2)),2,"0","R")
issolar = requestCheckvar(request("rdoS"),1)
sexflag= requestCheckvar(request("rdoSf"),1)

zipcode = requestCheckvar(request("zipcode"),32)
zipaddr = requestCheckvar(request("zipaddr"),128)
useraddr = requestCheckvar(request("useraddr"),128)
userphone = requestCheckvar(request("sUP"),16)
usercell = requestCheckvar(request("sUC"),16)
usermail = requestCheckvar(request("sUM"),128)
personalmail = requestCheckvar(request("sPM"),128)
msnmail = requestCheckvar(request("sMM"),128)
messenger = requestCheckvar(request("sNt"),128)
interphoneno = requestCheckvar(request("sCUP"),16)
extension = requestCheckvar(request("sCE"),4)
direct070 = requestCheckvar(request("sD070"),16)
if (part_sn = "13" or part_sn ="4" or part_sn ="5" or part_sn="6") then
	jobdetail = requestCheckvar(request("selO"),10)
else
	jobdetail = requestCheckvar(request("selC"),10)
end if

mywork = requestCheckvar(request("smywork"),150)

statediv = requestCheckvar(request("statediv"),1)
IF statediv = "" THEN statediv = "Y"
joinday = requestCheckvar(request("selJD_y"),4)&"-"&requestCheckvar(request("selJD_m"),2)&"-"&requestCheckvar(request("selJD_d"),2)
IF requestCheckvar(request("selRJD_y"),4) <> "" THEN
realjoinday= requestCheckvar(request("selRJD_y"),4)&"-"&requestCheckvar(request("selRJD_m"),2)&"-"&requestCheckvar(request("selRJD_d"),2)
END IF
retireday = requestCheckvar(request("selRD_y"),4)
IF retireday <> "" AND requestCheckvar(request("selRD_m"),2) <> "" THEN
	retireday = retireday &"-"&requestCheckvar(request("selRD_m"),2)
END IF
IF retireday <> "" AND requestCheckvar(request("selRD_d"),2) <> "" THEN
	retireday = retireday &"-"&requestCheckvar(request("selRD_d"),2)
END IF
retirereason	= requestCheckvar(request("retirereason"),4)

userimage=requestCheckvar(request("sUImg"),100)

emppass=requestCheckvar(request("sEP"),32)

department_id=requestCheckvar(request("department_id"),20)
if (department_id = "") then
	department_id = "NULL"
end if

rank_sn = requestCheckvar(request("selRank"),2)

changedate = requestCheckvar(request("chDate"),10)
logidx= requestCheckvar(request("logidx"),10)
existsUseridYN="N"

dim strMsg
dim objCmd, returnValue

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

'// 처리 분기 //
Select Case mode
Case "A"
	' 계약직관리에서도 사용해서 권한 체크 안함.
	'if not(isdispmember) then
	'	response.write "인사팀이거나 개인정보 접근권한 변경자만 접속 가능한 매뉴 입니다."
	'	response.end
	'end if

	'' '// 입사일만 입력하고 실제입사일이 입력되지 않으면 실제 입사일 입력(2013-07-08 skyer9)
	 if (joinday <> "") and (realjoinday = "") then
	 	realjoinday = joinday
	 end if

    if  emppass <> "" then
	    if chkPasswordComplexNonId(emppass)<>"" then
	    	response.write "<script language='javascript'>" &vbCrLf &_
	    					"	alert('" & chkPasswordComplexNonId(emppass) & "\n(사번로그인용) 비밀번호를 확인후 다시 시도해주세요.');" &vbCrLf &_
	    					" 	history.back();" &vbCrLf &_
	    					"</script>"
	    	dbget.close()	:	response.End
	    end if
	     emppass = md5(trim(emppass))
    end if

	strMsg = "사원정보가 등록되었습니다."

	'response.write "'"&username&"','"&emppass&"','"&juminno&"','"&birthday&"','"&issolar&"','"&sexflag&"','"&zipcode&"','"&zipaddr&"','"&useraddr&"','"&userphone&"','"&usercell&"','"&usermail&"','"&msnmail&"','"&interphoneno&"','"&extension&"','"&direct070&"','"&jobdetail&"','"&joinday&"','"&realjoinday&"','"&userimage&"','"&part_sn&"','"&posit_sn&"','"&job_sn&"','"&messenger&"', "&department_id&", '"&rank_sn&"','"&adminid&"','"& personalmail &"'"
	'response.end
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_partner].[dbo].sp_Ten_user_tenbyten_insert('"&username&"','"&emppass&"','"&juminno&"','"&birthday&"','"&issolar&"'"&_
					 ",'"&sexflag&"','"&zipcode&"','"&zipaddr&"','"&useraddr&"','"&userphone&"','"&usercell&"','"&usermail&"','"&msnmail&"','"&interphoneno&"'"&_
					 ",'"&extension&"','"&direct070&"','"&jobdetail&"','"&joinday&"','"&realjoinday&"','"&userimage&"','"&part_sn&"','"&posit_sn&"','"&job_sn&"'"&_
					 ",'"&messenger&"', "&department_id&", '"&rank_sn&"','"&adminid&"','"& personalmail &"','"& html2db(trim(gsshopuserid)) &"','"& frontid &"'"&_
					 ",'"& userNameEN &"')}"

		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
	Set objCmd = Nothing

	existsUseridYN="N"
	if not(frontid="" or isnull(frontid)) then
		sql = "select userid"
		sql = sql & " from db_user.dbo.tbl_user_n u with (nolock)"
		sql = sql & " where u.userid='"& frontid &"'"
		sql = sql & " 	and left(username,iif(len('" & username & "')>=3,3,2)) = left('" & username & "',iif(len('" & username & "')>=3,3,2)) "

		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.EOF  then
			existsUseridYN="Y"
		end if
		rsget.Close

		if existsUseridYN="N" then
			strMsg = strMsg & "\n\n입력하신 아이디가 프론트 텐바이텐 사이트에 존재하지 않거나 프론트의 이름이 달라 제외되고 저장 됩니다."
		end if
	end if

	if (returnValue = "1") then
			response.write	"<script  type='text/javascript'>" &_
							"	alert('" & strMsg & "');" &_
						 	"	opener.location.reload();" &_
						 	"	self.close();" &_
							"</script>"
	Else
		response.write	"<script  type='text/javascript'>" &_
						"	alert('처리중 에러가 발생했습니다.');" &_
						"	history.back();" &_
						"</script>"

	End If
	response.end

Case "U"
	' 계약직관리에서도 사용해서 권한 체크 안함.
	'if not(isdispmember) then
	'	response.write "인사팀이거나 개인정보 접근권한 변경자만 접속 가능한 매뉴 입니다."
	'	response.end
	'end if

	if (juminno = "-") then
		juminno = ""
	end if
	 if  emppass <> "" then
	    if chkPasswordComplexNonId(emppass)<>"" then
	    	response.write "<script language='javascript'>" &vbCrLf &_
	    					"	alert('" & chkPasswordComplexNonId(emppass) & "\n(사번로그인용) 비밀번호를 확인후 다시 시도해주세요.');" &vbCrLf &_
	    					" 	history.back();" &vbCrLf &_
	    					"</script>"
	    	dbget.close()	:	response.End
	    end if
	     emppass = md5(trim(emppass))
    end if
	strMsg = "사원정보가 수정되었습니다."
	IF retirereason = "" THEN retirereason = 1
		Dim blnAuth
		blnAuth = "N"
		IF C_PSMngPart or C_ADMIN_AUTH THEN blnAuth ="Y"

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_partner].[dbo].sp_Ten_user_tenbyten_update('"&empno&"', '"&username&"','"&emppass&"','"&juminno&"','"&birthday&"'"&_
					 ",'"&issolar&"','"&sexflag&"','"&zipcode&"','"&zipaddr&"','"&useraddr&"','"&userphone&"','"&usercell&"','"&usermail&"','"&msnmail&"'"&_
					 ",'"&interphoneno&"','"&extension&"','"&direct070&"','"&jobdetail&"','"&joinday&"','"&mywork&"','"&statediv&"','"&realjoinday&"'"&_
					 ",'"&retireday&"','"&userimage&"','"&part_sn&"','"&posit_sn&"','"&job_sn&"',"&retirereason&",'"&blnAuth&"','"&messenger&"'"&_
					 ",'"&userid&"',"&department_id&",'"&rank_sn&"','"&changedate&"','"&adminid&"','"& personalmail &"','"& html2db(trim(gsshopuserid)) &"'"&_
					 ",'"& frontid &"', '"& userNameEN &"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)

			.Execute, , adExecuteNoRecords
		End With
		returnValue = objCmd(0).Value
	Set objCmd = Nothing

  IF returnValue = 2  THEN
  		dbget.Close
  	%>
			<script  type="text/javascript">
				alert("계약직사원의 퇴사일이나 퇴사예정일은 월별급여를 마감한 시점에서는 \n등록/수정이 불가능합니다. 인사교육파트에 문의해주세요");
				history.back();
			</script>
	<%
		response.end
	END IF

	existsUseridYN="N"
	if not(frontid="" or isnull(frontid)) then
		sql = "select userid"
		sql = sql & " from db_user.dbo.tbl_user_n u with (nolock)"
		sql = sql & " where u.userid='"& frontid &"'"
		sql = sql & " 	and left(username,iif(len('" & username & "')>=3,3,2)) = left('" & username & "',iif(len('" & username & "')>=3,3,2)) "

		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.EOF  then
			existsUseridYN="Y"
		end if
		rsget.Close

		if existsUseridYN="N" then
			strMsg = strMsg & "\n\n입력하신 아이디가 프론트 텐바이텐 사이트에 존재하지 않거나 프론트의 이름이 달라 제외되고 저장 됩니다."
		end if
	end if

Case "D"
	' 계약직관리에서도 사용해서 권한 체크 안함.
	'if not(isdispmember) then
	'	response.write "인사팀이거나 개인정보 접근권한 변경자만 접속 가능한 매뉴 입니다."
	'	response.end
	'end if

		strMsg = "사원정보가 삭제되었습니다."
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_partner].[dbo].sp_Ten_user_tenbyten_delete('"&empno&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
		End With
		 returnValue = objCmd(0).Value

		 IF returnValue > 0 THEN
		 Set objCmd = Server.CreateObject("ADODB.COMMAND")
		 With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_partner].[dbo].sp_Ten_user_defaultpay_SetEnddate('"&empno&"','"&retireday&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
		End With
		returnValue = objCmd(0).Value
		END IF

Case "R"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call db_partner.dbo.sp_Ten_user_tenbyten_IDCheck('"&empno&"','"&userid&"','"&username&"')}"

		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
	Set objCmd = Nothing

	IF (returnValue ="1" ) THEN
%>
		<script langauge="javascript">
			opener.document.frm_member.hidID.value = "0";
			if(confirm("기존에 사용중인 아이디입니다. 기존사번에서의 아이디 사용을 중지하고 현재 사번에 등록하시겠습니까?")){
			opener.document.frm_member.hidID.value = "1";
			}
			self.close();
		</script>
<%
	ELSEIF (returnValue ="2") THEN
%>
		<script  type="text/javascript">
			alert("다른 사원명의로 사용중인 아이디입니다. 사용 불가능합니다.\n\n다른 아이디를 사용해주세요");
			opener.document.frm_member.hidID.value = "0";
			self.close();
		</script>
<%
	ELSE
%>
		<script  type="text/javascript">
			alert("사용가능한 아이디입니다.");
			opener.document.frm_member.hidID.value = "1";
			self.close();
		</script>
<%
	END IF
	dbget.Close
	response.end

' 직원명동일체크[어드민,프론트사이트]	' 2021.11.16 한용민 생성
Case "frontnamewebadmincheck"
	if frontid="" or isnull(frontid) then
		response.write "<script  type='text/javascript'>"
		response.write "	alert('프론트 텐바이텐사이트 아이디가 없습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if
	if empno="" or isnull(empno) then
		response.write "<script  type='text/javascript'>"
		response.write "	alert('사번이 없습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	set cMember = new CTenByTenMember
		cMember.Fempno = empno
		cMember.fnGetMemberData
		adminusername = cMember.fusername
	set cMember=nothing

	if adminusername="" or isnull(adminusername) then
		response.write "<script  type='text/javascript'>"
		response.write "	alert('어드민에 성함이 등록되어 있지 않습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	sql = "select username"
	sql = sql & " from db_user.dbo.tbl_user_n u with (nolock)"
	sql = sql & " where u.userid='"& frontid &"'"

	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF  then
		frontusername = rsget("username")
	end if
	rsget.Close

	if frontusername="" or isnull(frontusername) then
		response.write "<script  type='text/javascript'>"
		response.write "	alert('프론트 텐바이텐 사이트에 성함이 등록되어 있지 않습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
	.ActiveConnection = dbget
	.CommandType = adCmdText
	.CommandText = "{?= call [db_partner].[dbo].sp_Ten_user_tenbyten_chkFrontId('"& frontid &"','"& adminusername &"')}"
	.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
	.Execute, , adExecuteNoRecords
	End With
	returnValue = objCmd(0).Value
	Set objCmd = Nothing

	IF  returnValue = 0 THEN
		response.write "<script  type='text/javascript'>"
		response.write "	alert('텐바이텐아이디의 이름("& frontusername &")과 WEBADMIN 어드민 이름("& adminusername &")이 동일하지 않습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	END IF

	response.write "<script type='text/javascript'>"
	response.write "	alert('어드민과 텐바이텐 사이트의 이름이 일치 합니다.');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close() : response.end

' 퇴사이전으로 변경
Case "changenotretire"
	If not(C_ADMIN_AUTH or C_PSMngPart) Then
		response.write "<script  type='text/javascript'>"
		response.write "	alert('권한이 없습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if empno="" or isnull(empno) then
		response.write "<script  type='text/javascript'>"
		response.write "	alert('사번이 없습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	sql = "select statediv,retireday,retirereason"
	sql = sql & " from db_partner.dbo.tbl_user_tenbyten with (nolock)"
	sql = sql & " where empno='"& empno &"'"

	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF  then
		statediv = rsget("statediv")
		retireday = rsget("retireday")
		retirereason = rsget("retirereason")
	end if
	rsget.Close

	if not(statediv = "N" and retirereason <> "99") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('현재 상태가 퇴사상태가 아닙니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	sql = "update db_partner.dbo.tbl_user_tenbyten" & vbcrlf
	sql = sql & " set isusing=1" & vbcrlf
	sql = sql & " , statediv='Y'" & vbcrlf
	sql = sql & " , retireday = NULL" & vbcrlf
	sql = sql & " , retirereason = NULL where" & vbcrlf
	sql = sql & " empno = '"& empno &"'"

	'response.write sql & "<BR>"
	dbget.execute sql

	response.write "<script type='text/javascript'>"
	response.write "	alert('퇴사이전으로 변경되었습니다.');"
	response.write "	opener.location.reload();"
	response.write "	opener.focus();"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close() : response.end

' 직원명강제변경[어드민->프론트사이트]	' 2021.11.16 한용민 생성
Case "frontnamewebadminchange"
	If not(C_ADMIN_AUTH or C_PSMngPart) Then
		response.write "<script  type='text/javascript'>"
		response.write "	alert('권한이 없습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if frontid="" or isnull(frontid) then
		response.write "<script  type='text/javascript'>"
		response.write "	alert('프론트 텐바이텐사이트 아이디가 없습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if
	if empno="" or isnull(empno) then
		response.write "<script  type='text/javascript'>"
		response.write "	alert('사번이 없습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	set cMember = new CTenByTenMember
		cMember.Fempno = empno
		cMember.fnGetMemberData
		adminusername = cMember.fusername
		adminUserHp = replace(trim(cMember.Fusercell),"-","")
	set cMember=nothing

	sql = "select username, replace(isnull(usercell,''),'-','') as frontUserHp"
	sql = sql & " from db_user.dbo.tbl_user_n u with (nolock)"
	sql = sql & " where u.userid='"& frontid &"'"

	'response.write sql & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF  then
		frontusername = trim(rsget("username"))
		frontUserHp = trim(rsget("frontUserHp"))
	end if
	rsget.Close

	if adminUserHp<>frontUserHp then
		response.write "<script type='text/javascript'>"
		response.write "	alert('텐바이텐 사이트의 휴대폰번호("&frontUserHp&")와 직원정보의 휴대폰번호("&adminUserHp&")가 틀립니다.\n동일한 사람이 맞는지 확인하세요.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
	.ActiveConnection = dbget
	.CommandType = adCmdText
	.CommandText = "{?= call [db_partner].[dbo].sp_Ten_user_tenbyten_chkFrontId('"& frontid &"','"& adminusername &"')}"
	.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
	.Execute, , adExecuteNoRecords
	End With
	returnValue = objCmd(0).Value
	Set objCmd = Nothing

	IF returnValue = 0 THEN
		sql="update db_user.dbo.tbl_user_n set username='"& adminusername &"' where userid='"& frontid &"'"

		'response.write sql & "<BR>"
		dbget.execute sql

		response.write "<script type='text/javascript'>"
		response.write "	alert('텐바이텐 사이트의 직원명이 변경되었습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	else
		response.write "<script  type='text/javascript'>"
		response.write "	alert('변경불가.\n텐바이텐아이디의 이름("& frontusername &")과 WEBADMIN 어드민 이름("& adminusername &")이 동일합니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	END IF

Case "S"
	strMsg = "사원정보가 삭제되었습니다."
	sql = ""
	sql = sql & " UPDATE db_partner.dbo.tbl_user_tenbyten set "
	sql = sql & " isusing = 0 "
	sql = sql & " WHERE empno = '"&empno&"' "
	dbget.Execute sql, AssignedRow
	If AssignedRow > 0 Then
		returnValue = 1
	End If

' 프론트 직원아이디 퇴사처리. 쿠폰회수,회원등급변경	' 2018.04.17 한용민 생성
Case "RetireUser"
	If not(C_ADMIN_AUTH or C_PSMngPart) Then
		response.write "<script  type='text/javascript'>"
		response.write "	alert('권한이 없습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if

	strMsg = "처리 되었습니다."
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call db_partner.[dbo].[sp_Ten_user_tenbyten_RetireUser_Update]('"&userid&"','"&frontid&"','"&username&"')}"

		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
	Set objCmd = Nothing

	response.write	"<script  type='text/javascript'>" &_
					"	alert('처리완료되었습니다.');" &_
					"	opener.location.reload();" &_
					"	opener.focus();" &_
					"	self.close();" &_
					"</script>"
	dbget.close() : response.end
Case "LD" '//발령취소

		strMsg = "발령이 취소되었습니다."
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_partner].[dbo].usp_Ten_user_tenbyten_DelModLog('"&empno&"','"&logidx&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
		End With
		 returnValue = objCmd(0).Value
case "C" '이름체크(추후 삭제가능)
 	username =requestCheckvar(request("hidName"),32)
 	chkname = 0
 	strSql = " select username from db_partner.dbo.tbl_user_tenbyten where username = '"&username&"' and isusing=1"
 	 rsget.Open strSql, dbget, 1
 	 if not rsget.eof Then
 	 	 if not (isNull( rsget("username")) or  rsget("username") ="") then
 	 	 chkname = 1
 		end if
 	 end if
 	 rsget.close

 	 if chkname =1 then
 	 %>
		<script  type="text/javascript">
			alert("동일한 이름의 사용자가 있습니다. 사용불가능합니다.");
			parent.document.frm_member.hidNm.value = "0";
			location.href = "about:blank;";
		</script>
<%

	ELSE
%>
		<script  type="text/javascript">
			alert("사용가능한 이름입니다.");
			parent.document.frm_member.hidNm.value = "1";
			location.href = "about:blank;";
		</script>
<% end if
 	response.end

case "checkNameEN"		'영문이름체크(추후 삭제가능)
 	username =requestCheckvar(request("hiduserNameEN"),32)
 	chkname = 0
 	strSql = " select userNameEN from db_partner.dbo.tbl_user_tenbyten where userNameEN = '"&username&"' and isusing=1"
 	 rsget.Open strSql, dbget, 1
 	 if not rsget.eof Then
 	 	 if not (isNull( rsget("userNameEN")) or  rsget("userNameEN") ="") then
 	 	 chkname = 1
 		end if
 	 end if
 	 rsget.close

 	 if chkname =1 then
 	 %>
		<script  type="text/javascript">
			alert("동일한 영문이름의 사용자가 있습니다. 사용불가능합니다.");
			parent.document.frm_member.hiduserNameEN.value = "0";
			location.href = "about:blank;";
		</script>
<%

	ELSE
%>
		<script  type="text/javascript">
			alert("사용가능한 이름입니다.");
			parent.document.frm_member.hiduserNameEN.value = "1";
			location.href = "about:blank;";
		</script>
<% end if
 	response.end
End Select
	if (returnValue = "1") then
			response.write	"<script  type='text/javascript'>" &_
							"	alert('" & strMsg & "');" &_
							"location.href ='/admin/member/tenbyten/pop_member_modify.asp?sEPN="&empno&"'"&_
						 	"	//opener.location.reload();" &_
						 	"	//opener.focus();" &_
						 	"	//self.close();" &_
							"</script>"
	Else
		response.write	"<script  type='text/javascript'>" &_
						"	alert('처리중 에러가 발생했습니다.');" &_
						"	history.back();" &_
						"</script>"

	End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
