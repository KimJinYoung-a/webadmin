<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'           2023.10.30 한용민 수정(휴면계정정보표기. 휴면계정->일반계정 전환 로직 생성)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<%
dim userid, userseq, mail10x10, mailfinger, offlinemail, sms10x10, smsfinger, offlinesms, haveonlineaccount, haveofflineaccount
dim issameusercell, issameusermail, issameuserphone, mode, usercell, usermail, userphone, contents, complaintext
dim Enc_userpass, Enc_userpass64, sqlStr, refer
	mode 				= requestcheckvar(request("mode"),32)
	userid 				= requestcheckvar(request("userid"),32)
	userseq 			= request("userseq")
	haveonlineaccount 	= request("haveonlineaccount")
	haveofflineaccount 	= request("haveofflineaccount")
	issameusercell 		= request("issameusercell")
	issameuserphone 		= request("issameuserphone")
	issameusermail 		= request("issameusermail")
	mail10x10 = request("mail10x10")
	mailfinger = request("mailfinger")
	offlinemail = request("offlinemail")
	sms10x10 = request("sms10x10")
	smsfinger = request("smsfinger")
	offlinesms = request("offlinesms")
	complaintext = requestcheckvar(request("complaintext"),4000)
	refer = request.ServerVariables("HTTP_REFERER")

if (mode = "delonusercell") or ((mode = "deloffusercell") and (issameusercell = "Y")) then

	'// CS 메모 저장
	sqlStr = " select top 1 "
	sqlStr = sqlStr + " 	usercell, IsNull(usermail, '') as usermail "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_user].[dbo].tbl_user_n "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	userid = '" & userid & "' "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	usercell = ""
	usermail = ""
	if  not rsget.EOF  then
		usercell = rsget("usercell")
		usermail = rsget("usermail")
	end if
	rsget.close

	contents = "다른 고객의 요청으로 기존 핸드폰 번호(" + usercell + ")를 000-000-0000 로 변경"

    Call AddCsMemo("", "1", userid, session("ssBctId"), contents)
    response.write "<script>alert('CS MEMO에 저장되었습니다.')</script>"

' 일반전화번호 삭제(온라인과 오프라인이 같은 번호 일경우) 	' 2019.04.10 한용민 생성
elseif (mode = "delonuserphone") or ((mode = "delonuserphone") and (issameuserphone = "Y")) then
	'// CS 메모 저장
	sqlStr = " select top 1"
	sqlStr = sqlStr + " userphone, IsNull(usermail, '') as usermail"
	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_n"
	sqlStr = sqlStr + " where userid = '" & userid & "'"

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	userphone = ""
	usermail = ""
	if  not rsget.EOF  then
		userphone = rsget("userphone")
		usermail = rsget("usermail")
	end if
	rsget.close

	contents = "다른 고객의 요청으로 기존 전화번호(" + userphone + ")를 000-000-0000 로 변경"

    Call AddCsMemo("", "1", userid, session("ssBctId"), contents)
    response.write "<script>alert('CS MEMO에 변경내역이 저장되었습니다.')</script>"

elseif (mode = "delonusermail") then

	'// CS 메모 저장
	sqlStr = " select top 1 "
	sqlStr = sqlStr + " 	usercell, IsNull(usermail, '') as usermail "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_user].[dbo].tbl_user_n "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	userid = '" & userid & "' "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	usercell = ""
	usermail = ""
	if  not rsget.EOF  then
		usercell = rsget("usercell")
		usermail = rsget("usermail")
	end if
	rsget.close

	contents = "다른 고객의 요청으로 기존 이메일주소(" + usermail + ")를 삭제"

    Call AddCsMemo("", "1", userid, session("ssBctId"), contents)
    response.write "<script>alert('CS MEMO에 저장되었습니다.')</script>"

elseif (mode = "resetUserPass") then

	'' 임시비밀번호 생성 프로세스 변경 (기존 0000 -> 신규 6자리난수), 2015-07-15, skyer9
	'' 별도 고객안내 필요

	dim strRdm
	strRdm = RandomStr()
	call setTempPassword(userid,strRdm)

	contents = "고객의 요청으로 비밀번호를 " & strRdm & " 으로 임시비밀번호 생성"

    Call AddCsMemo("", "2", userid, session("ssBctId"), contents)
    response.write "<script>alert('CS MEMO에 저장되었습니다.')</script>"

' 온라인고객탈퇴처리	' 2021.11.23 한용민
elseif (mode = "delonuser") then
	if not(C_CSPowerUser or C_ADMIN_AUTH) then
		response.write "탈퇴 처리는 cs팀 파트장 이상 권한이 필요 합니다."
		response.write "<script type='text/javascript'>"
		response.write "	alert('탈퇴 처리는 cs팀 파트장 이상 권한이 필요 합니다..');"
		response.write "</script>"
		dbget.close() : response.end
	end if

	complaintext=trim(complaintext)
	if complaintext <> "" then
		if checkNotValidHTML(complaintext) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			dbget.close() : response.End
		end if
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('탈퇴 사유를 입력해주세요');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close() : response.End
	end If

    Call AddCsMemo("", "1", userid, session("ssBctId"), complaintext)
    response.write "<script type='text/javascript'>alert('CS MEMO에 저장되었습니다.')</script>"
end if

if (mode = "modifyuserinfo") then

	if (haveonlineaccount = "Y") then
		sqlStr = " update [db_user].[dbo].tbl_user_n " & VbCrlf
		sqlStr = sqlStr & " set email_10x10 = '" & mail10x10 & "' " & VbCrlf
		sqlStr = sqlStr & " ,email_way2way = '" & mailfinger & "' " & VbCrlf

		if (mail10x10 = "Y" or mailfinger = "Y") then
			sqlStr = sqlStr & " ,emailok = 'Y' " & VbCrlf
		else
			sqlStr = sqlStr & " ,emailok = 'N' " & VbCrlf
		end if

		sqlStr = sqlStr & " ,smsok = '" & sms10x10 & "' " & VbCrlf
		sqlStr = sqlStr & " ,smsok_fingers = '" & smsfinger & "' " & VbCrlf
		sqlStr = sqlStr & " where userid = '" & userid & "' "
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	end if

	if (haveofflineaccount = "Y") then
		sqlStr = " update db_shop.dbo.tbl_total_shop_user " & VbCrlf
		sqlStr = sqlStr & " set emailyn = '" & offlinemail & "' " & VbCrlf
		sqlStr = sqlStr & " ,smsyn = '" & offlinesms & "' " & VbCrlf
		sqlStr = sqlStr & " where userseq = " & CStr(userseq) & " " & VbCrlf
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	end if

elseif (mode = "delonusercell") then

	if (haveonlineaccount = "Y") then
		sqlStr = " update [db_user].[dbo].tbl_user_n " & VbCrlf
		sqlStr = sqlStr & " set usercell = '000-000-0000' " & VbCrlf
		sqlStr = sqlStr & " where userid = '" & userid & "' "
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	end if

	if (haveofflineaccount = "Y") and (issameusercell = "Y") then
		sqlStr = " update db_shop.dbo.tbl_total_shop_user " & VbCrlf
		sqlStr = sqlStr & " set hpno = '000-000-0000' " & VbCrlf
		sqlStr = sqlStr & " where userseq = " & CStr(userseq) & " " & VbCrlf
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	end if

' 일반전화번호 삭제 	' 2019.04.10 한용민 생성
elseif (mode = "delonuserphone") then
	if (haveonlineaccount = "Y") then
		sqlStr = " update [db_user].[dbo].tbl_user_n " & VbCrlf
		sqlStr = sqlStr & " set userphone = '000-000-0000' " & VbCrlf
		sqlStr = sqlStr & " where userid = '" & userid & "' "

		'response.write sqlStr
		dbget.execute sqlStr
	end if

	if (haveofflineaccount = "Y") and (issameuserphone = "Y") then
		sqlStr = " update db_shop.dbo.tbl_total_shop_user " & VbCrlf
		sqlStr = sqlStr & " set telno = '000-000-0000' " & VbCrlf
		sqlStr = sqlStr & " where userseq = " & CStr(userseq) & " " & VbCrlf

		'response.write sqlStr
		dbget.execute sqlStr
	end if

elseif (mode = "delonusermail") then

	if (haveonlineaccount = "Y") then
		sqlStr = " update [db_user].[dbo].tbl_user_n " & VbCrlf
		sqlStr = sqlStr & " set usermail = '' " & VbCrlf
		sqlStr = sqlStr & " where userid = '" & userid & "' "
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	end if

	if (haveofflineaccount = "Y") and (issameusermail = "Y") then
		sqlStr = " update db_shop.dbo.tbl_total_shop_user " & VbCrlf
		sqlStr = sqlStr & " set Email = '' " & VbCrlf
		sqlStr = sqlStr & " where userseq = " & CStr(userseq) & " " & VbCrlf
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	end if

elseif (mode = "deloffusercell") then

	if (haveofflineaccount = "Y") then
		sqlStr = " update db_shop.dbo.tbl_total_shop_user " & VbCrlf
		sqlStr = sqlStr & " set hpno = '000-000-0000' " & VbCrlf
		sqlStr = sqlStr & " where userseq = " & CStr(userseq) & " " & VbCrlf
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	end if

	if (haveonlineaccount = "Y") and (issameusercell = "Y") then
		sqlStr = " update [db_user].[dbo].tbl_user_n " & VbCrlf
		sqlStr = sqlStr & " set usercell = '000-000-0000' " & VbCrlf
		sqlStr = sqlStr & " where userid = '" & userid & "' "
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	end if

elseif (mode = "setuserdivto01") then

	sqlStr = " update " & VbCrlf
	sqlStr = sqlStr & " [db_user].[dbo].[tbl_logindata] " & VbCrlf
 	sqlStr = sqlStr & " set userdiv = '01' " & VbCrlf
 	sqlStr = sqlStr & " where userid = '" & userid & "' and userdiv = '05' "
	rsget.Open sqlStr,dbget,1
	'response.write sqlStr

' 온라인고객탈퇴처리	' 2021.11.23 한용민
elseif (mode = "delonuser") then
	if not(C_CSPowerUser or C_ADMIN_AUTH) then
		response.write "탈퇴 처리는 cs팀 파트장 이상 권한이 필요 합니다."
		response.write "<script type='text/javascript'>"
		response.write "	alert('탈퇴 처리는 cs팀 파트장 이상 권한이 필요 합니다..');"
		response.write "</script>"
		dbget.close() : response.end
	end if

	' 회원탈퇴처리
	sqlStr = "exec [db_user_Hold].[dbo].[usp_WEBADMIN_user_del] '" & userid & "',N'"& complaintext &"'" & VbCrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	response.write "<script type='text/javascript'>"
	response.write "	alert('탈퇴 처리 되었습니다.');"
	response.write "	opener.location.reload();"
	response.write "	self.close();"
	response.write "</script>"
	response.end

' 온라인 휴면고객을 일반회원으로 전환	' 2023.10.30 한용민
elseif (mode = "ChangeOnHoldUser") then

	' 휴면고객을 일반회원으로 전환
	sqlStr = "exec db_user_Hold.dbo.sp_Ten_HoldUserRevive '" & userid & "', NULL" & VbCrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	response.write "<script type='text/javascript'>"
	response.write "	alert('휴면고객->일반회원으로 전환 처리 되었습니다.');"
	response.write "	parent.location.reload();"
	response.write "</script>"
	response.end
end if

'//임시번호 생성
function RandomStr()
    dim str, strlen
    dim rannum, ix

	'// o 0 l 1 제외, 혼동의 소지가 있어서 제외.
    str = "abcdefghijkmnpqrstuvwxyz23456789"
    strlen = 6

    Randomize

    For ix = 1 to strlen
    	 rannum = Int((32 - 1 + 1) * Rnd + 1)
    	 RandomStr = RandomStr + Mid(str,rannum,1)
    Next
end Function

'//회원비번 수정
sub setTempPassword(userid,strRdm)
    dim sqlStr
    dim Enc_userpass, Enc_userpass64

    Enc_userpass = MD5(CStr(strRdm))
    Enc_userpass64 = SHA256(MD5(CStr(strRdm)))


    '##########################################################
    '임시비밀번호로 변경
    sqlStr = " update [db_user].[dbo].[tbl_logindata]" + vbCrlf
    sqlStr = sqlStr + " set userpass=''" + vbCrlf
    sqlStr = sqlStr + " ,Enc_userpass=''" + vbCrlf   ''2018/07/13
    sqlStr = sqlStr + " ,Enc_userpass64='" + Enc_userpass64 + "'" + vbCrlf
    sqlStr = sqlStr + " where userid='" + userid + "'"
    dbget.Execute(sqlStr)

    '##########################################################
end sub
%>
<script type="text/javascript">
	alert('저장 되었습니다.');
	location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
