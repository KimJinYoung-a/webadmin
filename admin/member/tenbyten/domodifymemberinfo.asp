<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%

dim mode
dim userid, part_sn, posit_sn, level_sn, job_sn, isusing, userdiv, userimage
dim frontid, empno, username, juminno, birthday, issolar, sexflag, zipcode, zipaddr, useraddr, userphone, usercell, usermail, msnmail, interphoneno, extension, direct070, jobdetail, statediv, joinday, retireday, regdate
dim olduserpass, newuserpass
dim bigo		'관련샵아이디

dim i, sql

dim isadmin
Dim mywork
dim bizsection_Cd

mode = requestCheckvar(request("mode"),32)
isadmin = requestCheckvar(request("isadmin"),1)

userid = requestCheckvar(request("userid"),32)
part_sn = requestCheckvar(request("part_sn"),4)
posit_sn = requestCheckvar(request("posit_sn"),4)
level_sn = requestCheckvar(request("level_sn"),4)
job_sn = requestCheckvar(request("job_sn"),4)
isusing = requestCheckvar(request("isusing"),1)
userdiv = requestCheckvar(request("userdiv"),1)
bigo = requestCheckvar(request("bigo"),32)

frontid = requestCheckvar(request("frontid"),32)
empno = requestCheckvar(request("empno"),32)
username = requestCheckvar(request("username"),32)
juminno = requestCheckvar(request("juminno"),14)
birthday = requestCheckvar(request("birthday"),10)
issolar = requestCheckvar(request("issolar"),1)
sexflag = requestCheckvar(request("sexflag"),1)
zipcode = requestCheckvar(request("zipcode"),32)
zipaddr= requestCheckvar(request("zipaddr"),128)
useraddr = requestCheckvar(request("useraddr"),128)
userphone = requestCheckvar(request("userphone"),16)
usercell = requestCheckvar(request("usercell"),16)
usermail = requestCheckvar(request("usermail"),128)
msnmail = requestCheckvar(request("msnmail"),128)
interphoneno = requestCheckvar(request("interphoneno"),16)
extension = requestCheckvar(request("extension"),4)
direct070 = requestCheckvar(request("direct070"),16)
jobdetail = requestCheckvar(request("jobdetail"),128)
statediv = requestCheckvar(request("statediv"),1)
joinday = requestCheckvar(request("joinday"),10)
retireday = requestCheckvar(request("retireday"),10)
regdate = requestCheckvar(request("regdate"),10)
mywork = requestCheckvar(request("mywork"),150)

olduserpass = requestCheckvar(request("olduserpass"),16)
newuserpass = requestCheckvar(request("newuserpass"),16)

userimage = requestCheckvar(request("userimage"),100)
bizsection_Cd= requestCheckvar(request("selBiz"),10)

userid = Replace(userid, " ", "")
frontid = Replace(frontid, " ", "")
username = Replace(username, " ", "")
empno = Replace(empno, " ", "")



'==============================================================================
dim oMember
Set oMember = new CTenByTenMember

if (isadmin = "N") then
	userid = session("ssBctId")
end if

oMember.Fuserid = userid

oMember.fnGetScmMyInfo



if (isadmin = "N") then
	username = oMember.Fusername
	frontid = oMember.Ffrontid
	usermail = oMember.Fusermail
	interphoneno = oMember.Finterphoneno
	'extension = oMember.Fextension
	'direct070 = oMember.Fdirect070
else
	if (IsNull(oMember.Fbirthday)) then
		birthday = "1950-01-01"
	else
		birthday = Left(oMember.Fbirthday, 10)
	end if
	issolar = oMember.Fissolar
	sexflag = oMember.Fsexflag
	msnmail = oMember.Fmsnmail
end if

if (mode = "base") then
	'기본정보 수정
	sql = "      update [db_partner].[dbo].tbl_user_tenbyten " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	username = '" & username & "', " & vbCrlf
	sql = sql & "	frontid = '" & frontid & "', " & vbCrlf
	sql = sql & "	birthday = '" & birthday & "', " & vbCrlf
	sql = sql & "	issolar = '" & issolar & "', " & vbCrlf
	sql = sql & "	sexflag = '" & sexflag & "', " & vbCrlf
	sql = sql & "	usermail = '" & usermail & "', " & vbCrlf
	sql = sql & "	msnmail = '" & msnmail & "', " & vbCrlf
	sql = sql & "	interphoneno = '" & interphoneno & "', " & vbCrlf
	sql = sql & "	extension = '" & extension & "', " & vbCrlf
	sql = sql & "	direct070 = '" & direct070 & "', " & vbCrlf
	sql = sql & "	userimage = '" & userimage & "', " & vbCrlf
	sql = sql & "	mywork = '" & mywork & "' " & vbCrlf
	sql = sql & "where userid = '" & userid & "' " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('수정되었습니다.');</script>"
	response.write "<script>document.location.href = '" & request.servervariables("http_referer") & "';</script>"

elseif (mode = "addr") then
	'비상연락망 수정
	sql = "      update [db_partner].[dbo].tbl_user_tenbyten " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	usercell = '" & usercell & "', " & vbCrlf
	sql = sql & "	userphone = '" & userphone & "', " & vbCrlf
	sql = sql & "	zipcode = '" & zipcode & "', " & vbCrlf
	sql = sql & "	zipaddr = '" & zipaddr & "', " & vbCrlf
	sql = sql & "	useraddr = '" & useraddr & "' " & vbCrlf
	sql = sql & "where userid = '" & userid & "' " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('수정되었습니다.');</script>"
	response.write "<script>document.location.href = '" & request.servervariables("http_referer") & "';</script>"

elseif (mode = "auth") then
	'권한 수정
	sql = "      update [db_partner].[dbo].tbl_partner " & vbCrlf
	sql = sql & "set lastInfoChgDT=getdate(), " & vbCrlf	
	sql = sql & "	level_sn = '" & level_sn & "', " & vbCrlf	
	sql = sql & "where id = '" & userid & "' " & vbCrlf
	dbget.Execute(sql)

	sql = "      update [db_partner].[dbo].tbl_user_tenbyten " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	part_sn = '" & part_sn & "', " & vbCrlf	
	sql = sql & "	posit_sn = '" & posit_sn & "', " & vbCrlf
	sql = sql & "	job_sn = '" & job_sn & "' " & vbCrlf
	sql = sql & "	jobdetail = '" & jobdetail & "' " & vbCrlf
	sql = sql & "where userid = '" & userid & "' " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('수정되었습니다.');</script>"
	response.write "<script>document.location.href = '" & request.servervariables("http_referer") & "';</script>"

elseif (mode = "mypass") then
	'비밀번호 수정(직원)

	sql = "select top 1 id " & vbCrlf
	sql = sql & "from [db_partner].[dbo].tbl_partner " & vbCrlf
	sql = sql & "where 1 = 1 " & vbCrlf
	sql = sql & "and id = '" & userid & "' " & vbCrlf
	sql = sql & "and password = '" & olduserpass & "' " & vbCrlf
	rsget.Open sql,dbget,1
	if (rsget.EOF or rsget.BOF) then
		newuserpass = ""
	end if
	rsget.Close

	if (newuserpass = "") then
		response.write "<script>alert('기존 비밀번호가 잘못 입력되었습니다.');</script>"
		response.write "<script>history.back();</script>"
	else
		'//패스워드 정책 검사
		if chkPasswordComplex(userid,newuserpass)<>"" then
			response.write "<script language='javascript'>alert('" & chkPasswordComplex(userid,newuserpass) & "\n비밀번호를 확인후 다시 시도해주세요.');</script>"
			response.write "<script>document.location.href = '" & request.servervariables("http_referer") & "';</script>"
		else
			sql = "      update [db_partner].[dbo].tbl_partner " & vbCrlf
			sql = sql & "set lastInfoChgDT=getdate(), " & vbCrlf
			sql = sql & "	password = '" & newuserpass & "' " & vbCrlf
			sql = sql & "where id = '" & userid & "' " & vbCrlf
			dbget.Execute(sql)

			response.write "<script>alert('수정되었습니다.');</script>"
			response.write "<script>document.location.href = '" & request.servervariables("http_referer") & "';</script>"
		end if
	end if
elseif (mode = "moreinfo") then
	'추가정보 수정
	sql = "      update [db_partner].[dbo].tbl_user_tenbyten " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	joinday = '" & joinday & "' " & vbCrlf
	sql = sql & "where userid = '" & userid & "' " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('수정되었습니다.');</script>"
	response.write "<script>document.location.href = '" & request.servervariables("http_referer") & "';</script>"
elseif (mode = "mywork") then
	sql = "      update [db_partner].[dbo].tbl_user_tenbyten " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	mywork = '" & mywork & "' " & vbCrlf
	sql = sql & "where empno = '" & empno & "' " & vbCrlf
	dbget.Execute(sql)

	response.write "<script>alert('수정되었습니다.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>window.close();</script>"
elseif (mode="biz") then  
	Dim objCmd, returnValue
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_partner].[dbo].sp_Ten_user_tenbyten_SetBizsection('"&empno&"','"&bizsection_Cd&"')}"			
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
		End With	
		 returnValue = objCmd(0).Value	
	IF 	returnValue <> 1 THEN
			Call Alert_return ("데이터 처리에 문제가 발생하였습니다.")
		response.end
	END IF	 
		Call Alert_move ("ERP부서가 등록되었습니다.", "member_modify.asp")
else
	'
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->