<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��Ʈ��� �������� ����
' History : 2007.07.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<%

dim mode, usermail, msnmail, interphoneno, extension, direct070, jobdetail, statediv, joinday, retireday, regdate
dim userid, empno, olduserpass, newuserpass, i, sql, bizsection_Cd, mywork, bigo
dim frontid, username, juminno, birthday, issolar, sexflag, zipcode, zipaddr, useraddr, userphone, usercell
dim part_sn, posit_sn, level_sn, job_sn, isusing, userdiv, userimage, messenger, userNameEN

empno = session("ssBctSn")

mode = requestCheckvar(request("mode"),32)
userNameEN = requestCheckvar(trim(request("userNameEN")),32)
userid = requestCheckvar(request("userid"),32)
part_sn = requestCheckvar(request("part_sn"),4)
posit_sn = requestCheckvar(request("posit_sn"),4)
level_sn = requestCheckvar(request("level_sn"),4)
job_sn = requestCheckvar(request("job_sn"),4)
isusing = requestCheckvar(request("isusing"),1)
userdiv = requestCheckvar(request("userdiv"),1)
bigo = requestCheckvar(request("bigo"),32)
frontid = requestCheckvar(request("frontid"),32)
username = requestCheckvar(request("username"),32)
juminno = requestCheckvar(request("juminno"),14)
birthday = requestCheckvar(request("birthday"),10)
issolar = requestCheckvar(request("issolar"),1)
sexflag = requestCheckvar(request("sexflag"),1)
zipcode = requestCheckvar(request("zipcode"),32)
zipaddr= stripHTML(requestCheckvar(request("zipaddr"),128))
useraddr = stripHTML(requestCheckvar(request("useraddr"),128))
userphone = requestCheckvar(request("userphone"),16)
usercell = requestCheckvar(request("usercell"),16)
usermail = stripHTML(requestCheckvar(request("usermail"),128))
msnmail = stripHTML(requestCheckvar(request("msnmail"),128))
messenger = stripHTML(requestCheckvar(request("messenger"),128))
interphoneno = requestCheckvar(request("interphoneno"),16)
extension = requestCheckvar(request("extension"),4)
direct070 = requestCheckvar(request("direct070"),16)
jobdetail = stripHTML(requestCheckvar(request("jobdetail"),128))
statediv = requestCheckvar(request("statediv"),1)
joinday = requestCheckvar(request("joinday"),10)
retireday = requestCheckvar(request("retireday"),10)
regdate = requestCheckvar(request("regdate"),10)
mywork = stripHTML(requestCheckvar(request("mywork"),150))

olduserpass = requestCheckvar(request("olduserpass"),16)
newuserpass = requestCheckvar(request("newuserpass"),16)

userimage = requestCheckvar(request("userimage"),100)
bizsection_Cd= requestCheckvar(request("selBiz"),10)

userid = Replace(userid, " ", "")
frontid = Replace(frontid, " ", "")
username = Replace(username, " ", "")
empno = Replace(empno, " ", "")


'==============================================================================
'' dim oMember
'' Set oMember = new CTenByTenMember

'' oMember.Fempno = empno

'' oMember.fnGetMemberData


if (mode = "base") then
    if checkNotValidHTML(msnmail) or checkNotValidHTML(messenger) or checkNotValidHTML(mywork) then
    	response.write "<script type='text/javascript'>alert('���뿡 ����ϽǼ� ���� �±װ� �ֽ��ϴ�.');</script>"
    	response.write "<script type='text/javascript'>document.location.href = '" & request.servervariables("http_referer") & "';</script>"
        dbget.close() : response.end
    end if

	'�⺻���� ����
	sql = "      update [db_partner].[dbo].tbl_user_tenbyten " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	birthday = '" & birthday & "', " & vbCrlf
	sql = sql & "	issolar = '" & issolar & "', " & vbCrlf
	sql = sql & "	sexflag = '" & sexflag & "', " & vbCrlf
	sql = sql & "	msnmail = '" & msnmail & "', " & vbCrlf
	sql = sql & "	messenger = '" & messenger & "', " & vbCrlf
	sql = sql & "	interphoneno = '" & interphoneno & "', " & vbCrlf
	sql = sql & "	extension = '" & extension & "', " & vbCrlf
	sql = sql & "	direct070 = '" & direct070 & "', " & vbCrlf
	sql = sql & "	userimage = '" & userimage & "', " & vbCrlf
	sql = sql & "	userNameEN = N'" & userNameEN & "', " & vbCrlf
	sql = sql & "	mywork = '" & mywork & "' " & vbCrlf
	sql = sql & "where empno = '" & empno & "' " & vbCrlf
	'response.write sql
	dbget.Execute(sql)

	response.write "<script type='text/javascript'>alert('�����Ǿ����ϴ�.');</script>"
	response.write "<script type='text/javascript'>document.location.href = '" & request.servervariables("http_referer") & "';</script>"

elseif (mode = "addr") then
    if checkNotValidHTML(userphone) or checkNotValidHTML(useraddr) then
    	response.write "<script type='text/javascript'>alert('���뿡 ����ϽǼ� ���� �±װ� �ֽ��ϴ�.');</script>"
    	response.write "<script type='text/javascript'>document.location.href = '" & request.servervariables("http_referer") & "';</script>"
        dbget.close() : response.end
    end if

	'��󿬶��� ����
	sql = "      update [db_partner].[dbo].tbl_user_tenbyten " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	usercell = '" & usercell & "', " & vbCrlf
	sql = sql & "	userphone = '" & userphone & "', " & vbCrlf
	sql = sql & "	zipcode = '" & zipcode & "', " & vbCrlf
	sql = sql & "	zipaddr = '" & zipaddr & "', " & vbCrlf
	sql = sql & "	useraddr = '" & useraddr & "' " & vbCrlf
	sql = sql & "where empno = '" & empno & "' " & vbCrlf
	dbget.Execute(sql)

	response.write "<script type='text/javascript'>alert('�����Ǿ����ϴ�.');</script>"
	response.write "<script type='text/javascript'>document.location.href = '" & request.servervariables("http_referer") & "';</script>"

elseif (mode = "auth") then
	'���� ����
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

	response.write "<script type='text/javascript'>alert('�����Ǿ����ϴ�.');</script>"
	response.write "<script type='text/javascript'>document.location.href = '" & request.servervariables("http_referer") & "';</script>"

elseif (mode = "mypass") then
	'��й�ȣ ����(����)

	sql = "select top 1 id " & vbCrlf
	sql = sql & " from [db_partner].[dbo].tbl_partner p " & vbCrlf
	sql = sql & " join [db_partner].[dbo].tbl_user_tenbyten t " & vbCrlf
	sql = sql & " on p.id = t.userid " & vbCrlf
	sql = sql & " where 1 = 1 " & vbCrlf
	sql = sql & " and t.empno = '" & empno & "' " & vbCrlf
	sql = sql & " and p.Enc_password64 = '" & SHA256(md5(olduserpass)) & "' " & vbCrlf  
	
	rsget.Open sql,dbget,1
	if (rsget.EOF or rsget.BOF) then
		newuserpass = ""
	else
	    userid=rsget("id")  ''2014/07/14 �߰�
	end if
	rsget.Close

	if (newuserpass = "") then
		response.write "<script type='text/javascript'>alert('���� ��й�ȣ�� �߸� �ԷµǾ����ϴ�.');</script>"
		response.write "<script type='text/javascript'>history.back();</script>"
		dbget.close() : response.end
	else
		'//�н����� ��å �˻�
		if chkPasswordComplex(empno,newuserpass)<>"" then
			response.write "<script type='text/javascript'>alert('" & chkPasswordComplex(empno,newuserpass) & "\n��й�ȣ�� Ȯ���� �ٽ� �õ����ּ���.');</script>"
			response.write "<script type='text/javascript'>document.location.href = '" & request.servervariables("http_referer") & "';</script>"
			dbget.close() : response.end

		else
			sql = "       update p " & vbCrlf
			sql = sql & " set p.Enc_password64 = '" & SHA256(md5(newuserpass)) & "' " & vbCrlf
			sql = sql & " ,p.Enc_password = '' " & vbCrlf
			sql = sql & " from [db_partner].[dbo].tbl_partner p " & vbCrlf
			sql = sql & " join [db_partner].[dbo].tbl_user_tenbyten t " & vbCrlf
			sql = sql & " on p.id = t.userid " & vbCrlf
			sql = sql & " where 1 = 1 " & vbCrlf
			sql = sql & " and t.empno = '" & empno & "' " & vbCrlf

			'response.write sql
			dbget.Execute(sql)
            
            ''���� �α��� ���� ���� //2014/07/14 '' tbl_user_tenbyten ����α��� ����
            sql = "exec db_partner.dbo.sp_Ten_Add_PartnerLoginLog '"&userid&"','"&Left(request.ServerVariables("REMOTE_ADDR"),16)&"','R','',0"

			'response.write sql
            dbget.Execute sql

			response.write "<script type='text/javascript'>alert('�����Ǿ����ϴ�.');</script>"
			response.write "<script type='text/javascript'>document.location.href = '" & request.servervariables("http_referer") & "';</script>"
			dbget.close() : response.end
		end if
	end if

elseif (mode = "myemppass") then
	'��й�ȣ ����(���)

	sql = "select top 1 empno " & vbCrlf
	sql = sql & "from [db_partner].[dbo].tbl_user_tenbyten " & vbCrlf
	sql = sql & "where 1 = 1 " & vbCrlf
	sql = sql & "and empno = '" & empno & "' " & vbCrlf
	sql = sql & "and Enc_emppass64 = '" & SHA256(md5(olduserpass)) & "' " & vbCrlf
	rsget.Open sql,dbget,1
	if (rsget.EOF or rsget.BOF) then
		newuserpass = ""
	end if
	rsget.Close

	if (newuserpass = "") then
		response.write "<script type='text/javascript'>alert('���� ��й�ȣ�� �߸� �ԷµǾ����ϴ�.');</script>"
		response.write "<script type='text/javascript'>history.back();</script>"
	    dbget.close() : response.end
	else
		'//�н����� ��å �˻�
		if chkPasswordComplex(empno,newuserpass)<>"" then
			response.write "<script language='javascript'>alert('" & chkPasswordComplex(empno,newuserpass) & "\n��й�ȣ�� Ȯ���� �ٽ� �õ����ּ���.');</script>"
			response.write "<script type='text/javascript'>document.location.href = '" & request.servervariables("http_referer") & "';</script>"
			dbget.close() : response.end
		else
			sql = "update [db_partner].[dbo].tbl_user_tenbyten " & vbCrlf
			sql = sql & " set lastEmpnoPwChgDT=getdate()" & vbCrlf
			sql = sql & " , Enc_emppass64 = '" & SHA256(md5(newuserpass)) & "' where" & vbCrlf
			sql = sql & " empno = '" & empno & "' " & vbCrlf

			'response.write sql & "<br>"
			dbget.Execute sql

			response.write "<script type='text/javascript'>alert('�����Ǿ����ϴ�.');</script>"
			response.write "<script type='text/javascript'>document.location.href = '" & request.servervariables("http_referer") & "';</script>"
			dbget.close() : response.end
		end if
	end if

elseif (mode = "moreinfo") then
	'�߰����� ����
	sql = "      update [db_partner].[dbo].tbl_user_tenbyten " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	joinday = '" & joinday & "' " & vbCrlf
	sql = sql & "where userid = '" & userid & "' " & vbCrlf
	dbget.Execute(sql)

	response.write "<script type='text/javascript'>alert('�����Ǿ����ϴ�.');</script>"
	response.write "<script type='text/javascript'>document.location.href = '" & request.servervariables("http_referer") & "';</script>"
elseif (mode = "mywork") then
	sql = "      update [db_partner].[dbo].tbl_user_tenbyten " & vbCrlf
	sql = sql & "set " & vbCrlf
	sql = sql & "	mywork = '" & mywork & "' " & vbCrlf
	sql = sql & "where empno = '" & empno & "' " & vbCrlf
	dbget.Execute(sql)

	response.write "<script type='text/javascript'>alert('�����Ǿ����ϴ�.');</script>"
	response.write "<script type='text/javascript'>opener.location.reload();</script>"
	response.write "<script type='text/javascript'>window.close();</script>"
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
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.")
		response.end
	END IF
		Call Alert_move ("ERP�μ��� ��ϵǾ����ϴ�.", "member_modify.asp")
else
	'
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->