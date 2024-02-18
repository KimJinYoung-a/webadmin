<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'           2023.10.30 �ѿ�� ����(�޸��������ǥ��. �޸����->�Ϲݰ��� ��ȯ ���� ����)
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

	'// CS �޸� ����
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

	contents = "�ٸ� ���� ��û���� ���� �ڵ��� ��ȣ(" + usercell + ")�� 000-000-0000 �� ����"

    Call AddCsMemo("", "1", userid, session("ssBctId"), contents)
    response.write "<script>alert('CS MEMO�� ����Ǿ����ϴ�.')</script>"

' �Ϲ���ȭ��ȣ ����(�¶��ΰ� ���������� ���� ��ȣ �ϰ��) 	' 2019.04.10 �ѿ�� ����
elseif (mode = "delonuserphone") or ((mode = "delonuserphone") and (issameuserphone = "Y")) then
	'// CS �޸� ����
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

	contents = "�ٸ� ���� ��û���� ���� ��ȭ��ȣ(" + userphone + ")�� 000-000-0000 �� ����"

    Call AddCsMemo("", "1", userid, session("ssBctId"), contents)
    response.write "<script>alert('CS MEMO�� ���泻���� ����Ǿ����ϴ�.')</script>"

elseif (mode = "delonusermail") then

	'// CS �޸� ����
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

	contents = "�ٸ� ���� ��û���� ���� �̸����ּ�(" + usermail + ")�� ����"

    Call AddCsMemo("", "1", userid, session("ssBctId"), contents)
    response.write "<script>alert('CS MEMO�� ����Ǿ����ϴ�.')</script>"

elseif (mode = "resetUserPass") then

	'' �ӽú�й�ȣ ���� ���μ��� ���� (���� 0000 -> �ű� 6�ڸ�����), 2015-07-15, skyer9
	'' ���� ���ȳ� �ʿ�

	dim strRdm
	strRdm = RandomStr()
	call setTempPassword(userid,strRdm)

	contents = "���� ��û���� ��й�ȣ�� " & strRdm & " ���� �ӽú�й�ȣ ����"

    Call AddCsMemo("", "2", userid, session("ssBctId"), contents)
    response.write "<script>alert('CS MEMO�� ����Ǿ����ϴ�.')</script>"

' �¶��ΰ�Ż��ó��	' 2021.11.23 �ѿ��
elseif (mode = "delonuser") then
	if not(C_CSPowerUser or C_ADMIN_AUTH) then
		response.write "Ż�� ó���� cs�� ��Ʈ�� �̻� ������ �ʿ� �մϴ�."
		response.write "<script type='text/javascript'>"
		response.write "	alert('Ż�� ó���� cs�� ��Ʈ�� �̻� ������ �ʿ� �մϴ�..');"
		response.write "</script>"
		dbget.close() : response.end
	end if

	complaintext=trim(complaintext)
	if complaintext <> "" then
		if checkNotValidHTML(complaintext) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
			response.write "</script>"
			dbget.close() : response.End
		end if
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('Ż�� ������ �Է����ּ���');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close() : response.End
	end If

    Call AddCsMemo("", "1", userid, session("ssBctId"), complaintext)
    response.write "<script type='text/javascript'>alert('CS MEMO�� ����Ǿ����ϴ�.')</script>"
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

' �Ϲ���ȭ��ȣ ���� 	' 2019.04.10 �ѿ�� ����
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

' �¶��ΰ�Ż��ó��	' 2021.11.23 �ѿ��
elseif (mode = "delonuser") then
	if not(C_CSPowerUser or C_ADMIN_AUTH) then
		response.write "Ż�� ó���� cs�� ��Ʈ�� �̻� ������ �ʿ� �մϴ�."
		response.write "<script type='text/javascript'>"
		response.write "	alert('Ż�� ó���� cs�� ��Ʈ�� �̻� ������ �ʿ� �մϴ�..');"
		response.write "</script>"
		dbget.close() : response.end
	end if

	' ȸ��Ż��ó��
	sqlStr = "exec [db_user_Hold].[dbo].[usp_WEBADMIN_user_del] '" & userid & "',N'"& complaintext &"'" & VbCrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	response.write "<script type='text/javascript'>"
	response.write "	alert('Ż�� ó�� �Ǿ����ϴ�.');"
	response.write "	opener.location.reload();"
	response.write "	self.close();"
	response.write "</script>"
	response.end

' �¶��� �޸���� �Ϲ�ȸ������ ��ȯ	' 2023.10.30 �ѿ��
elseif (mode = "ChangeOnHoldUser") then

	' �޸���� �Ϲ�ȸ������ ��ȯ
	sqlStr = "exec db_user_Hold.dbo.sp_Ten_HoldUserRevive '" & userid & "', NULL" & VbCrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	response.write "<script type='text/javascript'>"
	response.write "	alert('�޸��->�Ϲ�ȸ������ ��ȯ ó�� �Ǿ����ϴ�.');"
	response.write "	parent.location.reload();"
	response.write "</script>"
	response.end
end if

'//�ӽù�ȣ ����
function RandomStr()
    dim str, strlen
    dim rannum, ix

	'// o 0 l 1 ����, ȥ���� ������ �־ ����.
    str = "abcdefghijkmnpqrstuvwxyz23456789"
    strlen = 6

    Randomize

    For ix = 1 to strlen
    	 rannum = Int((32 - 1 + 1) * Rnd + 1)
    	 RandomStr = RandomStr + Mid(str,rannum,1)
    Next
end Function

'//ȸ����� ����
sub setTempPassword(userid,strRdm)
    dim sqlStr
    dim Enc_userpass, Enc_userpass64

    Enc_userpass = MD5(CStr(strRdm))
    Enc_userpass64 = SHA256(MD5(CStr(strRdm)))


    '##########################################################
    '�ӽú�й�ȣ�� ����
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
	alert('���� �Ǿ����ϴ�.');
	location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
