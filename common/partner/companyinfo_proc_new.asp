<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  공지사항 뷰
' History : 이상구 생성
'           2018.07.12 한용민 수정(ISMS대응 권한체크)
'####################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/ecContractApi_function.asp"-->
<% 
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim uid,company_name,email,manager_name,address
dim manager_address, tel, fax, userdiv
dim groupid,defaultsongjangdiv, c_userdiv, p_userdiv, pcuserdiv, mduserid
dim company_no_img, jungsan_acctno_img

groupid	= requestCheckvar(request("groupid"),6)
uid	= requestCheckvar(request("uid"),32)
defaultsongjangdiv = requestCheckvar(request("defaultsongjangdiv"),16)

company_name = requestCheckVar(html2db(request("company_name")),100)
email = requestCheckVar(html2db(request("email")),200)
manager_name = requestCheckVar(html2db(request("manager_name")),50)
address = requestCheckVar(html2db(request("address")),150)
manager_address = requestCheckVar(html2db(request("manager_address")),150)
tel	= requestCheckVar(html2db(request("tel")),50)
fax	= requestCheckVar(html2db(request("fax")),50)
pcuserdiv = requestCheckVar(request("pcuserdiv"),10)
mduserid = requestCheckVar(request("mduserid"),16)
if (pcuserdiv<>"") then
    p_userdiv = Trim(splitvalue(pcuserdiv,"_",0))
    c_userdiv = Trim(splitvalue(pcuserdiv,"_",1))
end if
dim applyallbrand
applyallbrand = requestCheckVar(request("applyallbrand"),10)

company_no_img = requestCheckVar(html2db(request("company_no_img")),150)
jungsan_acctno_img = requestCheckVar(html2db(request("jungsan_acctno_img")),150)

dim ceoname, company_no, zipcode, jungsan_gubun
dim jungsan_date,jungsan_bank,jungsan_acctno
dim jungsan_acctname, manager_phone, manager_hp
dim deliver_name, deliver_phone, deliver_email
dim deliver_hp, jungsan_name, jungsan_phone, jungsan_email
dim jungsan_hp, prtidx, jungsan_date_off
dim p_return_zipcode, p_return_address, p_return_address2
dim ogroupinfo, chkmwdiv, addmwdiv, addsellplace, addON_dlvtype, addON_dlvlimit, addON_dlvpay
ceoname			= requestCheckVar(html2db(request("ceoname")),50)
company_no  	= requestCheckVar(request("company_no"),20)
zipcode			= requestCheckVar(request("zipcode"),10)
jungsan_gubun 	= requestCheckVar(request("jungsan_gubun"),50)
jungsan_date 	= requestCheckVar(request("jungsan_date"),50)
jungsan_date_off 	= requestCheckVar(request("jungsan_date_off"),50)
jungsan_bank 	= requestCheckVar(html2db(request("jungsan_bank")),50)
jungsan_acctno 	= requestCheckVar(request("jungsan_acctno"),100)
jungsan_acctname = requestCheckVar(html2db(request("jungsan_acctname")),50)
manager_phone 	= requestCheckVar(request("manager_phone"),50)
manager_hp 		= requestCheckVar(request("manager_hp"),50)
deliver_name 	= requestCheckVar(html2db(request("deliver_name")),50)
deliver_phone 	= requestCheckVar(request("deliver_phone"),50)
deliver_email 	= requestCheckVar(request("deliver_email"),150)
deliver_hp 		= requestCheckVar(request("deliver_hp"),50)
jungsan_name 	= requestCheckVar(html2db(request("jungsan_name")),50)
jungsan_phone 	= requestCheckVar(request("jungsan_phone"),50)
jungsan_email 	= requestCheckVar(request("jungsan_email"),150)
jungsan_hp 		= requestCheckVar(request("jungsan_hp"),50)
prtidx 			= requestCheckVar(request("prtidx"),10)
jungsan_acctno = replace(jungsan_acctno,"-","")
jungsan_acctname = replace(Trim(jungsan_acctname)," ","")

dim company_zipcode, company_address, company_address2
dim company_tel, company_fax, return_zipcode, return_address, return_address2
dim manager_email
dim cs_name, cs_phone, cs_hp, cs_email

company_zipcode = requestCheckVar(request("company_zipcode"),10)
company_address = requestCheckVar(request("company_address"),150)
company_address2 = requestCheckVar(request("company_address2"),150)
company_tel = requestCheckVar(request("company_tel"),50)
company_fax = requestCheckVar(request("company_fax"),50)
return_zipcode = requestCheckVar(request("return_zipcode"),10)
return_address = requestCheckVar(request("return_address"),150)
return_address2 = requestCheckVar(request("return_address2"),150)
manager_email = requestCheckVar(request("manager_email"),150)

p_return_zipcode = requestCheckVar(request("p_return_zipcode"),10)
p_return_address = requestCheckVar(request("p_return_address"),150)
p_return_address2 = requestCheckVar(request("p_return_address2"),150)

cs_name = requestCheckVar(html2db(request("cs_name")),50)
cs_phone = requestCheckVar(html2db(request("cs_phone")),50)
cs_hp = requestCheckVar(html2db(request("cs_hp")),50)
cs_email = requestCheckVar(html2db(request("cs_email")),150)

dim vPurchaseType, vOffCateCode, vOffMDUserID
vPurchaseType				= Request("purchasetype")
vOffCateCode				= Request("offcatecode")
vOffMDUserID				= Request("offmduserid")

if not IsNumeric(prtidx) then prtidx=9999

dim company_upjong,company_uptae
company_upjong  = requestCheckVar(html2db(request("company_upjong")),100)
company_uptae   = requestCheckVar(html2db(request("company_uptae")),100)

dim subid
subid   = requestCheckVar(request("subid"),50)

dim mode
mode = requestCheckVar(request("mode"),30)

dim commission, password,passwordS
dim Enc_userpass, Enc_userpass64,Enc_2userpass64

commission = request("commission")
password = requestCheckVar(request("password"),32)
passwordS = requestCheckVar(request("passwordS"),32)

Enc_userpass = MD5(password)
Enc_userpass64 = SHA256(MD5(password))
Enc_2userpass64= SHA256(MD5(passwordS))

'####### 직원 연락처 Get. 웹훅 발송용. #######
	public Function fnGetMemberEmail(id)
		Dim strSql
		strSql = "	SELECT isNull(usermail,'') AS email FROM [db_partner].[dbo].tbl_user_tenbyten WHERE userid = '" & id & "' and userid <> '' "
		rsget.Open strSql,dbget,1

		IF not rsget.EOF THEN
			If rsget("email") = "" Then
				fnGetMemberEmail = ""
			Else
				fnGetMemberEmail = rsget("email")
			End If
		Else
			fnGetMemberEmail = ""
		END IF
		rsget.close
	End Function

'//패스워드 정책 검사(2008.12.15;허진원)
if mode="edit" or mode="addnewupchebrand" then
    if chkPasswordComplex(uid,password)<>"" then
        response.write "<script language='javascript'>" &vbCrLf &_
                        "	alert('" & chkPasswordComplex(uid,password) & "\n1차 비밀번호를 확인후 다시 시도해주세요.');" &vbCrLf &_
                        " 	history.back();" &vbCrLf &_
                        "</script>"
        dbget.close()	:	response.End
    end if
    
    if chkPasswordComplex(uid,passwordS)<>"" then
        response.write "<script language='javascript'>" &vbCrLf &_
                        "	alert('" & chkPasswordComplex(uid,passwordS) & "\n2차 비밀번호를 확인후 다시 시도해주세요.');" &vbCrLf &_
                        " 	history.back();" &vbCrLf &_
                        "</script>"
        dbget.close()	:	response.End
    end if
end if

dim socname_kor, socname, isusing, isextusing, streetusing
dim extstreetusing, specialbrand, maeipdiv, defaultmargine

socname_kor  = requestCheckVar(html2db(request("socname_kor")),150)
socname		 = requestCheckVar(html2db(request("socname")),150)
isusing		 = requestCheckVar(request("isusing"),10)
isextusing	 = requestCheckVar(request("isextusing"),10)
streetusing	 = requestCheckVar(request("streetusing"),10)
extstreetusing	 = requestCheckVar(request("extstreetusing"),10)
specialbrand	 = requestCheckVar(request("specialbrand"),10)
maeipdiv		 = requestCheckVar(request("maeipdiv"),10)
defaultmargine	 = requestCheckVar(request("defaultmargine"),20)

dim sqlStr, idExists
dim makerid, signtype, addmargin

''rw mode
On Error Resume Next
dbget.beginTrans

if mode="addnewupchebrand" then
    if checkNotValidHTML(company_zipcode) or checkNotValidHTML(company_address) or checkNotValidHTML(company_address2) or checkNotValidHTML(company_uptae) or checkNotValidHTML(company_upjong) then
    	response.write "<script>alert('사업자등록정보에 사용하실수 없는 태그가 있습니다.');</script>"
    	response.write "<script>document.location.href = '" & refer & "';</script>"
        response.end
    end if
    if checkNotValidHTML(company_tel) or checkNotValidHTML(company_fax) or checkNotValidHTML(return_zipcode) or checkNotValidHTML(return_address) or checkNotValidHTML(return_address2) then
    	response.write "<script>alert('파트너 기본정보에 사용하실수 없는 태그가 있습니다.');</script>"
    	response.write "<script>document.location.href = '" & refer & "';</script>"
        response.end
    end if
    if checkNotValidHTML(manager_name) or checkNotValidHTML(manager_phone) or checkNotValidHTML(manager_email) or checkNotValidHTML(manager_hp) or checkNotValidHTML(jungsan_name) or checkNotValidHTML(jungsan_phone) or checkNotValidHTML(jungsan_email) or checkNotValidHTML(jungsan_hp) then
    	response.write "<script>alert('파트너 담당자정보에 사용하실수 없는 태그가 있습니다.');</script>"
    	response.write "<script>document.location.href = '" & refer & "';</script>"
        response.end
    end if

	'// 아이디 중복 확인
	sqlStr = "select count(*) from [db_user].[dbo].tbl_logindata where userid='" & uid & "'"
	rsget.Open sqlStr,dbget,1
	if rsget(0)>0 then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('이미 정보 입력을 완료하였습니다.');" &vbCrLf &_
						" 	history.back();" &vbCrLf &_
						"</script>"
		response.End
	end if
	rsget.Close

    if (company_no = "888-00-00000") then
		'// 해외는 앞부분 888 이 고정이고 뒷부분 숫자는 자동증가

		idx = 0
		sqlStr = " select top 1 cast(right(replace(company_no , '-', ''), 7) as int) as idx "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_partner].[dbo].tbl_partner_group "
		sqlStr = sqlStr + " where Left(company_no, 3) = '888' and len(replace(company_no , '-', '')) = 10 "
		sqlStr = sqlStr + " order by cast(right(replace(company_no , '-', ''), 7) as int) desc "

		rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				idx = rsget("idx")
			end if
		rsget.Close

		idx = idx + 1
		idx = Format00(7, idx)
		company_no = "888-" + Left(idx, 2) + "-" + Right(idx, 5)

	end if

	'파트너 테이블 정산일 정보 가져오기
	sqlStr = "select top 1 isnull(p.jungsan_date,'') as jungsan_date, isnull(p.jungsan_date_off,'') as jungsan_date_off"
	sqlStr = sqlStr + ", isnull(p.signtype,0) as signtype , c.maeipdiv, c.defaultmargine"
	sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
	sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on p.id=c.userid"
	sqlStr = sqlStr + " where p.id='" & Cstr(uid) & "'"
	rsget.Open sqlStr,dbget,1
		if not rsget.Eof then
			jungsan_date = rsget("jungsan_date")
			jungsan_date_off = rsget("jungsan_date_off")
			signtype = rsget("signtype")
			maeipdiv = rsget("maeipdiv")
			addmargin = rsget("defaultmargine")
		end if
	rsget.Close

	''insert tbl_logindata
	sqlStr = "insert into [db_user].[dbo].tbl_logindata"
	sqlStr = sqlStr + "(userid,userpass,userdiv,lastlogin,Enc_userpass,Enc_userpass64,counter)" + vbCrlf
	sqlStr = sqlStr + " Values("
	sqlStr = sqlStr + " '" + (uid) + "'" + vbCrlf
	sqlStr = sqlStr + " ,'' " + vbCrlf
	sqlStr = sqlStr + ",'" + (c_userdiv) + "'" + vbCrlf
	sqlStr = sqlStr + ",getdate()" + vbCrlf
	sqlStr = sqlStr + ",''" + vbCrlf
	sqlStr = sqlStr + ",'" + (Enc_userpass64) + "'" + vbCrlf
	sqlStr = sqlStr + ",0" & ")"
	rsget.Open sqlStr,dbget,1

	''insert tbl_partner_group
	''Get Last Group ID
	if (groupid<>"") then
		sqlStr = "update [db_partner].[dbo].tbl_partner_group" + VbCrlf
		sqlStr = sqlStr + " set company_name='" + company_name + "'" + VbCrlf
		'' sqlStr = sqlStr + " ,company_no='" + company_no + "'" + VbCrlf       ''주석처리 2016/08/04
		sqlStr = sqlStr + " ,ceoname='" + ceoname + "'" + VbCrlf
		sqlStr = sqlStr + " ,company_uptae='" + company_uptae+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_upjong='" + company_upjong+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_zipcode='" + company_zipcode+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_address='" + company_address+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_address2='" + company_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_tel='" + company_tel+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_fax='" + company_fax+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode + "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_gubun='" + jungsan_gubun+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_bank='" + jungsan_bank+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_acctname='" + jungsan_acctname+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_acctno='" + jungsan_acctno+ "'" + VbCrlf

		sqlStr = sqlStr + " ,jungsan_date='" + jungsan_date+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_date_off='" + jungsan_date_off+ "'" + VbCrlf

		sqlStr = sqlStr + " ,manager_name='" + manager_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_phone='" + manager_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_hp='" + manager_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_email='" + manager_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_name='" + deliver_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_name='" + jungsan_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_phone='" + jungsan_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_hp='" + jungsan_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_email='" + jungsan_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate = getdate()" + VbCrlf
		sqlStr = sqlStr + " where groupid='" + groupid + "'"
''rw sqlStr
		rsget.Open sqlStr,dbget,1
	else
		sqlStr = "select top 1 groupid from [db_partner].[dbo].tbl_partner_group"
		sqlStr = sqlStr + " order by groupid desc"
		rsget.Open sqlStr,dbget,1
			if rsget.Eof then
				groupid = 1
			else
				groupid = rsget("groupid")
				groupid = Right(groupid,5)
				groupid = CLng(groupid) +1
			end if
		rsget.Close
		groupid = "G" + Format00(5,groupid)

		sqlStr = "insert into [db_partner].[dbo].tbl_partner_group"
		sqlStr = sqlStr + " (groupid, company_name, company_no, ceoname, company_uptae, "
		sqlStr = sqlStr + " company_upjong, company_zipcode, company_address, company_address2, "
		sqlStr = sqlStr + " company_tel, company_fax, return_zipcode, return_address, return_address2, "
		sqlStr = sqlStr + " jungsan_gubun, jungsan_bank, jungsan_date, jungsan_date_off, jungsan_acctname, jungsan_acctno, "
		sqlStr = sqlStr + " manager_name, manager_phone, manager_hp, manager_email, deliver_name, deliver_phone, "
		sqlStr = sqlStr + " deliver_hp, deliver_email, jungsan_name, jungsan_phone, jungsan_hp, jungsan_email, "
		sqlStr = sqlStr + " encCompNo, "
		sqlStr = sqlStr + " lastupdate)"
		sqlStr = sqlStr + " values('" + groupid + "'"
		sqlStr = sqlStr + " ,'" + company_name + "'"
		sqlStr = sqlStr + " ,'" + socialnoReplace(company_no) + "'"   ''2016/08/04
		sqlStr = sqlStr + " ,'" + ceoname + "'"
		sqlStr = sqlStr + " ,'" + company_uptae + "'"
		sqlStr = sqlStr + " ,'" + company_upjong + "'"
		sqlStr = sqlStr + " ,'" + company_zipcode + "'"
		sqlStr = sqlStr + " ,'" + company_address + "'"
		sqlStr = sqlStr + " ,'" + company_address2 + "'"
		sqlStr = sqlStr + " ,'" + company_tel + "'"
		sqlStr = sqlStr + " ,'" + company_fax + "'"
		sqlStr = sqlStr + " ,'" + return_zipcode + "'"
		sqlStr = sqlStr + " ,'" + return_address + "'"
		sqlStr = sqlStr + " ,'" + return_address2 + "'"
		sqlStr = sqlStr + " ,'" + jungsan_gubun + "'"
		sqlStr = sqlStr + " ,'" + jungsan_bank + "'"
		sqlStr = sqlStr + " ,'" + jungsan_date + "'"
		sqlStr = sqlStr + " ,'" + jungsan_date_off + "'"
		sqlStr = sqlStr + " ,'" + jungsan_acctname + "'"
		sqlStr = sqlStr + " ,'" + jungsan_acctno + "'"
		sqlStr = sqlStr + " ,'" + manager_name + "'"
		sqlStr = sqlStr + " ,'" + manager_phone + "'"
		sqlStr = sqlStr + " ,'" + manager_hp + "'"
		sqlStr = sqlStr + " ,'" + manager_email + "'"
		sqlStr = sqlStr + " ,'" + deliver_name + "'"
		sqlStr = sqlStr + " ,'" + deliver_phone + "'"
		sqlStr = sqlStr + " ,'" + deliver_hp + "'"
		sqlStr = sqlStr + " ,'" + deliver_email + "'"
		sqlStr = sqlStr + " ,'" + jungsan_name + "'"
		sqlStr = sqlStr + " ,'" + jungsan_phone + "'"
		sqlStr = sqlStr + " ,'" + jungsan_hp + "'"
		sqlStr = sqlStr + " ,'" + jungsan_email + "'"
		sqlStr = sqlStr + " ,[db_partner].[dbo].[uf_EncSOCNoPH1]('"&company_no&"')"  ''2016/08/04
		sqlStr = sqlStr + " ,getdate()"
		sqlStr = sqlStr + " )"

		dbget.Execute sqlStr

		if (LEN(Trim(replace(company_no,"-","")))=13) then
			sqlStr = "exec [db_cs].[dbo].[usp_Ten_partner_Enc_companyno] '"&groupid&"','"&company_no&"'"
			dbget.Execute sqlStr
		end if
	end if

	'// 아이디 중복 확인
	sqlStr = "select count(*) from [db_partner].[dbo].tbl_partner where id='" & uid & "'"
	rsget.Open sqlStr,dbget,1
	if rsget(0)>0 then
		rsget.Close
        ''업데이트 tbl_partner
        sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
        sqlStr = sqlStr + " set lastInfoChgDT=getdate(), company_name='" + company_name + "'" + VbCrlf
        sqlStr = sqlStr + " ,ceoname='" + ceoname + "'" + VbCrlf
        sqlStr = sqlStr + " ,company_no='" + socialnoReplace(company_no) + "'" + VbCrlf            ''주석처리 2016/08/04 주석제거 2016/08/24
        sqlStr = sqlStr + " ,company_upjong='" + company_upjong + "'" + VbCrlf
        sqlStr = sqlStr + " ,company_uptae='" + company_uptae + "'" + VbCrlf
        sqlStr = sqlStr + " ,zipcode='" + company_zipcode+ "'" + VbCrlf
        sqlStr = sqlStr + " ,address='" + company_address + "'" + VbCrlf
        sqlStr = sqlStr + " ,manager_address='" + company_address2 + "'" + VbCrlf
        sqlStr = sqlStr + " ,tel='" + company_tel + "'" + VbCrlf
        sqlStr = sqlStr + " ,fax='" + company_fax + "'" + VbCrlf
        sqlStr = sqlStr + " ,manager_name='" + manager_name + "'" + VbCrlf
        sqlStr = sqlStr + " ,email='" + manager_email + "'" + VbCrlf
        sqlStr = sqlStr + " ,manager_phone='" + manager_phone + "'" + VbCrlf
        sqlStr = sqlStr + " ,manager_hp='" + manager_hp + "'" + VbCrlf
        sqlStr = sqlStr + " ,deliver_name='" + deliver_name + "'" + VbCrlf
        sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
        sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
        sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_name='" + jungsan_name+ "'" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_phone='" + jungsan_phone+ "'" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_hp='" + jungsan_hp+ "'" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_email='" + jungsan_email+ "'" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_gubun='" + jungsan_gubun+ "'" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_bank='" + jungsan_bank+ "'" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_acctname='" + jungsan_acctname+ "'" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_acctno='" + jungsan_acctno+ "'" + VbCrlf
		sqlStr = sqlStr + " ,groupid='" + groupid+ "'" + VbCrlf
        if (jungsan_date<>"") then
            sqlStr = sqlStr + " ,jungsan_date='" + jungsan_date+ "'" + VbCrlf
        end if

        if (jungsan_date_off<>"") then
            sqlStr = sqlStr + " ,jungsan_date_off='" + jungsan_date_off+ "'" + VbCrlf
            sqlStr = sqlStr + " ,jungsan_date_frn='" + jungsan_date_off+ "'" + VbCrlf
        end if
        sqlStr = sqlStr + " ,return_zipcode='" + p_return_zipcode+ "'" + VbCrlf
        sqlStr = sqlStr + " ,return_address='" + p_return_address+ "'" + VbCrlf
        sqlStr = sqlStr + " ,return_address2='" + p_return_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " ,lastLoginDT=getdate()" + VbCrlf
		sqlStr = sqlStr + " ,lastPwChgDT=getdate()" + VbCrlf
		sqlStr = sqlStr + " ,Enc_password64='" + Enc_userpass64+ "'" + VbCrlf
		sqlStr = sqlStr + " ,Enc_2password64='" + Enc_2userpass64+ "'" + VbCrlf
		if (defaultsongjangdiv<>"") then
            sqlStr = sqlStr + ", defaultsongjangdiv=" + Cstr(defaultsongjangdiv) + vbCrlf
        end if
        sqlStr = sqlStr + " where id='" + Cstr(uid) + "'"
    ''rw sqlStr
        rsget.Open sqlStr,dbget,1
    else
        rsget.Close
        ''insert tbl_partner
        sqlStr = "insert into [db_partner].[dbo].tbl_partner" + vbCrlf
        sqlStr = sqlStr + "(id,Enc_password,Enc_password64,Enc_2password64,lastPwChgDT,lastLoginDT,userdiv,jungsan_date,groupid"+ vbCrlf
        sqlStr = sqlStr + ", deliver_name, deliver_phone, deliver_hp, deliver_email"+ vbCrlf
        sqlStr = sqlStr + ", return_zipcode, return_address, return_address2"+ vbCrlf
        sqlStr = sqlStr + ", purchasetype, offcatecode, offmduserid, selltype"+ vbCrlf
        if (defaultsongjangdiv<>"") then
            sqlStr = sqlStr + ", defaultsongjangdiv"+ vbCrlf
        end if
        sqlStr = sqlStr + ")" + vbCrlf
        sqlStr = sqlStr + " values('" + uid + "'" + vbCrlf
        sqlStr = sqlStr + " ,''" + vbCrlf                       '--  빈값으로변경
        sqlStr = sqlStr + " ,'" + Enc_userpass64 + "'" + vbCrlf '--암호화 고도화 2014/07/21
        sqlStr = sqlStr + " ,'" + Enc_2userpass64 + "'" + vbCrlf '--암호화 고도화 2014/07/21
        sqlStr = sqlStr + " ,getdate(),getdate()" + vbCrlf
		sqlStr = sqlStr + " ,'"&p_userdiv&"'" + vbCrlf
        sqlStr = sqlStr + " ,'" + jungsan_date + "'" + vbCrlf
        sqlStr = sqlStr + " ,'" + groupid + "'" + vbCrlf

        sqlStr = sqlStr + " ,'" + deliver_name + "'" + vbCrlf
        sqlStr = sqlStr + " ,'" + deliver_phone + "'" + vbCrlf
        sqlStr = sqlStr + " ,'" + deliver_hp + "'" + vbCrlf
        sqlStr = sqlStr + " ,'" + deliver_email + "'" + vbCrlf
        sqlStr = sqlStr + " ,'" + p_return_zipcode + "'" + vbCrlf         ''초기 반품주소는 사무실 주소와 동일하게 설정됩니다.
        sqlStr = sqlStr + " ,'" + p_return_address + "'" + vbCrlf
        sqlStr = sqlStr + " ,'" + p_return_address2 + "'" + vbCrlf
		if (Vpurchasetype<>"") then
		sqlStr = sqlStr + " ,"&Vpurchasetype&""+ vbCrlf
		else
		sqlStr = sqlStr + " ,1"+ vbCrlf
		end if
        sqlStr = sqlStr + " ,'" + Voffcatecode + "'" + vbCrlf
        sqlStr = sqlStr + " ,'" + Voffmduserid + "'" + vbCrlf
        sqlStr = sqlStr + " ,0"+ vbCrlf
        if (defaultsongjangdiv<>"") then
            sqlStr = sqlStr + " ,'" & defaultsongjangdiv + "'" + VbCrlf
        end if
        sqlStr = sqlStr + " )"
    'rw sqlStr
    'response.end
        rsget.Open sqlStr,dbget,1
	
		''같은 그룹 업체 업데이트.
		sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
		sqlStr = sqlStr + " set lastInfoChgDT=getdate(), company_name='" + company_name + "'" + VbCrlf
		sqlStr = sqlStr + " ,ceoname='" + ceoname + "'" + VbCrlf
		sqlStr = sqlStr + " ,company_no='" + socialnoReplace(company_no) + "'" + VbCrlf            ''주석처리 2016/08/04 주석제거 2016/08/24
		sqlStr = sqlStr + " ,company_upjong='" + company_upjong + "'" + VbCrlf
		sqlStr = sqlStr + " ,company_uptae='" + company_uptae + "'" + VbCrlf
		sqlStr = sqlStr + " ,zipcode='" + company_zipcode+ "'" + VbCrlf
		sqlStr = sqlStr + " ,address='" + company_address + "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_address='" + company_address2 + "'" + VbCrlf
		sqlStr = sqlStr + " ,tel='" + company_tel + "'" + VbCrlf
		sqlStr = sqlStr + " ,fax='" + company_fax + "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_name='" + manager_name + "'" + VbCrlf
		sqlStr = sqlStr + " ,email='" + manager_email + "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_phone='" + manager_phone + "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_hp='" + manager_hp + "'" + VbCrlf
		''sqlStr = sqlStr + " ,deliver_name='" + deliver_name + "'" + VbCrlf
		''sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		''sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		''sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_name='" + jungsan_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_phone='" + jungsan_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_hp='" + jungsan_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_email='" + jungsan_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_gubun='" + jungsan_gubun+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_bank='" + jungsan_bank+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_acctname='" + jungsan_acctname+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_acctno='" + jungsan_acctno+ "'" + VbCrlf

		if (jungsan_date<>"") then
			sqlStr = sqlStr + " ,jungsan_date='" + jungsan_date+ "'" + VbCrlf
		end if

		if (jungsan_date_off<>"") then
			sqlStr = sqlStr + " ,jungsan_date_off='" + jungsan_date_off+ "'" + VbCrlf
			sqlStr = sqlStr + " ,jungsan_date_frn='" + jungsan_date_off+ "'" + VbCrlf
		end if
		''sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode+ "'" + VbCrlf
		''sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
		''sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " where groupid='" + groupid + "'"
	'rw sqlStr
    'response.end
		rsget.Open sqlStr,dbget,1
	end if

	dim gubun, status, tidx

	gubun="temp"
	status = "2"
	if company_no_img<>"" and jungsan_acctno_img<>"" then gubun="newcompreg"

	sqlStr = "INSERT INTO [db_partner].[dbo].[tbl_partner_temp_info]" & VbCRLF
	sqlStr = sqlStr & "(reguserid,groupid,company_name,ceoname,company_no,jungsan_gubun,company_zipcode,company_address," & VbCRLF
	sqlStr = sqlStr & "company_address2,company_uptae,company_upjong,company_tel,company_fax,return_zipcode,return_address," & VbCRLF
	sqlStr = sqlStr & "return_address2,jungsan_bank,jungsan_acctno,jungsan_acctname,jungsan_date,jungsan_date_off," & VbCRLF
	sqlStr = sqlStr & "manager_name,manager_phone,manager_hp,manager_email,gubun," & VbCRLF
	sqlStr = sqlStr & "jungsan_name,jungsan_phone,jungsan_hp,jungsan_email,status,encCompNo)" & VbCRLF
	sqlStr = sqlStr & " values(" & VbCRLF
	sqlStr = sqlStr & "'" & mduserid & "','" & groupid & "','" & company_name & "','" & ceoname & "'"& VbCRLF
	sqlStr = sqlStr & ",'" & socialnoReplace(company_no) & "','" & jungsan_gubun & "','" & company_zipcode & "'" & VbCRLF
	sqlStr = sqlStr & ",'" & company_address & "','" & company_address2 & "','" & company_uptae & "'" & VbCRLF
	sqlStr = sqlStr & ",'" & company_upjong & "','" & company_tel & "','" & company_fax & "'" & VbCRLF
	sqlStr = sqlStr & ",'" & return_zipcode & "','" & return_address & "','" & return_address2 & "'" & VbCRLF
	sqlStr = sqlStr & ",'" & jungsan_bank & "','" & jungsan_acctno & "','" & jungsan_acctname & "'" & VbCRLF
	sqlStr = sqlStr & ",'" & jungsan_date & "','" & jungsan_date_off & "','" & manager_name & "'" & VbCRLF
	sqlStr = sqlStr & ",'" & manager_phone & "','" & manager_hp & "','" & manager_email & "'" & VbCRLF
	sqlStr = sqlStr & ",'" & gubun & "','" & jungsan_name & "','" & jungsan_phone & "'" & VbCRLF
	sqlStr = sqlStr & ",'" & jungsan_hp & "','" & jungsan_email & "','" & status & "'" & VbCRLF
	sqlStr = sqlStr & ",[db_partner].[dbo].[uf_EncSOCNoPH1]('"&company_no&"')" & VbCRLF    ''2016/08/04 추가
	sqlStr = sqlStr & ")"
	'response.write sqlStr & "<br>"
	dbget.Execute sqlStr

	sqlStr = " SELECT top 1 max(tidx) from [db_partner].[dbo].[tbl_partner_temp_info]"
	rsget.Open sqlStr,dbget,1
	IF Not rsget.EOF THEN
		tidx = rsget(0)
	END IF
	rsget.close

	if gubun="newcompreg" then
		sqlStr = "UPDATE [db_partner].[dbo].[tbl_partner_temp_info] SET " & vbCrLf
		sqlStr = sqlStr & " status = '3', " & vbCrLf
		sqlStr = sqlStr & " confirmuserid = '" & mduserid & "', " & vbCrLf
		sqlStr = sqlStr & " lastupdate = getdate() " & vbCrLf
		sqlStr = sqlStr & " WHERE tidx = '" & tidx & "'"
		rsget.Open sqlStr,dbget,1
	end if

	if (LEN(Trim(replace(company_no,"-","")))=13) then
		sqlStr = "exec [db_cs].[dbo].[usp_Ten_partner_temp_info_Enc_companyno] "&tidx&",'"&company_no&"'"
		dbget.Execute sqlStr
	end if

	sqlStr = "INSERT INTO [db_partner].[dbo].[tbl_partner_temp_makerid](tidx, makerid) VALUES('" & tidx & "','" & uid & "') " & vbCrLf
	IF sqlStr <> "" Then
		dbget.Execute sqlStr
	End IF

	'####### 첨부파일 저장 #######
	if company_no_img<>"" then
	sqlStr = "INSERT INTO [db_partner].[dbo].tbl_partner_temp_file (file_name, real_name, tidx)" & vbCrLf
	sqlStr = sqlStr & " values('" & company_no_img & "','" & company_no_img & "','" & tidx & "')"
	rsget.Open sqlStr,dbget,1
	end if
	if jungsan_acctno_img<>"" then
	sqlStr = " INSERT INTO [db_partner].[dbo].tbl_partner_temp_file (file_name, real_name, tidx)" & vbCrLf
	sqlStr = sqlStr & " values('" & jungsan_acctno_img & "', '" & jungsan_acctno_img & "', '" & tidx & "')"
	rsget.Open sqlStr,dbget,1
	end if

	''============== 전자계약 기본세팅 ==========================================================================================================
	dim chkCT11, chkCT12, oneContract, acctoken, reftoken, ecId, ecPwd, bcompno, oDftCTRPTypeDetail, userStatus
	dim contractType, contractContents, contractName, onoffgubun, subType, APIpath, strParam, ecAUser, ecBUser, ectypeSeq
	dim objXML, iRbody, jsResult, isDefaultContract, ctrKey, ctrNo, bufStr, con_status, con_info, tmpCallBack,strParam1
	dim A_COMPANY_NO, A_UPCHENAME, A_CEONAME, B_COMPANY_NO, B_UPCHENAME, B_CEONAME,DEFAULT_JUNGSANDATE,A_COMPANY_ADDR 
	dim ecCtrSeq, strErrMsg, ENDDATE, chkmwdivMExists, CONTRACT_DATE, B_COMPANY_ADDR, chkmwdivWExists
	bcompno = replace(company_no,"-","")
	chkCT11 = 1
	chkCT12 = 0
	if maeipdiv = "M" then chkCT12 = 1

	sqlStr = "select top 1 ecAUser" +vbcrlf
	sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master" +vbcrlf
	sqlStr = sqlStr & " where ecAUser<>''"
	sqlStr = sqlStr & " order by ctrKey desc"

	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		ecAUser = db2Html(rsget("ecAUser"))
	end if
	rsget.Close

	ecBUser = manager_name

	if signtype ="2" then 
		
		'token 가져오기(db에서)
		set oneContract = new CPartnerContract
			oneContract.fnGetContractToken
			acctoken = oneContract.Facctoken 	
			reftoken = oneContract.Freftoken 
  		set oneContract = nothing
  		
  		'token이 없을때 token 발행
 		if isNull(acctoken) then
 			call sbGetNewToken(ecId,ecPwd)
 			acctoken = Faccess_token
			if acctoken = "" Then
			%>
			<script type="text/javascript" language="javascript">
				alert( "전자계약 계정정보가 잘못입력되었습니다. 확인 후 다시 시도해주세요,");
				location.href = "<%=refer%>";				
			</script>
			<%	response.end					 	
			end if
 		end if	
 				 
 		'회원체크
 		userStatus = fnCheckUser(bcompno,acctoken)

 		if Fchkerror ="invalid_token" then
 			call sbGetRefToken(reftoken)
 			acctoken = Faccess_token
 			userStatus = fnCheckUser(bcompno,acctoken)
 		end if

		if userStatus <> "사용중" then
		%>
			<script type="text/javascript" language="javascript">
				alert( "[<%=userStatus%>]: LG U+ 전자계약 사이트에 가입되어있지 않습니다. 가입 확인 후 계약서 생성이 가능합니다,");
				location.href = "<%=refer%>";				
			</script>
		<% response.end
		end if
		
	    set oDftCTRPTypeDetail = new CPartnerContract
	    oDftCTRPTypeDetail.FRectContractType = DEFAULT_CONTRACTTYPE
    	oDftCTRPTypeDetail.FRectGroupID = groupid
    	oDftCTRPTypeDetail.getContractDetailProtoTypeWithGroupInfo

		A_COMPANY_NO = oDftCTRPTypeDetail.getDefaultValueByKey("$$A_COMPANY_NO$$")
		A_UPCHENAME = oDftCTRPTypeDetail.getDefaultValueByKey("$$A_UPCHENAME$$")
		A_CEONAME = oDftCTRPTypeDetail.getDefaultValueByKey("$$A_CEONAME$$")
		B_COMPANY_NO = oDftCTRPTypeDetail.getDefaultValueByKey("$$B_COMPANY_NO$$")
		B_UPCHENAME = oDftCTRPTypeDetail.getDefaultValueByKey("$$B_UPCHENAME$$")
		B_CEONAME = oDftCTRPTypeDetail.getDefaultValueByKey("$$B_CEONAME$$")
		DEFAULT_JUNGSANDATE = oDftCTRPTypeDetail.getDefaultValueByKey("$$DEFAULT_JUNGSANDATE$$")
		A_COMPANY_ADDR = oDftCTRPTypeDetail.getDefaultValueByKey("$$A_COMPANY_ADDR$$")
		B_COMPANY_ADDR = oDftCTRPTypeDetail.getDefaultValueByKey("$$B_COMPANY_ADDR$$")
		CONTRACT_DATE = oDftCTRPTypeDetail.getDefaultValueByKey("$$CONTRACT_DATE$$")
		ENDDATE = oDftCTRPTypeDetail.getDefaultValueByKey("$$ENDDATE$$")
    end if

    ''==//================================================================================================================================

	''==============계약서 작성 ============================================================================================================
	''기본계약서-----------------------------------------------------------------------------------------------------------------
	if chkCT11 = 1 and chkCT12 < 1 then
		contractType = DEFAULT_CONTRACTTYPE '수기형 계약서번호
		
		sqlStr = "select contractContents, contractName ,onoffgubun, subType" +vbcrlf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType" +vbcrlf
		sqlStr = sqlStr & " where contractType=" & contractType

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			contractContents = db2Html(rsget("contractContents"))
			contractName = db2Html(rsget("contractName"))
			onoffgubun = rsget("onoffgubun")
			subType    = rsget("subType")
		end if
		rsget.Close

		if signtype ="2" then ''전자결재 처리	 		      	
			ectypeSeq = Fec_defctrtype 'lg u+ 계약서번호		
			ecCtrSeq = 0
			APIpath =FecURL&"/api/createCont"

			strParam = "?type_seq="&ectypeSeq&"&cancel_limit=0&contract_dt="&CONTRACT_DATE&"&contract_key=&contract_money=0&expire_dt="&ENDDATE
			strParam = strParam&"&venderno="&A_COMPANY_NO&"&search_word="&server.URLEncode(B_UPCHENAME)&"&start_dt="&CONTRACT_DATE&"&title="&server.URLEncode(contractName)
			strParam = strParam&"&membList[0].company="&server.URLEncode(A_UPCHENAME)&"&membList[0].gubun=A&membList[0].users="&server.URLEncode(ecAUser)&"&membList[0].venderno="&A_COMPANY_NO
			strParam = strParam&"&membList[1].company="&server.URLEncode(B_UPCHENAME)&"&membList[1].gubun=B&membList[1].users="&server.URLEncode(ecBUser)&"&membList[1].venderno="&B_COMPANY_NO
			strParam = strParam&"&usertagList[0].tag_nm=JUNGSAN_DATE&usertagList[0].tag_vl="&server.URLEncode(DEFAULT_JUNGSANDATE)

			Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
					objXML.Open "GET", APIpath&strParam , False
					objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
					objXML.SetRequestHeader "Authorization", "Bearer " & acctoken
					objXML.Send()
					iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
					iRbody= replace(iRbody,"tmpCallBack({","{")
					iRbody = replace(iRbody,"})","}")
					If objXML.Status = "200" Then
						Set jsResult = JSON.parse(iRbody)
							con_status	= jsResult.status
							con_info= jsResult.info 
							if con_status ="succ" Then
								ecCtrSeq = con_info
							else
								if (con_info="001") then
									strErrMsg= "venderno 값 없음"
								elseif (con_info="002")then
									strErrMsg= "type_seq 값 없음"
								elseif (con_info="003")then
									strErrMsg= "title 값 없음" 
								elseif (con_info="004")then
									strErrMsg= "contract_dt 값 없음" 
								elseif (con_info="005")then
									strErrMsg= "rcontract_money 값 없음" 
								elseif (con_info="011")then
									strErrMsg= "membList(계약자 정보) 값 없음" 
								elseif (con_info="012")then
									strErrMsg= "membList(계약자 정보)가 10이상" 
								elseif (con_info="013")then
									strErrMsg= "계약자 구분 A(작성자) 정보와 계약서 본문의 사업자번호 다름" 
								elseif (con_info="014")then
									strErrMsg= "계약자 구분 값이 순차적이지 않음 (A,B,C,D...)" 
								elseif (con_info="015")then
									strErrMsg= "membList.venderno 값없음" 
								elseif (con_info="016")then
									strErrMsg=" membList.company 값없음" 
								elseif (con_info="020")then
									strErrMsg="venderno 에 사용자 존재하지 않음" 
								elseif (con_info="021")then
									strErrMsg=" 해당정보에 대한 문서가 존재하지 않음" 
								elseif (con_info="030")then
									strErrMsg="membList 에서 venderno 에 대한 사용자가 존재하지않음." 
								end if
							end if
					Set jsResult = Nothing
					End If
				Set objXML = Nothing
						
				'On Error Goto 0 
			if ecCtrSeq ="" or ecCtrSeq = 0 Then
				%>
				<script type="text/javascript" language="javascript">
				alert( "전자계약서 생성에 문제가 발생했습니다. 입력값 확인 후 다시 등록해주세요 - <%=strErrMsg%> ");
					location.href = "<%=refer%>";				
				</script>
				<% 
			response.end
			end if
		end if
		
		''기본계약서여부
		isDefaultContract = (subType=0)

		sqlStr = " select * from db_partner.dbo.tbl_partner_ctr_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("groupid") = groupid
		rsget("contractType") = contractType
		rsget("makerid") = CHKIIF(isDefaultContract,"",uid) '' 기본계약서는 사업자 기준 makerid
		rsget("ctrState") = 0  '' 수정중
		rsget("ctrNo") = ""
		rsget("regUserID") = mduserid
		rsget("ecCtrSeq") = ecCtrSeq
		rsget("ecauser") = ecAUser
		rsget("ecbuser") = ecBUser
		rsget.update
			ctrKey = rsget("ctrKey")
		rsget.close

		sqlStr = " insert into db_partner.dbo.tbl_partner_ctr_Detail"
		sqlStr = sqlStr&" (ctrKey,detailKey,detailValue)"
		sqlStr = sqlStr&" select "&ctrKey&",detailKey,"
		sqlStr = sqlStr&" (CASE WHEN detailKey='$$A_CEONAME$$' THEN '"&A_CEONAME&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$A_COMPANY_ADDR$$' THEN '"&A_COMPANY_ADDR&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$A_COMPANY_NO$$' THEN '"&A_COMPANY_NO&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$A_UPCHENAME$$' THEN '"&A_UPCHENAME&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$B_CEONAME$$' THEN '"&B_CEONAME&"'"
		sqlStr = sqlStr&"     WHEN detailKey='$$B_COMPANY_ADDR$$' THEN '"&html2db(B_COMPANY_ADDR)&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$B_COMPANY_NO$$' THEN '"&B_COMPANY_NO&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$B_UPCHENAME$$' THEN '"&B_UPCHENAME&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$CONTRACT_DATE$$' THEN '"&CONTRACT_DATE&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$DEFAULT_JUNGSANDATE$$' THEN '"&DEFAULT_JUNGSANDATE&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$ENDDATE$$' THEN '"&ENDDATE&"'"
		sqlStr = sqlStr&" 	  ELSE '' END)"
		sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_contractDetailType"
		sqlStr = sqlStr&" where contractType="&contractType
		dbget.Execute sqlStr

		ctrNo=CONTRACT_DATE
		bufStr  = CONTRACT_DATE
		bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일"
		contractContents = Replace(contractContents,"$$CONTRACT_DATE$$",bufStr)

		bufStr  = ENDDATE
		bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일" 
		contractContents = Replace(contractContents,"$$ENDDATE$$",bufStr) 

		contractContents = Replace(contractContents,"$$A_CEONAME$$",A_CEONAME)
		contractContents = Replace(contractContents,"$$A_COMPANY_ADDR$$",A_COMPANY_ADDR)
		contractContents = Replace(contractContents,"$$A_COMPANY_NO$$",A_COMPANY_NO)
		contractContents = Replace(contractContents,"$$A_UPCHENAME$$",A_UPCHENAME)
		contractContents = Replace(contractContents,"$$B_CEONAME$$",B_CEONAME)
		contractContents = Replace(contractContents,"$$B_COMPANY_ADDR$$",B_COMPANY_ADDR)
		contractContents = Replace(contractContents,"$$B_COMPANY_NO$$",B_COMPANY_NO)
		contractContents = Replace(contractContents,"$$B_UPCHENAME$$",B_UPCHENAME)
		contractContents = Replace(contractContents,"$$DEFAULT_JUNGSANDATE$$",DEFAULT_JUNGSANDATE)

		ctrNo = TRim(replace(replace(ctrNo," ",""),"-",""))
		ctrNo = ctrNo & "-" & Format00(2,contractType) & "-" & ctrKey

		sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
		sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
		sqlStr = sqlStr & " ,ctrNo='" & ctrNo & "'"
		sqlStr = sqlStr & " ,enddate='"&ENDDATE&"'"
		sqlStr = sqlStr & " ,ctrState=1" ''업체 오픈
        sqlStr = sqlStr & " ,sendUserID='" & mduserid & "'"
        sqlStr = sqlStr & " ,sendDate=getdate()"
		sqlStr = sqlStr & " where ctrKey=" & ctrKey
		dbget.Execute sqlStr
	end if
	'//-----------------------------------------------------------------------------------------------------------------

	'--직매입계약서-----------------------------------------------------------------------------------------------------
	if chkCT12 = 1 then
		contractType = DEFAULT_CONTRACTTYPE_M
		sqlStr = "select contractContents, contractName ,onoffgubun, subType" +vbcrlf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType" +vbcrlf
		sqlStr = sqlStr & " where contractType=" & contractType

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			contractContents = db2Html(rsget("contractContents"))
			contractName = db2Html(rsget("contractName"))
			onoffgubun = rsget("onoffgubun")
			subType    = rsget("subType")
		end if
		rsget.Close

		if signtype="2" then
			ectypeSeq = Fec_defctrtype_M
			ecCtrSeq = 0 
			APIpath =FecURL&"/api/createCont"

			strParam = "?type_seq="&ectypeSeq&"&cancel_limit=0&contract_dt="&CONTRACT_DATE&"&contract_key=&contract_money=0&expire_dt="&ENDDATE
			strParam = strParam&"&venderno="&A_COMPANY_NO&"&search_word="&server.URLEncode(B_UPCHENAME)&"&start_dt="&CONTRACT_DATE&"&title="&server.URLEncode(contractName)
			strParam = strParam&"&membList[0].company="&server.URLEncode(A_UPCHENAME)&"&membList[0].gubun=A&membList[0].users="&server.URLEncode(ecAUser)&"&membList[0].venderno="&A_COMPANY_NO
			strParam = strParam&"&membList[1].company="&server.URLEncode(B_UPCHENAME)&"&membList[1].gubun=B&membList[1].users="&server.URLEncode(ecBUser)&"&membList[1].venderno="&B_COMPANY_NO
			strParam = strParam&"&usertagList[0].tag_nm=JUNGSAN_DATE&usertagList[0].tag_vl="&server.URLEncode(DEFAULT_JUNGSANDATE)
			'On Error Resume Next

			Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
				objXML.Open "GET", APIpath&strParam , False
				objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
				objXML.SetRequestHeader "Authorization", "Bearer " & acctoken
				objXML.Send()
				iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
				iRbody= replace(iRbody,"tmpCallBack({","{")
				iRbody = replace(iRbody,"})","}")

				If objXML.Status = "200" Then
					Set jsResult = JSON.parse(iRbody)
						con_status	= jsResult.status
						con_info= jsResult.info 
						if con_status ="succ" Then
							ecCtrSeq = con_info
						else
							if (con_info="001") then
								strErrMsg= "venderno 값 없음"
							elseif (con_info="002")then
								strErrMsg= "type_seq 값 없음"
							elseif (con_info="003")then
								strErrMsg= "title 값 없음" 
							elseif (con_info="004")then
								strErrMsg= "contract_dt 값 없음" 
							elseif (con_info="005")then
								strErrMsg= "rcontract_money 값 없음" 
							elseif (con_info="011")then
								strErrMsg= "membList(계약자 정보) 값 없음" 
							elseif (con_info="012")then
								strErrMsg= "membList(계약자 정보)가 10이상" 
							elseif (con_info="013")then
								strErrMsg= "계약자 구분 A(작성자) 정보와 계약서 본문의 사업자번호 다름" 
							elseif (con_info="014")then
								strErrMsg= "계약자 구분 값이 순차적이지 않음 (A,B,C,D...)" 
							elseif (con_info="015")then
								strErrMsg= "membList.venderno 값없음" 
							elseif (con_info="016")then
								strErrMsg=" membList.company 값없음" 
							elseif (con_info="020")then
								strErrMsg="venderno 에 사용자 존재하지 않음" 
							elseif (con_info="021")then
								strErrMsg=" 해당정보에 대한 문서가 존재하지 않음" 
							elseif (con_info="030")then
								strErrMsg="membList 에서 venderno 에 대한 사용자가 존재하지않음." 
							end if	
						end if
					Set jsResult = Nothing
				End If
			Set objXML = Nothing

			'On Error Goto 0 
			if ecCtrSeq ="" or ecCtrSeq = 0 Then
				%>
				<script type="text/javascript" language="javascript">
					alert( "전자계약서 생성에 문제가 발생했습니다. 입력값 확인 후 다시 등록해주세요 - <%=strErrMsg%> ");
					location.href = "<%=refer%>";				
				</script>
				<%
			response.end
			end if
		end if
		''기본계약서여부
		isDefaultContract = (subType=0)

		sqlStr = " select * from db_partner.dbo.tbl_partner_ctr_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("groupid") = groupid
		rsget("contractType") = contractType
		rsget("makerid") = CHKIIF(isDefaultContract,"",uid) '' 기본계약서는 사업자 기준 makerid
		rsget("ctrState") = 0  '' 수정중
		rsget("ctrNo") = ""
		rsget("regUserID") = mduserid
		rsget("ecCtrSeq") = ecCtrSeq
		rsget("ecauser") = ecauser
		rsget("ecbuser") = ecbuser
		rsget.update
			ctrKey = rsget("ctrKey")
		rsget.close

		sqlStr = " insert into db_partner.dbo.tbl_partner_ctr_Detail"
		sqlStr = sqlStr&" (ctrKey,detailKey,detailValue)"
		sqlStr = sqlStr&" select "&ctrKey&",detailKey,"
		sqlStr = sqlStr&" (CASE WHEN detailKey='$$A_CEONAME$$' THEN '"&A_CEONAME&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$A_COMPANY_ADDR$$' THEN '"&A_COMPANY_ADDR&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$A_COMPANY_NO$$' THEN '"&A_COMPANY_NO&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$A_UPCHENAME$$' THEN '"&A_UPCHENAME&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$B_CEONAME$$' THEN '"&B_CEONAME&"'"
		sqlStr = sqlStr&"     WHEN detailKey='$$B_COMPANY_ADDR$$' THEN '"&html2db(B_COMPANY_ADDR)&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$B_COMPANY_NO$$' THEN '"&B_COMPANY_NO&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$B_UPCHENAME$$' THEN '"&B_UPCHENAME&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$CONTRACT_DATE$$' THEN '"&CONTRACT_DATE&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$DEFAULT_JUNGSANDATE$$' THEN '"&DEFAULT_JUNGSANDATE&"'"
		sqlStr = sqlStr&" 	  WHEN detailKey='$$ENDDATE$$' THEN '"&ENDDATE&"'"
		sqlStr = sqlStr&" 	  ELSE '' END)"
		sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_contractDetailType"
		sqlStr = sqlStr&" where contractType="&contractType
		dbget.Execute sqlStr

		ctrNo=CONTRACT_DATE
		bufStr  = CONTRACT_DATE
		bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일"
		contractContents = Replace(contractContents,"$$CONTRACT_DATE$$",bufStr)

		bufStr  = ENDDATE
		bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일" 
		contractContents = Replace(contractContents,"$$ENDDATE$$",bufStr) 

		contractContents = Replace(contractContents,"$$A_CEONAME$$",A_CEONAME)
		contractContents = Replace(contractContents,"$$A_COMPANY_ADDR$$",A_COMPANY_ADDR)
		contractContents = Replace(contractContents,"$$A_COMPANY_NO$$",A_COMPANY_NO)
		contractContents = Replace(contractContents,"$$A_UPCHENAME$$",A_UPCHENAME)
		contractContents = Replace(contractContents,"$$B_CEONAME$$",B_CEONAME)
		contractContents = Replace(contractContents,"$$B_COMPANY_ADDR$$",B_COMPANY_ADDR)
		contractContents = Replace(contractContents,"$$B_COMPANY_NO$$",B_COMPANY_NO)
		contractContents = Replace(contractContents,"$$B_UPCHENAME$$",B_UPCHENAME)
		contractContents = Replace(contractContents,"$$DEFAULT_JUNGSANDATE$$",DEFAULT_JUNGSANDATE)
		ctrNo = TRim(replace(replace(ctrNo," ",""),"-",""))
		ctrNo = ctrNo & "-" & Format00(2,contractType) & "-" & ctrKey

		sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
		sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
		sqlStr = sqlStr & " ,ctrNo='" & ctrNo & "'"
		sqlStr = sqlStr & " ,enddate='"&enddate&"'"
		sqlStr = sqlStr & " ,ctrState=1" ''업체 오픈
        sqlStr = sqlStr & " ,sendUserID='" & mduserid & "'"
        sqlStr = sqlStr & " ,sendDate=getdate()"
		sqlStr = sqlStr & " where ctrKey=" & ctrKey

		dbget.Execute sqlStr

	end if
	'//------------------------------------------------------------------------------------------------------------------------

    if (maeipdiv<>"") then
		dim addOF_ctrDate, addON_ctrDate, nMonth, addON_endDate

		if (Now()<"2014-01-01") then
			addON_ctrDate = "2014-01-01"
		else
			addON_ctrDate = Left(Now(),10)  ''Left(Buf,4)+"년 "+Mid(Buf,6,2)+"월 "+Mid(Buf,9,2)+"일" //계약서 내용만 치환
		end if

		nMonth = mid(addON_ctrDate,6,2)
		
		if (nMonth<=3) then
			addON_endDate = year(date())&"-06-30"
		elseif (nMonth>3 and nMonth<=6) then
			addON_endDate = year(date())&"-09-30"
		elseif (nMonth>6 and nMonth<=9) then
			addON_endDate = year(date())&"-12-31"
		elseif (nMonth>9 and nMonth<=12) then
			addON_endDate = year(dateadd("yyyy",1,date())) &"-03-31"
		end if

        SET ogroupInfo = new CPartnerGroup
        ogroupInfo.FRectGroupid = groupid
        if (groupid<>"") then
            ogroupInfo.GetOneGroupInfo
        end if

        if (ogroupInfo.FResultCount<1) then
            SET ogroupInfo = Nothing
            dbget.close()
			%>
			<script type="text/javascript" language="javascript">
				alert( "그룹정보가 없습니다. 입력값 확인 후 다시 등록해주세요 - <%=strErrMsg%> ");
				location.href = "<%=refer%>";				
			</script>
			<%
            response.end
        end if

        if (addOF_ctrDate<>"") and (addON_ctrDate="") then
            addON_ctrDate = addOF_ctrDate
        end if

        ''부속 합의서 등록
        '' 매입계약서인지 체크
        'For kk = 1 To Request.Form("maeipdiv").Count
            chkmwdiv = maeipdiv
			addmwdiv = maeipdiv
            addsellplace = "ON"

			'//전자계약
			if signtype="2" then
				dim defmargin, defdeliver, ismeaip	
				if (chkmwdiv="M")   then '' 매입/ 출고매입
					contractType = ADD_CONTRACTTYPE_M                
					ectypeSeq = Fec_addctrtype_M
					ismeaip ="기본공급율"
					defmargin = (100-CLNG(addmargin*100)/100)&" %"              
				else
					contractType = ADD_CONTRACTTYPE
					ectypeSeq = Fec_addctrtype
					ismeaip ="기본수수료"
					defmargin = (CLNG(addmargin*100)/100)&" %"
				end if

				sqlStr = "select contractContents, contractName ,onoffgubun, subType" &vbcrlf
				sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType" &vbcrlf
				sqlStr = sqlStr & " where contractType=" & contractType

				rsget.Open sqlStr,dbget,1
				if Not rsget.Eof then
					contractContents = db2Html(rsget("contractContents"))
					contractName = db2Html(rsget("contractName"))
					onoffgubun = rsget("onoffgubun")
					subType    = rsget("subType")
				end if
				rsget.Close
					
				''기본계약서여부
				isDefaultContract = (subType=0)
				dim defaultmargin,defaultdeliveryType,defaultFreebeasongLimit,defaultdeliverpay 
				dim mwName
				dim sellplacename
							
				if (addmwdiv="U") and (addON_dlvtype<>"") and (addON_dlvlimit<>"") and (addON_dlvpay<>"") then  
					defaultdeliveryType = addON_dlvtype
					defaultFreebeasongLimit = addON_dlvlimit
					defaultdeliverpay = addON_dlvpay
				end if

				if addsellplace ="ON" then
					if addmwdiv = "M" then
						mwName = "매입"
					elseif addmwdiv ="U" then
						mwName ="업체"
					elseif addmwdiv ="W" then
						mwName ="위탁"
					end if
					sellplacename = "온라인"
				else
					sqlStr = " SELECT comm_name FROM  db_jungsan.dbo.tbl_jungsan_comm_code where comm_cd = '"&addmwdiv&"'"
					rsget.Open sqlStr,dbget,1
					if not rsget.eof then
						mwName = rsget("comm_name")
					end if
					rsget.close
					sqlStr = " SELECT shopname FROM  db_shop.dbo.tbl_shop_user where userid = '"&addsellplace&"'"
					rsget.Open sqlStr,dbget,1
					if not rsget.eof then
						sellplaceName = rsget("shopname")&" 매장"
					end if
					rsget.close
				end if
				A_COMPANY_NO = replace(getDefaultContractValue("$$A_COMPANY_NO$$",ogroupInfo),"-","")
				A_UPCHENAME =getDefaultContractValue("$$A_UPCHENAME$$",ogroupInfo)
				A_CEONAME = getDefaultContractValue("$$A_CEONAME$$",ogroupInfo)
				A_COMPANY_ADDR = getDefaultContractValue("$$A_COMPANY_ADDR$$",ogroupInfo)
				B_COMPANY_NO = replace(getDefaultContractValue("$$B_COMPANY_NO$$",ogroupInfo) ,"-","")
				B_UPCHENAME = getDefaultContractValue("$$B_UPCHENAME$$",ogroupInfo)
				B_CEONAME = getDefaultContractValue("$$B_CEONAME$$",ogroupInfo)
				B_COMPANY_ADDR =getDefaultContractValue("$$B_COMPANY_ADDR$$",ogroupInfo)
				CONTRACT_DATE   =getDefaultContractValue("$$CONTRACT_DATE$$",ogroupInfo)
				ENDDATE   = getDefaultContractValue("$$ENDDATE$$",ogroupInfo)
				ecCtrSeq = 0

				APIpath =FecURL&"/api/createCont"

				strParam = "?type_seq="&ectypeSeq&"&cancel_limit=0&contract_dt="&CONTRACT_DATE&"&contract_key=&contract_money=0&expire_dt="&ENDDATE
				strParam = strParam&"&venderno="&A_COMPANY_NO&"&search_word="&server.URLEncode(B_UPCHENAME)&"&start_dt="&CONTRACT_DATE&"&title="&server.URLEncode(contractName) 
				strParam = strParam&"&membList[0].company="&server.URLEncode(A_UPCHENAME)&"&membList[0].gubun=A&membList[0].users="&server.URLEncode(ecAUser)&"&membList[0].venderno="&A_COMPANY_NO
				strParam = strParam&"&membList[1].company="&server.URLEncode(B_UPCHENAME)&"&membList[1].gubun=B&membList[1].users="&server.URLEncode(ecBUser)&"&membList[1].venderno="&B_COMPANY_NO 
				strParam = strParam&"&usertagList[0].tag_nm=TIT_ISMEAIP"&"&usertagList[0].tag_vl="&server.URLEncode(ismeaip)
				strParam = strParam&"&usertagList[1].tag_nm=VAL_MAKERID"&"&usertagList[1].tag_vl="&server.URLEncode(uid)
				strParam = strParam&"&usertagList[2].tag_nm=VAL_SELLPLACE"&"&usertagList[2].tag_vl="&server.URLEncode(sellplaceName)
				strParam = strParam&"&usertagList[3].tag_nm=VAL_MWDIV"&"&usertagList[3].tag_vl="&server.URLEncode(mwName)
				strParam = strParam&"&usertagList[4].tag_nm=VAL_DEFMARGIN"&"&usertagList[4].tag_vl="&server.URLEncode(defmargin)
				strParam = strParam&"&usertagList[5].tag_nm=VAL_DEFDELIVER"&"&usertagList[5].tag_vl="&server.URLEncode(defdeliver)
				'On Error Resume Next

				Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
					objXML.Open "GET", APIpath&strParam , False
					objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
					objXML.SetRequestHeader "Authorization", "Bearer " & acctoken
					objXML.Send()
					iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
					iRbody= replace(iRbody,"tmpCallBack({","{")
					iRbody = replace(iRbody,"})","}")

					If objXML.Status = "200" Then
						Set jsResult = JSON.parse(iRbody)
							con_status	= jsResult.status
							con_info= jsResult.info 
							if con_status ="succ" Then
								ecCtrSeq = con_info
							else
								if (con_info="001") then
									strErrMsg= "venderno 값 없음"
								elseif (con_info="002")then
									strErrMsg= "type_seq 값 없음"
								elseif (con_info="003")then
									strErrMsg= "title 값 없음" 
								elseif (con_info="004")then
									strErrMsg= "contract_dt 값 없음" 
								elseif (con_info="005")then
									strErrMsg= "rcontract_money 값 없음" 
								elseif (con_info="011")then
									strErrMsg= "membList(계약자 정보) 값 없음" 
								elseif (con_info="012")then
									strErrMsg= "membList(계약자 정보)가 10이상" 
								elseif (con_info="013")then
									strErrMsg= "계약자 구분 A(작성자) 정보와 계약서 본문의 사업자번호 다름" 
								elseif (con_info="014")then
									strErrMsg= "계약자 구분 값이 순차적이지 않음 (A,B,C,D...)" 
								elseif (con_info="015")then
									strErrMsg= "membList.venderno 값없음" 
								elseif (con_info="016")then
									strErrMsg=" membList.company 값없음" 
								elseif (con_info="020")then
									strErrMsg="venderno 에 사용자 존재하지 않음" 
								elseif (con_info="021")then
									strErrMsg=" 해당정보에 대한 문서가 존재하지 않음" 
								elseif (con_info="030")then
									strErrMsg="membList 에서 venderno 에 대한 사용자가 존재하지않음." 
								end if
							end if			
						Set jsResult = Nothing
					End If
				Set objXML = Nothing
									
					'On Error Goto 0 
				if ecCtrSeq ="" or ecCtrSeq = 0 Then
					%>
					<script type="text/javascript" language="javascript">
						alert( "전자계약서 생성에 문제가 발생했습니다. 입력값 확인 후 다시 등록해주세요 - <%=strErrMsg%> ");
						location.href = "<%=refer%>";				
					</script>
					<%
				response.end
				end if
		
				sqlStr = " select * from db_partner.dbo.tbl_partner_ctr_master where 1=0"
				rsget.Open sqlStr,dbget,1,3
				rsget.AddNew
				rsget("groupid") = groupid
				rsget("contractType") = contractType
				rsget("makerid") = CHKIIF(isDefaultContract,"",uid) '' 기본계약서는 사업자 기준 makerid
				rsget("ctrState") = 0  '' 수정중
				rsget("ctrNo") = ""
				rsget("regUserID") = mduserid
				rsget("ecCtrSeq") = ecCtrSeq
				rsget("ecauser") = ecAUser
				rsget("ecbuser") = ecBUser
				rsget.update
					ctrKey = rsget("ctrKey")
				rsget.close

				sqlStr = " insert into db_partner.dbo.tbl_partner_ctr_Detail"
				sqlStr = sqlStr&" (ctrKey,detailKey,detailValue)"
				sqlStr = sqlStr&" select "&ctrKey&",detailKey,"
				sqlStr = sqlStr&" (CASE WHEN detailKey='$$A_CEONAME$$' THEN '"&getDefaultContractValue("$$A_CEONAME$$",ogroupInfo)&"'"
				sqlStr = sqlStr&" 	  WHEN detailKey='$$A_COMPANY_ADDR$$' THEN '"&getDefaultContractValue("$$A_COMPANY_ADDR$$",ogroupInfo)&"'"
				sqlStr = sqlStr&" 	  WHEN detailKey='$$A_COMPANY_NO$$' THEN '"&getDefaultContractValue("$$A_COMPANY_NO$$",ogroupInfo)&"'"
				sqlStr = sqlStr&" 	  WHEN detailKey='$$A_UPCHENAME$$' THEN '"&getDefaultContractValue("$$A_UPCHENAME$$",ogroupInfo)&"'"
				sqlStr = sqlStr&" 	  WHEN detailKey='$$B_CEONAME$$' THEN '"&getDefaultContractValue("$$B_CEONAME$$",ogroupInfo)&"'"
				sqlStr = sqlStr&"     WHEN detailKey='$$B_COMPANY_ADDR$$' THEN '"&html2db(getDefaultContractValue("$$B_COMPANY_ADDR$$",ogroupInfo))&"'"
				sqlStr = sqlStr&" 	  WHEN detailKey='$$B_COMPANY_NO$$' THEN '"&getDefaultContractValue("$$B_COMPANY_NO$$",ogroupInfo)&"'"
				sqlStr = sqlStr&" 	  WHEN detailKey='$$B_UPCHENAME$$' THEN '"&html2db(getDefaultContractValue("$$B_UPCHENAME$$",ogroupInfo))&"'"
				sqlStr = sqlStr&" 	  WHEN detailKey='$$CONTRACT_DATE$$' THEN '"&addON_ctrDate&"'"
				sqlStr = sqlStr&" 	  WHEN detailKey='$$ENDDATE$$' THEN '"&addON_endDate&"'"
				sqlStr = sqlStr&" 	  ELSE '' END)"
				sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_contractDetailType"
				sqlStr = sqlStr&" where contractType="&contractType
				dbget.Execute sqlStr

				sqlStr = " insert into db_partner.dbo.tbl_partner_ctr_Sub"
				sqlStr = sqlStr & " (ctrKey,sellplace,mwdiv,defaultmargin,defaultDeliveryType,defaultFreeBeasongLimit,defaultDeliverPay)"
				sqlStr = sqlStr & " values("&ctrKey
				sqlStr = sqlStr & " ,'"&addsellplace&"'"
				sqlStr = sqlStr & " ,'"&addmwdiv&"'"
				sqlStr = sqlStr & " ,'"&addmargin&"'"
				if (addmwdiv="U") and (addON_dlvtype<>"") and (addON_dlvlimit<>"") and (addON_dlvpay<>"") then
					sqlStr = sqlStr & " ,'"&addON_dlvtype&"'"
					sqlStr = sqlStr & " ,'"&addON_dlvlimit&"'"
					sqlStr = sqlStr & " ,'"&addON_dlvpay&"'"
				else
					sqlStr = sqlStr & " ,NULL"
					sqlStr = sqlStr & " ,0"
					sqlStr = sqlStr & " ,0"
				end if
				sqlStr = sqlStr & ")"
				dbget.Execute sqlStr

				'' 계약서 DB 내용으로 치환
				if  (FillContractContentsByDB(ctrKey, contractContents)) then
					ctrNo = TRim(replace(replace(addON_ctrDate," ",""),"-",""))
					ctrNo = ctrNo & "-" & Format00(2,contractType) & "-" & ctrKey
					sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
					sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
					sqlStr = sqlStr & " ,ctrNo='" & ctrNo & "'"
					sqlStr = sqlStr & " ,enddate='"&addON_endDate&"'"
					sqlStr = sqlStr & " ,ctrState=1" ''업체 오픈
					sqlStr = sqlStr & " ,sendUserID='" & mduserid & "'"
					sqlStr = sqlStr & " ,sendDate=getdate()"
					sqlStr = sqlStr & " where ctrKey=" & ctrKey
					dbget.Execute sqlStr
				else
					response.write "계약서 작성실패"
				end if		       
			else							
				'if ((Not chkmwdivMExists) and ((chkmwdiv="M") or (chkmwdiv="B031"))) or ((Not chkmwdivWExists) and NOT ((chkmwdiv="M") or (chkmwdiv="B031"))) then
				if ((Not chkmwdivMExists) and ( chkmwdiv="M")) or ((Not chkmwdivWExists) and NOT (chkmwdiv="M")) then
					if (chkmwdiv="M")   then '' 매입/ 출고매입
						contractType = ADD_CONTRACTTYPE_M
						chkmwdivMExists = true
						ectypeSeq = Fec_addctrtype_M
					else
						contractType = ADD_CONTRACTTYPE
						chkmwdivWExists = true
						ectypeSeq = Fec_addctrtype
					end if

					sqlStr = "select contractContents, contractName ,onoffgubun, subType" +vbcrlf
					sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType" +vbcrlf
					sqlStr = sqlStr & " where contractType=" & contractType

					rsget.Open sqlStr,dbget,1
					if Not rsget.Eof then
						contractContents = db2Html(rsget("contractContents"))
						contractName = db2Html(rsget("contractName"))
						onoffgubun = rsget("onoffgubun")
						subType    = rsget("subType")
					end if
					rsget.Close

					''기본계약서여부
					isDefaultContract = (subType=0)
					sqlStr = " select * from db_partner.dbo.tbl_partner_ctr_master where 1=0"
					rsget.Open sqlStr,dbget,1,3
					rsget.AddNew
					rsget("groupid") = groupid
					rsget("contractType") = contractType
					rsget("makerid") = CHKIIF(isDefaultContract,"",uid) '' 기본계약서는 사업자 기준 makerid
					rsget("ctrState") = 0  '' 수정중
					rsget("ctrNo") = ""
					rsget("regUserID") = mduserid
					rsget("ecCtrSeq") = ecCtrSeq		
					rsget("ecauser") = ecAUser		
					rsget("ecbuser") = ecBuser	
					rsget.update
						ctrKey = rsget("ctrKey")
					rsget.close

					sqlStr = " insert into db_partner.dbo.tbl_partner_ctr_Detail"
					sqlStr = sqlStr&" (ctrKey,detailKey,detailValue)"
					sqlStr = sqlStr&" select "&ctrKey&",detailKey,"
					sqlStr = sqlStr&" (CASE WHEN detailKey='$$A_CEONAME$$' THEN '"&getDefaultContractValue("$$A_CEONAME$$",ogroupInfo)&"'"
					sqlStr = sqlStr&" WHEN detailKey='$$A_COMPANY_ADDR$$' THEN '"&getDefaultContractValue("$$A_COMPANY_ADDR$$",ogroupInfo)&"'"
					sqlStr = sqlStr&" WHEN detailKey='$$A_COMPANY_NO$$' THEN '"&getDefaultContractValue("$$A_COMPANY_NO$$",ogroupInfo)&"'"
					sqlStr = sqlStr&" WHEN detailKey='$$A_UPCHENAME$$' THEN '"&getDefaultContractValue("$$A_UPCHENAME$$",ogroupInfo)&"'"
					sqlStr = sqlStr&" WHEN detailKey='$$B_CEONAME$$' THEN '"&getDefaultContractValue("$$B_CEONAME$$",ogroupInfo)&"'"
					sqlStr = sqlStr&" WHEN detailKey='$$B_COMPANY_ADDR$$' THEN '"&html2db(getDefaultContractValue("$$B_COMPANY_ADDR$$",ogroupInfo))&"'"
					sqlStr = sqlStr&" WHEN detailKey='$$B_COMPANY_NO$$' THEN '"&getDefaultContractValue("$$B_COMPANY_NO$$",ogroupInfo)&"'"
					sqlStr = sqlStr&" WHEN detailKey='$$B_UPCHENAME$$' THEN '"&html2db(getDefaultContractValue("$$B_UPCHENAME$$",ogroupInfo))&"'"
					sqlStr = sqlStr&" WHEN detailKey='$$CONTRACT_DATE$$' THEN '"&addON_ctrDate&"'"
					sqlStr = sqlStr&" WHEN detailKey='$$ENDDATE$$' THEN '"&addON_endDate&"'"
					sqlStr = sqlStr&" ELSE '' END)"
					sqlStr = sqlStr&" from db_partner.dbo.tbl_partner_contractDetailType"
					sqlStr = sqlStr&" where contractType="&contractType
					dbget.Execute sqlStr

					''-----------------
					if ((chkmwdiv="M" or chkmwdiv="B031") and (addmwdiv="M" or addmwdiv="B031")) or ((chkmwdiv<>"M" and chkmwdiv<>"B031") and (addmwdiv<>"M" and addmwdiv<>"B031")) then
						addsellplace    = "ON"
						sqlStr = " insert into db_partner.dbo.tbl_partner_ctr_Sub"
						sqlStr = sqlStr & " (ctrKey,sellplace,mwdiv,defaultmargin,defaultDeliveryType,defaultFreeBeasongLimit,defaultDeliverPay)"
						sqlStr = sqlStr & " values("&ctrKey
						sqlStr = sqlStr & " ,'"&addsellplace&"'"
						sqlStr = sqlStr & " ,'"&addmwdiv&"'"
						sqlStr = sqlStr & " ,'"&addmargin&"'"
						if (addmwdiv="U") and (addON_dlvtype<>"") and (addON_dlvlimit<>"") and (addON_dlvpay<>"") then
							sqlStr = sqlStr & " ,'"&addON_dlvtype&"'"
							sqlStr = sqlStr & " ,'"&addON_dlvlimit&"'"
							sqlStr = sqlStr & " ,'"&addON_dlvpay&"'"
						else
							sqlStr = sqlStr & " ,NULL"
							sqlStr = sqlStr & " ,0"
							sqlStr = sqlStr & " ,0"
						end if
						sqlStr = sqlStr & ")"
						dbget.Execute sqlStr
					end if

					'' 계약서 DB 내용으로 치환
					if  (FillContractContentsByDB(ctrKey, contractContents)) then
						ctrNo = TRim(replace(replace(addON_ctrDate," ",""),"-",""))
						ctrNo = ctrNo & "-" & Format00(2,contractType) & "-" & ctrKey
						sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
						sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
						sqlStr = sqlStr & " ,ctrNo='" & ctrNo & "'"
						sqlStr = sqlStr & " ,enddate='"&addON_endDate&"'"
						sqlStr = sqlStr & " ,ctrState=1" ''업체 오픈
						sqlStr = sqlStr & " ,sendUserID='" & mduserid & "'"
						sqlStr = sqlStr & " ,sendDate=getdate()"
						sqlStr = sqlStr & " where ctrKey=" & ctrKey
						dbget.Execute sqlStr
					else
						'response.write "계약서 작성실패"
					end if
					''--------------------------
				end if

			end if
        'Next
        SET ogroupInfo = Nothing
    end if
	dim mailfrom, ocontract, oMdInfoList, mailcontent, mailtitle, manageUrl
	if application("Svr_Info")="Dev" then
		manageUrl = "http://testwebadmin.10x10.co.kr"
	else
		manageUrl = "http://webadmin.10x10.co.kr"
	end if
	sqlStr = "select IsNULL(p.usermail,'') as email from db_partner.dbo.tbl_user_tenbyten p"
	sqlStr = sqlStr & " where p.userid='" & mduserid & "'"
	sqlStr = sqlStr & " and p.userid<>''"
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		mailfrom = db2Html(rsget("email"))
	end if
	rsget.Close

	if (manager_hp<>"") then
        '' SMS 발송
        call SendNormalSMS_LINK(manager_hp,"1644-6030","[텐바이텐] 신규 계약서가 발송되었습니다. email 또는 SCM 업체계약관리 메뉴 참조")
    end if

    if (manager_email<>"") then
        '' 이메일 발송
        set ocontract = new CPartnerContract
        ocontract.FPageSize=50
        ocontract.FCurrPage = 1
        ocontract.FRectContractState = 1 ''오픈
        ocontract.FRectGroupID = groupid
        ocontract.FRectCtrKeyArr = ctrKey
        ocontract.GetNewContractList

        set oMdInfoList = new CPartnerContract
        oMdInfoList.FRectGroupID = groupid
        oMdInfoList.FRectContractState = 1 ''오픈
        oMdInfoList.FRectCtrKeyArr = ctrKey
        oMdInfoList.getContractEmailMdList(FALSE)

        mailtitle = "[텐바이텐] 신규 계약서가 발송 되었습니다."

        if signtype="2" then
        	mailcontent   = makeEcCtrMailContents(ocontract,oMdInfoList,False,manageUrl)
        else
        	mailcontent   = makeCtrMailContents(ocontract,oMdInfoList,False)
      	end if

        Call SendMail(mailfrom, manager_email, mailtitle, mailcontent)

        set ocontract=nothing
        set oMdInfoList=nothing
    end if
	''==//================================================================================================================================
    'response.write Err.Description
    'response.end
	If Err.Number = 0 Then
	        dbget.CommitTrans
			dim subject,title,text
			subject = "[" & Cstr(company_name) & "] 업체의 입점 등록이 완료되었습니다."
			title = "[" & Cstr(company_name) & "] 업체의 입점 등록이 완료되었습니다."
			text = "[" & Cstr(company_name) & "] 업체의 입점 등록이 완료되었습니다."
			call SendRadioWebHookMessage(fnGetMemberEmail(mduserid),"Admin",subject,title,text,"")
	Else
	        dbget.RollBackTrans
	        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\n입력한 값들이 너무 길지 않는지 확인바랍니다.\n주로 업태와 업종에서 에러가 자주 나타납니다.')</script>"
	        response.write "<script>history.back()</script>"
	        response.end
	End If
	on error Goto 0
else
	response.write "<script>alert('Error - 구분코드 없음 관리자 문의요망');</script>"
	response.write "<script>location.replace('" + refer + "');</script>"
	response.End
end if
%>

<script>alert('저장되었습니다.');self.close();</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->