<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
session.codePage = 65001
 %>
<%
'####################################################
' Description :  공지사항 뷰
' History : 이상구 생성
'           2018.07.12 한용민 수정(ISMS대응 권한체크)
'####################################################
%>
<!-- #include virtual="/lib/util/htmllib_UTF8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/function_utf8.asp" -->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/classes/partners/new_partnerusercls.asp"-->

<% 
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim uid,company_name,email,manager_name,address
dim manager_address, tel, fax, userdiv
dim groupid,defaultsongjangdiv, c_userdiv, p_userdiv, pcuserdiv, mduserid, idx
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

ceoname			= requestCheckVar(html2db(request("ceoname")),50)
company_no  	= requestCheckVar(trim(request("company_no")),20)
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
jungsan_name 	= requestCheckVar(html2db(trim(request("jungsan_name"))),50)
jungsan_phone 	= requestCheckVar(trim(request("jungsan_phone")),50)
jungsan_email 	= requestCheckVar(trim(request("jungsan_email")),150)
jungsan_hp 		= requestCheckVar(trim(request("jungsan_hp")),50)
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

dim commission, password	',passwordS
dim Enc_userpass, Enc_userpass64,Enc_2userpass64

commission = request("commission")
password = requestCheckVar(request("password"),32)
'passwordS = requestCheckVar(request("passwordS"),32)

Enc_userpass = MD5(password)
Enc_userpass64 = SHA256(MD5(password))
'Enc_2userpass64= SHA256(MD5(passwordS))

'####### 직원 연락처 Get. 웹훅 발송용. ########
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
                        "</script>"
        dbget.close()	:	session.codePage = 949 : response.End
    end if
    
    ' if chkPasswordComplex(uid,passwordS)<>"" then
    '     response.write "<script language='javascript'>" &vbCrLf &_
    '                     "	alert('" & chkPasswordComplex(uid,passwordS) & "\n2차 비밀번호를 확인후 다시 시도해주세요.');" &vbCrLf &_
    '                     "</script>"
    '     dbget.close()	:	session.codePage = 949 : response.End
    ' end if
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
dim opartner, makerid

''rw mode
On Error Resume Next
dbget.beginTrans

if mode="addnewupchebrand" then
    if checkNotValidHTML(company_zipcode) or checkNotValidHTML(company_address) or checkNotValidHTML(company_address2) or checkNotValidHTML(company_uptae) or checkNotValidHTML(company_upjong) then
    	response.write "<script>alert('사업자등록정보에 사용하실수 없는 태그가 있습니다.');</script>"
		session.codePage = 949
        response.end
    end if
    if checkNotValidHTML(company_tel) or checkNotValidHTML(company_fax) or checkNotValidHTML(return_zipcode) or checkNotValidHTML(return_address) or checkNotValidHTML(return_address2) then
    	response.write "<script>alert('파트너 기본정보에 사용하실수 없는 태그가 있습니다.');</script>"
		session.codePage = 949
        response.end
    end if
    if checkNotValidHTML(manager_name) or checkNotValidHTML(manager_phone) or checkNotValidHTML(manager_email) or checkNotValidHTML(manager_hp) or checkNotValidHTML(jungsan_name) or checkNotValidHTML(jungsan_phone) or checkNotValidHTML(jungsan_email) or checkNotValidHTML(jungsan_hp) then
    	response.write "<script>alert('파트너 담당자정보에 사용하실수 없는 태그가 있습니다.');</script>"
		session.codePage = 949
        response.end
    end if

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

	' 정산담당자 데이터 필수값으로.. 정산시 곤란
	if jungsan_name="" or isnull(jungsan_name) then
		response.write "<script type='text/javascript'>alert('정산담당자명을 입력해 주세요.');</script>"
		session.codePage = 949
		dbget.Close() : response.end
	end if
	if jungsan_phone="" or isnull(jungsan_phone) then
		response.write "<script type='text/javascript'>alert('정산담당자 전화번호를 입력해 주세요.');</script>"
		session.codePage = 949
		dbget.Close() : response.end
	end if
	if jungsan_hp="" or isnull(jungsan_hp) then
		response.write "<script type='text/javascript'>alert('정산담당자 이메일주소를 입력해 주세요.');</script>"
		session.codePage = 949
		dbget.Close() : response.end
	end if
	if jungsan_email="" or isnull(jungsan_email) then
		response.write "<script type='text/javascript'>alert('정산담당자 휴대폰번호를 입력해 주세요.');</script>"
		session.codePage = 949
		dbget.Close() : response.end
	end if

	'파트너 테이블 정산일 정보 가져오기
	sqlStr = "select top 1 isnull(jungsan_date,'') as jungsan_date, isnull(jungsan_date_off,'') as jungsan_date_off from [db_partner].[dbo].tbl_partner"
	sqlStr = sqlStr + " where id='" & Cstr(uid) & "'"
	rsget.Open sqlStr,dbget,1
		if not rsget.Eof then
			jungsan_date = rsget("jungsan_date")
			jungsan_date_off = rsget("jungsan_date_off")
		end if
	rsget.Close

	''insert tbl_logindata
	sqlStr = "update [db_user].[dbo].tbl_logindata"
	sqlStr = sqlStr + " set Enc_userpass64='" + (Enc_userpass64) + "'" + vbCrlf
	sqlStr = sqlStr + " where userid='" & Cstr(uid) & "'"
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
			session.codePage = 949
	        response.end
	End If
	on error Goto 0
else
	response.write "<script>alert('Error - 구분코드 없음 관리자 문의요망');</script>"
	session.codePage = 949
	response.End
end if
%>

<script>alert('저장되었습니다.');window.open('about:blank','_parent').parent.close();</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<% session.codePage = 949 %>