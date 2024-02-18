<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim uid,company_name,email,manager_name,address
dim manager_address, tel, fax, userdiv
dim groupid,defaultsongjangdiv
dim txoldpassword, txoldpasswordS, txnewpassword1,txnewpassword2, txnewpasswordS1,txnewpasswordS2

txoldpassword	= requestCheckvar(request("txoldpassword"),32)
txoldpasswordS	= requestCheckvar(request("txoldpasswordS"),32)
txnewpassword1 = requestCheckvar(request("txnewpassword1"),32)
txnewpassword2 = requestCheckvar(request("txnewpassword2"),32)
txnewpasswordS1 = requestCheckvar(request("txnewpasswordS1"),32)
txnewpasswordS2 = requestCheckvar(request("txnewpasswordS2"),32)

groupid		= requestCheckvar(request("groupid"),6)
uid			 = session("ssBctId")
defaultsongjangdiv = requestCheckvar(request("defaultsongjangdiv"),16)

company_name = html2db(request("company_name"))
email		 = html2db(request("email"))
manager_name = html2db(request("manager_name"))
address		 = html2db(request("address"))
manager_address = html2db(request("manager_address"))
tel			= html2db(request("tel"))
fax			= html2db(request("fax"))
userdiv 	= request("userdiv")

dim applyallbrand
applyallbrand 	= requestCheckVar(request("applyallbrand"),100)


dim ceoname, company_no, zipcode, jungsan_gubun
dim jungsan_date,jungsan_bank,jungsan_acctno
dim jungsan_acctname, manager_phone, manager_hp
dim deliver_name, deliver_phone, deliver_email
dim deliver_hp, jungsan_name, jungsan_phone, jungsan_email
dim jungsan_hp, prtidx

ceoname			= html2db(request("ceoname"))
company_no  	= requestCheckVar(request("company_no"),50)
zipcode			= requestCheckVar(request("zipcode"),30)
jungsan_gubun 	= requestCheckVar(request("jungsan_gubun"),100)
jungsan_date 	= requestCheckVar(request("jungsan_date"),100)
jungsan_bank 	= html2db(request("jungsan_bank"))
jungsan_acctno 	= requestCheckVar(request("jungsan_acctno"),200)
jungsan_acctname = requestCheckVar(html2db(request("jungsan_acctname")),100)
manager_phone 	= requestCheckVar(request("manager_phone"),100)
manager_hp 		= requestCheckVar(request("manager_hp"),100)
deliver_name 	= requestCheckVar(html2db(request("deliver_name")),100)
deliver_phone 	= requestCheckVar(request("deliver_phone"),100)
deliver_email 	= requestCheckVar(request("deliver_email"),300)
deliver_hp 		= requestCheckVar(request("deliver_hp"),100)
jungsan_name 	= html2db(request("jungsan_name"))
jungsan_phone 	= requestCheckVar(request("jungsan_phone"),100)
jungsan_email 	= requestCheckVar(request("jungsan_email"),300)
jungsan_hp 		= requestCheckVar(request("jungsan_hp"),100)
prtidx 			= requestCheckVar(request("prtidx"),100)


dim company_zipcode, company_address, company_address2
dim company_tel, company_fax, return_zipcode, return_address, return_address2
dim manager_email
dim cs_name, cs_phone, cs_hp, cs_email

company_zipcode = requestCheckVar(request("company_zipcode"),20)
company_address = requestCheckVar(request("company_address"),300)
company_address2 = requestCheckVar(request("company_address2"),300)
company_tel = requestCheckVar(request("company_tel"),100)
company_fax = requestCheckVar(request("company_fax"),100)
return_zipcode = requestCheckVar(request("return_zipcode"),20)
return_address = requestCheckVar(request("return_address"),300)
return_address2 = requestCheckVar(request("return_address2"),300)
manager_email = requestCheckVar(request("manager_email"),300)

cs_name = requestCheckVar(html2db(request("cs_name")),100)
cs_phone = requestCheckVar(html2db(request("cs_phone")),100)
cs_hp = requestCheckVar(html2db(request("cs_hp")),100)
cs_email = requestCheckVar(html2db(request("cs_email")),300)


if not IsNumeric(prtidx) then prtidx=9999

dim company_upjong,company_uptae
company_upjong  = requestCheckVar(html2db(request("company_upjong")),100)
company_uptae   = requestCheckVar(html2db(request("company_uptae")),150)

dim subid
subid   = requestCheckVar(request("subid"),100)

dim mode
mode = requestCheckVar(request("mode"),30)

dim commission,password
commission = requestCheckVar(request("commission"),20)
password = requestCheckVar(request("password"),150)

dim socname_kor, socname, isusing, isextusing, streetusing
dim extstreetusing, specialbrand, maeipdiv, defaultmargine

socname_kor  = requestCheckVar(html2db(request("socname_kor")),200)
socname		 = requestCheckVar(html2db(request("socname")),200)
isusing		 = requestCheckVar(request("isusing"),20)
isextusing	 = requestCheckVar(request("isextusing"),20)
streetusing	 = requestCheckVar(request("streetusing"),20)
extstreetusing	 = requestCheckVar(request("extstreetusing"),20)
specialbrand	 = requestCheckVar(request("specialbrand"),100)
maeipdiv		 = requestCheckVar(request("maeipdiv"),30)
defaultmargine	 = requestCheckVar(request("defaultmargine"),100)

dim sqlStr, idExists
dim Enc_userpass, Enc_userpass64,Enc_2userpass64

Enc_userpass = MD5(txnewpassword1)
Enc_userpass64 = SHA256(MD5(txnewpassword1)) 
Enc_2userpass64 = SHA256(MD5(txnewpasswordS1))

dim opartner, makerid, idcheck


'### 아이디 & 그룹코드 체킹
sqlStr = "select count(*) from [db_partner].[dbo].tbl_partner where id = '" & uid & "' and groupid = '" & groupid & "'"
rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
idcheck = rsget(0)
rsget.close
if idcheck < 1 then
	response.write "<script>alert('Error - ID와 그룹코드 불일치 관리자 문의요망');</script>"
	response.write "<script>location.replace('" + refer + "');</script>"
	dbget.close()	:	response.End
end if


if mode="groupedit" then
	if (groupid<>"") then
		sqlStr = "update [db_partner].[dbo].tbl_partner_group" + VbCrlf
		sqlStr = sqlStr + " set company_uptae='" + company_uptae+ "'" + VbCrlf
		''sqlStr = sqlStr + " ,company_no='" + company_no + "'" + VbCrlf
		''sqlStr = sqlStr + " ,ceoname='" + ceoname + "'" + VbCrlf
		''sqlStr = sqlStr + " ,company_name='" + company_name + "'" + VbCrlf
		sqlStr = sqlStr + " ,company_upjong='" + company_upjong+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_zipcode='" + company_zipcode+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_address='" + company_address+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_address2='" + company_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_tel='" + company_tel+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_fax='" + company_fax+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode + "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
		'''sqlStr = sqlStr + " ,jungsan_gubun='" + jungsan_gubun+ "'" + VbCrlf
		''정산정보 직접 수정불가
		''sqlStr = sqlStr + " ,jungsan_bank='" + jungsan_bank+ "'" + VbCrlf
		''sqlStr = sqlStr + " ,jungsan_acctname='" + jungsan_acctname+ "'" + VbCrlf
		''sqlStr = sqlStr + " ,jungsan_acctno='" + jungsan_acctno+ "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_name='" + manager_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_phone='" + manager_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_hp='" + manager_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_email='" + manager_email+ "'" + VbCrlf
		'배송담당자 정보는 브랜드별로만 수정가능(skyer9)
		'sqlStr = sqlStr + " ,deliver_name='" + deliver_name+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_name='" + jungsan_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_phone='" + jungsan_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_hp='" + jungsan_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_email='" + jungsan_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate = getdate()" + VbCrlf
		sqlStr = sqlStr + " where groupid='" + groupid + "'"
'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		''같은 그룹 업체 업데이트.
		sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
		sqlStr = sqlStr + " set company_upjong='" + company_upjong + "'" + VbCrlf
		''sqlStr = sqlStr + " ,ceoname='" + ceoname + "'" + VbCrlf
		''sqlStr = sqlStr + " ,company_no='" + company_no + "'" + VbCrlf
		''sqlStr = sqlStr + " ,company_name='" + company_name + "'" + VbCrlf
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
		''반품주소 및 브랜드별 배송담당자는 브랜드별로 수정가능
		'sqlStr = sqlStr + " ,deliver_name='" + deliver_name + "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_name='" + jungsan_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_phone='" + jungsan_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_hp='" + jungsan_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_email='" + jungsan_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,lastInfoChgDT = getdate()" + VbCrlf
		'''sqlStr = sqlStr + " ,jungsan_gubun='" + jungsan_gubun+ "'" + VbCrlf
		''정산정보 직접 수정불가
		''sqlStr = sqlStr + " ,jungsan_bank='" + jungsan_bank+ "'" + VbCrlf
		''sqlStr = sqlStr + " ,jungsan_acctname='" + jungsan_acctname+ "'" + VbCrlf
		''sqlStr = sqlStr + " ,jungsan_acctno='" + jungsan_acctno+ "'" + VbCrlf
		''반품주소 및 브랜드별 배송담당자는 브랜드별로 수정가능
		'sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
		''반품주소 및 브랜드별 배송담당자는 브랜드별로 수정가능
		'if (defaultsongjangdiv<>"") then
		'    sqlStr = sqlStr + " ,defaultsongjangdiv='" + defaultsongjangdiv+ "'" + VbCrlf
		'end if
		sqlStr = sqlStr + " where groupid='" + groupid + "'"

		rsget.Open sqlStr,dbget,1

	else
		response.write "<script>alert('Error - 그룹코드 없음 관리자 문의요망');</script>"
		response.write "<script>location.replace('" + refer + "');</script>"
		dbget.close()	:	response.End
	end if
elseif mode="brandedit" then


		sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
		sqlStr = sqlStr + " set deliver_name='" + deliver_name + "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " ,defaultsongjangdiv='" + defaultsongjangdiv+ "'" + VbCrlf
		sqlStr = sqlStr + " where groupid='" + groupid + "'"

		if (applyallbrand = "Y") then
			'같은 그룹 업체 업데이트.
		else
			sqlStr = sqlStr + " and id='" + uid + "'"
		end if
		rsget.Open sqlStr,dbget,1

		'그룹 반품담당자 정보를 가장 최근 브랜드등록정보로 덮어쒸운다.(skyer9)
		'과거 데이타를 그대로 두는것보다 그냥 덮어 쒸우는게 낫다.
		sqlStr = "update [db_partner].[dbo].tbl_partner_group" + VbCrlf
		sqlStr = sqlStr + " set deliver_name='" + deliver_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate = getdate()" + VbCrlf
		sqlStr = sqlStr + " where groupid='" + groupid + "'"
		rsget.Open sqlStr,dbget,1

elseif mode="modifyreturnaddress" then

		'로그인 아이디가 속한 그룹코드의 다른 브랜드정보 수정
		set opartner = new CPartnerUser

		opartner.FCurrpage = 1
		opartner.FRectDesignerID = session("ssBctId")
		opartner.FPageSize = 1
		opartner.GetOnePartnerNUser

		groupid = opartner.FOneItem.FGroupid

		makerid = requestCheckVar(request("makerid"),32)

		sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
		sqlStr = sqlStr + " set deliver_name='" + deliver_name + "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " ,defaultsongjangdiv='" + defaultsongjangdiv+ "'" + VbCrlf
		sqlStr = sqlStr + " where id='" + makerid + "' and groupid = '" + opartner.FOneItem.FGroupid + "' "
		'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		sqlStr = " IF EXISTS (SELECT TOP 1 brandid FROM [db_cs].[dbo].tbl_cs_brand_memo WHERE brandid = '" + CStr(makerid) + "') "
		sqlStr = sqlStr + " BEGIN "
		sqlStr = sqlStr + " 	UPDATE [db_cs].[dbo].tbl_cs_brand_memo "
		sqlStr = sqlStr + " 	set cs_name = '" + CStr(cs_name) + "', cs_phone = '" + CStr(cs_phone) + "', cs_hp = '" + CStr(cs_hp) + "', cs_email = '" + CStr(cs_email) + "', cs_modifyday = getdate(), cs_reguserid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + " 	WHERE brandid = '" + CStr(makerid) + "' "
		sqlStr = sqlStr + " END "
		sqlStr = sqlStr + " ELSE "
		sqlStr = sqlStr + " BEGIN "
		sqlStr = sqlStr + " 	INSERT INTO [db_cs].[dbo].tbl_cs_brand_memo(brandid, cs_name, cs_phone, cs_hp, cs_email, cs_modifyday, cs_reguserid) "
		sqlStr = sqlStr + " 	VALUES('" + CStr(makerid) + "', '" + CStr(cs_name) + "', '" + CStr(cs_phone) + "', '" + CStr(cs_hp) + "', '" + CStr(cs_email) + "', getdate(), '" + CStr(session("ssBctId")) + "') "
		sqlStr = sqlStr + " END "
		'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		'그룹 반품담당자 정보를 가장 최근 브랜드등록정보로 덮어쒸운다.(skyer9)
		'과거 데이타를 그대로 두는것보다 그냥 덮어 쒸우는게 낫다.
		sqlStr = "update [db_partner].[dbo].tbl_partner_group" + VbCrlf
		sqlStr = sqlStr + " set deliver_name='" + deliver_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate = getdate()" + VbCrlf
		sqlStr = sqlStr + " where groupid='" + groupid + "'"
		rsget.Open sqlStr,dbget,1

elseif mode="editpass" then
	if (txnewpassword1<>txnewpassword2) then
		response.write "<script>alert('패스워드가 일치하지 않습니다.');</script>"
		response.write "<script>location.replace('" + refer + "');</script>"
		dbget.close()	:	response.End
	end if

	'//패스워드 정책 검사(2008.12.15;허진원)
	if chkPasswordComplex(uid,txnewpassword1)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(uid,txnewpassword1) & "\n비밀번호를 확인후 다시 시도해주세요.');" &vbCrLf &_
						" 	location.replace('" & refer & "');" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	
	'//패스워드 정책 검사(2008.12.15;허진원)
	if chkPasswordComplex(uid,txnewpasswordS1)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(uid,txnewpasswordS1) & "\n비밀번호를 확인후 다시 시도해주세요.');" &vbCrLf &_
						" 	location.replace('" & refer & "');" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
    
    sqlStr = "select   id " & vbCrlf
	sqlStr = sqlStr & "from [db_partner].[dbo].tbl_partner   " & vbCrlf 
	sqlStr = sqlStr & "where  id = '" & uid & "' " & vbCrlf
	sqlStr = sqlStr & "and  Enc_password64 = '" & SHA256(md5(txoldpassword)) & "' " & vbCrlf  
	sqlStr = sqlStr & "and  (Enc_2password64 = '" & SHA256(md5(txoldpasswordS)) & "' or Enc_2password64 is Null) "
	rsget.Open sqlStr,dbget,1
	if (rsget.EOF or rsget.BOF) then
		txnewpassword1 = ""
	end if
	rsget.Close

	if (txnewpassword1 = "") then
		response.write "<script>alert('기존 비밀번호가 잘못 입력되었습니다.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End '' 추가 2014/07/15
	end if
	
	''약간수정 2016/10/19
	sqlStr = " IF Not Exists(select * from [db_user].[dbo].tbl_user_n where userid='"&uid&"')" + VbCrlf
    sqlStr = sqlStr + " BEGIN "
	sqlStr = sqlStr + "     update L" + VbCrlf
	sqlStr = sqlStr + "     set  Enc_userpass64='" + Enc_userpass64 + "'" + VbCrlf
	sqlStr = sqlStr + "     , Enc_userpass=''" + VbCrlf	
	sqlStr = sqlStr + "     from [db_user].[dbo].tbl_logindata L" + VbCrlf	
	sqlStr = sqlStr + "         inner Join [db_user].[dbo].tbl_user_c C" + VbCrlf	
	sqlStr = sqlStr + "         on L.userid=C.userid" + VbCrlf	
	sqlStr = sqlStr + "     where L.userid='" + uid + "'"  + VbCrlf	
	sqlStr = sqlStr + " END "+ VbCrlf	
	dbget.Execute sqlStr

	sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
	sqlStr = sqlStr + " set Enc_password64='" + Enc_userpass64 + "'" + VbCrlf
	sqlStr = sqlStr + " , Enc_password=''" + VbCrlf
	sqlStr = sqlStr + " ,  Enc_2password64='" + Enc_2userpass64 + "'" + VbCrlf 
	sqlStr = sqlStr + " where id='" + uid + "'" 
	dbget.Execute sqlStr
	
	''최종 로그인 일자 저장 //2014/07/14 '' tbl_user_tenbyten 사번로그인 제외
    sqlStr = "exec db_partner.dbo.sp_Ten_Add_PartnerLoginLog '"&uid&"','"&Left(request.ServerVariables("REMOTE_ADDR"),16)&"','R','',0"
    dbget.Execute sqlStr
else
	response.write "<script>alert('Error - 구분코드 없음 관리자 문의요망');</script>"
	response.write "<script>location.replace('" + refer + "');</script>"
	dbget.close()	:	response.End
end if
%>

<script>alert('저장되었습니다.');</script>
<script>location.replace('<%= refer %>');</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->