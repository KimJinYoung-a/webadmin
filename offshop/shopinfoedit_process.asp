<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim uid,company_name,email,manager_name,address
dim manager_address, tel, fax, userdiv
dim groupid,defaultsongjangdiv
dim txoldpassword,txoldpasswordS,txnewpassword1,txnewpassword2, txnewpasswordS1,txnewpasswordS2

txoldpassword	= requestCheckvar(request("txoldpassword"),32)
txoldpasswordS	= requestCheckvar(request("txoldpasswordS"),32)
txnewpassword1 = requestCheckvar(request("txnewpassword1"),32)
txnewpassword2 = requestCheckvar(request("txnewpassword2"),32)
txnewpasswordS1 = requestCheckvar(request("txnewpasswordS1"),32)
txnewpasswordS2 = requestCheckvar(request("txnewpasswordS2"),32)

groupid		= request("groupid")
uid			 = session("ssBctId")
defaultsongjangdiv = request("defaultsongjangdiv")

company_name = html2db(request("company_name"))
email		 = html2db(request("email"))
manager_name = html2db(request("manager_name"))
address		 = html2db(request("address"))
manager_address = html2db(request("manager_address"))
tel			= html2db(request("tel"))
fax			= html2db(request("fax"))
userdiv 	= request("userdiv")

dim ceoname, company_no, zipcode, jungsan_gubun
dim jungsan_date,jungsan_bank,jungsan_acctno
dim jungsan_acctname, manager_phone, manager_hp
dim deliver_name, deliver_phone, deliver_email
dim deliver_hp, jungsan_name, jungsan_phone, jungsan_email
dim jungsan_hp, prtidx

ceoname			= html2db(request("ceoname"))
company_no  	= request("company_no")
zipcode			= request("zipcode")
jungsan_gubun 	= request("jungsan_gubun")
jungsan_date 	= request("jungsan_date")
jungsan_bank 	= html2db(request("jungsan_bank"))
jungsan_acctno 	= request("jungsan_acctno")
jungsan_acctname = html2db(request("jungsan_acctname"))
manager_phone 	= request("manager_phone")
manager_hp 		= request("manager_hp")
deliver_name 	= html2db(request("deliver_name"))
deliver_phone 	= request("deliver_phone")
deliver_email 	= request("deliver_email")
deliver_hp 		= request("deliver_hp")
jungsan_name 	= html2db(request("jungsan_name"))
jungsan_phone 	= request("jungsan_phone")
jungsan_email 	= request("jungsan_email")
jungsan_hp 		= request("jungsan_hp")
prtidx 			= request("prtidx")


dim company_zipcode, company_address, company_address2
dim company_tel, company_fax, return_zipcode, return_address, return_address2
dim manager_email

company_zipcode = request("company_zipcode")
company_address = request("company_address")
company_address2 = request("company_address2")
company_tel = request("company_tel")
company_fax = request("company_fax")
return_zipcode = request("return_zipcode")
return_address = request("return_address")
return_address2 = request("return_address2")
manager_email = request("manager_email")


if not IsNumeric(prtidx) then prtidx=9999

dim company_upjong,company_uptae
company_upjong  = html2db(request("company_upjong"))
company_uptae   = html2db(request("company_uptae"))

dim subid
subid   = request("subid")

dim mode
mode = request("mode")

dim commission,password
commission = request("commission")
password = request("password")

dim socname_kor, socname, isusing, isextusing, streetusing
dim extstreetusing, specialbrand, maeipdiv, defaultmargine

socname_kor  = html2db(request("socname_kor"))
socname		 = html2db(request("socname"))
isusing		 = request("isusing")
isextusing	 = request("isextusing")
streetusing	 = request("streetusing")
extstreetusing	 = request("extstreetusing")
specialbrand	 = request("specialbrand")
maeipdiv		 = request("maeipdiv")
defaultmargine	 = request("defaultmargine")

dim shopname
shopname    = html2db(request("shopname"))
dim sqlStr, idExists
dim Enc_userpass, Enc_password64, Enc_2password64

Enc_userpass = MD5(txnewpassword1)
Enc_password64 = SHA256(MD5(txnewpassword1))
Enc_2password64 = SHA256(MD5(txnewpasswordS1))

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
'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		''같은 그룹 업체 업데이트.
		sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
		sqlStr = sqlStr + " set lastInfoChgDT=getdate(), company_upjong='" + company_upjong + "'" + VbCrlf
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
		sqlStr = sqlStr + " ,deliver_name='" + deliver_name + "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_name='" + jungsan_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_phone='" + jungsan_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_hp='" + jungsan_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_email='" + jungsan_email+ "'" + VbCrlf
		'''sqlStr = sqlStr + " ,jungsan_gubun='" + jungsan_gubun+ "'" + VbCrlf
		''정산정보 직접 수정불가
		''sqlStr = sqlStr + " ,jungsan_bank='" + jungsan_bank+ "'" + VbCrlf
		''sqlStr = sqlStr + " ,jungsan_acctname='" + jungsan_acctname+ "'" + VbCrlf
		''sqlStr = sqlStr + " ,jungsan_acctno='" + jungsan_acctno+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf

		if (defaultsongjangdiv<>"") then
		    sqlStr = sqlStr + " ,defaultsongjangdiv='" + defaultsongjangdiv+ "'" + VbCrlf
		end if
		sqlStr = sqlStr + " where groupid='" + groupid + "'"

		rsget.Open sqlStr,dbget,1


        sqlStr = "update [db_shop].[dbo].tbl_shop_user" + VbCrlf
    	sqlStr = sqlStr + " set shopname='" + shopname + "'," + VbCrlf
    	sqlStr = sqlStr + " shopphone='" + company_tel + "'," + VbCrlf
    	sqlStr = sqlStr + " shopzipcode='" + company_zipcode + "'," + VbCrlf
    	sqlStr = sqlStr + " shopaddr1='" + company_address + "'," + VbCrlf
    	sqlStr = sqlStr + " shopaddr2='" + company_address2 + "'," + VbCrlf
    	sqlStr = sqlStr + " manname='" + manager_name + "'," + VbCrlf
    	sqlStr = sqlStr + " manhp='" + manager_hp + "'," + VbCrlf
    	sqlStr = sqlStr + " manphone='" + manager_phone + "'," + VbCrlf
    	sqlStr = sqlStr + " manemail='" + manager_email + "'," + VbCrlf
    	sqlStr = sqlStr + " shopsocno='" + company_no + "'," + VbCrlf
    	sqlStr = sqlStr + " shopceoname='" + ceoname + "'" + VbCrlf
    	sqlStr = sqlStr + " where userid='" + uid + "'" + VbCrlf

    	rsget.Open sqlStr,dbget,1

	else
		response.write "<script>alert('Error - 그룹코드 없음 관리자 문의요망');</script>"
		response.write "<script>location.replace('" + refer + "');</script>"
		dbget.close()	:	response.End
	end if
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
	sqlStr = sqlStr & "and  Enc_password64 = '" & SHA256(md5(txoldpassword)) & "' " & vbCrlf  '''sqlStr = sqlStr & "and  Enc_password = '" & md5(txoldpassword) & "' " & vbCrlf
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
	
''  프론트 정보는 수정안함.
''	sqlStr = "update [db_user].[dbo].tbl_logindata" + VbCrlf
''	sqlStr = sqlStr + " set userpass='" + txnewpassword1 + "'" + VbCrlf
''	sqlStr = sqlStr + " , Enc_userpass='" + Enc_userpass + "'" + VbCrlf
''	sqlStr = sqlStr + " where userid='" + uid + "'"
''
''	rsget.Open sqlStr,dbget,1

	sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
	sqlStr = sqlStr + " set lastInfoChgDT=getdate(), Enc_password64='" + Enc_password64 + "'" + VbCrlf
	sqlStr = sqlStr + " ,Enc_password='" + Enc_userpass + "'" + VbCrlf
	sqlStr = sqlStr + " ,Enc_2password64='" + Enc_2password64 + "'" + VbCrlf 
	sqlStr = sqlStr + " where id='" + uid + "'"

	rsget.Open sqlStr,dbget,1

	' sqlStr = "update [db_shop].[dbo].tbl_shop_user" + VbCrlf
    ' sqlStr = sqlStr + " set Enc_shoppass='" + Enc_userpass + "'" + VbCrlf
    ' sqlStr = sqlStr + " where userid='" + uid + "'" + VbCrlf

    ' dbget.Execute sqlStr & "<Br>"
    
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