<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드 정보 저장
' History : 2018.08.03	정태훈
'			2022.02.24 한용민 수정(일반(간이)사업자, 원천징수, 해외사업자 체크생성후 저장하는 로직 생성)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/lib/util/base64unicode.asp"-->
<%

''response.write "<script>alert('소스보기로 실행결과를 볼 수 없다. 경고창을 띄우자.');</script>"

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim uid,company_name,manager_name,address
dim manager_address, tel, fax, userdiv, onlyflg, artistflg, kdesignflg
dim groupid, standardmdcatecode
dim catecode
dim partnerusing
dim defaultsongjangdiv
dim psocno
dim pcuserdiv, p_userdiv, c_userdiv, selltype, sellBizCd
dim padminUrl, padminId, padminPwd, pmallSellType, pcomType, taxevaltype, etcjungsantype, tplcompanyid
dim email, hp, businessgubun
dim idx, signtype

Dim vDefaultDeliveryType, vDefaultFreeBeasongLimit, vDefaultDeliverPay, vPurchaseType, vOffCateCode, vOffMDUserID
vDefaultDeliveryType		= Request("defaultdeliverytype")
vDefaultFreeBeasongLimit	= Request("defaultFreeBeasongLimit")
vDefaultDeliverPay			= Request("defaultDeliverPay")
vPurchaseType				= Request("purchasetype")
vOffCateCode				= Request("offcatecode")
vOffMDUserID				= Request("offmduserid")
standardmdcatecode	= requestCheckvar(request("standardmdcatecode"),3)
groupid		= request("groupid")
uid			 = request("uid")
onlyflg		= request("onlyflg")
artistflg		= request("artistflg")
kdesignflg		= request("kdesignflg")
catecode	= request("catecode")
pcuserdiv = requestCheckvar(request("pcuserdiv"),16)
selltype = requestCheckvar(request("selltype"),16)
sellBizCd = requestCheckvar(request("sellBizCd"),16)
padminUrl  = requestCheckvar(request("padminUrl"),160)
padminId  = requestCheckvar(request("padminId"),32)
padminPwd  = requestCheckvar(request("padminPwd"),32)
pmallSellType  = requestCheckvar(request("pmallSellType"),16)
pcomType = requestCheckvar(request("pcomType"),16)
taxevaltype = requestCheckvar(request("taxevaltype"),10)
etcjungsantype = requestCheckvar(request("etcjungsantype"),16)
    businessgubun      = requestCheckVar(request("businessgubun"),1)
email = requestCheckvar(request("email"),128)
hp = requestCheckvar(request("hp"),16)

signtype = requestCheckvar(request("signtype"),2)

if (pcuserdiv<>"") then
    p_userdiv = Trim(splitvalue(pcuserdiv,"_",0))
    c_userdiv = Trim(splitvalue(pcuserdiv,"_",1))
end if

if (selltype="") then
    selltype = "0"
end if

dim ceoname, company_no, zipcode, jungsan_gubun
dim jungsan_date,jungsan_bank,jungsan_acctno
dim jungsan_date_off, jungsan_date_frn
dim jungsan_acctname, manager_phone, manager_hp
dim deliver_name, deliver_phone, deliver_email
dim deliver_hp, jungsan_name, jungsan_phone, jungsan_email
dim jungsan_hp, prtidx, mduserid

deliver_name 	= requestCheckVar(request("deliver_name"),12)
deliver_phone 	= requestCheckVar(request("deliver_phone"),16)
deliver_email 	= requestCheckVar(request("deliver_email"),128)
deliver_hp 		= requestCheckVar(request("deliver_hp"),16)
prtidx 			= request("prtidx")
mduserid		= request("mduserid")
IF (prtidx="") then prtidx="9999"

tplcompanyid   = requestCheckvar(request("tplcompanyid"),32)
partnerusing= request("partnerusing")
defaultsongjangdiv = request("defaultsongjangdiv")
psocno = request("psocno")
company_name = html2db(request("company_name"))
manager_name = html2db(request("manager_name"))
address		 = html2db(request("address"))
manager_address = html2db(request("manager_address"))
tel			= html2db(request("tel"))
fax			= html2db(request("fax"))
userdiv 	= request("userdiv")
ceoname			= html2db(request("ceoname"))
company_no  	= Trim(request("company_no"))
zipcode			= request("zipcode")
jungsan_gubun 	= request("jungsan_gubun")
jungsan_date 	= request("jungsan_date")
jungsan_date_off= request("jungsan_date_off")
''사용안함.. off=frn 정산일
jungsan_date_frn= request("jungsan_date_frn")
jungsan_bank 	= html2db(request("jungsan_bank"))
jungsan_acctno 	= Trim(request("jungsan_acctno"))   '' trim 추가
jungsan_acctno  = replace(jungsan_acctno," ","")    ''  추가
jungsan_acctno  = replace(jungsan_acctno,"-","")    ''  추가
jungsan_acctname = html2db(request("jungsan_acctname"))
manager_phone 	= request("manager_phone")
manager_hp 		= request("manager_hp")
jungsan_name 	= html2db(request("jungsan_name"))
jungsan_phone 	= request("jungsan_phone")
jungsan_email 	= request("jungsan_email")
jungsan_hp 		= request("jungsan_hp")
dim company_zipcode, company_address, company_address2
dim company_tel, company_fax, return_zipcode, return_address, return_address2
dim manager_email

company_zipcode = request("company_zipcode")
company_address = request("company_address")
company_address2 = request("company_address2")
company_tel = request("company_tel")
company_fax = request("company_fax")
return_zipcode = requestCheckVar(request("return_zipcode"),8)
return_address = requestCheckVar(request("return_address"),128)
return_address2 = requestCheckVar(request("return_address2"),128)
manager_email = request("manager_email")

''if not IsNumeric(prtidx) then prtidx=9999

dim company_upjong,company_uptae
company_upjong  = Left(html2db(request("company_upjong")),32)
company_uptae   = Left(html2db(request("company_uptae")),25)
dim subid
subid   = request("subid")

dim mode
mode = request("mode")

dim commission
commission = request("commission")
if (commission="") then commission=0

dim socname_kor, socname, isusing, isextusing, streetusing, isoffusing
dim extstreetusing, specialbrand, maeipdiv, defaultmargine
dim M_margin, W_margin, U_margin

socname_kor  = html2db(trim(request("socname_kor")))
socname		 = html2db(trim(request("socname")))
isusing		 = request("isusing")
isextusing	 = request("isextusing")
streetusing	 = request("streetusing")
extstreetusing	 = request("extstreetusing")
specialbrand	 = request("specialbrand")
maeipdiv		 = request("maeipdiv")
defaultmargine	 = request("defaultmargine")
M_margin	 = request("M_margin")
W_margin	 = request("W_margin")
U_margin	 = request("U_margin")
isoffusing   = request("isoffusing")        ''2016/05/27

dim jungsan_bank_addCNT
dim j, jungsan_add_brand, jungsan_bank_add,jungsan_acctno_add,jungsan_acctname_add,jungsan_date_add,jungsan_date_off_add

	    
	    
if (isoffusing="") then isoffusing="N"
    
dim sqlStr, idExists
dim Enc_userpass, Enc_userpass64, Enc_2userpass64
dim prePrtIdx, password, passwordS

password="guest"

Enc_userpass = MD5(password)
Enc_userpass64 = SHA256(MD5(password))
Enc_2userpass64= SHA256(MD5(passwordS))

if mode = "edit" then
   rw "사용 중지 메뉴-관리자 문의 요망"
   dbget.Close() : response.end
end if

dim bufComno

''rw mode
On Error Resume Next
dbget.beginTrans

if (mode="brandedit") then
    '''/admin/member/popbrandinfoonly.asp


elseif mode="groupedit" then
    ''/admin/member/popupcheinfoonly.asp


elseif mode="modifyreturnaddress" then


elseif mode="modiprevmonthgroupid" Then


elseif mode="addnewupchebrand" then
    ''/admin/member/addnewbrand.asp

	'// 아이디 중복 확인
	sqlStr = "select count(*) from [db_user].[dbo].tbl_logindata with (nolock) where userid='" & uid & "'"
	rsget.Open sqlStr,dbget,1
	if rsget(0)>0 then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('이미 존재하거나 [일반고객]과 중복되는 아이디입니다. 다른 아이디를 입력해주세요.');" &vbCrLf &_
						" 	history.back();" &vbCrLf &_
						"</script>"
		response.End'dbget.close()	:	
	end if
	rsget.Close

	sqlStr = "select count(*) from [db_user].[dbo].tbl_deluser with (nolock) where userid='" & uid & "'"
	rsget.Open sqlStr,dbget,1
	if rsget(0)>0 then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('이미 존재하거나 [일반고객]과 중복되는 아이디입니다. 다른 아이디를 입력해주세요.');" &vbCrLf &_
						" 	history.back();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	rsget.Close

	sqlStr = "select count(*) from db_shop.dbo.tbl_shop_user with(nolock) where userid='" & uid & "'"
	rsget.Open sqlStr,dbget,1
	if rsget(0)>0 then
		response.write "<script type='text/javascript'>" &vbCrLf &_
						"	alert('이미 존재하는 아이디입니다. 다른 아이디를 입력해주세요.(오프라인)');" &vbCrLf &_
						" 	history.back();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	rsget.Close

	if (company_no = "888-00-00000") then
		'// 해외는 앞부분 888 이 고정이고 뒷부분 숫자는 자동증가

		idx = 0
		sqlStr = " select top 1 cast(right(replace(company_no , '-', ''), 7) as int) as idx "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_partner].[dbo].tbl_partner_group with (nolock)"
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

	''insert tbl_user_c
	sqlStr = "insert into [db_user].[dbo].tbl_user_c" & vbCrlf
	sqlStr = sqlStr + "(userid,socno,socname,birthday,socurl,ceoname," + vbCrlf
	sqlStr = sqlStr + "prcname," + vbCrlf
	sqlStr = sqlStr + "regdate,mileage,userdiv,catecode," + vbCrlf
	sqlStr = sqlStr + "isusing, isb2b, isextusing, vatinclude, maeipdiv," + vbCrlf
	sqlStr = sqlStr + "defaultmargine, socname_kor," & vbCrlf
	sqlStr = sqlStr + "coname,streetusing,extstreetusing,specialbrand,mduserid" + vbCrlf
	sqlStr = sqlStr + ",onlyflg,artistflg,kdesignflg,isoffusing" + vbCrlf

	'####### 업체조건배송인경우 얼마 미만 구매시 얼마 배송료 입력. 20110831
	If (maeipdiv = "U") or (pcuserdiv="999_50") Then
		sqlStr = sqlStr + ",defaultDeliveryType" + vbCrlf
		If vDefaultDeliveryType = "9" Then
			sqlStr = sqlStr + ",defaultFreeBeasongLimit,defaultDeliverPay" + vbCrlf
		End If
	End If

	sqlStr = sqlStr & " ,standardmdcatecode)Values(" & vbCrlf
	sqlStr = sqlStr + "'" + uid + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + socialnoBlank(company_no) + "'" + vbCrlf        '' 빈값처리. 2016/07/26 
	sqlStr = sqlStr + ",'" + socname + "'" + vbCrlf
	sqlStr = sqlStr + ",convert(varchar(10),getdate(),20)" + vbCrlf
'	sqlStr = sqlStr + ",'" + manager_email + "'" + vbCrlf					''업체가 등록으로 변경
	sqlStr = sqlStr + ",''" + vbCrlf
	sqlStr = sqlStr + ",'" + ceoname + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + manager_name + "'" + vbCrlf					''업체가 등록으로 변경
	sqlStr = sqlStr + ", getDate()"  + vbCrlf
	sqlStr = sqlStr + ",0" + vbCrlf
	sqlStr = sqlStr + ",'" + c_userdiv + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + catecode + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + isusing + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + "N" + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + isextusing + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + "Y" + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + maeipdiv + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + CStr(defaultmargine) + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + socname_kor + "'" + vbCrlf

	sqlStr = sqlStr + ",'" + company_name + "'" + vbCrlf					''업체가 등록으로 변경
	sqlStr = sqlStr + ",'" + streetusing + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + extstreetusing + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + specialbrand + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + mduserid + "'" + vbCrlf

	sqlStr = sqlStr + ",'" + onlyflg + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + artistflg + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + kdesignflg + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + isoffusing + "'" + vbCrlf
	
	If (maeipdiv = "U") or (pcuserdiv="999_50") Then
		If vDefaultDeliveryType = "null" Then
			sqlStr = sqlStr + ",null" + vbCrlf
		Else
			sqlStr = sqlStr + ",'" + vDefaultDeliveryType + "'" + vbCrlf
			If vDefaultDeliveryType = "9" Then
				sqlStr = sqlStr + ",'" + vDefaultFreeBeasongLimit + "','" + vDefaultDeliverPay + "'" + vbCrlf
			End If
		End IF
	End If
	sqlStr = sqlStr + ",'" & standardmdcatecode & "'" & vbCrlf
	sqlStr = sqlStr +  ")"

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	''insert tbl_partner
	sqlStr = "insert into [db_partner].[dbo].tbl_partner" + vbCrlf
	sqlStr = sqlStr + "(id, userdiv, company_no, signtype"+ vbCrlf
	if (jungsan_date<>"") then
	sqlStr = sqlStr + ", jungsan_date"+ vbCrlf
	end if
	if (jungsan_date_off<>"") then
	sqlStr = sqlStr + ", jungsan_date_off, jungsan_date_frn"+ vbCrlf
	end if
	sqlStr = sqlStr + ")" + vbCrlf
	sqlStr = sqlStr + " values('" + uid + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" & p_userdiv &"'" + vbCrlf
	sqlStr = sqlStr + " ,'" & company_no & "'" + vbCrlf
	sqlStr = sqlStr + " ,'" & signtype & "'" + vbCrlf
	if (jungsan_date<>"") then
	sqlStr = sqlStr + " ,'" + jungsan_date + "'" + vbCrlf
	end if
	if (jungsan_date_off<>"") then
	sqlStr = sqlStr + " ,'" & jungsan_date_off & "'"+ vbCrlf
	sqlStr = sqlStr + " ,'" + jungsan_date_off + "'" + vbCrlf
	end if
	sqlStr = sqlStr + " )"
 ''rw sqlStr
	 rsget.Open sqlStr,dbget,1

	''insert tbl_logindata
	sqlStr = "insert into [db_user].[dbo].tbl_logindata"
	sqlStr = sqlStr + "(userid,userpass,userdiv,lastlogin,Enc_userpass,Enc_userpass64,counter,userlevel)" + vbCrlf
	sqlStr = sqlStr + " Values("
	sqlStr = sqlStr + " '" + (uid) + "'" + vbCrlf
	sqlStr = sqlStr + " ,'' " + vbCrlf
	sqlStr = sqlStr + ",'" + (c_userdiv) + "'" + vbCrlf
	sqlStr = sqlStr + ",getdate()" + vbCrlf
	sqlStr = sqlStr + ",''" + vbCrlf
	sqlStr = sqlStr + ",'0'" + vbCrlf
	sqlStr = sqlStr + ",0,9" & ")"
	rsget.Open sqlStr,dbget,1


    ''매장(직영,가맹,도매)정보 저장
    if (pcuserdiv="501_21") or (pcuserdiv="502_21") or (pcuserdiv="503_21") then
		dim shopdiv
		Select Case pcuserdiv
			Case "501_21": shopdiv="1"	'직영
			Case "502_21": shopdiv="3"	'가맹
			Case "503_21": shopdiv="5"	'도매
		end Select
		
		sqlStr = "insert into db_shop.dbo.tbl_shop_user " + vbCrlf
		sqlStr = sqlStr + "(userid, userpass, shopname, shopdiv, vieworder, admindisplang, loginsite, viewsort) "+ vbCrlf
		sqlStr = sqlStr + " values "+ vbCrlf
		sqlStr = sqlStr + "('" + CStr(uid) + "'"+ vbCrlf
		sqlStr = sqlStr + ",''"+ vbCrlf
		sqlStr = sqlStr + ",'" + CStr(socname_kor) + "'"+ vbCrlf
		sqlStr = sqlStr + ",'" + CStr(shopdiv) + "'"+ vbCrlf
		sqlStr = sqlStr + ",'0'"+ vbCrlf
		sqlStr = sqlStr + ",'KOR'"+ vbCrlf
		sqlStr = sqlStr + ",''"+ vbCrlf
		sqlStr = sqlStr + ",'0')"+ vbCrlf
		dbget.Execute sqlStr
	end if

    ''제휴사 기타출고처
	''전송 데이터
	'sellBizCd, commission, taxevaltype, etcjungsantype, padminUrl, padminId, padminPwd, pmallSellType, pcomType, defaultmargine
elseif mode="addnewupchebrand2" then
   ''/admin/member/addnewbrand.asp

	'// 아이디 중복 확인
	sqlStr = "select count(*) from [db_user].[dbo].tbl_logindata with (nolock) where userid='" & uid & "'"
	rsget.Open sqlStr,dbget,1
	if rsget(0)>0 then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('이미 존재하거나 [일반고객]과 중복되는 아이디입니다. 다른 아이디를 입력해주세요.');" &vbCrLf &_
						" 	history.back();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	rsget.Close

	sqlStr = "select count(*) from [db_user].[dbo].tbl_deluser with (nolock) where userid='" & uid & "'"
	rsget.Open sqlStr,dbget,1
	if rsget(0)>0 then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('이미 존재하거나 [일반고객]과 중복되는 아이디입니다. 다른 아이디를 입력해주세요.');" &vbCrLf &_
						" 	history.back();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	rsget.Close

	if (company_no = "888-00-00000") then
		'// 해외는 앞부분 888 이 고정이고 뒷부분 숫자는 자동증가

		idx = 0
		sqlStr = " select top 1 cast(right(replace(company_no , '-', ''), 7) as int) as idx "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner_group with (nolock)"
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
		response.write "생성된 해외사업자 번호 : " & company_no & "<br>"
	end if

	''insert tbl_logindata
	sqlStr = "insert into [db_user].[dbo].tbl_logindata"
	sqlStr = sqlStr + "(userid,userpass,userdiv,lastlogin,Enc_userpass,Enc_userpass64,counter,userlevel)" + vbCrlf
	sqlStr = sqlStr + " Values("
	sqlStr = sqlStr + " '" + (uid) + "'" + vbCrlf
	sqlStr = sqlStr + " ,'' " + vbCrlf
	sqlStr = sqlStr + ",'" + (c_userdiv) + "'" + vbCrlf
	sqlStr = sqlStr + ",getdate()" + vbCrlf
	sqlStr = sqlStr + ",''" + vbCrlf
	sqlStr = sqlStr + ",'" + (Enc_userpass64) + "'" + vbCrlf
	sqlStr = sqlStr + ",0,9" & ")"
	rsget.Open sqlStr,dbget,1
''rw sqlStr

	''insert tbl_user_c
	sqlStr = "insert into [db_user].[dbo].tbl_user_c" & vbCrlf
	sqlStr = sqlStr + "(userid,socno,socname,birthday,socmail,socurl,ceoname," + vbCrlf
	sqlStr = sqlStr + "prcname," + vbCrlf
	sqlStr = sqlStr + "regdate,mileage,userdiv,catecode," + vbCrlf
	sqlStr = sqlStr + "isusing, isb2b, isextusing, vatinclude, maeipdiv," + vbCrlf
	sqlStr = sqlStr + "defaultmargine, socname_kor," & vbCrlf
	sqlStr = sqlStr + "coname,streetusing,extstreetusing,specialbrand,mduserid" + vbCrlf
	sqlStr = sqlStr + ",onlyflg,artistflg,kdesignflg,isoffusing" + vbCrlf

	'####### 업체조건배송인경우 얼마 미만 구매시 얼마 배송료 입력. 20110831
	If (maeipdiv = "U") or (pcuserdiv="999_50") Then
		sqlStr = sqlStr + ",defaultDeliveryType" + vbCrlf
		If vDefaultDeliveryType = "9" Then
			sqlStr = sqlStr + ",defaultFreeBeasongLimit,defaultDeliverPay" + vbCrlf
		End If
	End If

	sqlStr = sqlStr + " ,standardmdcatecode)Values("
	sqlStr = sqlStr + "'" + uid + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + socialnoBlank(company_no) + "'" + vbCrlf        '' 빈값처리. 2016/07/26 
	sqlStr = sqlStr + ",'" + socname + "'" + vbCrlf
	sqlStr = sqlStr + ",convert(varchar(10),getdate(),20)" + vbCrlf
	sqlStr = sqlStr + ",'" + manager_email + "'" + vbCrlf
	sqlStr = sqlStr + ",''" + vbCrlf
	sqlStr = sqlStr + ",'" + ceoname + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + manager_name + "'" + vbCrlf
	sqlStr = sqlStr + ", getDate()"  + vbCrlf
	sqlStr = sqlStr + ",0" + vbCrlf
	sqlStr = sqlStr + ",'" + c_userdiv + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + catecode + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + isusing + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + "N" + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + isextusing + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + "Y" + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + maeipdiv + "'" + vbCrlf
	sqlStr = sqlStr + "," + CStr(defaultmargine) + vbCrlf
	sqlStr = sqlStr + ",'" + socname_kor + "'" + vbCrlf

	sqlStr = sqlStr + ",'" + company_name + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + streetusing + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + extstreetusing + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + specialbrand + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + mduserid + "'" + vbCrlf

	sqlStr = sqlStr + ",'" + onlyflg + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + artistflg + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + kdesignflg + "'" + vbCrlf
	sqlStr = sqlStr + ",'" + isoffusing + "'" + vbCrlf
	
	If (maeipdiv = "U") or (pcuserdiv="999_50") Then
		If vDefaultDeliveryType = "null" Then
			sqlStr = sqlStr + ",null" + vbCrlf
		Else
			sqlStr = sqlStr + ",'" + vDefaultDeliveryType + "'" + vbCrlf
			If vDefaultDeliveryType = "9" Then
				sqlStr = sqlStr + ",'" + vDefaultFreeBeasongLimit + "','" + vDefaultDeliverPay + "'" + vbCrlf
			End If
		End IF
	End If

	sqlStr = sqlStr + ",'" & standardmdcatecode & "'" & vbCrlf
	sqlStr = sqlStr +  ")"
''rw sqlStr
	dbget.execute sqlStr

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
		sqlStr = "select top 1 groupid from [db_partner].[dbo].tbl_partner_group with (nolock)"
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
''rw sqlStr
		dbget.Execute sqlStr

		if (LEN(Trim(replace(company_no,"-","")))=13) then
			sqlStr = "exec [db_cs].[dbo].[usp_Ten_partner_Enc_companyno] '"&groupid&"','"&company_no&"'"
			dbget.Execute sqlStr
		end if
	end if

	''insert tbl_partner
	sqlStr = "insert into [db_partner].[dbo].tbl_partner" + vbCrlf
	sqlStr = sqlStr + "(id,Enc_password,Enc_password64,Enc_2password64,userdiv,jungsan_date,groupid"+ vbCrlf
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
	sqlStr = sqlStr + " ,'"&p_userdiv&"'" + vbCrlf
	sqlStr = sqlStr + " ,'" + jungsan_date + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" + groupid + "'" + vbCrlf

	sqlStr = sqlStr + " ,'" + deliver_name + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" + deliver_phone + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" + deliver_hp + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" + deliver_email + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" + return_zipcode + "'" + vbCrlf         ''초기 반품주소는 사무실 주소와 동일하게 설정됩니다.
	sqlStr = sqlStr + " ,'" + return_address + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" + return_address2 + "'" + vbCrlf
	sqlStr = sqlStr + " ,"&Vpurchasetype&""+ vbCrlf
	sqlStr = sqlStr + " ,'" + Voffcatecode + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" + Voffmduserid + "'" + vbCrlf
	sqlStr = sqlStr + " ,"&selltype&""+ vbCrlf
	if (defaultsongjangdiv<>"") then
	    sqlStr = sqlStr + " ,'" & defaultsongjangdiv + "'" + VbCrlf
	end if
	sqlStr = sqlStr + " )"
 ''rw sqlStr
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
''rw sqlStr
	rsget.Open sqlStr,dbget,1

    ''제휴사 기타출고처
    if (pcuserdiv="999_50") or (pcuserdiv="900_21") then
        sqlStr = " update [db_partner].[dbo].tbl_partner" + VbCrlf
		sqlStr = sqlStr + " set lastInfoChgDT=getdate(), sellBizCd='" + sellBizCd+ "'" + VbCrlf
		if (commission<>"") then
		    sqlStr = sqlStr + " ,commission="&commission/100&VbCrlf
	    end if
		if (taxevaltype<>"") then
		    sqlStr = sqlStr + " ,taxevaltype=" & taxevaltype & "" + VbCrlf
		end if
		sqlStr = sqlStr + " ,etcjungsantype='" & etcjungsantype & "'" + VbCrlf
		if (tplcompanyid<>"") then
		    sqlStr = sqlStr + " ,tplcompanyid='"&tplcompanyid&"'" + VbCrlf
		else
		    sqlStr = sqlStr + " ,tplcompanyid=NULL" + VbCrlf
	    end if
		sqlStr = sqlStr + " where id='"&uid&"'"
		dbget.Execute sqlStr
    end if

    ''제휴몰(온라인)인경우
    if (pcuserdiv="999_50") then

        sqlStr = " IF Exists(select * from db_partner.dbo.tbl_partner_addInfo where partnerid='"&uid&"')"+ VbCrlf
        sqlStr = sqlStr + " BEGIN"+ VbCrlf
        sqlStr = sqlStr + " update db_partner.dbo.tbl_partner_addInfo"+ VbCrlf
        sqlStr = sqlStr + " set padminUrl='"&padminUrl&"'"+ VbCrlf
		sqlStr = sqlStr + " ,padminId='"&padminId&"'"+ VbCrlf
		sqlStr = sqlStr + " ,padminPwd='"&padminPwd&"'"+ VbCrlf
		sqlStr = sqlStr + " ,pmallSellType='"&pmallSellType&"'"+ VbCrlf
		sqlStr = sqlStr + " ,pcomType='"&pcomType&"'"+ VbCrlf
		sqlStr = sqlStr + " where partnerid='"&uid&"'"+ VbCrlf
		sqlStr = sqlStr + " END "+ VbCrlf
		sqlStr = sqlStr + " ELSE "+ VbCrlf
		sqlStr = sqlStr + " BEGIN"+ VbCrlf
		sqlStr = sqlStr + " Insert Into db_partner.dbo.tbl_partner_addInfo"+ VbCrlf
		sqlStr = sqlStr + " (partnerid,padminUrl,padminId,padminPwd,pmallSellType,pcomType)"+ VbCrlf
		sqlStr = sqlStr + " values('"&uid&"'"+ VbCrlf
		sqlStr = sqlStr + " ,'"&padminUrl&"'"+ VbCrlf
		sqlStr = sqlStr + " ,'"&padminId&"'"+ VbCrlf
		sqlStr = sqlStr + " ,'"&padminPwd&"'"+ VbCrlf
		sqlStr = sqlStr + " ,'"&pmallSellType&"'"+ VbCrlf
		sqlStr = sqlStr + " ,'"&pcomType&"'"+ VbCrlf
		sqlStr = sqlStr + " )"+ VbCrlf
		sqlStr = sqlStr + " END"+ VbCrlf
		'rw  sqlStr
		dbget.Execute sqlStr
    end if

    ''2013/12/08 추가마진 관련 추가 eastone
    if (maeipdiv<>"") and (defaultmargine<>"") then
        sqlStr = " update [db_partner].[dbo].tbl_partner"
        sqlStr = sqlStr & " SET lastInfoChgDT=getdate(), "&maeipdiv&"_margin="&defaultmargine
        sqlStr = sqlStr + " where id='"&uid&"'"
		dbget.Execute sqlStr
    end if
end if

'response.write Err.Number
'response.end
	If Err.Number = 0 Then
	        dbget.CommitTrans
			dim title, lmscontents, qstring, kakaocontents, btnJson
			title = "[10x10]입점 관련 안내입니다."
			qstring = Server.UrlEncode(TBTEncryptUrl(Cstr(uid) + "|" +Cstr(pcuserdiv)))
			lmscontents = "안녕하세요. 텐바이텐입니다." & vbcrlf
			lmscontents = lmscontents & "입점을 환영합니다."  & vbcrlf& vbcrlf
			lmscontents = lmscontents & "아래 링크로 이동 하신 후 업체정보를 입력해 주시면,"  & vbcrlf
			lmscontents = lmscontents & "어드민(SCM) 로그인이 가능합니다."  & vbcrlf & vbcrlf
			lmscontents = lmscontents & "[사업자등록페이지 이동]"  & vbcrlf
			lmscontents = lmscontents & "https://scm.10x10.co.kr/common/partner/companyinfo.asp?qs="+qstring & vbcrlf& vbcrlf
			lmscontents = lmscontents & "감사합니다."

			kakaocontents = "텐바이텐 입점을 환영합니다." & vbcrlf & vbcrlf
			kakaocontents = kakaocontents & "아래 링크로 이동 하신 후 업체 정보를 입력해 주시면, "
			kakaocontents = kakaocontents & "어드민(SCM) 로그인이 가능합니다." & vbcrlf & vbcrlf
			kakaocontents = kakaocontents & "감사합니다." & vbcrlf
			btnJson = "{""button"":[{""name"":""업체정보 입력 바로가기"",""type"":""WL"", ""url_mobile"":""https://scm.10x10.co.kr/common/partner/companyinfo.asp?qs=" & trim(qstring) &"""}]}"
			'LMS발송
			'call SendNormalLMS(hp,title,"",lmscontents)
			'KAKAO발송
			call SendKakaoMsg_LINK(hp,"","A-0006",kakaocontents,"LMS",title,lmscontents,btnJson)
			'Email발송
			call sendmailPartnerJoin(email,"https://scm.10x10.co.kr/common/partner/companyinfo.asp?qs="+qstring)
	Else
	        dbget.RollBackTrans
	        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\n입력한 값들이 너무 길지 않는지 확인바랍니다.\n\n" & Err.description & "')</script>"
	        ''response.write "<script>history.back()</script>"
	        dbget.close()
	        response.end
	End If

	on error Goto 0


''브랜드정보동기화 140->138

dim IsLecturer

sqlStr = "select top 1 * from  [db_user].[dbo].tbl_user_c" + VbCrlf
sqlStr = sqlStr + " where userid='" + uid + "'" + VbCrlf
sqlStr = sqlStr + " and userdiv='14'" + VbCrlf

rsget.Open sqlStr, dbget, 1
if Not rsget.Eof then
	IsLecturer = true
else
    IsLecturer = false
end if
rsget.Close

if (IsLecturer) and  (pcuserdiv="9999_14") then
    Dim lec_yn                      : lec_yn = requestCheckVar(request("lec_yn"),10)
    Dim lec_margin                  : lec_margin = requestCheckVar(request("lec_margin"),10)
    Dim mat_margin                  : mat_margin = requestCheckVar(request("mat_margin"),10)
    Dim diy_yn                      : diy_yn = requestCheckVar(request("diy_yn"),10)
    Dim diy_margin                  : diy_margin = requestCheckVar(request("diy_margin"),10)
    Dim diy_dlv_gubun               : diy_dlv_gubun = requestCheckVar(request("diy_dlv_gubun"),10)

    if (lec_yn="N") then
        lec_margin=0
        mat_margin=0
    end if

    if (diy_yn="N") then
        diy_margin="0"
        diy_dlv_gubun=0
        vDefaultFreebeasongLimit=0
        vDefaultDeliverPay=0
    end if

    ''2016/08/25 추가
    if (mode="brandedit") then
        sqlStr = sqlStr + " update db_user.dbo.tbl_user_c"&vbCRLF
        sqlStr = sqlStr + " set defaultFreeBeasongLimit="&vDefaultFreebeasongLimit&vbCRLF
        sqlStr = sqlStr + " ,defaultDeliverPay="&vDefaultDeliverPay&vbCRLF
        sqlStr = sqlStr + " ,defaultDeliveryType="&diy_dlv_gubun  &vbCRLF        
        sqlStr = sqlStr + " where userid='"&uid&"'"&vbCRLF
        dbget.Execute sqlStr
    end if
        
    ''강좌/DIY 마진 관련.
    sqlStr = " If Exists (select * from db_academy.dbo.tbl_lec_user where lecturer_id='"&uid&"')"
    sqlStr = sqlStr & " BEGIN"
    sqlStr = sqlStr & "     update db_academy.dbo.tbl_lec_user" & VbCRLF
    sqlStr = sqlStr & "     set lecturer_name=convert(varchar(32),'"&HTML2DB(socname_kor)&"')" & VbCRLF
    sqlStr = sqlStr & "     ,en_name=convert(varchar(32),'"&HTML2DB(socname)&"')" & VbCRLF
    sqlStr = sqlStr & "     ,lec_yn='"&lec_yn&"'" & VbCRLF
    sqlStr = sqlStr & "     ,diy_yn='"&diy_yn&"'" & VbCRLF
    sqlStr = sqlStr & "     ,lec_margin="&lec_margin & VbCRLF
    sqlStr = sqlStr & "     ,mat_margin="&mat_margin & VbCRLF
    sqlStr = sqlStr & "     ,diy_margin="&diy_margin & VbCRLF
    sqlStr = sqlStr & "     ,diy_dlv_gubun="&diy_dlv_gubun & VbCRLF
    sqlStr = sqlStr & "     ,DefaultFreebeasongLimit="&vDefaultFreebeasongLimit & VbCRLF
    sqlStr = sqlStr & "     ,DefaultDeliveryPay="&vDefaultDeliverPay & VbCRLF
    sqlStr = sqlStr & "     where lecturer_id='"&uid&"'"
    sqlStr = sqlStr & " END"
    sqlStr = sqlStr & " ELSE"
    sqlStr = sqlStr & " BEGIN"
    sqlStr = sqlStr & "     insert into db_academy.dbo.tbl_lec_user" & VbCRLF
    sqlStr = sqlStr & "     (lecturer_id,lecturer_name,en_name, lec_yn,diy_yn,lec_margin,mat_margin,diy_margin,diy_dlv_gubun,DefaultFreebeasongLimit,DefaultDeliveryPay)"& VbCRLF
    sqlStr = sqlStr & "     values('"&uid&"'"
    sqlStr = sqlStr & "     ,convert(varchar(32),'"&HTML2DB(socname_kor)&"')" & VbCRLF
    sqlStr = sqlStr & "     ,convert(varchar(32),'"&HTML2DB(socname)&"')" & VbCRLF
    sqlStr = sqlStr & "     ,'"&lec_yn&"'" & VbCRLF
    sqlStr = sqlStr & "     ,'"&diy_yn&"'" & VbCRLF
    sqlStr = sqlStr & "     ,"&lec_margin&"" & VbCRLF
    sqlStr = sqlStr & "     ,"&mat_margin&"" & VbCRLF
    sqlStr = sqlStr & "     ,"&diy_margin&"" & VbCRLF
    sqlStr = sqlStr & "     ,"&diy_dlv_gubun&"" & VbCRLF
    sqlStr = sqlStr & "     ,"&vDefaultFreebeasongLimit&"" & VbCRLF
    sqlStr = sqlStr & "     ,"&vDefaultDeliverPay&"" & VbCRLF
    sqlStr = sqlStr & "     )" & VbCRLF
    sqlStr = sqlStr & " END"

    dbAcademyget.Execute sqlStr

    call copyFingersUserC(uid) '' 따로 뺌. 2016/05/16
end if

if (pcuserdiv="9999_15") then ''추가 2016/05/16
    call copyFingersUserC(uid)
end if

'' 따로 뺌. 2016/05/16
function copyFingersUserC(uid)
    dim sqlStr, idExists
    sqlStr = "select top 1 * from  [db3_common].[dbo].tbl_user_c"
    sqlStr = sqlStr + " where userid='" + uid + "'" + VbCrlf

    rsAcademyget.Open sqlStr, dbAcademyget, 1
    if Not rsAcademyget.Eof then
    	idExists = true
    else
        idExists = false
    end if
    rsAcademyget.Close

    if (idExists) then
        sqlStr = "update [db3_common].[dbo].tbl_user_c"
        sqlStr = sqlStr + " set socno=T.socno"
        sqlStr = sqlStr + " ,socname=T.socname"
        sqlStr = sqlStr + " ,catecode=T.catecode"
        sqlStr = sqlStr + " ,birthday=T.birthday"
        sqlStr = sqlStr + " ,socmail=T.socmail"
        sqlStr = sqlStr + " ,socurl=T.socurl"
        sqlStr = sqlStr + " ,ceoname=T.ceoname"
        sqlStr = sqlStr + " ,prcname=T.prcname"
        sqlStr = sqlStr + " ,zipcode=T.zipcode"
        sqlStr = sqlStr + " ,socaddr=T.socaddr"
        sqlStr = sqlStr + " ,socphone=T.socphone"
        sqlStr = sqlStr + " ,soccell=T.soccell"
        sqlStr = sqlStr + " ,socfax=T.socfax"
        sqlStr = sqlStr + " ,soctype=T.soctype"
        sqlStr = sqlStr + " ,socitem=T.socitem"
        sqlStr = sqlStr + " ,regdate=T.regdate"
        sqlStr = sqlStr + " ,mileage=T.mileage"
        sqlStr = sqlStr + " ,socicon=T.socicon"
        sqlStr = sqlStr + " ,soclogo=T.soclogo"
        sqlStr = sqlStr + " ,soccomment=T.soccomment"
        sqlStr = sqlStr + " ,isusing=T.isusing"
        sqlStr = sqlStr + " ,isb2b=T.isb2b"
        sqlStr = sqlStr + " ,userdiv=T.userdiv"
        sqlStr = sqlStr + " ,isoffusing=T.isoffusing"
        ''sqlStr = sqlStr + " ,isextusing=T.isextusing"				'// 수정불가 : 제휴몰별 판매설정 팝업창에서만 수정가능
        sqlStr = sqlStr + " ,vatinclude=T.vatinclude"
        sqlStr = sqlStr + " ,maeipdiv=T.maeipdiv"
        sqlStr = sqlStr + " ,defaultmargine=T.defaultmargine"
        sqlStr = sqlStr + " ,socname_kor=T.socname_kor"
        sqlStr = sqlStr + " ,coname=T.coname"
        sqlStr = sqlStr + " ,bankname=T.bankname"
        sqlStr = sqlStr + " ,acountname=T.acountname"
        sqlStr = sqlStr + " ,acountno=T.acountno"
        sqlStr = sqlStr + " ,prtidx=T.prtidx"
        sqlStr = sqlStr + " ,streetusing=T.streetusing"
        sqlStr = sqlStr + " ,extstreetusing=T.extstreetusing"
        sqlStr = sqlStr + " ,specialbrand=T.specialbrand"
        sqlStr = sqlStr + " ,visitcount=T.visitcount"
        sqlStr = sqlStr + " ,todayvisitcount=T.todayvisitcount"
        sqlStr = sqlStr + " ,recommendcount=T.recommendcount"
        sqlStr = sqlStr + " ,todayrecommendcount=T.todayrecommendcount"
        sqlStr = sqlStr + " ,itemcount=T.itemcount"
        sqlStr = sqlStr + " ,dgncomment=T.dgncomment"
        sqlStr = sqlStr + " ,titleimgurl=T.titleimgurl"
        sqlStr = sqlStr + " ,mduserid=T.mduserid"
        sqlStr = sqlStr + " ,topbrandcount=T.topbrandcount"
        sqlStr = sqlStr + " ,recenttopbrandyn=T.recenttopbrandyn"
        sqlStr = sqlStr + " ,modelitem=T.modelitem"
        sqlStr = sqlStr + " ,modelitem2=T.modelitem2"
        sqlStr = sqlStr + " ,modelimg=T.modelimg"
        sqlStr = sqlStr + " ,modelbimg=T.modelbimg"
        sqlStr = sqlStr + " ,modelbimg2=T.modelbimg2"
        sqlStr = sqlStr + " ,hitrank=T.hitrank"
        sqlStr = sqlStr + " ,smilerank=T.smilerank"
        sqlStr = sqlStr + " ,salerank=T.salerank"
        sqlStr = sqlStr + " ,giftrank=T.giftrank"
        sqlStr = sqlStr + " ,giftflg=T.giftflg"
        sqlStr = sqlStr + " ,hitflg=T.hitflg"
        sqlStr = sqlStr + " ,saleflg=T.saleflg"
        sqlStr = sqlStr + " ,smileflg=T.smileflg"
        sqlStr = sqlStr + " ,newflg=T.newflg"
        sqlStr = sqlStr + " ,defaultFreeBeasongLimit=T.defaultFreeBeasongLimit"  ''2016/06/16 추가
        sqlStr = sqlStr + " ,defaultDeliverPay=T.defaultDeliverPay"              ''2016/06/16 추가
        sqlStr = sqlStr + " ,defaultDeliveryType=T.defaultDeliveryType"          ''2016/06/16 추가
        sqlStr = sqlStr + " from (select * from [TENDB].[db_user].[dbo].tbl_user_c where userid='" + uid + "') as T"
        sqlStr = sqlStr + " where [db3_common].[dbo].tbl_user_c.userid='" + uid + "'"
        sqlStr = sqlStr + " and [db3_common].[dbo].tbl_user_c.userid=T.userid"

        dbAcademyget.Execute sqlStr
    else
        sqlStr = "insert into [db3_common].[dbo].tbl_user_c"
        sqlStr = sqlStr + " ("
        sqlStr = sqlStr + " userid "
        sqlStr = sqlStr + " , socno "
        sqlStr = sqlStr + " , socname "
        sqlStr = sqlStr + " , catecode "
        sqlStr = sqlStr + " , birthday "
        sqlStr = sqlStr + " , socmail "
        sqlStr = sqlStr + " , socurl "
        sqlStr = sqlStr + " , ceoname "
        sqlStr = sqlStr + " , prcname "
        sqlStr = sqlStr + " , zipcode "
        sqlStr = sqlStr + " , socaddr "
        sqlStr = sqlStr + " , socphone "
        sqlStr = sqlStr + " , soccell "
        sqlStr = sqlStr + " , socfax "
        sqlStr = sqlStr + " , soctype "
        sqlStr = sqlStr + " , socitem "
        sqlStr = sqlStr + " , regdate "
        sqlStr = sqlStr + " , mileage "
        sqlStr = sqlStr + " , socicon "
        sqlStr = sqlStr + " , soclogo "
        sqlStr = sqlStr + " , soccomment "
        sqlStr = sqlStr + " , isusing "
        sqlStr = sqlStr + " , isb2b "
        sqlStr = sqlStr + " , userdiv "
        sqlStr = sqlStr + " , isextusing "
        sqlStr = sqlStr + " , vatinclude "
        sqlStr = sqlStr + " , maeipdiv "
        sqlStr = sqlStr + " , defaultmargine "
        sqlStr = sqlStr + " , socname_kor "
        sqlStr = sqlStr + " , coname "
        sqlStr = sqlStr + " , bankname "
        sqlStr = sqlStr + " , acountname "
        sqlStr = sqlStr + " , acountno "
        sqlStr = sqlStr + " , prtidx "
        sqlStr = sqlStr + " , streetusing "
        sqlStr = sqlStr + " , extstreetusing "
        sqlStr = sqlStr + " , specialbrand "
        sqlStr = sqlStr + " , visitcount "
        sqlStr = sqlStr + " , todayvisitcount "
        sqlStr = sqlStr + " , recommendcount "
        sqlStr = sqlStr + " , todayrecommendcount "
        sqlStr = sqlStr + " , itemcount "
        sqlStr = sqlStr + " , dgncomment "
        sqlStr = sqlStr + " , titleimgurl "
        sqlStr = sqlStr + " , mduserid "
        sqlStr = sqlStr + " , topbrandcount "
        sqlStr = sqlStr + " , recenttopbrandyn "
        sqlStr = sqlStr + " , modelitem "
        sqlStr = sqlStr + " , modelitem2 "
        sqlStr = sqlStr + " , modelimg "
        sqlStr = sqlStr + " , modelbimg "
        sqlStr = sqlStr + " , modelbimg2 "
        sqlStr = sqlStr + " , hitrank "
        sqlStr = sqlStr + " , smilerank "
        sqlStr = sqlStr + " , salerank "
        sqlStr = sqlStr + " , giftrank "
        sqlStr = sqlStr + " , giftflg "
        sqlStr = sqlStr + " , hitflg "
        sqlStr = sqlStr + " , saleflg "
        sqlStr = sqlStr + " , smileflg "
        sqlStr = sqlStr + " , newflg "
        sqlStr = sqlStr + " , isoffusing "
        sqlStr = sqlStr + " , defaultFreeBeasongLimit "     ''2016/06/16 추가
        sqlStr = sqlStr + " , defaultDeliverPay "           ''2016/06/16 추가
        sqlStr = sqlStr + " , defaultDeliveryType "         ''2016/06/16 추가
        sqlStr = sqlStr + " )"
        sqlStr = sqlStr + " select "
        sqlStr = sqlStr + " userid "
        sqlStr = sqlStr + " , socno "
        sqlStr = sqlStr + " , socname "
        sqlStr = sqlStr + " , catecode "
        sqlStr = sqlStr + " , birthday "
        sqlStr = sqlStr + " , socmail "
        sqlStr = sqlStr + " , socurl "
        sqlStr = sqlStr + " , ceoname "
        sqlStr = sqlStr + " , prcname "
        sqlStr = sqlStr + " , zipcode "
        sqlStr = sqlStr + " , socaddr "
        sqlStr = sqlStr + " , socphone "
        sqlStr = sqlStr + " , soccell "
        sqlStr = sqlStr + " , socfax "
        sqlStr = sqlStr + " , soctype "
        sqlStr = sqlStr + " , socitem "
        sqlStr = sqlStr + " , regdate "
        sqlStr = sqlStr + " , mileage "
        sqlStr = sqlStr + " , socicon "
        sqlStr = sqlStr + " , soclogo "
        sqlStr = sqlStr + " , soccomment "
        sqlStr = sqlStr + " , isusing "
        sqlStr = sqlStr + " , isb2b "
        sqlStr = sqlStr + " , userdiv "
        sqlStr = sqlStr + " , isextusing "
        sqlStr = sqlStr + " , vatinclude "
        sqlStr = sqlStr + " , maeipdiv "
        sqlStr = sqlStr + " , defaultmargine "
        sqlStr = sqlStr + " , socname_kor "
        sqlStr = sqlStr + " , coname "
        sqlStr = sqlStr + " , bankname "
        sqlStr = sqlStr + " , acountname "
        sqlStr = sqlStr + " , acountno "
        sqlStr = sqlStr + " , prtidx "
        sqlStr = sqlStr + " , streetusing "
        sqlStr = sqlStr + " , extstreetusing "
        sqlStr = sqlStr + " , specialbrand "
        sqlStr = sqlStr + " , visitcount "
        sqlStr = sqlStr + " , todayvisitcount "
        sqlStr = sqlStr + " , recommendcount "
        sqlStr = sqlStr + " , todayrecommendcount "
        sqlStr = sqlStr + " , itemcount "
        sqlStr = sqlStr + " , dgncomment "
        sqlStr = sqlStr + " , titleimgurl "
        sqlStr = sqlStr + " , mduserid "
        sqlStr = sqlStr + " , topbrandcount "
        sqlStr = sqlStr + " , recenttopbrandyn "
        sqlStr = sqlStr + " , modelitem "
        sqlStr = sqlStr + " , modelitem2 "
        sqlStr = sqlStr + " , modelimg "
        sqlStr = sqlStr + " , modelbimg "
        sqlStr = sqlStr + " , modelbimg2 "
        sqlStr = sqlStr + " , hitrank "
        sqlStr = sqlStr + " , smilerank "
        sqlStr = sqlStr + " , salerank "
        sqlStr = sqlStr + " , giftrank "
        sqlStr = sqlStr + " , giftflg "
        sqlStr = sqlStr + " , hitflg "
        sqlStr = sqlStr + " , saleflg "
        sqlStr = sqlStr + " , smileflg "
        sqlStr = sqlStr + " , newflg "
        sqlStr = sqlStr + " , isoffusing "
        sqlStr = sqlStr + " , defaultFreeBeasongLimit "
        sqlStr = sqlStr + " , defaultDeliverPay "
        sqlStr = sqlStr + " , defaultDeliveryType "
        sqlStr = sqlStr + " from [TENDB].[db_user].[dbo].tbl_user_c where userid='" + uid + "'"

        dbAcademyget.Execute sqlStr
    end if
end function

function sendmailPartnerJoin(mailto, contents)
        dim mailfrom, mailtitle, mailcontent,dirPath,fileName
        dim fs,objFile

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "[10x10] 입점 관련 안내 메일입니다."

        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email/mailtemplate")
        fileName = dirPath&"\\mail_partner_join.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
		mailcontent = replace(mailcontent,":CONTENTSHTML:",contents)
        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end Function
%>

<script>alert('저장되었습니다.');</script>
<% if mode="addnewupchebrand" then %>
	<script>top.location.href='/admin/member/popbrandinfoonly_small.asp?designer=<%= uid %>';</script>
<% elseif mode="addnewupchebrand2" then %>
    <script>top.location.href='/admin/member/popbrandinfoonly.asp?designer=<%= uid %>';</script>
<% else %>
    <script>location.replace('<%= refer %>');</script>
<% end if %>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->
