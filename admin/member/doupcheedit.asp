<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드 정보
' History : 서동석 생성
'           2021.06.18 한용민 수정(담당자 휴대폰,이메일 인증정보 데이터쪽에도 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/RackCodeFunction.asp"-->
<%

''response.write "<script>alert('소스보기로 실행결과를 볼 수 없다. 경고창을 띄우자.');</script>"

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim uid,company_name,email,manager_name,address
dim manager_address, tel, fax, userdiv, onlyflg, artistflg, kdesignflg
dim groupid, idx, adminid, sql
dim catecode, standardmdcatecode
dim partnerusing
dim defaultsongjangdiv, exists_login_title, exists_login_gubun, exists_login_hp, exists_login_email
dim psocno, addetc_name,addetc_hp,addetc_email
dim pcuserdiv, p_userdiv, c_userdiv, selltype, sellBizCd
dim padminUrl, padminId, padminPwd, pmallSellType, pcomType, taxevaltype, etcjungsantype, tplcompanyid
Dim vDefaultDeliveryType, vDefaultFreeBeasongLimit, vDefaultDeliverPay, vPurchaseType, vOffCateCode, vOffMDUserID
dim etc_idx,etc_name,etc_hp,etc_email, tmpetc_idx,tmpetc_name,tmpetc_hp,tmpetc_email, i, etcaddyn
	etc_idx 	= replace(trim(Request("etc_idx"))," ","")
	etc_name 	= replace(trim(Request("etc_name"))," ","")
	etc_hp 	= replace(trim(Request("etc_hp"))," ","")
	etc_email 	= replace(trim(Request("etc_email"))," ","")
	addetc_name 	= requestCheckVar(trim(Request("addetc_name")),32)
	addetc_hp 	= requestCheckVar(trim(Request("addetc_hp")),32)
	addetc_email 	= requestCheckVar(trim(Request("addetc_email")),128)
	etcaddyn 	= requestCheckVar(trim(Request("etcaddyn")),1)
adminid = session("ssBctId")
vDefaultDeliveryType		= Request("defaultdeliverytype")
vDefaultFreeBeasongLimit	= Request("defaultFreeBeasongLimit")
vDefaultDeliverPay			= Request("defaultDeliverPay")
vPurchaseType				= Request("purchasetype")
vOffCateCode				= Request("offcatecode")
vOffMDUserID				= Request("offmduserid")
groupid		= request("groupid")
uid			 = request("uid")
company_name = stripHTML(html2db(request("company_name")))
email		 = stripHTML(html2db(request("email")))
manager_name = stripHTML(html2db(request("manager_name")))
address		 = stripHTML(html2db(request("address")))
manager_address = stripHTML(html2db(request("manager_address")))
tel			= html2db(request("tel"))
fax			= html2db(request("fax"))
userdiv 	= request("userdiv")
onlyflg		= request("onlyflg")
artistflg		= request("artistflg")
kdesignflg		= request("kdesignflg")
catecode	= request("catecode")
standardmdcatecode	= requestCheckvar(request("standardmdcatecode"),3)
partnerusing= request("partnerusing")
defaultsongjangdiv = request("defaultsongjangdiv")
psocno = request("psocno")
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
tplcompanyid   = requestCheckvar(request("tplcompanyid"),32)

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

ceoname			= stripHTML(html2db(request("ceoname")))
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

jungsan_acctname = stripHTML(html2db(request("jungsan_acctname")))
manager_phone 	= request("manager_phone")
manager_hp 		= request("manager_hp")
deliver_name 	= stripHTML(requestCheckVar(request("deliver_name"),12))
deliver_phone 	= requestCheckVar(request("deliver_phone"),16)
deliver_email 	= stripHTML(requestCheckVar(request("deliver_email"),128))
deliver_hp 		= requestCheckVar(request("deliver_hp"),16)
jungsan_name 	= stripHTML(html2db(trim(request("jungsan_name"))))
jungsan_phone 	= trim(request("jungsan_phone"))
jungsan_email 	= stripHTML(requestCheckVar(trim(request("jungsan_email")),128))
jungsan_hp 		= trim(request("jungsan_hp"))
prtidx 			= request("prtidx")
mduserid		= request("mduserid")

dim company_zipcode, company_address, company_address2
dim company_tel, company_fax, return_zipcode, return_address, return_address2
dim manager_email

company_zipcode = request("company_zipcode")
company_address = stripHTML(requestCheckVar(request("company_address"),128))
company_address2 = stripHTML(requestCheckVar(request("company_address2"),128))
company_tel = request("company_tel")
company_fax = request("company_fax")
return_zipcode = requestCheckVar(request("return_zipcode"),8)
return_address = stripHTML(requestCheckVar(request("return_address"),128))
return_address2 = stripHTML(requestCheckVar(request("return_address2"),128))
manager_email = requestCheckVar(request("manager_email"),128)


''if not IsNumeric(prtidx) then prtidx=9999
IF (prtidx="") then prtidx="9999"

dim company_upjong,company_uptae
company_upjong  = Left(html2db(request("company_upjong")),32)
company_uptae   = Left(html2db(request("company_uptae")),25)

dim subid
subid   = request("subid")

dim mode
mode = request("mode")

dim commission,password,passwordS
commission = request("commission")
password = requestCheckVar(request("password"),32)
'passwordS = requestCheckVar(request("passwordS"),32)
if (commission="") then commission=0

'//ISMS 권한 체크용(2021.06.01 원승현)
If Not(C_ADMIN_AUTH or C_MD_AUTH Or C_MD Or C_SYSTEM_Part Or C_MngPart Or C_OP or C_logics_Part or C_OFF_AUTH or C_CSUser) Then
	response.write "<script>alert('Error - 수정하실 수 있는 권한이 없습니다.');history.back();</script>"
	response.End
End If

'//패스워드 정책 검사(2008.12.15;허진원)
if mode="edit" or mode="addnewupchebrand" then
    	if chkPasswordComplex(uid,password)<>"" then
    		response.write "<script language='javascript'>" &vbCrLf &_
    						"	alert('" & chkPasswordComplex(uid,password) & "\n비밀번호를 확인후 다시 시도해주세요.');" &vbCrLf &_
    						" 	history.back();" &vbCrLf &_
    						"</script>"
    		dbget.close()	:	response.End
    	end if

		' if chkPasswordComplex(uid,passwordS)<>"" then
		' 	response.write "<script language='javascript'>" &vbCrLf &_
		' 					"	alert('" & chkPasswordComplex(uid,passwordS) & "\n2차 비밀번호를 확인후 다시 시도해주세요.');" &vbCrLf &_
		' 					" 	history.back();" &vbCrLf &_
		' 					"</script>"
		' 	dbget.close()	:	response.End
    	' end if
end if

dim socname_kor, socname, socname_use, isusing, isextusing, streetusing, isoffusing
dim extstreetusing, specialbrand, maeipdiv, defaultmargine
dim M_margin, W_margin, U_margin

socname_kor  = stripHTML(html2db(trim(request("socname_kor"))))
socname		 = stripHTML(html2db(trim(request("socname"))))
socname_use	 = request("socname_use")
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
if (socname_use="") then socname_use="E"

dim sqlStr, idExists
dim Enc_userpass, Enc_userpass64, Enc_2userpass64
dim prePrtIdx

Enc_userpass = MD5(password)
Enc_userpass64 = SHA256(MD5(password))
'Enc_2userpass64= SHA256(MD5(passwordS))
'if (IsNumeric(prtidx)) and (prtidx<>"") then
'    prtidx = Format00(4,prtidx)
'end if

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

	''sqlStr = "select top 1 userid,IsNULL(prtidx,'9999') as prtidx from [db_user].[dbo].tbl_user_c" + VbCrlf
	''sqlStr = sqlStr + " where userid='" + uid + "'" + VbCrlf
	''rsget.Open sqlStr, dbget, 1
	''if Not rsget.Eof then
	''	prePrtIdx = rsget("prtidx")
	''	prePrtIdx = Format00(4,prePrtIdx)
	''end if
	''rsget.Close

	sqlStr = "select top 1 id from [db_partner].[dbo].tbl_partner" + VbCrlf
	sqlStr = sqlStr + " where id='" + uid + "'" + VbCrlf

	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		idExists = true
	end if
	rsget.Close

	if (idExists) then
	    ''매입처 정보
	    if (pcuserdiv="9999_02") or (pcuserdiv="9999_14") or (pcuserdiv="9999_15") then  ''9999_15 추가 2016/05/16
    	    ''정산정보 관련 -> 그룹(업체) 정보에서 변경 가능.
    		sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
    	'''	sqlStr = sqlStr + " set password='" + password + "'" + VbCrlf  ''패스워드 변경 팝업창으로 2014/03/04
    		sqlStr = sqlStr + " set lastInfoChgDT=getdate(), M_margin=" + CStr(M_margin)  + VbCrlf
    		sqlStr = sqlStr + " ,W_margin=" + CStr(W_margin)  + VbCrlf
    		sqlStr = sqlStr + " ,U_margin=" + CStr(U_margin)  + VbCrlf
    		sqlStr = sqlStr + " ,deliver_name='" + deliver_name + "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_email='" + deliver_email + "'" + VbCrlf
    		sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,purchasetype='" + vPurchaseType+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,offcatecode='" + vOffCateCode+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,offmduserid='" + vOffMDUserID+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,selltype=" + selltype+ "" + VbCrlf
    		if (defaultsongjangdiv<>"") then
    		    sqlStr = sqlStr + " ,defaultsongjangdiv='" & defaultsongjangdiv + "'" + VbCrlf
    		end if
    		if (tplcompanyid<>"") then
    		    sqlStr = sqlStr + " ,tplcompanyid='"&tplcompanyid&"'" + VbCrlf
    		else
    		    sqlStr = sqlStr + " ,tplcompanyid=NULL" + VbCrlf
    	    end if
			sqlStr = sqlStr + " , lastadminid='"& adminid &"'" + VbCrlf
    		sqlStr = sqlStr + " where id='" + uid + "'" + VbCrlf
    		rsget.Open sqlStr, dbget, 1
    ''rw sqlStr

    		sqlStr = "update [db_user].[dbo].tbl_user_c" + VbCrlf
    		''sqlStr = sqlStr + " set prtidx='" + CStr(prtidx) + "'" + VbCrlf
    		sqlStr = sqlStr + " set socname_kor='" + CStr(socname_kor) + "'" + VbCrlf
    		sqlStr = sqlStr + " ,socname='" + CStr(socname)  + "'" + VbCrlf
			sqlStr = sqlStr + " ,socname_use='" + CStr(socname_use)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,isusing='" + CStr(isusing)  + "'" + VbCrlf
    		''sqlStr = sqlStr + " ,isextusing='" + CStr(isextusing)  + "'" + VbCrlf					'// 수정불가 : 제휴몰별 판매설정 팝업창에서만 수정가능
    		sqlStr = sqlStr + " ,streetusing='" + CStr(streetusing)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,extstreetusing='" + CStr(extstreetusing)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,specialbrand='" + CStr(specialbrand)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,maeipdiv='" + CStr(maeipdiv)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,defaultmargine=" + CStr(defaultmargine)  + VbCrlf
    		''sqlStr = sqlStr + " ,userdiv='" + CStr(c_userdiv)  + "'" + VbCrlf     ''수정불가(pcUserDiv)
    		sqlStr = sqlStr + " ,catecode='" + CStr(catecode)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,mduserid='" + CStr(mduserid)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,onlyflg='" + CStr(onlyflg)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,artistflg='" + CStr(artistflg)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,kdesignflg='" + CStr(kdesignflg)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,isoffusing='" + CStr(isoffusing)  + "'" + VbCrlf   '' 2016/05/27
			sqlStr = sqlStr + " ,standardmdcatecode='" + CStr(standardmdcatecode)  + "' where" + VbCrlf
    		sqlStr = sqlStr + " userid='" + uid + "'" + VbCrlf

			'response.write sqlStr & "<br>"
			dbget.execute sqlStr

            Call RF_SetBrandRackCode(uid, prtidx)

''''패스워드 변경 팝업창으로 2014/03/04
'            '' tbl_user_n 에 값이 없는경우만 && tbl_user_c 에 값이 있는경우만 // 20120813 서동석



    		sqlStr = "select top 1 groupid from [db_partner].[dbo].tbl_partner where id = '" & uid & "' "
    		rsget.Open sqlStr,dbget,1
    			if not rsget.Eof then
    				groupid = rsget("groupid")
    			else
    				groupid = ""
    			end if
    		rsget.Close

    		'그룹 대표 반품담당자 정보를 가장 최근 브랜드등록정보로 덮어쒸운다.(skyer9)
    		'과거 데이타를 그대로 두는것보다 그냥 덮어 쒸우는게 낫다.
    		sqlStr = "update [db_partner].[dbo].tbl_partner_group" + VbCrlf
    		sqlStr = sqlStr + " set deliver_name='" + deliver_name+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,lastupdate = getdate()" + VbCrlf
    		sqlStr = sqlStr + " where groupid='" + groupid + "'"
    		rsget.Open sqlStr,dbget,1

            ''상품랙코드 정보 업데이트.(매입처만 존재)
    		''if (CStr(prePrtIdx)<>CStr(prtidx)) then
    		''	sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
    		''	sqlStr = sqlStr + " set itemrackcode='" + CStr(prtidx) + "'" + VbCrlf
    		''	sqlStr = sqlStr + " where makerid='" + uid + "'" + VbCrlf
    		''	sqlStr = sqlStr + " and itemrackcode='" + CStr(prePrtIdx) + "'"

    		''	dbget.Execute sqlStr
    		''end if

	    end if

        ''협력업체
        if (pcuserdiv="902_21") then
	        sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
    		'''sqlStr = sqlStr + " set password='" + password + "'" + VbCrlf
    		'''sqlStr = sqlStr + " ,M_margin=" + CStr(M_margin)  + VbCrlf
    		'''sqlStr = sqlStr + " ,W_margin=" + CStr(W_margin)  + VbCrlf
    		'''sqlStr = sqlStr + " ,U_margin=" + CStr(U_margin)  + VbCrlf
    		sqlStr = sqlStr + " set lastInfoChgDT=getdate(), deliver_name='" + deliver_name + "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_email='" + deliver_email + "'" + VbCrlf
    		sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,purchasetype='" + vPurchaseType+ "'" + VbCrlf
    		'''sqlStr = sqlStr + " ,offcatecode='" + vOffCateCode+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,offmduserid='" + vOffMDUserID+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,selltype=" + selltype+ "" + VbCrlf
    		sqlStr = sqlStr + " ,sellBizCd='" + sellBizCd+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,commission="&commission/100&VbCrlf
    		if (defaultsongjangdiv<>"") then
    		    sqlStr = sqlStr + " ,defaultsongjangdiv='" & defaultsongjangdiv + "'" + VbCrlf
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
			sqlStr = sqlStr + " , lastadminid='"& adminid &"'" + VbCrlf
    		sqlStr = sqlStr + " where id='" + uid + "'" + VbCrlf

			'response.write sqlStr & "<br>"
			dbget.execute sqlStr

    		sqlStr = "update [db_user].[dbo].tbl_user_c" & VbCrlf
    		sqlStr = sqlStr & " set socname_kor='" & CStr(socname_kor) & "'" & VbCrlf
    		sqlStr = sqlStr & " ,socname='" & CStr(socname)  & "'" + VbCrlf
			sqlStr = sqlStr & " ,socname_use='" & CStr(socname_use)  & "'" + VbCrlf
			sqlStr = sqlStr & " where userid='" & uid + "'" & VbCrlf

			'response.write sqlStr & "<br>"
			dbget.execute sqlStr
        end if

	    '''매출처
	    if (pcuserdiv="501_21") or (pcuserdiv="502_21") or (pcuserdiv="503_21") or (pcuserdiv="900_21") or (pcuserdiv="999_50") then

	        sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
    		'''sqlStr = sqlStr + " set password='" + password + "'" + VbCrlf
    		'''sqlStr = sqlStr + " ,M_margin=" + CStr(M_margin)  + VbCrlf
    		'''sqlStr = sqlStr + " ,W_margin=" + CStr(W_margin)  + VbCrlf
    		'''sqlStr = sqlStr + " ,U_margin=" + CStr(U_margin)  + VbCrlf
    		sqlStr = sqlStr + " set lastInfoChgDT=getdate(), deliver_name='" + deliver_name + "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,deliver_email='" + deliver_email + "'" + VbCrlf
    		sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,purchasetype='" + vPurchaseType+ "'" + VbCrlf
    		'''sqlStr = sqlStr + " ,offcatecode='" + vOffCateCode+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,offmduserid='" + vOffMDUserID+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,selltype=" + selltype+ "" + VbCrlf
    		sqlStr = sqlStr + " ,sellBizCd='" + sellBizCd+ "'" + VbCrlf
    		sqlStr = sqlStr + " ,commission="&commission/100&VbCrlf
    		if (defaultsongjangdiv<>"") then
    		    sqlStr = sqlStr + " ,defaultsongjangdiv='" & defaultsongjangdiv + "'" + VbCrlf
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
			sqlStr = sqlStr + " , lastadminid='"& adminid &"'" + VbCrlf
    		sqlStr = sqlStr + " where id='" + uid + "'" + VbCrlf
    		rsget.Open sqlStr, dbget, 1
    ''rw sqlStr

    		sqlStr = "update [db_user].[dbo].tbl_user_c" + VbCrlf
    		''sqlStr = sqlStr + " set prtidx='" + CStr(prtidx) + "'" + VbCrlf
    		sqlStr = sqlStr + " set socname_kor='" + CStr(socname_kor) + "'" + VbCrlf
    		sqlStr = sqlStr + " ,socname='" + CStr(socname)  + "'" + VbCrlf
			sqlStr = sqlStr + " ,socname_use='" + CStr(socname_use)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,isusing='" + CStr(isusing)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,maeipdiv='" + CStr(maeipdiv)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,defaultmargine=" + CStr(defaultmargine)  + VbCrlf
    		sqlStr = sqlStr + " ,mduserid='" + CStr(mduserid)  + "'" + VbCrlf
    		sqlStr = sqlStr + " ,isoffusing='" + CStr(isoffusing)  + "'" + VbCrlf

    		if (pcuserdiv="999_50") then
    		    if (vDefaultDeliveryType<>"") and (vdefaultFreeBeasongLimit<>"") and (vdefaultDeliverPay<>"") then
        		    sqlStr = sqlStr + ",defaultDeliveryType='"&vDefaultDeliveryType&"'" + vbCrlf
            		sqlStr = sqlStr + ",defaultFreeBeasongLimit="&vdefaultFreeBeasongLimit&vbCrlf
            		sqlStr = sqlStr + ",defaultDeliverPay="&vdefaultDeliverPay&vbCrlf
            	end if
    		end if
    		sqlStr = sqlStr + " where userid='" + uid + "'" + VbCrlf

    		rsget.Open sqlStr, dbget, 1

    		'''sqlStr = sqlStr + " ,isextusing='" + CStr(isextusing)  + "'" + VbCrlf							'// 수정불가 : 제휴몰별 판매설정 팝업창에서만 수정가능
    		'''sqlStr = sqlStr + " ,streetusing='" + CStr(streetusing)  + "'" + VbCrlf
    		'''sqlStr = sqlStr + " ,extstreetusing='" + CStr(extstreetusing)  + "'" + VbCrlf
    		'''sqlStr = sqlStr + " ,specialbrand='" + CStr(specialbrand)  + "'" + VbCrlf
    	    '''sqlStr = sqlStr + " ,userdiv='" + CStr(c_userdiv)  + "'" + VbCrlf     ''수정불가(pcUserDiv)
    		'''sqlStr = sqlStr + " ,catecode='" + CStr(catecode)  + "'" + VbCrlf
    		'''sqlStr = sqlStr + " ,onlyflg='" + CStr(onlyflg)  + "'" + VbCrlf
    		'''sqlStr = sqlStr + " ,artistflg='" + CStr(artistflg)  + "'" + VbCrlf
    		'''sqlStr = sqlStr + " ,kdesignflg='" + CStr(kdesignflg)  + "'" + VbCrlf
    ''rw sqlStr

            ''제휴몰(온라인)인경우
            if (pcuserdiv="999_50") then

                sqlStr = " IF Exists(select * from db_partner.dbo.tbl_partner_addInfo where partnerid='"&uid&"')"+ VbCrlf
                sqlStr = sqlStr + " BEGIN"+ VbCrlf
                sqlStr = sqlStr + " update db_partner.dbo.tbl_partner_addInfo"+ VbCrlf
                sqlStr = sqlStr + " set padminUrl='"&padminUrl&"'"+ VbCrlf
        		sqlStr = sqlStr + " ,padminId='"&padminId&"'"+ VbCrlf
        		''sqlStr = sqlStr + " ,padminPwd='"&padminPwd&"'"+ VbCrlf ''제휴사 비번 사용안함
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

''            '' tbl_user_n 에 값이 없는경우만 && tbl_user_c 에 값이 있는경우만 // 20120813 서동석


    		sqlStr = "select top 1 groupid from [db_partner].[dbo].tbl_partner where id = '" & uid & "' "
    		rsget.Open sqlStr,dbget,1
    			if not rsget.Eof then
    				groupid = rsget("groupid")
    			else
    				groupid = ""
    			end if
    		rsget.Close

    		'그룹 대표 반품담당자 정보를 가장 최근 브랜드등록정보로 덮어쒸운다.(skyer9)
    		'과거 데이타를 그대로 두는것보다 그냥 덮어 쒸우는게 낫다.
    		if (groupid<>"") then
        		sqlStr = "update [db_partner].[dbo].tbl_partner_group" + VbCrlf
        		sqlStr = sqlStr + " set deliver_name='" + deliver_name+ "'" + VbCrlf
        		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
        		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
        		sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
        		sqlStr = sqlStr + " ,lastupdate = getdate()" + VbCrlf
        		sqlStr = sqlStr + " where groupid='" + groupid + "'"
        		rsget.Open sqlStr,dbget,1
            end if

''            ''상품랙코드 정보 업데이트.(매입처만 존재)
''    		if (CStr(prePrtIdx)<>CStr(prtidx)) then
''    			sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
''    			sqlStr = sqlStr + " set itemrackcode='" + CStr(prtidx) + "'" + VbCrlf
''    			sqlStr = sqlStr + " where makerid='" + uid + "'" + VbCrlf
''    			sqlStr = sqlStr + " and itemrackcode='" + CStr(prePrtIdx) + "'"
''
''    			dbget.Execute sqlStr
''    		end if
	    end if
	else
		rw "partner ID 가 존재하지 않습니다."
	end if

elseif mode="groupedit" then
    ''/admin/member/popupcheinfoonly.asp

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

    ''기존 사업자가 있을경우
    if  (psocno<>"") then
        Dim alreadySocNoExists
        if (Replace(psocno,"-","")<>Replace(company_no,"-","")) then
            sqlStr = "select count(*) as cnt from [db_partner].[dbo].tbl_partner_group"
            sqlStr = sqlStr &" where Replace(company_no,'-','')='"&Replace(company_no,"-","")&"'"

            rsget.Open sqlStr,dbget,1
                alreadySocNoExists = rsget("cnt")>0
            rsget.CLose
'rw alreadySocNoExists
            IF (alreadySocNoExists) then
                response.write "<script>alert('사업자 번호 변경 불가.("&company_no&") - 이미 존재하는 사업자 번호.');history.back();</script>"
                dbget.Close() : response.end
            end if
        end if
    end if

    ''주민번호 타입인지. 2016/08/04 수정---------------------------------------------------------------
    if (LEN(TRIM(replace(company_no,"-","")))=13) and (right(company_no,2)="**") then
        ' sqlStr = "select isNULL([db_partner].[dbo].[uf_DecSOCNoPH1](encCompNo),'') as DecCompNo"
        ' sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_group where groupid='"&groupid&"'"
        ' rsget.Open sqlStr,dbget,1
        ' if NOT rsget.Eof then
        '     bufComno = rsget("DecCompNo")
        ' end if
        ' rsget.CLose

		''암호화방식변경.
		sqlStr = "select isNULL(db_cs.[dbo].[uf_DecCompanyNoAES256](encCompNo64),'') as DecCompNo64"
        sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_group_adddata where groupid='"&groupid&"'"
        rsget.Open sqlStr,dbget,1
        if NOT rsget.Eof then
            bufComno = rsget("DecCompNo64")
        end if
        rsget.CLose


        if ((bufComno="") or (LEN(TRIM(replace(bufComno,"-","")))<>13) or (right(bufComno,2)="**")) then
            response.write "<script>alert('사업자(주민) 번호 오류 .("&bufComno&") - 관리자 문의 요망.');history.back();</script>"
            dbget.Close() : response.end
        end if

        company_no = bufComno
    end if
    '' ------------------------------------------------------------------------------------------------

	' 정산담당자 데이터 필수값으로.. 정산시 곤란
	if jungsan_name="" or isnull(jungsan_name) then
		response.write "<script type='text/javascript'>alert('정산담당자명을 입력해 주세요.');history.back();</script>"
		dbget.Close() : response.end
	end if
	if jungsan_phone="" or isnull(jungsan_phone) then
		response.write "<script type='text/javascript'>alert('정산담당자 전화번호를 입력해 주세요.');history.back();</script>"
		dbget.Close() : response.end
	end if
	if jungsan_hp="" or isnull(jungsan_hp) then
		response.write "<script type='text/javascript'>alert('정산담당자 이메일주소를 입력해 주세요.');history.back();</script>"
		dbget.Close() : response.end
	end if
	if jungsan_email="" or isnull(jungsan_email) then
		response.write "<script type='text/javascript'>alert('정산담당자 휴대폰번호를 입력해 주세요.');history.back();</script>"
		dbget.Close() : response.end
	end if

	if (groupid<>"") then
		sqlStr = "update [db_partner].[dbo].tbl_partner_group" + VbCrlf
		sqlStr = sqlStr + " set company_name='" + company_name + "'" + VbCrlf
		sqlStr = sqlStr + " ,company_no='" + socialnoReplace(company_no) + "'" + VbCrlf
		sqlStr = sqlStr + " ,ceoname='" + ceoname + "'" + VbCrlf
		sqlStr = sqlStr + " ,company_uptae='" + company_uptae+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_upjong='" + company_upjong+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_zipcode='" + company_zipcode+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_address='" + company_address+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_address2='" + company_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_tel='" + company_tel+ "'" + VbCrlf
		sqlStr = sqlStr + " ,company_fax='" + company_fax+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode + "'" + VbCrlf          ''사무실 주소
		sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_gubun='" + jungsan_gubun+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_bank='" + jungsan_bank+ "'" + VbCrlf
		
		if (C_ADMIN_AUTH) or (C_MngPart) then
		sqlStr = sqlStr + " ,jungsan_acctname='" + jungsan_acctname+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_acctno='" + jungsan_acctno+ "'" + VbCrlf
		end if

		sqlStr = sqlStr + " ,jungsan_date='" + jungsan_date+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_date_off='" + jungsan_date_off+ "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_name='" + manager_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_phone='" + manager_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_hp='" + manager_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,manager_email='" + manager_email+ "'" + VbCrlf
		'배송 담당자 정보는 브랜드별로만 수정 가능(skyer9)
		'sqlStr = sqlStr + " ,deliver_name='" + deliver_name+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_name='" + jungsan_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_phone='" + jungsan_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_hp='" + jungsan_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_email='" + jungsan_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate = getdate()" + VbCrlf
		sqlStr = sqlStr + " ,encCompNo=[db_partner].[dbo].[uf_EncSOCNoPH1]('"&company_no&"')"
		sqlStr = sqlStr + " where groupid='" + groupid + "'"
		dbget.Execute sqlStr
''rw sqlStr

		if (LEN(Trim(replace(company_no,"-","")))=13) then
			sqlStr = "exec [db_cs].[dbo].[usp_Ten_partner_Enc_companyno] '"&groupid&"','"&company_no&"'"
			dbget.Execute sqlStr
		end if


		sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
		sqlStr = sqlStr + " set lastInfoChgDT=getdate(), groupid='" + groupid + "'"
		sqlStr = sqlStr + " , lastadminid='"& adminid &"'" + VbCrlf
		sqlStr = sqlStr + " where id='" + uid + "'" + VbCrlf
		dbget.Execute sqlStr

		''같은 그룹 업체 업데이트.(반품주소,배송담당자정보는 제외)
		sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
		sqlStr = sqlStr + " set lastInfoChgDT=getdate(), company_name='" + company_name + "'" + VbCrlf
		sqlStr = sqlStr + " ,ceoname='" + ceoname + "'" + VbCrlf
		sqlStr = sqlStr + " ,company_no='" + socialnoReplace(company_no) + "'" + VbCrlf
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
		'sqlStr = sqlStr + " ,deliver_name='" + deliver_name + "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_name='" + jungsan_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_phone='" + jungsan_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_hp='" + jungsan_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_email='" + jungsan_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_gubun='" + jungsan_gubun+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_bank='" + jungsan_bank+ "'" + VbCrlf

		if (C_ADMIN_AUTH) or (C_MngPart) then
		sqlStr = sqlStr + " ,jungsan_acctname='" + jungsan_acctname+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_acctno='" + jungsan_acctno+ "'" + VbCrlf
	    end if

		if (jungsan_date<>"") then
		    sqlStr = sqlStr + " ,jungsan_date='" + jungsan_date+ "'" + VbCrlf
	    end if

	    if (jungsan_date_off<>"") then
		    sqlStr = sqlStr + " ,jungsan_date_off='" + jungsan_date_off+ "'" + VbCrlf
		    sqlStr = sqlStr + " ,jungsan_date_frn='" + jungsan_date_off+ "'" + VbCrlf
		end if

		'sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " , lastadminid='"& adminid &"'" + VbCrlf
		sqlStr = sqlStr + " where groupid='" + groupid + "'"

	    dbget.Execute sqlStr

	    ''브랜드별 정산벙보가 다른CASE //2016/12/14 추가
	    if (request.form("jungsan_add_brand").count>0) then
	        for j=1 to request.form("jungsan_add_brand").count
    	        jungsan_add_brand       = trim(request.form("jungsan_add_brand")(j))
    	        jungsan_bank_add        = trim(request.form("jungsan_bank_add")(j))
                jungsan_acctno_add      = trim(request.form("jungsan_acctno_add")(j))
                jungsan_acctname_add    = trim(request.form("jungsan_acctname_add")(j))
                jungsan_date_add        = trim(request.form("jungsan_date_add")(j))
                jungsan_date_off_add    = trim(request.form("jungsan_date_off_add")(j))

                ''rw jungsan_add_brand&"|"&jungsan_bank_add&"|"&jungsan_acctno_add&"|"&jungsan_acctname_add&"|"&jungsan_date_add&"|"&jungsan_date_off_add

				if jungsan_add_brand<>"" and jungsan_bank_add<>"" and jungsan_acctno_add<>"" and jungsan_acctname_add<>"" then
					sqlStr = " if not Exists(select * from db_partner.dbo.tbl_partner_addJungsanInfo where partnerid='"&jungsan_add_brand&"')" &vbCRLF
					sqlStr = sqlStr + " BEGIN"&vbCRLF
					sqlStr = sqlStr + " insert into db_partner.dbo.tbl_partner_addJungsanInfo"&vbCRLF
					sqlStr = sqlStr + " (partnerid,jungsan_bank"&vbCRLF

					if (C_ADMIN_AUTH) or (C_MngPart) then
					sqlStr = sqlStr + " ,jungsan_acctno,jungsan_acctname"&vbCRLF
					end if

					sqlStr = sqlStr + " ,jungsan_date,jungsan_date_off)"&vbCRLF
					sqlStr = sqlStr + " values('"&jungsan_add_brand&"'"&vbCRLF
					sqlStr = sqlStr + " ,'"&jungsan_bank_add&"'"&vbCRLF

					if (C_ADMIN_AUTH) or (C_MngPart) then
					sqlStr = sqlStr + " ,'"&jungsan_acctno_add&"'"&vbCRLF
					sqlStr = sqlStr + " ,'"&jungsan_acctname_add&"'"&vbCRLF
					end if

					sqlStr = sqlStr + " ,'"&jungsan_date_add&"'"&vbCRLF
					sqlStr = sqlStr + " ,'"&jungsan_date_off_add&"'"&vbCRLF
					sqlStr = sqlStr + " )"&vbCRLF
					sqlStr = sqlStr + " END"&vbCRLF
					sqlStr = sqlStr + " ELSE"&vbCRLF
					sqlStr = sqlStr + " BEGIN"&vbCRLF
					sqlStr = sqlStr + " update db_partner.dbo.tbl_partner_addJungsanInfo"&vbCRLF
					sqlStr = sqlStr + " set jungsan_bank='"&jungsan_bank_add&"'"&vbCRLF

					if (C_ADMIN_AUTH) or (C_MngPart) then
					sqlStr = sqlStr + " , jungsan_acctno='"&jungsan_acctno_add&"'"&vbCRLF
					sqlStr = sqlStr + " , jungsan_acctname='"&jungsan_acctname_add&"'"&vbCRLF
					end if

					sqlStr = sqlStr + " , jungsan_date='"&jungsan_date_add&"'"&vbCRLF
					sqlStr = sqlStr + " , jungsan_date_off='"&jungsan_date_off_add&"'"&vbCRLF
					sqlStr = sqlStr + " , lastupdate=getdate()"&vbCRLF
					sqlStr = sqlStr + " where partnerid='"&jungsan_add_brand&"'"&vbCRLF
					sqlStr = sqlStr + " END"&vbCRLF
					dbget.Execute sqlStr
				end if
            next
	    end if

	    sqlStr = "update P" + VbCrlf
        sqlStr = sqlStr + " set jungsan_bank=A.jungsan_bank" + VbCrlf

		if (C_ADMIN_AUTH) or (C_MngPart) then
        sqlStr = sqlStr + " ,jungsan_acctname=A.jungsan_acctname" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_acctno=A.jungsan_acctno" + VbCrlf
	    end if

        sqlStr = sqlStr + " ,jungsan_date=A.jungsan_date" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_date_off=A.jungsan_date_off" + VbCrlf
        sqlStr = sqlStr + " ,jungsan_date_frn=A.jungsan_date_off" + VbCrlf
		sqlStr = sqlStr + " , p.lastadminid='"& adminid &"'" + VbCrlf
        sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner P"
        sqlStr = sqlStr + "     Join db_partner.dbo.tbl_partner_addJungsanInfo A"
        sqlStr = sqlStr + "     on P.id=A.partnerid"
        sqlStr = sqlStr + " where P.groupid='" + groupid + "'"

        dbget.Execute sqlStr
	else
		''Get Last Group ID
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
		sqlStr = sqlStr + " jungsan_gubun, jungsan_bank, jungsan_date, jungsan_date_off"

		if (C_ADMIN_AUTH) or (C_MngPart) then
		sqlStr = sqlStr + " , jungsan_acctname, jungsan_acctno"
	    end if

		sqlStr = sqlStr + " ,manager_name, manager_phone, manager_hp, manager_email, deliver_name, deliver_phone, "
		sqlStr = sqlStr + " deliver_hp, deliver_email, jungsan_name, jungsan_phone, jungsan_hp, jungsan_email, "
		sqlStr = sqlStr + " encCompNo,"
		sqlStr = sqlStr + " lastupdate)"
		sqlStr = sqlStr + " values('" + groupid + "'"
		sqlStr = sqlStr + " ,'" + company_name + "'"
		sqlStr = sqlStr + " ,'" + socialnoReplace(company_no) + "'"
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

		if (C_ADMIN_AUTH) or (C_MngPart) then
		sqlStr = sqlStr + " ,'" + jungsan_acctname + "'"
		sqlStr = sqlStr + " ,'" + jungsan_acctno + "'"
	    end if

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
		sqlStr = sqlStr + " ,[db_partner].[dbo].[uf_EncSOCNoPH1]('"&company_no&"')"
		sqlStr = sqlStr + " ,getdate()"
		sqlStr = sqlStr + " )"

		dbget.Execute sqlStr

		if (LEN(Trim(replace(company_no,"-","")))=13) then
			sqlStr = "exec [db_cs].[dbo].[usp_Ten_partner_Enc_companyno] '"&groupid&"','"&company_no&"'"
			dbget.Execute sqlStr
		end if

		if uid<>"" then
			sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
			sqlStr = sqlStr + " set lastInfoChgDT=getdate(), groupid='" + groupid + "'"
			sqlStr = sqlStr + " , lastadminid='"& adminid &"'" + VbCrlf
			sqlStr = sqlStr + " where id='" + uid + "'" + VbCrlf

			rsget.Open sqlStr,dbget,1
		end if

		''같은 그룹 업체 업데이트.
		sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
		sqlStr = sqlStr + " set lastInfoChgDT=getdate(), company_name='" + company_name + "'" + VbCrlf
		sqlStr = sqlStr + " ,ceoname='" + ceoname + "'" + VbCrlf
		sqlStr = sqlStr + " ,company_no='" + socialnoReplace(company_no) + "'" + VbCrlf
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
		'sqlStr = sqlStr + " ,deliver_name='" + deliver_name + "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,deliver_email='" + deliver_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_name='" + jungsan_name+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_phone='" + jungsan_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_hp='" + jungsan_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_email='" + jungsan_email+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_gubun='" + jungsan_gubun+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_bank='" + jungsan_bank+ "'" + VbCrlf

		if (jungsan_date<>"") then
		    sqlStr = sqlStr + " ,jungsan_date='" + jungsan_date+ "'" + VbCrlf
	    end if

	    if (jungsan_date_off<>"") then
		    sqlStr = sqlStr + " ,jungsan_date_off='" + jungsan_date_off+ "'" + VbCrlf
		    sqlStr = sqlStr + " ,jungsan_date_frn='" + jungsan_date_off+ "'" + VbCrlf
		end if

		if (C_ADMIN_AUTH) or (C_MngPart) then
		sqlStr = sqlStr + " ,jungsan_acctname='" + jungsan_acctname+ "'" + VbCrlf
		sqlStr = sqlStr + " ,jungsan_acctno='" + jungsan_acctno+ "'" + VbCrlf
		end if

		'sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
		'sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " , lastadminid='"& adminid &"'" + VbCrlf
		sqlStr = sqlStr + " where groupid='" + groupid + "'"
		dbget.Execute sqlStr
	end if

	'manager_hp = replace(manager_hp,"-","")
	'jungsan_hp = replace(jungsan_hp,"-","")

	sql ="if exists(select userid from db_partner.dbo.tbl_partner_user with (nolock) where isusing='Y' and groupid ='"& groupid &"' and gubun=1)"
	sql = sql & " begin"
	sql = sql & "   update db_partner.dbo.tbl_partner_user set" & vbcrlf
	sql = sql & "   lastUpdate=getdate(),name=N'"& html2db(manager_name) &"',Title=N'일반담당자'" & vbcrlf
	sql = sql & " 	,hp=N'"& html2db(manager_hp) &"'" & vbcrlf
	sql = sql & " 	,email=N'"& html2db(manager_email) &"'" & vbcrlf
	sql = sql & "   where isusing='Y' and groupid ='"& groupid &"' and gubun=1"
	sql = sql & " end"
	sql = sql & " else"
	sql = sql & " begin"
	sql = sql & "   insert into db_partner.dbo.tbl_partner_user (groupid,userid,gubun,Title,name"
	sql = sql & "   ,hp,email"
	sql = sql & "   ,regdate,lastUpdate,isUsing)"
	sql = sql & "       select N'"& groupid &"',NULL,1,N'일반담당자',N'"& html2db(manager_name) &"'"
	sql = sql & "   	,N'"& html2db(manager_hp) &"',N'"& html2db(manager_email) &"'" & vbcrlf
	sql = sql & "       ,getdate(),getdate(),N'Y'"
	sql = sql & " end"

	'response.write sql & "<Br>"
	dbget.Execute sql

	sql ="if exists(select userid from db_partner.dbo.tbl_partner_user with (nolock) where isusing='Y' and groupid ='"& groupid &"' and gubun=2)"
	sql = sql & " begin"
	sql = sql & "   update db_partner.dbo.tbl_partner_user set" & vbcrlf
	sql = sql & "   lastUpdate=getdate(),name=N'"& html2db(jungsan_name) &"',Title=N'정산담당자'" & vbcrlf
	sql = sql & " 	,hp=N'"& html2db(jungsan_hp) &"'" & vbcrlf
	sql = sql & " 	,email=N'"& html2db(jungsan_email) &"'" & vbcrlf
	sql = sql & "   where isusing='Y' and groupid ='"& groupid &"' and gubun=2"
	sql = sql & " end"
	sql = sql & " else"
	sql = sql & " begin"
	sql = sql & "   insert into db_partner.dbo.tbl_partner_user (groupid,userid,gubun,Title,name"
	sql = sql & "   ,hp,email"
	sql = sql & "   ,regdate,lastUpdate,isUsing)"
	sql = sql & "       select N'"& groupid &"',NULL,2,N'정산담당자',N'"& html2db(jungsan_name) &"'"
	sql = sql & "   	,N'"& html2db(jungsan_hp) &"',N'"& html2db(jungsan_email) &"'" & vbcrlf
	sql = sql & "       ,getdate(),getdate(),N'Y'"
	sql = sql & " end"

	'response.write sql & "<Br>"
	dbget.Execute sql

	' 기존 추가담당자 입력되어 있던거 수정
	if etc_idx<>"" then
		etc_idx = split(etc_idx,",")
		etc_name = split(etc_name,",")
		etc_hp = split(etc_hp,",")
		etc_email = split(etc_email,",")
		for i = 0 to ubound(etc_idx)
			tmpetc_idx = trim(etc_idx(i))
			tmpetc_name = trim(etc_name(i))
			tmpetc_hp = trim(etc_hp(i))
			'tmpetc_hp = replace(tmpetc_hp,"-","")
			tmpetc_email = trim(etc_email(i))

			if isnull(tmpetc_idx) or tmpetc_idx="" or tmpetc_idx="0" then
				response.write "<script type='text/javascript'>"
				response.write "    alert('정상적인 경로로 시도해 주세요.');"
				response.write "    history.back();"
				response.write "</script>"
				response.End
			end if
			if isnull(tmpetc_name) or tmpetc_name="" then
				response.write "<script type='text/javascript'>"
				response.write "    alert('추가담당자명을 입력해주세요.');"
				response.write "    history.back();"
				response.write "</script>"
				response.End
			end if
			if isnull(tmpetc_hp) or tmpetc_hp="" then
				response.write "<script type='text/javascript'>"
				response.write "    alert('추가담당자의 휴대폰번호를 입력해 주세요.');"
				response.write "    history.back();"
				response.write "</script>"
				response.End
			end if
			if isnull(tmpetc_email) or tmpetc_email="" then
				response.write "<script type='text/javascript'>"
				response.write "    alert('추가담당자의 이메일 주소를 입력해 주세요.');"
				response.write "    history.back();"
				response.write "</script>"
				response.End
			end if

			exists_login_title=""
			exists_login_gubun=""
			exists_login_hp=""
			exists_login_email=""
			sql = "select top 1 isnull(title,'') as title, gubun, isnull(hp,'') as hp, isnull(email,'') as email"
			sql = sql & " from db_partner.dbo.tbl_partner_user with (nolock)"
			sql = sql & " where isusing='Y'"
			sql = sql & " and gubun=10"		' 추가담당자 내에서만 중복체크
			sql = sql & " and groupid ='"& groupid &"'"
			sql = sql & " and replace(hp,'-','')='"& html2db(replace(tmpetc_hp,"-","")) &"'"
			sql = sql & " and idx<>"& tmpetc_idx &""

			'response.write sql & "<Br>"
			rsget.CursorLocation = adUseClient
			rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
			if not rsget.EOF  then
				exists_login_title =rsget("title")
				exists_login_gubun =rsget("gubun")
				exists_login_hp =db2html(rsget("hp"))
				exists_login_email =db2html(rsget("email"))
			end if
			rsget.close

			if exists_login_title<>"" and exists_login_hp<>"" then
				response.write "<script type='text/javascript'>"
				response.write "    alert('추가담당자 휴대폰번호("& exists_login_hp &")가 이미 "& exists_login_title &" 에 중복으로 등록되어 있습니다.\n수정하시거나 삭제 부탁드립니다.');"
				response.write "    history.back();"
				response.write "</script>"
				response.End
			end if

			' 휴대폰번호 중복체크 성공시, 이메일 중복체크는 하지 않음(영세 업체의 경우 회사이메일을 공유하는 업체가 많다고 클레임). 차후 isms 심사때 이메일 중복체크 요청이 있을 경우 주석해제.
			'exists_login_title=""
			'exists_login_gubun=""
			'exists_login_hp=""
			'exists_login_email=""
			'sql = "select top 1 isnull(title,'') as title, gubun, isnull(hp,'') as hp, isnull(email,'') as email"
			'sql = sql & " from db_partner.dbo.tbl_partner_user with (nolock)"
			'sql = sql & " where isusing='Y'"
			'sql = sql & " and groupid ='"& groupid &"'"
			'sql = sql & " and email='"& html2db(tmpetc_email) &"'"
			'sql = sql & " and idx<>"& tmpetc_idx &""

			''response.write sql & "<Br>"
			'rsget.CursorLocation = adUseClient
			'rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
			'if not rsget.EOF  then
			'	exists_login_title =rsget("title")
			'	exists_login_gubun =rsget("gubun")
			'	exists_login_hp =db2html(rsget("hp"))
			'	exists_login_email =db2html(rsget("email"))
			'end if
			'rsget.close

			'if exists_login_title<>"" and exists_login_email<>"" then
			'	response.write "<script type='text/javascript'>"
			'	response.write "    alert('추가담당자 이메일주소("& exists_login_email &")가 이미 "& exists_login_title &" 에 중복으로 등록되어 있습니다.\n수정하시거나 삭제 부탁드립니다.');"
			'	response.write "    history.back();"
			'	response.write "</script>"
			'	response.End
			'end if

			sql = "update db_partner.dbo.tbl_partner_user set" & vbcrlf
			sql = sql & " lastUpdate=getdate(),name=N'"& html2db(tmpetc_name) &"',Title=N'추가담당자'" & vbcrlf
			sql = sql & " ,hp=N'"& html2db(tmpetc_hp) &"'" & vbcrlf
			sql = sql & " ,email=N'"& html2db(tmpetc_email) &"'" & vbcrlf
			sql = sql & " where isusing='Y' and groupid ='"& groupid &"' and gubun=10 and idx="& tmpetc_idx &""

			'response.write sql & "<Br>"
			dbget.Execute sql
		next
	end if

	' 추가담당자 신규 입력
	if etcaddyn="Y" then
		if isnull(addetc_name) or addetc_name="" then
			response.write "<script type='text/javascript'>"
			response.write "    alert('추가담당자명을 입력해주세요.');"
			response.write "    history.back();"
			response.write "</script>"
			response.End
		end if
		if isnull(addetc_hp) or addetc_hp="" then
			response.write "<script type='text/javascript'>"
			response.write "    alert('추가담당자의 휴대폰번호를 입력해 주세요.');"
			response.write "    history.back();"
			response.write "</script>"
			response.End
		end if
		if isnull(addetc_email) or addetc_email="" then
			response.write "<script type='text/javascript'>"
			response.write "    alert('추가담당자의 이메일 주소를 입력해 주세요.');"
			response.write "    history.back();"
			response.write "</script>"
			response.End
		end if
		'addetc_hp = replace(addetc_hp,"-","")

		exists_login_title=""
		exists_login_gubun=""
		exists_login_hp=""
		exists_login_email=""
		sql = "select top 1 isnull(title,'') as title, gubun, isnull(hp,'') as hp, isnull(email,'') as email"
		sql = sql & " from db_partner.dbo.tbl_partner_user with (nolock)"
		sql = sql & " where isusing='Y'"
		sql = sql & " and gubun=10"		' 추가담당자 내에서만 중복체크
		sql = sql & " and groupid ='"& groupid &"'"
		sql = sql & " and replace(hp,'-','')='"& html2db(replace(addetc_hp,"-","")) &"'"

		'response.write sql & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.EOF  then
			exists_login_title =rsget("title")
			exists_login_gubun =rsget("gubun")
			exists_login_hp =db2html(rsget("hp"))
			exists_login_email =db2html(rsget("email"))
		end if
		rsget.close

		if exists_login_title<>"" and exists_login_hp<>"" then
			response.write "<script type='text/javascript'>"
			response.write "    alert('신규로 등록하신 추가담당자 휴대폰번호("& exists_login_hp &")가 이미 "& exists_login_title &" 에 등록되어 있습니다.\n중복 등록은 불가 합니다.');"
			response.write "    history.back();"
			response.write "</script>"
			response.End
		end if

		' 휴대폰번호 중복체크 성공시, 이메일 중복체크는 하지 않음(영세 업체의 경우 회사이메일을 공유하는 업체가 많다고 클레임). 차후 isms 심사때 이메일 중복체크 요청이 있을 경우 주석해제.
		'exists_login_title=""
		'exists_login_gubun=""
		'exists_login_hp=""
		'exists_login_email=""
		'sql = "select top 1 isnull(title,'') as title, gubun, isnull(hp,'') as hp, isnull(email,'') as email"
		'sql = sql & " from db_partner.dbo.tbl_partner_user with (nolock)"
		'sql = sql & " where isusing='Y'"
		'sql = sql & " and groupid ='"& groupid &"'"
		'sql = sql & " and email='"& html2db(addetc_email) &"'"

		''response.write sql & "<Br>"
		'rsget.CursorLocation = adUseClient
		'rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		'if not rsget.EOF  then
		'	exists_login_title =rsget("title")
		'	exists_login_gubun =rsget("gubun")
		'	exists_login_hp =db2html(rsget("hp"))
		'	exists_login_email =db2html(rsget("email"))
		'end if
		'rsget.close

		'if exists_login_title<>"" and exists_login_email<>"" then
		'	response.write "<script type='text/javascript'>"
		'	response.write "    alert('신규로 등록하신 추가담당자 이메일주소("& exists_login_email &")가 이미 "& exists_login_title &" 에 등록되어 있습니다.\n중복 등록은 불가 합니다.');"
		'	response.write "    history.back();"
		'	response.write "</script>"
		'	response.End
		'end if

		sql = "insert into db_partner.dbo.tbl_partner_user (groupid,userid,gubun,Title,name"
		sql = sql & " ,hp,email"
		sql = sql & " ,regdate,lastUpdate,isUsing)"
		sql = sql & "   select N'"& groupid &"',NULL,10,N'추가담당자',N'"& html2db(addetc_name) &"'"
		sql = sql & "   ,N'"& html2db(addetc_hp) &"',N'"& html2db(addetc_email) &"'" & vbcrlf
		sql = sql & "   ,getdate(),getdate(),N'Y'"

		' response.write sql & "<Br>"
		dbget.Execute sql
	end if

elseif mode="etc_del" then

	if (groupid<>"") then
		if etc_idx<>"" then
			if isnull(etc_idx) or etc_idx="" or etc_idx="0" then
				response.write "<script type='text/javascript'>"
				response.write "    alert('정상적인 경로로 시도해 주세요.\n지정된 번호가 없습니다.');"
				response.write "</script>"
				response.End
			end if

			sql = "update db_partner.dbo.tbl_partner_user set" & vbcrlf
			sql = sql & " lastUpdate=getdate(),isusing='N'" & vbcrlf
			sql = sql & " where isusing='Y' and groupid ='"& groupid &"' and gubun=10 and idx="& etc_idx &""

			'response.write sql & "<Br>"
			dbget.Execute sql
		end if
	else
		response.write "<script>alert('Error - 그룹코드 없음 관리자 문의요망');</script>"
		dbget.close()	:	response.End
	end if

elseif mode="modifyreturnaddress" then

		uid = requestCheckVar(request("uid"),32)

		defaultsongjangdiv = requestCheckVar(request("defaultsongjangdiv"),4)

		deliver_name = requestCheckVar(request("deliver_name"),12)
		deliver_phone = requestCheckVar(request("deliver_phone"),16)
		deliver_hp = requestCheckVar(request("deliver_hp"),16)
		deliver_email = requestCheckVar(request("deliver_email"),128)

		return_zipcode = requestCheckVar(request("return_zipcode"),8)
		return_address = requestCheckVar(request("return_address"),128)
		return_address2 = requestCheckVar(request("return_address2"),128)

		'deliver_hp = replace(deliver_hp,"-","")

		sqlStr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
		sqlStr = sqlStr + " set lastInfoChgDT=getdate(), defaultsongjangdiv='" & defaultsongjangdiv + "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_name='" + deliver_name + "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_phone='" + deliver_phone+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_hp='" + deliver_hp+ "'" + VbCrlf
		sqlStr = sqlStr + " ,deliver_email='" + deliver_email + "'" + VbCrlf
		sqlStr = sqlStr + " ,return_zipcode='" + return_zipcode+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address='" + return_address+ "'" + VbCrlf
		sqlStr = sqlStr + " ,return_address2='" + return_address2+ "'" + VbCrlf
		sqlStr = sqlStr + " , lastadminid='"& adminid &"'" + VbCrlf
		sqlStr = sqlStr + " where id='" + uid + "'"
		rsget.Open sqlStr,dbget,1

		sqlStr = "select top 1 groupid from [db_partner].[dbo].tbl_partner where id = '" & uid & "' "
		rsget.Open sqlStr,dbget,1
			if not rsget.Eof then
				groupid = rsget("groupid")
			else
				groupid = ""
			end if
		rsget.Close

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

		'response.write sqlStr

		sql ="if exists(select userid from db_partner.dbo.tbl_partner_user with (nolock) where isusing='Y' and groupid ='"& groupid &"' and gubun=3 and userid='"& uid &"')"
		sql = sql & " begin"
		sql = sql & "   update db_partner.dbo.tbl_partner_user set" & vbcrlf
		sql = sql & "   lastUpdate=getdate(),name=N'"& html2db(deliver_name) &"',Title=N'배송담당자'" & vbcrlf
		sql = sql & " ,hp=N'"& html2db(deliver_hp) &"'" & vbcrlf
		sql = sql & " ,email=N'"& html2db(deliver_email) &"'" & vbcrlf
		sql = sql & "   where isusing='Y' and groupid ='"& groupid &"' and gubun=3 and userid='"& uid &"'"
		sql = sql & " end"
		sql = sql & " else"
		sql = sql & " begin"
		sql = sql & "   insert into db_partner.dbo.tbl_partner_user (groupid,userid,gubun,Title,name"
		sql = sql & "   ,hp,email"
		sql = sql & "   ,regdate,lastUpdate,isUsing)"
		sql = sql & "       select N'"& groupid &"',N'"& uid &"',3,N'배송담당자',N'"& html2db(deliver_name) &"'"
		sql = sql & "   	,N'"& html2db(deliver_hp) &"',N'"& html2db(deliver_email) &"'" & vbcrlf
		sql = sql & "       ,getdate(),getdate(),N'Y'"
		sql = sql & " end"

		'response.write sql & "<Br>"
		dbget.Execute sql

elseif mode="modiprevmonthgroupid" Then

		sqlStr = " update m1 "
		sqlStr = sqlStr + " set m1.groupid = '" & groupid & "', m1.lastupdate = getdate() "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_partner].[dbo].[tbl_monthly_brandInfo] m1 "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m1.makerid = '" & uid & "' "
		sqlStr = sqlStr + " 	and m1.yyyymm >= convert(varchar(7), dateadd(m, -1, getdate()), 121) "
		sqlStr = sqlStr + " 	and m1.groupid <> '" & groupid & "' "
		rsget.Open sqlStr,dbget,1

elseif mode="addnewupchebrand" then
    ''/admin/member/addnewbrand.asp

	'// 아이디 중복 확인
	sqlStr = "select count(*) from [db_user].[dbo].tbl_logindata with(nolock) where userid='" & uid & "'"
	rsget.Open sqlStr,dbget,1
	if rsget(0)>0 then
		response.write "<script type='text/javascript'>" &vbCrLf &_
						"	alert('이미 존재하거나 [일반고객]과 중복되는 아이디입니다. 다른 아이디를 입력해주세요.');" &vbCrLf &_
						" 	history.back();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	rsget.Close

	sqlStr = "select count(*) from [db_user].[dbo].tbl_deluser with(nolock) where userid='" & uid & "'"
	rsget.Open sqlStr,dbget,1
	if rsget(0)>0 then
		response.write "<script type='text/javascript'>" &vbCrLf &_
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
	sqlStr = sqlStr + "defaultmargine, socname_kor,socname_use, " & vbCrlf
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
	sqlStr = sqlStr + ",'" + socname_use + "'" + vbCrlf

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
	sqlStr = sqlStr + "(id,Enc_password,Enc_password64,userdiv,jungsan_date,groupid"+ vbCrlf	' ,Enc_2password64
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
	'sqlStr = sqlStr + " ,'" + Enc_2userpass64 + "'" + vbCrlf '--암호화 고도화 2014/07/21
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
	sqlStr = sqlStr + " , lastadminid='"& adminid &"'" + VbCrlf
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
		sqlStr = sqlStr + " , lastadminid='"& adminid &"'" + VbCrlf
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

    ''2013/12/08 추가마진 관련 추가 eastone
    if (maeipdiv<>"") and (defaultmargine<>"") then
        sqlStr = " update [db_partner].[dbo].tbl_partner"
        sqlStr = sqlStr & " SET lastInfoChgDT=getdate(), "&maeipdiv&"_margin="&defaultmargine
		sqlStr = sqlStr + " , lastadminid='"& adminid &"'" + VbCrlf
        sqlStr = sqlStr + " where id='"&uid&"'"
		dbget.Execute sqlStr
    end if

	' 직원, 브랜드 변경로그
	fnChkauthlog "", uid, "10", "SCM 브랜드 권한생성", adminid
end if
' response.write Err.Description
'response.end
	If Err.Number = 0 Then
	        dbget.CommitTrans
	Else
	        dbget.RollBackTrans
	        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\n입력한 값들이 너무 길지 않는지 확인바랍니다.\n주로 업태와 업종에서 에러가 자주 나타납니다.')</script>"
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

%>

<script>alert('저장되었습니다.');</script>
<% if mode="addnewupchebrand" then %>
    <script>top.location.href='/admin/member/popbrandinfoonly.asp?designer=<%= uid %>';</script>
<% else %>
    <script>location.replace('<%= refer %>');</script>
<% end if %>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->
