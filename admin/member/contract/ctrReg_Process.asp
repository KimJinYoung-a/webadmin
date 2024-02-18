<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드 계약 관리
' Hieditor : 2009.04.07 서동석 생성
'			 2010.05.26 한용민 수정
' 			 2017.06.23 정윤정 U+전자계약 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/ecContractApi_function.asp"-->
<%
	dim mode, makerid, contractType, ctrState , contractEtcContetns , onoffgubun, subType, ctrcount
	dim mailfrom, mailto, mailtitle, mailcontent, innerContents ,CurrState,NextState, sendOpenMail
	dim ctrKey, groupid, isDefaultContract, ctrNo
	dim addmwdiv, addsellplace, addmargin, addON_ctrDate, addOF_ctrDate, addON_dlvtype, addON_dlvlimit, addON_dlvpay, ctrSubKey
	dim addONOF_endDate , ictrMon
	dim AssignedRow
	dim ogroupInfo
	dim i,kk, chkmwdiv, chkmwdivMExists, chkmwdivWExists
	dim chkCtr, cnt
	dim ckHp,mngHp,ckEmail,mngEmail, noMatCnt, iCtrKeyArr
	dim chkCT11, chkCT12, chkCT13
	dim addOF_endDate,addON_endDate
	dim signtype,ectypeSeq
	dim ecAUser, ecBUser,enddate
	dim bcompno
	dim ecId, ecPwd
	Dim access_token, token_type, refresh_token
	Dim APIpath, jsResult
	Dim objXML, xmlDOM, iRbody, strParam, mngName

	ctrcount=0

    ctrKey          = requestCheckvar(request("ctrKey"),32)
	mode            = requestCheckvar(request("mode"),32)
	makerid         = requestCheckvar(request("makerid"),32)
	contractType    = requestCheckvar(request("contractType"),10)
	contractEtcContetns = request("contractEtcContetns")
	CurrState       = requestCheckvar(request("CurrState"),32)
	NextState       = requestCheckvar(request("NextState"),32)
	sendOpenMail    = requestCheckvar(request("sendOpenMail"),32)

	groupid         = requestCheckvar(request("groupid"),10)
    addmwdiv        = requestCheckvar(request("addmwdiv"),500)
    addsellplace    = requestCheckvar(request("addsellplace"),500)
    addmargin       = requestCheckvar(request("addmargin"),500)
    addON_ctrDate   = requestCheckvar(request("addON_ctrDate"),10)
    addOF_ctrDate   = requestCheckvar(request("addOF_ctrDate"),10)
    addON_endDate   = requestCheckvar(request("addON_endDate"),10)
    addOF_endDate   = requestCheckvar(request("addOF_endDate"),10)

    addON_dlvtype   = requestCheckvar(request("addON_dlvtype"),10)
    addON_dlvlimit  = requestCheckvar(request("addON_dlvlimit"),10)
    addON_dlvpay    = requestCheckvar(request("addON_dlvpay"),10)

    ctrSubKey       = requestCheckvar(request("ctrSubKey"),500)
    chkCtr = requestCheckvar(request("chkCtr"),500)

    ckHp        = requestCheckvar(request("ckHp"),10)
    mngHp       = requestCheckvar(request("mngHp"),20)
    ckEmail     = requestCheckvar(request("ckEmail"),10)
    mngEmail    = requestCheckvar(request("mngEmail"),100)
    noMatCnt    = requestCheckvar(request("noMatCnt"),10)
	mngName		= requestCheckvar(request("mngName"),50)

    chkCT11     = requestCheckvar(request("chkCT11"),1)
    chkCT12     = requestCheckvar(request("chkCT12"),1)
	chkCT13     = requestCheckvar(request("chkCT13"),1)

   if chkCT11 = "" then chkCT11 = 0
   if chkCT12 = "" then chkCT12 = 0
   if chkCT13 = "" then chkCT13 = 0

   	signtype 		= requestCheckvar(request("signtype"),1)
   	ecAUser = requestCheckvar(request("ecAUser"),32)
   	ecBUser = requestCheckvar(request("ecBUser"),32)
   	ecId =  requestCheckvar(request("LgUID"),32)
   	ecPwd =  requestCheckvar(request("LgUPW"),32)
 '   addOn_ecAuser	   = requestCheckvar(request("addOn_ecAuser"),32)
'		addOn_ecBuser	   = requestCheckvar(request("addOn_ecBuser"),32)

    bcompno = replace(requestCheckvar(request("bcompno"),32),"-","")

	dim sqlStr , objItem, contractExists , contractContents ,contractName, HtmlcontractEtcContetns
	dim bufStr, refer
	dim ocontract,oMdInfoList
	dim userStatus
	dim oneContract,acctoken,reftoken
	dim A_COMPANY_NO, A_UPCHENAME, A_CEONAME, B_COMPANY_NO, B_UPCHENAME, B_CEONAME,DEFAULT_JUNGSANDATE,A_COMPANY_ADDR,B_COMPANY_ADDR,CONTRACT_DATE
	dim ecCtrSeq, strErrMsg
	dim authexCount, authexType, docuSignSendContents, docuSignContractData
	dim docuSignEnvelopeId, docuSignStatus, docuSignStatusDateTime, docuSignUri
	refer = request.ServerVariables("HTTP_REFERER")

	if (mode="reg") then
		'// 기본 구성 및 검증
		Select Case signtype
			''==============U+ 전자계약 기본세팅 ================
			Case "2"
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
						Response.write "<script type='text/javascript' language='javascript'>alert('전자계약 계정정보가 잘못입력되었습니다. 확인 후 다시 시도해주세요,');location.href = '"&refer&"';</script>"
						response.end
					end if
				end if

				'회원체크
				userStatus= 	fnCheckUser(bcompno,acctoken)

				if Fchkerror ="invalid_token" then
					call sbGetRefToken(reftoken)
					acctoken = Faccess_token
					userStatus= 	fnCheckUser(bcompno,acctoken)
				end if

				if userStatus <> "사용중" then
					Response.write "<script type='text/javascript' language='javascript'>alert('["&userStatus&"]: 업체가 LG U+ 전자계약 사이트에 가입되어있지 않습니다. 가입 확인후 계약서 생성이 가능합니다,');location.href = '"&refer&"';</script>"
					response.end
				end if

				if chkCT11 =1  or chkCT12 =1 then
					For Each objItem In Request.Form
						if (Left(objItem,2)="$$") and (Right(objItem,2)="$$") then
							if  objItem = "$$A_COMPANY_NO$$" then
								A_COMPANY_NO = replace(Request.Form(objItem),"-","")
							elseif  objItem = "$$A_UPCHENAME$$" then
								A_UPCHENAME = Request.Form(objItem)
							elseif  objItem = "$$A_CEONAME$$" then
								A_CEONAME = Request.Form(objItem)
							elseif  objItem = "$$A_COMPANY_ADDR$$" then
								A_COMPANY_ADDR = Request.Form(objItem)
							elseif  objItem = "$$B_COMPANY_NO$$" then
								B_COMPANY_NO = replace(Request.Form(objItem) ,"-","")
							elseif  objItem = "$$B_UPCHENAME$$" then
								B_UPCHENAME = Request.Form(objItem)
							elseif  objItem = "$$B_CEONAME$$" then
								B_CEONAME = Request.Form(objItem)
							elseif objItem = "$$B_COMPANY_ADDR$$"	then
								B_COMPANY_ADDR = Request.Form(objItem)
							elseif objItem = "$$CONTRACT_DATE$$" then
								CONTRACT_DATE   = Request.Form(objItem)
							'	CONTRACT_DATE   = Left(CONTRACT_DATE,4) & "년 " & Mid(CONTRACT_DATE,6,2) & "월 " & Mid(CONTRACT_DATE,9,2) & "일"
							elseif objItem ="$$ENDDATE$$"			then
								ENDDATE   = Request.Form(objItem)
							'	ENDDATE   = Left(ENDDATE,4) & "년 " & Mid(ENDDATE,6,2) & "월 " & Mid(ENDDATE,9,2) & "일"
							elseif objItem = "$$DEFAULT_JUNGSANDATE$$" then
								DEFAULT_JUNGSANDATE = 	Request.Form(objItem)
							end if
						end if
					Next
				end if
		
			''============== DocuSign 전자서명 ================
			Case "3"

				if mngEmail="" then
					Response.Write "<script>alert(""수령자의 이메일이 없습니다. 확인 후 다시 시도해주세요."");history.back();</script>"
					Response.END
				end if

				if ecBUser="" then
					Response.Write "<script>alert(""계약자명이 없습니다. 확인 후 다시 시도해주세요."");history.back();</script>"
					Response.END
				end if
		end SELECT

		''==============================================================
		''============  계약서 작성  ====================================
		''==============================================================

		''기본계약서-----------------------------------------------------------------------------------------------------------------
		if chkCT11 = 1 then

			contractType = DEFAULT_CONTRACTTYPE '수기형 계약서번호

			''DocuSign 계약일 경우 새로운 계약서 번호를 사용
			If signtype = "3" Then
				contractType = DEFAULT_NEWCONTRACTTYPE '거래기본계약서(2021.11)
			End If

			'' 계약서 폼을 불러온다.
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

			''LG U+ 전자결재
			if signtype ="2" then
				ectypeSeq = Fec_defctrtype 'lg u+ 계약서번호
				ecCtrSeq = 0

				dim con_status, con_info, tmpCallBack,strParam1
				APIpath =FecURL&"/api/createCont"

				strParam = "?type_seq="&ectypeSeq&"&cancel_limit=0&contract_dt="&CONTRACT_DATE&"&contract_key=&contract_money=0&expire_dt="&ENDDATE
				strParam = strParam&"&venderno="&A_COMPANY_NO&"&search_word="&server.URLEncode(B_UPCHENAME)&"&start_dt="&CONTRACT_DATE&"&title="&server.URLEncode(contractName)
				strParam = strParam&"&membList[0].company="&server.URLEncode(A_UPCHENAME)&"&membList[0].gubun=A&membList[0].users="&server.URLEncode(ecAUser)&"&membList[0].venderno="&A_COMPANY_NO
				strParam = strParam&"&membList[1].company="&server.URLEncode(B_UPCHENAME)&"&membList[1].gubun=B&membList[1].users="&server.URLEncode(ecBUser)&"&membList[1].venderno="&B_COMPANY_NO
				strParam = strParam&"&usertagList[0].tag_nm=JUNGSAN_DATE&usertagList[0].tag_vl="&server.URLEncode(DEFAULT_JUNGSANDATE)
				'On Error Resume Next

				Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
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
							strErrMsg = getLgEcErrMessage(con_info)
						end if

					Set jsResult = Nothing
				End If

				Set objXML = Nothing

				'On Error Goto 0
				if ecCtrSeq ="" or ecCtrSeq = 0 Then
					Response.Write "<script>alert(""전자계약서 생성에 문제가 발생했습니다. 입력값 확인 후 다시 등록해주세요 - " & strErrMsg & """);location.href=""" & refer & """;</script>"
					Response.END
				end if
			end if

			''기본계약서여부(db상에 subType이 0일 경우 기본계약서임)
			isDefaultContract = (subType=0)

			sqlStr = " select * from db_partner.dbo.tbl_partner_ctr_master where 1=0"
			rsget.Open sqlStr,dbget,1,3
			rsget.AddNew
				rsget("groupid")            = groupid
				rsget("contractType")       = contractType
				rsget("makerid")            = CHKIIF(isDefaultContract,"",makerid) '' 기본계약서는 사업자 기준 makerid
				rsget("ctrState")           = 0  '' 수정중
				rsget("ctrNo")              = ""
				rsget("regUserID")          = session("ssBctID")
				rsget("ecCtrSeq")			= ecCtrSeq
				rsget("ecauser")			= ecauser
				rsget("ecbuser")			= ecbuser
				rsget("signType")			= getSignTypeCode(signtype) '' SignType Code 가져오기
			rsget.update
				ctrKey = rsget("ctrKey")
			rsget.close
			
			'If signtype = "3" Then
				'' 기본 계약서 DocuSign은 기본 폼값에 개인정보 수집 관련 계약서 값을 추가해준다.
				'contractContents = contractContents&getPriContractContentsDocuSign()
			'End If

			'' 기본 계약서 폼(contractContents)값을 기준으로 request 받은 값으로 치환해준다.
			For Each objItem In Request.Form
				if (Left(objItem,2)="$$") and (Right(objItem,2)="$$") then
					sqlStr = " insert into db_partner.dbo.tbl_partner_ctr_Detail"
					sqlStr = sqlStr & " (ctrKey, detailKey, detailValue)"
					sqlStr = sqlStr & " values("
					sqlStr = sqlStr & " " & ctrKey
					sqlStr = sqlStr & " ,'" & objItem & "'"
					sqlStr = sqlStr & " ,'" & Newhtml2db(Request.Form(objItem)) & "'"
					sqlStr = sqlStr & " )"
					dbget.Execute sqlStr
					if (objItem="$$CONTRACT_DATE$$") then
						bufStr  = Request.Form(objItem)
						bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일"
						'' DocuSign일경우 해당 일자 변형 하면 안됨
						If signtype <> "3" Then
							contractContents = Replace(contractContents,objItem,bufStr)
						End If
					elseif  (objItem="$$ENDDATE$$")   then
						enddate = Request.Form(objItem)
						bufStr  = enddate
						bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일"
						contractContents = Replace(contractContents,objItem,bufStr)
					else
						contractContents = Replace(contractContents,objItem,Request.Form(objItem))
					end if

					if (objItem="$$CONTRACT_DATE$$") then 
						ctrNo=Request.Form(objItem)
					End If
				end if
			Next

			ctrNo = TRim(replace(replace(ctrNo," ",""),"-",""))
			ctrNo = ctrNo & "-" & Format00(2,contractType) & "-" & ctrKey

			'' 치환한 값을 해당 파트너 정보에 업데이트 해준다.
			sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
			sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
			sqlStr = sqlStr & " ,ctrNo='" & ctrNo & "'"
			sqlStr = sqlStr & ", enddate='"&enddate&"'"
			sqlStr = sqlStr & " where ctrKey=" & ctrKey
			dbget.Execute sqlStr

		end if

		'--직매입계약서-----------------------------------------------------------------------------------------------------
		if chkCT12 = 1 then
			contractType = DEFAULT_CONTRACTTYPE_M
			''DocuSign 계약일 경우 새로운 계약서 번호를 사용
			If signtype = "3" Then
				contractType = DEFAULT_NEWCONTRACTTYPE_M '물품공급계약서(직매입)(2021.11)
			End If			

			''직매입계약서 기본폼을 불러온다.
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

			''LG U+ 전자결재
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

				Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
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
						strErrMsg = getLgEcErrMessage(con_info)
					end if

					Set jsResult = Nothing
				End If
				Set objXML = Nothing

				'On Error Goto 0
				if ecCtrSeq ="" or ecCtrSeq = 0 Then
					Response.Write "<script>alert(""전자계약서 생성에 문제가 발생했습니다. 입력값 확인 후 다시 등록해주세요 - " & strErrMsg & """);location.href=""" & refer & """;</script>"
					Response.END
				end if
			end if

			''기본계약서여부(직매입 계약서는 기본계약서로 분류)
			isDefaultContract = (subType=0)

			sqlStr = " select * from db_partner.dbo.tbl_partner_ctr_master where 1=0"
			rsget.Open sqlStr,dbget,1,3
				rsget.AddNew
				rsget("groupid")            = groupid
				rsget("contractType")       = contractType
				rsget("makerid")            = CHKIIF(isDefaultContract,"",makerid) '' 기본계약서는 사업자 기준 makerid
				rsget("ctrState")           = 0  '' 수정중
				rsget("ctrNo")              = ""
				rsget("regUserID")          = session("ssBctID")
				rsget("ecCtrSeq")			= ecCtrSeq
				rsget("ecauser")			= ecauser
				rsget("ecbuser")			= ecbuser
				rsget("signType")			= getSignTypeCode(signtype) '' SignType Code 가져오기				
			rsget.update
				ctrKey = rsget("ctrKey")
			rsget.close

			'' 직매입 계약서 DocuSign은 기본 폼값에 개인정보 수집 관련 계약서 값을 추가해준다.
			'If signtype = "3" Then
			'	contractContents = contractContents&getPriContractContentsDocuSign()
			'End If
			'' 직매입 계약서 폼(contractContents)값을 기준으로 request 받은 값으로 치환해준다.			
			For Each objItem In Request.Form
				if (Left(objItem,2)="$$") and (Right(objItem,2)="$$") then
					sqlStr = " insert into db_partner.dbo.tbl_partner_ctr_Detail"
					sqlStr = sqlStr & " (ctrKey, detailKey, detailValue)"
					sqlStr = sqlStr & " values("
					sqlStr = sqlStr & " " & ctrKey
					sqlStr = sqlStr & " ,'" & objItem & "'"
					sqlStr = sqlStr & " ,'" & Newhtml2db(Request.Form(objItem)) & "'"
					sqlStr = sqlStr & " )"
					dbget.Execute sqlStr

					if (objItem="$$CONTRACT_DATE$$") then
						bufStr  = Request.Form(objItem)
						bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일"
						'' DocuSign일경우 해당 일자 변형 하면 안됨
						If signtype <> "3" Then						
							contractContents = Replace(contractContents,objItem,bufStr)
						End If
					elseif  (objItem="$$ENDDATE$$")   then
						enddate = Request.Form(objItem)
						bufStr  = enddate
						bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일"
						contractContents = Replace(contractContents,objItem,bufStr)
					else
						contractContents = Replace(contractContents,objItem,Request.Form(objItem))
					end if

					if (objItem="$$CONTRACT_DATE$$") then 
						ctrNo=Request.Form(objItem)
					End If
				end if
			Next

			ctrNo = TRim(replace(replace(ctrNo," ",""),"-",""))
			ctrNo = ctrNo & "-" & Format00(2,contractType) & "-" & ctrKey

			'' 치환한 값을 해당 파트너 정보에 업데이트 해준다.
			sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
			sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
			sqlStr = sqlStr & " ,ctrNo='" & ctrNo & "'"
			sqlStr = sqlStr & ", enddate='"&enddate&"'"
			sqlStr = sqlStr & " where ctrKey=" & ctrKey
			dbget.Execute sqlStr
		end if
		'//------------------------------------------------------------------------------------------------------------------------

		'--특약 계약서-----------------------------------------------------------------------------------------------------
		if chkCT13 = 1 then
			contractType = SPECIALAPPOINTMENTCONTRACTTYPE '특약계약서(2022.02)

			''특약계약서 기본폼을 불러온다.
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

			''기본계약서여부(특약 계약서는 기본계약서로 분류)
			isDefaultContract = (subType=0)

			sqlStr = " select * from db_partner.dbo.tbl_partner_ctr_master where 1=0"
			rsget.Open sqlStr,dbget,1,3
				rsget.AddNew
				rsget("groupid")            = groupid
				rsget("contractType")       = contractType
				rsget("makerid")            = CHKIIF(isDefaultContract,"",makerid) '' 기본계약서는 사업자 기준 makerid
				rsget("ctrState")           = 0  '' 수정중
				rsget("ctrNo")              = ""
				rsget("regUserID")          = session("ssBctID")
				rsget("ecCtrSeq")			= ecCtrSeq
				rsget("ecauser")			= ecauser
				rsget("ecbuser")			= ecbuser
				rsget("signType")			= getSignTypeCode(signtype) '' SignType Code 가져오기				
			rsget.update
				ctrKey = rsget("ctrKey")
			rsget.close

			'' 특약 계약서 폼(contractContents)값을 기준으로 request 받은 값으로 치환해준다.			
			For Each objItem In Request.Form
				if (Left(objItem,2)="$$") and (Right(objItem,2)="$$") then
					sqlStr = " insert into db_partner.dbo.tbl_partner_ctr_Detail"
					sqlStr = sqlStr & " (ctrKey, detailKey, detailValue)"
					sqlStr = sqlStr & " values("
					sqlStr = sqlStr & " " & ctrKey
					sqlStr = sqlStr & " ,'" & objItem & "'"
					sqlStr = sqlStr & " ,'" & Newhtml2db(Request.Form(objItem)) & "'"
					sqlStr = sqlStr & " )"
					dbget.Execute sqlStr

					if (objItem="$$CONTRACT_DATE$$") then
						bufStr  = Request.Form(objItem)
						bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일"
						'' DocuSign일경우 해당 일자 변형 하면 안됨
						If signtype <> "3" Then						
							contractContents = Replace(contractContents,objItem,bufStr)
						End If
					elseif  (objItem="$$ENDDATE$$")   then
						enddate = Request.Form(objItem)
						bufStr  = enddate
						bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일"
						contractContents = Replace(contractContents,objItem,bufStr)
					else
						contractContents = Replace(contractContents,objItem,Replace(Request.Form(objItem), Chr(13)&Chr(10), "<br>"))
					end if

					if (objItem="$$CONTRACT_DATE$$") then 
						ctrNo=Request.Form(objItem)
					End If
				end if
			Next

			ctrNo = TRim(replace(replace(ctrNo," ",""),"-",""))
			ctrNo = ctrNo & "-" & Format00(2,contractType) & "-" & ctrKey

			'' 치환한 값을 해당 파트너 정보에 업데이트 해준다.
			sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
			sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
			sqlStr = sqlStr & " ,ctrNo='" & ctrNo & "'"
			sqlStr = sqlStr & ", enddate='"&enddate&"'"
			sqlStr = sqlStr & " where ctrKey=" & ctrKey
			dbget.Execute sqlStr
		end if
		'//------------------------------------------------------------------------------------------------------------------------

		''공인전자서명 면제 요청서 추가--------------------------------------------------------------------------------------------
		''해당 요청서는 거래기본계약서 또는 물품공급계약서(직매입)를 작성할때 해당 groupid에 공인전자서명 면제 요청서가 없으면 생성해준다.
		'' 공인 전자 서명 폼 그대로 사용하는지? 확인 면제 요청 사유가 디폴트로 있어야 될 것 같음(체크할것)
		'' 공인 전자 서명 폼이 미완성이라 당분간은 사용안함 2022.02.14
		If FALSE Then
			''현재 groupid 기준 해당 요청서가 있는지 확인
			If (chkCT11 = 1 Or chkCT12 = 1) then
				''2022년 1월 25일 기준 DocuSign에만 해당 요청서 추가		
				If signtype = "3" Then
					authexType = AUTHEX_NEWCONTRACTTYPE '공인전자서명 면제 요청서(2021.11)

					''groupid 기준으로 공인전자서명 면제 요청서가 있는지 확인
					sqlStr = "select count(ctrKey) as cnt "+vbcrlf
					sqlStr = sqlStr & " FROM db_partner.dbo.tbl_partner_ctr_master  "+vbcrlf
					sqlStr = sqlStr & " WHERE groupid='"&groupid&"' AND contractType='"&authexType&"' And ctrState >= 0 "
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
						authexCount = rsget("cnt")
					rsget.close

					''공인전자서명 면제 요청서가 없으면
					If authexCount < 1 Then
						''공인전자서명 기본폼을 불러온다.
						sqlStr = "select contractContents, contractName ,onoffgubun, subType" +vbcrlf
						sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType" +vbcrlf
						sqlStr = sqlStr & " where contractType=" & authexType
						rsget.Open sqlStr,dbget,1
						if Not rsget.Eof then
							contractContents = db2Html(rsget("contractContents"))
							contractName = db2Html(rsget("contractName"))
							onoffgubun = rsget("onoffgubun")
							subType    = rsget("subType")
						end if
						rsget.Close

						''기본계약서여부(공인전자서명 면제 요청서는 브랜드 아이디 기반이 아니므로 기본 계약서로 분류)
						isDefaultContract = (subType=0)

						sqlStr = " select * from db_partner.dbo.tbl_partner_ctr_master where 1=0"
						rsget.Open sqlStr,dbget,1,3
							rsget.AddNew
							rsget("groupid")            = groupid
							rsget("contractType")       = authexType
							rsget("makerid")            = CHKIIF(isDefaultContract,"",makerid) '' 기본계약서는 사업자 기준 makerid
							rsget("ctrState")           = 0  '' 수정중
							rsget("ctrNo")              = ""
							rsget("regUserID")          = session("ssBctID")
							rsget("ecCtrSeq")			= ecCtrSeq
							rsget("ecauser")			= ecauser
							rsget("ecbuser")			= ecbuser
							rsget("signType")			= getSignTypeCode(signtype) '' SignType Code 가져오기						
						rsget.update
							ctrKey = rsget("ctrKey")
						rsget.close

						'' 기본 계약서 폼(contractContents)값을 기준으로 request 받은 값으로 치환해준다.
						For Each objItem In Request.Form
							if (Left(objItem,2)="$$") and (Right(objItem,2)="$$") then
								sqlStr = " insert into db_partner.dbo.tbl_partner_ctr_Detail"
								sqlStr = sqlStr & " (ctrKey, detailKey, detailValue)"
								sqlStr = sqlStr & " values("
								sqlStr = sqlStr & " " & ctrKey
								sqlStr = sqlStr & " ,'" & objItem & "'"
								sqlStr = sqlStr & " ,'" & Newhtml2db(Request.Form(objItem)) & "'"
								sqlStr = sqlStr & " )"
								dbget.Execute sqlStr

								if (objItem="$$CONTRACT_DATE$$") then
									bufStr  = Request.Form(objItem)
									bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일"
									If signType <> "3" Then
										contractContents = Replace(contractContents,objItem,bufStr)
									End If
								elseif  (objItem="$$ENDDATE$$")   then
									enddate = Request.Form(objItem)
									bufStr  = enddate
									bufStr  = Left(bufStr,4) & "년 " & Mid(bufStr,6,2) & "월 " & Mid(bufStr,9,2) & "일"
									contractContents = Replace(contractContents,objItem,bufStr)
								else
									contractContents = Replace(contractContents,objItem,Request.Form(objItem))
								end if

								if (objItem="$$CONTRACT_DATE$$") then 
									ctrNo=Request.Form(objItem)
								End If
							end if
						Next

						ctrNo = TRim(replace(replace(ctrNo," ",""),"-",""))
						ctrNo = ctrNo & "-" & Format00(2,contractType) & "-" & ctrKey

						'' 치환한 값을 해당 파트너 정보에 업데이트 해준다.
						sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
						sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
						sqlStr = sqlStr & " ,ctrNo='" & ctrNo & "'"
						sqlStr = sqlStr & ", enddate='"&enddate&"'"
						sqlStr = sqlStr & " where ctrKey=" & ctrKey
						dbget.Execute sqlStr					
					End If
				End If
			End If
		End If
		
		'' 부속합의서 값이 있을경우
		if (addmwdiv<>"") then

			SET ogroupInfo = new CPartnerGroup
			ogroupInfo.FRectGroupid = groupid
			if (groupid<>"") then
				ogroupInfo.GetOneGroupInfo
			end if

			if (ogroupInfo.FResultCount<1) then
				SET ogroupInfo = Nothing
				dbget.close()
				response.write "그룹정보가 없습니다."
				response.end
			end if

			if (addOF_ctrDate<>"") and (addON_ctrDate="") then
				addON_ctrDate = addOF_ctrDate
			end if

			''부속 합의서 등록
			'' 매입계약서인지 체크
			For kk = 1 To Request.Form("addmwdiv").Count
				chkmwdiv     = Request.Form("addmwdiv")(kk)
				addmwdiv        = Request.Form("addmwdiv")(kk)
				addsellplace    = Request.Form("addsellplace")(kk)
				addmargin       = Request.Form("addmargin")(kk)

				'//LG U+ 전자계약
				if signtype="2" then
					dim defmargin, defdeliver	,ismeaip
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
						elseif addmwdiv ="U"	 then
							mwName ="업체"
						elseif addmwdiv ="W"	 then
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
					strParam = strParam&"&usertagList[1].tag_nm=VAL_MAKERID"&"&usertagList[1].tag_vl="&server.URLEncode(makerid)
					strParam = strParam&"&usertagList[2].tag_nm=VAL_SELLPLACE"&"&usertagList[2].tag_vl="&server.URLEncode(sellplaceName)
					strParam = strParam&"&usertagList[3].tag_nm=VAL_MWDIV"&"&usertagList[3].tag_vl="&server.URLEncode(mwName)
					strParam = strParam&"&usertagList[4].tag_nm=VAL_DEFMARGIN"&"&usertagList[4].tag_vl="&server.URLEncode(defmargin)
					strParam = strParam&"&usertagList[5].tag_nm=VAL_DEFDELIVER"&"&usertagList[5].tag_vl="&server.URLEncode(defdeliver)
					'On Error Resume Next

					Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
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
							strErrMsg = getLgEcErrMessage(con_info)
						end if

						Set jsResult = Nothing
					End If
					Set objXML = Nothing

					'On Error Goto 0
					if ecCtrSeq ="" or ecCtrSeq = 0 Then
						Response.Write "<script>alert(""전자계약서 생성에 문제가 발생했습니다. 입력값 확인 후 다시 등록해주세요 - " & strErrMsg & """);location.href=""" & refer & """;</script>"
						Response.END
					end if

					sqlStr = " select * from db_partner.dbo.tbl_partner_ctr_master where 1=0"
					rsget.Open sqlStr,dbget,1,3
						rsget.AddNew
						rsget("groupid")		= groupid
						rsget("contractType")	= contractType
						rsget("makerid")		= CHKIIF(isDefaultContract,"",makerid) '' 기본계약서는 사업자 기준 makerid
						rsget("ctrState")		= 0  '' 수정중
						rsget("ctrNo")			= ""
						rsget("regUserID")		= session("ssBctID")
						rsget("ecCtrSeq")		= ecCtrSeq
						rsget("ecauser")		= ecAUser
						rsget("ecbuser")		= ecBUser
						rsget("signType")		= getSignTypeCode(signtype) '' SignType Code 가져오기						
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
						sqlStr = sqlStr & " where ctrKey=" & ctrKey
						dbget.Execute sqlStr
					else
						response.write "계약서 작성실패"
					end if

				else
					'// 수기계약 Or DocuSign 처리
					'if ((Not chkmwdivMExists) and ((chkmwdiv="M") or (chkmwdiv="B031"))) or ((Not chkmwdivWExists) and NOT ((chkmwdiv="M") or (chkmwdiv="B031"))) then
					if ((Not chkmwdivMExists) and ( chkmwdiv="M")  ) or ((Not chkmwdivWExists) and NOT  (chkmwdiv="M")   ) then
						if (chkmwdiv="M")   then '' 매입/ 출고매입
							contractType = ADD_CONTRACTTYPE_M
							''DocuSign 계약일 경우 새로운 계약서 번호를 사용
							If signtype = "3" Then
								contractType = ADD_NEWCONTRACTTYPE_M ''물품공급계약(직매입) 부속합의서(2021.11)
							End If
							chkmwdivMExists = true
							ectypeSeq = Fec_addctrtype_M
						else
							contractType = ADD_CONTRACTTYPE
							''DocuSign 계약일 경우 새로운 계약서 번호를 사용
							If signtype = "3" Then
								contractType = ADD_NEWCONTRACTTYPE ''거래기본계약 부속합의서(2021.11)
							End If							
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
							rsget("groupid")			= groupid
							rsget("contractType")		= contractType
							rsget("makerid")			= CHKIIF(isDefaultContract,"",makerid) '' 기본계약서는 사업자 기준 makerid
							rsget("ctrState")			= 0  '' 수정중
							rsget("ctrNo")				= ""
							rsget("regUserID")			= session("ssBctID")
							rsget("ecCtrSeq")			= ecCtrSeq
							rsget("ecauser")			= ecAUser
							rsget("ecbuser")			= ecBuser
							rsget("signType")			= getSignTypeCode(signtype) '' SignType Code 가져오기							
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

						''-----------------
						For i = 1 To Request.Form("addmwdiv").Count
							addmwdiv        = Request.Form("addmwdiv")(i)
							if ((chkmwdiv="M" or chkmwdiv="B031") and (addmwdiv="M" or addmwdiv="B031")) or ((chkmwdiv<>"M" and chkmwdiv<>"B031") and (addmwdiv<>"M" and addmwdiv<>"B031")) then

								addsellplace    = Request.Form("addsellplace")(i)
								addmargin       = Request.Form("addmargin")(i)

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
						Next

						'' 계약서 DB 내용으로 치환
						if  (FillContractContentsByDB(ctrKey, contractContents)) then

							ctrNo = TRim(replace(replace(addON_ctrDate," ",""),"-",""))
							ctrNo = ctrNo & "-" & Format00(2,contractType) & "-" & ctrKey

							sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
							sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
							sqlStr = sqlStr & " ,ctrNo='" & ctrNo & "'"
							sqlStr = sqlStr & " ,enddate='"&addON_endDate&"'"
							sqlStr = sqlStr & " where ctrKey=" & ctrKey

							dbget.Execute sqlStr
						else
							response.write "계약서 작성실패"
						end if
						''--------------------------
					end if
				end if
			Next

			SET ogroupInfo = Nothing
		end if

		response.write "<script>alert('등록 되었습니다.\n\n확인 하신후 오픈하시기 바랍니다.');</script>"
		response.write "<script>location.href = '" & refer & "'</script>"
		dbget.close()	:	response.End

	elseif (mode="edt") then
		''수정 가능상태 Check
		sqlStr = "select contractType, ctrState from db_partner.dbo.tbl_partner_ctr_master"
		sqlStr = sqlStr & " where ctrKey=" & ctrKey

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			ctrState   = rsget("ctrState")
			contractType    = rsget("contractType")
			contractExists = (ctrState>=1)  ''오픈하면 수정 못함/ 삭제후 재작성.
		end if
		rsget.Close

		if (contractExists) then
			response.write "<script>alert('수정 가능 상태가 아닙니다.\n관리자 문의 요망.');history.back();</script>"
			dbget.close()	:	response.End
		end if


		sqlStr = "select t.contractContents, t.contractName from "
		sqlStr = sqlStr & " db_partner.dbo.tbl_partner_contractType t"
		sqlStr = sqlStr & "     Join db_partner.dbo.tbl_partner_ctr_master c"
		sqlStr = sqlStr & "     on c.contractType=t.contractType"
		sqlStr = sqlStr & " where c.ctrKey=" & ctrKey


		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			contractContents = db2Html(rsget("contractContents"))
			contractName = db2Html(rsget("contractName"))
		end if
		rsget.Close

	'    sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
	'    sqlStr = sqlStr & " set contractEtcContetns='" & Newhtml2db(contractEtcContetns) & "'"
	'    sqlStr = sqlStr & " where contractID=" & contractID
	'
	'    dbget.Execute sqlStr

		For Each objItem In Request.Form
			''response.write objItem & "," & Request.Form(objItem) & "<br>"
			if (Left(objItem,2)="$$") and (Right(objItem,2)="$$") then
				sqlStr = " update db_partner.dbo.tbl_partner_ctr_Detail"
				sqlStr = sqlStr & " set detailValue='" & Newhtml2db(Request.Form(objItem)) & "'"
				sqlStr = sqlStr & " where ctrKey=" & ctrKey
				sqlStr = sqlStr & " and detailKey='" & objItem & "'"

				dbget.Execute sqlStr

				if (objItem="$$CONTRACT_DATE$$") then
					bufStr  = Request.Form(objItem)
					bufStr  = Left(bufStr,4) & "년" & Mid(bufStr,6,2) & "월" & Mid(bufStr,9,2) & "일"
					contractContents = Replace(contractContents,objItem,bufStr)
				else
					contractContents = Replace(contractContents,objItem,Request.Form(objItem))
				end if

				if (objItem="$$CONTRACT_DATE$$") then ctrNo=Request.Form(objItem)
			end if
		Next

		For i = 1 To Request.Form("addmwdiv").Count
			addmwdiv        = Request.Form("addmwdiv")(i)

			addsellplace    = Request.Form("addsellplace")(i)
			addmargin       = Request.Form("addmargin")(i)
			ctrSubKey       = Request.Form("ctrSubKey")(i)

			sqlStr = " update db_partner.dbo.tbl_partner_ctr_Sub"
			sqlStr = sqlStr & " set sellplace='"&addsellplace&"'"
			sqlStr = sqlStr & " ,mwdiv='"&addmwdiv&"'"
			sqlStr = sqlStr & " ,defaultmargin='"&addmargin&"'"
			sqlStr = sqlStr & " where ctrKey="&ctrKey
			sqlStr = sqlStr & " and ctrSubKey="&ctrSubKey
	''rw sqlStr
			dbget.Execute sqlStr
		Next

		'' 계약서 DB 내용으로 치환
		if  (FillContractContentsByDB(ctrKey, contractContents)) then
			''계약서 번호 생성. YYYYMMDD(계약일)-contractType-contractID
			ctrNo = TRim(replace(replace(ctrNo," ",""),"-",""))
			ctrNo = ctrNo & "-" & Format00(2,contractType) & "-" & ctrKey
			contractContents = Replace(contractContents,"$$CONTRACT_NO$$",ctrNo)

			sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
			sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
			sqlStr = sqlStr & " ,ctrNo='" & ctrNo & "'"
			if (ctrState=-2) then
				sqlStr = sqlStr & " ,ctrState=0"
			end if
			sqlStr = sqlStr & " where ctrKey=" & ctrKey


			dbget.Execute sqlStr
		else
			response.write "계약서 작성실패"
		end if

		response.write "<script>alert('수정 되었습니다.\n\n확인 하신후 재발송 하시기 바랍니다.');</script>"
		response.write "<script>opener.location.reload(); location.replace('" & refer & "');</script>"
		dbget.close()	:	response.End

	elseif (mode="del") then
		''수정 가능상태 Check
		sqlStr = "select contractType, ctrState from db_partner.dbo.tbl_partner_ctr_master"
		sqlStr = sqlStr & " where ctrKey=" & ctrKey

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			ctrState   = rsget("ctrState")
			contractType    = rsget("contractType")
			contractExists = (ctrState>=3)
		end if
		rsget.Close

	' 임시삭제
	'    if Not C_ADMIN_AUTH then
	'        if (contractExists) then
	'            response.write "<script>alert('삭제 가능 상태가 아닙니다.\n관리자 문의 요망.');history.back();</script>"
	'            dbget.close()	:	response.End
	'        end if
	'    End if


		sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
		sqlStr = sqlStr & " set ctrState=-1"
		sqlStr = sqlStr & " ,finUserID='"&session("ssBctID")&"'" ''최종처리
		sqlStr = sqlStr & " ,deleteDate=getdate()"
		sqlStr = sqlStr & " where ctrKey=" & ctrKey
	'  if Not C_ADMIN_AUTH then
	'     sqlStr = sqlStr & " and ctrState<3"
	'  end if

		dbget.Execute sqlStr,AssignedRow

		if (AssignedRow>0) then
			response.write "<script>alert('삭제 되었습니다.');</script>"
			response.write "<script>opener.location.reload(); window.close();</script>"
			dbget.close()	:	response.End
		else
			response.write "<script>alert('수정 중 오류가 발생 하였습니다.');</script>"
		end if
		response.write "<script>opener.location.reload(); location.replace('" & refer & "');</script>"
		dbget.close()	:	response.End

	elseif (mode="fin") then
		''수정 가능상태 Check
		sqlStr = "select contractType, ctrState from db_partner.dbo.tbl_partner_ctr_master"
		sqlStr = sqlStr & " where ctrKey=" & ctrKey

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			ctrState   = rsget("ctrState")
			contractType    = rsget("contractType")
			contractExists = (ctrState<1)
		end if
		rsget.Close

		if (contractExists) then
			response.write "<script>alert('완료 가능 상태가 아닙니다.\n관리자 문의 요망.');history.back();</script>"
			dbget.close()	:	response.End
		end if


		sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
		sqlStr = sqlStr & " set ctrState=7"
		sqlStr = sqlStr & " ,finUserID='"&session("ssBctID")&"'"
		sqlStr = sqlStr & " ,finishDate=getdate()"
		sqlStr = sqlStr & " where ctrKey=" & ctrKey
		sqlStr = sqlStr & " and ctrState<7 and ctrState>=1"

		dbget.Execute sqlStr,AssignedRow

		if (AssignedRow>0) then
			response.write "<script>alert('계약 완료 되었습니다.');</script>"
			response.write "<script>opener.location.reload(); window.close();</script>"
			dbget.close()	:	response.End
		else
			response.write "<script>alert('수정 중 오류가 발생 하였습니다.');</script>"
		end if
		response.write "<script>opener.location.reload(); location.replace('" & refer & "');</script>"
		dbget.close()	:	response.End

	elseif (mode="state0") then
		''수정 가능상태 Check
		sqlStr = "select contractType, ctrState from db_partner.dbo.tbl_partner_ctr_master"
		sqlStr = sqlStr & " where ctrKey=" & ctrKey

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			ctrState   = rsget("ctrState")
			contractType    = rsget("contractType")
			contractExists = (ctrState>3)
		end if
		rsget.Close

		if (contractExists) then
			response.write "<script>alert('수정 가능 상태가 아닙니다.\n관리자 문의 요망.');history.back();</script>"
			dbget.close()	:	response.End
		end if


		sqlStr = " update db_partner.dbo.tbl_partner_ctr_master"
		sqlStr = sqlStr & " set ctrState=0"
		sqlStr = sqlStr & " where ctrKey=" & ctrKey
		sqlStr = sqlStr & " and ctrState<7 and ctrState>=1"

		dbget.Execute sqlStr,AssignedRow

		if (AssignedRow>0) then
			response.write "<script>alert('수정 완료 되었습니다.');</script>"
			response.write "<script>opener.location.reload(); window.close();</script>"
			dbget.close()	:	response.End
		else
			response.write "<script>alert('수정 중 오류가 발생 하였습니다.');</script>"
		end if
		response.write "<script>opener.location.reload(); location.replace('" & refer & "');</script>"
		dbget.close()	:	response.End

	elseif (mode="ctropen") then
	'    if (session("ssBctID")<>"icommang") then
	'        response.write "<script>alert('현재 오픈할 수 없습니다.- 관리자 문의요망');</script>"
	'        dbget.close()	:	response.End
	'    end if

		''이메일 체크
		if (ckEmail<>"") and (mngEmail<>"") then
			if (mngEmail="") or (InStr(mngEmail,"@")<0) or (Len(mngEmail)<8) then
				response.write "<script>alert('업체 담당자 Email 주소가 유효하지 않습니다.');</script>"
				response.write "<script>location.replace('" & refer & "');</script>"
				dbget.close() : response.End
			end if

			sqlStr = "select IsNULL(p.usermail,'') as email from db_partner.dbo.tbl_user_tenbyten p"
			sqlStr = sqlStr & " where p.userid='"&session("ssBctID")&"'"
			sqlStr = sqlStr & " and p.userid<>''"

			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				mailfrom = db2Html(rsget("email"))
			end if
			rsget.Close

			mailfrom = Trim(mailfrom)

			if (mailfrom="") or (InStr(mailfrom,"@")<0) or (Len(mailfrom)<8) then
				response.write "<script>alert('발송자 Email  주소가 유효하지 않습니다.마이 정보에서 Email 수정 후 사용하시기 바랍니다.(등록된 이메일주소:"&mailfrom&")');</script>"
				response.write "<script>location.replace('" & refer & "');</script>"
				dbget.close()	:	response.End
			end if
		end if

		dim con_error
		cnt = Request.Form("chkCtr").Count

		set oneContract = new CPartnerContract
			oneContract.fnGetContractToken
			acctoken = oneContract.Facctoken
			reftoken = oneContract.Freftoken
		set oneContract = nothing

		for i=1 to cnt
			chkCtr = Request.Form("chkCtr")(i)

			sqlStr = "select  m.ecBUser, g.company_no , m.ecCtrSeq"
				sqlStr =  sqlStr & " from db_partner.dbo.tbl_partner_ctr_master as m "
				sqlStr =  sqlStr & " inner join db_partner.dbo.tbl_partner_group G on m.groupid = g.groupid"
				sqlStr = sqlStr & " where ctrKey ="&chkCtr
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				ecBUser= rsget("ecBUser")
				B_COMPANY_NO= replace(rsget("company_no"),"-","")
				ecCtrSeq = rsget("ecCtrSeq")
			end if
			rsget.Close

			if ecCtrSeq <>"0" and not isNull(ecCtrSeq) then  '전자계약일 경우
				con_status =  fnCheckCont(ecCtrSeq,B_COMPANY_NO,ecBUser,acctoken)
				if 	Fchkerror ="invalid_token"		then
						call sbGetRefToken(reftoken)
						acctoken = Faccess_token
						con_status =  fnCheckCont(ecCtrSeq,B_COMPANY_NO,ecBUser,acctoken)
				end if

				if con_status<> "succ" Then
					Response.write "<script type='text/javascript' language='javascript'>alert('전자계약서 오픈에 문제가 발생했습니다. 입력값 확인 후 다시 처리해주세요 - "&FErrMsg&"');location.href = '"&refer&"';</script>"
					response.end
				end if
			end if

			sqlstr = " update db_partner.dbo.tbl_partner_ctr_master"&VbCRLF
			sqlstr = sqlstr & " set ctrState=1"&VbCRLF                              ''업체 오픈
			sqlstr = sqlstr & " ,sendUserID='"&session("ssBctID")&"'"&VbCRLF
			sqlstr = sqlstr & " ,sendDate=getdate()"
			sqlstr = sqlstr & " where ctrKey="&chkCtr&VbCRLF
			sqlstr = sqlstr & " and ctrState=0"&VbCRLF ''수정중만 오픈가능

			dbget.Execute  sqlstr, AssignedRow

			if (AssignedRow>0) then ''and (noMatCnt>0)
				''마진변경
				sqlstr = " exec db_partner.[dbo].[sp_Ten_partner_AddContract_MarginUpdateByKey] "&chkCtr&",'"&session("ssBctID")&"'"
			' rw sqlstr&"<BR>"
				dbget.Execute  sqlstr
			end if

			iCtrKeyArr = iCtrKeyArr&chkCtr&","
		next

		if Right(iCtrKeyArr,1)="," then iCtrKeyArr=Left(iCtrKeyArr,Len(iCtrKeyArr)-1)

		if (ckHp<>"") and (mngHp<>"") then
			'' SMS 발송
			''call SendNormalSMS(mngHp,"1644-6030","[텐바이텐] 신규 계약서가 발송되었습니다. email 또는 SCM 업체계약관리 메뉴 참조")
			call SendNormalSMS_LINK(mngHp,"1644-6030","[텐바이텐] 신규 계약서가 발송되었습니다. email 또는 SCM 업체계약관리 메뉴 참조")
		end if

		if (ckEmail<>"") and (mngEmail<>"") then
			'' 이메일 발송
			set ocontract = new CPartnerContract
			ocontract.FPageSize=50
			ocontract.FCurrPage = 1
			ocontract.FRectContractState = 1 ''오픈
			ocontract.FRectGroupID = groupid
			ocontract.FRectCtrKeyArr = iCtrKeyArr
			ocontract.GetNewContractList

			set oMdInfoList = new CPartnerContract
			oMdInfoList.FRectGroupID = groupid
			oMdInfoList.FRectContractState = 1 ''오픈
			oMdInfoList.FRectCtrKeyArr = iCtrKeyArr
			oMdInfoList.getContractEmailMdList(FALSE)

			mailtitle       = "[텐바이텐] 신규 계약서가 발송 되었습니다."
			if signtype="2" then
				mailcontent   = makeEcCtrMailContents(ocontract,oMdInfoList,False,manageUrl)
			else
				mailcontent   = makeCtrMailContents(ocontract,oMdInfoList,False)
			end if

			Call SendMail(mailfrom, mngEmail, mailtitle, mailcontent)

			set ocontract=nothing
			set oMdInfoList=nothing
		end if

		if (application("Svr_Info")	= "Dev") then
			response.write mailcontent
		else
			response.write "<script>alert('계약서가 발송되었습니다.');</script>"
			response.write "<script>opener.location.reload(); window.close();</script>"
			dbget.close()	:	response.End
		end if

	elseif (mode="ctropendocusign") then
		'' 도큐사인 발송
	'    if (session("ssBctID")<>"icommang") then
	'        response.write "<script>alert('현재 오픈할 수 없습니다.- 관리자 문의요망');</script>"
	'        dbget.close()	:	response.End
	'    end if

		''이메일 체크
		if (ckEmail<>"") and (mngEmail<>"") then
			if (mngEmail="") or (InStr(mngEmail,"@")<0) or (Len(mngEmail)<8) then
				response.write "<script>alert('업체 담당자 Email 주소가 유효하지 않습니다.');</script>"
				response.write "<script>location.replace('" & refer & "');</script>"
				dbget.close() : response.End
			end if

			sqlStr = "select IsNULL(p.usermail,'') as email from db_partner.dbo.tbl_user_tenbyten p"
			sqlStr = sqlStr & " where p.userid='"&session("ssBctID")&"'"
			sqlStr = sqlStr & " and p.userid<>''"

			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				mailfrom = db2Html(rsget("email"))
			end if
			rsget.Close

			mailfrom = Trim(mailfrom)

			if (mailfrom="") or (InStr(mailfrom,"@")<0) or (Len(mailfrom)<8) then
				response.write "<script>alert('발송자 Email  주소가 유효하지 않습니다.마이 정보에서 Email 수정 후 사용하시기 바랍니다.(등록된 이메일주소:"&mailfrom&")');</script>"
				response.write "<script>location.replace('" & refer & "');</script>"
				dbget.close()	:	response.End
			end if
		end if

		If Trim(mngName="") Then
			response.write "<script>alert('수신자 이름이 유효하지 않습니다.업체 정보를 확인해주세요.');</script>"
			response.write "<script>location.replace('" & refer & "');</script>"
			dbget.close()	:	response.End
		End If

		cnt = Request.Form("chkCtr").Count
		docuSignContractData = ""
		for i=1 to cnt
			chkCtr = Request.Form("chkCtr")(i)

			sqlStr = "select  m.*, c.contractName "
				sqlStr =  sqlStr & " from db_partner.dbo.tbl_partner_ctr_master as m "
				sqlStr =  sqlStr & " inner join db_partner.dbo.tbl_partner_contracttype c on m.contractType = c.contractType"
				sqlStr = sqlStr & " where m.ctrKey ="&chkCtr
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				contractContents = db2Html(rsget("contractContents"))
				contractName = db2Html(rsget("contractName"))
			end if
			rsget.Close
			docuSignContractData = docuSignContractData&",{""documentName"": """&contractName&""",""html"": """&aspJsonStringEscape(contractContents)&"""}"

			'' ImageBase64 치환
			docuSignContractData = replace(docuSignContractData, "$$IMAGE1$$", DocuSignStampBase64)
		next

		docuSignContractData = Right(docuSignContractData,len(docuSignContractData)-1)
		docuSignSendContents = ""

		docuSignSendContents = "{"
		docuSignSendContents = docuSignSendContents & """body"": ""텐바이텐 기본 계약서입니다.\n내용 확인하시고 싸인해주세요.\n담당자의 확인이 필요한경우 "&Trim(mailfrom)&"로 이메일 주세요."","		
		docuSignSendContents = docuSignSendContents & """email"": """&Trim(mngEmail)&""","
		docuSignSendContents = docuSignSendContents & """htmlDocumentList"": ["&docuSignContractData&"],"
'		docuSignSendContents = docuSignSendContents & """imageList"": "
'		docuSignSendContents = docuSignSendContents & "[{"
'		docuSignSendContents = docuSignSendContents & """base64Image"": """&DocuSignStampBase64&""","
'		docuSignSendContents = docuSignSendContents & """pattern"": ""$$IMAGE1$$"""
'		docuSignSendContents = docuSignSendContents & "}],"
		docuSignSendContents = docuSignSendContents & """name"": """&Trim(mngName)&""","
		docuSignSendContents = docuSignSendContents & """signDatePattern"": ""$$CONTRACT_DATE$$"","
		docuSignSendContents = docuSignSendContents & """signPattern"": ""$$SIGN_PATTERN$$"","
		docuSignSendContents = docuSignSendContents & """subject"": ""텐바이텐 계약서"""
		docuSignSendContents = docuSignSendContents & "}"

		'response.write docuSignSendContents
		'response.end


		Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
		Session.CodePage = 65001
		'Set objXML = CreateObject("Msxml2.ServerXMLHTTP")
		'objXML.SetTimeouts 40000, 40000, 40000, 40000
		objXML.Open "POST", FecDocuURL&"/api/contract/v1/docusign/htmlSign", False
		objXML.setRequestHeader "Content-Type", "application/json"
		if (application("Svr_Info")	<> "Dev") then
			objXML.SetRequestHeader "api-key-v1", ""+CStr(adminApiKey)+""
		End If
		objXML.Send docuSignSendContents
		Session.CodePage = 949
		iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")

		'response.write objXML.Status&"<br>"&iRbody
		
		If objXML.Status = "200" Then
			Set jsResult = JSON.parse(iRbody)
			docuSignEnvelopeId = jsResult.envelopeId
			docuSignStatus = jsResult.status
			docuSignStatusDateTime = jsResult.statusDateTime
			docuSignUri = jsResult.uri
			Set jsResult = Nothing
		Else
			response.write "<script>alert('DocuSign 통신중 오류가 발생하였습니다.\nErrorCode("&objXML.Status&")');</script>"
			response.write "<script>location.replace('" & refer & "');</script>"
			dbget.close() : response.End
		End If
		Set objXML = Nothing

		for i=1 to cnt
			chkCtr = Request.Form("chkCtr")(i)

			sqlstr = " update db_partner.dbo.tbl_partner_ctr_master"&VbCRLF
			sqlstr = sqlstr & " set ctrState=1"&VbCRLF                              ''업체 오픈
			sqlstr = sqlstr & " ,sendUserID='"&session("ssBctID")&"'"&VbCRLF
			sqlstr = sqlstr & " ,sendDate=getdate()"
			sqlstr = sqlstr & " ,docuSignId='"&docuSignEnvelopeId&"'"
			sqlstr = sqlstr & " ,docuSignUri='"&docuSignUri&"'"
			sqlstr = sqlstr & " ,docuSignSenddate='"&docuSignStatusDateTime&"'"									
			sqlstr = sqlstr & " where ctrKey="&chkCtr&VbCRLF
			sqlstr = sqlstr & " and ctrState=0"&VbCRLF ''수정중만 오픈가능
			dbget.Execute  sqlstr, AssignedRow

			if (AssignedRow>0) then ''and (noMatCnt>0)
				''마진변경
				sqlstr = " exec db_partner.[dbo].[sp_Ten_partner_AddContract_MarginUpdateByKey] "&chkCtr&",'"&session("ssBctID")&"'"
			' rw sqlstr&"<BR>"
				dbget.Execute  sqlstr
			end if
			
			iCtrKeyArr = iCtrKeyArr&chkCtr&","		
		next

		if Right(iCtrKeyArr,1)="," then iCtrKeyArr=Left(iCtrKeyArr,Len(iCtrKeyArr)-1)

		if (ckHp<>"") and (mngHp<>"") then
			'' SMS 발송
			''call SendNormalSMS(mngHp,"1644-6030","[텐바이텐] 신규 계약서가 발송되었습니다. email 또는 SCM 업체계약관리 메뉴 참조")
			call SendNormalSMS_LINK(mngHp,"1644-6030","[텐바이텐] 신규 계약서가 발송되었습니다. email 참조요망.")
		end if

		'' DocuSign에서 보내는 메일을 사용하기 때문에 우리쪽에서 메일은 보내지 않음
		'if (ckEmail<>"") and (mngEmail<>"") then
			'' 이메일 발송
		'	set ocontract = new CPartnerContract
		'	ocontract.FPageSize=50
		'	ocontract.FCurrPage = 1
		'	ocontract.FRectContractState = 1 ''오픈
		'	ocontract.FRectGroupID = groupid
		'	ocontract.FRectCtrKeyArr = iCtrKeyArr
		'	ocontract.GetNewContractList

		'	set oMdInfoList = new CPartnerContract
		'	oMdInfoList.FRectGroupID = groupid
		'	oMdInfoList.FRectContractState = 1 ''오픈
		'	oMdInfoList.FRectCtrKeyArr = iCtrKeyArr
		'	oMdInfoList.getContractEmailMdList(FALSE)

		'	mailtitle       = "[텐바이텐] 신규 계약서가 발송 되었습니다."
		'	if signtype="2" then
		'		mailcontent   = makeEcCtrMailContents(ocontract,oMdInfoList,False,manageUrl)
		'	else
		'		mailcontent   = makeCtrMailContents(ocontract,oMdInfoList,False)
		'	end if

		'	Call SendMail(mailfrom, mngEmail, mailtitle, mailcontent)

		'	set ocontract=nothing
		'	set oMdInfoList=nothing
		'end if

		response.write "<script>alert('계약서가 발송되었습니다.');</script>"
		response.write "<script>opener.location.reload(); window.close();</script>"
		dbget.close()	:	response.End

	elseif (mode="rjtCtr") then
		sqlStr = " if NOT Exists(select * from db_partner.dbo.tbl_partner_ctr_Hold where makerid='"&makerid&"' and onoffgbn='"&addsellplace&"')" & VbCRLF
		sqlStr = sqlStr & " BEGIN"& VbCRLF
		sqlStr = sqlStr & "     insert into db_partner.dbo.tbl_partner_ctr_Hold"
		sqlStr = sqlStr & "     (makerid,onoffgbn,holdregid)"
		sqlStr = sqlStr & "     values('"&makerid&"','"&addsellplace&"','"&session("ssBctID")&"')"
		sqlStr = sqlStr & " END" & VbCRLF

		dbget.Execute sqlStr,AssignedRow

		if (AssignedRow>0) then
			response.write "<script>alert('계약 보류 브랜드로 지정되었습니다.');</script>"
			response.write "<script>location.replace('" & refer & "');</script>"
			dbget.close()	:	response.End
		else
			response.write "<script>alert('계약 보류 지정 중 오류가 발생 하였습니다.');</script>"
		end if

	elseif (mode="rjtCtrDel") then
		sqlStr = " delete from db_partner.dbo.tbl_partner_ctr_Hold "
		sqlStr = sqlStr & " where makerid='"&makerid&"'"
		sqlStr = sqlStr & " and onoffgbn='"&addsellplace&"'"
		''sqlStr = sqlStr & " and holdregid='"&session("ssBctID")&"'"

		dbget.Execute sqlStr,AssignedRow

		if (AssignedRow>0) then
			response.write "<script>alert('계약 보류 브랜드로 지정이 해제되었습니다.');</script>"
			response.write "<script>location.replace('" & refer & "');</script>"
			dbget.close()	:	response.End
		else
			response.write "<script>alert('계약 보류 해제 중 오류가 발생 하였습니다.');</script>"
		end if

	elseif (mode="ctrfin") then
		''계약완료 상태 변경

	elseif (mode ="ecstate") then
		dim  arrList, intLoop
		dim ecCtrState
		sqlStr = " select  m.ctrKey, ecctrseq, g.company_no, ecBUser , m.ctrstate "
		sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master as m with (nolock)"
		sqlStr = sqlStr & "	inner join db_partner.dbo.tbl_partner_group as g with (nolock) on m.groupid = g.groupid"
		sqlStr = sqlStr & "	where m.groupid = '"&groupid&"'"
		sqlStr = sqlStr & "	and m.ctrstate<>-1"		' 삭제는 제낀다
		sqlStr = sqlStr & "	 and ecCtrseq > 0 	"
		rsget.Open sqlStr,dbget,1
		if not rsget.eof Then
			arrList = rsget.getrows()
		end if
		rsget.close

		if isArray(arrList) Then
			'token 가져오기(db에서)
			set oneContract = new CPartnerContract
					oneContract.fnGetContractToken
					acctoken = oneContract.Facctoken
					reftoken = oneContract.Freftoken
			set oneContract = nothing

			'token이 없을때 token 발행
			if isNull(acctoken) then
				'call sbGetNewToken(ecId,ecPwd)
				'acctoken = Faccess_token
				'if acctoken = "" Then
				Response.write "<script type='text/javascript' language='javascript'>alert('전자계약 계정정보가 잘못입력되었습니다. 확인 후 다시 시도해주세요,');location.href = '"&refer&"';</script>"
				response.end
				'  end if
			end if

			for intLoop = 0 To uBound(arrList,2)
				ecCtrState =  fnViewEcCont(arrList(1,intLoop),replace(arrList(2,intLoop),"-",""),arrList(3,intLoop),acctoken)

				if Fchkerror ="invalid_token" then
						call sbGetRefToken(reftoken)
						acctoken = Faccess_token
						ecCtrState =  fnViewEcCont(arrList(1,intLoop),replace(arrList(2,intLoop),"-",""),arrList(3,intLoop),acctoken)
				end if

				if ecCtrState = "" then
					Response.write "<script type='text/javascript' language='javascript'>alert('전자계약상태 업데이트에 문제가 발생했습니다. 확인 후 다시 시도해주세요,');location.href = '"&refer&"';</script>"
					response.end
				end if

				sqlStr = "update db_partner.dbo.tbl_partner_ctr_master set ctrstate = "&GetContractEcState(ecCtrState)&", lastupdate =getdate()"
				sqlstr = sqlstr & " where ctrKey="&arrList(0,intLoop)&" and ctrstate <> " &GetContractEcState(ecCtrState)
				dbget.Execute  sqlstr, 1
			next
		end if
		Response.write "<script type='text/javascript' language='javascript'>alert('전자계약서 상태가 업데이트 되었습니다.');location.href = '"&refer&"';</script>"
		response.end

	elseif (mode ="docustate") then
		dim  docuStatusAdminCodeConversion, docuErrorStatusValue
		docuErrorStatusValue = ""
		sqlStr = " select  m.ctrKey, m.docuSignId, m.ctrstate "
		sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master as m with (nolock)"
		sqlStr = sqlStr & "	where m.groupid = '"&groupid&"'"
		sqlStr = sqlStr & "	and m.ctrstate<>-1"		' 삭제는 제낀다
		sqlStr = sqlStr & "	and ISNULL(m.docuSignId,'') <> '' "
		sqlStr = sqlStr & "	and signType='D' "
		rsget.Open sqlStr,dbget,1
		if not rsget.eof Then
			Do Until rsget.eof
				If rsget("docuSignId") <> "" Then
					Session.CodePage = 65001
					Set objXML= CreateObject("MSXML2.XMLHTTP.3.0")
					'Set objXML = CreateObject("Msxml2.ServerXMLHTTP")
					'objXML.SetTimeouts 40000, 40000, 40000, 40000
					objXML.Open "GET", FecDocuURL&"/api/contract/v1/docusign/envelope/"&rsget("docuSignId"), False
					objXML.setRequestHeader "Content-Type", "application/json"
					if (application("Svr_Info")	<> "Dev") then
						objXML.SetRequestHeader "api-key-v1", ""+CStr(adminApiKey)+""
					End If					
					objXML.Send
					Session.CodePage = 949
					iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")

					If objXML.Status = "200" Then
						Set jsResult = JSON.parse(iRbody)
						docuSignEnvelopeId = jsResult.envelopeId
						docuSignStatus = jsResult.status
						docuSignStatusDateTime = jsResult.statusDateTime
						docuSignUri = jsResult.uri
						Set jsResult = Nothing

						Select Case Trim(docuSignStatus)
							case "created"
								docuStatusAdminCodeConversion = 1		
							case "sent"
								docuStatusAdminCodeConversion = 1
							case "delivered"
								docuStatusAdminCodeConversion = 1
							case "signed"
								docuStatusAdminCodeConversion = 6
							case "declined"
								docuStatusAdminCodeConversion = 2
							case "completed"
								docuStatusAdminCodeConversion = 7
							'case "faxpending" '' 텐바이텐에선 사용안함
							'case "autoresponded" '' 텐바이텐에선 사용안함
							Case Else
								docuStatusAdminCodeConversion = rsget("ctrstate")
						End Select					

						sqlStr = "update db_partner.dbo.tbl_partner_ctr_master set ctrstate = "&docuStatusAdminCodeConversion&", lastupdate =getdate()"
						sqlstr = sqlstr & " where groupid = '"&groupid&"' AND ctrkey='"&rsget("ctrKey")&"' and ctrstate<>-1 and ISNULL(docuSignId,'') <> '' and signType='D' AND DocuSignId='"&rsget("docuSignId")&"' "
						dbget.Execute  sqlstr, 1						
					Else
						docuErrorStatusValue = docuErrorStatusValue &","& objXML.Status
						'response.write "<script>alert('DocuSign 통신중 오류가 발생하였습니다.\nErrorCode("&objXML.Status&")');</script>"
						'response.write "<script>location.replace('" & refer & "');</script>"
						'dbget.close() : response.End
					End If
					Set objXML = Nothing

				End If
			rsget.movenext
			loop
		Else
			Response.write "<script type='text/javascript' language='javascript'>alert('요청한 groupid로 등록된 DocuSign 문서가 없습니다.');location.href = '"&refer&"';</script>"
			response.end
		end if
		rsget.close

		'If Trim(docuSignEnvelopeId) = "" Then
		'	Response.write "<script type='text/javascript' language='javascript'>alert('DocuSign 상태를 확인할 수 없습니다.');location.href = '"&refer&"';</script>"
		'	response.end
		'End If
		If Trim(docuErrorStatusValue) <> "" Then
			Response.write "<script type='text/javascript' language='javascript'>alert('DocuSign 상태가 업데이트 되었지만\n상태 업데이트 중 통신 오류가 발생한 건이 있습니다.("&docuErrorStatusValue&")');location.href = '"&refer&"';</script>"
			response.end
		Else
			Response.write "<script type='text/javascript' language='javascript'>alert('DocuSign 상태가 업데이트 되었습니다.');location.href = '"&refer&"';</script>"
			response.end
		End If

	elseif mode="ecuser" then
		if False and ecBUser ="" then
			Response.write "<script type='text/javascript' language='javascript'>alert('수정할 담당자명이 없습니다. 다시 시도해주세요');history.back();</script>"
			response.end
		end if

		sqlStr = " select  count(ctrstate) as cnt "
		sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master with (nolock)"
		sqlStr = sqlStr & "	where ctrstate < 7 and ctrstate >= 0 and  ecCtrSeq <> '0' and  ecCtrSeq is not null and groupid = '"&groupid&"'"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.eof Then
			ctrcount = rsget("cnt")
		end if
		rsget.close

		if ctrcount < 1 then
			Response.write "<script type='text/javascript' language='javascript'>alert('계약상태가 계약완료인 계약서는 전자계약담당자를 수정 하실수 없습니다.');history.back();</script>"
			response.end
		end if

		sqlStr ="update db_partner.dbo.tbl_partner_ctr_master set ecBUser ='"&trim(ecBUser)&"', lastupdate =getdate() where ctrstate < 7 and ctrstate >= 0 and  ecCtrSeq <> '0' and  ecCtrSeq is not null and groupid = '"&groupid&"' "
		dbget.Execute  sqlstr, 1

		Response.write "<script type='text/javascript' language='javascript'>alert('전자계약서 담당자명이 수정되었습니다.');location.href = '"&refer&"';</script>"
		response.end
	else
		response.write "<script>alert('정의되지 않았습니다. - " & mode & "');</script>"
	end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
