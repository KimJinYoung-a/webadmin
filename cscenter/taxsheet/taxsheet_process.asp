<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/cscenter/lib/TaxSheetFunc.asp"-->
<%

	'// 변수 선언 //
	dim taxIdx, busiIdx, mode, menupos
	dim page, searchDiv, searchKey, searchString, param
	dim SQL, msg, retURL

	dim oTax
	dim taxSheetIssueType

	'// 파라메터 접수 //
	taxIdx = request("taxIdx")
	if taxIdx="" and request("chkSelect")<>"" then
		taxIdx = request("chkSelect")
	end if
	busiIdx = request("busiIdx")
	mode = request("mode")
	menupos = request("menupos")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")

	param = "&menupos=" & menupos & "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수


	'// 전송값 체크 함수 //
	Function chkTRStr(str)
		dim stmp
		'stmp = Replace(str,"&", "#amp;")
		'stmp = Replace(str,"=", "#equ;")
		stmp = Server.URLEncode(str)
		chkTRStr = stmp
	End Function


	'///// 모드별 분기 처리 /////
	Select Case mode
		'사업자등록증 확인 처리
		Case "BusiOk"
			SQL =	" Update db_order.[dbo].tbl_busiinfo Set " &_
					"	confirmYn='Y '" &_
					" Where busiIdx=" & busiIdx
			msg = "사업자등록증을 확인처리하였습니다."
			retURL = "tax_view.asp?taxIdx=" & taxIdx & param

		'요청서 출력 처리
		Case "sheetOk"
			msg = fnTaxSheetPrint(taxIdx)
			if msg="OK" then
				SQL=""
				msg = "세금계산서를 발행하였습니다."
				retURL = "tax_view.asp?taxIdx=" & taxIdx & param
			else
			    ''response.write "msg:"&msg
				response.write	"<script language='javascript'>" &_
								"	alert('" & msg & "');" &_
								"	history.back();" &_
								"</script>"
				dbget.close()	:	response.End
			end if

		'요청서 일괄출력 처리
		Case "BatchOk"
			dim arrIdx, lp
			arrIdx = split(taxIdx,",")
			for lp=0 to ubound(arrIdx)
				msg = fnTaxSheetPrint(arrIdx(lp))
				if msg<>"OK" then
					response.write	"<script language='javascript'>" &_
									"	alert('" & msg & " [" & arrIdx(lp) & "]');" &_
									"	history.back();" &_
									"</script>"
					dbget.close()	:	response.End
				end if
			next

			SQL=""
			msg = "세금계산서 총 " & ubound(arrIdx)+1 & " 건 발행하였습니다."
			retURL = "tax_list.asp?taxIdx=" & taxIdx & param

		'요청서 삭제
		Case "sheetDel"
			SQL =	" Update db_order.[dbo].tbl_taxSheet Set " &_
					"	delYn = 'Y' " &_
					" Where taxIdx=" & taxIdx
			msg = "세금계산서 요청 내용을 삭제하였습니다."
			retURL = "tax_list.asp?taxIdx=" & param
		Case Else

	End Select

	'// DB처리 및 페이지 이동

	'트랜젝션 시작
	if SQL<>"" then
		dbget.beginTrans
		dbget.Execute(SQL)

		if (mode = "sheetDel") then

			set oTax = new CTax
			oTax.FRecttaxIdx = taxIdx

			oTax.GetTaxRead

			taxSheetIssueType = GetTaxSheetIssueType(taxIdx)

			if (taxSheetIssueType = "orderserial") then
			    SQL = " update [db_order].[dbo].tbl_order_master" & VbCrlf
			    SQL = SQL & " set " & VbCrlf
			    SQL = SQL & " authcode = (case when accountdiv in ('7', '20') then NULL else authcode end) " + VbCrlf
			    SQL = SQL & " , cashreceiptreq = NULL " + VbCrlf
			    SQL = SQL & " where orderserial='" & oTax.FTaxList(0).Forderserial & "'"
			    dbget.Execute SQL

			elseif (taxSheetIssueType = "etcmeachul") then
				SQL = " update db_shop.dbo.tbl_fran_meachuljungsan_master" + vbCrlf
				SQL = SQL + " set issuestatecd = NULL " + vbCrlf
				SQL = SQL + " where idx=" & oTax.FTaxList(0).Forderidx
				rsget.Open SQL, dbget, 1

				if CStr(oTax.FTaxList(0).Forderidx) = "0" then
					SQL = " update db_shop.dbo.tbl_fran_meachuljungsan_master" + vbCrlf
					SQL = SQL + " set issuestatecd = NULL " + vbCrlf
					SQL = SQL + " where idx in (select matchlinkkey from db_order.dbo.tbl_taxSheet_Match where matchtype = 'E' and taxidx = " & oTax.FTaxList(0).Ftaxidx & ") "
					rsget.Open SQL, dbget, 1
				end if
			end if


			''if (Not IsNull(oTax.FTaxList(0).Forderserial)) and (oTax.FTaxList(0).Forderserial <> "") then '고객 세금계산서 신청
			''	if (Left(oTax.FTaxList(0).Forderserial, 2) <> "SO") then
			''	end if
			''end if
		end if
	end if

	'오류검사 및 반영
	If Err.Number = 0 Then
		if SQL<>"" then
			dbget.CommitTrans				'커밋(정상)
		end if

		response.write	"<script language='javascript'>" &_
						"	alert('" & msg & "');" &_
						"	self.location='" & retURL & "';" &_
						"</script>"
	Else
	    dbget.RollBackTrans				'롤백(에러발생시)

		response.write	"<script language='javascript'>" &_
						"	alert('처리중 에러가 발생했습니다.');" &_
						"	history.back();" &_
						"</script>"
	End If

	'### 세금계산서 발급 함수 ###
	Function fnTaxSheetPrint(taxIdx)

		dim oTax, strSql, isueDate
	    dim preIssuedCnt
	    ''기발행내역 있는지 체크 (중복체크) : 세금계산서.
	    strSql = "select count(*) as cnt from db_order.[dbo].tbl_taxSheet"
        strSql = strSql&" where orderserial in (select orderserial from db_order.[dbo].tbl_taxSheet where taxidx="&taxIdx&")"
        strSql = strSql&" and isueYn='Y' and DelYn='N'"
        strSql = strSql&" and taxidx<>"&taxIdx&""

        rsget.Open strSql, dbget, 1
           preIssuedCnt  = rsget("cnt")
        rsget.Close

        if (preIssuedCnt>0) then
            errMsg = "기존 발행 내역이 존재 합니다."
            fnTaxSheetPrint = errMsg
            exit function
        end if

        ''현금영수증 신청내역 있는지 체크
        strSql = 	" Select count(*) as cnt " &_
			" From [db_log].[dbo].tbl_cash_receipt " &_
			" Where orderserial in (select orderserial from db_order.[dbo].tbl_taxSheet where taxidx="&taxIdx&")" &_
			"	and cancelyn='N'" &_
			"	and resultcode='00'"

		rsget.Open strSql, dbget, 1
           preIssuedCnt  = rsget("cnt")
        rsget.Close

        if (preIssuedCnt>0) then
            errMsg = "현금영수증 발행 내역이 존재 합니다."
            fnTaxSheetPrint = errMsg
            exit function
        end if


		Dim objXMLHTTP
		Dim reqParam '호출파라미터
		Dim tmpArr, errMsg
		Dim retval  '호출결과
		Dim tax_no  '세금계산서 번호
		Dim tel1, tel2, tel3, sIdx
		Dim cur_c_corp_no, cur_u_user_no
		errMsg = "OK"

		'# 기본 내용 접수
		set oTax = new CTax
		oTax.FRecttaxIdx = taxIdx
		oTax.GetTaxRead

		'전화번호 분리
		if db2html(oTax.FTaxList(0).FrepTel)<>"" then
		    oTax.FTaxList(0).FrepTel = Replace(oTax.FTaxList(0).FrepTel,")","-")
			if instr(db2html(oTax.FTaxList(0).FrepTel),"-")<1 then
				tmpArr = array(left(db2html(oTax.FTaxList(0).FrepTel),3),mid(db2html(oTax.FTaxList(0).FrepTel),4,4),right(db2html(oTax.FTaxList(0).FrepTel),4))
			else
				tmpArr = split(db2html(oTax.FTaxList(0).FrepTel),"-")
			end if
			tel1 = tmpArr(0)
			tel2 = tmpArr(1)
			tel3 = tmpArr(2)
		end if
		sIdx = Replace(Date(),"-", "") & taxIdx

		'작성일 지정(2009.01.06;허진원)
		if datediff("m",oTax.FTaxList(0).Fipkumdate,date())=0 then
			'입급일이 현재달과 같으면 입금일로
			isueDate = dateSerial(Year(oTax.FTaxList(0).Fipkumdate),Month(oTax.FTaxList(0).Fipkumdate),Day(oTax.FTaxList(0).Fipkumdate))
		elseif datediff("m",oTax.FTaxList(0).Fipkumdate,date())=1 and datediff("d",date(),dateSerial(year(date),month(date),5))>=0 then
			'입급일이 지난달이면서 당월 5일 이전이라면 입금일로
			isueDate = dateSerial(Year(oTax.FTaxList(0).Fipkumdate),Month(oTax.FTaxList(0).Fipkumdate),Day(oTax.FTaxList(0).Fipkumdate))
		else
			'그렇지 않으면 금월 1일로 세팅
			isueDate = dateSerial(Year(date),Month(date),1)
		end if

		'테스트/실서버 구분
		if (application("Svr_Info")="Dev") then
    		cur_c_corp_no = "10001568"		'(기업 번호 - test:10001568)
    		cur_u_user_no = "1000794"		'(담당자 번호 - test:1000794)
    	else
    		cur_c_corp_no = "57911"			'(기업 번호 - Real:57911)
    		cur_u_user_no = "261746"		'(담당자 번호 - Real:261746)
    	end if

	    '전송 파라메터 구성
	    '"&item_nm=" & chkTRStr(LeftB(db2html(oTax.FTaxList(0).Fitemname),40)) &_ (상품명 통일;2006-06-26;허진원)
		reqParam =	"uniq_id=" & sIdx &_
	    			"&biz_no=" & Replace(oTax.FTaxList(0).FbusiNo,"-", "") &_
	    			"&corp_nm=" & chkTRStr(db2html(oTax.FTaxList(0).FbusiName)) &_
	    			"&ceo_nm=" & chkTRStr(db2html(oTax.FTaxList(0).FbusiCEOName)) &_
	    			"&biz_status=" & chkTRStr(db2html(oTax.FTaxList(0).FbusiType)) &_
	    			"&biz_type=" & chkTRStr(db2html(oTax.FTaxList(0).FbusiItem)) &_
	    			"&addr=" & chkTRStr(db2html(oTax.FTaxList(0).FbusiAddr)) &_
	    			"&dam_nm=" & chkTRStr(db2html(oTax.FTaxList(0).FrepName)) &_
	    			"&email=" & chkTRStr(db2html(oTax.FTaxList(0).FrepEmail)) &_
	    			"&hp_no1=" & tel1 &_
	    			"&hp_no2=" & tel2 &_
	    			"&hp_no3=" & tel3 &_
	    			"&serial_no=" & sIdx &_
	    			"&item_count=1" &_
	    			"&item_nm=" & chkTRStr("디자인사무용품") &_
	    			"&item_qty=1" &_
	    			"&item_price=" & oTax.FTaxList(0).FtotalPrice-oTax.FTaxList(0).FtotalTax &_
	    			"&item_vat=" & oTax.FTaxList(0).FtotalTax &_
	    			"&cur_c_corp_no=" & cur_c_corp_no &_
	    			"&cur_u_user_no=" & cur_u_user_no &_
					"&api_no=1" &_
					"&enc_yn=N" &_
					"&cur_biz_no=2118700620" &_
					"&cur_corp_nm=" & chkTRStr("(주)텐바이텐") &_
					"&cur_ceo_nm=" & chkTRStr("이창우") &_
					"&cur_biz_status=" & chkTRStr("서비스") &_
					"&cur_biz_type=" & chkTRStr("전자상거래") &_
					"&cur_addr=" & chkTRStr("서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐") &_
					"&cur_dam_nm=" & chkTRStr("텐바이텐") &_
					"&cur_email=" & chkTRStr("customer@10x10.co.kr") &_
					"&cur_hp_no1=02" &_
					"&cur_hp_no2=1644" &_
					"&cur_hp_no3=6030" &_
					"&write_date=" & replace(isueDate,"-","") &_
					"&sb_type=01" &_
					"&pc_gbn=C" &_
					"&tax_type=01" &_
					"&bill_type=01" &_
					"&approve_type=11" &_
					"&final_status=12" &_
					"&Cash_amt=" & oTax.FTaxList(0).FtotalPrice
		'' "&credit_amt=" & oTax.FTaxList(0).FtotalPrice 서동석 추가 : 외상미수금으로 표시
		'response.Write reqParam
		'dbget.close()	:	response.End

		'XML통신 컨퍼넌트 선언
		Set objXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
		if (application("Svr_Info")="Dev") then
			objXMLHTTP.Open "POST",	"http://ifs.neoport.net:8383/tx_create.req", False
		else
			objXMLHTTP.Open "POST",	"http://web1.neoport.net:8383/tx_create.req", False
		end if
		objXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		objXMLHTTP.Send reqParam
'response.write "retval="&retval
		retval = trim(objXMLHTTP.responseText)
		retval = replace(retval,Vbcrlf,"")
		retval = replace(retval,Vbcr,"")
		retval = replace(retval,Vblf,"")
		Set objXMLHTTP = Nothing


	    If retval <> "" Then
	        tmpArr = Split(retval, "|")
	        tax_no = Clng(tmpArr(0))

	        if tax_no<0 then
	        	errMsg = retval
	        else
	        	errMsg = "OK"
	        end if
	    else
			errMsg = "통신중 오류가 발생했습니다."
	    End If

		if errMsg="OK" then
			'데이터 처리
			strSql =	" Update db_order.[dbo].tbl_busiinfo Set " &_
					"	confirmYn='Y '" &_
					" Where busiIdx=" & oTax.FTaxList(0).FbusiIdx & ";" & chr(13)&chr(10)

			strSql = strSql & " Update db_order.[dbo].tbl_taxSheet Set " &_
						"	neoTaxNo = '" & tax_no & "' " &_
						"	,curUserId = '" & Session("ssBctId") & "' " &_
						"	,printDate = getdate() " &_
						"	,isueYn = 'Y' " &_
						"	,isueDate = '" & isueDate & "' " &_
						" Where taxIdx=" & taxIdx
			dbget.Execute(strSql)
		end if

		set oTax = Nothing

		'// 결과 반납
		fnTaxSheetPrint = errMsg
	End Function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

