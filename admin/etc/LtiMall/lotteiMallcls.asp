<%
'' 배송정책  3만원 이하 2500
CONST CMAXMARGIN = 15			'' MaxMagin임.. '(롯데iMall 10%)
CONST CMAXLIMITSELL = 5        '' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CMALLNAME = "lotteimall"
CONST CLTIMALLMARGIN = 11       ''마진 11%
CONST CHEADCOPY = "Design Your Life! 새로운 일상을 만드는 감성생활브랜드 텐바이텐" ''생활 감성채널 텐바이텐
CONST CPREFIXITEMNAME ="[텐바이텐]"
CONST CitemGbnKey ="K1099999" ''상품구분키 ''하나로 통일
CONST CUPJODLVVALID = TRUE   ''업체 조건배송 등록 가능여부

CONST ENTP_CODE = "011799"                                    '' 협력사코드
CONST MD_CODE   = "0168"                                      '' MD_Code
CONST BRAND_CODE   = "1099329"                                '' 롯데에 받아야함
CONST BRAND_NAME   = "텐바이텐(10x10)"                        '' 롯데에 받아야함
CONST MAKECO_CODE  = "9999"                                   '' 롯데에 받아야함
CONST CDEFALUT_STOCK = 99       '' 재고관리 수량 기본 99 (한정 아닌경우)

Class CLotteiMallItem
	Public FLastUpdate
	Public FisUsing

	'담당MD
	Public FMDCode
	Public FMDName
	Public FSellFeeType
	Public FNormalSellFee
	Public FEventSellFee

	'MD상품군
	Public FgroupCode               ''' 롯데iMall =>LCode. 50000000 : 전문몰
	Public FSuperGroupName
	Public FGroupName

	'롯데닷컴 카테고리
	Public FitemGbnKey
	Public FitemGbnNm

	Public FDispNo
	Public FDispNm

	Public FDispLrgNm
	Public FDispMidNm
	Public FDispSmlNm
	Public FDispThnNm

	Public FGbnLrgNm
	Public FGbnMidNm
	Public FGbnSmlNm
	Public FGbnThnNm
	Public FCateIsUsing
	Public FItemcnt

	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FtenCateName
	Public Fdisptpcd

	'롯데닷컴 브랜드
	Public FlotteBrandCd
	Public FlotteBrandName
	Public FTenMakerid
	Public FTenBrandName

	'롯데닷컴 상품목록
	Public FLTiMallRegdate
	Public FLTiMallLastUpdate
	Public FLTiMallGoodNo				'실상품번호
	Public FLTiMallTmpGoodNo			'임시상품번호
	Public FLTiMallPrice
	Public FLTiMallSellYn
	Public FregUserid
	Public FLotteDispCnt
	Public FCateMapCnt
	Public FLTiMallStatCd				'상품등록상태
	Public FregedOptCnt
	Public FrctSellCNT
	Public FaccFailCNT              '등록수정 오류 횟수
	Public FlastErrStr              '최종오류

	'텐바이텐 상품목록
	Public Fitemid
	Public Fitemname
	Public FitemDiv
	Public FsmallImage
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fmakerid
	Public Fregdate
	Public ForgPrice
	Public ForgSuplyCash
	Public FSellCash
	Public FBuyCash
	Public FsellYn
	Public FsaleYn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public Fkeywords
	Public ForderComment
	Public FoptionCnt
	Public Fsourcearea
	Public Fmakername
	Public Fitemcontent
	Public FUsingHTML
	Public Fdeliverytype
	Public Fvatinclude
	Public Fdefaultdeliverytype
	Public FdefaultfreeBeasongLimit
	Public FrequireMakeDay
	public FmaySoldOut

	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FOutmallstandardMargin

	Public FoptAddPrcCnt
	Public FoptAddPrcRegType

	Public FRectMode    ''??
	Public Fidx
	Public FNewitemname
	Public FItemnameChange
	Public FItemoption
	Public FOptaddprice
	Public FOptionname
	Public FOptlimitno
	Public FOptlimitsold
	Public FOptsellyn
	Public FRegedOptionname
	Public FRegedItemname
	Public FSpecialPrice
	Public FStartDate
	Public FEndDate
	Public FPurchasetype

	Public Function getRealItemname
		If FitemnameChange = "" Then
			getRealItemname = FNewitemname
		Else
			getRealItemname = FItemnameChange
		End If
	End Function

	Function getLimitEa()
		dim ret : ret = (FLimitno-FLimitSold)
		if (ret<1) then ret=0
		getLimitEa = ret
	End Function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<font color=blue>한정:"&getLimitEa&"</font>"
	    End if
	End Function

	'// 품절여부
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Function getNOREST_ALLOW_MONTH()
	    '1~29만원 : 일시불
	    '30~59만원 : 5개월
	    '60~99만원 이하 : 7개월
	    '100만원 이상 : 10개월
	    Dim retVal : retVal = ""
	    If (FSellCash < 300000) Then
	        exit function
	    ElseIf (FSellCash < 600000) Then
	        getNOREST_ALLOW_MONTH = "5"
	    ElseIf (FSellCash < 1000000) Then
	        getNOREST_ALLOW_MONTH = "7"
	    ElseIf (FSellCash >= 1000000) Then
	        getNOREST_ALLOW_MONTH = "10"
	    End If
	End Function

	Function getItemNameFormat()
		Dim buf
		buf = replace(FItemName,"'","")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","프로")
		buf = replace(buf,"[무료배송]","")
		buf = replace(buf,"[무료 배송]","")
		getItemNameFormat = buf
	End Function

	''옵션구분명 - :안됨 max20Byte
	Function getGOODSDT_NmFormat(idtname)
		Dim buf
		buf = Replace(db2Html(idtname),":","")
		buf = Replace(buf,"디자인을 선택해주세요","디자인 선택")
		buf = Replace(buf,"디자인을 선택 하세요","디자인 선택")
		buf = Replace(buf,"디자인을 선택해 주세요","디자인 선택")
		buf = Replace(buf,"디자인을 골라주세요","디자인 선택")
		buf = Replace(buf,"다이어리 선택하기!","다이어리 선택")
		getGOODSDT_NmFormat = Trim(buf)
	End Function

	Function getLTiMallSuplyPrice()
	    getLTiMallSuplyPrice = CLNG(FSellCash*(100-CLTIMALLMARGIN)/100)
	End Function

	Function getDispGubunNm()
		getDispGubunNm = getDisptpcdName
	End Function

	Public Function getDisptpcdName
        if (Fdisptpcd="10") then
            getDisptpcdName = "일반"
        elseif (Fdisptpcd="11") then
            getDisptpcdName = "브랜드"
        elseif (Fdisptpcd="12") then
            getDisptpcdName = "<font color='blue'>전문</font>"
        elseif (Fdisptpcd="99") then
            getDisptpcdName = "<font color='red'>신규</font>"
        else
            getDisptpcdName = Fdisptpcd
        end if
	End Function

	Public Function getDeliverytypeName
		If (Fdeliverytype = "9") Then
			getDeliverytypeName = "<font color='blue'>[조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
		ElseIf (Fdeliverytype = "7") then
			getDeliverytypeName = "<font color='red'>[업체착불]</font>"
		ElseIf (Fdeliverytype = "2") then
			getDeliverytypeName = "<font color='blue'>[업체]</font>"
		Else
			getDeliverytypeName = ""
		End If
	End Function

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// 검색어배열
	Public Function getItemKeywordArray(sno)
		Dim arrRst, arrRst2
		If trim(Fkeywords) = "" Then Exit Function

		arrRst = split(Fkeywords,",")
		If ubound(arrRst) = 0 Then
			'구분이 공백일 경우
			arrRst2 = split(arrRst(0), " ")
			If ubound(arrRst2) > 0 Then
				arrRst = split(Fkeywords, " ")
			End If
		End If

		If ubound(arrRst) >= sno Then
			getItemKeywordArray = trim(arrRst(sno))
		Else
			getItemKeywordArray = ""
		End If
	End Function

	'// 검색어
	Public Function getItemKeyword()
		Dim arrRst, arrRst2, q, p, r, divBound1, divBound2, divBound3, Keyword1, Keyword2, Keyword3, strRst
		If trim(Fkeywords) = "" Then Exit Function

		If Len(Fkeywords) > 50 Then
			arrRst = Split(Fkeywords,",")
			If Ubound(arrRst) = 0 then
				'구분이 공백일 경우
				arrRst2 = split(arrRst(0)," ")
				If Ubound(arrRst2) > 0 then
					arrRst = split(Fkeywords," ")
				'2013-10-22 김진영 수정..ex)826121, 826124
				Else
					'구분이 세미콜론일 경우
					arrRst2 = split(arrRst(0),";")
					If Ubound(arrRst2) > 0 then
						arrRst = split(Fkeywords,";")
					End If
				End If
			End If
			'키워드 1
			divBound1 = CLng(Ubound(arrRst)/3)
			For q = 0 to divBound1
				Keyword1 = Keyword1&arrRst(q)&","
			Next
			If Right(keyword1,1) = "," Then
				keyword1 = Left(keyword1,Len(keyword1)-1)
			End If

			'키워드 2
			divBound2 = divBound1 + 1
			For p = divBound2 to divBound2 + divBound1
				Keyword2 = Keyword2&arrRst(p)&","
			Next
			If Right(keyword2,1) = "," Then
				keyword2 = Left(keyword2,Len(keyword2)-1)
			End If

			'키워드 3
			divBound3 = divBound2 + divBound1
			For r = divBound3 to Ubound(arrRst)
				Keyword3 = Keyword3&arrRst(r)&","
			Next
			If Right(keyword3,1) = "," Then
				keyword3 = Left(keyword3,Len(keyword3)-1)
			End If

			strRst = ""
			strRst = strRst & "&sch_kwd_1_nm="&Keyword1
			strRst = strRst & "&sch_kwd_2_nm="&Keyword2
			strRst = strRst & "&sch_kwd_3_nm="&Keyword3
			getItemKeyword = strRst
		Else
			strRst = ""
			strRst = strRst & "&sch_kwd_1_nm="&Fkeywords
			strRst = strRst & "&sch_kwd_2_nm="
			strRst = strRst & "&sch_kwd_3_nm="
			getItemKeyword = strRst
		End If
	End Function

	''//상품명 변경 파라메터 생성(롯데닷컴과 파라매타명이 다름)
	Public Function GetLtiMallItemNameEditParameter()
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo
		strRst = strRst & "&goods_no=" & FLTiMallGoodNo
		strRst = strRst & "&goods_nm=" & Trim(getItemNameFormat)
		strRst = strRst & "&chg_caus_cont=api 상품명 변경"
		GetLtiMallItemNameEditParameter = strRst
	End Function

    '// 가격 수정 파라메터 생성
    Public Function getLtiMallItemPriceEditParameter()
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo
		strRst = strRst & "&strGoodsNo=" & FLTiMallGoodNo
		strRst = strRst & "&strReqSalePrc=" & GetRaiseValue(MustPrice/10)*10
		getLtiMallItemPriceEditParameter = strRst
    End Function

	Public Function MustPrice
		Dim GetTenTenMargin
		'2013-07-25 김진영//텐텐마진이 iMALL의 마진보다 작을 때 orgprice로 전송 시작
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < FOutmallstandardMargin Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		'2013-07-25 김진영//텐텐마진이 iMALL의 마진보다 작을 때 orgprice로 전송 끝
	End Function

	'// 텐바이텐 상품옵션 검사
	Public Function checkTenItemOptionValid()
		Dim strSql, chkRst, chkMultiOpt
		Dim cntType, cntOpt
		chkRst = true
		chkMultiOpt = false

		If FoptionCnt > 0 Then
			'// 이중옵션확인
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				cntType = rsget.RecordCount
			End If
			rsget.Close

			If chkMultiOpt Then
				'// 이중옵션 일때
				strSql = "Select optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						cntOpt = ubound(split(db2Html(rsget("optionname")), ",")) + 1
						If cntType <> cntOpt then
							chkRst = false
						End If
						rsget.MoveNext
					Loop
				Else
					chkRst = false
				End If
				rsget.Close
			Else
				'// 단일옵션일 때
				strSql = "Select optionTypeName, optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1
				If (rsget.EOF or rsget.BOF) Then
					chkRst = false
				End If
				rsget.Close
			End If
		End If
		'//결과 반환
		checkTenItemOptionValid = chkRst
	End Function

	'// 롯데닷컴 판매여부 반환
	Public Function getLTiMallSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn = "Y" and FisUsing = "Y" Then
			If FLimitYn="N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) Then
				getLTiMallSellYn = "Y"
			Else
				getLTiMallSellYn = "N"
			End if
		Else
			getLTiMallSellYn = "N"
		End If
	End Function

	'// 롯데아이몰 등록상태 반환 // 사용안함
	Public Function getLotteItemStatCd()
	    Select Case FLTiMallStatCd
		    Case "0"
				getLotteItemStatCd = "등록예정"
			Case "10"
				getLotteItemStatCd = "전송시도"         ''통신시도( 구 임시등록)
			Case "20"
				getLotteItemStatCd = "승인요청"         ''1차등록
			Case "30"
				getLotteItemStatCd = "승인완료"
			Case "40"
				getLotteItemStatCd = "반려"
			Case "50"
				getLotteItemStatCd = "승인불가"
			Case "51"
				getLotteItemStatCd = "재승인요청"
			Case "52"
				getLotteItemStatCd = "수정요청"
			CASE ELSE
			    getLotteItemStatCd = FLTiMallStatCd
		End Select
	End Function

	'// 롯데아이몰 등록상태 반환
	public function getLTIMallStatCDName()
	    Select Case FLTiMallStatCd
		    Case 0
				getLTIMallStatCDName = "등록예정"
			Case 1
				getLTIMallStatCDName = "전송시도"         ''통신시도( 구 임시등록)
			Case 20
				getLTIMallStatCDName = "승인요청"         ''1차등록
			Case 7
				getLTIMallStatCDName = "승인완료"
			Case -1
				getLTIMallStatCDName = "등록실패"
 			CASE -9
 				getLTIMallStatCDName = "미등록"
			CASE "40"
				getLTIMallStatCDName = "반려"
			CASE "50"
				getLTIMallStatCDName = "승인불가"
			CASE "51"
				getLTIMallStatCDName = "재승인요청"
			CASE "52"
				getLTIMallStatCDName = "수정요청"
			CASE ELSE
			    getLTIMallStatCDName = FLTiMallStatCd
		End Select


    end function

	'// 롯데아이몰 판매여부 반환
	Public Function getLotteiMallSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getLotteiMallSellYn = "Y"
			Else
				getLotteiMallSellYn = "N"
			End If
		Else
			getLotteiMallSellYn = "N"
		End If
	End Function

	Public Function getLimitLotteEa()
		Dim ret
		ret = FLimitNo - FLimitSold - 5
		If (ret < 1) Then ret = 0
		getLimitLotteEa = ret
	End Function

	'// 상품등록 파라메터 생성
	Public Function getLotteiMallItemRegParameter(isEdit)
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)사용자인증키
		If (isEdit) Then
		   strRst = strRst & "&goods_req_no="&FLTiMallTmpGoodNo
		End If
		strRst = strRst & "&brnd_no=" & BRAND_CODE											'(*)브랜드코드
		strRst = strRst & "&goods_nm=" & Trim(getItemNameFormat)							'(*)전시상품명
'		strRst = strRst & "&sch_kwd_1_nm=" & getItemKeywordArray(0)							'키워드1
'		strRst = strRst & "&sch_kwd_2_nm=" & getItemKeywordArray(1)							'키워드2
'		strRst = strRst & "&sch_kwd_3_nm=" & getItemKeywordArray(2)							'키워드3
		strRst = strRst & getItemKeyword
		strRst = strRst & "&mdl_no="															'모델번호(?)
		strRst = strRst & "&pur_shp_cd=3" 													'(*)매출형태(1.직매입, 4.특정, 3.특정판매)	롯데닷컴은 2(판매분매입)로 설정되어있음..아이몰엔 2가 없는데..그래서 4로 놓긴했는데; ''3일듯: 현아 확인
		strRst = strRst & "&sale_shp_cd=10" 												'(*)판매형태코드(10:정상)
		strRst = strRst & "&sale_prc=" & cLng(GetRaiseValue(FSellCash/10)*10)				'(*)판매가
		strRst = strRst & "&mrgn_rt="&CLTIMALLMARGIN 										'(*)마진율(7/1일 시스템 개편하면서 11로 바뀐다함..)
'		strRst = strRst & "&pur_prc=" & cLng(FSellCash*0.88)									'공급가(REQUEST 파람에는 없으나 샘플파일 넘길때는 있던데??) :: 안넣어도 등록가능
'		strRst = strRst & "&tdf_sct_cd=1" 													'(*)과면세코드(1:과세, 2:면세)	'2013-11-11 18:09 김진영 수정//롯데닷컴처럼 모두 과세로 되어있던 상태 수정
		strRst = strRst & "&tdf_sct_cd="&CHKIIF(FVatInclude="N","2","1")					'(*)과면세코드(1:과세, 2:면세)
		strRst = strRst & getLotteiMallCateParamToReg()											'(*)MD상품군 및 해당 전시카테고리(상품수정에서 카테고리 변경이 안 됨..2013-07-02 전시카테고리 수정API로 수정
		If (FLimitYn="Y") then
		    strRst = strRst & "&inv_mgmt_yn=Y"												'(*)재고관리여부(롯데닷컴처럼 변형) 2013-06-24 김진영
			If FoptionCnt = 0 then
				strRst = strRst & "&inv_qty="&getLimitLotteEa()								'재고수량
			End If
		Else
			strRst = strRst & "&inv_mgmt_yn=Y" 												'(*)재고관리여부(롯데닷컴처럼 변형) 2013-06-24 김진영
			If FoptionCnt = 0 then
			    strRst = strRst & "&inv_qty="&CDEFALUT_STOCK								'디폴트 수량 99로
			End if
		End If
		strRst = strRst & getLotteiMallOptionParamToReg()									'옵션명 및 옵션상세 :: 단품번호 추가
		strRst = strRst & "&add_choc_tp_cd_10="													'날짜선택형옵션
		If FitemDiv = "06" Then
			strRst = strRst & "&add_choc_tp_cd_20=주문제작상품"						 		'입력형옵션
		End If

		If FitemDiv="06" or FitemDiv="16" then
			strRst = strRst & "&exch_rtgs_sct_cd=10"																					'교환/반품여부 10:불가능 / 20:가능
		Else
			strRst = strRst & "&exch_rtgs_sct_cd=20"																					'교환/반품여부 10:불가능 / 20:가능
		End If

		strRst = strRst & "&dlv_proc_tp_cd=1" 												'(*)배송유형(1:업체배송, 3:센터배송, 4:센터경유, 6:e-쿠폰배송)
		strRst = strRst & "&gift_pkg_yn=N" 													'(*)선물포장여부
		strRst = strRst & "&dlv_mean_cd=10" 												'(*)배송수단(10:택배 ,11:명절퀵배송 ,40:현장수령 ,50:DHL ,60:해외우편 ,70:일반우편 ,80:등기우편)
		strRst = strRst & getLotteiMallGoodDLVDtParams										'(*)배송상품구분 및 배송기일
		strRst = strRst & "&imps_rgn_info_val="													'배송불가지역(10:서울,수도권, 21:지방, 22:도서지역, 23:인천영종도, 30:제주) 여러개의경우:(콜론)으로 구분하여 전송 한개라도 콜론으로 전송
		strRst = strRst & "&byr_age_lmt_cd=0" 												'(*)구입자나이제한(0:전체, 19:19세이상)
		If Fitemid = "407171" or Fitemid = "788038" or Fitemid = "785541" or Fitemid = "785540" or Fitemid = "785542" or Fitemid = "679670" or Fitemid = "620503" or Fitemid = "590196" or Fitemid = "221081" Then
		strRst = strRst & "&dlv_polc_no=" & tenDlvFreeCd									'(*)배송정책번호(???) tenDlvCd는 inc_dailyAuthCheck.asp에서 정의 (API_TEST에서 따옴)
		Else
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd										'(*)배송정책번호(???) tenDlvCd는 inc_dailyAuthCheck.asp에서 정의 (API_TEST에서 따옴)
		End If
		strRst = strRst & "&corp_dlvp_sn=44764"						 						'(*)반품지(???) (API_TEST에서 따옴)
		strRst = strRst & "&corp_rls_pl_sn=44765"						 					'(*)출고지(???) (API_TEST에서 따옴)
		strRst = strRst & "&orpl_nm=" & chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea)	'(*)원산지
		strRst = strRst & "&mfcp_nm=" & chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)		'(*)제조사
		strRst = strRst & "&impr_nm="						 								'판매자(???)
		strRst = strRst & "&img_url=" & FbasicImage											'(*)대표이미지URL
		strRst = strRst & getLotteiMallAddImageParamToReg()									'부가이미지URL
		strRst = strRst & getLotteiMallItemContParamToReg()									'(*)상세설명
		strRst = strRst & "&md_ntc_2_FCONT="												'MD공지
		strRst = strRst & "&brnd_intro_cont=Design Your Life! 새로운 일상을 만드는 감성생활브랜드 텐바이텐"		'브랜드 설명
'2013-10-10 김진영 수정..주의사항 땜시 상품등록/수정오류 났었음
		ForderComment = replace(ForderComment,"&nbsp;"," ")
		ForderComment = replace(ForderComment,"&nbsp"," ")
		ForderComment = replace(ForderComment,"&"," ")
		ForderComment = replace(ForderComment,chr(13)," ")
		ForderComment = replace(ForderComment,chr(10)," ")
		ForderComment = replace(ForderComment,chr(9)," ")
		strRst = strRst & "&att_mtr_cont=" &URLEncodeUTF8(ForderComment)						'주의사항
		strRst = strRst & "&as_cont="															'AS정보
		strRst = strRst & "&gft_nm="															'사은품명
		strRst = strRst & "&gft_aply_strt_dtime="												'사은품시작일시
		strRst = strRst & "&gft_aply_end_dtime="												'사은품종료일시
		strRst = strRst & "&gft_fcont="															'사은품정보
		strRst = strRst & "&corp_goods_no=" & Fitemid										'업체상품번호
		strRst = strRst & "&sum_pkg_psb_yn=Y"												'합포장가능여부(자체배송만Y ,N) ==> 우선은 Y로..
	    strRst = strRst & getLotteiMallItemInfoCdToReg()   ''진영
		getLotteiMallItemRegParameter = strRst
	End Function

    Public Function getLotteiMallItemEditParameter2()
		Dim strRst
		strRst = getLotteiMallItemRegParameter(true)
		getLotteiMallItemEditParameter2 = strRst
    End Function

	'// 상품수정 파라메터 생성
	Public Function getLotteiMallItemEditParameter()
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)사용자인증키
		strRst = strRst & "&goods_no=" & FLtiMallGoodNo										'(*)롯데아이몰 상품번호
		strRst = strRst & "&brnd_no=" & BRAND_CODE											'(*)브랜드코드
'		strRst = strRst & "&sch_kwd_1_nm=" & getItemKeywordArray(0)							'키워드1
'		strRst = strRst & "&sch_kwd_2_nm=" & getItemKeywordArray(1)							'키워드2
'		strRst = strRst & "&sch_kwd_3_nm=" & getItemKeywordArray(2)							'키워드3
		strRst = strRst & getItemKeyword
		strRst = strRst & "&mdl_no="															'모델번호(?)
		strRst = strRst & "&pur_shp_cd=3" 													'(*)매출형태(1.직매입, 4.특정, 3.특정판매)	롯데닷컴은 2(판매분매입)로 설정되어있음..아이몰엔 2가 없는데..그래서 4로 놓긴했는데; ''3일듯: 현아 확인
'		strRst = strRst & "&tdf_sct_cd=1" 													'(*)과면세코드(1:과세, 2:면세)	'2013-11-11 18:09 김진영 수정//롯데닷컴처럼 모두 과세로 되어있던 상태 수정
		strRst = strRst & "&tdf_sct_cd="&CHKIIF(FVatInclude="N","2","1")					'(*)과면세코드(1:과세, 2:면세)
		strRst = strRst & getLotteiMallCateParamToReg()										'(*)해당 전시카테고리(MD상품군 파라매타도 넘기는 데 괜찮을지 몰겠음..매뉴얼엔 MD상품군 넘기는 파라매타가 없음..진영맘대로)
		strRst = strRst & getLotteiMallOptionParamToEdit()
		strRst = strRst & "&add_choc_tp_cd_10="												'날짜선택형옵션
		If FitemDiv = "06" Then
			strRst = strRst & "&add_choc_tp_cd_20=주문제작상품"						 		'입력형옵션
		End If

		If FitemDiv="06" or FitemDiv="16" then
			strRst = strRst & "&exch_rtgs_sct_cd=10"																					'교환/반품여부 10:불가능 / 20:가능
		Else
			strRst = strRst & "&exch_rtgs_sct_cd=20"																					'교환/반품여부 10:불가능 / 20:가능
		End If
		strRst = strRst & "&dlv_proc_tp_cd=1" 												'(*)배송유형(1:업체배송, 3:센터배송, 4:센터경유, 6:e-쿠폰배송)
		strRst = strRst & "&gift_pkg_yn=N" 													'(*)선물포장여부
		strRst = strRst & "&dlv_mean_cd=10" 												'(*)배송수단(10:택배 ,11:명절퀵배송 ,40:현장수령 ,50:DHL ,60:해외우편 ,70:일반우편 ,80:등기우편)
		strRst = strRst & getLotteiMallGoodDLVDtParams										'(*)배송상품구분 및 배송기일
		strRst = strRst & "&imps_rgn_info_val="													'배송불가지역(10:서울,수도권, 21:지방, 22:도서지역, 23:인천영종도, 30:제주) 여러개의경우:(콜론)으로 구분하여 전송 한개라도 콜론으로 전송
		strRst = strRst & "&byr_age_lmt_cd=0" 												'(*)구입자나이제한(0:전체, 19:19세이상)
		If Fitemid = "407171" or Fitemid = "788038" or Fitemid = "785541" or Fitemid = "785540" or Fitemid = "785542" or Fitemid = "679670" or Fitemid = "620503" or Fitemid = "590196" or Fitemid = "221081" Then
		strRst = strRst & "&dlv_polc_no=" & tenDlvFreeCd									'(*)배송정책번호(???) tenDlvCd는 inc_dailyAuthCheck.asp에서 정의 (API_TEST에서 따옴)
		Else
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd										'(*)배송정책번호(???) tenDlvCd는 inc_dailyAuthCheck.asp에서 정의 (API_TEST에서 따옴)
		End If
		strRst = strRst & "&corp_dlvp_sn=44764"						 						'(*)반품지(???) (API_TEST에서 따옴)
		strRst = strRst & "&corp_rls_pl_sn=44765"						 					'(*)출고지(???) (API_TEST에서 따옴)
		strRst = strRst & "&orpl_nm=" & chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea)	'(*)원산지
		strRst = strRst & "&mfcp_nm=" & chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)		'(*)제조사
		strRst = strRst & "&impr_nm="						 									'판매자(???)
		strRst = strRst & "&img_url=" & FbasicImage											'(*)대표이미지URL
		strRst = strRst & getLotteiMallAddImageParamToReg()									'부가이미지URL
		strRst = strRst & getLotteiMallItemContParamToReg()									'(*)상세설명
		strRst = strRst & "&md_ntc_2_FCONT="													'MD공지
		strRst = strRst & "&brnd_intro_cont=Design Your Life! 새로운 일상을 만드는 감성생활브랜드 텐바이텐"		'브랜드 설명
'2013-10-10 김진영 수정..주의사항 땜시 상품등록/수정오류 났었음
		ForderComment = replace(ForderComment,"&nbsp;"," ")
		ForderComment = replace(ForderComment,"&nbsp"," ")
		ForderComment = replace(ForderComment,"&"," ")
		ForderComment = replace(ForderComment,chr(13)," ")
		ForderComment = replace(ForderComment,chr(10)," ")
		ForderComment = replace(ForderComment,chr(9)," ")
		strRst = strRst & "&attd_mtr_cont=" &URLEncodeUTF8(ForderComment)						'주의사항
		strRst = strRst & "&as_cont="															'AS정보
		strRst = strRst & "&gft_nm="															'사은품명
		strRst = strRst & "&gft_aply_strt_dtime="												'사은품시작일시
		strRst = strRst & "&gft_aply_end_dtime="												'사은품종료일시
		strRst = strRst & "&gft_fcont="															'사은품정보
		strRst = strRst & "&corp_goods_no=" & Fitemid										'업체상품번호
		strRst = strRst & "&sum_pkg_psb_yn=Y"												'합포장가능여부(자체배송만Y ,N) ==> 우선은 Y로..
	    strRst = strRst & getLotteiMallItemInfoCdToReg()   ''진영
		'결과 반환
		getLotteiMallItemEditParameter = strRst
	End Function

	Public Function getLotteiMallAddOptParameter(nm, dc)
		Dim strRst
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)사용자인증키
		strRst = strRst & "&goods_no=" & FLtiMallGoodNo										'(*)롯데아이몰 상품번호
		strRst = strRst & "&opt_nm=" & nm													'(*)롯데아이몰 추가할 옵션명
		strRst = strRst & "&item_nm=" & dc													'(*)롯데아이몰 추가할 옵션종류명
		getLotteiMallAddOptParameter = strRst
	End Function

	Public Function getLotteiMallOptionParamToEdit()
		Dim ret : ret = ""
		Dim i
		Dim strSql, arrRows, iErrStr
		Dim isOptionExists
		Dim mayOptionCnt : mayOptionCnt = 0
		Dim item_sale_stat_cd,outmalloptcode, optLimit
		Dim item_noStr, item_sale_stat_cdStr, inv_qtyStr, optDanPoomCD, corp_item_no
		Dim optValidExists : optValidExists = false
		Dim preMaxOutmalloptcode : preMaxOutmalloptcode=-1

		strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_ltimall '"&CMallName&"'," & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
		    arrRows = rsget.getRows
		End If
		rsget.close

		isOptionExists = isArray(arrRows)
		If (isOptionExists) Then
		    mayOptionCnt = UBound(ArrRows,2)
		    mayOptionCnt = mayOptionCnt + 1
		End If

		If (FregedOptCnt <> mayOptionCnt) Then
		    rw "FregedOptCnt="&FregedOptCnt&".."&"mayOptionCnt="&mayOptionCnt
		    CALL LtiMallOneItemCheckStock(Fitemid,iErrStr)
		End If

		ret = ""
		If (Not isOptionExists) Then										'단일상품인 경우
		    rw "getLimitLotteEa="&getLimitLotteEa
		    If (FLimitYn="Y") Then
			    ret = ret & "&inv_mgmt_yn=Y"
			    ret = ret & "&inv_qty="&getLimitLotteEa()
    		    ret = ret & "&item_no=0"
    		    ret = ret & "&item_sale_stat_cd=10"
			Else
				ret = ret & "&inv_mgmt_yn=Y"
				ret = ret & "&inv_qty="&CDEFALUT_STOCK
    		    ret = ret & "&item_no=0"
    		    ret = ret & "&item_sale_stat_cd=10"
			END IF
		Else																'옵션이 있는 경우
		    ''ret = ret&"&item_mgmt_yn=Y"
		    If FLimitYn="Y" Then
			    ret = ret&"&inv_mgmt_yn=Y"
			Else
			    ret = ret&"&inv_mgmt_yn=Y"
		    End If

		    For i = 0 To UBound(ArrRows, 2)
		        if (ArrRows(11,i)=1) then ''기등록옵션만 돌림
    		        item_sale_stat_cd = "10"									''10:판매진행,20:품절,30:판매종료
    			    outmalloptcode = ArrRows(15,i)
    			    If IsNULL(outmalloptcode) or outmalloptcode = "" Then
    			        outmalloptcode = preMaxOutmalloptcode + 1
    			    Else
    			        If (preMaxOutmalloptcode > outmalloptcode) then
    			            preMaxOutmalloptcode = preMaxOutmalloptcode
    			        Else
    			            preMaxOutmalloptcode = outmalloptcode
    			        End If
    			    End If

    				If FLimitYn = "Y" Then
    					If ArrRows(4,i)-5 > 100 Then							'2013-07-04 김진영 수정..한정상품이라도 수량이 100개가 넘는다면 CDEFALUT_STOCK로 고정
    						optLimit = CDEFALUT_STOCK
    					Else
    				    	optLimit = ArrRows(4,i)-5
    					End If
    				Else
    				    optLimit = CDEFALUT_STOCK
    				End If

    				If (optLimit < 1) then optLimit = 0
    				If (ArrRows(6,i) = "N") or (ArrRows(7,i) = "N") Then item_sale_stat_cd="20"
    				If (FLimitYn = "Y") and (optLimit < 1) Then item_sale_stat_cd="20"

    				If ((ArrRows(11,i)="1") and (ArrRows(12,i)="1")) or (ArrRows(13,i)="1") Then
    				    optLimit=0
    				    item_sale_stat_cd="20"
    				End If

    				item_noStr = item_noStr & "&item_no="&outmalloptcode
    				item_sale_stat_cdStr = item_sale_stat_cdStr & "&item_sale_stat_cd="&item_sale_stat_cd
    				inv_qtyStr = inv_qtyStr & "&inv_qty="&optLimit
    				optDanPoomCD = FItemid&"_"&ArrRows(1,i)
    				corp_item_no = corp_item_no & "&corp_item_no="&optDanPoomCD
    				If (item_sale_stat_cd = "10") Then optValidExists = TRUE
    			end if
		    Next
		    ret = ret&item_noStr&item_sale_stat_cdStr&inv_qtyStr&corp_item_no
		End If

		If (Not isOptionExists) Then   ''옵션이 없으면.
			If getLTiMallSellYn = "Y" Then											'판매상태			(*:10:판매,20:품절)
				ret = ret & "&sale_stat_cd=10"
			Else
			    FSellyn="N"
				ret = ret & "&sale_stat_cd=20"
			End If
		Else
		    If (optValidExists) and (getLTiMallSellYn = "Y") Then					''판매중 이고 옵션 판매가능이면.
		        ret = ret & "&sale_stat_cd=10"
		    Else
		        rw "None Exists Valid Option"
		        FSellyn="N"
		        ret = ret & "&sale_stat_cd=20"
		    End If
		End if
		getLotteiMallOptionParamToEdit = ret
	End Function

	'// 상품등록: MD상품군 및 전시 카테고리 파라메터 생성(상품등록용)
	Public Function getLotteiMallCateParamToReg()
		Dim strSql, strRst, i, ogrpCode
		strSql = ""
		strSql = strSql & " SELECT TOP 6 c.groupCode, m.dispNo, c.disptpcd "
		strSql = strSql & " FROM db_item.dbo.tbl_lotteimall_cate_mapping as m "
		strSql = strSql & " INNER JOIN db_temp.dbo.tbl_lotteimall_Category as c on m.DispNO = c.DispNO "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
	    strSql = strSql & " and c.isusing='Y'"
		strSql = strSql & " ORDER BY disptpcd ASC "           ''''//일반몰을 기본 카테고리로..
		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) Then
		    ogrpCode = rsget("groupCode")
			strRst = "&md_gsgr_no=" & ogrpCode
			i = 0
			Do until rsget.EOF
				If (rsget("disptpcd")="10") then
				    strRst = strRst & "&disp_no=" & rsget("dispNo")			'기본 전시카테고리
				Else
				    IF (ogrpCode=rsget("groupCode")) then
					    strRst = strRst & "&disp_no_b=" & rsget("dispNo") 	'추가 전시카테고리
					End IF
			    End If
				rsget.MoveNext
				i = i + 1
			Loop
		End If
		rsget.Close
		getLotteiMallCateParamToReg = strRst
	End Function

	'//전시 카테고리 파라메터 수정(상품수정용)
	Public Function getLotteiMallCateParamToEdit()
		Dim strSql, strRst, i, ogrpCode
		strRst = "subscriptionId=" & ltiMallAuthNo											'(*)사용자인증키
		strRst = strRst & "&goods_no=" & FLtiMallGoodNo										'(*)롯데아이몰 상품번호

		strSql = ""
		strSql = strSql & " SELECT TOP 6 c.groupCode, m.dispNo, c.disptpcd "
		strSql = strSql & " FROM db_item.dbo.tbl_lotteimall_cate_mapping as m "
		strSql = strSql & " INNER JOIN db_temp.dbo.tbl_lotteimall_Category as c on m.DispNO = c.DispNO "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
	    strSql = strSql & " and c.isusing='Y'"
		strSql = strSql & " ORDER BY disptpcd ASC "           ''''//일반몰을 기본 카테고리로..
		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) Then
		    ogrpCode = rsget("groupCode")
			i = 0
			Do until rsget.EOF
				If (rsget("disptpcd")="10") then
				    strRst = strRst & "&disp_no=" & rsget("dispNo")			'기본 전시카테고리
				Else
				    IF (ogrpCode=rsget("groupCode")) then
					    strRst = strRst & "&disp_no_b=" & rsget("dispNo") 	'추가 전시카테고리
					End IF
			    End If
				rsget.MoveNext
				i = i + 1
			Loop
		End If
		rsget.Close
		strRst = strRst & "&chg_aft_fcont=전시카테고리변경"									'(*)변경사유
		getLotteiMallCateParamToEdit = strRst
	End Function


	'// 상품등록: 옵션 파라메터 생성(상품등록용)
	Public function getLotteiMallOptionParamToReg()
		dim strSql, strRst, i, optYn, optNm, optDc, chkMultiOpt, optLimit, optDanPoomCD
		chkMultiOpt = false
		optYn = "N"
		If FoptionCnt > 0 Then
			'// 이중옵션일 때
			'#옵션명 생성
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			optNm = ""
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				optYn = "Y"
				Do until rsget.EOF
					optNm = optNm & Replace(db2Html(rsget("optionTypeName")),":","")
					rsget.MoveNext
					If Not(rsget.EOF) then optNm = optNm & ":"
				Loop
			end if
			rsget.Close

			'#옵션내용 생성
			If chkMultiOpt Then
				strSql = ""
				strSql = strSql & " SELECT optionname, (optlimitno-optlimitsold) as optLimit, itemoption, itemid "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " wHERE itemid = " & FItemid
				strSql = strSql & " and isUsing = 'Y' and optsellyn = 'Y' "
				strSql = strSql & " and optaddprice = 0 "
				'''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) " ''일단 입력
				rsget.Open strSql,dbget,1

				optDc = ""
				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit-5
					    optDanPoomCD = rsget("itemid")&"_"&rsget("itemoption")
					    If (optLimit < 1) Then optLimit = 0
					    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
						optDc = optDc & Replace(Replace(db2Html(rsget("optionname")),":",""),"'","") & "," & optLimit & "," & optDanPoomCD
						rsget.MoveNext
						If Not(rsget.EOF) Then optDc = optDc & ":"
					Loop
				End If
				rsget.Close
			End If

			'// 단일옵션일 때
			If Not(chkMultiOpt) Then
				strSql = ""
				strSql = strSql & " SELECT optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, itemoption, itemid "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid = " & FItemid
				strSql = strSql & " and isUsing = 'Y' and optsellyn = 'Y' "
				strSql = strSql & " and optaddprice = 0 "
				''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				If Not(rsget.EOF or rsget.BOF) then
					optYn = "Y"
					If db2Html(rsget("optionTypeName")) <> "" Then
						optNm = Replace(db2Html(rsget("optionTypeName")),":","")
					Else
						optNm = "옵션"
					End If
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit - 5
					    optDanPoomCD = rsget("itemid")&"_"&rsget("itemoption")
					    If (optLimit < 1) Then optLimit = 0
					    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK   ''2013/06/12 재고관리여부 모두 Y로 변경 되므로

						optDc = optDc & Replace(Replace(Replace(db2Html(rsget("optionname")),":",""),",",""),"'","") & "," & optLimit & "," & optDanPoomCD
						rsget.MoveNext
						If Not(rsget.EOF) Then optDc = optDc & ":"
					Loop
				End If
				rsget.Close
			End If
		End If
		strRst = strRst & "&item_mgmt_yn=" & optYn						'단품관리여부(옵션)
		strRst = strRst & "&opt_nm=" & optNm							'옵션명
		strRst = strRst & "&item_list=" & optDc							'옵션상세
		getLotteiMallOptionParamToReg = strRst
	End Function

	Public Function getLotteiMallGoodDLVDtParams()
		dim strRst
		strRst = ""
		If (FtenCateLarge="055") or (FtenCateLarge="040") then ''가구/패브릭 15일로
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=15"
		ElseIf (FtenCateLarge="080") or (FtenCateLarge="100") then  ''우먼/베이비 5일
			strRst = strRst & "&dlv_goods_sct_cd=03" 																						'배송상품구분		(*:주문제작03)
			strRst = strRst & "&dlv_dday=5"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="001" or FtenCateMid="004")) then  ''수납/생활> 옷/이불수납 or 주방수납 10일 - 현아씨요청 2013/01/22
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="025") and (FtenCateMid="107")) then  ''디지털 > 기타 스마트기기 케이스  10일 - 현아씨요청 2013/01/22
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="050") and (FtenCateMid="777")) then   ''홈/데코 > 거울   - 미희씨요청 2013/03/08
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="001")) then    ''HOME > 수납/생활 > 보관/정리용품 > 수납장 			주문제작15일 045&cdm=002&cds=001
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=15"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="002")) then    ''HOME > 수납/생활 > 보관/정리용품 > 틈새수납장			주문제작10일
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="005")) then    ''HOME > 수납/생활 > 보관/정리용품 > 잡지꽂이 			주문제작10일
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="001")) then    ''HOME > 수납/생활 > 데코수납 > 우드박스 				주문제작10일
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="007")) then    ''HOME > 수납/생활 > 데코수납 > 인터폰박스 			               주문제작10일
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="050") and (FtenCateMid="060") and (FtenCateSmall="070")) then    ''HOME > 홈/데코 > 소품박스/바구니 > 인터폰박스			주문제작10일 cdl=050&cdm=060&cds=070
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="110") and (FtenCateMid="090") and (FtenCateSmall="040")) then    ''HOME > 감성채널 > DIY > 나무로만들기 				주문제작10일 110&cdm=090&cds=040
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
		ElseIf ((FtenCateLarge="045") and (FtenCateMid="010")) then   ''수납/생활 > 디자인선반  - 미희씨요청 2013/03/08
			strRst = strRst & "&dlv_goods_sct_cd=03"
			strRst = strRst & "&dlv_dday=10"
'		ElseIf (FtenCateLarge="025")  then  ''디지털 10일 - 미희씨요청 2013/01/17
'		    strRst = strRst & "&dlv_goods_sct_cd=03" 																						'배송상품구분		(*:주문제작03)
'		    strRst = strRst & "&dlv_dday=10"
		ElseIf ((FitemDiv="06") or (FitemDiv="16")) then    ''주문(후)제작상품
			strRst = strRst & "&dlv_goods_sct_cd=03"
			If (FrequireMakeDay>7) then
				    strRst = strRst & "&dlv_dday="&CStr(FrequireMakeDay)
			ElseIf (FrequireMakeDay<1) then
				    strRst = strRst & "&dlv_dday=7"
			Else
				    strRst = strRst & "&dlv_dday="&(FrequireMakeDay+1)
			End If
		Else
			strRst = strRst & "&dlv_goods_sct_cd=01" 																						'배송상품구분		(*:일반상품)
			strRst = strRst & "&dlv_dday=3" 																								'배송기일			(*:3일이내)
		End If
		getLotteiMallGoodDLVDtParams = strRst
	End Function

	'// 상품등록: 상품추가이미지 파라메터 생성(상품등록용)
	Public Function getLotteiMallAddImageParamToReg()
		Dim strRst, strSQL, i
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget("imgType")="0" then
					strRst = strRst & "&img_url" & i & "=" & Server.URLEncode("http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400"))
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next
		End If
		rsget.Close
		getLotteiMallAddImageParamToReg = strRst
	End Function

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getLotteiMallItemContParamToReg()
		Dim strRst, strSQL, strtextVal
		strRst = Server.URLEncode("<div align=""center"">")
		'2014-01-17 10:00 김진영 탑 이미지 추가
		strRst = strRst & Server.URLEncode("<p><a href=""http://www.lotteimall.com/display/viewDispShop.lotte?disp_no=5100455"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_Ltimall.jpg""></a></p><br>")
		'#기본 상품설명
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,"&nbsp;"," ")
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,"&nbsp"," ")
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,"&"," ")
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,chr(13)," ")
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,chr(10)," ")
		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,chr(9)," ")
''		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,"""","&quot;")
''		oiMall.FItemList(i).Fitemcontent = replace(oiMall.FItemList(i).Fitemcontent,"'","&#39;")
'%BE%C8%B3%E7
'%EC%95%88%EB%85%95
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & URLEncodeUTF8(oiMall.FItemList(i).Fitemcontent & "<br>")
				'strRst = strRst & (oiMall.FItemList(i).Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & URLEncodeUTF8(oiMall.FItemList(i).Fitemcontent & "<br>")
				'strRst = strRst & (oiMall.FItemList(i).Fitemcontent & "<br>")
			Case Else
				strRst = strRst & URLEncodeUTF8(oiMall.FItemList(i).Fitemcontent & "<br>")
				'strRst = strRst & (ReplaceBracket(oiMall.FItemList(i).Fitemcontent) & "<br>")
		End Select

		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
				If rsget("imgType") = "1" Then
					strRst = strRst & Server.URLEncode("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage & """ border=""0""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage2 & """ border=""0""><br>")

		'#배송 주의사항
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_LTimall.jpg"">")
		strRst = strRst & Server.URLEncode("</div>")
		getLotteiMallItemContParamToReg = "&dtl_info_fcont=" & strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = Server.URLEncode(rsget("textVal"))
			strRst = Server.URLEncode("<div align=""center""><p><a href=""http://www.lotteimall.com/display/viewDispShop.lotte?disp_no=5100455"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_Ltimall.jpg""></a></p><br>") & strtextVal & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_LTimall.jpg""></div>")
			getLotteiMallItemContParamToReg = "&dtl_info_fcont=" & strRst
		End If
		rsget.Close
	End Function

	Public Function getLotteiMallItemInfoCdToReg()
		Dim anjunInfo
        ''안전인증정보(애매함)
		If (Fsafetyyn="Y" and FsafetyDiv<>0) Then
			If (FsafetyDiv=10) Then											'국가통합인증(KC마크)
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=31"					'KS인증
				anjunInfo = anjunInfo & "&sft_cert_org_cd=31"					'한국표준협회
			Elseif (FsafetyDiv=20) Then										'전기용품 안전인증
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=21"					'전기용품안전인증
				anjunInfo = anjunInfo & "&sft_cert_org_cd=21"					'한국전기전자시험연구원
			Elseif (FsafetyDiv=30) Then										'KPS 안전인증 표시
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=21"					'전기용품안전인증
				anjunInfo = anjunInfo & "&sft_cert_org_cd=21"					'한국전기전자시험연구원
			Elseif (FsafetyDiv=40) Then										'KPS 자율안전 확인 표시
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=22"					'전기용품자율안전확인신고
				anjunInfo = anjunInfo & "&sft_cert_org_cd=22"					'한국전자파연구원
			Elseif (FsafetyDiv=50) Then										'KPS 어린이 보호포장 표시
				anjunInfo = anjunInfo & "&sft_cert_sct_cd=31"					'KS인증
				anjunInfo = anjunInfo & "&sft_cert_org_cd=31"					'한국표준협회
			Else
				anjunInfo = ""
			End if
			anjunInfo = anjunInfo & "&sft_cert_no="&Server.URLEncode(FsafetyNum)
		End If

		Dim strRst, strSQL
		Dim mallinfoDiv,mallinfoCd,infoContent, mallinfoCdAll, bufTxt

		'동일모델의 출시년월 뽑는 쿼리
		Dim YM, ConvertYM, SD
		strSQL = ""
		strSQL = strSQL & " SELECT top 1 F.infocontent, IC.safetyDiv " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv=M.mallinfoDiv  " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd AND F.itemid='"&Fitemid&"' " & vbcrlf
		strSQL = strSQL & " where IC.itemid='"&Fitemid&"' and M.mallinfocd = '10011' " & vbcrlf
		rsget.Open strSql,dbget,1

		If Not(rsget.EOF or rsget.BOF) then
			YM = rsget("infocontent")
			SD = rsget("safetyDiv")
		Else
			YM = "X"
			SD = "X"
		End If
		rsget.Close

		If YM <> "X" Then
		    YM = replace(YM,".","")
		    YM = replace(YM,"/","")
		    YM = replace(YM,"-","")
		    YM = replace(YM," ","")
		    YM = TRIM(YM)

			If isNumeric(Ym) Then
				ConvertYM = Clng(YM)
			Else
				ConvertYM = "X"
			End If
		Else
			ConvertYM = YM
		End If

		strSQL = ""
		strSQL = strSQL & " SELECT TOP 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE " & vbcrlf

		If SD = "10" Then
			'출시년월의 값이 없는 경우
			If ConvertYM = "X" Then
				strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN IC.safetyNum " & vbcrlf
			'출시년월의 값이 있는 경우
			Else
				'출시년월이 2012년 7월 이전인 경우
				If ConvertYM < 201207 Then
					strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN '해당없음' " & vbcrlf	 '(맵핑코드가 KCC인증이고), (10x10에서 안전인증코드여부가 Y, 구분이 KC(10), 201207전)일 때
				'출시년월이 2012년 7월 이후인 경우
				ElseIf ConvertYM >= 201207 Then
					strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') AND (LEFT(IC.safetyDiv,3)='KCC') THEN IC.safetyNum " & vbcrlf '(맵핑코드가 KCC인증이고), (10x10에서 안전인증코드여부가 Y, 구분이 KC(10), 201207후)일 때
				End If
			End If
		End If
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10063') THEN IC.safetyNum " & vbcrlf		'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10063') THEN '해당없음'  " & vbcrlf		'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10205') THEN IC.safetyNum " & vbcrlf		'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10205') THEN '해당없음'  " & vbcrlf		'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') = 'Y') AND (M.mallinfoCd= '10206') THEN 'KC 안전인증 필'  " & vbcrlf	'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'') <> 'Y') AND (M.mallinfoCd= '10206') THEN '해당없음'  " & vbcrlf		'10206은 KC인증
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn= 'N') THEN '해당없음'  " & vbcrlf		'(맵핑코드가 KCC인증이고), (10x10에서 안전인증코드여부가 N)일 때
		strSQL = strSQL & " 	 WHEN M.infoCd='00001' THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN M.infoCd='00002' THEN '원산지와 동일' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='19008') THEN '제공함' " & vbcrlf				'귀금속의 가공지
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='19008') THEN '제공하지 않음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='18008') THEN '기능성 심사 필' " & vbcrlf		'화장품의 기능성 화장품 여부
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='18008') THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN F.chkDiv='Y' AND (M.infoCd='17008') THEN '식품위생법에 따른 수입신고필함' " & vbcrlf		'식품위생법에 따른 수입신고 여부	20130215진영 추가
		strSQL = strSQL & " 	 WHEN F.chkDiv='N' AND (M.infoCd='17008') THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='M' THEN replace(F.infocontent,'.','') " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='C' AND F.chkDiv='N' THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN replace(F.infocontent,'1644-6030','1644-6035') " & vbcrlf
		strSQL = strSQL & " 	 WHEN c.infoCd='02004' and F.infocontent='' then '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 ELSE F.infocontent " & vbcrlf
		strSQL = strSQL & " END AS infoContent, L.shortVal " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON '1'+IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd AND F.itemid='"&Fitemid&"' " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_OutMall_etcLink as L on L.mallid = M.mallid and L.linkgbn='infoDiv21Lotte' and L.itemid ='"&FItemid&"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'lotteimall' AND IC.itemid='"&Fitemid&"' " & vbcrlf
		strSQL = strSQL & " ORDER BY M.infocd ASC"
		rsget.Open strSQL,dbget,1
		Dim mat_name, mat_percent, mat_place, material

		If Not(rsget.EOF or rsget.BOF) then
			mallinfoDiv = "&ec_goods_artc_cd="&Server.URLEncode(rsget("mallinfoDiv"))						'상품품목코드
			Do until rsget.EOF
				mallinfoCd = rsget("mallinfoCd")
				infoContent  = rsget("infoContent")
				infoContent  = replace(infoContent,"%", "％")
				infoContent  = replace(infoContent,chr(13), "")
				infoContent  = replace(infoContent,chr(10), "")
				infoContent  = replace(infoContent,chr(9), " ")
				If mallinfoCd="10085" Then
					If isNull(rsget("shortVal")) = FALSE Then
						material = Split(rsget("shortVal"),"!!^^")
						mat_name	= material(0)
						mat_percent	= material(1)
						mat_place	= material(2)

						bufTxt = "&mmtr_nm="&mat_name														'주원료명
						bufTxt = bufTxt&"&cmps_rt="&mat_percent												'함량
						bufTxt = bufTxt&"&mmtr_orpl_nm="&mat_place											'원료원산지
					End If
				End If
				mallinfoCdAll = mallinfoCdAll & "&"&mallinfoCd&"=" &infoContent								'상품품목별 항목정보
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		strRst = anjunInfo & mallinfoDiv & mallinfoCdAll & bufTxt
		getLotteiMallItemInfoCdToReg = strRst
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end class

class CLotteiMall
	public FItemList()

	public FResultCount
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectMdCode
	public FRectDspNo
	public FRectIsMapping

	public FRectSDiv
	public FRectKeyword
	public FRectOrderby
	public FRectGrpCode

	public FRectCDL
	public FRectCDM
	public FRectCDS

    public FRectMode

	public FRectItemID
	public FRectItemName
	public FRectMakerid
	public FRectLotteNotReg
	public FRectMatchCate
	''public FRectMatchCateNotCheck
	public FRectSellYn
	public FRectLimitYn
	public FRectSailYn
	public FRectLTiMallGoodNo
	public FRectLTiMallTmpGoodNo
	public FRectMinusMigin
	public FRectonlyValidMargin
	Public FRectStartMargin
	Public FRectEndMargin
	public FRectIsSoldOut
	public FRectExpensive10x10
	public FRectLotteYes10x10No
	public FRectLotteNo10x10Yes
	public FRectOnreginotmapping
	public FRectNotJehyu
	public FRectEventid
	public FRectdiffPrc
	public FRectdisptpcd
    public FRectCateUsingYn

	Public FRectExtNotReg
	Public FRectIsReged
	Public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectExcTrans
	Public FRectPriceOption
	Public FRectIsOption
	Public FRectLtimallYes10x10No
	Public FRectReqEdit
	Public FRectPurchasetype
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectIsextusing
	Public FRectCisextusing
	Public FRectRctsellcnt
	Public FRectLtimallNo10x10Yes

    ''정렬순서
    public FRectOrdType
    public FRectoptAddprcExists
    public FRectoptAddPrcRegTypeNone
    public FRectoptAddprcExistsExcept
    public FRectoptExists
    public FRectoptnotExists
    public FRectregedOptNull

    public FRectFailCntExists
    public FRectFailCntOverExcept
    public FRect10000_Over
    public FRectExtSellYn
    public FRectInfoDiv
	Public FRectisMadeHand
	Public FRectIsSpecialPrice

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'--------------------------------------------------------------------------------

    '// 담당MD목록
	Public Sub getLotte_MDList
		Dim sqlStr,i
		sqlStr = " select count(MDCode) as cnt, CEILING(CAST(Count(MDCode) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + "From db_temp.dbo.tbl_lotteiMall_MDInfo "
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotteiMall_MDInfo "
		sqlStr = sqlStr + " order by MDCode asc"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				set FItemList(i) = new CLotteiMallItem
					FItemList(i).FMDCode		= rsget("MDCode")
					FItemList(i).FMDName		= db2html(rsget("MDName"))
					FItemList(i).FSellFeeType	= rsget("SellFeeType")
					FItemList(i).FNormalSellFee	= rsget("NormalSellFee")
					FItemList(i).FEventSellFee	= rsget("EventSellFee")
					FItemList(i).FisUsing		= rsget("isUsing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    '// 담당MD상품군 목록
    Public Sub getLotte_MDGrpList
		Dim sqlStr, i
		sqlStr = " select count(groupCode) as cnt, CEILING(CAST(Count(groupCode) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " From db_temp.dbo.tbl_lotteiMall_MDCateGrp "
		sqlStr = sqlStr + " Where MDCode='" & FRectMdCode & "'"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotteiMall_MDCateGrp "
		sqlStr = sqlStr + " Where MDCode='" & FRectMdCode & "'"
		sqlStr = sqlStr + " order by groupCode asc"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CLotteiMallItem
					FItemList(i).FgroupCode			= rsget("groupCode")
					FItemList(i).FSuperGroupName	= db2html(rsget("SuperGroupName"))
					FItemList(i).FGroupName			= db2html(rsget("GroupName"))
					FItemList(i).FisUsing			= rsget("isUsing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'--------------------------------------------------------------------------------
	'// 텐바이텐-롯데아이몰 카테고리 :: 전문몰 카테고리가 매핑 되어야 함..
	Public Sub getTenLotteimallCateList
		Dim sqlStr, addSql, i, odySql

		If FRectCDL <> "" Then
			addSql = addSql & " and s.code_large='" & FRectCDL & "'"
		End If

		If FRectCDM <> "" then
			addSql = addSql & " and s.code_mid='" & FRectCDM & "'"
		End If

		If FRectCDS <> "" then
			addSql = addSql & " and s.code_small='" & FRectCDS & "'"
		End If

		if FRectDspNo <> "" then
			addSql = addSql & " and cm.dispNo='" & FRectDspNo & "'"
		end if

		If FRectIsMapping = "Y" then
			addSql = addSql & " and cm.DispNo is Not null "
		ElseIf FRectIsMapping = "N" then
			addSql = addSql & " and cm.DispNo is null "
		End If

		If FRectKeyword <> "" Then
			Select Case FRectSDiv
				Case "LCD"	'롯데아이몰 전시코드 검색
					addSql = addSql & " and cm.DispNo='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End If

		If FRectOrderby <> "" Then
			Select Case FRectOrderby
				Case "1"	'카테고리순
					odySql = odySql & " ORDER BY s.code_large, s.code_mid, s.code_small, disptpcd desc "
				Case "2"	'상품수
					odySql = odySql & " ORDER BY W.itemcnt DESC, s.code_large,s.code_mid,s.code_small ASC, disptpcd desc "
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_lotteimall_cate_mapping as cm on cm.tenCateLarge = s.code_large and cm.tenCateMid = s.code_mid and cm.tenCateSmall = s.code_small "
		If FRectdisptpcd <> "" Then
			sqlStr = sqlStr & " JOIN db_temp.dbo.tbl_lotteimall_Category as lc on lc.DispNo = cm.DispNo and lc.disptpcd='" & FRectdisptpcd &"'"
	    Else
			sqlStr = sqlStr & " LEFT JOIN db_temp.dbo.tbl_lotteimall_Category as lc on lc.DispNo = cm.DispNo "
        End If
		sqlStr = sqlStr & " Where 1 = 1 " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT cate_large, cate_mid, cate_small, count(*) as itemcnt "
		sqlStr = sqlStr & " INTO #categoryTBL "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item "
		sqlStr = sqlStr & " WHERE isusing = 'Y' and sellyn = 'Y' "
		sqlStr = sqlStr & " group by cate_large, cate_mid, cate_small "
		dbget.execute sqlStr

		sqlStr = ""
		sqlStr = sqlStr & " select top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " s.code_large, s.code_mid, s.code_small "
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large = s.code_large) as large_nm "
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large = s.code_large and code_mid = s.code_mid) as mid_nm "
		sqlStr = sqlStr & " ,code_nm as small_nm "
		sqlStr = sqlStr & " ,cm.DispNo, lc.DispNm, lc.DispLrgNm, lc.DispMidNm, lc.DispSmlNm, lc.DispThnNm, lc.groupCode, lc.disptpcd "
		sqlStr = sqlStr & " ,lc.isusing, W.itemcnt"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s "
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_lotteimall_cate_mapping as cm on cm.tenCateLarge = s.code_large and cm.tenCateMid = s.code_mid and cm.tenCateSmall = s.code_small "
		If FRectdisptpcd <> "" Then
			sqlStr = sqlStr & " JOIN db_temp.dbo.tbl_lotteimall_Category as lc on lc.DispNo = cm.DispNo and lc.disptpcd='" & FRectdisptpcd &"'"
	    Else
			sqlStr = sqlStr & " LEFT JOIN db_temp.dbo.tbl_lotteimall_Category as lc on lc.DispNo = cm.DispNo "
        End If
		If FRectdisptpcd <> "" Then
            sqlStr = sqlStr & " and lc.disptpcd='" & FRectdisptpcd &"'"
        End If
		sqlStr = sqlStr & " LEFT JOIN #categoryTBL as W on W.cate_large = s.code_large and s.code_mid = W.cate_mid and s.code_small = W.cate_small  " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & odySql
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CLotteiMallItem
					FItemList(i).FtenCDLName	= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName	= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName	= db2html(rsget("small_nm"))
					FItemList(i).FDispNo		= rsget("DispNo")
					FItemList(i).FDispNm		= db2html(rsget("DispNm"))
					FItemList(i).FtenCateLarge	= rsget("code_large")
					FItemList(i).FtenCateMid	= rsget("code_mid")
					FItemList(i).FtenCateSmall	= rsget("code_small")
					FItemList(i).FgroupCode		= rsget("groupCode")
					FItemList(i).FDispLrgNm		= db2html(rsget("DispLrgNm"))
					FItemList(i).FDispMidNm		= db2html(rsget("DispMidNm"))
					FItemList(i).FDispSmlNm		= db2html(rsget("DispSmlNm"))
					FItemList(i).FDispThnNm		= db2html(rsget("DispThnNm"))
	                FItemList(i).Fdisptpcd      = rsget("disptpcd")
	                FItemList(i).FCateisusing   = rsget("isusing")
					FItemList(i).FItemcnt		= rsget("itemcnt")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// 롯데아이몰 카테고리
	Public Sub getLTiMallCategoryList
		Dim sqlStr, addSql, i

		If FRectDspNo <> "" Then
			addSql = addSql & " and c.dispNo=" & FRectDspNo
		End If

		If FRectGrpCode <> "" Then
			addSql = addSql & " and c.groupCode=" & FRectGrpCode
		End If

        If FRectdisptpcd <> "" Then
            addSql = addSql & " and c.disptpcd='" & FRectdisptpcd &"'"
        End If

		If FRectKeyword <> "" Then
			Select Case FRectSDiv
				Case "LCD"	'롯데아이몰 전시코드 검색
					addSql = addSql & " and c.DispNo='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(롯데아이몰 세분류명)
					addSql = addSql & " and ((c.dispNm like '%" & FRectKeyword & "%') or (c.dispsmlNm like '%" & FRectKeyword & "%'))"
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(c.DispNo) as cnt, CEILING(CAST(Count(c.DispNo) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_lotteimall_Category as c "
		sqlStr = sqlStr & " WHERE c.DispMidNm not like '%1300K%' and dispLrgNm not like '%1300K%' " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT distinct top " & CStr(FPageSize*FCurrPage) & " c.* "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_lotteimall_Category as c "
		sqlStr = sqlStr & " WHERE c.DispMidNm not like '%1300K%' and dispLrgNm not like '%1300K%' " & addSql
		sqlStr = sqlStr & " ORDER BY c.DispLrgNm, c.DispMidNm, c.DispSmlNm, c.DispThnNm, c.DispNo "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				set FItemList(i) = new CLotteiMallItem
					FItemList(i).FDispNo		= rsget("DispNo")
					FItemList(i).FDispNm		= db2html(rsget("DispNm"))
					FItemList(i).FDispLrgNm		= db2html(rsget("DispLrgNm"))
					FItemList(i).FDispMidNm		= db2html(rsget("DispMidNm"))
					FItemList(i).FDispSmlNm		= db2html(rsget("DispSmlNm"))
					FItemList(i).FDispThnNm		= db2html(rsget("DispThnNm"))
	                FItemList(i).Fdisptpcd      = rsget("disptpcd")
					FItemList(i).FisUsing		= rsget("isUsing")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'--------------------------------------------------------------------------------

	'// 롯데닷컴 브랜드
	public Sub getLotteBrandList
		dim sqlStr, addSql, i

		if FRectMakerid<>"" then
			addSql = addSql & " and b.TenMakerid='" & FRectMakerid & "'"
		end if

		if FRectKeyword<>"" then
			Select Case FRectSDiv
				Case "LCD"	'롯데닷컴 브랜드코드 검색
					addSql = addSql & " and b.lotteBrandCD='" & FRectKeyword & "'"
				Case "TCD"	'텐바이텐 브랜드ID 검색
					addSql = addSql & " and b.TenMakerid='" & FRectKeyword & "'"
				Case "BNM"	'브랜드명(텐바이텐명)
					addSql = addSql & " and c.socname_kor like '%" & FRectKeyword & "%'"
			End Select
		end if

		sqlStr = " select count(b.TenMakerid) as cnt, CEILING(CAST(Count(b.TenMakerid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_user.dbo.tbl_user_c as c "
		sqlStr = sqlStr + " 	Join db_item.dbo.tbl_lotte_brand_mapping as b "
		sqlStr = sqlStr + " 		on c.userid=b.TenMakerid "
		sqlStr = sqlStr + " Where 1=1 " & addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " b.*, c.socname_kor "
		sqlStr = sqlStr + " from db_user.dbo.tbl_user_c as c "
		sqlStr = sqlStr + " 	Join db_item.dbo.tbl_lotte_brand_mapping as b "
		sqlStr = sqlStr + " 		on c.userid=b.TenMakerid "
		sqlStr = sqlStr + " Where 1=1 " & addSql
		sqlStr = sqlStr + " order by b.regdate desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteiMallItem

				FItemList(i).FlotteBrandCd		= rsget("lotteBrandCd")
				FItemList(i).FlotteBrandName	= db2html(rsget("lotteBrandName"))
				FItemList(i).FTenMakerid		= rsget("TenMakerid")
				FItemList(i).FTenBrandName		= db2html(rsget("socname_kor"))
				FItemList(i).FisUsing			= rsget("isUsing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'--------------------------------------------------------------------------------
	'옵션추가금액 상품 리스트
	Public Sub getLTiMallAddOptionRegedItemList
		Dim sqlStr, addSql, i
		'브랜드검색
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'상품명 검색
		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End if

		'롯데아이몰 상품번호 검색
        If (FRectLtimallGoodNo <> "") then
            If Right(Trim(FRectLtimallGoodNo) ,1) = "," Then
            	FRectLtimallGoodNo = Replace(FRectLtimallGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallGoodNo in ('" & replace(Left(FRectLtimallGoodNo, Len(FRectLtimallGoodNo)-1),",","','") & "')"
            Else
				FRectLtimallGoodNo = Replace(FRectLtimallGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallGoodNo in ('" & replace(FRectLtimallGoodNo,",","','") & "')"
            End If
        End If

		'롯데아이몰 승인전 상품번호 검색
        If (FRectLtimallTmpGoodNo <> "") then
            If Right(Trim(FRectLtimallTmpGoodNo) ,1) = "," Then
            	FRectItemid = Replace(FRectLtimallTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallTmpGoodNo in ('" & replace(Left(FRectLtimallTmpGoodNo, Len(FRectLtimallTmpGoodNo)-1),",","','") & "')"
            Else
				FRectLtimallTmpGoodNo = Replace(FRectLtimallTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallTmpGoodNo in ('" & replace(FRectLtimallTmpGoodNo,",","','") & "')"
            End If
        End If

		'카테고리 검색
		If FRectCDL <> "" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if
		If FRectCDM <> "" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if
		If FRectCDS <> "" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End If

		'등록여부 검색
		Select Case FRectExtNotReg
			Case "Q"	''등록실패
				addSql = addSql & " and J.LtimallStatCd = '-1'"
			Case "J"	'등록예정이상
				addSql = addSql & " and J.LtimallStatCd >= '0'"
			Case "W"	'등록예정
				addSql = addSql & " and J.LtimallStatCd = '0'"
		    Case "A"	'전송시도중오류
				addSql = addSql & " and J.LtimallStatCd = '1'"
			Case "C"	'반려
			    addSql = addSql & " and J.LtimallStatCd = '40'"
			Case "F"	'등록완료(임시)
			    addSql = addSql & " and J.LtimallStatCd = '20'"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.LtimallStatCd = '7'"
				addSql = addSql & " and J.LtimallGoodNo is Not Null"
			Case "R"	'수정요망		'스케줄링에서 사용
				addSql = addSql & " and (J.LtimallStatCd = '7')"
				addSql = addSql & " and J.LtimallLastUpdate < i.lastupdate"
				addSql = addSql & " and isnull(J.LtimallGoodNo, '') <> '' "
		End Select

		'미등록 라디오버튼 클릭 시
		Select Case FRectIsReged
			Case "N"	'등록예정이상
			    addSql = addSql & " and J.midx is NULL  and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
		End Select

		'판매여부 검색
		Select Case FRectSellYn
			Case "Y"	addSql = addSql & " and i.sellYn='Y'"			'판매
			Case "N"	addSql = addSql & " and i.sellYn in ('S','N')"	'품절
		End Select

		'텐바이텐 한정여부 검색
		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		'텐바이텐 세일여부 검색
		If FRectSailYn <> "" Then
			addSql = addSql & " and i.sailYn = '" & FRectSailYn & "'"
		End If

		'역마진 및 마진 CMAXMARGIN 이상 검색
		If (FRectonlyValidMargin <> "") Then
			IF (FRectonlyValidMargin = "Y") Then
				addSql = addSql & " and Round(((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100,0) >= " & CMAXMARGIN & VbCrlf
			Else
				addSql = addSql & " and Round(((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100,0) < " & CMAXMARGIN & VbCrlf
			End If
		End If

		If (FRectStartMargin <> "") OR (FRectEndMargin <> "") Then
			If (FRectStartMargin <> "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and ("
				addSql = addSql & " 	convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
				addSql = addSql & " 	and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
				addSql = addSql & " ) "
			ElseIf (FRectStartMargin <> "") And (FRectEndMargin = "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
			ElseIf (FRectStartMargin = "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
			End If
		End If

		'주문제작 여부 검색
		If FRectisMadeHand <> "" Then
			If (FRectisMadeHand = "Y") Then
				addSql = addSql & " and i.itemdiv in ('06', '16')" & VbCrlf
			Else
				addSql = addSql & " and i.itemdiv not in ('06', '16')" & VbCrlf
			End If
		End if

		'옵션 여부 검색
		If FRectIsOption <> "" Then
			If FRectIsOption = "optAll" Then			'옵션전체
				addSql = addSql & " and i.optioncnt > 0"
			ElseIf FRectIsOption = "optaddpricey" Then	'추가금액Y
				addSql = addSql & " and i.optioncnt > 0"
 				addSql = addSql & " and J.optAddPrcCnt > 0"
			ElseIf FRectIsOption = "optaddpricen" Then	'추가금액N
				addSql = addSql & " and i.optioncnt > 0"
				addSql = addSql & " and isNULL(J.optAddPrcCnt,0)=0"
			ElseIf FRectIsOption = "optN" Then			'단품
				addSql = addSql & " and i.optioncnt = 0"
			ElseIf FRectIsOption = "optAddPrcRegType" Then
				addSql = addSql & " and J.optAddPrcCnt > 0"
				addSql = addSql & " and J.optAddPrcRegType = 0"
			End If
		End If

		'텐바이텐 품목정보 검색
		If (FRectInfoDiv <> "") then
			If (FRectInfoDiv = "YY") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
			ElseIf (FRectInfoDiv = "NN") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')=''"
			Else
				addSql = addSql & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End If
		End If

		'텐바이텐 등록제외 브랜드 제외 검색
		If (FRectNotinmakerid <> "") then
			If (FRectNotinmakerid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteimall') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteimall') "
			End If
		End If

		'텐바이텐 등록제외 상품 제외 검색
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteimall') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteimall') "
			End If
		End If

		'제휴몰 전송제외 상품 검색
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
				addSql = addSql & " or i.isExtUsing='N' "
				addSql = addSql & " or uc.isExtUsing='N' "
				addSql = addSql & " or i.deliveryType = 7 "
				addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				addSql = addSql & " or i.itemdiv = '21' "
				addSql = addSql & " or i.deliverfixday in ('C','X','G') "
				addSql = addSql & " or i.itemdiv >= 50 "
				addSql = addSql & " or i.itemdiv = '08' "
				addSql = addSql & " or i.itemdiv = '09' "
				addSql = addSql & " or i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or not (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " or not ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "F") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and i.itemdiv <> '21' "											'// 딜상품
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// 꽃배달, 화물배달, 해외직구
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.itemdiv <> '08' "											'// 티켓(강좌) 상품
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present상품
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and 'Y' = (CASE WHEN i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or J.accFailCnt > 0 "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "N") Then
				addSql = addSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'lotteimall') "
				addSql = addSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'lotteimall') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and i.itemdiv <> '21' "											'// 딜상품
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// 꽃배달, 화물배달, 해외직구
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.itemdiv <> '08' "											'// 티켓(강좌) 상품
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present상품
				addSql = addSql & " and i.cate_large <> '999' "										'// 카테고리 미지정
				addSql = addSql & " and i.cate_large <> '' "										'// 카테고리 미지정
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and isNULL(ct.infodiv,'') not in ('','18','20','22') "
				addSql = addSql & " and not ('1'+ct.infoDiv in ('107', '108', '109', '110', '111', '112', '113', '114', '123') and not exists (select top 1 tr.itemid from db_item.dbo.tbl_safetycert_tenReg tr where tr.itemid = i.itemid and isnull(TR.certNum, '') <> '')) "
				addSql = addSql & " and not (i.optioncnt > 0 and exists (select top 1 r.itemid from [db_item].[dbo].tbl_OutMall_regedoption R where R.mallid = 'lotteimall' and R.itemid = i.itemid and R.itemoption = '0000')) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	not exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.itemid=i.itemid and o.optAddPrice > 0) "
				addSql = addSql & " ) "
			End If
		End If

		'옵션추가금액New
		If (FRectPriceOption <> "") then
			If (FRectPriceOption = "Y") Then
				addSql = addSql & " and i.itemid in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			ElseIf (FRectPriceOption = "N") Then
				addSql = addSql & " and i.itemid not in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			End If
		End If

		'롯데아이몰 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.LtimallSellYn <> 'X'"
			Else
				addSql = addSql & " and J.LtimallSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'롯데아이몰 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(c.mapCnt, '') <> ''"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(c.mapCnt, '') = ''"
		End Select

        '롯데아이몰 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.LtimallPrice is Not Null and J.LtimallPrice < i.sellcash"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.LtimallPrice is Not Null and i.sellcash <> J.LtimallPrice "
		End If

		'롯데아이몰판매,  10x10 품절
		If (FRectLtimallYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.LtimallSellYn='Y'"
		End If

		'롯데아이몰품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectLtimallNo10x10Yes <> "" Then
			addSql = addSql & " and (J.LtimallSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.LtimallLastUpdate < i.lastupdate "
		End If

		'스케줄링에서 사용 실패횟수 제한
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.LtimallLastUpdate, 10) "
		End If

		'배송구분
		If (FRectDeliverytype <> "") Then
			addSql = addSql & " and i.deliverytype='" & FRectDeliverytype & "'"
		End If

		'거래구분
		If FRectMWDiv = "MW" Then
			addSql = addSql & " and (i.mwdiv='M' or i.mwdiv='W')"
		ElseIf FRectMWDiv<>"" Then
			addSql = addSql & " and i.mwdiv='"& FRectMWDiv & "'"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = i.itemid and M.mallid = '"&CMALLNAME&"' "
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = '"&CMALLNAME&"' "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_ltimallAddoption_regitem as J on J.midx = M.idx "
		Else
			sqlStr = sqlStr & "	JOIN db_etcmall.dbo.tbl_ltimallAddoption_regitem as J on J.midx = M.idx "
		End If
		sqlStr = sqlStr & "	LEFT Join db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'ltiMall' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " WHERE 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LTiMallGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  ''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'주문제작
			End If
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//제휴몰 판매만 허용
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//착불배송 상품 제거
		End If
		sqlStr = sqlStr & addSql
		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " & CStr(FPageSize*FCurrPage) & "  M.idx, isnull(M.itemnameChange, '') as itemnameChange, isnull(M.newitemname, '') as newitemname, i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.sellcash, i.buycash"
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.LTiMallRegdate, J.LTiMallLastUpdate, J.LTiMallGoodNo, J.LTiMallTmpGoodNo, J.LTiMallPrice, J.LTiMallSellYn, J.regUserid, IsNULL(J.LTiMallStatCd,-9) as LTiMallStatCd "
		sqlStr = sqlStr & "	, c.mapCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, i.itemdiv"
		sqlStr = sqlStr & "	, o.itemoption , o.optaddprice, o.optionname, o.optlimitno, o.optlimitsold, o.optsellyn "
		sqlStr = sqlStr & "	, M.optionname as regedOptionname, M.itemname as regedItemname "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = i.itemid and M.mallid = '"&CMALLNAME&"' "
		sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = '"&CMALLNAME&"' "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_etcmall.dbo.tbl_ltimallAddoption_regitem as J on J.midx = M.idx "
		Else
			sqlStr = sqlStr & "	JOIN db_etcmall.dbo.tbl_ltimallAddoption_regitem as J on J.midx = M.idx "
		End If
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'ltiMall' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & "	LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " where 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LTiMallGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  ''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'주문제작
			End If
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//제휴몰 판매만 허용
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//착불배송 상품 제거
    	End If
		sqlStr = sqlStr & addSql
		If (FRectOrdType = "LS") AND (FRectLotteNotReg = "F") Then
			sqlStr = sqlStr & " ORDER BY J.lastStatCheckDate, J.LtiMallLastupdate"
		ElseIf (FRectLotteNotReg = "F") Then
		    sqlStr = sqlStr & " ORDER BY J.LtiMallLastupdate "
		ElseIf (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.itemid DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC" '' m.regdate desc
	    End If
		rsget.pagesize = FPageSize
'rw sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CLotteiMallItem
					FItemList(i).Fidx				= rsget("idx")
					FItemList(i).FNewitemname		= rsget("newitemname")
					FItemList(i).FItemnameChange	= rsget("itemnameChange")
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).FLTiMallRegdate	= rsget("LTiMallRegdate")
					FItemList(i).FLTiMallLastUpdate	= rsget("LTiMallLastUpdate")
					FItemList(i).FLTiMallGoodNo		= rsget("LTiMallGoodNo")
					FItemList(i).FLTiMallTmpGoodNo	= rsget("LTiMallTmpGoodNo")
					FItemList(i).FLTiMallPrice		= rsget("LTiMallPrice")
					FItemList(i).FLTiMallSellYn		= rsget("LTiMallSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FLTiMallStatCd		= rsget("LTiMallStatCd")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FoptionCnt         = rsget("optionCnt")
	                FItemList(i).FrctSellCNT        = rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT		= rsget("accFailCNT")
	                FItemList(i).FlastErrStr		= rsget("lastErrStr")
	                FItemList(i).FinfoDiv           = rsget("infoDiv")
	                FItemList(i).Fitemdiv			= rsget("itemdiv")
					FItemList(i).FItemoption		= rsget("itemoption")
					FItemList(i).FOptaddprice		= rsget("optaddprice")
					FItemList(i).FOptionname		= rsget("optionname")
					FItemList(i).FOptlimitno		= rsget("optlimitno")
					FItemList(i).FOptlimitsold		= rsget("optlimitsold")
					FItemList(i).FOptsellyn			= rsget("optsellyn")
					FItemList(i).FRegedOptionname	= rsget("regedOptionname")
					FItemList(i).FRegedItemname		= rsget("regedItemname")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// 롯데iMall 상품 목록 // 수정시 조건이 달라야 함..
	Public Sub getLTiMallRegedItemList
		Dim sqlStr, addSql, i
		'브랜드검색
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'상품명 검색
		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End if

		'롯데아이몰 상품번호 검색
        If (FRectLtimallGoodNo <> "") then
            If Right(Trim(FRectLtimallGoodNo) ,1) = "," Then
            	FRectLtimallGoodNo = Replace(FRectLtimallGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallGoodNo in ('" & replace(Left(FRectLtimallGoodNo, Len(FRectLtimallGoodNo)-1),",","','") & "')"
            Else
				FRectLtimallGoodNo = Replace(FRectLtimallGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallGoodNo in ('" & replace(FRectLtimallGoodNo,",","','") & "')"
            End If
        End If

		'롯데아이몰 승인전 상품번호 검색
        If (FRectLtimallTmpGoodNo <> "") then
            If Right(Trim(FRectLtimallTmpGoodNo) ,1) = "," Then
            	FRectLtimallTmpGoodNo = Replace(FRectLtimallTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallTmpGoodNo in ('" & replace(Left(FRectLtimallTmpGoodNo, Len(FRectLtimallTmpGoodNo)-1),",","','") & "')"
            Else
				FRectLtimallTmpGoodNo = Replace(FRectLtimallTmpGoodNo,",,",",")
            	addSql = addSql & " and J.LtimallTmpGoodNo in ('" & replace(FRectLtimallTmpGoodNo,",","','") & "')"
            End If
        End If

		'카테고리 검색
		If FRectCDL <> "" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if
		If FRectCDM <> "" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if
		If FRectCDS <> "" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End If

		'등록여부 검색
		Select Case FRectExtNotReg
			Case "Q"	''등록실패
				addSql = addSql & " and J.LtimallStatCd = '-1'"
			Case "J"	'등록예정이상
				addSql = addSql & " and J.LtimallStatCd >= '0'"
			Case "W"	'등록예정
				addSql = addSql & " and J.LtimallStatCd = '0'"
		    Case "A"	'전송시도중오류
				addSql = addSql & " and J.LtimallStatCd = '1'"
			Case "C"	'반려
			    addSql = addSql & " and J.LtimallStatCd = '40'"
			Case "F"	'등록완료(임시)
			    addSql = addSql & " and J.LtimallStatCd = '20'"
			Case "D"	'등록완료(전시)
			    addSql = addSql & " and J.LtimallStatCd = '7'"
				addSql = addSql & " and J.LtimallGoodNo is Not Null"
			Case "R"	'수정요망		'스케줄링에서 사용
				addSql = addSql & " and (J.LtimallStatCd = '7')"
				addSql = addSql & " and J.LtimallLastUpdate < i.lastupdate"
				addSql = addSql & " and isnull(J.LtimallGoodNo, '') <> '' "
		End Select

		'미등록 라디오버튼 클릭 시
		Select Case FRectIsReged
			Case "N"	'등록예정이상
			    addSql = addSql & " and J.itemid is NULL  and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
		End Select

		'판매여부 검색
		Select Case FRectSellYn
			Case "Y"	addSql = addSql & " and i.sellYn='Y'"			'판매
			Case "N"	addSql = addSql & " and i.sellYn in ('S','N')"	'품절
		End Select

		'텐바이텐 한정여부 검색
		If FRectLimitYn <> "" Then
			addSql = addSql & " and i.limitYn = '" & FRectLimitYn & "'"
		End If

		'텐바이텐 세일여부 검색
		If FRectSailYn <> "" Then
			addSql = addSql & " and i.sailYn = '" & FRectSailYn & "'"
		End If

		'역마진 및 마진 CMAXMARGIN 이상 검색
		If (FRectonlyValidMargin <> "") Then
			IF (FRectonlyValidMargin = "Y") Then
				addSql = addSql & " and i.sellcash<>0"
				addSql = addSql & " and i.sellcash - i.buycash > 0 "
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&CMAXMARGIN & VbCrlf
			Else
				addSql = addSql & " and i.sellcash<>0"
				addSql = addSql & " and i.sellcash - i.buycash > 0 "
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<"&CMAXMARGIN & VbCrlf
			End If
		End If

		If (FRectStartMargin <> "") OR (FRectEndMargin <> "") Then
			If (FRectStartMargin <> "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and ("
				addSql = addSql & " 	convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
				addSql = addSql & " 	and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
				addSql = addSql & " ) "
			ElseIf (FRectStartMargin <> "") And (FRectEndMargin = "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)>="&FRectStartMargin & VbCrlf
			ElseIf (FRectStartMargin = "") And (FRectEndMargin <> "") Then
				addSql = addSql & " and convert(int, ((i.sellcash-i.buycash)/(CASE WHEN i.sellcash=0 THEN 1 ELSE i.sellcash END))*100)<="&FRectEndMargin & VbCrlf
			End If
		End If

		'주문제작 여부 검색
		If FRectisMadeHand <> "" Then
			If (FRectisMadeHand = "Y") Then
				addSql = addSql & " and i.itemdiv in ('06', '16')" & VbCrlf
			Else
				addSql = addSql & " and i.itemdiv not in ('06', '16')" & VbCrlf
			End If
		End if

		'옵션 여부 검색
		If FRectIsOption <> "" Then
			If FRectIsOption = "optAll" Then			'옵션전체
				addSql = addSql & " and i.optioncnt > 0"
			ElseIf FRectIsOption = "optaddpricey" Then	'추가금액Y
				addSql = addSql & " and i.optioncnt > 0"
 				addSql = addSql & " and J.optAddPrcCnt > 0"
			ElseIf FRectIsOption = "optaddpricen" Then	'추가금액N
				addSql = addSql & " and i.optioncnt > 0"
				addSql = addSql & " and isNULL(J.optAddPrcCnt,0)=0"
			ElseIf FRectIsOption = "optN" Then			'단품
				addSql = addSql & " and i.optioncnt = 0"
			ElseIf FRectIsOption = "optAddPrcRegType" Then
				addSql = addSql & " and J.optAddPrcCnt > 0"
				addSql = addSql & " and J.optAddPrcRegType = 0"
			End If
		End If

		'텐바이텐 품목정보 검색
		If (FRectInfoDiv <> "") then
			If (FRectInfoDiv = "YY") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')<>''"
			ElseIf (FRectInfoDiv = "NN") Then
				addSql = addSql & " and isNULL(ct.infodiv,'')=''"
			Else
				addSql = addSql & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End If
		End If

		'텐바이텐 등록제외 브랜드 제외 검색
		If (FRectNotinmakerid <> "") then
			If (FRectNotinmakerid = "Y") Then
				addSql = addSql & " and i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
			ElseIf (FRectNotinmakerid = "N") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
			End If
		End If

		'텐바이텐 등록제외 상품 제외 검색
		If (FRectNotinitemid <> "") then
			If (FRectNotinitemid = "Y") Then
				addSql = addSql & " and i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
			ElseIf (FRectNotinitemid = "N") Then
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
			End If
		End If

		'제휴몰 전송제외 상품 검색
		If (FRectExcTrans <> "") then
			If (FRectExcTrans = "Y") Then
				addSql = addSql & " and 'Y' = (CASE WHEN i.isusing='N' "
				addSql = addSql & " or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
				addSql = addSql & " or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
				addSql = addSql & " or i.isExtUsing='N' "
				addSql = addSql & " or uc.isExtUsing='N' "
				addSql = addSql & " or i.deliveryType = 7 "
				addSql = addSql & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				addSql = addSql & " or i.itemdiv = '21' "
				addSql = addSql & " or i.deliverfixday in ('C','X','G') "
				addSql = addSql & " or i.itemdiv >= 50 "
				addSql = addSql & " or i.itemdiv = '08' "
				addSql = addSql & " or i.itemdiv = '09' "
				addSql = addSql & " or i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or not (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " or not ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "F") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and i.itemdiv <> '21' "											'// 딜상품
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// 꽃배달, 화물배달, 해외직구
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.itemdiv <> '08' "											'// 티켓(강좌) 상품
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present상품
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and 'Y' = (CASE WHEN i.cate_large = '999' "
				addSql = addSql & " or i.cate_large='' "
				addSql = addSql & " or J.accFailCnt > 0 "
				addSql = addSql & " THEN 'Y' ELSE 'N' END) "
			ElseIf (FRectExcTrans = "N") Then
				addSql = addSql & " and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='lotteimall') "
				addSql = addSql & " and i.isusing='Y' "
				addSql = addSql & " and i.isExtUsing='Y' "											'// 외부몰허용상품
				addSql = addSql & " and uc.isExtUsing='Y' "
				addSql = addSql & " and i.deliveryType <> 7 "										'// 업체착불
				addSql = addSql & " and i.itemdiv <> '21' "											'// 딜상품
				addSql = addSql & " and i.deliverfixday not in ('C','X','G') "						'// 꽃배달, 화물배달, 해외직구
				addSql = addSql & " and not ((i.deliveryType = 9) and (i.sellcash < 10000)) "		'// 판매가(할인가) 1만원 미만
				addSql = addSql & " and i.itemdiv <> '08' "											'// 티켓(강좌) 상품
				addSql = addSql & " and i.itemdiv <> '09' "											'// Present상품
				addSql = addSql & " and i.cate_large <> '999' "										'// 카테고리 미지정
				addSql = addSql & " and i.cate_large <> '' "										'// 카테고리 미지정
				addSql = addSql & " and i.itemdiv < 50 "
				addSql = addSql & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.optsellyn='Y' and o.itemid=i.itemid and (o.optlimityn <> 'Y' or (o.optlimitno-o.optlimitsold)>5)) "
				addSql = addSql & " ) "
				addSql = addSql & " and isNULL(ct.infodiv,'') not in ('','18','20','22') "
				addSql = addSql & " and not ('1'+ct.infoDiv in ('107', '108', '109', '110', '111', '112', '113', '114', '123') and not exists (select top 1 tr.itemid from db_item.dbo.tbl_safetycert_tenReg tr where tr.itemid = i.itemid and isnull(TR.certNum, '') <> '')) "
				addSql = addSql & " and not (i.optioncnt > 0 and exists (select top 1 r.itemid from [db_item].[dbo].tbl_OutMall_regedoption R where R.mallid = 'lotteimall' and R.itemid = i.itemid and R.itemoption = '0000')) "
				addSql = addSql & " and ( "
				addSql = addSql & " 	i.optioncnt = 0 "
				addSql = addSql & " 	or "
				addSql = addSql & " 	not exists(SELECT top 1 o.itemid FROM [db_item].[dbo].tbl_item_option o WHERE o.isUsing='Y' and o.itemid=i.itemid and o.optAddPrice > 0) "
				addSql = addSql & " ) "
			End If
		End If

        '특가 상품 여부
        If (FRectIsSpecialPrice <> "") then
            If (FRectIsSpecialPrice = "Y") Then
				addSql = addSql & " and (GETDATE() > mi.startDate and GETDATE() <= mi.endDate) "
            End If
        End If

		'옵션추가금액New
		If (FRectPriceOption <> "") then
			If (FRectPriceOption = "Y") Then
				addSql = addSql & " and i.itemid in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			ElseIf (FRectPriceOption = "N") Then
				addSql = addSql & " and i.itemid not in (SELECT itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists]) "
			End If
		End If

		'롯데아이몰 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and J.LtimallSellYn <> 'X'"
			Else
				addSql = addSql & " and J.LtimallSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		'등록수정오류상품
		Select Case FRectFailCntExists
			Case "Y"	'오류1회이상
				addSql = addSql & " and J.accFailCNT>0"
			Case "N"	'오류0회
				addSql = addSql & " and J.accFailCNT=0"
		End Select

		'롯데아이몰 카테고리 매칭 여부
		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				addSql = addSql & " and isnull(c.mapCnt, '') <> ''"
			Case "N"	'미매칭
				addSql = addSql & " and isnull(c.mapCnt, '') = ''"
		End Select

        '롯데아이몰 < 10x10 가격
		If (FRectexpensive10x10 <> "") Then
			addSql = addSql & " and J.LtimallPrice is Not Null and J.LtimallPrice < i.sellcash"
		End If

		'가격상이전체보기
		If FRectdiffPrc <> "" Then
			addSql = addSql & " and J.LtimallPrice is Not Null and i.sellcash <> J.LtimallPrice "
		End If

		'롯데아이몰판매,  10x10 품절
		If (FRectLtimallYes10x10No <> "") Then
			addSql = addSql & " and i.sellyn<>'Y'"
			addSql = addSql & " and J.LtimallSellYn='Y'"
		End If

		'롯데아이몰품절&텐바이텐판매가능(판매중,한정>=10) 상품보기
		If FRectLtimallNo10x10Yes <> "" Then
			addSql = addSql & " and (J.LtimallSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>"&CMAXLIMITSELL&")))"
		End If

		'수정요망상품보기(최종업데이트일 기준)
		If FRectReqEdit <> "" Then
			addSql = addSql & " and J.LtimallLastUpdate < i.lastupdate "
		End If

		'스케줄링에서 사용 실패횟수 제한
		If (FRectFailCntOverExcept <> "") Then
			addSql = addSql & " and J.accFailCNT < "&FRectFailCntOverExcept
		End If

		'스케줄링에서 사용 라스트업데이트 기준 수정
		If (FRectOrdType = "LU") Then
		    addSql = addSql & " and isnull(J.lastStatCheckDate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(J.LtimallLastUpdate, 10) "
		End If

		'배송구분
		If (FRectDeliverytype <> "") Then
			addSql = addSql & " and i.deliverytype='" & FRectDeliverytype & "'"
		End If

		'거래구분
		If FRectMWDiv = "MW" Then
			addSql = addSql & " and (i.mwdiv='M' or i.mwdiv='W')"
		ElseIf FRectMWDiv<>"" Then
			addSql = addSql & " and i.mwdiv='"& FRectMWDiv & "'"
		End If

		'제휴 사용 여부(상품)
		If (FRectIsextusing <> "") Then
			addSql = addSql & " and i.isextusing='" & FRectIsextusing & "'"
		End If

		'제휴 사용 여부(브랜드)
		If (FRectCisextusing <> "") Then
			addSql = addSql & " and uc.isextusing='" & FRectCisextusing & "'"
		End If

		'3개월 판매량
		Select Case FRectRctsellcnt
			Case "0"	'0
				addSql = addSql & " and isnull(J.rctSellCnt, 0) = 0 "
			Case "1"	'1개이상
				addSql = addSql & " and isnull(J.rctSellCnt, 0) >= 1 "
		End Select

		'구매유형
		If (FRectPurchasetype <> "") Then
			Select Case FRectPurchasetype
				Case "101"
                    addSql = addSql & " and p.purchasetype in (4, 5, 6, 7, 8) "
				Case "356"	'0
					addSql = addSql & " and p.purchasetype in (3, 5, 6) "
				Case Else
					addSql = addSql & " and p.purchasetype='" & FRectPurchasetype & "'"
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_LTiMall_regItem as J on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_LTiMall_regItem as J on J.itemid = i.itemid "
		End If
		sqlStr = sqlStr & "	LEFT Join db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'ltiMall' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & " LEFT join db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LTiMallGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  ''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'주문제작
			End If
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//제휴몰 판매만 허용
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//착불배송 상품 제거
		End If
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
''rw sqlStr
		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.itemid, i.itemname, i.smallImage "
		sqlStr = sqlStr & "	, i.makerid, i.regdate, i.lastUpdate, i.orgPrice, i.orgSuplyCash, i.sellcash, i.buycash"
		sqlStr = sqlStr & "	, i.sellYn, i.sailyn, i.LimitYn, i.LimitNo, i.LimitSold, i.deliverytype, i.optionCnt"
		sqlStr = sqlStr & "	, J.LTiMallRegdate, J.LTiMallLastUpdate, J.LTiMallGoodNo, J.LTiMallTmpGoodNo, J.LTiMallPrice, J.LTiMallSellYn, J.regUserid, IsNULL(J.LTiMallStatCd,-9) as LTiMallStatCd "
		sqlStr = sqlStr & "	, c.mapCnt, J.regedOptCnt, J.rctSellCNT, J.accFailCNT, J.lastErrStr "
		sqlStr = sqlStr & " ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr & "	, Ct.infoDiv, J.optAddPrcCnt, J.optAddPrcRegType, i.itemdiv, mi.mustPrice as specialPrice, mi.startDate, mi.endDate, p.purchasetype "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		If (FRectIsReged = "N") OR (FRectIsReged = "A") Then		'//미등록이 아니면 JOIN
			sqlStr = sqlStr & "	LEFT JOIN db_item.dbo.tbl_LTiMall_regItem as J on J.itemid = i.itemid "
		Else
			sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_LTiMall_regItem as J on J.itemid = i.itemid "
		End If
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_OutMall_CateMap_Summary as c on c.mallid = 'ltiMall' and c.tenCateLarge = i.cate_large and c.tenCateMid = i.cate_mid and c.tenCateSmall = i.cate_small "
		sqlStr = sqlStr & "	LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents as ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_outmall_mustPriceItem as mi with (nolock) on mi.itemid = i.itemid and mi.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " where 1 = 1"
		If (FRectIsReged <> "N" and FRectExtNotReg <> "Q")  Then		'// 미등록도 아니고 등록실패도 아니면 조건 없음
			If FRectIsReged = "Q" Then
				sqlStr = sqlStr & " and J.LTiMallGoodNo is Not Null "
				sqlStr = sqlStr & " and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>5)) "
				sqlStr = sqlStr & " and 'N' = (CASE WHEN i.isusing='N'  "
				sqlStr = sqlStr & " or i.isExtUsing='N' "
				sqlStr = sqlStr & " or uc.isExtUsing='N' "
				sqlStr = sqlStr & " or ((i.deliveryType = 9) and (i.sellcash < 10000)) "
				sqlStr = sqlStr & " or i.sellyn<>'Y' "
				sqlStr = sqlStr & " or i.deliverfixday in ('C','X','G') "
				sqlStr = sqlStr & " or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large='' "
				sqlStr = sqlStr & " or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "
				sqlStr = sqlStr & " THEN 'Y' ELSE 'N' END) "
			End If
		Else
    		sqlStr = sqlStr & " and i.isusing='Y' "
    		sqlStr = sqlStr & " and i.deliverfixday not in ('C','X','G') "
    		sqlStr = sqlStr & " and i.basicimage is not null "
    		sqlStr = sqlStr & " and i.itemdiv<50 "  ''and i.itemdiv<>'08'
    		sqlStr = sqlStr & " and i.itemdiv not in ('08','09')"
    		sqlStr = sqlStr & " and i.cate_large<>'' "
		    sqlStr = sqlStr & " and ((i.cate_large <> '999') or ((i.cate_large='999')and(i.makerid='ftroupe'))) " & VBCRLF
    		sqlStr = sqlStr & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
			If FRectExtNotReg <> "" Then
				sqlStr = sqlStr & " and i.sellcash>=1000 "  & VBCRLF
				'sqlStr = sqlStr & " and i.itemdiv<>'06'" & VBCRLF				'주문제작
			End If
    		sqlStr = sqlStr & "	and uc.isExtUsing='Y'"  ''20130304 브랜드 제휴사용여부 Y만.
			sqlStr = sqlStr & " and i.isExtUsing='Y'"														'//제휴몰 판매만 허용
			sqlStr = sqlStr & " and i.deliverytype not in ('7')"											'//착불배송 상품 제거
    	End If
		sqlStr = sqlStr & addSql
		If (FRectOrdType = "LS") AND (FRectLotteNotReg = "F") Then
			sqlStr = sqlStr & " ORDER BY J.lastStatCheckDate, J.LtiMallLastupdate"
		ElseIf (FRectLotteNotReg = "F") Then
		    sqlStr = sqlStr & " ORDER BY J.LtiMallLastupdate "
		ElseIf (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC "
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY J.rctSellCNT DESC, i.itemscore DESC, J.itemid DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC" '' m.regdate desc
	    End If
		rsget.pagesize = FPageSize
'rw sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CLotteiMallItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).FOrgSuplycash		= rsget("OrgSuplycash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).FLTiMallRegdate	= rsget("LTiMallRegdate")
					FItemList(i).FLTiMallLastUpdate	= rsget("LTiMallLastUpdate")
					FItemList(i).FLTiMallGoodNo		= rsget("LTiMallGoodNo")
					FItemList(i).FLTiMallTmpGoodNo	= rsget("LTiMallTmpGoodNo")
					FItemList(i).FLTiMallPrice		= rsget("LTiMallPrice")
					FItemList(i).FLTiMallSellYn		= rsget("LTiMallSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FLTiMallStatCd		= rsget("LTiMallStatCd")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
					If Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If
	                FItemList(i).FoptionCnt         = rsget("optionCnt")
	                FItemList(i).FregedOptCnt       = rsget("regedOptCnt")
	                FItemList(i).FrctSellCNT        = rsget("rctSellCNT")
	                FItemList(i).FaccFailCNT		= rsget("accFailCNT")
	                FItemList(i).FlastErrStr		= rsget("lastErrStr")
	                FItemList(i).FinfoDiv           = rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt      = rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType  = rsget("optAddPrcRegType")
	                FItemList(i).Fitemdiv		  	= rsget("itemdiv")
					FItemList(i).FSpecialPrice		= rsget("specialPrice")
					FItemList(i).FStartDate	      	= rsget("startDate")
					FItemList(i).FEndDate			= rsget("endDate")
					FItemList(i).FPurchasetype		= rsget("purchasetype")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    ''' 등록되지 말아야 될 상품..
    public Sub getLtiMallreqExpireItemList
		dim sqlStr, addSql, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_LtiMall_regItem as m on i.itemid=m.itemid and m.LTiMallGoodNo is Not Null and m.LTiMallSellYn = 'Y' "                ''' 롯데 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid = c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid = ct.itemid"
		sqlStr = sqlStr & " WHERE (i.isusing <> 'Y' or i.isExtUsing <> 'Y' or i.deliverytype in ('7') "
		'//조건배송 10000원 이상
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        END IF
		sqlStr = sqlStr & " 	or i.deliverfixday in ('C','X','G') "
		sqlStr = sqlStr & " 	or i.itemdiv='08'"
		sqlStr = sqlStr & " 	or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','22')"  ''화장품, 식품류 제외
        sqlStr = sqlStr & " )"

        ''//연동 제외상품
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt<getdate()"
        sqlStr = sqlStr & "     and edDt>getdate()"
        sqlStr = sqlStr & "     and mallgubun='"&CMALLNAME&"'"
        sqlStr = sqlStr & " )"

        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'롯데아이몰 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and m.LtimallSellYn <> 'X'"
			Else
				addSql = addSql & " and m.LtimallSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		''2013/05/29 추가
		If (FRectInfoDiv <> "") Then
			If (FRectInfoDiv = "YY") then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'')<>''"
			Elseif (FRectInfoDiv = "NN") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'')=''"
			Else
				sqlStr = sqlStr & " and ct.infodiv='"&FRectInfoDiv&"'"
			End if
		End If
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " i.* "
		sqlStr = sqlStr & "	, m.LTiMallRegdate, m.LTiMallLastUpdate, m.LTiMallGoodNo, m.LTiMallTmpGoodNo, m.LTiMallPrice, m.LTiMallSellYn, m.regUserid, m.LTiMallStatCd "
		sqlStr = sqlStr & "	, 1 as mapCnt "
		sqlStr = sqlStr & " ,c.defaultdeliverytype, c.defaultfreeBeasongLimit"
		sqlStr = sqlStr & " ,ct.infoDiv, m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_LtiMall_regItem as m on i.itemid=m.itemid and m.LTiMallGoodNo is Not Null and m.LTiMallSellYn= 'Y' "                ''' 롯데 판매중인거만.
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item_contents ct on i.itemid=ct.itemid"
		sqlStr = sqlStr & " WHERE (i.isusing<>'Y' or i.isExtUsing<>'Y' "
		sqlStr = sqlStr & " 	or i.deliverytype in ('7') "
		'//조건배송 10000원 이상
		IF (CUPJODLVVALID) then
		    sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<10000) )" ''
		ELSE
            sqlStr = sqlStr & " or ((i.deliveryType=9) and (i.sellcash<isNULL(c.defaultFreebeasongLimit,0)) )" ''
        ENd IF
		sqlStr = sqlStr & "		or i.deliverfixday in ('C','X','G') "
		sqlStr = sqlStr & "     or i.itemdiv='08'"
		sqlStr = sqlStr & "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		sqlStr = sqlStr & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		sqlStr = sqlStr & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		sqlStr = sqlStr & "		or c.isExtUsing='N'"
		sqlStr = sqlStr & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		sqlStr = sqlStr & "		or isNULL(ct.infodiv,'') in ('','18','20','22')"
        sqlStr = sqlStr & " )"

        ''//연동 제외상품 //디비로 만들어야 할듯.
        sqlStr = sqlStr & " and i.itemid not in ("
        sqlStr = sqlStr & "     select itemid from db_temp.dbo.tbl_jaehyumall_not_edit_itemid"
        sqlStr = sqlStr & "     where stDt < getdate()"
        sqlStr = sqlStr & "     and edDt > getdate()"
        sqlStr = sqlStr & "     and mallgubun = '"&CMALLNAME&"'"
        sqlStr = sqlStr & " )"

        sqlStr = sqlStr & " and i.makerid<>'ftroupe'"  ''2013/07/19 ftroupe 예외처리

        If FRectMakerid <> "" Then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		End if

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		'롯데아이몰 판매여부
		If (FRectExtSellYn<>"") then
			If (FRectExtSellYn = "YN") Then
				addSql = addSql & " and m.LtimallSellYn <> 'X'"
			Else
				addSql = addSql & " and m.LtimallSellYn='" & FRectExtSellYn & "'"
			End if
		End If

		''2013/05/29 추가
		If (FRectInfoDiv <> "") Then
			If (FRectInfoDiv = "YY") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'') <> ''"
			Elseif (FRectInfoDiv = "NN") Then
				sqlStr = sqlStr & " and isNULL(ct.infodiv,'') = ''"
			Else
				sqlStr = sqlStr & " and ct.infodiv = '"&FRectInfoDiv&"'"
			End if
		End If
		sqlStr = sqlStr & " ORDER BY m.regdate DESC, i.itemid DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				set FItemList(i) = new CLotteiMallItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).FOrgSuplycash		= rsget("OrgSuplycash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")

					FItemList(i).FLTiMallRegdate	= rsget("LTiMallRegdate")
					FItemList(i).FLTiMallLastUpdate	= rsget("LTiMallLastUpdate")
					FItemList(i).FLTiMallGoodNo		= rsget("LTiMallGoodNo")
					FItemList(i).FLTiMallTmpGoodNo	= rsget("LTiMallTmpGoodNo")
					FItemList(i).FLTiMallPrice		= rsget("LTiMallPrice")
					FItemList(i).FLTiMallSellYn		= rsget("LTiMallSellYn")
					FItemList(i).FregUserid			= rsget("regUserid")
					FItemList(i).FLTiMallStatCd		= rsget("LTiMallStatCd")
					FItemList(i).FCateMapCnt		= rsget("mapCnt")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
	                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
	                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")

					If Not(FItemList(i).FsmallImage = "" or isNull(FItemList(i).FsmallImage)) Then
						FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
					Else
						FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
					End If

	                FItemList(i).FinfoDiv 			= rsget("infoDiv")
	                FItemList(i).FoptAddPrcCnt      = rsget("optAddPrcCnt")
	                FItemList(i).FoptAddPrcRegType  = rsget("optAddPrcRegType")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'--------------------------------------------------------------------------------
	'// 미등록 상품 목록(등록용)
	Public Sub getLTiMallNotRegItemList
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			'''2013-07-25 김진영 옵션 추가금액 있는경우, 옵션금액 팝업에서 설정한 것만
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & " select itemid from ("
            addSql = addSql & "     select o.itemid"
            addSql = addSql & " 	,count(*) as optCNT"
            addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	from db_item.dbo.tbl_item_option as o "
            addSql = addSql & " 	left join db_item.dbo.tbl_LTiMall_regItem as RR on o.itemid = RR.itemid and RR.itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	where o.itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and o.isusing='Y'"
            addSql = addSql & " 	and isnull(RR.optAddPrcRegType,'') = '0'"
            addSql = addSql & " 	group by o.itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " where optAddCNT>0"
            addSql = addSql & " or (optCnt-optNotSellCnt<1)"
            addSql = addSql & " )"

            ''' 2013/05/29 특정품목 등록 불가 (화장품, 식품류)
            addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18','20','22')"
		End If
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, '"&CitemGbnKey&"' as itemGbnKey"
		strSql = strSql & "	, isNULL(R.LtiMallStatCD,-9) as LtiMallStatCD"

		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotteimall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " JOIN db_user.dbo.tbl_user_c UC on i.makerid = UC.userid"
		''strSql = strSql & " 	Join db_item.dbo.tbl_LTiMall_cateGbn_mapping G"
		''strSql = strSql & " 		on G.tenCateLarge=i.cate_large and G.tenCateMid=i.cate_mid and G.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_LtiMall_regItem R on i.itemid=R.itemid"
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = 'lotteimall' "
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
		ELSE
		    strSql = strSql & "	and (i.deliveryType <> 9)"
	    END IF
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "			'플라워/화물배송/해외직구 상품 제외
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & "	and UC.isExtUsing <> 'N'"
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		''strSql = strSql & "     and i.sellcash=i.orgprice"              '''당분간 할인 안하는것만.. // 가격수정 모듈 없음..?
		''strSql = strSql & " 	and (i.orgprice<>0 and ((i.orgprice-i.orgSuplyCash)/i.orgprice)*100>=" & CMAXMARGIN & ")"							'역마진 상품 제외
		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and i.itemid not in (Select itemid From db_item.dbo.tbl_LtiMall_regItem where LtiMallStatCD>3) "	''LtiMallStatCD>=3 등록완료이상은 등록안됨.										'롯데등록상품 제외
		''strSql = strSql & "		and cm.mapCnt is Not Null "	& addSql
		strSql = strSql & addSql																				'카테고리 매칭 상품만
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  Then
			Do until rsget.EOF
				Set FItemList(i) = new CLotteiMallItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).FtenCateLarge		= rsget("cate_large")
					FItemList(i).FtenCateMid		= rsget("cate_mid")
					FItemList(i).FtenCateSmall		= rsget("cate_small")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FitemDiv			= rsget("itemdiv")
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).ForgSuplyCash		= rsget("orgSuplyCash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FisUsing			= rsget("isusing")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).Fkeywords			= rsget("keywords")
					FItemList(i).Fvatinclude        = rsget("vatinclude")
					FItemList(i).ForderComment		= db2html(rsget("ordercomment"))
					FItemList(i).FoptionCnt			= rsget("optionCnt")
					FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
					FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
					FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
					FItemList(i).Fsourcearea		= db2html(rsget("sourcearea"))
					FItemList(i).Fmakername			= db2html(rsget("makername"))
					FItemList(i).FUsingHTML			= rsget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsget("itemcontent"))
	                FItemList(i).FitemGbnKey        = rsget("itemGbnKey")
	                FItemList(i).FLtiMallStatCD     = rsget("LtiMallStatCD")
	                FItemList(i).FRectMode			= FRectMode

	                FItemList(i).FinfoDiv			= rsget("infoDiv")
	                FItemList(i).Fsafetyyn			= rsget("safetyyn")
	                FItemList(i).FsafetyDiv			= rsget("safetyDiv")
	                FItemList(i).FsafetyNum			= rsget("safetyNum")
					FItemList(i).FOutmallstandardMargin	= rsget("outmallstandardMargin")
					i = i + 1
					rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'--------------------------------------------------------------------------------
	'// 롯데iMall 상품 목록(수정용)
	public Sub getLTiMallEditedItemList
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		ElseIf FRectNotJehyu = "Y" Then
			'제휴몰 상품이 아닌것
			addSql = " and i.isExtUsing='N' "
		Else
			'수정된 상품만
			addSql = " and m.LtiMallLastUpdate < i.lastupdate"
		End If

        ''//연동 제외상품
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt < getdate()"
        addSql = addSql & "     and edDt > getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, m.LtiMallGoodNo, m.LtiMallTmpGoodNo, m.LtiMallSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut"
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_LtiMall_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotteimall_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		''If (FRectMatchCateNotCheck <> "on") Then
		if (FRectMatchCate="Y") THEN  '' eastone 수정 2013/09/01
		    strSql = strSql & " and cm.mapCnt is Not Null "
	    End If
		strSql = strSql & addSql
		strSql = strSql & " and isNULL(m.LtiMallTmpGoodNo, m.LtiMallGoodNo) is Not Null "									'#등록 상품만
''rw strSql
		rsget.Open strSql,dbget,1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		if not rsget.EOF Then
			Do until rsget.EOF
				Set FItemList(i) = new CLotteiMallItem
					FItemList(i).Fitemid			= rsget("itemid")
					FItemList(i).FtenCateLarge		= rsget("cate_large")
					FItemList(i).FtenCateMid		= rsget("cate_mid")
					FItemList(i).FtenCateSmall		= rsget("cate_small")
					FItemList(i).Fitemname			= db2html(rsget("itemname"))
					FItemList(i).FitemDiv			= rsget("itemdiv")
					FItemList(i).FsmallImage		= rsget("smallImage")
					FItemList(i).Fmakerid			= rsget("makerid")
					FItemList(i).Fregdate			= rsget("regdate")
					FItemList(i).FlastUpdate		= rsget("lastUpdate")
					FItemList(i).ForgPrice			= rsget("orgPrice")
					FItemList(i).ForgSuplyCash		= rsget("orgSuplyCash")
					FItemList(i).FSellCash			= rsget("sellcash")
					FItemList(i).FBuyCash			= rsget("buycash")
					FItemList(i).FsellYn			= rsget("sellYn")
					FItemList(i).FsaleYn			= rsget("sailyn")
					FItemList(i).FisUsing			= rsget("isusing")
					FItemList(i).FLimitYn			= rsget("LimitYn")
					FItemList(i).FLimitNo			= rsget("LimitNo")
					FItemList(i).FLimitSold			= rsget("LimitSold")
					FItemList(i).Fkeywords			= rsget("keywords")
					FItemList(i).ForderComment		= db2html(rsget("ordercomment"))
					FItemList(i).FoptionCnt			= rsget("optionCnt")
					FItemList(i).FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
					FItemList(i).FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
					FItemList(i).FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
					FItemList(i).Fsourcearea		= db2html(rsget("sourcearea"))
					FItemList(i).Fmakername			= db2html(rsget("makername"))
					FItemList(i).FUsingHTML			= rsget("usingHTML")
					FItemList(i).Fitemcontent		= db2html(rsget("itemcontent"))
					FItemList(i).FLTiMallGoodNo		= rsget("LtiMallGoodNo")
					FItemList(i).FLtiMallTmpGoodNo	= rsget("LtiMallTmpGoodNo")
					FItemList(i).FLtiMallSellYn		= rsget("LtiMallSellYn")
					FItemList(i).Fvatinclude        = rsget("vatinclude")
	                FItemList(i).FoptionCnt         = rsget("optionCnt")
	                FItemList(i).FregedOptCnt       = rsget("regedOptCnt")
	                FItemList(i).FaccFailCNT        = rsget("accFailCNT")
	                FItemList(i).FlastErrStr        = rsget("lastErrStr")
	                ''FItemList(i).Fcorp_dlvp_sn      = rsget("returnCode")
	                FItemList(i).Fdeliverytype      = rsget("deliverytype")
	                FItemList(i).FrequireMakeDay    = rsget("requireMakeDay")

	                FItemList(i).FinfoDiv       = rsget("infoDiv")
	                FItemList(i).Fsafetyyn      = rsget("safetyyn")
	                FItemList(i).FsafetyDiv     = rsget("safetyDiv")
	                FItemList(i).FsafetyNum     = rsget("safetyNum")
	                FItemList(i).FmaySoldOut    = rsget("maySoldOut")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'--------------------------------------------------------------------------------
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end class

'// MD상품군 선택상자 출력
Function printLotteCateGrpSelectBox(fnm,selcd)
	Dim strSql, rstStr
	rstStr = "<Select name='" & fnm & "' class='select'>"
	rstStr = rstStr & "<option value=''>전체</option>"
	strSql = "Select * From db_temp.dbo.tbl_lotteiMall_MDCateGrp Where isUsing='Y'"
	rsget.Open strSql,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		Do Until rsget.EOF
			If cStr(rsget("groupCode")) = cStr(selcd) Then
				rstStr = rstStr & "<option value='" & rsget("groupCode") & "' selected>" & rsget("groupName")& "</option>"
			Else
				rstStr = rstStr & "<option value='" & rsget("groupCode") & "'>" & rsget("groupName")& "</option>"
			End If
			rsget.MoveNext
		Loop
	End If
	rsget.Close
	rstStr = rstStr & "</select>"
	printLotteCateGrpSelectBox = rstStr
End Function

'// 상품이미지 존재여부 검사
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

Function getLtiMallItemIdByTenItemID(iitemid)
	Dim sqlStr, retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT isNULL(ltiMallGoodNo, ltiMallTmpGoodNo) as ltiMallGoodNo " & VBCRLF
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_LTiMall_regItem" & VBCRLF
	sqlStr = sqlStr & " WHERE itemid = "&iitemid & VBCRLF

	rsget.Open sqlStr,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		retVal = rsget("ltiMallGoodNo")
	End If
	rsget.Close

	If IsNULL(retVal) Then retVal = ""
	getLtiMallItemIdByTenItemID = retVal
End Function

Function getLtiMallTmpItemIdByTenItemID(iitemid)
	Dim sqlStr, retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT ltiMallTmpGoodNo, isnull(ltiMallGoodNo,'') as ltiMallGoodNo " & VBCRLF
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_LTiMall_regItem" & VBCRLF
	sqlStr = sqlStr & " WHERE itemid = "&iitemid & VBCRLF
	rsget.Open sqlStr,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		If rsget("ltiMallGoodNo") <> "" Then
			retVal = "전시상품"
		Else
			retVal = rsget("ltiMallTmpGoodNo")
		End If
	End If
	rsget.Close

	If IsNULL(retVal) Then retVal = ""
	getLtiMallTmpItemIdByTenItemID = retVal
End Function

''//상품명 변경 파라메터 생성(롯데닷컴과 파라매타명이 다름)
Function fnGetLtiMallItemNameEditParameter(iLotteGoodNo, iItemName)
	Dim strRst
	strRst = "subscriptionId=" & ltiMallAuthNo
	strRst = strRst & "&goods_no=" & iLotteGoodNo
	strRst = strRst & "&goods_nm=" & Trim(iItemName)
	strRst = strRst & "&chg_caus_cont=api 상품명 변경"
	fnGetLtiMallItemNameEditParameter = strRst
End Function

Function getOutmallstandardMargin
	Dim sqlStr
	sqlStr = ""
	sqlStr = sqlStr & " SELECT TOP 1 isNull(outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin " & VBCRLF
	sqlStr = sqlStr & " FROM db_partner.dbo.tbl_partner_addInfo " & VBCRLF
	sqlStr = sqlStr & " WHERE partnerid = '"& CMALLNAME &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getOutmallstandardMargin = rsget("outmallstandardMargin")
	Else
		getOutmallstandardMargin = CMAXMARGIN
	End If
	rsget.Close
End Function
%>
