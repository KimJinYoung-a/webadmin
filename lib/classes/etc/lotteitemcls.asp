<%
CONST CMAXMARGIN = 15			'' MaxMagin임.. '(롯데닷컴 11%)
CONST CMAXLIMITSELL = 5         '' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CMALLNAME = "lotteCom"
CONST CDEFALUT_STOCK = 99       '' 재고관리 수량 기본 99 (한정 아닌경우)

class CLotteItem
	Public FRegId
	Public FLastupdateId

	public FLastUpdate
	public FisUsing

	'담당MD
	public FMDCode
	public FMDName
	public FSellFeeType
	public FNormalSellFee
	public FEventSellFee

	'MD상품군
	public FgroupCode
	public FSuperGroupName
	public FGroupName

	'롯데닷컴 카테고리
	public FDispNo
	public FDispNm
	public FDispLrgNm
	public FDispMidNm
	public FDispSmlNm
	public FDispThnNm
	public FtenCateLarge
	public FtenCateMid
	public FtenCateSmall
	public FtenCDLName
	public FtenCDMName
	public FtenCDSName
	public FtenCateName
    public Fdisptpcd
    public FCateisusing

	'롯데닷컴 브랜드
	public FlotteBrandCd
	public FlotteBrandName
	public FTenMakerid
	public FTenBrandName

	'롯데닷컴 상품목록
	public FLotteRegdate
	public FLotteLastUpdate
	public FLotteGoodNo				'실상품번호
	public FLotteTmpGoodNo			'임시상품번호
	public FLottePrice
	public FLotteSellYn
	public FregUserid
	public FLotteDispCnt
	public FCateMapCnt
	public FLotteStatCd				'상품등록상태
    public FrctSellCNT              '6개월판매량
    public FregedOptCnt             '롯데등록옵션
    public FaccFailCNT              '등록수정 오류 횟수
    public FlastErrStr              '최종오류

    '''public Fcorp_dlvp_sn             '반품주소지코드

	'텐바이텐 상품목록
	public Fitemid
	public Fitemname
	public FitemDiv
	public FsmallImage
	public FbasicImage
	public FmainImage
	public FmainImage2
	public Fmakerid
	public Fregdate
	public ForgPrice
	public ForgSuplyCash
	public FSellCash
	public FBuyCash
	public FsellYn
	public FsaleYn
	public FLimitYn
	public FLimitNo
	public FLimitSold
	public Fkeywords
	public ForderComment
	public FoptionCnt
	public Fsourcearea
	public Fmakername
	public Fitemcontent
	public FUsingHTML
    public Fdeliverytype
    public FrequireMakeDay
    public Fdefaultdeliverytype
    public FdefaultfreeBeasongLimit
	Public Fsocname_kor
    ''품목정보 및 안전인증정보.
    public FinfoDiv
    public Fsafetyyn
    public FsafetyDiv
    public FsafetyNum

    public FoptAddPrcCnt
    public FoptAddPrcRegType

    public FmaySoldOut ''제휴몰 사용안함등

    public function getLotteReturnDLVCode
        dim sqlStr
        dim retReturnCode

        getLotteReturnDLVCode = "72125"    ''기본 반품지 코드(물류)
    Exit function '' 물류로 재통합 2013/02/25
        if Not (Fdeliverytype="2" or Fdeliverytype="7" or Fdeliverytype="9") then Exit function  ''자체배송이면 물류


'        ''임시 막음
'        if (Fcorp_dlvp_sn<>"72125") and (Fcorp_dlvp_sn<>"113045") and (Fcorp_dlvp_sn<>"113044") and (Fcorp_dlvp_sn<>"114747") then
'            Fcorp_dlvp_sn = "72125"
'        end if

        '' 반품주소지가 같은지 검사 방식으로 변경
        sqlStr = " select R.returnCode from db_item.dbo.tbl_OutMall_BrandReturnCode R"
        sqlStr = sqlStr & "  	Join db_temp.dbo.tbl_jaehyumall_returnInfo T"
        sqlStr = sqlStr & "	on R.makerid='"&Fmakerid&"'"
        sqlStr = sqlStr & "	and R.returnCode=T.returnCode"
        sqlStr = sqlStr & "	Join db_partner.dbo.tbl_partner p"
        sqlStr = sqlStr & "	on p.id=R.makerid"
        sqlStr = sqlStr & " where replace(T.returnAddress,' ','')=replace(replace(p.return_zipCode,'-','') +  p.return_address + p.return_address2,' ','')"

        rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
		    retReturnCode = rsget("returnCode")
        end if
        rsget.close

        if isNULL(retReturnCode) then Exit function
        if (retReturnCode="") then Exit function

        getLotteReturnDLVCode = CStr(retReturnCode)
    end function

    public function getDisptpcdName
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
    end function

    public function getDeliverytypeName
        if (Fdeliverytype="9") then
            getDeliverytypeName = "<font color='blue'>[조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
        elseif (Fdeliverytype="7") then
            getDeliverytypeName = "<font color='red'>[업체착불]</font>"
        elseif (Fdeliverytype="2") then
            getDeliverytypeName = "<font color='blue'>[업체]</font>"
        else
            getDeliverytypeName = ""
        end if
    end function

	'// 품절여부
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	end function

    public function getItemNameFormat()
        dim buf
        buf = replace(FItemName,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","프로")
        buf = replace(buf,"[무료배송]","")
        buf = replace(buf,"[무료 배송]","")
        getItemNameFormat = buf
    end function

    public function getLimitEa()
        dim ret
        ret = FLimitNo-FLimitSold

        if (ret<1) then ret=0
        getLimitEa = ret
    end function

    public function getLimitLotteEa()
        dim ret
        ret = FLimitNo-FLimitSold-5
		'2013-10-10 진영 추가..한정 중 1000개가 넘으면 최대 999로 넣겠다고 롯데에서 요청메일 옴
		If ret > 1000 Then
			 ret = 999
		End If

        if (ret<1) then ret=0
        getLimitLotteEa = ret
    end function

	'// 검색어배열
	public function getItemKeywordArray(sno)
		dim arrRst, arrRst2
		if trim(Fkeywords)="" then exit Function

		arrRst = split(Fkeywords,",")
		if ubound(arrRst)=0 then
			'구분이 공백일 경우
			arrRst2 = split(arrRst(0)," ")
			if ubound(arrRst2)>0 then
				arrRst = split(Fkeywords," ")
			end if
		end if

		if ubound(arrRst)>=sno then
			getItemKeywordArray=trim(arrRst(sno))
		else
			getItemKeywordArray=""
		end if
	end function

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
				End If
			End If

			'키워드 1
			divBound1 = Cint(Ubound(arrRst)/3)
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
			strRst = strRst & "&sch_kwd_1_nm=" & Server.URLEncode(Keyword1)
			strRst = strRst & "&sch_kwd_2_nm=" & Server.URLEncode(Keyword2)
			strRst = strRst & "&sch_kwd_3_nm=" & Server.URLEncode(Keyword3)
			getItemKeyword = strRst
		Else
			strRst = ""
			strRst = strRst & "&sch_kwd_1_nm="&Server.URLEncode(Fkeywords)
			strRst = strRst & "&sch_kwd_2_nm="
			strRst = strRst & "&sch_kwd_3_nm="
			getItemKeyword = strRst
		End If
	End Function

	'// 상품등록 파라메터 생성
	public Function getLotteItemRegParameter(isEdit)
		dim strRst
		strRst = "subscriptionId=" & lotteAuthNo																						'롯데닷컴 인증번호	(*)
		if (isEdit) then
		   strRst = strRst & "&goods_req_no="&FLotteTmpGoodNo
		end if
		strRst = strRst & "&brnd_no=" & tenBrandCd																						'브랜드코드			(*)
'		strRst = strRst & "&goods_nm=" & Server.URLEncode(Trim(getItemNameFormat))																	'상품명				(*)
		strRst = strRst & "&goods_nm=" & Server.URLEncode(Fsocname_kor) & " " & Server.URLEncode(Trim(getItemNameFormat))
'		strRst = strRst & "&sch_kwd_1_nm=" & Server.URLEncode(getItemKeywordArray(0))													'키워드1
'		strRst = strRst & "&sch_kwd_2_nm=" & Server.URLEncode(getItemKeywordArray(1))													'키워드2
'		strRst = strRst & "&sch_kwd_3_nm=" & Server.URLEncode(getItemKeywordArray(2))													'키워드3
		strRst = strRst & getItemKeyword												'키워드 //2013-08-09 김진영 수정
		strRst = strRst & "&pmct_fix_cd=2"																					 			'원가고정여부		(*:마진율고정)
		strRst = strRst & "&pur_shp_cd=3" 		'' 2=>3 특정																						'매출형태			(*:판매분매입)
		strRst = strRst & "&sale_shp_cd=10" 																							'판매형태			(*:정상)
		strRst = strRst & "&sale_prc=" & cLng(GetRaiseValue(FSellCash/10)*10)																				'판매가(원판매가)	(*) 소비자가로 설정한 이유..?
		strRst = strRst & "&mrgn_rt=12" 																								'마진율				(*:11%) ==> 2013/01/01 12%
		strRst = strRst & "&pur_prc=" & cLng(FSellCash*0.88)																			'공급가				(*)
		strRst = strRst & "&tdf_sct_cd=1" 																								'과면세코드			(*:과세)
		strRst = strRst & "&card_dsct_yn=Y"																					 			'롯데카드할인여부	(*:사용)
		if getLotteSellYn="Y" then										 																'판매상태			(*:10:판매,20:품절)
			strRst = strRst & "&sale_stat_cd=10"
		else
			strRst = strRst & "&sale_stat_cd=20"
		end if

		IF (FLimitYn="Y") then
		    strRst = strRst & "&inv_mgmt_yn=Y"
		    if FoptionCnt=0 then
		        strRst = strRst & "&inv_qty="&getLimitLotteEa()
		    end if
		ELSE
		    ''2013/06/12 재고관리여부 모두 Y로 변경
    		strRst = strRst & "&inv_mgmt_yn=Y" 																							'재고관리여부		(*:관리안함)
    		if FoptionCnt=0 then
    		    strRst = strRst & "&inv_qty="&CDEFALUT_STOCK
    		end if
    	END IF

		if FitemDiv="06" then
			strRst = strRst & "&add_choc_tp_cd_20=" & Server.URLEncode("주문제작상품") 													'문구입력형옵션
		end if
		strRst = strRst & "&dlv_proc_tp_cd=1" 																							'배송유형			(*:직송)
		strRst = strRst & "&box_pkg_yn=Y" 																								'배송Box포장		(*:포장가능)
		strRst = strRst & "&fut_msg_yn=N" 																								'선물포장여부		(*:불가)
		strRst = strRst & "&shop_fwd_msg_yn=N"	 																						'매장전달			(*:사용안함)
		strRst = strRst & "&dlv_mean_cd=10" 																							'배송수단			(*:택배)

    	strRst = strRst & getLotteGoodDLVDtParams
		strRst = strRst & "&dlvp_stn_grp_cd=01" 																						'배송가능지역		(*:전국)
		strRst = strRst & "&byr_age_lmt_cd=0" 																							'구입자나이제한		(*:전체)
		strRst = strRst & "&exch_rtgs_sct_cd=21" 																						'교환반품구분		(*:분담교환)
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd																					'배송정책번호		(*:전체)
		strRst = strRst & "&corp_dlvp_sn=" & getLotteReturnDLVCode 																			 			'반품배송지코드		(*:물류센터)
		strRst = strRst & "&dcom_asgn_rtgs_hdc_use_yn=Y"                                                                 ''반품교환 지정택배 사용여부  dcom_asgn_rtgs_hdc_use_yn
		strRst = strRst & "&orpl_nm=" & Server.URLEncode(chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea))	'원산지				(*)
		strRst = strRst & "&mfcp_nm=" & Server.URLEncode(chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername))		'제조사				(*)
		strRst = strRst & "&img_url=" & Server.URLEncode(FbasicImage)																	'대표이미지URL		(*)
		strRst = strRst & "&attd_mtr_cont=" & Server.URLEncode(ForderComment)															'주문시 주의사항(-)
		'strRst = strRst & "&md_ntc_2_fcont=" & server.URLEncode("<img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg"">")	'MD공지
		strRst = strRst & "&brnd_intro_cont=" & server.URLEncode("Design Your Life! 새로운 일상을 만드는 감성생활브랜드 텐바이텐")		'브랜드 설명
		strRst = strRst & "&corp_goods_no=" & Fitemid																					'업체상품번호(텐텐상품코드)
		strRst = strRst & getLotteCateParamToReg()																						'MD상품군 및 해당 전시카테고리
		strRst = strRst & getLotteOptionParamToReg()																					'상품옵션 여부 및 옵션내용
		strRst = strRst & getLotteItemContParamToReg()																					'상품상세설명		(*)
		strRst = strRst & getLotteAddImageParamToReg()																					'상품 추가 이미지
        strRst = strRst & getLotteItemInfoCdToReg()   ''진영
		'결과 반환
		getLotteItemRegParameter = strRst
	end Function

    public Function getLotteItemEditParameter2()
        dim strRst
        strRst = getLotteItemRegParameter(true)

        getLotteItemEditParameter2 = strRst
    end function

	'// 상품수정 파라메터 생성
	public Function getLotteItemEditParameter()
		dim strRst
		strRst = "subscriptionId=" & lotteAuthNo																						'롯데닷컴 인증번호	(*)
		strRst = strRst & "&goods_no=" & FLotteGoodNo																					'롯뎃닷컴 상품번호	(*)
		strRst = strRst & "&brnd_no=" & tenBrandCd																						'브랜드코드			(*)
		strRst = strRst & "&goods_nm=" & Server.URLEncode(Fsocname_kor) & " " & Server.URLEncode(Trim(getItemNameFormat))				'상품명				(*)
		''if (FItemid="443553") then strRst = strRst & "."
'		strRst = strRst & "&sch_kwd_1_nm=" & Server.URLEncode(getItemKeywordArray(0))													'키워드1
'		strRst = strRst & "&sch_kwd_2_nm=" & Server.URLEncode(getItemKeywordArray(1))													'키워드2
'		strRst = strRst & "&sch_kwd_3_nm=" & Server.URLEncode(getItemKeywordArray(2))													'키워드3
		strRst = strRst & getItemKeyword																								'키워드 //2013-08-09 김진영 수정
		strRst = strRst & "&pur_shp_cd=2" 																								'매출형태			(*:판매분매입)

		''getLotteOptionParamToEdit 에 포함
'		if getLotteSellYn="Y" then										 																'판매상태			(*:10:판매,20:품절)
'			strRst = strRst & "&sale_stat_cd=10"
'		else
'			strRst = strRst & "&sale_stat_cd=20"
'		end if

		''getLotteOptionParamToEdit 에 포함
'		IF (FLimitYn="Y") then
'		    strRst = strRst & "&inv_mgmt_yn=Y"
'		    if FoptionCnt=0 then
'		        strRst = strRst & "&inv_qty="&getLimitLotteEa()
'		    end if
'		ELSE
'    		strRst = strRst & "&inv_mgmt_yn=N" 																								'재고관리여부		(*:관리안함)
'    	END IF

		if FitemDiv="06" then
		    ''수정 쪽에 이게 들어가믄 오류..? // 이값이 없어도 주문문구 옵션이 사라지지는 않는듯.
			'strRst = strRst & "&add_choc_tp_cd_20=" & Server.URLEncode("주문제작상품") 													'문구입력형옵션
	        ''strRst = strRst & "&add_choc_tp_cd_20="
		end if
		strRst = strRst & "&dlv_proc_tp_cd=1" 																							'배송유형			(*:직송)
		strRst = strRst & "&box_pkg_yn=Y" 																								'배송Box포장		(*:포장가능)
		strRst = strRst & "&fut_msg_yn=N" 																								'선물포장여부		(*:불가)
		strRst = strRst & "&shop_fwd_msg_yn=N"	 																						'매장전달			(*:사용안함)
		strRst = strRst & "&dlv_mean_cd=10" 																							'배송수단			(*:택배)

    	strRst = strRst & getLotteGoodDLVDtParams
		strRst = strRst & "&dlvp_stn_grp_cd=01" 																						'배송가능지역		(*:전국)
		strRst = strRst & "&byr_age_lmt_cd=0" 																							'구입자나이제한		(*:전체)
		strRst = strRst & "&exch_rtgs_sct_cd=21" 																						'교환반품구분		(*:분담교환)
		strRst = strRst & "&dlv_polc_no=" & tenDlvCd																					'배송정책번호		(*:전체)
		strRst = strRst & "&corp_dlvp_sn=" & getLotteReturnDLVCode 																			 			'반품배송지코드		(*:물류센터)
		strRst = strRst & "&dcom_asgn_rtgs_hdc_use_yn=Y"                                                                                               ''반품교환 지정택배 사용여부 2013/02/26
		strRst = strRst & "&orpl_nm=" & Server.URLEncode(chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea))	'원산지				(*)
		strRst = strRst & "&mfcp_nm=" & Server.URLEncode(chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername))		'제조사				(*)
		strRst = strRst & "&img_url=" & Server.URLEncode(FbasicImage)																	'대표이미지URL		(*)
		'strRst = strRst & "&md_ntc_2_fcont=" & server.URLEncode("<img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg"">")	'MD공지
		strRst = strRst & "&brnd_intro_cont=" & server.URLEncode("Design Your Life! 새로운 일상을 만드는 감성생활브랜드 텐바이텐")		'브랜드 설명
		if (Fitemid="409536") then
		    strRst = strRst & "&attd_mtr_cont="
		else
    		strRst = strRst & "&attd_mtr_cont=" & Server.URLEncode(ForderComment)															'주문시 주의사항
    	end if
		strRst = strRst & "&corp_goods_no=" & Fitemid																					'업체상품번호(텐텐상품코드)
		strRst = strRst & getLotteItemContParamToReg()																					'상품상세설명		(*)
		strRst = strRst & getLotteAddImageParamToReg()																					'상품 추가 이미지
        strRst = strRst & getLotteOptionParamToEdit()
		strRst = strRst & getLotteItemInfoCdToReg()																						'상품품목정보/2012-11-02김진영 생성
        ''strRst = strRst & getLotteCateParamToReg() ''''''''20120831

		'결과 반환
		getLotteItemEditParameter = strRst
	end Function

    '// 가격 수정 파라메터 생성
    public Function getLotteItemPriceEditParameter()
        ''http://openapidev.lotte.com/openapi/updateGoodsNmOpenApi.lotte?subscriptionId=[인증키]&strGoodsNo=XXXXXX&strReqSalePrc=[상품가격]&strChgCausCont=[변경사유]
        dim strRst
        strRst = "subscriptionId=" & lotteAuthNo
        strRst = strRst & "&strGoodsNo=" & FLotteGoodNo
        strRst = strRst & "&strReqSalePrc=" & GetRaiseValue(MustPrice/10)*10
        ''strRst = strRst & "&mrgn_rt=12"
        ''strRst = strRst & "&pur_prc=" & cLng(FSellCash*0.88)
        ''strRst = strRst & "&strChgCausCont=" & Server.URLEncode("가격변동")

        getLotteItemPriceEditParameter = strRst
    end function

	Public Function MustPrice
		Dim GetTenTenMargin
		'2013-07-25 김진영//텐텐마진이 닷컴의 마진보다 작을 때 orgprice로 전송 시작
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If GetTenTenMargin < CMAXMARGIN Then
			MustPrice = Forgprice
		Else
			MustPrice = FSellCash
		End If
		'2013-07-25 김진영//텐텐마진이 닷컴의 마진보다 작을 때 orgprice로 전송 끝
	End Function


    ''//상품명 변경 파라메터 생성
    public Function getLotteItemNameEditParameter()
        dim strRst
        strRst = "subscriptionId=" & lotteAuthNo
        strRst = strRst & "&strGoodsNo=" & FLotteGoodNo
        strRst = strRst & "&strGoodsNm=" & Server.URLEncode(Trim(getItemNameFormat))
        strRst = strRst & "&strMblGoodsNm=" & Server.URLEncode(Trim(getItemNameFormat))
        strRst = strRst & "&strChgCausCont=" & Server.URLEncode("api 상품명 변경")
        getLotteItemNameEditParameter = strRst
    end function

    public function getLotteGoodDLVDtParams()
        dim strRst
        strRst = ""

		If Fmakerid = "bugangit10" Then		 '2013-07-26 현아주임님 요청..특정 브랜드는 무조건 일반으로 돌리라고 해서..
		    strRst = strRst & "&dlv_goods_sct_cd=01"
    		strRst = strRst & "&dlv_dday=3"
    		getLotteGoodDLVDtParams = strRst
    		Exit Function
		End If

		If Fitemid = "305876" OR Fitemid = "303848" OR Fitemid = "305877" OR Fitemid = "305878" Then
		    strRst = strRst & "&dlv_goods_sct_cd=03"
    		strRst = strRst & "&dlv_dday=10"
    		getLotteGoodDLVDtParams = strRst
			Exit Function
		End If

        if (FtenCateLarge="055") or (FtenCateLarge="040") then ''가구/패브릭 15일로
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=15"
		elseif (FtenCateLarge="080") or (FtenCateLarge="100") then  ''우먼/베이비 5일
		    strRst = strRst & "&dlv_goods_sct_cd=03" 																						'배송상품구분		(*:주문제작03)
		    strRst = strRst & "&dlv_dday=5"
		elseif ((FtenCateLarge="045") and (FtenCateMid="001" or FtenCateMid="004")) then  ''수납/생활> 옷/이불수납 or 주방수납 10일 - 현아씨요청 2013/01/22
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
		elseif ((FtenCateLarge="025") and (FtenCateMid="107")) then  ''디지털 > 기타 스마트기기 케이스  10일 - 현아씨요청 2013/01/22
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
	    elseif ((FtenCateLarge="050") and (FtenCateMid="777")) then   ''홈/데코 > 거울   - 미희씨요청 2013/03/08
	        strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
		elseif ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="001")) then    ''HOME > 수납/생활 > 보관/정리용품 > 수납장 			주문제작15일 045&cdm=002&cds=001
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=15"
		elseif ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="002")) then    ''HOME > 수납/생활 > 보관/정리용품 > 틈새수납장			주문제작10일
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="045") and (FtenCateMid="002") and (FtenCateSmall="005")) then    ''HOME > 수납/생활 > 보관/정리용품 > 잡지꽂이 			주문제작10일
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
		elseif ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="001")) then    ''HOME > 수납/생활 > 데코수납 > 우드박스 				주문제작10일
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="045") and (FtenCateMid="006") and (FtenCateSmall="007")) then    ''HOME > 수납/생활 > 데코수납 > 인터폰박스 			               주문제작10일
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="050") and (FtenCateMid="060") and (FtenCateSmall="070")) then    ''HOME > 홈/데코 > 소품박스/바구니 > 인터폰박스			주문제작10일 cdl=050&cdm=060&cds=070
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="110") and (FtenCateMid="090") and (FtenCateSmall="040")) then    ''HOME > 감성채널 > DIY > 나무로만들기 				주문제작10일 110&cdm=090&cds=040
            strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
        elseif ((FtenCateLarge="045") and (FtenCateMid="010")) then   ''수납/생활 > 디자인선반  - 미희씨요청 2013/03/08
	        strRst = strRst & "&dlv_goods_sct_cd=03"
		    strRst = strRst & "&dlv_dday=10"
'		elseif (FtenCateLarge="025")  then  ''디지털 10일 - 미희씨요청 2013/01/17
'		    strRst = strRst & "&dlv_goods_sct_cd=03" 																						'배송상품구분		(*:주문제작03)
'		    strRst = strRst & "&dlv_dday=10"
		elseif ((FitemDiv="06") or (FitemDiv="16")) then    ''주문(후)제작상품
		    strRst = strRst & "&dlv_goods_sct_cd=03"
		    if (FrequireMakeDay>7) then
		        strRst = strRst & "&dlv_dday="&CStr(FrequireMakeDay)
		    elseif (FrequireMakeDay<1) then
		        strRst = strRst & "&dlv_dday=7"
		    else
		        strRst = strRst & "&dlv_dday="&(FrequireMakeDay+1)
		    end if
		else
		    strRst = strRst & "&dlv_goods_sct_cd=01" 																						'배송상품구분		(*:일반상품)
    		strRst = strRst & "&dlv_dday=3" 																								'배송기일			(*:3일이내)
    	end if

    	getLotteGoodDLVDtParams = strRst
    end function

	'// 상품등록: MD상품군 및 전시 카테고리 파라메터 생성(상품등록용)
	public function getLotteCateParamToReg()
		dim strSql, strRst, i, ogrpCode
		strSql = "Select top 6 c.groupCode, m.dispNo, c.disptpcd "
		strSql = strSql & " from db_item.dbo.tbl_lotte_cate_mapping as m "
		strSql = strSql & " 	join db_temp.dbo.tbl_lotte_Category as c "
		strSql = strSql & " 		on m.DispNO=c.DispNO "
		strSql = strSql & " where tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " 	and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " 	and tenCateSmall='" & FtenCateSmall & "' "
	''strSql = strSql & " 	and c.disptpcd<>'99'"
	    strSql = strSql & " 	and c.isusing='Y'"
		strSql = strSql & " order by (CASE WHEN c.disptpcd='12' THEN 'ZZ' ELSE c.disptpcd END) desc"           ''''//전문몰을 기본 카테고리로..
		rsget.Open strSql,dbget,1

		if Not(rsget.EOF or rsget.BOF) then
		    ogrpCode = rsget("groupCode")
			strRst = "&md_gsgr_no=" & ogrpCode				            'md상품군 코드 (md상품군 코드 같은 카테고리만 등록가능)
            ''if (rsget("groupCode")="1598") then
            ''    strRst = "&md_gsgr_no=" & "1596"
            ''end if

			i=0
			Do until rsget.EOF
				''if i=0 then
				''	strRst = strRst & "&disp_no=" & rsget("dispNo")		'기본 전시카테고리
				''else
				''	strRst = strRst & "&disp_no_b=" & rsget("dispNo") 	'추가 전시카테고리
				''end if

				if (rsget("disptpcd")="12") then                        ''전문몰 카테고리를 기본으로 하라고.. /2012/06/14
				    strRst = strRst & "&disp_no=" & rsget("dispNo")		'기본 전시카테고리
				else
				    IF (ogrpCode=rsget("groupCode")) then
    				    strRst = strRst & "&disp_no_b=" & rsget("dispNo") 	'추가 전시카테고리
    				End IF
			    end if
				rsget.MoveNext
				i=i+1
			Loop
		end if

		rsget.Close

		getLotteCateParamToReg = strRst
	end function

	'//전시 카테고리 파라메터 수정(상품수정용)
	Public Function getLotteCateParamToEdit()
		Dim strSql, strRst, i, ogrpCode
		strRst = "subscriptionId=" & lotteAuthNo											'(*)사용자인증키
		strRst = strRst & "&strGoodsNo=" & FLotteGoodNo										'(*)롯데닷컴 상품번호

		strSql = ""
		strSql = strSql & " SELECT TOP 6 c.groupCode, m.dispNo, c.disptpcd "
		strSql = strSql & " FROM db_item.dbo.tbl_lotte_cate_mapping as m "
		strSql = strSql & " INNER JOIN db_temp.dbo.tbl_lotte_Category as c on m.DispNO = c.DispNO "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
	    strSql = strSql & " and c.isusing='Y'"
		strSql = strSql & " order by (CASE WHEN c.disptpcd='12' THEN 'ZZ' ELSE c.disptpcd END) desc"           ''''//전문몰을 기본 카테고리로..
		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) Then
		    ogrpCode = rsget("groupCode")
			i = 0
			Do until rsget.EOF
				If (rsget("disptpcd")="12") then
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
		getLotteCateParamToEdit = strRst
	End Function

    public function getLotteOptionParamToEditNew()
        dim ret : ret = ""
        dim i
        dim strSql, arrRows, iErrStr
        dim isOptionExists
        dim mayOptionCnt : mayOptionCnt = 0
        dim item_sale_stat_cd,outmalloptcode, optLimit
        dim item_noStr, item_sale_stat_cdStr, inv_qtyStr
        dim optValidExists : optValidExists = false
        dim preMaxOutmalloptcode : preMaxOutmalloptcode=-1

        strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_Lotte '"&CMallName&"'," & FItemid
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
        rsget.Open strSql, dbget
        if Not(rsget.EOF or rsget.BOF) then
            arrRows = rsget.getRows
        end if
        rsget.close

        isOptionExists = isArray(arrRows)
        if (isOptionExists) then
            mayOptionCnt = UBound(ArrRows,2)
            mayOptionCnt = mayOptionCnt + 1
        end if

        ''if (FoptionCnt<>mayOptionCnt) then
        if (FregedOptCnt<>mayOptionCnt) then
            ''재고조회.
            rw "FregedOptCnt="&FregedOptCnt&".."&"mayOptionCnt="&mayOptionCnt

            CALL LotteOneItemCheckStock(Fitemid,iErrStr)
        end if

        ret = ""
        if (Not isOptionExists) then
            IF (FLimitYn="Y") then
    		    ret = ret & "&inv_mgmt_yn=Y"
    		    ret = ret & "&item_no=0"
    		    ret = ret & "&item_sale_stat_cd=10"
    		    ret = ret & "&inv_qty="&getLimitLotteEa()                                                                                   '재고관리 할 경우 단품이 있어야..
    		ELSE
    		    ''2013/06/12 재고관리여부 모두 Y로 변경
        		ret = ret & "&inv_mgmt_yn=Y" 																								'재고관리여부
    		    ret = ret & "&item_no=0"												''(추가/아래라인까지..)2013-07-16 김진영 이 항목 필수..없으니 재고수정이 안 됨
    		    ret = ret & "&item_sale_stat_cd=10"
        		ret = ret & "&inv_qty="&CDEFALUT_STOCK  ''2013/06/12추가
        	END IF
        else
            if FLimitYn="Y" then
			    ret = ret&"&inv_mgmt_yn=Y"
			else
			    ret = ret&"&inv_mgmt_yn=Y" ''2013/06/12 수정
		    end if

            For i =0 To UBound(ArrRows,2)
                item_sale_stat_cd = "10"    ''10:판매진행,20:품절,30:판매종료
			    outmalloptcode = ArrRows(2,i)
			    if IsNULL(outmalloptcode) then
			        outmalloptcode=preMaxOutmalloptcode+1
			    else
			        if (preMaxOutmalloptcode>outmalloptcode) then
			            preMaxOutmalloptcode=preMaxOutmalloptcode
			        else
			            preMaxOutmalloptcode=outmalloptcode
			        end if
			    end if

			    if FLimitYn="Y" then
			        optLimit = ArrRows(4,i)-5
	        		'2013-10-10 진영 추가..한정 중 1000개가 넘으면 최대 999로 넣겠다고 롯데에서 요청메일 옴
			        If optLimit > 1000 Then
			        	optLimit = 999
			        End If
			    else
			        optLimit = CDEFALUT_STOCK
			    end if
			    if (optLimit<1) then optLimit=0
			    if (ArrRows(6,i)="N") or (ArrRows(7,i)="N") then item_sale_stat_cd="20"
			    if (FLimitYn="Y") and (optLimit<1) then item_sale_stat_cd="20"

			    if ((ArrRows(11,i)="1") and (ArrRows(12,i)="1")) or (ArrRows(13,i)="1") then
			        optLimit=0
			        item_sale_stat_cd="20"
			    end if

			    item_noStr = item_noStr & "&item_no="&outmalloptcode
			    item_sale_stat_cdStr = item_sale_stat_cdStr & "&item_sale_stat_cd="&item_sale_stat_cd
			    inv_qtyStr = inv_qtyStr & "&inv_qty="&optLimit
			    if (item_sale_stat_cd = "10") then optValidExists=TRUE
            next

            ret = ret&item_noStr&item_sale_stat_cdStr&inv_qtyStr
        end if

	    ''rw ret
	    if (Not isOptionExists) then   ''옵션이 없으면.
    		if getLotteSellYn="Y" then										 																'판매상태			(*:10:판매,20:품절)
    			ret = ret & "&sale_stat_cd=10"
    		else
    		    FSellyn="N"
    			ret = ret & "&sale_stat_cd=20"
    		end if
    	else
    	    if (optValidExists) and (getLotteSellYn="Y") then  ''판매중 이고 옵션 판매가능이면.
    	        ret = ret & "&sale_stat_cd=10"
    	    else
    	        rw "None Exists Valid Option"
    	        FSellyn="N"
    	        ret = ret & "&sale_stat_cd=20"
    	    end if
        end if

        getLotteOptionParamToEditNew = ret
    end function

    '// 상품수정: 옵션 파라메터 생성(상품수정용) :: 기 옵션이 정의 되어 있어야 함..
    ''옵션명 수정이 안됨.. 관련 API 있는지 문의.
    public function getLotteOptionParamToEdit()
        if (TRUE) or (FItemID="138371") or (FItemID="295139") or (FItemID="830724") or (FItemID="830728") or (FItemID="830816") or (FItemID="830795") then
            getLotteOptionParamToEdit=getLotteOptionParamToEditNew
            exit function
        end if

        dim ret : ret = ""
        dim strSql
        dim optYn, item_noStr, item_sale_stat_cdStr, inv_qtyStr, optLimit, item_sale_stat_cd, outmalloptcode
        dim optValidExists : optValidExists = FALSE
        ''getLotteOptionParamToEdit = ret
        ''Exit function

        ''옵션이 있으나 롯데 저장된 옵션 내역이 없으면..
        if (FoptionCnt>0) then
            ''재고조회.
        end if

'        if (FItemid=379468) then
'            ret = ret&"&inv_mgmt_yn=Y"
'            ret = ret&"&item_no=0&item_no=1&item_no=2"
'            ret = ret&"&item_sale_stat_cd=20&item_sale_stat_cd=20&item_sale_stat_cd=10"
'            ret = ret&"&inv_qty=0&inv_qty=0&inv_qty=2"
'        end if
'
'        if (FItemid=330185) then
'            ret = ret&"&inv_mgmt_yn=Y"
'            ret = ret&"&item_no=0&item_no=1&item_no=2&item_no=3&item_no=4&item_no=5&item_no=6&item_no=7&item_no=8&item_no=9"
'            ret = ret&"&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10"
'            ret = ret&"&inv_qty=10&inv_qty=10&inv_qty=10&inv_qty=10&inv_qty=5&inv_qty=5&inv_qty=10&inv_qty=10&inv_qty=10&inv_qty=10"
'        end if

        ''553443

        ''&inv_mgmt_yn=Y&item_no=0&item_no=1&item_no=2&item_no=3&item_no=4&item_no=5&item_no=6&item_no=7&item_no=8&item_no=9&item_no=10&item_no=11&item_no=12&item_no=13&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=20&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&item_sale_stat_cd=10&inv_qtyStr=5&inv_qtyStr=6&inv_qtyStr=9&inv_qtyStr=7&inv_qtyStr=11&inv_qtyStr=5&inv_qtyStr=8&inv_qtyStr=3&inv_qtyStr=27&inv_qtyStr=0&inv_qtyStr=13&inv_qtyStr=11&inv_qtyStr=4&inv_qtyStr=10


        strSql = "Select o.itemoption, (CASE WHEN convert(varchar(18),o.optionTypeName)<>o.optionTypeName THEN '옵션선택' ELSE o.optionTypeName END) as optionTypeName, o.optionname, (o.optlimitno-o.optlimitsold) as optLimit, o.optlimityn, o.isUsing, o.optsellyn "
		strSql = strSql & " ,R.outmalloptcode"
		strSql = strSql & " ,(CASE WHEN  R.itemoption is NULL THEN 0 ELSE 1 END) as preged"
		strSql = strSql & " From [db_item].[dbo].tbl_item_option o"
		strSql = strSql & "     left join [db_item].[dbo].tbl_OutMall_regedoption R"
		strSql = strSql & "     on o.itemid=R.itemid"
		strSql = strSql & "     and o.itemoption=R.itemoption"
		strSql = strSql & "     and R.itemoption<>''"
		strSql = strSql & "     and R.mallid='"&CMALLNAME&"'"
		strSql = strSql & " where o.itemid=" & FItemid
		strSql = strSql & " 	and o.optaddprice=0 "                     '''추가금액 불가.
		strSql = strSql & " 	and ((o.isUsing='Y') or (R.itemid is Not NULL)) "  '
		strSql = strSql & " 	and R.outmalloptcode is Not NULL"         '''등록 옵션만 수정 가능. :: 애초에 모든 옵션을 등록해야할듯..
		strSql = strSql & " and isNULL(R.outmallsellyn,'')<>'X'"

		'''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "

		rsget.Open strSql,dbget,1

		if Not(rsget.EOF or rsget.BOF) then
			optYn = "Y"
			if FLimitYn="Y" then
			    ret = ret&"&inv_mgmt_yn=Y"
			else
			    ret = ret&"&inv_mgmt_yn=Y"
		    end if

			Do until rsget.EOF
			    item_sale_stat_cd = "10"    ''10:판매진행,20:품절,30:판매종료
			    outmalloptcode = rsget("outmalloptcode")
			    optLimit = rsget("optLimit")-5
                if (FLimitYn<>"Y") then optLimit=CDEFALUT_STOCK

			    if (optLimit<1) then optLimit=0
			    if (rsget("isUsing")="N") or (rsget("optsellyn")="N") then item_sale_stat_cd="20"
			    if (FLimitYn="Y") and (optLimit<1) then item_sale_stat_cd="20"

			    item_noStr = item_noStr & "&item_no="&outmalloptcode
			    item_sale_stat_cdStr = item_sale_stat_cdStr & "&item_sale_stat_cd="&item_sale_stat_cd
			    inv_qtyStr = inv_qtyStr & "&inv_qty="&optLimit

			    if (item_sale_stat_cd = "10") then optValidExists=TRUE
			    rsget.MoveNext
			Loop
		end if
		rsget.Close


		if optYn <> "Y" then
    		IF (FLimitYn="Y") then
    		    ret = ""
    		    ret = ret & "&inv_mgmt_yn=Y"
    		    ret = ret & "&item_no=0"
    		    ret = ret & "&item_sale_stat_cd=10"
    		    ret = ret & "&inv_qty="&getLimitLotteEa()
    		ELSE
        		ret = ret & "&inv_mgmt_yn=Y" 																								'재고관리여부
        		ret = ret & "&inv_qty="&CDEFALUT_STOCK
        	END IF

		else
		    ret = ret&item_noStr&item_sale_stat_cdStr&inv_qtyStr
	    end if

	    ''rw ret
	    if optYn <> "Y" then   ''옵션이 없으면.
    		if getLotteSellYn="Y" then										 																'판매상태			(*:10:판매,20:품절)
    			ret = ret & "&sale_stat_cd=10"
    		else
    			ret = ret & "&sale_stat_cd=20"
    		end if
    	else
    	    if (optValidExists) and (getLotteSellYn="Y") then  ''판매중 이고 옵션 판매가능이면.
    	        ret = ret & "&sale_stat_cd=10"
    	    else
    	        ret = ret & "&sale_stat_cd=20"
    	    end if
        end if

        getLotteOptionParamToEdit = ret
    end function

	'// 상품등록: 옵션 파라메터 생성(상품등록용)
	public function getLotteOptionParamToReg()
		dim strSql, strRst, i, optYn, optNm, optDc, chkMultiOpt, optLimit
		chkMultiOpt = false
		optYn = "N"

		if FoptionCnt>0 then
			'// 이중옵션일 때
			'#옵션명 생성
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget

			optNm = ""
			if Not(rsget.EOF or rsget.BOF) then
				chkMultiOpt = true
				optYn = "Y"
				Do until rsget.EOF
					optNm = optNm & Replace(db2Html(rsget("optionTypeName")),":","")
					rsget.MoveNext
					if Not(rsget.EOF) then optNm = optNm & ":"
				Loop
			end if
			rsget.Close

			'#옵션내용 생성
			if chkMultiOpt then
				strSql = "Select optionname, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				'''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) " ''일단 입력

				rsget.Open strSql,dbget,1

				optDc = ""
				if Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit-5
					    if (optLimit<1) then optLimit=0
					    if (FLimitYN<>"Y") then optLimit=CDEFALUT_STOCK   ''2013/06/12 재고관리여부 모두 Y로 변경 되므로

						optDc = optDc & Replace(Replace(db2Html(rsget("optionname")),":",""),"'","") & "," & optLimit
						rsget.MoveNext
						if Not(rsget.EOF) then optDc = optDc & ":"
					Loop
				end if
				rsget.Close
			end if


			'// 단일옵션일 때
			if Not(chkMultiOpt) then
				strSql = "Select optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				if Not(rsget.EOF or rsget.BOF) then
					optYn = "Y"
					if db2Html(rsget("optionTypeName"))<>"" then
						optNm = Replace(db2Html(rsget("optionTypeName")),":","")
					else
						optNm = "옵션"
					end if
					Do until rsget.EOF
					    optLimit = rsget("optLimit")
					    optLimit = optLimit-5
					    if (optLimit<1) then optLimit=0
					    if (FLimitYN<>"Y") then optLimit=CDEFALUT_STOCK   ''2013/06/12 재고관리여부 모두 Y로 변경 되므로

						optDc = optDc & Replace(Replace(Replace(db2Html(rsget("optionname")),":",""),",",""),"'","") & "," & optLimit
						rsget.MoveNext
						if Not(rsget.EOF) then optDc = optDc & ":"
					Loop
				end if
				rsget.Close
			end if
		end if

		strRst = strRst & "&item_mgmt_yn=" & optYn						'단품관리여부(옵션)
		strRst = strRst & "&opt_nm=" & server.URLEncode(optNm)			'옵션명
		strRst = strRst & "&item_list=" & server.URLEncode(optDc)		'옵션상세

		getLotteOptionParamToReg = strRst
	end function

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	public function getLotteItemContParamToReg()
		dim strRst, strSQL
		strRst = Server.URLEncode("<div align=""center"">")
		'2014-01-17 10:00 김진영 탑 이미지 추가
		strRst = strRst & Server.URLEncode("<p><a href=""http://www.lotte.com/display/viewDispShop.lotte?disp_no=5293948"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_lotteCom.jpg""></a></p>")
		'#기본 상품설명
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & Server.URLEncode(oLotteitem.FItemList(i).Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & Server.URLEncode(nl2br(oLotteitem.FItemList(i).Fitemcontent) & "<br>")
			Case Else
				strRst = strRst & Server.URLEncode(nl2br(ReplaceBracket(oLotteitem.FItemList(i).Fitemcontent)) & "<br>")
		End Select

		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		if Not(rsget.EOF or rsget.BOF) then
			Do Until rsget.EOF
				if rsget("imgType")="1" then
					strRst = strRst & Server.URLEncode("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0""><br>")
				end if
				rsget.MoveNext
			Loop
		end if

		rsget.Close

		'#기본 상품 설명이미지
		if ImageExists(FmainImage) then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage & """ border=""0""><br>")
		if ImageExists(FmainImage2) then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage2 & """ border=""0""><br>")

		'#배송 주의사항
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg"">")

		strRst = strRst & Server.URLEncode("</div>")

		getLotteItemContParamToReg = "&dtl_info_fcont=" & strRst

		''660877 db_item.dbo.tbl_OutMall_etcLink 쿼리 후 값이 존재하면 시작
		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.Open strSQL, dbget
		if Not(rsget.EOF or rsget.BOF) then
			strRst = Server.URLEncode(""&rsget("textVal")&"")
			strRst = Server.URLEncode("<div align=""center""><p><a href=""http://www.lotte.com/display/viewDispShop.lotte?disp_no=5293948"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_lotteCom.jpg""></a></p>") & strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info.jpg""></div>")
			getLotteItemContParamToReg = "&dtl_info_fcont=" & strRst
		End If
		rsget.Close
		''660877 db_item.dbo.tbl_OutMall_etcLink 쿼리 후 값이 존재하면 끝

'		if (FItemID="502049") or (FItemID="660877") then
'		    ''getLotteItemContParamToReg = "&dtl_info_fcont="&Server.URLEncode("<div></div>")     ''
'		    getLotteItemContParamToReg = ""                                           '' 연동안하려면 파라메터 없이 또는빈값
'		end if
	end function

	'// 상품등록: 상품추가이미지 파라메터 생성(상품등록용)
	public function getLotteAddImageParamToReg()
		dim strRst, strSQL, i

		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		if Not(rsget.EOF or rsget.BOF) then
			for i=1 to rsget.RecordCount
				if rsget("imgType")="0" then
					strRst = strRst & "&img_url" & i & "=" & Server.URLEncode("http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400"))
				end if
				rsget.MoveNext
				if i>=5 then Exit For
			next
		end if

		rsget.Close

		getLotteAddImageParamToReg = strRst
	end function

    ''국산여부
    public Function isMadeInKorea(iVal)
        isMadeInKorea = False
        if isNULL(iVal) then Exit Function
        iVal = Trim(iVal)

        if (iVal="한국") or (iVal="대한민국") or (iVal="KOREA") or (iVal="KOREA 한국") or (iVal="KOREA 국내산") then
            isMadeInKorea = True
        end if

        if (iVal="국산") or (iVal="국내제작") or (iVal="한국OEM") or (iVal="국내제작(한국)") or (iVal="국내") then
            isMadeInKorea = True
        end if

        if (iVal="한국 / 코지아트") then
            isMadeInKorea = True
        end if
    end Function

	'2012/11/02 김진영 생성 상품품목정보 파라메타
	Public Function getLotteItemInfoCdToReg()
		Dim strRst, strSQL
		Dim anjunInfo, mallinfoDiv, mallinfoCdAll,mallinfoCd, infoCDVal, psourceArea
        Dim bufTxt : bufTxt=""

        ''안전인증정보
		If (Fsafetyyn="Y" and FsafetyDiv<>0) Then
			anjunInfo = anjunInfo & "&sft_cert_tgt_yn=Y"
			If (FsafetyDiv=10) Then
				anjunInfo = anjunInfo & "&kps_1_no="&Server.URLEncode(FsafetyNum)
			Elseif (FsafetyDiv=20) Then
				anjunInfo = anjunInfo & "&kps_1_no="
				anjunInfo = anjunInfo & "&kps_2_no="&Server.URLEncode(FsafetyNum)
			Elseif (FsafetyDiv=30) Then
				anjunInfo = anjunInfo & "&kps_1_no="
				anjunInfo = anjunInfo & "&kps_3_no="&Server.URLEncode(FsafetyNum)
			Elseif (FsafetyDiv=40) Then
				anjunInfo = anjunInfo & "&kps_1_no="
				anjunInfo = anjunInfo & "&kps_4_no="&Server.URLEncode(FsafetyNum)
			Elseif (FsafetyDiv=50) Then
				anjunInfo = anjunInfo & "&kps_1_no="
				anjunInfo = anjunInfo & "&kps_5_no="&Server.URLEncode(FsafetyNum)
			Else
				anjunInfo = ""
			End if
		End If

'		strSql = ""
'		strSql = strSql & " SELECT top 100 M.* " & vbcrlf
'		strSql = strSql & " ,isNULL(CASE WHEN M.infocd='00000' then 'N' " '''db_item.dbo.[fn_LotteCom_SaftyFormat](IC.safetyyn,IC.safetyDiv,IC.safetyNum) " & vbcrlf
'		strSql = strSql & " 	  WHEN M.infocd='99999' then M.infoETC"& vbcrlf
'		strSql = strSql & " 	  WHEN M.infocd='00005' then '해당없음'"& vbcrlf
'		strSql = strSql & " 	  WHEN M.infocd='00006' then '상품 상세 참고'"& vbcrlf
'		strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='N' THEN '해당없음' " & vbcrlf
'		strSql = strSql & " 	  WHEN c.infotype='P' THEN replace(c.infoDesc,'1644-6030','1644-6035') " & vbcrlf
'		'2014-07-14 16:07 김진영 하단 추가. 조아름 요청 "품질보증기준" 앞에 텍스트 문구 추가
'		strSql = strSql & " 	  WHEN (c.infoItemName= '품질보증기준') and (isnull(ET.itemid, '') <> '') THEN '품질보증기준 관련법 및 소비자 분쟁해결 기준에 따름, ' + F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
'		strSql = strSql & "  ELSE F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
'		strSql = strSql & "  END,'') as infoCDVal " & vbcrlf
'		strSql = strSql & "  , L.shortVal, ET.itemid as ETText " & vbcrlf
'		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
'		strSql = strSql & " Join db_item.dbo.tbl_item_contents IC " & vbcrlf
'		strSql = strSql & " on IC.infoDiv=M.mallinfoDiv " & vbcrlf
'		strSql = strSql & " left Join db_item.dbo.tbl_item_infoCode c " & vbcrlf
'		strSql = strSql & " on M.infocd=c.infocd " & vbcrlf
'		strSql = strSql & " Left Join db_item.dbo.tbl_item_infoCont F " & vbcrlf
'		strSql = strSql & " on M.infocd=F.infocd and F.itemid=" & FItemid & vbcrlf
'		strSql = strSql & " left join db_item.dbo.tbl_item_infoCont F2 " & vbcrlf
'		strSql = strSql & " on M.infocdAdd=F2.infocd and F2.itemid=" & FItemid & vbcrlf
'		strSql = strSql & " left join db_item.dbo.tbl_OutMall_etcLink as L " & vbcrlf
'		strSql = strSql & " on ((L.mallid = M.mallid) OR (isnull(L.mallid,'') = '')) and L.itemid =" & FItemid & vbcrlf
'		strSql = strSql & " and L.linkgbn='infoDiv21Lotte'" & vbcrlf
'		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_Outmall_etcTextItem as ET on IC.itemid = ET.itemid and ET.mallid = '"&CMALLNAME&"' " & vbcrlf
'		strSql = strSql & " where M.mallid = '"&CMALLNAME&"' AND IC.itemid=" & FItemid
'		strSql = strSql & " order by M.mallinfoCd"

		strSql = ""
		strSql = strSql & " SELECT top 100 M.* " & vbcrlf
		strSql = strSql & " ,isNULL(CASE WHEN M.infocd='99999' then M.infoETC"& vbcrlf
		strSql = strSql & " 	  WHEN M.infocd='00005' then '해당없음'"& vbcrlf
		strSql = strSql & " 	  WHEN M.infocd='00006' then '상품 상세 참고'"& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0106' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0507' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0701' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0703' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0801' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0803' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0901' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0903' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1001' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1003' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1007' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1101' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1107' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1201' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1206' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1301' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1306' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1401' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1407' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1409' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1601' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1602' and F.chkDiv='Y' THEN 'Y'+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1602' and F.chkDiv='N' THEN 'N'+'//번호없음'" & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1604' and F.chkDiv='Y' THEN F.infocontent+'//'+F.infocontent " & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1604' and F.chkDiv='N' THEN '해당없음//해당없음' " & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1608' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1701' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1801' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1803' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1805' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1901' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2001' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2002' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2004' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2008' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2102' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2103' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2104' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2202' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2203' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2204' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2208' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2209' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2301' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2303' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2306' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2310' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2401' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2502' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2607' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='3501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='3504' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='N' THEN '"&trim(html2db(Fmakername))&"//해당없음'" & vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='Y' THEN '"&trim(html2db(Fmakername))& "//" &trim(html2db(Fmakername))&"'" & vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='P' THEN '텐바이텐 고객행복센터//1644-6035' " & vbcrlf
		strSql = strSql & " 	  WHEN (c.infoItemName= '품질보증기준') and (isnull(ET.itemid, '') <> '') THEN '품질보증기준 관련법 및 소비자 분쟁해결 기준에 따름, ' + F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
		strSql = strSql & "  ELSE F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
		strSql = strSql & "  END,'') as infoCDVal " & vbcrlf
		strSql = strSql & "  , L.shortVal, ET.itemid as ETText " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents IC on IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c on M.infocd=c.infocd " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F on M.infocd=F.infocd and F.itemid=" & FItemid & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd=F2.infocd and F2.itemid=" & FItemid & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_OutMall_etcLink as L on ((L.mallid = M.mallid) OR (isnull(L.mallid,'') = '')) and L.itemid =" & FItemid & vbcrlf
		strSql = strSql & " and L.linkgbn='infoDiv21Lotte'" & vbcrlf
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_Outmall_etcTextItem as ET on IC.itemid = ET.itemid and ET.mallid = '"&CMALLNAME&"' " & vbcrlf
		strSql = strSql & " WHERE M.mallid = '"&CMALLNAME&"' AND IC.itemid=" & FItemid
		strSql = strSql & " and M.infoCd not in ('00000', '99999') "
		strSql = strSql & " ORDER BY M.mallinfoCd"
'response.write strSql
'response.end
		rsget.Open strSql,dbget,1
		Dim mat_name, mat_percent, mat_place, material

		psourceArea = ""
		If Not(rsget.EOF or rsget.BOF) then
			mallinfoDiv = "&ec_goods_artc_cd="&Server.URLEncode(rsget("mallinfoDiv"))
			Do until rsget.EOF
			    mallinfoCd = rsget("mallinfoCd")
			    infoCDVal  = rsget("infoCDVal")

			    IF mallinfoCd="2105" Then
			    	If isNull(rsget("shortVal")) = FALSE Then
						material = Split(rsget("shortVal"),"!!^^")
						mat_name	= material(0)
						mat_percent	= material(1)
						mat_place	= material(2)

				        bufTxt = "&mmtr_nm="&Server.URLEncode(""&mat_name&"")
				        bufTxt = bufTxt&"&cmps_rt="&Server.URLEncode(""&mat_percent&"")
				        bufTxt = bufTxt&"&mmtr_orpl_nm="&Server.URLEncode(""&mat_place&"")
			    	End If
			    Else
			        if (mallinfoCd="3503") then
			            '' 품목(35 기타인경우) 원산지값 저장
			            psourceArea = rsget("infoCDVal")
			        end if

			        if (mallinfoCd="3504") then  ''국산인경우 해당없음으로 넣어달라고 함. 수입인경우만 표시 // 2014-07-14 14:30 3504일 때 "브랜드네임,해당없음" 유미희 처리요청
			            if (isMadeInKorea(psourceArea)) then
			                infoCDVal = Fmakername & "//해당없음"
    			        end if
			        end if
        		    mallinfoCdAll = mallinfoCdAll & "&"&mallinfoCd&"=" &Server.URLEncode(infoCDVal)
    			End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		strRst = anjunInfo & mallinfoDiv & mallinfoCdAll & bufTxt
		getLotteItemInfoCdToReg = strRst
	End Function

	'2014/12/15 김진영 생성 상품품목정보수정 파라메타
	Public Function getLotteItemInfoCdToEdt()
		Dim strRst, strSQL, strRst2
		Dim anjunInfo, mallinfoCdAll,mallinfoCd, infoCDVal, psourceArea
        Dim bufTxt : bufTxt=""

		strSql = ""
		strSql = strSql & " SELECT top 100 M.* " & vbcrlf
		strSql = strSql & " ,isNULL(CASE WHEN M.infocd='99999' then M.infoETC"& vbcrlf
		strSql = strSql & " 	  WHEN M.infocd='00005' then '해당없음'"& vbcrlf
		strSql = strSql & " 	  WHEN M.infocd='00006' then '상품 상세 참고'"& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0106' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0507' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0701' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0703' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0801' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0803' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0901' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='0903' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1001' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1003' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1007' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1101' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1107' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1201' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1206' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1301' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1306' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1401' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1407' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1409' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1601' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1602' and F.chkDiv='Y' THEN 'Y'+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1602' and F.chkDiv='N' THEN 'N'+'//번호없음'" & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1604' and F.chkDiv='Y' THEN F.infocontent+'//'+F.infocontent " & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1604' and F.chkDiv='N' THEN '해당없음//해당없음' " & vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1608' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1701' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1801' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1803' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1805' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='1901' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2001' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2002' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2004' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2008' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2102' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2103' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2104' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2202' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2203' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2204' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2208' THEN F.infocontent+'//'+F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2209' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2301' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2303' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2306' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2310' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2401' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2502' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='2607' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='3501' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN M.mallinfocd='3504' THEN F.infocontent+'//'+F.infocontent "& vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='N' THEN '"&trim(Fmakername)&"//해당없음'" & vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='C' and F.chkDiv='Y' THEN '"&trim(Fmakername)& "//" &trim(Fmakername)&"'" & vbcrlf
		strSql = strSql & " 	  WHEN c.infotype='P' THEN '텐바이텐 고객행복센터//1644-6035' " & vbcrlf
		strSql = strSql & " 	  WHEN (c.infoItemName= '품질보증기준') and (isnull(ET.itemid, '') <> '') THEN '품질보증기준 관련법 및 소비자 분쟁해결 기준에 따름, ' + F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
		strSql = strSql & "  ELSE F.infocontent + isNULL(F2.infocontent,'') " & vbcrlf
		strSql = strSql & "  END,'') as infoCDVal " & vbcrlf
		strSql = strSql & "  , L.shortVal, ET.itemid as ETText " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents IC on IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c on M.infocd=c.infocd " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F on M.infocd=F.infocd and F.itemid=" & FItemid & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd=F2.infocd and F2.itemid=" & FItemid & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_OutMall_etcLink as L on ((L.mallid = M.mallid) OR (isnull(L.mallid,'') = '')) and L.itemid =" & FItemid & vbcrlf
		strSql = strSql & " and L.linkgbn='infoDiv21Lotte'" & vbcrlf
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_Outmall_etcTextItem as ET on IC.itemid = ET.itemid and ET.mallid = '"&CMALLNAME&"' " & vbcrlf
		strSql = strSql & " WHERE M.mallid = '"&CMALLNAME&"' AND IC.itemid=" & FItemid
		strSql = strSql & " and M.infoCd not in ('00000', '99999') "
		strSql = strSql & " ORDER BY M.mallinfoCd"
'		rw strSql
'		response.end
		rsget.Open strSql,dbget,1
		Dim mat_name, mat_percent, mat_place, material
		psourceArea = ""
		If Not(rsget.EOF or rsget.BOF) then
			strRst2 = "&ec_goods_artc_cd="&Server.URLEncode(rsget("mallinfoDiv"))
			Do until rsget.EOF
			    mallinfoCd = rsget("mallinfoCd")
			    infoCDVal  = rsget("infoCDVal")

			    IF mallinfoCd="2105" Then
			    	If isNull(rsget("shortVal")) = FALSE Then
						material = Split(rsget("shortVal"),"!!^^")
						mat_name	= material(0)
						mat_percent	= material(1)
						mat_place	= material(2)

				        bufTxt = "&mmtr_nm="&Server.URLEncode(""&mat_name&"")
				        bufTxt = bufTxt&"&cmps_rt="&Server.URLEncode(""&mat_percent&"")
				        bufTxt = bufTxt&"&mmtr_orpl_nm="&Server.URLEncode(""&mat_place&"")
			    	End If
			    Else
			        if (mallinfoCd="3503") then
			            '' 품목(35 기타인경우) 원산지값 저장
			            psourceArea = rsget("infoCDVal")
			        end if

			        if (mallinfoCd="3504") then  ''국산인경우 해당없음으로 넣어달라고 함. 수입인경우만 표시 // 2014-07-14 14:30 3504일 때 "브랜드네임,해당없음" 유미희 처리요청
			            if (isMadeInKorea(psourceArea)) then
			                infoCDVal = Fmakername & "//해당없음"
    			        end if
			        end if
        		    mallinfoCdAll = mallinfoCdAll & "&"&mallinfoCd&"=" &Server.URLEncode(infoCDVal)
    			End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		strRst = "subscriptionId=" & lotteAuthNo
		strRst = strRst & "&goods_no=" & FLotteGoodNo
		strRst2 = strRst2 & "&chg_caus_cont=" & Server.URLEncode("api 상품품목 변경")
		getLotteItemInfoCdToEdt = strRst & bufTxt & strRst2 & mallinfoCdAll
	End Function

	'// 텐바이텐 상품옵션 검사
	public function checkTenItemOptionValid()
		dim strSql, chkRst, chkMultiOpt
		dim cntType, cntOpt
		chkRst = true
		chkMultiOpt = false

		if FoptionCnt>0 then
			'// 이중옵션확인
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget

			if Not(rsget.EOF or rsget.BOF) then
				chkMultiOpt = true
				cntType = rsget.RecordCount
			end if
			rsget.Close

			if chkMultiOpt then
				'// 이중옵션 일때
				strSql = "Select optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				if Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						cntOpt = ubound(split(db2Html(rsget("optionname")),","))+1
						if cntType<>cntOpt then
							chkRst = false
						end if
						rsget.MoveNext
					Loop
				else
					chkRst = false
				end if
				rsget.Close
			Else
				'// 단일옵션일 때
				strSql = "Select optionTypeName, optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				''strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.Open strSql,dbget,1

				if (rsget.EOF or rsget.BOF) then
					chkRst = false
				end if
				rsget.Close
			end if
		end if

		'//결과 반환
		checkTenItemOptionValid = chkRst

	end Function

	'// 롯데닷컴 판매여부 반환
	public function getLotteSellYn()
		'판매상태 (10:판매진행, 20:품절)
		if FsellYn="Y" and FisUsing="Y" then
			if FLimitYn="N" or (FLimitYn="Y" and FLimitNo-FLimitSold>=CMAXLIMITSELL) then
				getLotteSellYn = "Y"
			else
				getLotteSellYn = "N"
			end if
		else
			getLotteSellYn = "N"
		end if
	end Function

	'// 롯데닷컴 등록상태 반환
	public function getLotteItemStatCd()
		Select Case FLotteStatCd
		    Case "00"
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
		End Select
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end class

class CLotte
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
	public FRectGrpCode

	public FRectCDL
	public FRectCDM
	public FRectCDS

	public FRectItemID
	public FRectItemName
	public FRectMakerid
	public FRectLotteNotReg
	public FRectMatchCate
	''public FRectMatchCateNotCheck
	public FRectSellYn
	public FRectLimitYn
	public FRectSailYn
	public FRectLotteGoodNo
	public FRectLotteTmpGoodNo
	public FRectMinusMigin
	public FRectonlyValidMargin
	public FRectIsSoldOut
	public FRectExpensive10x10
	public FRectLotteYes10x10No
	public FRectLotteNo10x10Yes
	public FRectOnreginotmapping
	public FRectNotJehyu
	public FRectEventid
	public FRectdiffPrc
	public FRectdisptpcd
    public FRectoptAddprcExists
    public FRectoptAddPrcRegTypeNone
    public FRectoptAddprcExistsExcept
    public FRectoptExists
    public FRectoptnotExists
    public FRectisMadeHand
    public FRectregedOptNull
    public FRectOrdType
    public FRectFailCntExists
    public FRectFailCntOverExcept
    public FRectLimitOver
    public FRectExtSellYn
    public FRectInfoDivYn
    public FRectOnlyNotUsingCheck

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
    public Sub getLotte_MDList
		dim sqlStr,i
		sqlStr = " select count(MDCode) as cnt, CEILING(CAST(Count(MDCode) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + "From db_temp.dbo.tbl_lotte_MDInfo "

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotte_MDInfo "
		sqlStr = sqlStr + " order by MDCode asc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem
				FItemList(i).FMDCode		= rsget("MDCode")
				FItemList(i).FMDName		= db2html(rsget("MDName"))
				FItemList(i).FSellFeeType	= rsget("SellFeeType")
				FItemList(i).FNormalSellFee	= rsget("NormalSellFee")
				FItemList(i).FEventSellFee	= rsget("EventSellFee")
				FItemList(i).FisUsing		= rsget("isUsing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


    '// 담당MD상품군 목록
    public Sub getLotte_MDGrpList
		dim sqlStr, i

		sqlStr = " select count(groupCode) as cnt, CEILING(CAST(Count(groupCode) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " From db_temp.dbo.tbl_lotte_MDCateGrp "
		sqlStr = sqlStr + " Where MDCode='" & FRectMdCode & "'"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotte_MDCateGrp "
		sqlStr = sqlStr + " Where MDCode='" & FRectMdCode & "'"
		sqlStr = sqlStr + " order by groupCode asc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem

				FItemList(i).FgroupCode			= rsget("groupCode")
				FItemList(i).FSuperGroupName	= db2html(rsget("SuperGroupName"))
				FItemList(i).FGroupName			= db2html(rsget("GroupName"))
				FItemList(i).FisUsing			= rsget("isUsing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'--------------------------------------------------------------------------------
	'// 텐바이텐-롯데닷컴 카테고리 :: 전문몰 카테고리가 매핑 되어야 함..
	public Sub getTenLotteCateList
		dim sqlStr, addSql, i

		if FRectCDL<>"" then
			addSql = addSql & " and s.code_large='" & FRectCDL & "'"
		end if
		if FRectCDM<>"" then
			addSql = addSql & " and s.code_mid='" & FRectCDM & "'"
		end if
		if FRectCDS<>"" then
			addSql = addSql & " and s.code_small='" & FRectCDS & "'"
		end if
		if FRectDspNo<>"" then
			addSql = addSql & " and cm.dispNo='" & FRectDspNo & "'"
		end if

		if FRectIsMapping="Y" then
			addSql = addSql & " and cm.DispNo is Not null "
		elseif FRectIsMapping="N" then
			addSql = addSql & " and cm.DispNo is null "
		end if

		if FRectKeyword<>"" then
			Select Case FRectSDiv
				Case "LCD"	'롯데닷컴 전시코드 검색
					addSql = addSql & " and cm.DispNo='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(텐바이텐 소분류명)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		end if

		sqlStr = " select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_item.dbo.tbl_cate_small as s "
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_lotte_cate_mapping as cm "
		sqlStr = sqlStr + " 		on cm.tenCateLarge=s.code_large "
		sqlStr = sqlStr + " 			and cm.tenCateMid=s.code_mid "
		sqlStr = sqlStr + " 			and cm.tenCateSmall=s.code_small "
		If FRectdisptpcd <> "" Then
			sqlStr = sqlStr & " 	Join db_temp.dbo.tbl_lotte_Category as lc on lc.DispNo = cm.DispNo and lc.disptpcd='" & FRectdisptpcd &"'"
		Else
			sqlStr = sqlStr & " 	left Join db_temp.dbo.tbl_lotte_Category as lc on lc.DispNo = cm.DispNo "
		End If
'		sqlStr = sqlStr + " 	left Join db_temp.dbo.tbl_lotte_Category as lc "
'		sqlStr = sqlStr + " 		on lc.DispNo=cm.DispNo "
'		if FRectdisptpcd<>"" then
'            sqlStr = sqlStr & " and lc.disptpcd='" & FRectdisptpcd &"'"
'        end if


		sqlStr = sqlStr + " Where 1=1 " & addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " 	s.code_large,s.code_mid,s.code_small "
		sqlStr = sqlStr + " 	,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm "
		sqlStr = sqlStr + " 	,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "
		sqlStr = sqlStr + " 	,code_nm as small_nm "
		sqlStr = sqlStr + " 	,cm.DispNo, lc.DispNm, lc.DispLrgNm, lc.DispMidNm, lc.DispSmlNm, lc.DispThnNm, lc.groupCode, lc.disptpcd "
		sqlStr = sqlStr + " 	,lc.isusing"
		sqlStr = sqlStr + " from db_item.dbo.tbl_cate_small as s "
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_lotte_cate_mapping as cm "
		sqlStr = sqlStr + " 		on cm.tenCateLarge=s.code_large "
		sqlStr = sqlStr + " 			and cm.tenCateMid=s.code_mid "
		sqlStr = sqlStr + " 			and cm.tenCateSmall=s.code_small "

		If FRectdisptpcd <> "" Then
			sqlStr = sqlStr & " 	Join db_temp.dbo.tbl_lotte_Category as lc on lc.DispNo = cm.DispNo and lc.disptpcd='" & FRectdisptpcd &"'"
		Else
			sqlStr = sqlStr & " 	left Join db_temp.dbo.tbl_lotte_Category as lc on lc.DispNo = cm.DispNo "
		End If

'		sqlStr = sqlStr + " 	left Join db_temp.dbo.tbl_lotte_Category as lc "
'		sqlStr = sqlStr + " 		on lc.DispNo=cm.DispNo "
'		if FRectdisptpcd<>"" then
'            sqlStr = sqlStr & " and lc.disptpcd='" & FRectdisptpcd &"'"
'        end if

		sqlStr = sqlStr + " Where 1=1 and s.display_yn = 'Y' " & addSql
		sqlStr = sqlStr & " ORDER BY s.code_large, s.code_mid, s.code_small, disptpcd desc "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem

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
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'// 롯데닷컴 카테고리
	public Sub getLotteCategoryList
		dim sqlStr, addSql, i

		if FRectDspNo<>"" then
			addSql = addSql & " and c.dispNo=" & FRectDspNo
		end if

'		if FRectIsMapping="Y" then
'			addSql = addSql & " and m.DispNo is Not null "
'		elseif FRectIsMapping="N" then
'			addSql = addSql & " and m.DispNo is null "
'		end if

		if FRectGrpCode<>"" then
			addSql = addSql & " and c.groupCode=" & FRectGrpCode
		end if

        if FRectdisptpcd<>"" then
            addSql = addSql & " and c.disptpcd='" & FRectdisptpcd &"'"
        end if

		if FRectKeyword<>"" then
			Select Case FRectSDiv
				Case "LCD"	'롯데닷컴 전시코드 검색
					addSql = addSql & " and c.DispNo='" & FRectKeyword & "'"
'				Case "TCD"	'텐바이텐 카테고리코드 검색(대중소 통합코드 9자리)
'					addSql = addSql & " and m.tenCateLarge&m.tenCateMid&m.tenCateSmall='" & FRectKeyword & "'"
				Case "CNM"	'카테고리명(롯데 세분류명)
					''addSql = addSql & " and c.dispNm like '%" & FRectKeyword & "%'"
					addSql = addSql & " and ((c.dispNm like '%" & FRectKeyword & "%') or (c.dispsmlNm like '%" & FRectKeyword & "%') or (c.dispLrgNm like '%" & FRectKeyword & "%') )"
			End Select
		end if

		sqlStr = " select count(c.DispNo) as cnt, CEILING(CAST(Count(c.DispNo) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotte_Category as c "
'		sqlStr = sqlStr + " 	Left Join db_item.dbo.tbl_lotte_cate_mapping as m "
'		sqlStr = sqlStr + " 		on c.DispNo=m.DispNo "
		sqlStr = sqlStr + " Where c.DispMidNm<>'바보사랑' and c.groupcode in (1590,1596,1597,1598,1599,13795) " & addSql



		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

'		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " c.*, m.tenCateLarge, m.tenCateMid, m.tenCateSmall, s.code_nm "
		sqlStr = " select distinct top " + CStr(FPageSize*FCurrPage) + " c.* "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_lotte_Category as c "
'		sqlStr = sqlStr + " 	Left Join db_item.dbo.tbl_lotte_cate_mapping as m "
'		sqlStr = sqlStr + " 		on c.DispNo=m.DispNo "
'		sqlStr = sqlStr + " 	Left Join db_item.dbo.tbl_cate_small as s "
'		sqlStr = sqlStr + " 		on s.code_large=m.tenCateLarge "
'		sqlStr = sqlStr + " 		and s.code_mid=m.tenCateMid "
'		sqlStr = sqlStr + " 		and s.code_small=m.tenCateSmall "
		sqlStr = sqlStr + " Where c.DispMidNm<>'바보사랑' and c.groupcode in (1590,1596,1597,1598,1599,13795) " & addSql
		sqlStr = sqlStr + " order by c.DispLrgNm, c.DispMidNm, c.DispSmlNm, c.DispThnNm, c.DispNo"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem

				FItemList(i).FDispNo		= rsget("DispNo")
				FItemList(i).FDispNm		= db2html(rsget("DispNm"))
				FItemList(i).FDispLrgNm		= db2html(rsget("DispLrgNm"))
				FItemList(i).FDispMidNm		= db2html(rsget("DispMidNm"))
				FItemList(i).FDispSmlNm		= db2html(rsget("DispSmlNm"))
				FItemList(i).FDispThnNm		= db2html(rsget("DispThnNm"))
                FItemList(i).Fdisptpcd      = rsget("disptpcd")

'				FItemList(i).FtenCateLarge	= rsget("tenCateLarge")
'				FItemList(i).FtenCateMid	= rsget("tenCateMid")
'				FItemList(i).FtenCateSmall	= rsget("tenCateSmall")
'				FItemList(i).FtenCateName	= db2html(rsget("code_nm"))

				FItemList(i).FisUsing		= rsget("isUsing")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


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
		if Cint(FCurrPage)>Cint(FTotalPage) then
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
				set FItemList(i) = new CLotteItem

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

	'// 롯데닷컴 상품 목록 // 수정시 조건이 달라야 함..
	public Sub getLotteRegedItemList
		dim sqlStr, addSql, i

		if FRectMakerid<>"" then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		end if

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		if FRectItemName<>"" then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		end if

		if FRectLotteGoodNo<>"" then
			addSql = addSql & " and m.LotteGoodNo='" & FRectLotteGoodNo & "'"
		end if

		if FRectCDL<>"" then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		end if
		if FRectCDM<>"" then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		end if
		if FRectCDS<>"" then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		end if

		'롯데닷컴 임시상품번호 검색
		If FRectLotteTmpGoodNo <> "" Then
			addSql = addSql & " and m.LotteTmpGoodNo = '" & FRectLotteTmpGoodNo & "'"
		End If

		Select Case FRectLotteNotReg
			Case "M"	'미등록
				addSql = addSql & " and m.itemid is Null and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>=5)) "
			Case "W"	'등록예정
			    addSql = addSql & " and m.lotteStatCd='00'"                 ''등록예정
				addSql = addSql & " and m.LotteTmpGoodNo is Null"
				addSql = addSql & " and m.LotteGoodNo is Null"
			Case "F"	'등록완료(임시)
				addSql = addSql & " and m.LotteTmpGoodNo is Not Null"
				addSql = addSql & " and m.LotteGoodNo is Null"
			Case "D"	'등록완료(전시)
				addSql = addSql & " and m.LotteTmpGoodNo is Not Null"
				addSql = addSql & " and m.LotteGoodNo is Not Null"
			Case "R"	'수정요망
		        addSql = addSql & " and m.LotteGoodNo is Not NULL"
		        addSql = addSql & " and m.LotteLastUpdate < i.lastupdate"
		    Case "J"	'반려
		        addSql = addSql & " and m.lotteStatCd='40'"                 ''반려
		        addSql = addSql & " and m.LotteTmpGoodNo is Not NULL"
		        addSql = addSql & " and m.LotteGoodNo is  NULL"
		End Select

		Select Case FRectMatchCate
			Case "Y"	'매칭완료
				''addSql = addSql & " and c.mapCnt is Not Null"
				addSql = addSql & " and isNULL(c.mapCnt,0)>0"
			Case "N"	'미매칭
				addSql = addSql & " and c.mapCnt is Null"

		End Select

		Select Case FRectSellYn
			Case "Y"	'판매
				addSql = addSql & " and i.sellYn='Y'"
			Case "N"	'품절
				addSql = addSql & " and i.sellYn in ('S','N')"
		End Select

		if FRectLimitYn<>"" then
			addSql = addSql & " and i.limitYn='" & FRectLimitYn & "'"
		end if

		Select Case FRectSailYn
			Case "Y"	'세일Y
				addSql = addSql & " and i.sailYn='Y'"
			Case "N"	'품절
				addSql = addSql & " and i.sailYn='N'"
		End Select

		if (FRectMinusMigin<>"") then
		   addSql = addSql & " and i.sellcash<>0"
		   addSql = addSql & " and ((i.sellcash-i.buycash)/i.sellcash)*100<"&CMAXMARGIN & VbCrlf
		   addSql = addSql & " and m.LotteSellYn= 'Y' " '''  조건 추가.
		else
		   IF (FRectonlyValidMargin<>"") then
		        addSql = addSql & " and i.sellcash<>0"
		        addSql = addSql & " and ((i.sellcash-i.buycash)/i.sellcash)*100>="&CMAXMARGIN & VbCrlf
		   END IF
		   ''addSql = addSql & " and m.LotteSellYn<> 'X' " '''  조건 추가.
		end if

		if FRectExpensive10x10 <> "" then
		   addSql = addSql & " and m.LottePrice is Not Null and i.sellcash > m.LottePrice "
		end if

        if FRectdiffPrc <> "" then
		   addSql = addSql & " and m.LottePrice is Not Null and i.sellcash <> m.LottePrice "
		end if

		if FRectLotteYes10x10No <> "" then
		   ''addSql = addSql & " and m.LottePrice is Not Null and i.sellcash > m.LottePrice "
		   addSql = addSql & " and m.LottePrice is Not Null and (m.LotteSellYn= 'Y' and i.sellyn <> 'Y')"
		end if

		if FRectLotteNo10x10Yes <> "" then
		   addSql = addSql & " and m.LottePrice is Not Null and (m.LotteSellYn= 'N' and i.sellyn='Y' and (i.limityn='N' or (i.limityn='Y' and i.limitno-i.limitsold>="&CMAXLIMITSELL&")))"
		end if


		if FRectOnreginotmapping <> "" then
		    addSql = addSql & " and m.LotteTmpGoodNo is Not Null and IsNULL(c.mapCnt,0)>0" '''c.mapCnt is Null
		end if

		if FRectEventid<>"" then
			addSql = addSql & " and i.itemid in (Select itemid From [db_event].[dbo].tbl_eventitem Where evt_code='" & FRectEventid & "')" + VbCrlf
		end if

		iF (FRectOrdType="LU") then		'2014-01-13 김진영 추가..상품최근업데이트가 없어서..
		    addSql = addSql & " and isnull(m.itemTableUpdateChkdate,'') = '' "
		    addSql = addSql & " and Left(i.lastupdate, 10) <> Left(m.LotteLastUpdate,10) "
		end if

		''옵션추가금액 존재상품.
		if (FRectoptAddprcExists<>"") and (FRectLotteNotReg<>"M") then
		    addSql = addSql & " and m.optAddPrcCnt>0"
'		    addSql = addSql & " and i.itemid in ("
'		    addSql = addSql & "     select distinct ii.itemid "
'		    addSql = addSql & "     from db_item.dbo.tbl_item ii "
'		    addSql = addSql & "     Join db_item.dbo.tbl_item_option o "
'		    addSql = addSql & "     on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
'		    addSql = addSql & " )"
		end if

		if (FRectoptAddPrcRegTypeNone<>"") then          ''옵션추가금액상품 미설정 상품.
		    addSql = addSql & " and m.optAddPrcCnt>0"
		    addSql = addSql & " and m.optAddPrcRegType=0"
		end if

		''옵션추가금액 존재상품 제외
		if (FRectoptAddprcExistsExcept<>"") then
		    addSql = addSql & " and i.itemid Not in ("
		    addSql = addSql & "     select distinct ii.itemid "
		    addSql = addSql & "     from db_item.dbo.tbl_item ii "
		    addSql = addSql & "     Join db_item.dbo.tbl_item_option o "
		    addSql = addSql & "     on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
		    addSql = addSql & " )"
		end if

		if (FRectoptExists<>"") then
            addSql = addSql & " and i.optioncnt>0"
        end if

		if (FRectoptnotExists<>"") then
            addSql = addSql & " and i.optioncnt=0"
        end if

		If FRectisMadeHand<>"" then
			if (FRectisMadeHand="Y") then
				addSql = addSql & " and i.itemdiv in ('06', '16')" & VbCrlf
			Else
				addSql = addSql & " and i.itemdiv not in ('06', '16')" & VbCrlf
			End If
		End if

        if (FRectregedOptNull<>"") then
            addSql = addSql & " and isNULL(m.regedOptCnt,0)=0"
            addSql = addSql & " and i.optioncnt>0"
        end if

        if (FRectFailCntExists<>"") then
            addSql = addSql & " and m.accFailCNT>0"
        end if

        if (FRectFailCntOverExcept<>"") then
            addSql = addSql & " and m.accFailCNT<"&FRectFailCntOverExcept
        end if

        if (FRectLimitOver<>"") then
            addSql = addSql & " and ((i.limityn='N') or ((i.limityn='Y') and (i.limitno-i.limitsold>"&FRectLimitOver&")))"
        end if

        if (FRectLotteYes10x10No = "" ) then
    		'//제휴몰 판매만 허용
    		addSql = addSql & " and i.isExtUsing='Y'"
    		'//착불배송 상품 제거
    		addSql = addSql & " and i.deliverytype not in ('7')"
    		'//조건배송 10000원 이상
'            addSql = addSql + " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
        end if

        if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        addSql = addSql + " and m.lotteSellYn<>'X'"
		    else
		        addSql = addSql + " and m.lotteSellYn='" & FRectExtSellYn & "'"
		    end if
		end if

		if (FRectInfoDivYn<>"") then
		    if FRectInfoDivYn="Y" then
		        addSql = addSql + " and isNULL(ct.infoDiv,'')<>''"
		    elseif FRectInfoDivYn="N" then
    		    addSql = addSql + " and isNULL(ct.infoDiv,'')=''"
    		else
    		    addSql = addSql + " and ct.infoDiv='"&FRectInfoDivYn&"'"
    		end if
		end if

		sqlStr = " select count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_item.dbo.tbl_item as i "
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_lotte_regItem as m "
		sqlStr = sqlStr + " 		on i.itemid=m.itemid "
		sqlStr = sqlStr + " 	left Join db_item.dbo.tbl_OutMall_CateMap_Summary as c " ''where disptpcd='12' 조건추가 브랜드 전문몰
		sqlStr = sqlStr + " 		on c.mallid='"&CMALLNAME&"' and c.tenCateLarge=i.cate_large and c.tenCateMid=i.cate_mid and c.tenCateSmall=i.cate_small "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item_contents as ct "
		sqlStr = sqlStr + " 	on i.itemid=ct.itemid"
		sqlStr = sqlStr + " 	left join db_user.dbo.tbl_user_c uc"
		sqlStr = sqlStr + " 	on i.makerid=uc.userid"
		sqlStr = sqlStr + " 	left join db_partner.dbo.tbl_partner p"
		sqlStr = sqlStr + " 	on i.makerid=p.id"
		sqlStr = sqlStr + " where 1=1 "

		''if (FRectLotteNotReg="D") or (FRectLotteNotReg="R") then  ''이미 상품이 등록된 CASE
		IF (FRectLotteNotReg<>"M") and (FRectLotteNotReg<>"W") THEN ''미등록/등록가능:M, 등록예정:W

		ELSE
			IF (FRectLotteNotReg="M") Then	'2014-01-24 이문재이사님요청//간이과세는 미등록 리스트에 뿌리지 않게..
				sqlStr = sqlStr + " 	and p.jungsan_gubun not in ('간이과세', '면세') "			
			End If
		    ''if (FRectLotteYes10x10No = "" ) then
        		sqlStr = sqlStr + "     and i.isusing='Y' "
        		sqlStr = sqlStr + "     and i.deliverfixday not in ('C','X') "
        		sqlStr = sqlStr + "     and i.basicimage is not null "
        		sqlStr = sqlStr + "     and i.itemdiv<50 and i.itemdiv<>'08' "
        		sqlStr = sqlStr + "     and i.cate_large<>'' "
        		sqlStr = sqlStr + "     and i.cate_large<>'999' "
        		sqlStr = sqlStr + "		and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
        		sqlStr = sqlStr + "		and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
        		sqlStr = sqlStr + "		and uc.isExtusing='Y'" '' 브랜드 제휴사용여부Y만
        	''ENd IF
    	ENd IF


    	sqlStr = sqlStr + "     and i.sellcash>0 " & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " i.* "
		sqlStr = sqlStr + "		, m.LotteRegdate, m.LotteLastUpdate, m.LotteGoodNo, m.LotteTmpGoodNo, m.LottePrice, m.LotteSellYn, m.regUserid, m.lotteStatCd "
		sqlStr = sqlStr + "		, c.mapCnt , m.rctSellCNT, m.regedOptCnt, m.accFailCNT, m.lastErrStr "
		sqlStr = sqlStr + "		, Ct.infoDiv"
		sqlStr = sqlStr + "     ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr + "     ,m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr + " from db_item.dbo.tbl_item as i "
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_lotte_regItem as m "
		sqlStr = sqlStr + " 		on i.itemid=m.itemid "
		sqlStr = sqlStr + " 	left Join db_item.dbo.tbl_OutMall_CateMap_Summary as c " ''where disptpcd='12' 조건추가 브랜드 전문몰
		sqlStr = sqlStr + " 		on c.mallid='"&CMALLNAME&"' and c.tenCateLarge=i.cate_large and c.tenCateMid=i.cate_mid and c.tenCateSmall=i.cate_small "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item_contents as ct "
		sqlStr = sqlStr + " 	on i.itemid=ct.itemid"
		sqlStr = sqlStr + " 	left join db_user.dbo.tbl_user_c uc"
		sqlStr = sqlStr + " 	on i.makerid=uc.userid"
		sqlStr = sqlStr + " 	left join db_partner.dbo.tbl_partner p"
		sqlStr = sqlStr + " 	on i.makerid=p.id"
		sqlStr = sqlStr + " where 1=1 "
		''if (FRectLotteNotReg="D") or (FRectLotteNotReg="R") then  ''이미 상품이 등록된 CASE
		IF (FRectLotteNotReg<>"M") and (FRectLotteNotReg<>"W") THEN ''미등록/등록가능:M, 등록예정:W

		ELSE
			IF (FRectLotteNotReg="M") Then	'2014-01-24 이문재이사님요청//간이과세는 미등록 리스트에 뿌리지 않게..
				sqlStr = sqlStr + " 	and p.jungsan_gubun not in ('간이과세', '면세') "
			End If
		    ''if (FRectLotteYes10x10No = "" ) then
        		sqlStr = sqlStr + "     and i.isusing='Y' "
        		sqlStr = sqlStr + "     and i.deliverfixday not in ('C','X') "
        		sqlStr = sqlStr + "     and i.basicimage is not null "
        		sqlStr = sqlStr + "     and i.itemdiv<50 and i.itemdiv<>'08' "
        		sqlStr = sqlStr + "     and i.cate_large<>'' "
        		sqlStr = sqlStr + "     and i.cate_large<>'999' "
        		sqlStr = sqlStr + "		and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
        		sqlStr = sqlStr + "		and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
        		sqlStr = sqlStr + "		and uc.isExtusing='Y'" '' 브랜드 제휴사용여부Y만
        	''end if
        end if

        sqlStr = sqlStr + "     and i.sellcash>0 " & addSql

		'sqlStr = sqlStr + " order by i.itemid desc "


		IF (FRectOrdType="B") then
		    sqlStr = sqlStr + " order by i.itemscore desc, i.itemid desc"
		ELSEIF (FRectOrdType="BM") then
		    sqlStr = sqlStr + " order by m.rctSellCNT desc,i.itemscore desc, m.regdate desc"
		ELSEIF (FRectOrdType="LU") then		'2014-01-13 김진영 추가..상품최근업데이트가 없어서..
		    sqlStr = sqlStr + " order by i.lastupdate desc, i.itemscore desc, i.itemid desc "
		ELSE
		    sqlStr = sqlStr + " order by i.itemid desc "
	    ENd IF
'rw sqlStr
''response.end

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem

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

				FItemList(i).FLotteRegdate		= rsget("LotteRegdate")
				FItemList(i).FLotteLastUpdate	= rsget("LotteLastUpdate")
				FItemList(i).FLotteGoodNo		= rsget("LotteGoodNo")
				FItemList(i).FLotteTmpGoodNo	= rsget("LotteTmpGoodNo")
				FItemList(i).FLottePrice		= rsget("LottePrice")
				FItemList(i).FLotteSellYn		= rsget("LotteSellYn")
				FItemList(i).FregUserid			= rsget("regUserid")
				FItemList(i).FLotteStatCd		= rsget("lotteStatCd")
				FItemList(i).FCateMapCnt		= rsget("mapCnt")
				FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")

				if Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
				else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				end if
				FItemList(i).FoptionCnt       = rsget("optionCnt")
                FItemList(i).FrctSellCNT      = rsget("rctSellCNT")
                FItemList(i).FregedOptCnt     = rsget("regedOptCnt")
                FItemList(i).FaccFailCNT      = rsget("accFailCNT")
                FItemList(i).FlastErrStr      = rsget("lastErrStr")

                FItemList(i).FinfoDiv       = rsget("infoDiv")
                FItemList(i).FoptAddPrcCnt      = rsget("optAddPrcCnt")
                FItemList(i).FoptAddPrcRegType  = rsget("optAddPrcRegType")
                FItemList(i).FItemdiv  = rsget("itemdiv")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

    ''' 등록되지 말아야 될 상품..
    public Sub getLotteReqExpireItemList
		dim sqlStr, addSql, i


		sqlStr = " select count(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr + " from db_item.dbo.tbl_item as i "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_lotte_regItem as m "
		sqlStr = sqlStr + " 		on i.itemid=m.itemid "
		sqlStr = sqlStr + " 		and m.LotteGoodNo is Not Null"
		''sqlStr = sqlStr + " 		and m.LotteSellYn= 'Y' "                    ''' 롯데 판매중인거만.
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item_contents as ct "
		sqlStr = sqlStr + " 	on i.itemid=ct.itemid"
		sqlStr = sqlStr + " 	left join db_user.dbo.tbl_user_c uc"
		sqlStr = sqlStr + " 	on i.makerid=uc.userid"
		sqlStr = sqlStr + " where 1=1"

		if (FRectOnlyNotUsingCheck="on") then
		    sqlStr = sqlStr + "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
    		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
    		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
    		sqlStr = sqlStr + "		or uc.isExtUsing='N'"
            sqlStr = sqlStr + "	)"
            ''//연동 제외상품
            sqlStr = sqlStr & " and i.itemid not in ("
            sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
            sqlStr = sqlStr & "     where stDt<getdate()"
            sqlStr = sqlStr & "     and edDt>getdate()"
            sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
            sqlStr = sqlStr & "     and linkgbn='donotEdit'"
            sqlStr = sqlStr & " )"
		else
    		sqlStr = sqlStr + "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
    		sqlStr = sqlStr + "     or i.deliverytype in ('7') "
    		'//조건배송 10000원 이상
            sqlStr = sqlStr + "     or ((i.deliveryType=9) and (i.sellcash<10000))"
    		sqlStr = sqlStr + "     or i.deliverfixday in ('C','X') "
    		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
    		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
    		sqlStr = sqlStr + "		or uc.isExtUsing='N'"
            sqlStr = sqlStr + "	)"

            ''//연동 제외상품
            sqlStr = sqlStr & " and i.itemid not in ("
            sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
            sqlStr = sqlStr & "     where stDt<getdate()"
            sqlStr = sqlStr & "     and edDt>getdate()"
            sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
            sqlStr = sqlStr & "     and linkgbn='donotEdit'"
            sqlStr = sqlStr & " )"
        end if

        if FRectMakerid<>"" then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		end if

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        sqlStr = sqlStr + " and m.lotteSellYn<>'X'"
		    else
		        sqlStr = sqlStr + " and m.lotteSellYn='" & FRectExtSellYn & "'"
		    end if
		end if

		Select Case FRectSellYn
			Case "Y"	'판매
				sqlStr = sqlStr & " and i.sellYn='Y'"
			Case "N"	'품절
				sqlStr = sqlStr & " and i.sellYn in ('S','N')"
		End Select

		if (FRectInfoDivYn<>"") then
		    if FRectInfoDivYn="Y" then
		        sqlStr = sqlStr + " and isNULL(ct.infoDiv,'')<>''"
		    elseif FRectInfoDivYn="N" then
    		    sqlStr = sqlStr + " and isNULL(ct.infoDiv,'')=''"
    		else
    		    sqlStr = sqlStr + " and ct.infoDiv='"&FRectInfoDivYn&"'"
    		end if
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " i.* "
		sqlStr = sqlStr + "		, m.LotteRegdate, m.LotteLastUpdate, m.LotteGoodNo, m.LotteTmpGoodNo, m.LottePrice, m.LotteSellYn, m.regUserid, m.lotteStatCd "
		sqlStr = sqlStr + "		, 1 as mapCnt, m.rctSellCNT "
		sqlStr = sqlStr + "		, ct.infoDiv"
		sqlStr = sqlStr + "     ,uc.defaultdeliverytype, uc.defaultfreeBeasongLimit"
		sqlStr = sqlStr + "     , m.optAddPrcCnt, m.optAddPrcRegType"
		sqlStr = sqlStr + " from db_item.dbo.tbl_item as i "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_lotte_regItem as m "
		sqlStr = sqlStr + " 		on i.itemid=m.itemid "
		sqlStr = sqlStr + " 		and m.LotteGoodNo is Not Null"
		''sqlStr = sqlStr + " 		and m.LotteSellYn= 'Y' "                ''' 롯데 판매중인거만.
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item_contents as ct "
		sqlStr = sqlStr + " 	on i.itemid=ct.itemid"
		sqlStr = sqlStr + " 	left join db_user.dbo.tbl_user_c uc"
		sqlStr = sqlStr + " 	on i.makerid=uc.userid"
		sqlStr = sqlStr + " where 1=1"

		if (FRectOnlyNotUsingCheck="on") then
		    sqlStr = sqlStr + "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
    		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
    		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
    		sqlStr = sqlStr + "		or uc.isExtUsing='N'"
            sqlStr = sqlStr + "	)"
            ''//연동 제외상품
            sqlStr = sqlStr & " and i.itemid not in ("
            sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
            sqlStr = sqlStr & "     where stDt<getdate()"
            sqlStr = sqlStr & "     and edDt>getdate()"
            sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
            sqlStr = sqlStr & "     and linkgbn='donotEdit'"
            sqlStr = sqlStr & " )"
		else
    		sqlStr = sqlStr + "     and (i.isusing<>'Y' or i.isExtUsing<>'Y' "
    		sqlStr = sqlStr + "     or i.deliverytype in ('7') "
    		'//조건배송 10000원 이상
            sqlStr = sqlStr + "     or ((i.deliveryType=9) and (i.sellcash<10000))"
    		sqlStr = sqlStr + "     or i.deliverfixday in ('C','X') "
    		sqlStr = sqlStr + "     or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
    		sqlStr = sqlStr + "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
    		sqlStr = sqlStr + "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
    		sqlStr = sqlStr + "		or uc.isExtUsing='N'"
            sqlStr = sqlStr + "	)"

            ''//연동 제외상품 //디비로 만들어야 할듯.
            sqlStr = sqlStr & " and i.itemid not in ("
            sqlStr = sqlStr & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
            sqlStr = sqlStr & "     where stDt<getdate()"
            sqlStr = sqlStr & "     and edDt>getdate()"
            sqlStr = sqlStr & "     and mallid='"&CMALLNAME&"'"
            sqlStr = sqlStr & "     and linkgbn='donotEdit'"
            sqlStr = sqlStr & " )"
        end if

        if FRectMakerid<>"" then
			sqlStr = sqlStr & " and i.makerid='" & FRectMakerid & "'"
		end if

        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlStr = sqlStr & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        sqlStr = sqlStr + " and m.lotteSellYn<>'X'"
		    else
		        sqlStr = sqlStr + " and m.lotteSellYn='" & FRectExtSellYn & "'"
		    end if
		end if

		Select Case FRectSellYn
			Case "Y"	'판매
				sqlStr = sqlStr & " and i.sellYn='Y'"
			Case "N"	'품절
				sqlStr = sqlStr & " and i.sellYn in ('S','N')"
		End Select

		if (FRectInfoDivYn<>"") then
		    if FRectInfoDivYn="Y" then
		        sqlStr = sqlStr + " and isNULL(ct.infoDiv,'')<>''"
		    elseif FRectInfoDivYn="N" then
    		    sqlStr = sqlStr + " and isNULL(ct.infoDiv,'')=''"
    		else
    		    sqlStr = sqlStr + " and ct.infoDiv='"&FRectInfoDivYn&"'"
    		end if
		end if

		sqlStr = sqlStr + " order by m.regdate desc, i.itemid desc "
''rw sqlStr
''response.end

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem

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

				FItemList(i).FLotteRegdate		= rsget("LotteRegdate")
				FItemList(i).FLotteLastUpdate	= rsget("LotteLastUpdate")
				FItemList(i).FLotteGoodNo		= rsget("LotteGoodNo")
				FItemList(i).FLotteTmpGoodNo	= rsget("LotteTmpGoodNo")
				FItemList(i).FLottePrice		= rsget("LottePrice")
				FItemList(i).FLotteSellYn		= rsget("LotteSellYn")
				FItemList(i).FregUserid			= rsget("regUserid")
				FItemList(i).FLotteStatCd		= rsget("lotteStatCd")
				FItemList(i).FCateMapCnt		= rsget("mapCnt")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdefaultdeliverytype = rsget("defaultdeliverytype")
                FItemList(i).FdefaultfreeBeasongLimit = rsget("defaultfreeBeasongLimit")

				if Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
				else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				end if
                FItemList(i).FrctSellCNT      = rsget("rctSellCNT")
                FItemList(i).FinfoDiv       = rsget("infoDiv")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'--------------------------------------------------------------------------------
	'// 미등록 상품 목록(등록용)
	public Sub getLotteNotRegItemList
		dim sqlStr, addSql, i

		if FRectItemID<>"" then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"

			'''2013-07-25 김진영 옵션 추가금액 있는경우, 옵션금액 팝업에서 설정한 것만
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & " select itemid from ("
            addSql = addSql & "     select o.itemid"
            addSql = addSql & " 	,count(*) as optCNT"
            addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	from db_item.dbo.tbl_item_option as o "
            addSql = addSql & " 	left join db_item.dbo.tbl_lotte_regItem as RR on o.itemid = RR.itemid and RR.itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	where o.itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and o.isusing='Y'"
            addSql = addSql & " 	and isnull(RR.optAddPrcRegType,'') = '0'"
            addSql = addSql & " 	group by o.itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " where optAddCNT>0"
            addSql = addSql & " or (optCnt-optNotSellCnt<1)"
            addSql = addSql & " )"
		end if

		strSql = "Select top " & FPageSize & " i.* "
		strSql = strSql & "		, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay, UC.socname_kor "
		''strSql = strSql & "		, B.returnCode"
		strSql = strSql & "		, C.infoDiv,isNULL(C.safetyyn,'N') as safetyyn,isNULL(C.safetyDiv,0) as safetyDiv,C.safetyNum " '' 품목정보 및 안전인증정보. 추가 20121102
		strSql = strSql & " From db_item.dbo.tbl_item as i "
		strSql = strSql & " 	join db_item.dbo.tbl_item_contents as c "
		strSql = strSql & " 		on i.itemid=c.itemid "
		strSql = strSql & " 	left Join (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotte_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm "
		strSql = strSql & " 		on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " 	JOIN db_user.dbo.tbl_user_c UC on i.makerid = UC.userid"
		strSql = strSql & " 	JOIN db_partner.dbo.tbl_partner p on i.makerid = p.id"
		''strSql = strSql & " 	left Join db_item.dbo.tbl_OutMall_BrandReturnCode B"
		''strSql = strSql & " 	    on i.makerid=B.makerid and B.mallid='"&CMALLNAME&"'"
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & "     and p.jungsan_gubun not in ('간이과세', '면세')"		'2014-01-24 김진영// 이문재 이사님요청. 간이과세, 면세인 브랜드는 등록에서 뺴기
		strSql = strSql & "     and i.isExtUsing='Y' "
		strSql = strSql & "     and i.deliverytype not in ('7')"
		strSql = strSql & " 	and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		strSql = strSql & "     and i.sellyn='Y' "
		strSql = strSql & "     and i.deliverfixday not in ('C','X') "																				'플라워/화물배송 상품 제외
		strSql = strSql & "     and i.basicimage is not null "
		strSql = strSql & "     and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & "     and i.cate_large<>'' "
		strSql = strSql & "     and i.cate_large<>'999' "
		strSql = strSql & "     and i.sellcash>0 "
		strSql = strSql & "	and UC.isExtUsing <> 'N'"
		strSql = strSql & "     and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		''strSql = strSql & "     and i.sellcash=i.orgprice"              '''당분간 할인 안하는것만.. // 가격수정 모듈 없음..?
		''strSql = strSql & " 	and (i.orgprice<>0 and ((i.orgprice-i.orgSuplyCash)/i.orgprice)*100>=" & CMAXMARGIN & ")"							'역마진 상품 제외
		strSql = strSql & " 	and (i.sellcash<>0 and ((i.sellcash-i.buycash)/i.sellcash)*100>=" & CMAXMARGIN & ")"
		strSql = strSql & "		and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "		and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "		and i.itemid not in (Select itemid From db_item.dbo.tbl_lotte_regItem  where lottestatCD not in ('00','10') ) "		    '롯데등록상품 제외  ,'10' -- 오류
		strSql = strSql & "		and cm.mapCnt is Not Null "	& addSql																				'카테고리 매칭 상품만
''rw strSql
		rsget.Open strSql,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CLotteItem
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
                ''FItemList(i).Fcorp_dlvp_sn      = rsget("returnCode")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")

                FItemList(i).FrequireMakeDay    = rsget("requireMakeDay")

                FItemList(i).FinfoDiv       = rsget("infoDiv")
                FItemList(i).Fsafetyyn      = rsget("safetyyn")
                FItemList(i).FsafetyDiv     = rsget("safetyDiv")
                FItemList(i).FsafetyNum     = rsget("safetyNum")
                FItemList(i).Fsocname_kor	= rsget("socname_kor")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


	'--------------------------------------------------------------------------------
	'// 롯데닷컴 상품 목록(수정용)
	public Sub getLotteEditedItemList
		dim sqlStr, addSql, i

		if FRectItemID<>"" then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		elseif FRectNotJehyu="Y" then
			'제휴몰 상품이 아닌것
			addSql = " and i.isExtUsing='N' "
		else
			'수정된 상품만
			addSql = " and m.LotteLastUpdate<i.lastupdate"
		end if

        ''//연동 제외상품
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt<getdate()"
        addSql = addSql & "     and edDt>getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = "Select top " & FPageSize & " i.* "
		strSql = strSql & "		, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay, uc.socname_kor "
		strSql = strSql & "		, m.LotteGoodNo, m.LotteTmpGoodNo, m.LotteSellYn, isNULL(m.regedOptCnt,0) as regedOptCnt "
		strSql = strSql & "		, m.accFailCNT, m.lastErrStr "
		''strSql = strSql & "		, B.returnCode"
		strSql = strSql & "		, C.infoDiv,isNULL(C.safetyyn,'N') as safetyyn,isNULL(C.safetyDiv,0) as safetyDiv,C.safetyNum " '' 품목정보 및 안전인증정보. 추가 20121102

        strSql = strSql & "		,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType=9) and (i.sellcash<10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv>=50 or i.itemdiv='08' or i.cate_large='999' or i.cate_large=''"
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		THEN 'Y' ELSE 'N' END) as maySoldOut"

		strSql = strSql & " From db_item.dbo.tbl_item as i "
		strSql = strSql & " 	join db_item.dbo.tbl_item_contents as c "
		strSql = strSql & " 		on i.itemid=c.itemid "
		strSql = strSql & " 	join db_item.dbo.tbl_lotte_regItem as m "
		strSql = strSql & " 		on i.itemid=m.itemid "
		strSql = strSql & " 	left Join (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_lotte_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm "
		strSql = strSql & " 		on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " 	left join db_user.dbo.tbl_user_c uc"
		strSql = strSql & " 		on i.makerid=uc.userid"

		''strSql = strSql & " 	left Join db_item.dbo.tbl_OutMall_BrandReturnCode B"
		''strSql = strSql & " 	    on i.makerid=B.makerid and B.mallid='"&CMALLNAME&"'"
		strSql = strSql & " Where 1=1"
		''if (FRectMatchCateNotCheck<>"on") then
		IF (FRectMatchCate="Y") then '' eastone 수정 2013/09/01
		    strSql = strSql & " and cm.mapCnt is Not Null "
	    end if
		strSql = strSql & addSql
		''strSql = strSql & " and m.LotteGoodNo is Not Null "									'#등록 상품만
		strSql = strSql & " and isNULL(m.LotteTmpGoodNo,m.LotteGoodNo) is Not Null "									'#등록 상품만
''rw strSql
		rsget.Open strSql,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CLotteItem
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
				FItemList(i).FLotteGoodNo		= rsget("LotteGoodNo")
				FItemList(i).FLotteTmpGoodNo	= rsget("LotteTmpGoodNo")
				FItemList(i).FLotteSellYn		= rsget("LotteSellYn")

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
                FItemList(i).Fsocname_kor    = rsget("socname_kor")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	Public Sub OnlyLotteNotUpdateMakeridList
		Dim sqlStr, i, addsql

		If FRectMakerId <> "" Then
			addSql = " and makerid in ('" & FRectMakerId & "')"
		End If

		sqlStr = " select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " From db_temp.dbo.tbl_Lotte_not_in_makerid_By_KimJinYoung "
		sqlStr = sqlStr & " Where 1=1 " & addsql

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage)>Cint(FTotalPage) Then
			FResultCount = 0
			Exit sub
		End If

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_Lotte_not_in_makerid_By_KimJinYoung "
		sqlStr = sqlStr & " Where 1=1 " & addsql
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLotteItem
					FItemList(i).FMakerid		= rsget("makerid")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FRegId			= rsget("regId")
					FItemList(i).FLastupdate	= rsget("lastupdate")
					FItemList(i).FLastupdateId	= rsget("lastupdateId")
					FItemList(i).FIsUsing		= rsget("isUsing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub

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
	dim strSql, rstStr
	rstStr = "<Select name='" & fnm & "' class='select'>"
	rstStr = rstStr & "<option value=''>전체</option>"

	strSql = "Select * From db_temp.dbo.tbl_lotte_MDCateGrp Where isUsing='Y'"
	rsget.Open strSql,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			if cStr(rsget("groupCode"))=cStr(selcd) then
				rstStr = rstStr & "<option value='" & rsget("groupCode") & "' selected>" & rsget("groupName")& "  ["& rsget("groupCode") &"]</option>"
			else
				rstStr = rstStr & "<option value='" & rsget("groupCode") & "'>" & rsget("groupName")& "  ["& rsget("groupCode") &"]</option>"
			end if
			rsget.MoveNext
		Loop
	end if
	rsget.Close

	rstStr = rstStr & "</select>"

	printLotteCateGrpSelectBox = rstStr
end Function

'// 상품이미지 존재여부 검사
function ImageExists(byval iimg)
	if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
		ImageExists = false
	else
		ImageExists = true
	end if
end function

Function GetRaiseValue(value)
    If Fix(value) < value Then
    GetRaiseValue = Fix(value) + 1
    Else
    GetRaiseValue = Fix(value)
    End If
End Function

function getLotteItemIdByTenItemID(iitemid)
    dim sqlStr, retVal
    sqlStr = " select isNULL(lotteGoodNo,lotteTmpGoodNo) as lotteGoodNo "&VbCRLF
    sqlStr = sqlStr & " from db_item.dbo.tbl_lotte_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&iitemid&VbCRLF

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    retVal = rsget("lotteGoodNo")
	end if
	rsget.Close

	if IsNULL(retVal) then retVal=""
	getLotteItemIdByTenItemID = retVal
end function

''//상품명 변경 파라메터 생성
Function fnGetLotteItemNameEditParameter(iLotteGoodNo,iItemName)
    dim strRst
    strRst = "subscriptionId=" & lotteAuthNo
    strRst = strRst & "&strGoodsNo=" & iLotteGoodNo
    strRst = strRst & "&strGoodsNm=" & Server.URLEncode(Trim(iItemName))
    strRst = strRst & "&strMblGoodsNm=" & Server.URLEncode(Trim(iItemName))
    strRst = strRst & "&strChgCausCont=" & Server.URLEncode("api 상품명 변경")
    fnGetLotteItemNameEditParameter = strRst
end function

Function getLotteTmpItemIdByTenItemID(iitemid)
	Dim sqlStr, retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT lotteTmpGoodNo, isnull(lotteGoodNo,'') as lotteGoodNo " & VBCRLF
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_lotte_regItem" & VBCRLF
	sqlStr = sqlStr & " WHERE itemid = "&iitemid & VBCRLF
	rsget.Open sqlStr,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		If rsget("lotteGoodNo") <> "0" Then
			retVal = "전시상품"
		Else
			retVal = rsget("lotteTmpGoodNo")
		End If
	End If
	rsget.Close

	If IsNULL(retVal) Then retVal = ""
	getLotteTmpItemIdByTenItemID = retVal
End Function

%>