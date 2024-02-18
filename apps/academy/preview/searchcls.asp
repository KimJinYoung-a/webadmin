<%
''CS_EUCKR <=> CS_UTF8
''group by 필드는 널이 없게 하자.
'''--------------------------------------------------------------------------------------
DIM G_ORGSCH_ADDR		'## 인덱싱서버
DIM G_1STSCH_ADDR		'## 1차조회서버
DIM G_2NDSCH_ADDR		'## 2차조회서버
DIM G_3RDSCH_ADDR		'## 3차조회서버
Dim G_4THSCH_ADDR		'## 4차조회서버

DIM G_SCH_TIME : G_SCH_TIME=formatdatetime(now(),4)

IF (application("Svr_Info") = "Dev") THEN
     G_1STSCH_ADDR = "61.252.133.10"  ''"110.93.128.109" ''
     G_2NDSCH_ADDR = "61.252.133.10"
     G_3RDSCH_ADDR = "61.252.133.10"
     G_4THSCH_ADDR = "61.252.133.10"
     G_ORGSCH_ADDR = "61.252.133.10"
ELSE
     G_1STSCH_ADDR = "192.168.0.206"        ''192.168.0.210  :: 검색페이지(search.asp)   '
     G_2NDSCH_ADDR = "192.168.0.206"        ''192.168.0.207  :: 카테고리, 상품, 브랜드
     G_3RDSCH_ADDR = "192.168.0.206"        ''192.168.0.209  :: GiftPlus , scn_dt_itemDispColor :: 확인.
     G_4THSCH_ADDR = "192.168.0.206"        ''192.168.0.208  :: mobile 6:10분에 인덱싱 시작 카피
     G_ORGSCH_ADDR = "192.168.0.206"        ''192.168.0.206  
END IF

''sample in doc
function escapeQuery( istr )
	dim ret, c, i
	ret = ""
	For i=1 To Len(istr)
		c = Mid(istr,i,1)
		select case c
		case "\"
			ret = ret & "\\"
		case "'"
			ret = ret & "\'"
		case chr(34)
			ret = ret & "\" & chr(34)
		case "*"
			ret = ret & "\*"
		case else
			ret = ret & c
		end select
	Next
	escapeQuery = ret
end function

function getTimeChkAddr(defaultAddr)
    '''6시10분 1차섭 인덱싱 및 2차서버로 Copy
    '''6시50분~ 1차=>3차서버로 Copy
    getTimeChkAddr = defaultAddr

    IF (defaultAddr=G_4THSCH_ADDR) THEN
        IF (G_SCH_TIME>"06:00:00") and (G_SCH_TIME<"06:40:00") then
            getTimeChkAddr = G_2NDSCH_ADDR
        END IF
    ELSE
        IF (G_SCH_TIME>"06:40:00") and (G_SCH_TIME<"07:00:00") then
            getTimeChkAddr = G_4THSCH_ADDR
        END IF
    END IF
end function

function debugQuery(iDocruzer,Scn,iSearchQuery,iSortQuery,iFTotalCount,iFResultcount)
  ''exit function
    IF Not (application("Svr_Info")="Dev") THEN
        exit function
    ENd IF

    dim itime
    Call iDocruzer.GetResult_SearchTime(itime) '소요시간
    'rw "-------------------------------"
    'rw Scn
    'rw iSearchQuery
    'rw iSortQuery
    'rw "FTotalCount:"&iFTotalCount
    'rw "FResultcount:"&iFResultcount
    'rw "GetResult_SearchTime:"&itime
end function
'''--------------------------------------------------------------------------------------

Class SearchGroupByItems

	Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB

	PUBLIC FImageSmall
	PUBLIC FSubTotal

	PUBLIC FCateCode
	PUBLIC FCateName
	PUBLIC FCateCd1
	PUBLIC FCateCd2
	PUBLIC FCateCd3
	PUBLIC FCateDepth

	PUBLIC FcolorCode
	PUBLIC FcolorName
	PUBLIC FcolorIcon

	PUBLIC FStyleCd
	PUBLIC FStyleName

	PUBLIC FAttribCd
	PUBLIC FAttribName

	PUBLIC FminPrice
	PUBLIC FmaxPrice
	PUBLIC FnewCateLarge
	PUBLIC FnewCateMid

End Class


'###################################	상품 관련 검색	######################################################################
Class SearchItemCls

	Private SUB Class_initialize()
        ''기본 1차 서버.------------------------
		SvrAddr = getTimeChkAddr(G_1STSCH_ADDR)
		''--------------------------------------

		SvrPort = "6167"'DocSvrPort

		AuthCode = "" '인증값

		Logs = "" '로그값

		FResultCount = 0
		FTotalCount = 0
		FPageSize = 10
		FCurrPage = 1
		FPageSize = 30
		FRectColsSize =5
		FLogsAccept = false

	End SUB

	Private SUB Class_Terminate()

	End SUB

	dim FItemList
	dim FPageSize
	dim FCurrPage
	dim FScrollCount
	dim FResultCount
	dim FTotalCount
	dim FTotalPage

	dim FRectSearchTxt		'검색어
	dim FRectSearchItemDiv	'카테고리 검색 범위 (y:기본 카테고리만, n:추가 카테고리 폼함)
	dim FRectSearchCateDep	'카테고리 검색 범위 (X:해당 카테고리만, T:하위 카테고리 포함)
	dim FRectPrevSearchTxt	'이전 검색어
	dim FRectExceptText		'제외어
	dim FRectSortMethod		'정렬방식 (ne:신상품, be:인기상품, lp:낮은가격, hp:높은가격, hs:할인률, br:상품후기, ws:위시수)
	dim FRectSearchFlag 	'검색범위 (sc:세일쿠폰, ea:사용기전체, ep:포토사용기, ne:신상품, fv:위시상품, pk:포장서비스)
	dim FRectNotMakerID		'makerid 제외

	dim FRectMakerid		'업체 아이디
	dim FRectCateCode		'카테고리코드
	dim FListDiv			'카테고리/검색 구분용
	dim FSellScope			'판매가능 상품검색 여부
	dim FGroupScope			'검색시 그룹핑 범위 (1:1depth, 2:2depth, 3:3depth)
	dim FdeliType			'배송방법 (FD:무료배송, TN:텐바이텐 배송, FT:무료+텐바이텐 배송, WD:해외배송)

	dim FcolorCode			'상품컬러칩
	dim FstyleCd			'상품스타일
	dim FattribCd			'상품속성

	dim FminPrice			'가격최소값
	dim FmaxPrice			'가격최대값
	dim FSalePercentHigh	'할인율 최대값
	dim FSalePercentLow		'할인율 최소값

	dim FCheckResearch 		'결과내 재검색 체크용
	dim FRectColsSize		'결과 리스트 열수
	dim FLogsAccept			'추가 로그 저장 여부

	dim FarrCate			'복수 카테고리
	dim FisTenOnly			'텐바이텐 전용상품
	dim FisLimit			'한정판매상품
	dim FisFreeBeasong
	dim FRectCPidx			'쿠폰idx
	
	dim FRectSearchGubun	'검색구분.상품,강좌,매거진

	Private SvrAddr
	Private SvrPort
	Private AuthCode
	Private Logs
	Private Scn
	private strQuery
	Private Order
	Private StartNum

	Private SearchQuery
	Private SortQuery
    
    public function InitDocruzer(iDocruzer)
        InitDocruzer = FALSE
        IF ( iDocruzer.BeginSession() < 0 ) THEN
			EXIT function
		End If
        
        IF NOT DocSetOption(iDocruzer) THEN
			EXIT function
		End If
		InitDocruzer = TRUE
    End function

    public function DocSetOption(iDocruzer)
        dim ret 
        ret = iDocruzer.SetOption(iDocruzer.OPTION_REQUEST_CHARSET_UTF8,1)
        DocSetOption = (ret>=0)
    end function
    
    

	''/검색 조건 설정
	FUNCTION getSearchQuery(byref query)
		dim strQue, arrCCD, arrSCD, arrACD, lp

		'### 검색구분에 따른 기본값 확인 및 설정 ###
		Select Case FListDiv
			Case "search"
				'검색 페이지 결과
				IF (FRectSearchTxt="" or isNull(FRectSearchTxt)) Then EXIT FUNCTION
			Case "list"
				'카테고리 목록
				IF (FRectCateCode="" or isNull(FRectCateCode)) Then EXIT FUNCTION
			Case "fulllist"
				'카테고리없는 전체.
			Case Else
				EXIT FUNCTION
		End Select

		'### 검색조건 생성 ###

		'@ 검색어(키워드)
		IF FRectSearchTxt<>"" Then
			FRectSearchTxt = chgCoinedKeyword(FRectSearchTxt)
			FRectSearchTxt = escapeQuery(FRectSearchTxt)  ''2015 추가
			
			IF FRectExceptText<>"" Then
			    FRectExceptText = escapeQuery(FRectExceptText)  ''2015 추가
				strQue = getQrCon(strQue) & "(idx_itemname='" & FRectSearchTxt & " ! " & FRectExceptText & "' BOOLEAN) "	'제외어
			else
				strQue = getQrCon(strQue) & "idx_itemname='" & FRectSearchTxt & "'  allword "	'키워드검색(동의어 포함) synonym
				'strQue = getQrCon(strQue) & "idx_itemname='" & FRectSearchTxt & "'  natural "		'자연어 검색(동의어 포함) synonym
			End if
		END IF
		
		'@ 검색 제외 브랜드
		IF FRectNotMakerID <> "" Then
			strQue = strQue & getQrCon(strQue) & "idx_makerid != '" & FRectNotMakerID & "' "
		End IF

		'@ 카테고리 검색 범위 idx_isDefault 삭제
		''IF FRectSearchItemDiv="y" Then
		''	''strQue = strQue & getQrCon(strQue) & "idx_isDefault='y' "
		''END IF

		'@ 카테고리
		IF FRectCateCode<>"" Then
			if FRectSearchCateDep="X" then
				strQue = strQue & getQrCon(strQue) & "idx_catecode='" & FRectCateCode & "'"
			else
			    IF FRectSearchItemDiv="y" Then ''기본카테고리
			        strQue = strQue & getQrCon(strQue) & "idx_catecode like '" & FRectCateCode & "*'"
			    else                           ''추가카테검색
			        strQue = strQue & getQrCon(strQue) & "idx_catecodeExt like '" & FRectCateCode & "*'"
			    end if
			end if
		END IF

		'@ 복수 카테고리
		IF FarrCate<>"" THEN
			dim arrCt, lpCt
			if right(FarrCate,1)="," then FarrCate=left(FarrCate,len(FarrCate)-1)
			arrCt = split(FarrCate,",")
			strQue = strQue & getQrCon(strQue) & "("
			for lpCt=0 to ubound(arrCt)
				if FRectSearchCateDep="X" then
					strQue = strQue & " idx_catecode='" & RequestCheckVar(LCase(trim(arrCt(lpCt))),18) & "' "
				else
					strQue = strQue & " idx_catecode like '" & RequestCheckVar(LCase(trim(arrCt(lpCt))),18) & "*' "
				end if
				if lpCt<ubound(arrCt) then strQue = strQue & " or "
			next
			strQue = strQue & " )"
		END IF

		'@ 검색범위
		IF FRectSearchFlag<>"" THEN
			Select Case FRectSearchFlag
				Case "sc"	'세일쿠폰
					strQue= strQue & getQrCon(strQue) & "(idx_saleyn='Y' or idx_itemcouponyn='Y') "
				Case "ea"	'전체사용기
					strQue= strQue & getQrCon(strQue) & "(idx_evalcnt>0) "
				Case "ep"	'포토사용기
					strQue= strQue & getQrCon(strQue) & "(idx_evalcntPhoto>0) "
				Case "ne"	'신상품
					strQue = strQue & getQrCon(strQue) & "idx_newyn='Y' "
				Case "fv"	'위시상품
					strQue = strQue & getQrCon(strQue) & "(idx_favcount>0) "
				Case "pk"	'포장서비스
					strQue = strQue & getQrCon(strQue) & "idx_pojangok='Y' "
			End Select
		END IF

		'@ 브랜드
		IF FRectMakerid<>"" THEN
			dim arrMkr, lpMkr
			if right(FRectMakerid,1)="," then FRectMakerid=left(FRectMakerid,len(FRectMakerid)-1)
			arrMkr = split(FRectMakerid,",")
			strQue = strQue & getQrCon(strQue) & "("
			for lpMkr=0 to ubound(arrMkr)
				strQue = strQue & " idx_makerid='" & RequestCheckVar(LCase(trim(arrMkr(lpMkr))),32) & "'  "
				if lpMkr<ubound(arrMkr) then strQue = strQue & " or "
			next
			strQue = strQue & " )"
		END IF

		'@ 가격범위
		if FminPrice<>"" then
			strQue = strQue & getQrCon(strQue) & "idx_sellcash>='" & FminPrice & "'"
		end if
		if FmaxPrice<>"" then
			strQue = strQue & getQrCon(strQue) & "idx_sellcash<='" & FmaxPrice & "'"
		end if

		'@ 할인범위
		IF FSalePercentHigh<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_salepercent >=" & (1-FSalePercentHigh)*100 & " "
		End IF
		IF FSalePercentLow<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_salepercent <" & (1-FSalePercentLow)*100 & " "
		End IF

		'@ 배송방법
		Select Case FdeliType
			Case "FD"	'무료배송
				strQue = strQue & getQrCon(strQue) & "isFreeBeasong='Y'"
			Case "TN"	'텐바이텐배송
				strQue = strQue & getQrCon(strQue) & "(deliverytype='1' or deliverytype='4')"
			Case "FT"	'무료 + 텐바이텐배송
				strQue = strQue & getQrCon(strQue) & "(deliverytype='1' or deliverytype='4') and isFreeBeasong='Y'"
			Case "WD"	'해외배송
				strQue = strQue & getQrCon(strQue) & "isAboard='Y'"
		end Select

		'@ 텐바이텐 전용상품
		IF FisTenOnly="only" Then
			strQue = strQue & getQrCon(strQue) & "idx_tenOnlyYn='Y' "
		End IF

		'@ 한정상품
		IF FisLimit="limit" Then
			strQue = strQue & getQrCon(strQue) & "idx_limityn='Y' "
		End IF

		'@ 무료배송상품
		IF FisFreeBeasong="free" Then
			strQue = strQue & getQrCon(strQue) & "idx_isFreeBeasong='Y' "
		End If

		'@ 상품 판매 범위
		IF FSellScope="Y" Then
			strQue = strQue & getQrCon(strQue) & "idx_sellyn='Y' "
		ELSE
			strQue = strQue & getQrCon(strQue) & "(idx_sellyn='Y' or idx_sellyn='S') "
		End IF

		'@ 쿠폰번호
		IF FRectCPidx<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_curritemcouponidx='" & FRectCPidx & "'"
		END IF

        ''2015 추가 string list group by 사용시 해당필드에 널값이 있는경우 빈값 대신 000 넣음.
        IF  scn="scn_dt_itemAcademyAttribGroup" then
            strQue = strQue & getQrCon(strQue) & "idx_attribgrp!='000' "
        ELSEIF  scn="scn_dt_itemAcademyCateGroup" then
            if (FGroupScope="2") then
                strQue = strQue & getQrCon(strQue) & "idx_cd2grp!='000' "
            elseif (FGroupScope="3") then
                strQue = strQue & getQrCon(strQue) & "idx_cd3grp!='000' "
            end if
        END IF
        
		query = strQue
	End FUNCTION

	Sub getSortQuery(byref query)
		dim strQue

		'// 중복 상품 제거(중복 등록 카테고리일경우) 2015 제거
		'IF (FRectCateCode<>"" and FRectSearchItemDiv<>"y") Then '' 추가 카테고리 검색시
    	'	'strQue = " GROUP BY itemid"
    	'END IF

		'// 정렬
		IF FRectSortMethod="ne" THEN '신상품
			strQue = strQue & " ORDER BY itemid desc"
		ELSEIF FRectSortMethod="be" THEN '인기상품
			if FRectSearchFlag="fv" then
				'위시상품보기는 정렬을 위시순으로
				strQue = strQue & " ORDER BY favcount desc,itemscore desc,itemid desc"
			else
				strQue = strQue & " ORDER BY itemscore desc,itemid desc"
			end if
		ELSEIF FRectSortMethod="lp" THEN '낮은가격
			strQue = strQue & " ORDER BY sellcash "
		ELSEIF FRectSortMethod="hp" THEN'높은가격
			strQue = strQue & " ORDER BY sellcash desc"
		ELSEIF FRectSortMethod="hs" THEN '핫세일 (할인율이 높은순)
			strQue = strQue & " ORDER BY salepercent desc, saleprice desc"
		ELSEIF FRectSortMethod="br" THEN '베스트후기순
			strQue = strQue & " ORDER BY evalcnt desc,itemid desc"
		ELSEIF FRectSortMethod="ws" THEN '위시순
			strQue = strQue & " ORDER BY favcount desc,itemid desc"
		ELSEIF FRectSortMethod="pj" THEN '인기포장순
			strQue = strQue & " ORDER BY pojangcnt desc,itemid desc"
		ELSE
			strQue = strQue & " ORDER BY itemid desc"
		END IF
		
		query = strQue
	End Sub

	Function getQrCon(query)
		if Not(query="" or isNull(query)) then
			getQrCon = " and "
		end if
	End Function

	'// 상품 이미지 폴더 반환(컬러코드 유무에 따라 일반/컬러칩 구분)
	Function getItemImageUrl()
		IF application("Svr_Info")	= "Dev" THEN
				getItemImageUrl = "http://testimage.thefingers.co.kr"
		Else
				getItemImageUrl = "http://image.thefingers.co.kr"
		End If
	end function

	'####### 상품 검색 - 검색 엔진 ######
	PUBLIC SUB getSearchList()

		DIM Scn
		DIM Docruzer,ret

		DIM Logs ,iRows
		DIM arrData,arrSize, retMatchCd, retMatchVal

		'// 검색 결과 출력 시나리오명
		if FcolorCode="" or FcolorCode="0" then
			Scn= "scn_dt_itemAcademy"		'일반 상품 검색
		else
			'Scn= "scn_dt_itemColor"		'상품 컬러별 검색
			'Scn= "scn_dt_itemAcademyColor"	'상품 컬러별 검색(전시카테고리)
			Scn= "scn_dt_itemAcademy"		    '일반 상품 검색 통일 2015
		end if

		StartNum = (FCurrPage -1)*FPageSize '// 검색시작 Row

		CALL getSearchQuery(SearchQuery)	'// 검색 쿼리생성
		CALL getSortQuery(SortQuery)		'// 정렬 쿼리 생성
		'Response.Write SearchQuery &"<Br>"
		IF SearchQuery="" THEN
			EXIT SUB
		END IF

		IF (FLogsAccept) and (FRectSearchTxt<>"") and (FCurrPage="1") THEN
            'Logs = "상품+^" & FRectSearchTxt & "]##" & FRectSearchTxt & "||" & FRectPrevSearchTxt  	'// 로그값
            
            ''2015 search4
            '기본:[사이트@카테고리+사용자$성별코드|연령|검색어타입(서비스)|첫검색|페이지번호|정렬순^이전검색어##검색어] ''기본
            Dim iLOG_SITE : iLOG_SITE = "FINGERS"
            Dim iLOG_CATE : iLOG_CATE = "MOB"
            Dim iLOG_USER : iLOG_USER = GetUserLevelStr(GetLoginUserLevel) '' 회원등급을 사용
            Dim iLOG_SEX  : iLOG_SEX  = "" '' 0비로그인,1남성,2여성
            Dim iLOG_AGE  : iLOG_AGE  = "" '' 0비로그인,1:10대,2:20대,3:30대,4:40대,5:50대
            Dim iLOG_STYPE : iLOG_STYPE = "" '' 서비스 사용안함 X
            Dim iLOG_FIRST : iLOG_FIRST = "" '' 첫검색/재검색 사용안함 X  FCheckResearch
            
            Logs = iLOG_SITE&"@"                ''[ @
            Logs = Logs&iLOG_CATE&"+"           ''@ +
            Logs = Logs&iLOG_USER&"$"           ''+ $
            Logs = Logs&iLOG_SEX&"|"            ''$ |
            Logs = Logs&iLOG_AGE&"|"            ''| | 
            Logs = Logs&iLOG_STYPE&"|"          ''| | 
            Logs = Logs&iLOG_FIRST&"|"          ''| | 
            Logs = Logs&FCurrPage&"|"           ''| | 
            Logs = Logs&FRectSortMethod&"^"     ''| ^ 
            Logs = Logs&FRectPrevSearchTxt&"##" ''^ ##
            Logs = Logs&FRectSearchTxt          ''## ]
            
		END IF

		'##### 대상서버선택 '### 상품검색인경우 1차서버
		SvrAddr = getTimeChkAddr(G_1STSCH_ADDR)

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF
		
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)
'response.write Docruzer.GetErrorMessage()
'response.end
		IF( ret < 0 ) THEN
			dbACADEMYget.execute "EXECUTE db_academy.dbo.[sp_Academy_DocLog] @ErrMsg ='"& html2db(SearchQuery) & "[" & html2db(Docruzer.GetErrorMessage()) &"]'"

			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING

			IF FListDiv<>"search" THEN
				'// 1번 서버 에러시 2번에서 구동(2번도 에러면 Skip)
				if (SvrAddr = G_1STSCH_ADDR) then
					SvrAddr = G_2NDSCH_ADDR  ''"192.168.0.108"
					if (G_1STSCH_ADDR<>G_2NDSCH_ADDR) then  ''추가 2013/09
					    call getSearchList()
				    end if
				end if
			END IF

			EXIT SUB
		END IF

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수
		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		'Response.write "검색결과수 : " & FTotalCount & "<br>"
		IF( FResultCount <= 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB 'Response.write "GetResult_RowSize: " & Docruzer.GetErrorMessage()
		END IF

		FTotalPage =  Cdbl(FTotalCount\FPageSize)
		IF  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) THEN
			FtotalPage = FtotalPage +1
		END IF

		REDIM FItemList(FResultCount)

		FOR iRows=0 to FResultCount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.GetErrorMessage()
				EXIT FOR
			END IF

			SET FItemList(iRows) = NEW CCategoryPrdItem

				FItemList(iRows).FCateCode = arrData(0)
				FItemList(iRows).FarrCateCd = arrData(2)
				FItemList(iRows).FItemDiv	= arrData(3)
				FItemList(iRows).FItemid = arrData(4)
				FItemList(iRows).FItemName = db2html(arrData(5))
				FItemList(iRows).FKeyWords = db2html(arrData(6))
				FItemList(iRows).FSellCash = arrData(7)
				FItemList(iRows).FOrgPrice = arrData(8)
				FItemList(iRows).FMakerId = arrData(9)
				FItemList(iRows).FBrandName = db2html(arrData(10))
			    FItemList(iRows).FImageBasic 	= getItemImageUrl & "/diyitem/webimage/basic/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(11))
				FItemList(iRows).FImageMask 	= getItemImageUrl & "/diyitem/webimage/mask/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(12))
				FItemList(iRows).FImageList 	= getItemImageUrl & "/diyitem/webimage/list/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(13))
				FItemList(iRows).FImageList120 	= getItemImageUrl & "/diyitem/webimage/list120/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(14))
				FItemList(iRows).FImageIcon1 	= getItemImageUrl & "/diyitem/webimage/icon1/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(15))
				FItemList(iRows).FImageIcon2 	= getItemImageUrl & "/diyitem/webimage/icon2/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(16))
				FItemList(iRows).FImageSmall	= getItemImageUrl & "/diyitem/webimage/small/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(17))
				if (arrData(47)<>"") then ''2015 추가 (추가 이미지)
				    if (FcolorCode="" or FcolorCode="0") then
				    	FItemList(iRows).FAddimage      = getItemImageUrl & "/diyitem/webimage/add1/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" & db2html(arrData(47))
				    else
				    	FItemList(iRows).FAddimage      = replace(getItemImageUrl,"/color","/image") & "/diyitem/webimage/add1/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" & db2html(arrData(47))
				    end if
			    elseif (arrData(12)<>"") then	'추가이미지가 없고 마스킹이 있는 경우 마스킹을 추가이미지로 처리(2015.09.07; 허진원)
			    	FItemList(iRows).FAddimage      = FItemList(iRows).FImageMask
			    end if
				FItemList(iRows).FSellyn = arrData(18)
				FItemList(iRows).FSaleyn = arrData(19)
				FItemList(iRows).FLimityn = arrData(20)
				FItemList(iRows).FRegdate = dateserial(left(arrData(21),4),mid(arrData(21),5,2),mid(arrData(21),7,2))
				IF arrData(22)<>"" Then
					FItemList(iRows).FReipgodate= dateserial(left(arrData(22),4),mid(arrData(22),5,2),mid(arrData(22),7,2))
				End IF
				FItemList(iRows).FItemcouponyn = arrData(23)
				FItemList(iRows).FItemCouponValue = arrData(24)
				FItemList(iRows).FItemCouponType = arrData(25)
				FItemList(iRows).FEvalCnt = arrData(26)
				FItemList(iRows).FEvalcnt_Photo = arrData(27)
				FItemList(iRows).FfavCount = arrData(28)
				FItemList(iRows).FItemScore = arrData(29)
				FItemList(iRows).FtenOnlyYn = arrData(33)

                FItemList(iRows).Frecentsellcount = arrData(48) ''//2015 추가
                FItemList(iRows).FPojangOk = arrData(49)		''//2015.10.07
				If isNull(arrData(50)) OR arrData(50) = "" Then
					FItemList(iRows).FImgProfile = ""
				Else
					FItemList(iRows).FImgProfile = "http://" & CHKIIF(application("Svr_Info")="Dev","test","") & "image.thefingers.co.kr/corner/newImage_profile/thumbimg3/t3_" & arrData(50)
				End If
				FItemList(iRows).FRealKeyword = arrData(51)			'실제 저장된 키워드만 뽑음.

                'FItemList(iRows).FcolorCd = arrData(35)
                
                if (application("Svr_Info")<>"Dev") then
                    ''FItemList(iRows).FImageBasic = getStonThumbImgURL(FItemList(iRows).FImageBasic,300,200,true,false)
                    FItemList(iRows).FImageBasic = getStonReSizeImg(FItemList(iRows).FImageBasic,410,"",100)
                end if
			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

	'####### 상품 검색 카테고리별 카운팅  ######
	PUBLIC SUB getGroupbyCategoryList()

		'// 검색 결과 출력 시나리오명
		Scn= "scn_dt_itemAcademyCateGroup"		'일반 상품 검색

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// 검색시작 Row
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성
        
		'//그룹 범위별 지정(정렬 쿼리 생성)
		Select Case FGroupScope
			Case "1"
				SortQuery = " GROUP BY idx_cd1grp order by idx_cd1grp " 
			Case "2"
				SortQuery = " GROUP BY idx_cd2grp order by idx_cd2grp "
			Case "3"
				SortQuery = " GROUP BY idx_cd3grp order by idx_cd3grp "
			Case Else
				SortQuery = " GROUP BY idx_cd1grp order by idx_cd1grp "
		end Select
		
		IF SearchQuery="" Then
			EXIT SUB
		End If
		
		dim Rowids,Scores
		FTotalCount = 0

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF
		
 'response.write "group : " & SearchQuery & SortQuery & "<br>"
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
		    'response.write "ERR:"&Docruzer.GetErrorMessage()
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF


		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
'CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)

		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new SearchGroupByItems
				FItemList(iRows).FCateCode	= arrData(0)
				FItemList(iRows).FCateName	= arrData(1)
				FItemList(iRows).FCateCd1	= arrData(2)
				FItemList(iRows).FCateCd2	= arrData(3)
				FItemList(iRows).FCateCd3	= arrData(4)
				''2015 변경
				
				if Len(FItemList(iRows).FCateCd1)>=4 then
				    FItemList(iRows).FCateCd1	= Mid(FItemList(iRows).FCateCd1,5,255)              '' sort4+[code3]
				end if
				    
				if Len(FItemList(iRows).FCateCd2)>=11 then
				    FItemList(iRows).FCateCd2	= Mid(FItemList(iRows).FCateCd2,9+3,255)            '' sort4+4+code3+[code3]
				end if
				    
				if Len(FItemList(iRows).FCateCd3)>=18 then
				    FItemList(iRows).FCateCd3	= Mid(FItemList(iRows).FCateCd3,13+3+3,255)         '' sort4+4+4+code3+[code3]
				end if
				
				FItemList(iRows).FCateDepth	= arrData(5)
                
                ''rw FItemList(iRows).FCateCd1&"|"&FItemList(iRows).FCateCd2&"|"&FItemList(iRows).FCateCd3&"|"&FItemList(iRows).FCateDepth
                
				FItemList(iRows).FImageSmall = getItemImageUrl & "/small/" & GetImageSubFolderByItemid(arrData(6)) & "/" &db2html( arrData(7))
				FItemList(iRows).FSubTotal 	= Scores(iRows)
				FTotalCount = FTotalCount + Scores(iRows)
			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT

		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

	
	'####### 상품 검색 총 카운팅  ######
	PUBLIC SUB getTotalCount()

		'// 검색 결과 출력 시나리오명
		Scn= "scn_dt_itemAcademy"		'일반 상품 검색
		
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// 검색시작 Row
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성

		'IF (FRectCateCode<>"" and FRectSearchItemDiv<>"y") Then
		'    ' SortQuery = " GROUP BY itemid" ''2013 조건추가 '' 필요없음 2015
		'else
    	'	' SortQuery = " "	'// 정렬 쿼리 생성
    	'end if

		IF SearchQuery="" Then
			EXIT SUB
		End If

		dim Rowids,Scores

        '// 기본으로 검색2번 섭, 검색어가 있다면 1번섭 사용
        ''---------------------------------------------------------------------------------------------------------
        if (FRectCateCode<>"") or (FRectMakerid<>"") and (FRectSearchTxt="")  then
            SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR) ''G_4THSCH_ADDR
        end if
        ''---------------------------------------------------------------------------------------------------------

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF
 
	'rw "group : " & SearchQuery & SortQuery & "<br>" ''
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _  
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)
		
						
		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)


		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB
    

	PUBLIC FUNCTION HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	END FUNCTION

	PUBLIC FUNCTION HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	END FUNCTION

	PUBLIC FUNCTION StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	END FUNCTION

End Class


'###################################	강좌 관련 검색	######################################################################
Class SearchlecCls

	Private SUB Class_initialize()
        ''기본 1차 서버.------------------------
		SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR)
		''--------------------------------------

		SvrPort = "6167"'DocSvrPort

		AuthCode = "" '인증값

		Logs = "" '로그값

		FResultCount = 0
		FTotalCount = 0
		FPageSize = 10
		FCurrPage = 1
		FPageSize = 30
		FRectColsSize =5
		FLogsAccept = false

	End SUB

	Private SUB Class_Terminate()

	End SUB

	dim FItemList
	dim FPageSize
	dim FCurrPage
	dim FScrollCount
	dim FResultCount
	dim FTotalCount
	dim FTotalPage

	dim FRectSearchTxt		'검색어
	dim FRectSearchItemDiv	'카테고리 검색 범위 (y:기본 카테고리만, n:추가 카테고리 폼함)
	dim FRectSearchCateDep	'카테고리 검색 범위 (X:해당 카테고리만, T:하위 카테고리 포함)
	dim FRectPrevSearchTxt	'이전 검색어
	dim FRectExceptText		'제외어
	dim FRectSortMethod		'정렬방식 (ne:신상품, be:인기상품, lp:낮은가격, hp:높은가격, hs:할인률, br:상품후기, ws:위시수)
	dim FRectSearchFlag 	'검색범위 (sc:세일쿠폰, ea:사용기전체, ep:포토사용기, ne:신상품, fv:위시상품, pk:포장서비스)

	dim FRectMakerid		'업체 아이디
	dim FRectCateCode		'카테고리코드
	dim FListDiv			'카테고리/검색 구분용
	dim FSellScope			'판매가능 상품검색 여부
	dim FGroupScope			'검색시 그룹핑 범위 (1:1depth, 2:2depth, 3:3depth)
	dim FdeliType			'배송방법 (FD:무료배송, TN:텐바이텐 배송, FT:무료+텐바이텐 배송, WD:해외배송)

	dim FminPrice			'가격최소값
	dim FmaxPrice			'가격최대값
	dim FSalePercentHigh	'할인율 최대값
	dim FSalePercentLow		'할인율 최소값

	dim FCheckResearch 		'결과내 재검색 체크용
	dim FRectColsSize		'결과 리스트 열수
	dim FLogsAccept			'추가 로그 저장 여부

	dim FarrCate			'복수 카테고리
	dim FisTenOnly			'텐바이텐 전용상품
	dim FisLimit			'한정판매상품
	dim FisFreeBeasong
	
	dim FRectSearchGubun	'검색구분.상품,강좌,매거진
	
	dim FRectCPidx			'쿠폰idx
	dim FRectCateCodeLarge
	dim FRectCateCodeMid
	dim FRectClassCD			'클래스구분 원데이10 위클리20
	dim FRectClassPlaceCD	'장소구분 핑거스10 스튜디오20
	dim FRectLecCost			'수강료
	dim FRectLecTime			'시간대
	dim FRectCodeLarge		'대카테고리
	dim FRectCodeMid			'중카테고리

	Private SvrAddr
	Private SvrPort
	Private AuthCode
	Private Logs
	Private Scn
	private strQuery
	Private Order
	Private StartNum

	Private SearchQuery
	Private SortQuery
    
    public function InitDocruzer(iDocruzer)
        InitDocruzer = FALSE
        IF ( iDocruzer.BeginSession() < 0 ) THEN
			EXIT function
		End If
        
        IF NOT DocSetOption(iDocruzer) THEN
			EXIT function
		End If
		InitDocruzer = TRUE
    End function

    public function DocSetOption(iDocruzer)
        dim ret 
        ret = iDocruzer.SetOption(iDocruzer.OPTION_REQUEST_CHARSET_UTF8,1)
        DocSetOption = (ret>=0)
    end function
    
    

	''/검색 조건 설정
	FUNCTION getSearchQuery(byref query)
		dim strQue, arrCCD, arrSCD, arrACD, lp
		dim icurrTime : icurrTime=replace(LEFT(now(),10),"-","")&"000000"

		Select Case FListDiv
			Case "search"
				'검색 페이지 결과
				IF (FRectSearchTxt="" or isNull(FRectSearchTxt)) Then EXIT FUNCTION
			Case "list"
				'카테고리 목록
				'IF (FRectCateCode="" or isNull(FRectCateCode)) Then EXIT FUNCTION
			Case "fulllist"
				'카테고리없는 전체.
			Case Else
				EXIT FUNCTION
		End Select

		'### 검색조건 생성 ###

		'@ 검색어(키워드)
		IF FRectSearchTxt<>"" Then
			FRectSearchTxt = chgCoinedKeyword(FRectSearchTxt)
			FRectSearchTxt = escapeQuery(FRectSearchTxt)  ''2015 추가
			
			IF FRectExceptText<>"" Then
			    FRectExceptText = escapeQuery(FRectExceptText)  ''2015 추가
				strQue = getQrCon(strQue) & "(idx_keywords='" & FRectSearchTxt & " ! " & FRectExceptText & "' BOOLEAN) "	'제외어
			else
				strQue = getQrCon(strQue) & "idx_keywords='" & FRectSearchTxt & "'  allword "	'키워드검색(동의어 포함) synonym
				'strQue = getQrCon(strQue) & "idx_itemname='" & FRectSearchTxt & "'  natural "		'자연어 검색(동의어 포함) synonym
			End if
		END IF

		'@ 쿠폰번호
		IF FRectCPidx<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_currlecturercouponidx='" & FRectCPidx & "'"
		END IF

		'@ 카테고리
		IF FRectCateCodeLarge<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_newcatelarge='" & FRectCateCodeLarge & "'"
		END IF
		
		IF FRectCateCodeMid<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_newcatemid='" & FRectCateCodeMid & "'"
		END IF
		
		'@ 클래스구분
		IF FRectClassCD<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_catecd1='" & FRectClassCD & "'"
		END IF

		'@FSellScope ''reg_startday~reg_endday
		if FSellScope="Y" then
		    strQue = strQue & getQrCon(strQue) & "idx_reg_startday<='"&icurrTime&"'"
		    strQue = strQue & getQrCon(strQue) & "idx_reg_endday>='"&icurrTime&"'"
		    strQue = strQue & getQrCon(strQue) & "idx_sellyn='Y'"
		end if

		If FRectLecTime<>"" Then
			If FRectLecTime = "1" Then
				strQue = strQue & getQrCon(strQue) & "(idx_lec_startday1_time>'899' and idx_lec_startday1_time<'1100')"	'오전: 09:00~10:59
			ElseIf FRectLecTime = "2" Then
				strQue = strQue & getQrCon(strQue) & "(idx_lec_startday1_time>'1099' and idx_lec_startday1_time<'1400')"	'점심: 11:00~13:59
			ElseIf FRectLecTime = "3" Then
				strQue = strQue & getQrCon(strQue) & "(idx_lec_startday1_time>'1399' and idx_lec_startday1_time<'1800')"	'오후: 14:00~17:59
			ElseIf FRectLecTime = "4" Then
				strQue = strQue & getQrCon(strQue) & "(idx_lec_startday1_time>'1799' and idx_lec_startday1_time<'2400')"	'저녁: 18:00~23:59
			End If
		END IF
		
		'@ 수강료
		IF FRectLecCost<>"" Then
			If FRectLecCost = "1" Then
				strQue = strQue & getQrCon(strQue) & "(idx_lec_cost<'50001')"
			ElseIf FRectLecCost = "2" Then
				strQue = strQue & getQrCon(strQue) & "(idx_lec_cost<'100001')"
			ElseIf FRectLecCost = "3" Then
				strQue = strQue & getQrCon(strQue) & "(idx_lec_cost<'150001')"
			ElseIf FRectLecCost = "4" Then
				strQue = strQue & getQrCon(strQue) & "(idx_lec_cost<'200001')"
			End If
		END IF
		

		'@ 검색범위
		IF FRectSearchFlag<>"" THEN
			Select Case FRectSearchFlag
				Case "sc"	'세일쿠폰
					'strQue= strQue & getQrCon(strQue) & "(idx_saleyn='Y' or idx_itemcouponyn='Y') "
					strQue= strQue & getQrCon(strQue) & "(idx_itemcouponyn='Y') "
				Case "ea"	'전체사용기
					strQue= strQue & getQrCon(strQue) & "(idx_evalcnt>0) "
				'Case "ep"	'포토사용기
				'	strQue= strQue & getQrCon(strQue) & "(idx_evalcntPhoto>0) "
				Case "ne"	'신상품
					strQue = strQue & getQrCon(strQue) & "idx_newyn='Y' "
				'Case "fv"	'위시상품
				'	strQue = strQue & getQrCon(strQue) & "(idx_favcount>0) "
			End Select
		END IF
        
		query = strQue
	End FUNCTION

	Sub getSortQuery(byref query)
		dim strQue

		'// 정렬
		IF FRectSortMethod="ne" THEN '신상품
			strQue = strQue & " ORDER BY idx desc"
		ELSEIF FRectSortMethod="be" THEN '인기상품
			if FRectSearchFlag="fv" then
				'위시상품보기는 정렬을 위시순으로
				strQue = strQue & " ORDER BY favcount desc,brandrank asc,idx desc"
			else
				strQue = strQue & " ORDER BY limit_sold desc, brandrank asc,idx desc"
			end if
		ELSEIF FRectSortMethod="lp" THEN '낮은가격
			strQue = strQue & " ORDER BY lec_cost "
		ELSEIF FRectSortMethod="hp" THEN'높은가격
			strQue = strQue & " ORDER BY lec_cost desc"
		ELSEIF FRectSortMethod="br" THEN '베스트후기순
			strQue = strQue & " ORDER BY evalcnt desc,idx desc"
		ELSEIF FRectSortMethod="ws" THEN '위시순
			strQue = strQue & " ORDER BY favcount desc,idx desc"
		ELSEIF FRectSortMethod="mi" THEN '마감임박순
			strQue = strQue & " ORDER BY limit_per asc,idx desc"
		ELSE
			strQue = strQue & " ORDER BY idx desc"
		END IF
		
		query = strQue
	End Sub

	Function getQrCon(query)
		if Not(query="" or isNull(query)) then
			getQrCon = " and "
		end if
	End Function

	'// 상품 이미지 폴더 반환(컬러코드 유무에 따라 일반/컬러칩 구분)
	Function getItemImageUrl()
		IF application("Svr_Info")	= "Dev" THEN
				getItemImageUrl = "http://testimage.thefingers.co.kr"
		Else
				getItemImageUrl = "http://image.thefingers.co.kr"
		End If
	end function

	'####### 강좌 검색 - 검색 엔진 ######
	PUBLIC SUB getSearchList()

		DIM Scn
		DIM Docruzer,ret

		DIM Logs ,iRows
		DIM arrData,arrSize, retMatchCd, retMatchVal

		'// 검색 결과 출력 시나리오명
		Scn= "scn_dt_lecAcademy"		'일반 상품 검색

		StartNum = (FCurrPage -1)*FPageSize '// 검색시작 Row

		CALL getSearchQuery(SearchQuery)	'// 검색 쿼리생성
		CALL getSortQuery(SortQuery)		'// 정렬 쿼리 생성
		''Response.Write SearchQuery &"<Br>"
		'IF SearchQuery="" THEN
		'	EXIT SUB
		'END IF

        ''로그 않쌓음. 2016/08/25 eastone
		IF (FALSE) and (FLogsAccept) and (FRectSearchTxt<>"") and (FCurrPage="1") THEN
            'Logs = "상품+^" & FRectSearchTxt & "]##" & FRectSearchTxt & "||" & FRectPrevSearchTxt  	'// 로그값
            
            ''2015 search4
            '기본:[사이트@카테고리+사용자$성별코드|연령|검색어타입(서비스)|첫검색|페이지번호|정렬순^이전검색어##검색어] ''기본
            Dim iLOG_SITE : iLOG_SITE = "FINGERS"
            Dim iLOG_CATE : iLOG_CATE = "MOB"
            Dim iLOG_USER : iLOG_USER = GetUserLevelStr(GetLoginUserLevel) '' 회원등급을 사용
            Dim iLOG_SEX  : iLOG_SEX  = "" '' 0비로그인,1남성,2여성
            Dim iLOG_AGE  : iLOG_AGE  = "" '' 0비로그인,1:10대,2:20대,3:30대,4:40대,5:50대
            Dim iLOG_STYPE : iLOG_STYPE = "" '' 서비스 사용안함 X
            Dim iLOG_FIRST : iLOG_FIRST = "" '' 첫검색/재검색 사용안함 X  FCheckResearch
            
            Logs = iLOG_SITE&"@"                ''[ @
            Logs = Logs&iLOG_CATE&"+"           ''@ +
            Logs = Logs&iLOG_USER&"$"           ''+ $
            Logs = Logs&iLOG_SEX&"|"            ''$ |
            Logs = Logs&iLOG_AGE&"|"            ''| | 
            Logs = Logs&iLOG_STYPE&"|"          ''| | 
            Logs = Logs&iLOG_FIRST&"|"          ''| | 
            Logs = Logs&FCurrPage&"|"           ''| | 
            Logs = Logs&FRectSortMethod&"^"     ''| ^ 
            Logs = Logs&FRectPrevSearchTxt&"##" ''^ ##
            Logs = Logs&FRectSearchTxt          ''## ]
            
		END IF

		'##### 대상서버선택 '### 강좌검색인경우 2차서버
		SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR)

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF
		
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)
'response.write Docruzer.GetErrorMessage()
'response.end
		IF( ret < 0 ) THEN
		    rw "err1"
			dbACADEMYget.execute "EXECUTE db_academy.dbo.[sp_Academy_DocLog] @ErrMsg ='"& html2db(SearchQuery) & "[" & html2db(Docruzer.GetErrorMessage()) &"]'"

			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING

			IF FListDiv<>"search" THEN
				'// 1번 서버 에러시 2번에서 구동(2번도 에러면 Skip)
				if (SvrAddr = G_1STSCH_ADDR) then
					SvrAddr = G_2NDSCH_ADDR  ''"192.168.0.108"
					if (G_1STSCH_ADDR<>G_2NDSCH_ADDR) then  ''추가 2013/09
					    call getSearchList()
				    end if
				end if
			END IF

			EXIT SUB
		END IF

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수
		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		'Response.write "검색결과수 : " & FTotalCount & "<br>"
		IF( FResultCount <= 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB 'Response.write "GetResult_RowSize: " & Docruzer.GetErrorMessage()
		END IF

		FTotalPage =  Cdbl(FTotalCount\FPageSize)
		IF  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) THEN
			FtotalPage = FtotalPage +1
		END IF

		REDIM FItemList(FResultCount)

		FOR iRows=0 to FResultCount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.GetErrorMessage()
				EXIT FOR
			END IF

			SET FItemList(iRows) = NEW LectuerListItem

				FItemList(iRows).FLecIdx = arrData(0)					'강좌번호
				FItemList(iRows).FnewCateLarge = arrData(1)			'강좌 대카테고리
				FItemList(iRows).FnewCateMid = arrData(2)				'강좌 중카테고리
				FItemList(iRows).FClassCD = arrData(3)					'클래스 구분
				FItemList(iRows).FClassPlaceCD = arrData(4)			'강의장소 구분
				FItemList(iRows).FWeClassYN = arrData(5)				'위클래스
				FItemList(iRows).FLecStartWeek = arrData(6)			'강좌일 요일(숫자)
				FItemList(iRows).FlecStartDate = arrData(7)			'강좌 날짜(ex 20011010)
				FItemList(iRows).FCateSortNo = arrData(8)				'중카테고리 정렬번호
				FItemList(iRows).FLecTitle =db2html(arrData(9))		'강좌명
				FItemList(iRows).FLecCost = arrData(10)				'수강료
				FItemList(iRows).FMatCost = arrData(11)				'재료비
				FItemList(iRows).FLectCount = arrData(12)				'강좌 횟수
				FItemList(iRows).FLecperiod = db2html(arrData(13))	'강좌 기간
				FItemList(iRows).Flec_startday1 = arrData(14)			'강좌 시작 첫번째
				FItemList(iRows).Flimit_count = arrData(15)			'최소인원
				FItemList(iRows).FLecLimitCount = arrData(15)
				FItemList(iRows).Flimit_sold = arrData(16)			'판매량
				FItemList(iRows).FLecLimitSold = arrData(16)
				FItemList(iRows).Flecturer_id = arrData(17)			'강사 아이디
				FItemList(iRows).FMainimg = fingersImgUrl & "/lectureitem/main/" & GetImageSubFolderByItemid(FItemList(iRows).FLecIdx) & "/" & arrData(18)		'제품설명이미지 300x250
				FItemList(iRows).FStoryimg = fingersImgUrl & "/lectureitem/story1/" & GetImageSubFolderByItemid(FItemList(iRows).FLecIdx) & "/" & arrData(19)	'상품페이지직사각(3:2)이미지
				FItemList(iRows).FSmallimg = fingersImgUrl & "/lectureitem/small/" & GetImageSubFolderByItemid(FItemList(iRows).FLecIdx) & "/" & arrData(20)		'정사각 아이콘 50x50
				FItemList(iRows).FIconimg1 = fingersImgUrl & "/lectureitem/icon1/" & GetImageSubFolderByItemid(FItemList(iRows).FLecIdx) & "/" & arrData(21)		'정사각 아이콘 200x200
				FItemList(iRows).Foblong_img2 = fingersImgUrl & "/lectureitem/obl2/" & GetImageSubFolderByItemid(FItemList(iRows).FLecIdx) & "/" & arrData(22)	'직사각(3:2) 아이콘 240x160
				FItemList(iRows).Foblong_img3 = fingersImgUrl & "/lectureitem/obl3/" & GetImageSubFolderByItemid(FItemList(iRows).FLecIdx) & "/" & arrData(23)	'직사각(3:2) 아이콘 180x120
				FItemList(iRows).FMatincludeYN = arrData(24)			'재료비포함여부    ==> N(재료비 현장(별도)결제), Y(재료비 함께결제), X(재료비없음)
				FItemList(iRows).FReg_yn = arrData(25)					'등록가능 여부(Y등록 가능 N,등록종료)
				FItemList(iRows).FRegStartDate = arrData(26)			'접수시작일
				FItemList(iRows).FReg_startday = fnDateFormatOutPut(arrData(26))
				FItemList(iRows).FRegEndDate = arrData(27)			'접수종료일
				FItemList(iRows).FReg_endday = fnDateFormatOutPut(arrData(27))
				FItemList(iRows).Flecturer_regdate = fnDateFormatOutPut(arrData(28))		'강사 등록일
				FItemList(iRows).FValCount = arrData(29)				'해당강사 수강평가수
				FItemList(iRows).Flinit_per = arrData(30)
				FItemList(iRows).Fsellyn = arrData(31)					'사용여부 수정 관련
				FItemList(iRows).Flecturercouponyn = arrData(32)		'강좌 쿠폰 여부
				FItemList(iRows).Flecturercouponvalue = arrData(33)	'강좌 쿠폰 값
				FItemList(iRows).Flecturercoupontype = arrData(34)	'강좌 쿠폰 타입. 1:% 2:원
				FItemList(iRows).Fkeyword = db2html(arrData(35))		'키워드
				'FItemList(iRows). = arrData(36)						'강좌번호
				FItemList(iRows).Fnewyn = arrData(37)					'신규강좌 WHEN i.regdate>dateadd(day,-14,convert(varchar(10),getdate(),121)) THEN 'Y'
				FItemList(iRows).Fcatename = arrData(38)				'카테고리명 (ex 만지기^^구체관절/비스크돌)
				FItemList(iRows).Fisbestbrand = arrData(39)			'[db_academy].[dbo].[tbl_user_c_Academy].hitflg as isBestBrand
				FItemList(iRows).Fbestrank = arrData(40)				'[db_academy].[dbo].[tbl_user_c_Academy].hitRank as brandRank
				'FItemList(iRows). = arrData(41)						'replicate('0', 4 - LEN(convert(varchar(9),cm1.orderno))) + convert(varchar(9),cm1.orderno) as cd1sort
				'FItemList(iRows). = arrData(42)						'replicate('0', 4 - LEN(convert(varchar(9),cm2.orderno))) + convert(varchar(9),cm2.orderno) as cd2sort
				FItemList(iRows).FCate1and2 = arrData(43)				'카테고리 대 & 중. (ex 1020)
				FItemList(iRows).FLecturer_name = db2html(arrData(44))'강사 이름
				FItemList(iRows).Fsocname = arrData(45)				'[db_academy].[dbo].[tbl_user_c_Academy].socname
				FItemList(iRows).Fsocname_kor = arrData(46)			'[db_academy].[dbo].[tbl_user_c_Academy].socname_kor
				'FItemList(iRows).Fimage_profile = arrData(47)			'[db_academy].[dbo].[tbl_corner_good].image_profile
				If isNull(arrData(47)) OR arrData(47) = "" Then
					FItemList(iRows).Fimage_profile = ""
				Else
					FItemList(iRows).Fimage_profile = "http://" & CHKIIF(application("Svr_Info")="Dev","test","") & "image.thefingers.co.kr/corner/newImage_profile/thumbimg3/t3_" & arrData(47)
				End If
				FItemList(iRows).Fevalcnt = arrData(48)				'해당강좌 수강평가수
				FItemList(iRows).FLecStartday1time = arrData(49)		'시작 시간
				If isNull(arrData(50)) OR arrData(50) = "" Then
					FItemList(iRows).Fmorollingimg1 = fingersImgUrl & "/lectureitem/obl1/" & GetImageSubFolderByItemid(FItemList(iRows).FLecIdx) & "/" & arrData(54)	'리스트용이미지
				Else
					FItemList(iRows).Fmorollingimg1 = fingersImgUrl & "/lectureitem/morolling1/" & GetImageSubFolderByItemid(FItemList(iRows).FLecIdx) & "/" & arrData(50)	'리스트용이미지
				End If
				If Not isNull(arrData(51)) Then
					If arrData(51) = "0" Then
						FItemList(iRows).FRealFinishOX = "o"				'진짜종료인지
					Else
						FItemList(iRows).FRealFinishOX = "x"				'진짜종료인지
					End If
				End If
				FItemList(iRows).FRealKeyword = arrData(52)			'실제 저장된 키워드만 뽑음.
				FItemList(iRows).FoptionCnt = arrData(55)			'강좌 일정(옵션)수
				

			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

	'####### 상품 검색 카테고리별 카운팅  ######
	PUBLIC SUB getGroupbyCategoryList()

		'// 검색 결과 출력 시나리오명
		Scn= "scn_dt_lecAcademyCateGroup"		'일반 상품 검색

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// 검색시작 Row
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성
        
        IF SearchQuery="" Then
			EXIT SUB
		End If
		
		'//그룹 범위별 지정(정렬 쿼리 생성)
		Select Case FGroupScope
			Case "1"
				SortQuery = " GROUP BY idx_newcatelarge order by idx_newcatelarge " 
			Case "2"
				SortQuery = " GROUP BY idx_newcatemid order by idx_newcatemid "
			Case Else
				SortQuery = " GROUP BY idx_newcatelarge order by idx_newcatelarge "
		end Select

		dim Rowids,Scores
		FTotalCount = 0

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF
		
 'response.write "group : " & SearchQuery & SortQuery & "<br>"
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
		    response.write "ERR:"&Docruzer.GetErrorMessage()
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

'Response.write FResultcount & "!!"
		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)

		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new SearchGroupByItems
				'FItemList(iRows).FCateCode	= arrData(0)
				'FItemList(iRows).FCateName	= arrData(1)
				'FItemList(iRows).FCateCd1	= arrData(2)
				'FItemList(iRows).FCateCd2	= arrData(3)
				'FItemList(iRows).FCateCd3	= arrData(4)
				
				FItemList(iRows).FCateName		= arrData(0)
				FItemList(iRows).FnewCateLarge	= arrData(1)
				FItemList(iRows).FnewCateMid	= arrData(2)
				''2015 변경
				
				'if Len(FItemList(iRows).FCateCd1)>=4 then
				'    FItemList(iRows).FCateCd1	= Mid(FItemList(iRows).FCateCd1,5,255)              '' sort4+[code3]
				'end if
				    
				'if Len(FItemList(iRows).FCateCd2)>=11 then
				'    FItemList(iRows).FCateCd2	= Mid(FItemList(iRows).FCateCd2,9+3,255)            '' sort4+4+code3+[code3]
				'end if
				    
				'if Len(FItemList(iRows).FCateCd3)>=18 then
				'    FItemList(iRows).FCateCd3	= Mid(FItemList(iRows).FCateCd3,13+3+3,255)         '' sort4+4+4+code3+[code3]
				'end if
				
				'FItemList(iRows).FCateDepth	= arrData(5)
                
                ''rw FItemList(iRows).FCateCd1&"|"&FItemList(iRows).FCateCd2&"|"&FItemList(iRows).FCateCd3&"|"&FItemList(iRows).FCateDepth
                
				'FItemList(iRows).FImageSmall = getItemImageUrl & "/small/" & GetImageSubFolderByItemid(arrData(6)) & "/" &db2html( arrData(7))
				FItemList(iRows).FSubTotal 	= Scores(iRows)
				FTotalCount = FTotalCount + Scores(iRows)
			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT

		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

	
	'####### 상품 검색 총 카운팅  ######
	PUBLIC SUB getTotalCount()

		'// 검색 결과 출력 시나리오명
		Scn= "scn_dt_lecAcademy"		'일반 상품 검색
		
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// 검색시작 Row
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성

		IF SearchQuery="" Then
			EXIT SUB
		End If

		dim Rowids,Scores

        '// 기본으로 검색2번 섭, 검색어가 있다면 1번섭 사용
        ''---------------------------------------------------------------------------------------------------------
        if (FRectCateCode<>"") or (FRectMakerid<>"") and (FRectSearchTxt="")  then
            SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR) ''G_4THSCH_ADDR
        end if
        ''---------------------------------------------------------------------------------------------------------

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF
 
	'rw "group : " & SearchQuery & SortQuery & "<br>" ''
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _  
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)
		
						
		If( ret < 0 ) Then
		    rw "err2:"&Docruzer.GetErrorMessage
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)


		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB
    
    
	public function IsRegFinished()
		dim nowday, nextday , thisday, betweenday
		dim yyyy1, mm1, dd1, yyyy2 ,mm2, dd2

		''IsRegFinished
		''	E는 접수기간 종료,강제종료,
		''	B는 마감 임박
		''	W는 정원초과(대기자 신청가능)
		if FReg_yn="N" then
			IsRegFinished = "E"
			exit function
		end if

		nowday = now()
		yyyy1 = Cstr(Year(FRegEndDate))
		mm1 = Cstr(Month(FRegEndDate))
		dd1 = Cstr(day(FRegEndDate))

		yyyy2 = Cstr(Year(now()))
		mm2 = Cstr(Month(now()))
		dd2 = Cstr(day(now()))

		thisday = DateSerial(yyyy2, mm2, dd2)

		nextday = DateSerial(yyyy1, mm1 , dd1+ 1)

		betweenday = DateDiff("d",thisday,FRegEndDate)

		if (Flimit_count-Flimit_sold)<1 then
			IsRegFinished = "E"
			exit function
		elseif (Flimit_count-Flimit_sold)<= 5  then
			If FRealFinishOX = "o" Then
				IsRegFinished = "E"
				exit function
			Else
				IsRegFinished = "B"
			End If
		end if

		if (FRegStartDate>nowday) or (FRegEndDate<nowday) then
			IsRegFinished = "E"
			exit function
		elseif (betweenday <= 3) then
			IsRegFinished = "B"
		end if

	end function
	
	'// 강좌 쿠폰 여부
	public Function IsCouponlecturer()
			IsCouponlecturer = (FlecturerCouponYN="Y")
	end Function
	
	'신규강사 여부
	public Function isNewLecturer()
		if Datediff("m",FLecturer_Regdate,date())<1 then
			isNewLecturer = true
		else
			isNewLecturer = false
		end if
	end Function
    

	PUBLIC FUNCTION HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	END FUNCTION

	PUBLIC FUNCTION HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	END FUNCTION

	PUBLIC FUNCTION StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	END FUNCTION

End Class


'// 내검색어 저장처리 (쿠키, 최근 5개 저장→세션으로 변경;2014.07.29:허진원)
Sub procMySearchKeyword(kwd)
	Dim arrCKwd, rstKwd, i, excKwd
	'''arrCKwd = request.Cookies("search")("keyword")
	arrCKwd = session("myKeyword")
	arrCKwd = split(arrCKwd,",")
	''excKwd = "update,select,insert,and,union,from,alter,shutdown,kill,declare,exec,having,;,--"		'쿠키저장 제외 단어 (쿠키 인젝션)

	rstKwd = trim(kwd)
	if ubound(arrCKwd)>-1 then
		for i=0 to ubound(arrCKwd)
			if not(chkArrValue(excKwd,lcase(arrCKwd(i)))) then
				if arrCKwd(i)<>trim(kwd) then rstKwd = rstKwd & "," & arrCKwd(i)
			end if
			if i>4 then Exit For
		next
	end if

	'쿠키저장
	''response.Cookies("search").domain = "10x10.co.kr"
	''''response.cookies("search").Expires = Date + 3	'3일간 쿠키 저장 => 브라우저
	''response.Cookies("search")("keyword") = rstKwd
	session("myKeyword") = rstKwd
end Sub

'// 신조어/동의어 변환 처리 (신조어가 및 동의어가 안되는 문제 있을때 사용)
Function chgCoinedKeyword(kwd)
	dim arrChgTxt, arrItm
	arrChgTxt = split("반8||ban8",",")

	for each arrItm in arrChgTxt
		arrItm = split(arrItm,"||")
		if ubound(arrItm)>0 then
			kwd = Replace(kwd,arrItm(0),arrItm(1))
		end if
	next

	chgCoinedKeyword = kwd
end Function


'// 추가 카테고리 번호 추출 (추가카테고리에서 해당 카테고리 번호만 추출)
Function getArrayDispCate(vDisp,vArr)
	Dim vRst, i

	if vArr="" or isNull(vArr) or vDisp="" or isNull(vDisp) then Exit Function

	vArr = replace(trim(vArr)," ",",")
	vRst = split(vArr,",")

	if Not(isArray(vRst)) then Exit Function

	for i=0 to ubound(vRst)
		if inStr(vRst(i),vDisp)>0 then
			getArrayDispCate = vRst(i)
			Exit function
		end if
	next
end Function


'// 카테고리 Histoty 출력(2016 Ver.)
function fnPrnCategoryHistorymultiV16(vCode, vDiv, byRef vCateCnt, vCallBack)
	dim strHistory, strLink, SQL, i, j
	j = (len(vCode)/3)
    
	'히스토리 기본
	if vDiv="A" then
		strHistory = "<a href=""#"" onclick=""" & vCallBack & "(''); return false;"">전체</a>"
	end if

	i = 0
	'// 카테고리 이름 접수
	SQL = "SELECT ([db_academy].[dbo].[getCateCodeFullDepthName_Academy]('" & vCode & "'))"
	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open SQL, dbACADEMYget, adOpenForwardOnly, adLockReadOnly  '' 수정.2015/08/12

	if NOT(rsACADEMYget.EOF or rsACADEMYget.BOF) then
		If Not(isNull(rsACADEMYget(0))) Then

			for i = 1 to j
				If (GetPolderName(1) = "diyshop" OR GetPolderName(1) = "lecture") Then
					If i > 0 Then		'특정 폴더내 페이지의 카테고리 Depth 점프 단계 (이하는 표시안함)
						strHistory = strHistory & "<a href=""#"" onclick=""" & vCallBack & "('" &  Left(vCode,(3*i)) & "'); return false;"">"
						strHistory = strHistory & Split(db2html(rsACADEMYget(0)),"^^")(i-1)
						
						If i=j Then strHistory = strHistory & " ()" End If
		
						strHistory = strHistory & "</a>"
					End If
				Else
					strHistory = strHistory & "<a href=""#"" onclick=""" & vCallBack & "('" &  Left(vCode,(3*i)) & "'); return false;"">"
					strHistory = strHistory & Split(db2html(rsACADEMYget(0)),"^^")(i-1)
					
					If i=j Then strHistory = strHistory & " ()" End If
	
					strHistory = strHistory & "</a>"
				End If
			next
		End If
	end if
	
	rsACADEMYget.Close
	
	If (GetPolderName(1) = "diyshop" OR GetPolderName(1) = "lecture") Then
		vCateCnt = i - 2
	Else
		vCateCnt = i
	End If
	
	fnPrnCategoryHistorymultiV16 = strHistory
end function


'// Lec 카테고리 Histoty 출력
function fnPrnLecCategoryHistorymulti(vCodelarge, vCodemid, vDiv, byRef vCateCnt)
	dim strHistory, strLink, SQL, i, j
	
	if vCodemid <> "" Then
		j = 2
	else
		j = 1
	end if

	'히스토리 기본
	if vDiv="A" then
		strHistory = "<a href=""#"" onclick=""goLecCateList('',''); return false;"">전체</a>"
	end if

	i = 0
	'// 카테고리 이름 접수
	SQL = "SELECT ([db_academy].[dbo].[getLecCateCodeFullDepthName_Academy]('" & vCodelarge & "','" & vCodemid & "'))"
	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open SQL, dbACADEMYget, adOpenForwardOnly, adLockReadOnly  '' 수정.2015/08/12

	if NOT(rsACADEMYget.EOF or rsACADEMYget.BOF) then
		If Not(isNull(rsACADEMYget(0))) Then

			for i = 1 to j
				strHistory = strHistory & "<a href=""#"" onclick=""goLecCateList('" & vCodelarge & "','" & CHKIIF(i=1,"",vCodemid) & "'); return false;"">"
				strHistory = strHistory & Split(db2html(rsACADEMYget(0)),"^^")(i-1)
				
				If i=j Then strHistory = strHistory & " ()" End If

				strHistory = strHistory & "</a>"
			next
		End If
	end if
	
	rsACADEMYget.Close
	vCateCnt = i
	
	fnPrnLecCategoryHistorymulti = strHistory
end function


'// Magazine 카테고리 Histoty 출력
function fnPrnMagaCategoryHistorymulti(vCode, vDiv, byRef vCateCnt)
	dim strHistory, strLink, SQL, i, j

	'히스토리 기본
	if vDiv="A" then
		strHistory = "<a href=""#"" onclick=""goMagaCateList(''); return false;"">전체</a>"
	end if

	i = 0
	'// 카테고리 이름 접수
	SQL = "SELECT ([db_academy].[dbo].[getMagaCateCodeFullDepthName_Academy]('" & vCode & "'))"
	rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open SQL, dbACADEMYget, adOpenForwardOnly, adLockReadOnly  '' 수정.2015/08/12

	if NOT(rsACADEMYget.EOF or rsACADEMYget.BOF) then
		If Not(isNull(rsACADEMYget(0))) Then
			strHistory = strHistory & "<a href=""#"" onclick=""goMagaCateList('" & vCode & "'); return false;"">" & db2html(rsACADEMYget(0)) & " ()" & "</a>"
			i = 1
		End If
	end if
	
	rsACADEMYget.Close
	vCateCnt = i
	
	fnPrnMagaCategoryHistorymulti = strHistory
end function


function fnKeywordFormatChange(v,db)
	Dim vTemp, i, vDB
	If db = "lec" Then
		vDB = "&searchgb=lecAcademy"
	ElseIf db = "mag" Then
		vDB = "&searchgb=magazine"
	Else
		vDB = "&searchgb=itemAcademy"
	End If
	
	For i = LBound(Split(v,",")) To UBound(Split(v,","))
		vTemp = vTemp & "<a href=""/search/search_result.asp?rect=" & Trim(Split(v,",")(i)) & vDB & """>#" & Trim(Split(v,",")(i)) & "</a>"
	Next
	
	fnKeywordFormatChange = vTemp
end function


Function fnMyWishCheck(arr, itemid)
	Dim i, torf
	torf = False
	IF isArray(arr) THEN
		For i =0 To UBound(arr,2)
			If CStr(arr(0,i)) =  CStr(itemid) Then
				torf = True
				Exit For
			End If
		Next
	End If
	fnMyWishCheck = torf
End Function


Function fnDateFormatOutPut(d)
	Dim tmp
	If d <> "" Then
		If Len(d) = 14 Then
			tmp = Left(d,4) & "-" & Mid(d,5,2) & "-" & Mid(d,7,2)
			tmp = CDate(tmp)
		End If
	End If
	fnDateFormatOutPut = tmp
End Function
%>
