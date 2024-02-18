<%
''CS_EUCKR <=> CS_UTF8
''group by 필드는 널이 없게 하자.
'''--------------------------------------------------------------------------------------
DIM G_KSCOLORCD : G_KSCOLORCD = Array("023","001","002","010","021","003","004","024","019","005","016","006","007","020","008","018","017","009","011","012","022","013","014","015","025","026","027","028","029","030","031")
DIM G_KSCOLORNM : G_KSCOLORNM = Array("와인","빨강","주황","갈색","카멜","노랑","베이지","아이보리","카키","초록","민트","연파랑","파랑","네이비","보라","연보라","베이비핑크","핑크","흰색","옅은회색","짙은회색","검정","은색","금색","체크","스트라이프","도트","플라워","드로잉","애니멀","기하학")

Dim G_KSSTYLECD : G_KSSTYLECD = Array("010","020","030","040","050","060","070","080","090")
Dim G_KSSTYLENM : G_KSSTYLENM = Array("클래식","큐티","댄디","모던","내추럴","오리엔탈","팝","로맨틱","빈티지")

DIM G_ORGSCH_ADDR
DIM G_1STSCH_ADDR
DIM G_2NDSCH_ADDR
DIM G_3RDSCH_ADDR
Dim G_4THSCH_ADDR

DIM G_SCH_TIME : G_SCH_TIME=formatdatetime(now(),4)

IF (application("Svr_Info") = "Dev") THEN
     G_1STSCH_ADDR = "192.168.50.10"  ''"110.93.128.109" ''
     G_2NDSCH_ADDR = "192.168.50.10"
     G_3RDSCH_ADDR = "192.168.50.10"
     G_4THSCH_ADDR = "192.168.50.10"
     G_ORGSCH_ADDR = "192.168.50.10"
ELSE
     G_1STSCH_ADDR = "192.168.0.210"        ''192.168.0.210  :: 검색페이지(search.asp)   '
     G_2NDSCH_ADDR = "192.168.0.207"        ''192.168.0.207  :: 카테고리, 상품, 브랜드   ''문제 발생시 208 로.
     G_3RDSCH_ADDR = "192.168.0.209"        ''192.168.0.209  :: GiftPlus , scn_dt_itemDispColor :: 확인.
     G_4THSCH_ADDR = "192.168.0.208"        ''192.168.0.208  :: mobile 6:10분에 인덱싱 시작 카피
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
		case "("
			ret = ret & "\("
		case ")"
			ret = ret & "\)"
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
  exit function
    IF Not (application("Svr_Info")="Dev") THEN
        exit function
    ENd IF

    dim itime
    Call iDocruzer.GetResult_SearchTime(itime) '소요시간
    rw "-------------------------------"
    rw Scn
    rw iSearchQuery
    rw iSortQuery
    rw "FTotalCount:"&iFTotalCount
    rw "FResultcount:"&iFResultcount
    rw "GetResult_SearchTime:"&itime
end function
'''--------------------------------------------------------------------------------------
''2015
'' StringList 검색 결과중  컬러검새 관련
function getDocArrMatchCodeVal(iRectArr,iResultCdArr,iResultValArr,byref retMatchCd,byref retMatchVal)
    dim findvalArr : findvalArr = split(trim(iRectArr),",")
    dim rsltCd : rsltCd = split(iResultCdArr," ")
    dim rsltVal : rsltVal = split(iResultValArr," ")
    dim findval, i,j

    for i=LBound(findvalArr) to Ubound(findvalArr)
        findval = findvalArr(i)
        for j=LBound(rsltCd) to Ubound(rsltCd)
            if (rsltCd(j)=findval) then
                retMatchCd  = rsltCd(j)
                retMatchVal = rsltVal(j)
                exit for
            end if
        next
    next
end function

function getCdPosVal(iCD,icdArr,iValArr)
    dim i,iPos
    getCdPosVal = ""
    if Not isArray(icdArr) then Exit function
    if Not isArray(iValArr) then Exit function
    if UBound(icdArr)<>UBound(iValArr) then Exit function

    iPos = -1
    for i=LBound(icdArr) to UBound(icdArr)
        if (iCD=icdArr(i)) then
            iPos = i
            Exit For
        end if
    next
    if iPos<0 then Exit function
    getCdPosVal=iValArr(iPos)
end function


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

End Class

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

	dim FRectItemid			'상품코드	'진영2017-03-28
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

    public FRectSynonymAssign
	public FRectExpandSearch
	public FRectIdxrect

	public function getRetSearchQuery
		getRetSearchQuery = SearchQuery
	end function

	public function getRetSortQuery
		getRetSortQuery = SortQuery
	end function

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

	public function GetLevelUpCount()

		if (FCurrRank<FLastRank) then
			GetLevelUpCount = CStr(FLastRank-FCurrRank)
		elseif (FCurrRank=FLastRank) and (FLastRank=0) then
			GetLevelUpCount = ""
		elseif (FCurrRank=FLastRank) then
			GetLevelUpCount = ""
		elseif (FCurrRank>FLastRank) and (FLastRank=0) then
			GetLevelUpCount = ""
		else
			GetLevelUpCount = CStr(FCurrRank-FLastRank)
			if FCurrRank-FLastRank>=FCurrPos then
				GetLevelUpCount = ""
			end if
		end if
	end function

	public function GetLevelUpArrow()
		if (FCurrRank<FLastRank) then
			GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2013/award/ico_rank_up.gif' alt='순위 상승' /> " & GetLevelUpCount()
		elseif (FCurrRank=FLastRank) and (FLastRank=0) then
			GetLevelUpArrow = ""
		elseif (FCurrRank=FLastRank) then
			'GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2010/bestaward/ico_none.gif' align='absmiddle' style='display:inline;'> <font class='eng11px00'><b>0</b></font>"
			GetLevelUpArrow = ""
		elseif (FCurrRank>FLastRank) and (FLastRank=0) then
			GetLevelUpArrow = ""
		else
			GetLevelUpArrow = "<img src='http://fiximage.10x10.co.kr/web2013/award/ico_rank_down.gif' alt='순위 하락' /> " & GetLevelUpCount()
			if FCurrRank-FLastRank>=FCurrPos then
				'GetLevelUpArrow = "<font class='eng11px00'><b>0</b></font>"
				GetLevelUpArrow = ""
			end if
		end if
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
			Case "colorlist"
				'컬러 검색 결과
				if (FcolorCode="" and FcolorCode="0") Then EXIT FUNCTION
			Case "brand"
				'브랜드 상품 목록
				IF (FRectMakerid="" or isNull(FRectMakerid)) Then EXIT FUNCTION
				FRectSearchItemDiv = "y"
			Case "salelist"
				'할인상품 목록
				FRectSearchFlag = "sc"
			Case "newlist"
				'신상품 목록
				FRectSearchFlag = "ne"
				IF FRectCateCode="" Then FRectSearchItemDiv = "y"
			Case "bestlist"
				'베스트상품 목록
				FRectSearchItemDiv = "y"
			Case "aboard"
				'해외판매 상품 목록
				FdeliType = "WD"
			Case "fulllist"
				'카테고리없는 전체.
			Case Else
				EXIT FUNCTION
		End Select

		'### 검색조건 생성 ###

		if (FRectIdxrect="") then FRectIdxrect="idx_itemname" ''신규방식

		'@ 검색어(키워드)
		If FRectSearchTxt<>"" Then
			FRectSearchTxt = chgCoinedKeyword(FRectSearchTxt)
			FRectSearchTxt = escapeQuery(FRectSearchTxt)  ''2015 추가
			IF FRectExceptText<>"" Then
			    FRectExceptText = escapeQuery(FRectExceptText)  ''2015 추가
				strQue = getQrCon(strQue) & "("&FRectIdxrect&"='" & FRectSearchTxt & " ! " & FRectExceptText & "' BOOLEAN) "	'제외어
			else
				if (FRectExpandSearch="allwordOradjacent") then
					strQue = " aliasing "&FRectIdxrect&" as name1, "&FRectIdxrect&" as name2 (name1='"&FRectSearchTxt&"' alladjacent "&CHKIIF(FRectSynonymAssign<>"","synonym","")&" or name2='"&FRectSearchTxt&"' allword "&CHKIIF(FRectSynonymAssign<>"","synonym","")&") "
				else
					strQue = getQrCon(strQue) & ""&FRectIdxrect&"='" & FRectSearchTxt & "'  "	
					if (FRectExpandSearch<>"") then
						strQue = strQue & " "&FRectExpandSearch&" "
					end if

					IF (FRectExpandSearch<>"") and (FRectSynonymAssign<>"") then
						strQue = strQue & " synonym "	'키워드검색(동의어 포함 => synonym )
					end if
					'strQue = getQrCon(strQue) & "idx_itemname='" & FRectSearchTxt & "'  natural "		'자연어 검색(동의어 포함) synonym
				end if
			End if
		End If
'		If FRectSearch <> "" AND FRectSearchTxt<>"" Then
'			If (FRectSearch = "keyword") OR (FRectSearch = "itemname") Then
'				FRectSearchTxt = chgCoinedKeyword(FRectSearchTxt)
'				FRectSearchTxt = escapeQuery(FRectSearchTxt)  ''2015 추가
'				IF FRectExceptText<>"" Then
'				    FRectExceptText = escapeQuery(FRectExceptText)  ''2015 추가
'					strQue = getQrCon(strQue) & "(idx_itemname='" & FRectSearchTxt & " ! " & FRectExceptText & "' BOOLEAN) "	'제외어
'				else
'					strQue = getQrCon(strQue) & "idx_itemname='" & FRectSearchTxt & "'  allword "	'키워드검색(동의어 포함) synonym
'					'strQue = getQrCon(strQue) & "idx_itemname='" & FRectSearchTxt & "'  natural "		'자연어 검색(동의어 포함) synonym
'				End if
'			ElseIf FRectSearch = "itemid" Then
'				strQue = getQrCon(strQue) & "idx_itemid=" & FRectSearchTxt
'				'strQue = getQrCon(strQue) & "idx_itemid=279397 or idx_itemid=145983 "
'			End If
'		End If

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

		'@ 컬러검색
		if Not(FcolorCode="" or isNull(FcolorCode) or FcolorCode="0") then
			arrCCD = split(FcolorCode,",")
			if ubound(arrCCD)>0 then
				'다중 컬러코드 적용
				strQue = strQue & getQrCon(strQue) & "(idx_colorCd='"&replace(FcolorCode,","," ")&"' anystring)"  ''2015 변경
'				strQue = strQue & getQrCon(strQue) & "("
'				for lp=0 to ubound(arrCCD)
'					if getNumeric(arrCCD(lp))<>"" then
'						if lp>0 then strQue = strQue & " or "
'						strQue = strQue & "idx_colorCd='" & getNumeric(arrCCD(lp)) & "'"
'					end if
'				next
'				strQue = strQue & ")"
			else
				'단일 컬러코드 적용
				strQue = strQue & getQrCon(strQue) & "idx_colorCd='" & getNumeric(arrCCD(0)) & "'"
			end if
		end if

		'@ 스타일 조건 쿼리
		if Not(FstyleCd="" or isNull(FstyleCd)) then
			arrSCD = split(FstyleCd,",")
			if ubound(arrSCD)>0 then
				'다중 스타일코드 적용
				strQue = strQue & getQrCon(strQue) & "(idx_styleCd='"&replace(FstyleCd,","," ")&"' anystring)"  ''2015 변경
'				strQue = strQue & getQrCon(strQue) & "("
'				for lp=0 to ubound(arrSCD)
'					if getNumeric(arrSCD(lp))<>"" then
'						if lp>0 then strQue = strQue & " or "
'						strQue = strQue & "idx_styleCd='" & getNumeric(arrSCD(lp)) & "'"
'					end if
'				next
'				strQue = strQue & ")"
			else
				'단일 스타일코드 적용
				strQue = strQue & getQrCon(strQue) & "idx_styleCd='" & getNumeric(arrSCD(0)) & "'"
			end if
		end if

		'@ 상품속성 조건 쿼리
		if Not(FattribCd="" or isNull(FattribCd)) then
			arrACD = split(FattribCd,",")
			if ubound(arrACD)>0 then
				'다중 속성코드 적용
				strQue = strQue & getQrCon(strQue) & "(idx_attribCd='"&replace(FattribCd,","," ")&"' anystring)"  ''2015 변경
'				strQue = strQue & getQrCon(strQue) & "("
'				for lp=0 to ubound(arrACD)
'					if getNumeric(arrACD(lp))<>"" then
'						if lp>0 then strQue = strQue & " or "
'						strQue = strQue & "idx_attribCd='" & getNumeric(arrACD(lp)) & "'"
'					end if
'				next
'				strQue = strQue & ")"
			else
				'단일 속성코드 적용
				strQue = strQue & getQrCon(strQue) & "idx_attribCd='" & getNumeric(arrACD(0)) & "'"
			end if
		end if

		'@ 상품 판매 범위
'		IF FSellScope="Y" Then
'			strQue = strQue & getQrCon(strQue) & "idx_sellyn='Y' "
'		ELSE
'			strQue = strQue & getQrCon(strQue) & "(idx_sellyn='Y' or idx_sellyn='S') "
'		End IF

		If FSellScope <> "" Then
			strQue = strQue & getQrCon(strQue) & "idx_sellyn='"& FSellScope &"' "
		Else
			strQue = strQue & getQrCon(strQue) & "(idx_sellyn='Y' or idx_sellyn='S' or idx_sellyn='N')  "
		End If

        ''2015 추가 string list group by 사용시 해당필드에 널값이 있는경우 빈값 대신 000 넣음.
        IF scn="scn_dt_itemDispStyleGroup" then
            strQue = strQue & getQrCon(strQue) & "idx_styleCd!='000' "
        ELSEIF  scn="scn_dt_itemDispAttribGroup" then
            strQue = strQue & getQrCon(strQue) & "idx_attribgrp!='000' "
        ELSEIF  scn="scn_dt_itemDispColorGroup" then
            strQue = strQue & getQrCon(strQue) & "idx_colorgrp!='000' "
        ELSEIF  scn="scn_dt_itemDispCateGroup" then
            if (FGroupScope="2") then
                strQue = strQue & getQrCon(strQue) & "idx_cd2grp!='000' "
            elseif (FGroupScope="3") then
                strQue = strQue & getQrCon(strQue) & "idx_cd3grp!='000' "
            end if
        END IF

		' 제외검색어 추가 : 조건절 가장 뒤에 위치해있어야한다. 이 뒤로 뭐 추가하지 말고 앞으로 추가할것
		IF FRectSearchTxt<>"" and (scn="scn_dt_itemDisp" or scn="scn_dt_itemDispCateGroup"  or scn="scn_dt_itemDispBrandGroup"  or scn="scn_dt_itemDispColorGroup"  or scn="scn_dt_itemDispStyleGroup" or scn="scn_dt_itemDispAttribGroup"  or scn="") then 
			strQue = strQue & "exclude by keylist(0) "
		end if

		query = strQue
	
'rw strQue
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
				strQue = strQue & " ORDER BY $MATCHFIELD(cateboostkeylist,bestkeylist) desc,itemscore desc,itemid desc"
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
		ELSEIF FRectSortMethod="bs0" THEN '판매수순(기존)
		    strQue = strQue & " ORDER BY sellCnt desc,sellcash desc,itemid desc"
		ELSEIF FRectSortMethod="bs1" THEN '판매수순(검색어)
			strQue = strQue & " ORDER BY sellCnt desc,$MATCHFIELD(cateboostkeylist,bestkeylist) desc, itemid desc"
		ELSEIF FRectSortMethod="bs2" THEN '판매수순(검색어)
			strQue = strQue & " ORDER BY sellCnt desc,$MATCHFIELD(cateboostkeylist) desc, itemid desc"
		ELSEIF FRectSortMethod="bs3" THEN 'TEST
			strQue = strQue & " ORDER BY sellCnt desc,$MATCHFIELD(cateboostkeylist,bestkeylist) desc,sellcash desc, itemid desc"
		ELSEIF FRectSortMethod="bs4" THEN 'TEST
			strQue = strQue & " ORDER BY $MATCHFIELD(cateboostkeylist,bestkeylist) desc,sellCnt desc,itemscore desc, itemid desc"
		ELSEIF FRectSortMethod="bs5" THEN 'TEST
			strQue = strQue & " ORDER BY $MATCHFIELD(cateboostkeylist) desc,sellCnt desc,sellcash desc, itemid desc"
		ELSEIF FRectSortMethod="bs6" THEN 'TEST
			strQue = strQue & " ORDER BY $CATEGORYFIELD( recomkeyword(1) seasonboost_groupid(3) categorynamelist(0) bestkeylist(2) makerid(4), '"&replace(TRIM(FRectSearchTxt)," ","")&"') desc, sellCnt desc, itemscore desc, itemid desc"
		ELSEIF FRectSortMethod="be6" THEN 'TEST
			strQue = strQue & " ORDER BY $CATEGORYFIELD( recomkeyword(1) seasonboost_groupid(3) categorynamelist(0) bestkeylist(2) makerid(4) , '"&replace(TRIM(FRectSearchTxt)," ","")&"') desc, itemscore desc, itemid desc"
		ELSEIF FRectSortMethod="bs7" THEN 'TEST
			strQue = strQue & " ORDER BY $CATEGORYFIELD( recomkeyword(1) seasonboost_groupid(3) categorynamelist(0) bestkeylist(2) makerid(4) , '"&replace(TRIM(FRectSearchTxt)," ","")&"') desc,$MATCHFIELD(name1,name2) desc, sellCnt desc, itemscore desc, itemid desc"
		ELSEIF FRectSortMethod="be7" THEN 'TEST
			strQue = strQue & " ORDER BY $CATEGORYFIELD( recomkeyword(1) seasonboost_groupid(3) categorynamelist(0) bestkeylist(2) makerid(4) , '"&replace(TRIM(FRectSearchTxt)," ","")&"') desc,$MATCHFIELD(name1,name2) desc, itemscore desc, itemid desc"
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
			if FcolorCode="" or FcolorCode="0" then
				getItemImageUrl = "http://webimage.10x10.co.kr/image"
			else
				getItemImageUrl = "http://webimage.10x10.co.kr/color"
			end if
		Else
			if FcolorCode="" or FcolorCode="0" then
				getItemImageUrl = "http://webimage.10x10.co.kr/image"
			else
				getItemImageUrl = "http://webimage.10x10.co.kr/color"
			end if
		End If
	end function

	'####### 상품 검색 - 검색 엔진 ######
	PUBLIC SUB getSearchList()

		DIM Scn
		DIM Docruzer,ret

		DIM Logs ,iRows
		DIM arrData,arrSize, retMatchCd, retMatchVal

        if (FPageSize>300) then FPageSize=300  ''2016/11/16추가

		'// 검색 결과 출력 시나리오명
		if FcolorCode="" or FcolorCode="0" then
			Scn= "scn_dt_itemDisp"		'일반 상품 검색
		else
			'Scn= "scn_dt_itemColor"		'상품 컬러별 검색
			'Scn= "scn_dt_itemDispColor"	'상품 컬러별 검색(전시카테고리)
			Scn= "scn_dt_itemDisp"		    '일반 상품 검색 통일 2015
		end if

		StartNum = (FCurrPage -1)*FPageSize '// 검색시작 Row

		CALL getSearchQuery(SearchQuery)	'// 검색 쿼리생성
		CALL getSortQuery(SortQuery)		'// 정렬 쿼리 생성
		''Response.Write SearchQuery &"<Br>"
		IF SearchQuery="" THEN
			EXIT SUB
		END IF

		IF (FLogsAccept) and (FRectSearchTxt<>"") and (FCurrPage="1") THEN

            'Logs = "상품+^" & FRectSearchTxt & "]##" & FRectSearchTxt & "||" & FRectPrevSearchTxt  	'// 로그값

            ''2015 search4
            '기본:[사이트@카테고리+사용자$성별코드|연령|검색어타입(서비스)|첫검색|페이지번호|정렬순^이전검색어##검색어] ''기본
            Dim iLOG_SITE : iLOG_SITE = "WEB"
            Dim iLOG_CATE : iLOG_CATE = "RECT"
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

        ''상품검색/브랜드검색이 아닌경우 2차서버로.
        ''---------------------------------------------------------------------------------------------------------
        if (FRectCateCode<>"") or (FRectMakerid<>"") and (FRectSearchTxt="")  then
            'response.write "2차서버<br>"
             SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR) ''G_4THSCH_ADDR ''G_2NDSCH_ADDR
        end if
        ''칼라칩검색/브랜드검색 3차
        if (Scn= "scn_dt_itemDispColor") or (FRectMakerid<>"") then
        	'response.write "3차서버<br>"
        	SvrAddr = getTimeChkAddr(G_3RDSCH_ADDR)
        end if
        ''---------------------------------------------------------------------------------------------------------

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

		IF( ret < 0 ) THEN
		    rw "err1"
			dbget.execute "EXECUTE db_log.dbo.sp_Ten_DocLog @ErrMsg ='"& html2db(SearchQuery) & "[" & html2db(Docruzer.GetErrorMessage()) &"]'"

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
				if (FcolorCode="" or FcolorCode="0") then
				    FItemList(iRows).FImageBasic 	= getItemImageUrl & "/basic/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(11))
    				FItemList(iRows).FImageMask 	= getItemImageUrl & "/mask/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(12))
    				FItemList(iRows).FImageList 	= getItemImageUrl & "/list/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(13))
    				FItemList(iRows).FImageList120 	= getItemImageUrl & "/list120/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(14))
    				FItemList(iRows).FImageIcon1 	= getItemImageUrl & "/icon1/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(15))
    				FItemList(iRows).FImageIcon2 	= getItemImageUrl & "/icon2/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(16))
    				FItemList(iRows).FImageSmall	= getItemImageUrl & "/small/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(17))
				else
				    retMatchCd  = ""
				    retMatchVal = ""

				    Call getDocArrMatchCodeVal(FcolorCode,arrData(35),arrData(44),retMatchCd,retMatchVal)

				    FItemList(iRows).FImageBasic 	= getItemImageUrl & "/basic/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &retMatchVal
    				FItemList(iRows).FImageMask 	= FItemList(iRows).FImageBasic
    				FItemList(iRows).FImageList 	= getThumbImgFromURL(FItemList(iRows).FImageBasic,100,100,"true","false") 'getStonReSizeImg(FItemList(iRows).FImageBasic,100,100,85)
    				FItemList(iRows).FImageList120 	= getThumbImgFromURL(FItemList(iRows).FImageBasic,120,120,"true","false") 'getStonReSizeImg(FItemList(iRows).FImageBasic,120,120,85)
    				FItemList(iRows).FImageIcon1 	= getThumbImgFromURL(FItemList(iRows).FImageBasic,200,200,"true","false") 'getStonReSizeImg(FItemList(iRows).FImageBasic,200,200,85)
    				FItemList(iRows).FImageIcon2 	= getThumbImgFromURL(FItemList(iRows).FImageBasic,150,150,"true","false") 'getStonReSizeImg(FItemList(iRows).FImageBasic,150,150,85)
    				FItemList(iRows).FImageSmall	= getThumbImgFromURL(FItemList(iRows).FImageBasic,50,50,"true","false") 'getStonReSizeImg(FItemList(iRows).FImageBasic,50,50,85)
    			end if
				if (arrData(47)<>"") then ''2015 추가 (추가 이미지)
				    if (FcolorCode="" or FcolorCode="0") then
				    	FItemList(iRows).FAddimage      = getItemImageUrl & "/add1/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" & db2html(arrData(47))
				    else
				    	FItemList(iRows).FAddimage      = replace(getItemImageUrl,"/color","/image") & "/add1/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" & db2html(arrData(47))
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
                FItemList(iRows).FAllCateName = arrData(41)		''//2015.10.07

                'FItemList(iRows).FcolorCd = arrData(35)
			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

	'####### 상품 검색 카테고리별 카운팅  ######
	PUBLIC SUB getGroupbyCategoryList()

		'// 검색 결과 출력 시나리오명
		Scn= "scn_dt_itemDispCateGroup"		'일반 상품 검색

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
				SortQuery = " GROUP BY cd1grp order by cd1grp "
			Case "2"
				SortQuery = " GROUP BY cd2grp order by cd2grp "
			Case "3"
				SortQuery = " GROUP BY cd3grp order by cd3grp "
			Case Else
				SortQuery = " GROUP BY cd1grp order by cd1grp "
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

	'####### 상품 검색 브랜드별 카운팅  ######
	PUBLIC SUB getGroupbyBrandList()

		'// 검색 결과 출력 시나리오명
		''Scn= "scn_dt_itemDisp"		'일반 상품 검색

		Scn= "scn_dt_itemDispBrandGroup"

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// 검색시작 Row
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성

		'//그룹 범위별 지정(정렬 쿼리 생성)
		'SortQuery = " GROUP BY makerid order by brandname "
		SortQuery = " GROUP BY makerid order by count(*) desc" ''desc $RELEVANCE

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
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수

		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new CCategoryPrdItem
				FItemList(iRows).FMakerID		= arrData(0)
				FItemList(iRows).FBrandName		= arrData(1)
				FItemList(iRows).FImageSmall	= getItemImageUrl & "/small/" & GetImageSubFolderByItemid(arrData(2)) & "/" &db2html( arrData(3))
				FItemList(iRows).FisBestBrand	= arrData(4)
				FItemList(iRows).FItemScore 	= Scores(iRows)
				FItemList(iRows).FCurrRank		= arrData(5)

			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT

		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB


	'####### 상품 스타일별 카운팅  ######
	PUBLIC SUB getGroupbyStyleList()
		'// 검색 결과 출력 시나리오명
		Scn= "scn_dt_itemDispStyleGroup"		'상품스타일 검색

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim groupKeyCnt,groupKeyVal,groupsize

		dim FseekTime

		StartNum = 0 						'// 검색시작 Row
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성

		'//그룹 범위 지정(정렬 쿼리 생성)
		SortQuery = " GROUP BY styleCd order by styleCd "

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

		ret = Docruzer.Search(SvrAddr&":"&SvrPort, _
						Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,"",StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
            'response.write Docruzer.GetErrorMessage()
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		ret = Docruzer.GetResult_GroupBy(FResultcount,groupKeyCnt,groupKeyVal,groupsize,100)

		IF( ret < 0 ) THEN
			'Response.write "GetResult_Row: " & Docruzer.msg
			SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
			EXIT Sub
		END IF

		if (groupKeyCnt<1) then
		    SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
		    Exit Sub
		end if

CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)
        dim bufStyle

		FOR iRows = 0 to FResultcount -1
		    bufStyle = groupKeyVal(iRows,0)
			SET FItemList(iRows) = new SearchGroupByItems
			FItemList(iRows).FStyleCd		= bufStyle
	        'FItemList(iRows).FStyleName		= buf

			FItemList(iRows).FSubTotal 	= groupsize(iRows)
			FTotalCount = FTotalCount + groupsize(iRows)
		NEXT

		SET groupKeyVal = NOTHING
        SET groupsize   = NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB



    '####### 상품 속성별 카운팅  ######
	PUBLIC SUB getGroupbyAttribList()

		'// 검색 결과 출력 시나리오명
		Scn= "scn_dt_itemDispAttribGroup"		'상품속성 검색

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim groupKeyCnt,groupKeyVal,groupsize

		dim FseekTime

		StartNum = 0 						'// 검색시작 Row
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성

		'//그룹 범위 지정(정렬 쿼리 생성)
		SortQuery = " GROUP BY attribgrp order by attribgrp "

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

		''response.write "group : " & SearchQuery & SortQuery & "<br>"
		ret = Docruzer.Search(SvrAddr&":"&SvrPort, _
						Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,"",StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
		    response.write "ERR:"&Docruzer.GetErrorMessage()
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

        ret = Docruzer.GetResult_GroupBy(FResultcount,groupKeyCnt,groupKeyVal,groupsize,100)

        IF( ret < 0 ) THEN
			'Response.write "GetResult_Row: " & Docruzer.msg
			SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
			EXIT Sub
		END IF

		if (groupKeyCnt<1) then
		    SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
		    Exit Sub
		end if
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)
        Dim bufAttr

		FOR iRows = 0 to FResultcount -1
            bufAttr = groupKeyVal(iRows,0)
            SET FItemList(iRows) = new SearchGroupByItems
			FItemList(iRows).FAttribCd		= LEFT(bufAttr,6)
			FItemList(iRows).FAttribName	= replace(Mid(bufAttr,7,500),"_"," ")

            'rw "FAttribCd:"&FItemList(iRows).FAttribCd
            'rw "FAttribName:"&FItemList(iRows).FAttribName

			FItemList(iRows).FSubTotal 	= groupsize(iRows)
			FTotalCount = FTotalCount + groupsize(iRows)
		NEXT

		SET groupKeyVal = NOTHING
        SET groupsize   = NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB


	'####### 상품 검색 총 카운팅  ######
	PUBLIC SUB getTotalCount()

		'// 검색 결과 출력 시나리오명
		if FcolorCode="" or FcolorCode="0" then
			Scn= "scn_dt_itemDisp"		'일반 상품 검색
		else
			'Scn= "scn_dt_itemDispColor"		'상품 컬러별 검색(전시카테고리)
			Scn= "scn_dt_itemDisp"		'일반 상품 검색
		end if
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
        ''칼라칩검색/브랜드검색 3차
        if (Scn= "scn_dt_itemDispColor") or (FRectMakerid<>"") then SvrAddr = getTimeChkAddr(G_3RDSCH_ADDR)
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

    '####### 상품색상별 카운팅  ######
	PUBLIC SUB getTotalItemColorCount()

        '' 컬러칩 박스.
		'// 검색 결과 출력 시나리오명
		Scn= "scn_dt_itemDispColorGroup"		'상품 컬러별 검색(전시카테고리)

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim groupKeyCnt,groupKeyVal,groupsize

		dim FseekTime

		StartNum = 0 						'// 검색시작 Row
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성
		SortQuery = "Group by colorgrp Order by colorgrp "	'// 정렬 쿼리 생성

		IF SearchQuery="" Then
			EXIT SUB
		End If

		dim Rowids,Scores
		FTotalCount = 0
        ''---------------------------------------------------------------------------------------------------------
        if (FRectCateCode<>"") or (FRectMakerid<>"") and (FRectSearchTxt="")  then
            SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR)
        end if
        ''칼라칩검색/브랜드검색 3차
        if (Scn= "scn_dt_itemDispColor") or (FRectMakerid<>"") then SvrAddr = getTimeChkAddr(G_3RDSCH_ADDR)
        ''---------------------------------------------------------------------------------------------------------

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

	    ret = Docruzer.Search(SvrAddr&":"&SvrPort, _
						Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,"",StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		ret = Docruzer.GetResult_GroupBy(FResultcount,groupKeyCnt,groupKeyVal,groupsize,100)

		IF( ret < 0 ) THEN
			'Response.write "GetResult_Row: " & Docruzer.msg
			SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
			EXIT Sub
		END IF

		if (groupKeyCnt<1) then            ''그룹 갯수는 1개;
		    SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
		    Exit Sub
		end if

CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)
        dim bufColor, pos1

		FOR iRows = 0 to FResultcount -1
            bufColor = groupKeyVal(iRows,0)

            ''rw "groupKeyVal:"&bufColor
            if (Len(bufColor)>=6) then
    			SET FItemList(iRows) = new SearchGroupByItems
				FItemList(iRows).FcolorCode		= mid(bufColor,4,3)  ''colorgrp
				FItemList(iRows).FcolorName		= getCdPosVal(FItemList(iRows).FcolorCode,G_KSCOLORCD,G_KSCOLORNM)                                  ''mid(bufColor,8,pos1)
				FItemList(iRows).FcolorIcon		= "http://webimage.10x10.co.kr/color/colorchip/" & "chip"&CLNG(FItemList(iRows).FcolorCode)&".gif"  ''mid(bufColor,pos1+1,255)

				FItemList(iRows).FSubTotal 	= groupsize(iRows)
				FTotalCount = FTotalCount + groupsize(iRows)
            end if

		NEXT

        SET groupKeyVal = NOTHING
        SET groupsize   = NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB


	'####### 상품가격 범위 (최소가격,최대가격)  ######
	PUBLIC SUB getItemPriceRange()

		'// 검색 결과 출력 시나리오명
		Scn= "scn_dt_itemDisp"		'상품일반 검색

		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// 검색시작 Row
		FPageSize = 2						'// 결과 값
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성

		IF SearchQuery="" Then EXIT SUB

		dim Rowids,Scores

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF

		'//정렬 쿼리 생성 : 최소가격
		''SortQuery = " order by sellcash asc "
		SortQuery = " extract by minmax(sellcash)"  '2015 search4 style

		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, AuthCode, Logs, Scn, SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(0)

		if (FResultcount=1) or (FResultcount=2) then
            SET FItemList(0) = new SearchGroupByItems

			ret = Docruzer.GetResult_Row( arrData, arrSize, 0 )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				FResultcount = 0
				EXIT Sub
			END IF


			FItemList(0).FminPrice		= arrData(7)
			FItemList(0).FmaxPrice		= arrData(7)

		    if (FResultcount=2) then '' 결과가 1개인 경우가 있음
    		    ret = Docruzer.GetResult_Row( arrData, arrSize, 1 )

    			IF( ret < 0 ) THEN
    				'Response.write "GetResult_Row: " & Docruzer.msg
    				FResultcount = 0
    				EXIT Sub
    			END IF

    			FItemList(0).FmaxPrice		= arrData(7)
            end if

			SET arrData = NOTHING
			SET arrSize = NOTHING

	    end if
		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

'		SET FItemList(0) = new SearchGroupByItems
'
'		if FResultcount>0 then
'			ret = Docruzer.GetResult_Row( arrData, arrSize, 0 )
'			IF ret>=0 THEN
'				FItemList(0).FminPrice		= arrData(7)
'			end if
'			SET arrData = NOTHING
'			SET arrSize = NOTHING
'		end if
'

exit sub

		'//정렬 쿼리 생성 : 최대가격
		SortQuery = " order by sellcash desc "
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, AuthCode, Logs, Scn, SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		if FResultcount>0 then
			ret = Docruzer.GetResult_Row( arrData, arrSize, 0 )
			IF ret>=0 THEN
				FItemList(0).FmaxPrice		= arrData(7)
			end if
			SET arrData = NOTHING
			SET arrSize = NOTHING
		end if

		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

	'' ## 형태소 분석
	'' iextractLevel (2 or 5) 2:기본값
	public Function ExtractKeyword(byval iextractLevel)
		Dim Docruzer,ret
		Dim iRows
		Dim out_count,out_kwd
		Dim MaxCnt : MaxCnt =10
        dim ret_val, i, j

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106번으로 일단

		ret = Docruzer.ExtractKeyword _
						(SvrAddr&":"&SvrPort,_
						out_count,out_kwd,MaxCnt,"h1" _
						,FRectSearchTxt _
						,0,Docruzer.LC_KOREAN, Docruzer.CS_UTF8,CLNG(iextractLevel))

		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			rw "ERR"
			SET Docruzer = NOTHING
			EXIT FUNCTION

		END IF
		ret_val = ""
        for i=0 to out_count-1
            ret_val=ret_val&out_kwd(i)&vbCRLF
        next

        ret_val=TRim(ret_val)
        if (Right(ret_val,1)=vbCRLF) then ret_val=Left(ret_val,Len(ret_val)-1)

		ExtractKeyword = ret_val
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING
	end Function

	public Function OLD_ExtractKeyword()
		Dim Docruzer,ret
		Dim iRows
		Dim out_count,out_kwd
		Dim MaxCnt : MaxCnt =10
        dim ret_val, i, j

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106번으로 일단

		ret = Docruzer.ExtractKeyword _
						(SvrAddr&":"&SvrPort,_
						out_count,out_kwd,MaxCnt,"h1" _
						,replace(FRectSearchTxt," ","") _
						,0,Docruzer.LC_KOREAN, Docruzer.CS_UTF8,2)

		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			rw "ERR"
			SET Docruzer = NOTHING
			EXIT FUNCTION

		END IF
		ret_val = ""
        for i=0 to out_count-1
            ret_val=ret_val&out_kwd(i)&vbCRLF
        next

        ret_val=TRim(ret_val)
        if (Right(ret_val,1)=vbCRLF) then ret_val=Left(ret_val,Len(ret_val)-1)

		OLD_ExtractKeyword = ret_val
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING
	end Function

    '####### 동의어 2018/03/30  ######
	PUBLIC FUNCTION GetSynonymList()

		Dim Docruzer,ret
		Dim iRows
		Dim term_count,synonym_count,synonym_list
		Dim MaxCnt : MaxCnt =10
        dim ret_val, i, j

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106번으로 일단

		ret = Docruzer.GetSynonymList _
						(SvrAddr&":"&SvrPort,_
						term_count,synonym_count,synonym_list,_
						MaxCnt,replace(FRectSearchTxt," ","") _
						,1,1,1,4,2,2)

		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			rw "ERR"
			SET Docruzer = NOTHING
			EXIT FUNCTION

		END IF
		ret_val = ""
        for i=0 to term_count-1
			ret_val = ret_val&"["
            for j=0 to synonym_count(i)-1
                ret_val=ret_val&synonym_list(i,j)&","
            next
			ret_val=ret_val&"]"
            if (i<term_count-1) then ret_val=ret_val&vbCRLF
        next

        ret_val=TRim(ret_val)
		ret_val=replace(ret_val,",]","]")
        if (Right(ret_val,1)=",") then ret_val=Left(ret_val,Len(ret_val)-1)

		GetSynonymList = ret_val
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION

	'####### 추천검색어  ######
	PUBLIC FUNCTION getRecommendKeyWords()

		Dim Docruzer,ret
		Dim iRows
		Dim arrData,arrSize
		Dim MaxCnt : MaxCnt =5

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106번으로 일단

		ret = Docruzer.RecommendKeyWord _
						(SvrAddr&":"&SvrPort,_
						arrSize,arrData,_
						MaxCnt,replace(FRectSearchTxt," ",""),0)

		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT FUNCTION

		END IF

		getRecommendKeyWords = arrData
		SET arrData = NOTHING
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION

	'####### 인기검색어  ######
	PUBLIC FUNCTION getPopularKeyWords()

		DIM Docruzer,ret
		DIM iRows
		DIM arrData,arrSize
		DIM MaxCnt : MaxCnt =FPageSize

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106번으로 일단

		ret = Docruzer.getPopularKeyword _
						(SvrAddr&":"&SvrPort,_
						arrSize,arrData,_
						MaxCnt,4)
		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT FUNCTION
		END IF
		getPopularKeyWords = arrData
		SET arrData = NOTHING
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION

	'####### 인기검색어 (추가정보) ######
	PUBLIC FUNCTION getPopularKeyWords2(byRef arDt, byRef arTg)

		DIM Docruzer,ret
		DIM iRows
		DIM arrSize, arrData, arrTag
		DIM MaxCnt : MaxCnt =FPageSize

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT FUNCTION
		END IF

        ''SvrAddr = G_3RDSCH_ADDR ''G_ORGSCH_ADDR  '' 106번으로 일단
        SvrAddr = G_ORGSCH_ADDR

		'인기 검색어 (추가정보)
		ret = Docruzer.getPopularKeyword2 _
						(SvrAddr&":"&SvrPort,_
						arrSize,arrData,arrTag,_
						MaxCnt,4)

		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT FUNCTION
		END IF

		arDt = arrData
		arTg = arrTag

		SET arrData = NOTHING
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION

    '####### 실시간인기검색어:K-search  ######
	PUBLIC FUNCTION getRealtimePopularKeyWords(byRef arDt, byRef arTg)

		DIM Docruzer,ret
		DIM iRows
		DIM arrData,arrSize, arrTags
		DIM MaxCnt : MaxCnt =FPageSize

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106번으로 일단

		ret = Docruzer.getRealTimePopularKeyword _
						(SvrAddr&":"&SvrPort,_
						arrSize,arrData,arrTags,_
						MaxCnt,1,0)                     ''' 0 file / 1 memory
		IF( ret < 0 ) THEN
		    'rw "TTT:"&Docruzer.GetErrorMessage()
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT FUNCTION
		END IF
		arDt = arrData
		arTg = arrTags
		SET arrData = NOTHING
		SET arrTags = NOTHING
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION

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

'###############################
'### 고객 상품조회 이력      ###
'###############################
Class ViewHistoryCls
public FItemList()

	public FResultCount
	public FPageSize
	public FCurrpage
	public FTotalCount
	public FTotalPage
	public FScrollCount

	public FRectUserID

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
		FPageSize = 100
		FCurrpage = 1
	End Sub

	Private Sub Class_Terminate()
		'//
	End Sub

	public Function GetRandomUserID()
		dim sql, i

		sql = " select top 1 userid "
		sql = sql & " from ( "
		sql = sql & " 	select top 100 userid "
		sql = sql & " 	from "
		sql = sql & " 	[db_EVT].[dbo].[tbl_itemevent_userLogData_BACK] "
		sql = sql & " 	where type = 'item' "
		sql = sql & " 	order by idx desc "
		sql = sql & " ) T "
		sql = sql & " order by newid() "
		rsEVTget.Open SQL,dbEVTget,1

		GetRandomUserID = ""
		if  not rsEVTget.EOF  then
			GetRandomUserID = rsEVTget("userid")
		end if
		rsEVTget.close

	End Function

    public Sub getMyTodayViewListNew()
        dim sqlStr, i

        sqlStr = " Select top " & CStr(FPageSize*FCurrPage) & " * From "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + " 	select max(L.idx) as idx, i.cate_large, i.cate_mid, i.cate_small "
        sqlStr = sqlStr + " 		,i.itemid, i.itemname, i.sellcash, i.sellyn, i.limityn, i.limitno, i.limitsold "
        sqlStr = sqlStr + " 		,i.itemgubun, i.deliverytype, i.evalcnt, i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue,i.curritemcouponidx "
        sqlStr = sqlStr + " 		,i.basicimage, i.smallimage, i.listimage, i.listimage120, i.icon1image, i.icon2image, (Case When isNull(i.frontMakerid,'')='' then i.makerid else i.frontMakerid end) as makerid, i.brandname, max(L.regdate) as regdate, i.sailyn, i.sailprice "
		sqlStr = sqlStr + " 		,i.orgprice, i.specialuseritem, i.optioncnt, i.itemdiv, i.dispcate1, i.sellSTDate, [db_analyze_data_raw].[dbo].[getDisplayCateName](convert(bigint, left(d.catecode, 9))) as code_nm "
		''sqlStr = sqlStr + " 		, [db_analyze_data_raw].[dbo].[getKeywordsFromItemName](i.itemname) as kwds "
        sqlStr = sqlStr + " 	From db_evt.[dbo].[tbl_itemevent_userLogData_FrontRecent] L "
        sqlStr = sqlStr + " 	inner join db_analyze_data_raw.dbo.tbl_item i on L.itemid = i.itemid "
		sqlStr = sqlStr + " 	left join [db_analyze_data_raw].[dbo].[tbl_display_cate_item] d on i.itemid = d.itemid and d.isDefault = 'Y' "
        sqlStr = sqlStr + " 	Where L.type in ('item', 'itemrect') And L.userid='"&FRectUserID&"' "
        sqlStr = sqlStr + " 	group by i.cate_large, i.cate_mid, i.cate_small "
        sqlStr = sqlStr + " 		,i.itemid, i.itemname, i.sellcash, i.sellyn, i.limityn, i.limitno, i.limitsold "
        sqlStr = sqlStr + " 		,i.itemgubun, i.deliverytype, i.evalcnt, i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue,i.curritemcouponidx "
        sqlStr = sqlStr + " 		,i.basicimage, i.smallimage, i.listimage, i.listimage120, i.icon1image, i.icon2image, (Case When isNull(i.frontMakerid,'')='' then i.makerid else i.frontMakerid end), i.brandname, i.sailyn, i.sailprice "
		sqlStr = sqlStr + " 		,i.orgprice, i.specialuseritem, i.evalcnt, i.optioncnt, i.itemdiv, i.dispcate1, i.sellSTDate, [db_analyze_data_raw].[dbo].[getDisplayCateName](convert(bigint, left(d.catecode, 9))) "
		''sqlStr = sqlStr + " 		,[db_analyze_data_raw].[dbo].[getKeywordsFromItemName](i.itemname) "
        sqlStr = sqlStr + " )AA Where itemid is not null "
		sqlStr = sqlStr + " order by idx "

        rsEVTget.pagesize = FPageSize
        rsEVTget.Open sqlStr, dbEVTget, 1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsEVTget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)
		i=0
		if Not rsEVTget.Eof then
			rsEVTget.absolutepage = FCurrPage
			do until rsEVTget.eof
				set FItemList(i) = new CCategoryPrdItem

				FItemList(i).FItemID       = rsEVTget("itemid")
				FItemList(i).FItemName     = db2html(rsEVTget("itemname"))
				FItemList(i).FItemDiv 	= rsEVTget("itemdiv")		'상품 속성

				FItemList(i).FSellcash     = rsEVTget("sellcash")
				FItemList(i).FSellYn       = rsEVTget("sellyn")
				FItemList(i).FLimitYn      = rsEVTget("limityn")
				FItemList(i).FLimitNo      = rsEVTget("limitno")
				FItemList(i).FLimitSold    = rsEVTget("limitsold")
				FItemList(i).Fitemgubun    = rsEVTget("itemgubun")
				FItemList(i).FDeliverytype = rsEVTget("deliverytype")

				FItemList(i).Fevalcnt       = rsEVTget("evalcnt")
				FItemList(i).Fitemcouponyn  = rsEVTget("itemcouponyn")
				FItemList(i).Fitemcoupontype 	= rsEVTget("itemcoupontype")
				FItemList(i).Fitemcouponvalue 	= rsEVTget("itemcouponvalue")
				FItemList(i).Fcurritemcouponidx = rsEVTget("curritemcouponidx")

				FItemList(i).FImageBasic = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsEVTget("basicimage")
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsEVTget("smallimage")
				FItemList(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsEVTget("listimage")
				FItemList(i).FImageList120 = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsEVTget("listimage120")

				If FItemList(i).FItemDiv="21" Then
				FItemList(i).FImageicon2 = "http://webimage.10x10.co.kr/image/icon2/" + rsEVTget("icon2image")
				Else
				FItemList(i).FImageicon2 = "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(FItemList(i).FItemid) + "/" + rsEVTget("icon2image")
				End If
				FItemList(i).FMakerID   = rsEVTget("makerid")
				FItemList(i).FBrandName = db2html(rsEVTget("brandname"))
				FItemList(i).FRegdate   = rsEVTget("regdate")

				FItemList(i).FSaleYn    = rsEVTget("sailyn")
				FItemList(i).FSailPrice = rsEVTget("sailprice")
				FItemList(i).FOrgPrice   = rsEVTget("orgprice")
				FItemList(i).FSpecialuseritem = rsEVTget("specialuseritem")
				FItemList(i).Fevalcnt = rsEVTget("evalcnt")
				FItemList(i).FOptioncnt	= rsEVTget("optioncnt")
				FItemList(i).FCateName = rsEVTget("code_nm")
				FItemList(i).Fsdate	= rsEVTget("sellSTDate")
		''		FItemList(i).Fkeywords	= rsEVTget("kwds")

				rsEVTget.movenext
				i=i+1
			loop
		end if
		rsEVTget.close
    end Sub

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

'###############################
'### 상품후기 검색           ###
'###############################
Class SearchItemEvaluate
	public FItemList()

	public FResultCount
	public FPageSize
	public FCurrpage
	public FTotalCount
	public FTotalPage
	public FScrollCount

	public FRectDispCate
	public FRectSort
	public FRectMode
	public FRectArrItemid

	Private Sub Class_Initialize()
		redim preserve FItemList(0)
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'// 베스트 리뷰 선택 목록(검색用) 접수 : 상품당 최근 1개 리뷰 목록 //
	public Sub GetBestReviewArrayList()
		dim sql, i

		if FRectArrItemid="" then
			Exit Sub
		end if

		'// 목록 접수 //
		sql = " SELECT e.idx, e.userid, e.regdate, e.itemid " + vbcrlf
		sql = sql & " , e.contents, (Case When isNull(i.frontMakerid,'')='' then i.makerid else i.frontMakerid end) as makerid, i.brandname, e.TotalPoint as Tpoint " + vbcrlf
		sql = sql + " , isnull(e.Point_function,0) as Point_function "
		sql = sql + " , isnull(e.Point_Design,0) as Point_Design "
		sql = sql + " , isnull(e.Point_Price,0) as Point_Price "
		sql = sql + " , isnull(e.Point_satisfy,0) as Point_satisfy "
		sql = sql & " , i.itemname, i.sellcash, i.orgprice, i.sellyn, i.sailyn, i.limityn, i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue " + vbcrlf
		sql = sql & " , i.listimage120, i.evalcnt, i.itemscore, e.File1, e.File2, i.icon1image, i.evalcnt, ic.favcount, i.tenOnlyYn  " + vbcrlf
		sql = sql & " , (case when isnull(ee.itemid,'') <> '' then 'Y' else 'N' end) as Eval_excludeyn"
		sql = sql & " FROM db_board.[dbo].tbl_item_evaluate e " + vbcrlf
		sql = sql & " JOIN  [db_item].[dbo].tbl_item i  " + vbcrlf
		sql = sql & " 	on e.itemid = i.itemid " + vbcrlf
		sql = sql & " JOIN  [db_item].[dbo].tbl_item_contents ic  " + vbcrlf
		sql = sql & " 	on i.itemid = ic.itemid " + vbcrlf
		sql = sql + " left join db_board.dbo.tbl_Item_Evaluate_exclude ee"
		sql = sql + " 	on e.itemid=ee.itemid"
		sql = sql & " WHERE e.idx in (" & vbcrlf
		sql = sql & "	 Select idx from (" & vbcrlf
		sql = sql & "		select itemid, max(idx) as idx " & vbcrlf
		sql = sql & "		from db_board.[dbo].tbl_item_evaluate" & vbcrlf
		sql = sql & "		where itemid in (" & FRectArrItemid & ")" & vbcrlf
		sql = sql & "			and isusing='Y'" & vbcrlf

		if FRectMode="photo" then
			sql = sql & " and (File1 is Not Null or File2 is Not Null) " + vbcrlf
		end if

		sql = sql & "		group by itemid) as T )"
		Select Case FRectSort
			Case "ne"
				'신상순
				sql = sql & " ORDER BY e.itemid DESC  " + vbcrlf
			Case "be"
				'인기상품 순
				sql = sql & " ORDER BY i.itemscore DESC  " + vbcrlf
			Case "lp"
				'낮은 가격순
				sql = sql & " ORDER BY i.sellcash asc  " + vbcrlf
			Case "hp"
				'높은 가격순
				sql = sql & " ORDER BY i.sellcash desc  " + vbcrlf
			Case "hs"
				'높은 할인순
				sql = sql & " ORDER BY (i.sellcash/i.orgprice) desc  " + vbcrlf
			Case else
				sql = sql & " ORDER BY e.itemid DESC  " + vbcrlf
		End Select

		rsget.Open SQL,dbget,1

		FResultCount = rsget.RecordCount
		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			Do Until rsget.EOF or rsget.BOF
				set FItemList(i) = new CCategoryPrdItem

				FItemList(i).fEval_excludeyn 	= rsget("Eval_excludeyn")
				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fitemid			= rsget("itemid")
				FItemList(i).Fuserid			= rsget("userid")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Fitemname			= db2html(rsget("itemname"))
				FItemList(i).Fmakerid			= db2html(rsget("makerid"))
				FItemList(i).Fbrandname			= db2html(rsget("brandname"))
				FItemList(i).Fevalcnt			= rsget("evalcnt")
				FItemList(i).Fcontents			= db2html(rsget("contents"))
				FItemList(i).FOrgprice			= rsget("orgprice")
				FItemList(i).FSellYn			= rsget("sellyn")
				FItemList(i).FSaleYn			= rsget("sailyn")
				FItemList(i).FLimitYn			= rsget("limityn")
				FItemList(i).FSellCash			= rsget("sellcash")
				FItemList(i).FPoints			= rsget("TPoint")
				FItemList(i).FPoint_fun			= rsget("Point_Function")
				FItemList(i).FPoint_dgn			= rsget("Point_Design")
				FItemList(i).FPoint_prc			= rsget("Point_Price")
				FItemList(i).FPoint_stf			= rsget("Point_Satisfy")
				FItemList(i).Fitemcouponyn		= rsget("itemcouponyn")
				FItemList(i).Fitemcoupontype	= rsget("itemcoupontype")
				FItemList(i).FItemCouponValue	= rsget("itemcouponvalue")
				FItemList(i).FTenOnlyYn			= rsget("tenOnlyYn")

				if Not(rsget("File1")="" or isNull(rsget("File1"))) then
					FItemList(i).FImageIcon1		= "http://imgstatic.10x10.co.kr/goodsimage/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("File1")
				end if
				if Not(rsget("File2")="" or isNull(rsget("File2"))) then
					FItemList(i).FImageIcon2		= "http://imgstatic.10x10.co.kr/goodsimage/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("File2")
				end if
				FItemList(i).FImageList120    = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage120")
				FItemList(i).FIcon1Image	  = "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon1image")
				FItemList(i).FEvalcnt		  = rsget("evalcnt")
				FItemList(i).Ffavcount		  = rsget("favcount")

				rsget.moveNext
				i = i + 1
			Loop
		end if
		rsget.close
	end Sub

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



'###############################
'### 이벤트 검색             ###
'###############################

Class SearchEventItems

	Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB

	PUBLIC Fevt_code
	PUBLIC Fevt_bannerimg
	PUBLIC Fevt_startdate
	PUBLIC Fevt_enddate
	PUBLIC Fevt_kind
	PUBLIC Fbrand
	PUBLIC Fevt_LinkType
	PUBLIC Fevt_bannerlink
	PUBLIC Fetc_itemid
	PUBLIC Fetc_itemimg
	PUBLIC Ficon1image
	PUBLIC Fevt_name
	PUBLIC Fevt_tag
	PUBLIC Fevt_subcopyK
	PUBLIC Fissale
	PUBLIC Fisgift
	PUBLIC Fisitemps
	PUBLIC Fiscoupon
	PUBLIC FisOnlyTen
	PUBLIC Fisoneplusone
	PUBLIC Fisfreedelivery
	PUBLIC Fisbookingsell
	PUBLIC Fiscomment
	PUBLIC Fevt_state

End Class

Class SearchEventCls
    ''검색페이지 이벤트.
	Private SUB Class_initialize()

        '' 기본 1차서버--------------------------------------------------------------------------------------------
		SvrAddr = getTimeChkAddr(G_1STSCH_ADDR)
		''---------------------------------------------------------------------------------------------------------

		SvrPort = "6167"'DocSvrPort
		AuthCode = "" '인증값
		Logs = "" '로그값

		FResultCount = 0
		FTotalCount = 0
		FPageSize = 10
		FCurrPage = 1
		FPageSize = 30

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
	dim FRectExceptText		'제외어
	dim FRectChannel		'검색 채널 (W:isWeb, M:isMobile, A:isApp)

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

	'####### 이벤트 검색 ######
	PUBLIC SUB getEventList()

		dim Scn : Scn= "scn_dt_event2015" 		'// 검색 결과 출력 시나리오명
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = (FCurrPage -1)*FPageSize	'// 검색시작 Row

		'// 검색 쿼리생성 (이벤트는 제외어 검색 하지말것 서버죽음 ㄷㄷ)
		'IF FRectExceptText<>"" Then
		'	SearchQuery = " (idx_eventKeyword='" & FRectSearchTxt & " ! " & FRectExceptText & "' BOOLEAN) "	'제외어
		'else
			SearchQuery = " idx_eventKeyword='" & FRectSearchTxt & "'  allword "	'키워드검색
		'End if

		SearchQuery = SearchQuery &_
					" and idx_evt_state='7' " &_
					" and idx_evt_startdate<='" & Replace(date(),"-","") & "000000' " &_
					" and idx_evt_enddate>='" & Replace(date(),"-","") & "000000' "

		Select Case FRectChannel
			Case "W"
				SearchQuery = SearchQuery & " and idx_evt_isWeb=1 "
			Case "M"
				SearchQuery = SearchQuery & " and idx_evt_isMobile=1 "
			Case "A"
				SearchQuery = SearchQuery & " and idx_evt_isApp=1 "
			Case Else
				SearchQuery = SearchQuery & " and idx_evt_isWeb=1 "
		End Select

		'//그룹 범위별 지정(정렬 쿼리 생성)
		SortQuery = "Order by evt_startdate desc "

		dim Rowids,Scores

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

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF


		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수

		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new SearchEventItems
				FItemList(iRows).Fevt_code			= arrData(0)
				FItemList(iRows).Fevt_bannerimg		= arrData(1)
				FItemList(iRows).Fevt_startdate		= dateSerial(left(arrData(2),4),mid(arrData(2),5,2),mid(arrData(2),7,2))	'### ASP 날짜형태로 변환
				FItemList(iRows).Fevt_enddate		= dateSerial(left(arrData(3),4),mid(arrData(3),5,2),mid(arrData(3),7,2))	'### ASP 날짜형태로 변환
				FItemList(iRows).Fevt_kind			= arrData(4)
				FItemList(iRows).Fbrand				= arrData(5)
				FItemList(iRows).Fevt_LinkType		= arrData(6)
				FItemList(iRows).Fevt_bannerlink	= arrData(7)
				FItemList(iRows).Fetc_itemid		= arrData(8)
				FItemList(iRows).Fetc_itemimg		= arrData(9)
				FItemList(iRows).Ficon1image		= arrData(10)
				FItemList(iRows).Fevt_name			= arrData(11)
				FItemList(iRows).Fevt_tag			= arrData(12)
				FItemList(iRows).Fevt_subcopyK		= arrData(13)
				FItemList(iRows).Fissale			= arrData(14)
				FItemList(iRows).Fisgift			= arrData(15)
				FItemList(iRows).Fisitemps			= arrData(16)
				FItemList(iRows).Fiscoupon			= arrData(17)
				FItemList(iRows).FisOnlyTen			= arrData(18)
				FItemList(iRows).Fisoneplusone		= arrData(19)
				FItemList(iRows).Fisfreedelivery	= arrData(20)
				FItemList(iRows).Fisbookingsell		= arrData(21)
				FItemList(iRows).Fiscomment			= arrData(22)
				FItemList(iRows).Fevt_state			= arrData(23)

			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT

		SET Rowids= NOTHING
		SET Scores= NOTHING

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
end Class




'###############################
'### PLAY 검색             ###
'###############################

Class SearchPlayItems

	Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB

	PUBLIC Fplaycate
	PUBLIC Fidx
	PUBLIC FlistImage
	PUBLIC Fplayname
	PUBLIC Fopendate
	PUBLIC FsortNo
	PUBLIC Fcont
	PUBLIC Ftag

	Function getPlayCateNm()
		Select Case Fplaycate
			Case "1"
				getPlayCateNm = "Ground"
			Case "2"
				getPlayCateNm = "Style+"
			Case "3"
				getPlayCateNm = "Color trend"
			Case "4"
				getPlayCateNm = "Design fingers"
			Case "5"
				getPlayCateNm = "그림일기"
			Case "6"
				getPlayCateNm = "Video clip"
			Case "7"
				getPlayCateNm = "T-episode"
		End Select
	End function

	Function getPlayCateLink()
		Select Case Fplaycate
			Case "1"
				getPlayCateLink = "/play/playGround?gidx="
			Case "2"
				getPlayCateLink = "/play/playStylePlusView.asp?idx="
			Case "3"
				getPlayCateLink = "/play/playColorTrendView.asp?ctcode="
			Case "4"
				getPlayCateLink = "/play/playdesignfingers.asp?fingerid="
			Case "5"
				getPlayCateLink = "/play/playPicDiary.asp?idx="
			Case "6"
				getPlayCateLink = "/play/playVideoClip.asp?idx="
			Case "7"
				getPlayCateLink = "/play/playtEpisodePhotopick.asp?"
		End Select
	End function

End Class

Class SearchPlayCls
    ''검색페이지 Play.
	Private SUB Class_initialize()

        '' 기본 1차서버--------------------------------------------------------------------------------------------
		SvrAddr = getTimeChkAddr(G_1STSCH_ADDR)
		''---------------------------------------------------------------------------------------------------------

		SvrPort = "6167"'DocSvrPort
		AuthCode = "" '인증값
		Logs = "" '로그값

		FResultCount = 0
		FTotalCount = 0
		FPageSize = 10
		FCurrPage = 1
		FPageSize = 30

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
	dim FRectExceptText		'제외어

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

	'####### PLAY 검색 ######
	PUBLIC SUB getPlayList()
        'rw "getPlayList - SKIP"
        'Exit SUB

		dim Scn : Scn= "scn_dt_play2013" 		'// 검색 결과 출력 시나리오명
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = (FCurrPage -1)*FPageSize	'// 검색시작 Row

		'// 검색 쿼리생성 (Play는 제외어 검색 하지말것 서버죽음 ㄷㄷ)
		SearchQuery = " idx_playKeyword='" & FRectSearchTxt & "'  allword "	'키워드검색

		'//그룹 범위별 지정(정렬 쿼리 생성)
		SortQuery = "Order by opendate desc, sortNo, idx desc "

		dim Rowids,Scores

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

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF


		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수

		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new SearchPlayItems
				FItemList(iRows).Fplaycate	= arrData(0)
				FItemList(iRows).Fidx		= arrData(1)
				FItemList(iRows).FlistImage	= arrData(2)
				FItemList(iRows).Fplayname	= arrData(3)
				FItemList(iRows).Fopendate	= dateSerial(left(arrData(4),4),mid(arrData(4),5,2),mid(arrData(4),7,2))	'### ASP 날짜형태로 변환
				FItemList(iRows).FsortNo	= arrData(5)
				FItemList(iRows).Fcont		= arrData(6)
				FItemList(iRows).Ftag		= arrData(7)

			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT

		SET Rowids= NOTHING
		SET Scores= NOTHING

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
end Class



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
%>
