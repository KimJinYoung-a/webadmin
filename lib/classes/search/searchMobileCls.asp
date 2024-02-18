<%
''CS_EUCKR => CS_UTF8
''모바일 class U8로 변경
'''--------------------------------------------------------------------------------------
DIM G_KSCOLORCD : G_KSCOLORCD = Array("023","001","002","010","021","003","004","024","019","005","016","006","007","020","008","018","017","009","011","012","022","013","014","015","025","026","027","028","029","030","031")
DIM G_KSCOLORNM : G_KSCOLORNM = Array("와인","빨강","주황","갈색","카멜","노랑","베이지","아이보리","카키","초록","민트","연파랑","파랑","네이비","보라","연보라","베이비핑크","핑크","흰색","옅은회색","짙은회색","검정","은색","금색","체크","스트라이프","도트","플라워","드로잉","애니멀","기하학")

Dim G_KSSTYLECD : G_KSSTYLECD = Array("010","020","030","040","050","060","070","080","090")
Dim G_KSSTYLENM : G_KSSTYLENM = Array("클래식","큐티","댄디","모던","내추럴","오리엔탈","팝","로맨틱","빈티지")

DIM G_ORGSCH_ADDR , GG_ORGSCH_ADDR
DIM G_1STSCH_ADDR , GG_1STSCH_ADDR
DIM G_2NDSCH_ADDR , GG_2NDSCH_ADDR
DIM G_3RDSCH_ADDR , GG_3RDSCH_ADDR
Dim G_4THSCH_ADDR , GG_4THSCH_ADDR

DIM G_SCH_TIME : G_SCH_TIME=formatdatetime(now(),4)

IF (application("Svr_Info") = "Dev") THEN
     G_1STSCH_ADDR = "192.168.50.10"  ''"110.93.128.109" ''
     G_2NDSCH_ADDR = "192.168.50.10"
     G_3RDSCH_ADDR = "192.168.50.10"
     G_4THSCH_ADDR = "192.168.50.10"
     G_ORGSCH_ADDR = "192.168.50.10"
ELSE
     G_1STSCH_ADDR = "192.168.0.210"        ''192.168.0.210  :: www 검색페이지(search.asp)   '
     G_2NDSCH_ADDR = "192.168.0.207"        ''192.168.0.207  :: www 카테고리, 상품, 브랜드
     G_3RDSCH_ADDR = "192.168.0.209"        ''192.168.0.209  :: app 
     G_4THSCH_ADDR = "192.168.0.208"        ''192.168.0.208  :: mobile 6:10분에 인덱싱 시작 카피
     G_ORGSCH_ADDR = "192.168.0.206"        ''192.168.0.206
END IF

GG_1STSCH_ADDR = G_1STSCH_ADDR
GG_2NDSCH_ADDR = G_2NDSCH_ADDR
GG_3RDSCH_ADDR = G_3RDSCH_ADDR
GG_4THSCH_ADDR = G_4THSCH_ADDR
GG_ORGSCH_ADDR = G_ORGSCH_ADDR

'' 2017/10/09 추가 =================================================================================================
'' 오류 발생시 Application("G_4THSCH_ADDR")=G_ORGSCH_ADDR 로 변경. 
'' 이후 오류 복구시 Application("G_4THSCH_ADDR")="" 이부분 주석 해제후 G_4THSCH_ADDR 로 치환할것. 이후 다시 주석처리
'' Application("G_4THSCH_ADDR")=""
if (Application("G_4THSCH_ADDR")="") then
    Application("G_4THSCH_ADDR")=G_4THSCH_ADDR
end if

G_4THSCH_ADDR=Application("G_4THSCH_ADDR")

''apps\appCom\wish\protoV3\searchKeyword2017.asp 에서 사용하므로 이것도 추가
if (Application("G_3RDSCH_ADDR")="") then
    Application("G_3RDSCH_ADDR")=G_3RDSCH_ADDR
end if

G_3RDSCH_ADDR=Application("G_3RDSCH_ADDR")
''response.write G_4THSCH_ADDR
'' ==================================================================================================================

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
	PUBLIC FCateName1
	PUBLIC FCateName2
	PUBLIC FCateName3
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
		SvrAddr = getTimeChkAddr(G_4THSCH_ADDR)
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
	dim FRectSortMethod		'정렬방식 (ne:신상품, be:인기상품, lp:낮은가격, hp:높은가격, hs:할인률, br:상품후기, ws:위시수, bs:판매수)
	dim FRectSearchFlag 	'검색범위 (sc:세일쿠폰, ea:사용기전체, ep:포토사용기, ne:신상품, fv:위시상품, pk:포장서비스)
	dim FRectColorExclude	'제외 컬러
	dim FRectDispExclude	'제외 카테

	dim FRectMakerid		'업체 아이디
	dim FRectCateCode		'카테고리코드
	dim FListDiv			'카테고리/검색 구분용
	dim FSellScope			'판매가능 상품검색 여부
	dim FGroupScope			'검색시 그룹핑 범위 (1:1depth, 2:2depth, 3:3depth)
	dim FdeliType			'배송방법 (FD:무료배송, TN:텐바이텐 배송, FT:무료+텐바이텐 배송, WD:해외배송)
	dim FRectItemID

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
	dim FSubShopCd			'서브샵(100:다이어리스토리)
	dim FGiftDiv			'사은품 종류 ( R: 다이어리 사은품 )
	dim FNewYn				'신규 상품 여부 Y/N
	dim FBestYn				'베스트 상품 여부 Y/N
	dim FawardType			'베스트 리스트 타입 period : 기간별 ,
	Dim FRectNoDealItem		'검색 시 Deal 상품 제외

	dim fRectadultType		' 검색시 성인용품

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

	Function removeDupVal(ByVal varArr, idx)
		'*******************************
		' 배열 요소 중복제거
		' - param
		' varArr : 배열
		' idx    : 코드 그룹 분할 인덱스
		' - return
		' 중복제거 배열
		'*******************************	
		Dim dic, item, rtnVal

		Set dic = CreateObject("Scripting.Dictionary")
		dic.removeall
		dic.CompareMode = 0

		For Each item In varArr
			If not dic.Exists(Left(item, idx)) Then dic.Add Left(item, idx), Left(item, idx)
		Next

		rtnVal = dic.keys
		Set dic = Nothing
		removeDupVal = rtnVal
	End Function

	function generateAttrCode(rawData, codeIdx)
		'*******************************
		' 속성 쿼리 그룹 생성
		' - param
		' rawData : 속성 파라미터 ex) "301001,301002,301003,301004,305001,304001"
		' idx     : 코드 그룹 분할 인덱스
		' - return
		' 그룹 분할 쿼리 ex) "and (idx_attribCd='301001 301002 301003 301004' anystring) and (idx_attribCd='305001' anystring) and (idx_attribCd='304001' anystring)"
		'*******************************		
		dim attArr : attArr = split(rawData,",")
		if(not isArray(attArr)) then exit function
		dim groupCodeArr : groupCodeArr = removeDupVal(attArr, codeIdx)
		dim attrStr, item, key, attr
		dim dic
		set dic = Server.CreateObject("Scripting.Dictionary")

		'초기화
		for Each item in groupCodeArr
			if item <> "" then
				dic.Add item, ""
			end if
		next

		' ex)
		' key - value
		'===========================
		' 301 : 301001 301002
		' 302 : 302001 302002
		' 305 : 305001 305002 305002
		for each key in dic
			for each attr in attArr
				if Left(attr, codeIdx) = key Then
					dic.item(key) = dic.item(key) + attr + " "
				end if
			next
		next

		for each key in dic
			attrStr = attrStr & " and " & "(idx_attribCd='"& trim(dic.item(key)) &"' anystring)"
		next
		generateAttrCode = attrStr
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
			Case "saleonly"
				'할인상품 목록(무료배송쿠폰제외)
				FRectSearchFlag = "os"
			Case "fulllist"
				'카테고리없는 전체.
			Case "subshop"
				'서브샵
				IF (FSubShopCd="" or isNull(FSubShopCd)) Then EXIT FUNCTION				 				
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
				strQue = getQrCon(strQue) & "idx_itemname='" & FRectSearchTxt & "'  allword synonym "	'키워드검색(동의어 포함) synonym
				'strQue = getQrCon(strQue) & "idx_itemname='" & FRectSearchTxt & "'  natural "		'자연어 검색(동의어 포함) synonym
			End if
		END IF

		'############# 딜상품 제외 테스트 조건(실섭 적용, 스테이징 테스트 완료 후 삭제) ######################
		'// 2019-09-24 원승현 18주년 취향 이벤트에서 사용하기 위해 
		If FRectNoDealItem="Y" Then
			strQue = strQue & getQrCon(strQue) & "idx_itemdiv != '21' "
		End if

		'@ 카테고리 검색 범위
		'IF FRectSearchItemDiv="y" Then
		'	''strQue = strQue & getQrCon(strQue) & "idx_isDefault='y' "
		'END IF
		
		'@상품코드
		IF FRectItemID<>"" Then
			strQue = strQue & getQrCon(strQue) & FRectItemID
		END IF

		'@ 카테고리
		IF FRectCateCode<>"" Then
			if FRectSearchCateDep="X" then
				strQue = strQue & getQrCon(strQue) & "idx_catecode='" & FRectCateCode & "'"
			else
				IF FRectSearchItemDiv="y" Then ''기본카테고리
			        ''strQue = strQue & getQrCon(strQue) & "idx_catecode like '" & FRectCateCode & "*'"  ''2017/10/27 검색성능 관련 - konan 조현복
			        strQue = strQue & getQrCon(strQue) & "idx_catecodelist='" & FRectCateCode & "'"
			    else                           ''추가카테검색
			        ''strQue = strQue & getQrCon(strQue) & "idx_catecodeExt like '" & FRectCateCode & "*'"  ''2017/10/27 검색성능 관련 - konan 조현복
			        strQue = strQue & getQrCon(strQue) & "idx_catecodeExt='" & FRectCateCode & "'"
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
					''strQue = strQue & " idx_catecode like '" & RequestCheckVar(LCase(trim(arrCt(lpCt))),18) & "*' "  ''2017/10/27 검색성능 관련 - konan 조현복
					strQue = strQue & " idx_catecodelist='" & RequestCheckVar(LCase(trim(arrCt(lpCt))),18) & "' "
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
				Case "os"	'할인전용 (무료배송 쿠폰 제외)
					strQue= strQue & getQrCon(strQue) & "(idx_saleyn='Y' or (idx_itemcouponyn='Y' and idx_isFreeBeasong='N')) "
				Case "pk"	'포장서비스
					strQue = strQue & getQrCon(strQue) & "idx_pojangok='Y' and (deliverytype='1' or deliverytype='4') "
				Case "scpk"	'포장서비스 & 세일쿠폰
					strQue = strQue & getQrCon(strQue) & "((idx_saleyn='Y' or idx_itemcouponyn='Y') and idx_pojangok='Y') and (deliverytype='1' or deliverytype='4') "
				Case "qq"   '바로배송 Quick
				    strQue = strQue & getQrCon(strQue) & "(idx_deliverfixday='Q') "
			End Select
		END IF

		'@ 서브샵
		IF FSubShopCd<>"" THEN
			strQue= strQue & getQrCon(strQue) & "subshoplist = '"& FSubShopCd &"' "
		END IF 		

		IF FGiftDiv>"" THEN
			Select Case FGiftDiv
				Case "R"	'다이어리사은품
					strQue= strQue & getQrCon(strQue) & "idx_giftdiv='R' "
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
			'// 해외직구배송작업추가(원승현)
			Case "QT"	'바로배송
				strQue = strQue & getQrCon(strQue) & "idx_deliverfixday='Q'"
			Case "DT"	'해외직구
				strQue = strQue & getQrCon(strQue) & "idx_deliverfixday='G'"
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

		' 서브샵 속성코드 별도처리 2019-08-29
		' 속성별 and 조건으로 그룹화
		' ex) and (idx_attribCd='301001 301002 301003 301004' anystring) and (idx_attribCd='305001' anystring) and (idx_attribCd='304001' anystring)
		IF FSubShopCd = "100" THEN '일단 다이어리 스토리만
			strQue = strQue & generateAttrCode(FattribCd, 3)	
		else
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
		end if

		'@ 상품 판매 범위
		IF FSellScope="Y" Then
			strQue = strQue & getQrCon(strQue) & "idx_sellyn='Y' "
		ELSE
			strQue = strQue & getQrCon(strQue) & "(idx_sellyn='Y' or idx_sellyn='S') "
		End IF
        
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

		IF FawardType = "period" or FawardType = "userlevel" then 
			strQue = strQue & getQrCon(strQue) & "idx_bestyn='Y' "
		end if 

		IF fRectadultType<>"" THEN
			strQue= strQue & getQrCon(strQue) & "idx_adultType = "& fRectadultType &" "
		END IF 	

		query = strQue
	End FUNCTION

	Sub getSortQuery(byref query)
		dim strQue

		'// 중복 상품 제거(중복 등록 카테고리일경우) 2015 제거
		'IF (FRectCateCode<>"" and FRectSearchItemDiv<>"y") Then '' 추가 카테고리 검색시
    	'	strQue = " GROUP BY itemid"
    	'END IF

		'// 정렬
		IF FRectSortMethod="ne" THEN '신상품
			strQue = strQue & " ORDER BY regdate desc, itemid desc"
		ELSEIF FRectSortMethod="be" THEN '인기상품
			if (FListDiv<>"search") then
			    strQue = strQue & " ORDER BY itemscore desc,itemid desc"
			else
    			''2018/03/26 A/B TEST 종료
    			strQue = strQue & " ORDER BY $MATCHFIELD(cateboostkeylist,bestkeylist) desc, itemscore desc,itemid desc"
    		end if
		ELSEIF FRectSortMethod="vv" THEN ''MATCHFIELD 필드 사용. 2018/03/20 TEST
		    strQue = strQue & " ORDER BY $MATCHFIELD(bestkeylist) desc, itemscore desc,itemid desc"
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
		ELSEIF FRectSortMethod="bs" THEN '판매수순
			if (FListDiv<>"search") then '' 분리 2019/06/14
		    	strQue = strQue & " ORDER BY sellCnt desc,sellcash desc,itemid desc"
			else
				strQue = strQue & " ORDER BY $MATCHFIELD(cateboostkeylist,bestkeylist) desc, sellCnt desc, itemscore desc, itemid desc"
			end if
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
        Dim iDocErrMsg
        
        if (FPageSize>300) then FPageSize=300  ''2016/11/16추가
            
		'// 검색 결과 출력 시나리오명
		if FcolorCode="" or FcolorCode="0" then
			Scn= "scn_dt_itemDisp"		'일반 상품 검색
		else
			'Scn= "scn_dt_itemColor"		'상품 컬러별 검색
			'Scn= "scn_dt_itemDispColor"		'상품 컬러별 검색(전시카테고리)
			Scn= "scn_dt_itemDisp"		    '일반 상품 검색 통일 2015
		end if

		StartNum = (FCurrPage -1)*FPageSize '// 검색시작 Row

		CALL getSearchQuery(SearchQuery)	'// 검색 쿼리생성
		CALL getSortQuery(SortQuery)		'// 정렬 쿼리 생성
		''Response.Write SearchQuery &"<Br>"
		IF SearchQuery="" THEN
			EXIT SUB
		END IF

		IF (FLogsAccept) and (FRectSearchTxt<>"") and (FCurrPage="1") THEN ''1페이지만 남김
            ''Logs = ("상품+^" & FRectSearchTxt & "]##" & FRectSearchTxt & "||" & FRectPrevSearchTxt ) 	'// 로그값
            ''if (now()>"2015-03-05") then
                Dim iLOG_SITE : iLOG_SITE = "MOB"
                Dim iLOG_CATE : iLOG_CATE = "RECT" 
                Dim iLOG_USER : iLOG_USER = GetUserStrlarge(GetLoginUserLevel) '' 회원등급을 사용
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
            ''end if
            
		END IF

        ''상품검색/브랜드검색이 아닌경우 2차서버로.
        ''---------------------------------------------------------------------------------------------------------
        'if (FRectCateCode<>"") or (FRectMakerid<>"") and (FRectSearchTxt="")  then
        '    'response.write "2차서버<br>"
        '     SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR) 
        'end if
        ''칼라칩검색/브랜드검색 3차
        'if (Scn= "scn_dt_itemDispColor") or (FRectMakerid<>"") then
        '	'response.write "3차서버<br>"
        '	SvrAddr = getTimeChkAddr(G_3RDSCH_ADDR)
        'end if
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
			dbget.execute "EXECUTE db_log.dbo.sp_Ten_DocLog @ErrMsg ='["&SvrAddr&"]"& html2db(Docruzer.GetErrorMessage())&"["&Request.ServerVariables("REMOTE_ADDR")&"]["&Request.ServerVariables("LOCAL_ADDR")&"]["& html2db(SearchQuery)&"]'"
			
            iDocErrMsg = Docruzer.GetErrorMessage()
            if (InStr(iDocErrMsg,"recv queue full")>0) or (InStr(iDocErrMsg,"socket time out")>0) or (InStr(iDocErrMsg,"cannot connect to server")>0) then
                if (Application("G_4THSCH_ADDR")=GG_4THSCH_ADDR) then
                    Application("G_4THSCH_ADDR") = GG_ORGSCH_ADDR
                elseif (Application("G_4THSCH_ADDR")=GG_ORGSCH_ADDR) then
                    Application("G_4THSCH_ADDR") = GG_2NDSCH_ADDR
                end if
            end if
        
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING

''			IF FListDiv<>"search" THEN
''				'// 1번 서버 에서시 2번에서 구동(2번도 에러면 Skip)
''				if (SvrAddr = G_1STSCH_ADDR) then
''					SvrAddr = G_2NDSCH_ADDR  ''"192.168.0.108"
''					if (G_1STSCH_ADDR<>G_2NDSCH_ADDR) then
''					    call getSearchList()
''				    end if
''				end if
''			END IF

			EXIT SUB
		END IF

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수
		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
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
				FItemList(iRows).FItemDiv	= arrData(3)
				FItemList(iRows).FItemid = arrData(4)
				FItemList(iRows).FItemName = db2html(arrData(5))
				FItemList(iRows).FKeyWords = db2html(arrData(6))
				FItemList(iRows).FSellCash = arrData(7)
				FItemList(iRows).FOrgPrice = arrData(8)
				FItemList(iRows).FMakerId = arrData(9)
				FItemList(iRows).FImageIcon1 	= getItemImageUrl & "/icon1/" & GetImageSubFolderByItemid(FItemList(iRows).FItemid) & "/" &db2html(arrData(15))

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

		If FRectDispExclude <> "" Then
			SearchQuery = SearchQuery & FRectDispExclude '" and idx_catecode != '123' "
		End If

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
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF


		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)

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
				
				if Len(FItemList(iRows).FCateCd1)>=4 then
				    FItemList(iRows).FCateCd1	= Mid(FItemList(iRows).FCateCd1,5,255)              '' sort4+[code3]
				end if
				    
				if Len(FItemList(iRows).FCateCd2)>=11 then
				    FItemList(iRows).FCateCd2	= Mid(FItemList(iRows).FCateCd2,9+3,255)            '' sort4+4+code3+[code3]
				end if
				    
				if Len(FItemList(iRows).FCateCd3)>=18 then
				    FItemList(iRows).FCateCd3	= Mid(FItemList(iRows).FCateCd3,13+3+3,255)         '' sort4+4+4+code3+[code3]
				end if
				
				FItemList(iRows).FCateName1 = Replace(Split(arrData(1),"^^")(0),",","")
				
				if UBound(Split(arrData(1),"^^")) > 0 then
					FItemList(iRows).FCateName2 = Replace(Split(arrData(1),"^^")(1),",","")
					
					if UBound(Split(arrData(1),"^^")) > 1 then
						FItemList(iRows).FCateName3 = Replace(Split(arrData(1),"^^")(2),",","")
					else
						FItemList(iRows).FCateName3 = ""
					end if
				else
					FItemList(iRows).FCateName2 = ""
				end if
				
				FItemList(iRows).FCateDepth	= arrData(5)

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
		Scn= "scn_dt_itemDispBrandGroup"		'일반 상품 검색
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = (FCurrPage -1)*FPageSize '// 검색시작 Row
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성

		'//그룹 범위별 지정(정렬 쿼리 생성)
		'SortQuery = " GROUP BY makerid order by brandname "
		SortQuery = " GROUP BY makerid order by count(*) desc"

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

		'response.write "group : " & SearchQuery & SortQuery & "<br>"
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
            CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT Sub
		END IF
		
		if (groupKeyCnt<1) then
		    SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
            CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
		    Exit Sub
		end if
		
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
		dim arrData,arrSize

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

		'response.write "group : " & SearchQuery & SortQuery & "<br>"
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
            CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT Sub
		END IF
		
		if (groupKeyCnt<1) then
		    SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
            CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
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
			'Scn= "scn_dt_itemColor"		'상품 컬러별 검색
			'Scn= "scn_dt_itemDispColor"		'상품 컬러별 검색(전시카테고리)
			Scn= "scn_dt_itemDisp"		'일반 상품 검색  //2015/03통일
		end if
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// 검색시작 Row
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성

		'IF (FRectCateCode<>"" and FRectSearchItemDiv<>"y") Then
		''    SortQuery = " GROUP BY itemid" ''2013 조건추가
		'else
    	''	SortQuery = " "	'// 정렬 쿼리 생성
    	'end if

		IF SearchQuery="" Then
			EXIT SUB
		End If

		dim Rowids,Scores

        '// 기본으로 검색2번 섭, 검색어가 있다면 1번섭 사용
        ''---------------------------------------------------------------------------------------------------------
        'if (FRectCateCode<>"") or (FRectMakerid<>"") and (FRectSearchTxt="")  then
        '    SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR) 
        'end if
        ''칼라칩검색/브랜드검색 3차
        'if (Scn= "scn_dt_itemDispColor") or (FRectMakerid<>"") then SvrAddr = getTimeChkAddr(G_3RDSCH_ADDR)
        ''---------------------------------------------------------------------------------------------------------

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

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수


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
		
		If FRectColorExclude <> "" Then
			SearchQuery = SearchQuery & FRectColorExclude
		End If
		'response.write SearchQuery & "<br>!"
		SortQuery = "Group by colorgrp Order by colorgrp "	'// 정렬 쿼리 생성

		IF SearchQuery="" Then
			EXIT SUB
		End If

		dim Rowids,Scores
		FTotalCount = 0
        ''---------------------------------------------------------------------------------------------------------
        'if (FRectCateCode<>"") or (FRectMakerid<>"") and (FRectSearchTxt="")  then
        '    SvrAddr = getTimeChkAddr(G_2NDSCH_ADDR)
        'end if
        ''칼라칩검색/브랜드검색 3차
        'if (Scn= "scn_dt_itemDispColor") or (FRectMakerid<>"") then SvrAddr = getTimeChkAddr(G_3RDSCH_ADDR)
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
            CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT Sub
		END IF
		
		if (groupKeyCnt<1) then            ''그룹 갯수는 1개;
		    SET groupKeyVal = NOTHING
            SET groupsize   = NOTHING
            CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
		    Exit Sub
		end if
		
		REDIM FItemList(FResultCount)
        dim bufColor, pos1
        
        FOR iRows = 0 to FResultcount -1
            bufColor = groupKeyVal(iRows,0)
            
       ''response.write "groupKeyVal:"&bufColor
            if (Len(bufColor)>=6) then
    			SET FItemList(iRows) = new SearchGroupByItems
				FItemList(iRows).FcolorCode		= mid(bufColor,4,3)
				FItemList(iRows).FcolorName		= getCdPosVal(FItemList(iRows).FcolorCode,G_KSCOLORCD,G_KSCOLORNM)                                  ''mid(bufColor,8,pos1)
				FItemList(iRows).FcolorIcon		= "http://webimage.10x10.co.kr/color/colorchip/" & "chip"&CLNG(FItemList(iRows).FcolorCode)&".gif"  ''mid(bufColor,pos1+1,255)

				FItemList(iRows).FSubTotal 	= groupsize(iRows)
				FTotalCount = FTotalCount + groupsize(iRows)
            end if

		NEXT
		
		SET groupKeyVal = NOTHING
        SET groupsize   = NOTHING
        
        
'		FOR iRows = 0 to FResultcount -1
'
'			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )
'
'			IF( ret < 0 ) THEN
'				'Response.write "GetResult_Row: " & Docruzer.msg
'				EXIT FOR
'			END IF
'
'			SET FItemList(iRows) = new SearchGroupByItems
'				FItemList(iRows).FcolorCode		= arrData(0)
'				FItemList(iRows).FcolorName		= arrData(1)
'				FItemList(iRows).FcolorIcon		= "http://webimage.10x10.co.kr/color/colorchip/" & arrData(2)
'
'				FItemList(iRows).FSubTotal 	= Scores(iRows)
'				FTotalCount = FTotalCount + Scores(iRows)
'
'			SET arrData = NOTHING
'			SET arrSize = NOTHING
'
'		NEXT
'
'		SET Rowids= NOTHING
'		SET Scores= NOTHING

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

	End SUB


	'####### 추천검색어  ######
	PUBLIC FUNCTION getRecommendKeyWords()

		Dim Docruzer,ret
		Dim iRows
		Dim arrData,arrSize
		Dim MaxCnt : MaxCnt =10

		if (Trim(FRectSearchTxt)="") then Exit FUNCTION 

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT FUNCTION
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106번으로 일단 G_2NDSCH_ADDR '' G_ORGSCH_ADDR

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

	'####### 인기검색어 (일반)  ######
	PUBLIC FUNCTION getPopularKeyWords()

		DIM Docruzer,ret
		DIM iRows
		DIM arrData,arrSize
		DIM MaxCnt : MaxCnt =FPageSize

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT FUNCTION
		END IF

        SvrAddr = G_ORGSCH_ADDR  '' 106번으로 일단 G_2NDSCH_ADDR '' G_ORGSCH_ADDR

		'인기 검색어 (일반)
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

        SvrAddr = G_ORGSCH_ADDR  '' 106번으로 일단 G_2NDSCH_ADDR '' G_ORGSCH_ADDR

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

	END Function
	
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
		
		If isapp = "1" Then
			SvrAddr = G_3RDSCH_ADDR '//app
		Else 
			SvrAddr = G_4THSCH_ADDR '//mobile
		End If 

		ret = Docruzer.getRealtimePopularKeyWord _
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

	PUBLIC FUNCTION getRealtimePopularKeyWordsAppOnly(byRef arDt, byRef arTg)

		DIM Docruzer,ret
		DIM iRows
		DIM arrData,arrSize, arrTags
		DIM MaxCnt : MaxCnt =FPageSize

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF
		
		SvrAddr = G_3RDSCH_ADDR '//app

		ret = Docruzer.getRealtimePopularKeyWord _
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
		sql = sql & " , i.itemname, i.sellcash, i.orgprice, i.sellyn, i.sailyn, i.limityn, i.itemcouponyn, i.itemcoupontype, i.itemcouponvalue " + vbcrlf
		sql = sql & " , i.listimage120, i.evalcnt, i.itemscore, e.File1, e.File2, i.icon1image, i.evalcnt, ic.favcount, i.tenOnlyYn  " + vbcrlf
		sql = sql & " FROM db_board.[dbo].tbl_item_evaluate e " + vbcrlf
		sql = sql & " JOIN  [db_item].[dbo].tbl_item i  " + vbcrlf
		sql = sql & " 	on e.itemid = i.itemid " + vbcrlf
		sql = sql & " JOIN  [db_item].[dbo].tbl_item_contents ic  " + vbcrlf
		sql = sql & " 	on i.itemid = ic.itemid " + vbcrlf
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

                ''FItemList(i).Fidx				= rsget("idx")
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
	'// 검색결과 이벤트 배너영역 수정
	Public Fevt_subname

End Class

Class SearchEventCls
    ''검색페이지 이벤트.
	Private SUB Class_initialize()

        '' 기본 1차서버--------------------------------------------------------------------------------------------
		SvrAddr = getTimeChkAddr(G_4THSCH_ADDR)
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
	dim FRectGubun

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

		if (Trim(FRectSearchTxt)="") then Exit Sub 

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
					
		If FRectGubun = "mktevt" Then
				SearchQuery = SearchQuery & " and idx_evt_kind = '28' "
		ElseIf FRectGubun = "all" Then
				SearchQuery = SearchQuery & " and (idx_evt_kind = '1' or idx_evt_kind = '19' or idx_evt_kind = '29' or idx_evt_kind = '13' or idx_evt_kind = '28') "
		Else FRectGubun = "planevt"
				SearchQuery = SearchQuery & " and (idx_evt_kind = '1' or idx_evt_kind = '19' or idx_evt_kind = '29' or idx_evt_kind = '13') "
		End If
'response.write SearchQuery
		Select Case FRectChannel
			Case "W"
				SearchQuery = SearchQuery & " and idx_evt_isWeb=1 "
			Case "M"
				SearchQuery = SearchQuery & " and idx_evt_isMobile=1 "
			Case "A"
				SearchQuery = SearchQuery & " and idx_evt_isApp=1 "
			Case Else
				SearchQuery = SearchQuery & " and idx_evt_isMobile=1 "
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
'Response.write "GetResult_Row: " & ret
		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF


		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)

		REDIM FItemList(FResultCount)

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수
'Response.write "GetResult_Row: " & FTotalCount
		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new SearchEventItems
				FItemList(iRows).Fevt_code			= arrData(0)
				FItemList(iRows).Fevt_bannerimg		= arrData(24) ''(pcWeb:1,M/A:24)
				FItemList(iRows).Fevt_startdate		= dateSerial(left(arrData(2),4),mid(arrData(2),5,2),right(arrData(2),2))	'### ASP 날짜형태로 변환
				FItemList(iRows).Fevt_enddate		= dateSerial(left(arrData(3),4),mid(arrData(3),5,2),right(arrData(3),2))	'### ASP 날짜형태로 변환
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
				'// 검색결과 이벤트 배너영역 수정
				FItemList(iRows).Fevt_subname		= arrData(28)

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
'### 브랜드 전체 리스트 검색   ###
'###############################

Class SearchBrandItems

	Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB

	PUBLIC Fuserid
	public Fkor_div
	public Feng_div
	PUBLIC Fsocname
	PUBLIC Fsocname_kor
	PUBLIC Frecommendcount
	PUBLIC Ftodayrecommendcount
	PUBLIC Fhitflg
	PUBLIC Fsaleflg
	PUBLIC Fsmileflg
	PUBLIC Fnewflg
	PUBLIC Fgiftflg
	PUBLIC Fonlyflg
	PUBLIC Fartistflg
	PUBLIC Fkdesignflg
	public Fsort_kor
	public Fsort_eng


End Class

Class SearchBrandCls
    ''검색페이지 이벤트.
	Private SUB Class_initialize()

        '' 기본 1차서버--------------------------------------------------------------------------------------------
		SvrAddr = getTimeChkAddr(G_4THSCH_ADDR)
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
	dim FRectWord
	dim FRectSort

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
    
    
	'####### 브랜드 전체리스트 검색 ######
	PUBLIC SUB getBrandList()

		dim Scn : Scn= "scn_dt_brandlist2017" 		'// 검색 결과 출력 시나리오명
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = (FCurrPage -1)*FPageSize	'// 검색시작 Row
		'// 검색 쿼리생성 (이벤트는 제외어 검색 하지말것 서버죽음 ㄷㄷ)

		If FRectWord <> "" Then
			SearchQuery = " idx_kor_div='" & FRectWord & "' "
		End If

		If FRectSearchTxt <> "" Then
			SearchQuery = " idx_brandName='" & FRectSearchTxt & "'  allword "	'키워드검색
		End If
		'SearchQuery = SearchQuery &_
		'			" and evt_state='7' " &_
		'			" and evt_startdate<='" & Replace(date(),"-","") & "000000' " &_
		'			" and evt_enddate>='" & Replace(date(),"-","") & "000000' "
'response.write SearchQuery
		'//그룹 범위별 지정(정렬 쿼리 생성)
		If FRectSort <> "" Then
			SortQuery = "Order by " & FRectSort & " asc "
		Else
			SortQuery = "Order by sort_kor asc "
		End If
'response.write SortQuery
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
			'Response.write "GetResult_Row: " & Docruzer.msg
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF


		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)

		REDIM FItemList(FResultCount)

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수
'response.write FResultCount
		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new SearchBrandItems
				FItemList(iRows).Fuserid					= arrData(0)
				FItemList(iRows).Fkor_div				= arrData(1)
				FItemList(iRows).Feng_div				= arrData(2)
				FItemList(iRows).Fsocname				= arrData(3)
				FItemList(iRows).Fsocname_kor			= arrData(4)
				FItemList(iRows).Frecommendcount		= arrData(5)
				FItemList(iRows).Ftodayrecommendcount	= arrData(6)
				FItemList(iRows).Fhitflg					= arrData(7)
				FItemList(iRows).Fsaleflg				= arrData(8)
				FItemList(iRows).Fsmileflg				= arrData(9)
				FItemList(iRows).Fnewflg					= arrData(10)
				FItemList(iRows).Fgiftflg				= arrData(11)
				FItemList(iRows).Fonlyflg				= arrData(12)
				FItemList(iRows).Fartistflg				= arrData(13)
				FItemList(iRows).Fkdesignflg			= arrData(14)
				FItemList(iRows).Fsort_kor				= arrData(15)
				FItemList(iRows).Fsort_eng				= arrData(16)

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
'### 플레잉 전체 리스트 검색   ###
'###############################

Class SearchPlayingItems

	Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB

	PUBLIC Fdidx
	public Ftitle
	public Fpc_bgcolor
	PUBLIC Fmo_bgcolor
	PUBLIC Fstate
	PUBLIC Fstartdate
	PUBLIC Ftitlestyle
	PUBLIC Fsubcopy
	PUBLIC Fkeyword
	PUBLIC Fimg28

End Class

Class SearchPlayingCls
    ''검색페이지 이벤트.
	Private SUB Class_initialize()

        '' 기본 1차서버--------------------------------------------------------------------------------------------
		SvrAddr = getTimeChkAddr(G_4THSCH_ADDR)
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
	dim FRectWord

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
    
    
	'####### 플레잉 전체리스트 검색 ######
	PUBLIC SUB getPlayingList2017()

		dim Scn : Scn= "scn_dt_playing2017" 		'// 검색 결과 출력 시나리오명
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		if Trim(FRectSearchTxt)="" then Exit Sub

		StartNum = (FCurrPage -1)*FPageSize	'// 검색시작 Row

		If FRectSearchTxt <> "" Then
			SearchQuery = " idx_playingKeyword='" & FRectSearchTxt & "'  allword "	'키워드검색
		End If
		SearchQuery = SearchQuery &_
					" and idx_state='7' " &_
					" and idx_startdate<='" & Replace(date(),"-","") & "000000' "
'response.write SearchQuery
		'//그룹 범위별 지정(정렬 쿼리 생성)
		SortQuery = "Order by didx desc "

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
			'Response.write "GetResult_Row: " & Docruzer.msg
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF


		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)

		REDIM FItemList(FResultCount)

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수
'response.write FResultCount
		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new SearchPlayingItems
				FItemList(iRows).Fdidx			= arrData(0)
				FItemList(iRows).Ftitle			= arrData(1)
				FItemList(iRows).Fpc_bgcolor	= arrData(2)
				FItemList(iRows).Fmo_bgcolor	= arrData(3)
				FItemList(iRows).Fstate			= arrData(4)
				FItemList(iRows).Fstartdate		= arrData(5)
				FItemList(iRows).Ftitlestyle	= arrData(6)
				FItemList(iRows).Fsubcopy		= arrData(7)
				FItemList(iRows).Fkeyword		= arrData(8)
				FItemList(iRows).Fimg28			= arrData(9)

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
'### PIECE 전체 리스트 검색   ###
'###############################

Class SearchPieceItems

	Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB

	PUBLIC Fidx
	PUBLIC Ffidx
	public Fgubun
	public Fbannergubun
	PUBLIC FnoticeYN
	PUBLIC Flistimg
	PUBLIC Flisttext
	PUBLIC Fshorttext
	PUBLIC Flisttitle
	PUBLIC Fadminid
	PUBLIC Fusertype
	PUBLIC Fetclink
	PUBLIC Fsnsbtncnt
	PUBLIC Foccupation
	PUBLIC Fnickname
	PUBLIC Fstartdate
	PUBLIC Fenddate
	PUBLIC Fisusing
	PUBLIC Fdeleteyn
	PUBLIC Ftagtext
	PUBLIC Fpitem
	PUBLIC Fpieceidx

End Class

Class SearchPieceCls
	Private SUB Class_initialize()

        '' 기본 1차서버--------------------------------------------------------------------------------------------
		SvrAddr = getTimeChkAddr(G_4THSCH_ADDR)
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
	dim FRectWord
	dim FRectSearchGubun
	dim FRectAdminID
	dim FRectIsOpening

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
    
    
	'####### PIECE 전체리스트 검색 ######
	PUBLIC SUB getPieceList2017()

		dim Scn : Scn= "scn_dt_piece2017" 		'// 검색 결과 출력 시나리오명
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime
		

		StartNum = (FCurrPage -1)*FPageSize	'// 검색시작 Row

		If FRectIsOpening <> "" Then
			SearchQuery = " idx_noticeYN='Y' "
		Else
			SearchQuery = " idx_noticeYN='N' "
		End If

		If FRectSearchTxt <> "" Then
			SearchQuery = " idx_pieceKeyword='" & FRectSearchTxt & "'  allword "	'키워드검색
		End If

		If FRectAdminID <> "" Then
			SearchQuery = " idx_adminid='" & FRectAdminID & "' "
		End If

		If FRectSearchGubun = "list" Then
			SearchQuery = SearchQuery & ""
		ElseIf FRectSearchGubun = "allsearch" Then
			SearchQuery = SearchQuery & " and idx_gubun = '1' "
		ElseIf FRectSearchGubun = "tagsearch" Then
			SearchQuery = SearchQuery & ""
		End If
		'// 오픈시에 날짜 주석 풀기
		SearchQuery = SearchQuery & " and idx_startdate<='" & Replace(date(),"-","") & ""& Num2Str(Hour(now),2,"0","R") & Num2Str(Minute(now),2,"0","R") & Num2Str(Second(now),2,"0","R") &"' "
'					" and idx_state='7' " &_
'response.write SearchQuery
		'//그룹 범위별 지정(정렬 쿼리 생성)
		SortQuery = "Order by startdate desc "
		'SortQuery = "Order by idx desc "

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
			'Response.write "GetResult_Row: " & Docruzer.msg
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF


		Call Docruzer.GetResult_RowSize(FResultcount) '검색 결과 수
		Call Docruzer.GetResult_Rowid(Rowids,Scores)

		REDIM FItemList(FResultCount)

		Call Docruzer.GetResult_TotalCount(FTotalCount) '검색결과 총 수
'response.write FResultCount
		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new SearchPieceItems
				FItemList(iRows).Fidx			= arrData(0)
				FItemList(iRows).Ffidx			= arrData(1)
				FItemList(iRows).Fgubun			= arrData(2)
				FItemList(iRows).Fbannergubun	= arrData(3)
				FItemList(iRows).FnoticeYN		= arrData(4)
				FItemList(iRows).Flistimg		= arrData(5)
				FItemList(iRows).Flisttext		= arrData(6)
				FItemList(iRows).Fshorttext		= arrData(7)
				FItemList(iRows).Flisttitle		= arrData(8)
				FItemList(iRows).Fadminid		= arrData(9)
				FItemList(iRows).Fusertype		= arrData(10)
				FItemList(iRows).Fetclink		= arrData(11)
				FItemList(iRows).Fsnsbtncnt		= arrData(12)
				FItemList(iRows).Foccupation	= arrData(13)
				FItemList(iRows).Fnickname		= arrData(14)
				FItemList(iRows).Fstartdate		= arrData(15)
				FItemList(iRows).Fenddate		= arrData(16)
				FItemList(iRows).Fisusing		= arrData(17)
				FItemList(iRows).Fdeleteyn		= arrData(18)
				FItemList(iRows).Ftagtext		= arrData(19)
				FItemList(iRows).Fpitem			= arrData(20)
				FItemList(iRows).Fpieceidx		= arrData(21)

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

'// 내 위시 상품 목록(검색 결과에서 상품목록 전송)
Sub getMyFavItemList(uid,iid,byRef sWArr)
  'Exit Sub ''사용안함 2014/09/23
	dim strSQL
	strSQL = "execute [db_my10x10].[dbo].[sp_Ten_MyWishSearchItemNew] '" & CStr(uid) & "', '" & cStr(iid) & "'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSQL,dbget,adOpenForwardOnly,adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		sWArr = rsget.getRows()
	end if
	rsget.Close
end Sub

Function fnIsMyFavItem(arr,itemid)
Dim i, r
	r = False
	If isArray(arr) Then
		For i=0 To UBound(arr,2)
			If InStr((","&arr(0,i)&","),(","&itemid&",")) > 0 Then
				r = True
				Exit For
			End If
		Next
	End If
	fnIsMyFavItem = r
End Function

'// 내검색어 저장처리 (쿠키, 최근 10개 저장)
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
			if i>9 then Exit For
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

Function fnGetSearchEvent(Kywd)
	Dim oGrEvt, i, result, ename
	'// 이벤트 검색결과
	set oGrEvt = new SearchEventCls
	oGrEvt.FRectSearchTxt = Kywd
	oGrEvt.FRectChannel = "A"		'검색 채널 (W:isWeb, M:isMobile, A:isApp)
	oGrEvt.FCurrPage = 1
	oGrEvt.FPageSize = 1
	oGrEvt.FScrollCount =10
	oGrEvt.FRectGubun = "all"
	oGrEvt.getEventList

	if oGrEvt.FResultCount>0 then

		FOR i = 0 to oGrEvt.FResultCount-1
			
			If oGrEvt.FItemList(i).Fevt_kind = "28" Then
				result = "이벤트$$"
			Else
				result = "기획전$$"
			End If
			
			ename = split(db2html(oGrEvt.FItemList(i).Fevt_name),"|")(0) & "$$"
			
			result = result & ename & oGrEvt.FItemList(i).Fevt_code
			
		Next
	End If

	Set oGrEvt = nothing
	fnGetSearchEvent = result
End function
%>