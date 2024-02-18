<%
'###########################################################
' Description :  SCM 우편번호 클래스
' History : 2016.07.01 한용민 프론트 이전 생성
'###########################################################
%>
<%
''CS_EUCKR <=> CS_UTF8
''group by 필드는 널이 없게 하자.
'''--------------------------------------------------------------------------------------

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
     G_1STSCH_ADDR = "192.168.0.210"        ''192.168.0.210  :: (X)
     G_2NDSCH_ADDR = "192.168.0.207"        ''192.168.0.207  :: 
     G_3RDSCH_ADDR = "192.168.0.209"        ''192.168.0.209  :: (X)
     G_4THSCH_ADDR = "192.168.0.208"        ''192.168.0.208  :: 
     G_ORGSCH_ADDR = "192.168.0.206"        ''192.168.0.206
END IF

GG_1STSCH_ADDR = G_1STSCH_ADDR
GG_2NDSCH_ADDR = G_2NDSCH_ADDR
GG_3RDSCH_ADDR = G_3RDSCH_ADDR
GG_4THSCH_ADDR = G_4THSCH_ADDR
GG_ORGSCH_ADDR = G_ORGSCH_ADDR

if (Application("G_ZIPSCH_ADDR")="") or (Application("G_ZIPSCH_ADDR")=GG_1STSCH_ADDR) or (Application("G_ZIPSCH_ADDR")=GG_3RDSCH_ADDR) then
	Application("G_ZIPSCH_ADDR")=G_ORGSCH_ADDR
end if
G_ORGSCH_ADDR = Application("G_ZIPSCH_ADDR")

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

Class CZipPrdItem
    Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB
	
	PUBLIC Fidx
    PUBLIC Fzipcode
    PUBLIC Fsido
    PUBLIC Fgungu
    PUBLIC Feupmyun
    PUBLIC Froad
    PUBLIC Fisunder
    PUBLIC Fbuilding_no
    PUBLIC Fbuilding_sub
    PUBLIC Fofficial_bld
    PUBLIC Fdong
    PUBLIC Fri
    PUBLIC Fdong_admin
    PUBLIC Fissan
    PUBLIC Fjibun_main
    PUBLIC Fjibun_sub
    PUBLIC Fold_zipcode
    PUBLIC Fzipgroup   '2016/07/14 left(zipcode,2)
    
End Class
	
Class CGroupbySidoItem

	Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB

	PUBLIC Fzipgroup
	PUBLIC Fsido
	PUBLIC Fgungu
	PUBLIC FCNT
	
End Class



Class SearchItemCls

	Private SUB Class_initialize()
        ''기본 1차 서버.------------------------
		SvrAddr = getTimeChkAddr(G_ORGSCH_ADDR)
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
	dim FRectZipgroup               '' 그루핑코드
	
	dim FRectSearchSiDo 	'SiDo
	dim FRectSearchGunGu	'gungu
	dim FRectSearchRoad     'Road
	dim FRectSearchBuilding_no  'Building_no
	dim FRectSearchDong         'Dong
	dim FRectSearchJibun_main   'Jibun_main
	dim FRectSearchJibun_sub    'Jibun_sub
	
	dim FRectPrevSearchTxt	'이전 검색어
	
	dim FRectExceptText		'제외어
	
	dim FRectSortMethod		'정렬방식 

	dim FCheckResearch 		'결과내 재검색 체크용
	dim FRectColsSize		'결과 리스트 열수
	dim FLogsAccept			'추가 로그 저장 여부


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

		'### 검색조건 생성 ###

		'@ 검색어(키워드)
		IF FRectSearchTxt<>"" Then
			FRectSearchTxt = chgCoinedKeyword(FRectSearchTxt)
			FRectSearchTxt = escapeQuery(FRectSearchTxt)  ''2015 추가
		''FRectSearchTxt=replace(FRectSearchTxt," "," & ")
		''rw FRectSearchTxt
			IF FRectExceptText<>"" Then
			    FRectExceptText = escapeQuery(FRectExceptText)  ''2015 추가
				strQue = getQrCon(strQue) & "(idx_searchtxt='" & FRectSearchTxt & " ! " & FRectExceptText & "' BOOLEAN) "	'제외어
			else
				strQue = getQrCon(strQue) & "idx_searchtxt='" & FRectSearchTxt & "'  allword synonym"	'키워드검색(동의어 포함) synonym
				'strQue = getQrCon(strQue) & "idx_searchtxt='" & FRectSearchTxt & "'  BOOLEAN "
				'strQue = getQrCon(strQue) & "idx_searchtxt like '" & FRectSearchTxt & "*'   "
			End if
		END IF

        '@idx_zipgroup
        IF FRectZipgroup<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_zipgroup='"&FRectZipgroup&"' "
		End IF
		
		'@ SiDo
		IF FRectSearchSiDo<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_sido='"&FRectSearchSiDo&"' "
		End IF
		
		'@ GunGu
		IF FRectSearchGunGu<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_gungu='"&FRectSearchGunGu&"' "
		End IF
		
        '@ Road
		IF FRectSearchRoad<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_road='"&FRectSearchRoad&"' "
		End IF
		
		'@ Building_no
		IF FRectSearchBuilding_no<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_building_no='"&FRectSearchBuilding_no&"' "
		End IF

        '@ Dong
		IF FRectSearchDong<>"" Then
			strQue = strQue & getQrCon(strQue) & "dong='"&FRectSearchDong&"' "
		End IF
		
		'@ Jibun_main
		IF FRectSearchJibun_main<>"" Then
			strQue = strQue & getQrCon(strQue) & "jibun_main='"&FRectSearchJibun_main&"' "
		End IF
		
		'@ Jibun_sub
		IF FRectSearchJibun_sub<>"" Then
			strQue = strQue & getQrCon(strQue) & "jibun_sub='"&FRectSearchJibun_sub&"' "
		End IF
		
        
		query = strQue
	End FUNCTION

	Sub getSortQuery(byref query)
		dim strQue

        strQue = ""
        
		'// 정렬
		IF FRectSortMethod="zipcode" THEN 'zipcode
			strQue = strQue & " ORDER BY zipcode"
		ELSEIF FRectSortMethod="sido" THEN 'sido
			strQue = strQue & " ORDER BY sido"
		ELSEIF FRectSortMethod="idx" THEN 'idx
			strQue = strQue & " ORDER BY idx"
		ELSE
		    strQue = strQue & " ORDER BY $RELEVANCE DESC"  ''적합도순
		END IF
		
		query = strQue
	End Sub

	Function getQrCon(query)
		if Not(query="" or isNull(query)) then
			getQrCon = " and "
		end if
	End Function

	'####### 우편번호  검색 - 검색 엔진 ######
	PUBLIC SUB getSearchList()

		DIM Scn
		DIM Docruzer,ret

		DIM Logs ,iRows
		DIM arrData,arrSize, retMatchCd, retMatchVal
		Dim iDocErrMsg

		'// 검색 결과 출력 시나리오명
		Scn= "scn_dt_zipcode"		    

		StartNum = (FCurrPage -1)*FPageSize '// 검색시작 Row

		CALL getSearchQuery(SearchQuery)	'// 검색 쿼리생성
		CALL getSortQuery(SortQuery)		'// 정렬 쿼리 생성
		''Response.Write SearchQuery &"<Br>"
		IF SearchQuery="" THEN
			EXIT SUB
		END IF
		

		IF (FALSE) and (FLogsAccept) and (FRectSearchTxt<>"") and (FCurrPage="1") THEN
		    
            'Logs = "상품+^" & FRectSearchTxt & "]##" & FRectSearchTxt & "||" & FRectPrevSearchTxt  	'// 로그값
            ''2015 search4
            '기본:[사이트@카테고리+사용자$성별코드|연령|검색어타입(서비스)|첫검색|페이지번호|정렬순^이전검색어##검색어] ''기본
            Dim iLOG_SITE : iLOG_SITE = "ZIP"
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
		    IF (application("Svr_Info")="Dev") THEN
    		    rw "err1"
    		    rw SearchQuery
    		    rw Docruzer.GetErrorMessage()
    		END IF
    		
			'dbget.execute "EXECUTE db_log.dbo.sp_Ten_DocLog @ErrMsg ='"& html2db(SearchQuery) & "[" & html2db(Docruzer.GetErrorMessage()) &"]'"

			dbget.execute "EXECUTE db_log.dbo.sp_Ten_DocLog @ErrMsg ='["&SvrAddr&"]"& html2db(Docruzer.GetErrorMessage())&"["&Request.ServerVariables("REMOTE_ADDR")&"]["&Request.ServerVariables("LOCAL_ADDR")&"]["& html2db(SearchQuery)&"]'"

			iDocErrMsg = Docruzer.GetErrorMessage()
            if (InStr(iDocErrMsg,"recv queue full")>0) or (InStr(iDocErrMsg,"socket time out")>0) or (InStr(iDocErrMsg,"cannot connect to server")>0) or (InStr(iDocErrMsg,"scn_dt_zipcode' doesn't exist")>0) then
                IF (SvrAddr=Application("G_ZIPSCH_ADDR")) then
                    if (Application("G_ZIPSCH_ADDR")= GG_ORGSCH_ADDR) then
                        Application("G_ZIPSCH_ADDR") = GG_2NDSCH_ADDR
                    elseif (Application("G_ZIPSCH_ADDR")=GG_2NDSCH_ADDR) then
                        Application("G_ZIPSCH_ADDR") = GG_4THSCH_ADDR
					elseif (Application("G_ZIPSCH_ADDR")=GG_4THSCH_ADDR) then
                        Application("G_ZIPSCH_ADDR") = GG_ORGSCH_ADDR
                    end if
                end if
            end if

			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING

'				'// 1번 서버 에러시 2번에서 구동(2번도 에러면 Skip)
'				if (SvrAddr = G_1STSCH_ADDR) then
'					SvrAddr = G_2NDSCH_ADDR  ''"192.168.0.108"
'					if (G_1STSCH_ADDR<>G_2NDSCH_ADDR) then  ''추가 2013/09
'					    call getSearchList()
'				    end if
'				end if

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

			SET FItemList(iRows) = NEW CZipPrdItem
                
                FItemList(iRows).Fidx           = arrData(0)
                FItemList(iRows).Fzipcode       = arrData(1)
                FItemList(iRows).Fsido          = arrData(2)
                FItemList(iRows).Fgungu         = arrData(3)
                FItemList(iRows).Feupmyun       = arrData(4)
                FItemList(iRows).Froad          = arrData(5)
                FItemList(iRows).Fisunder       = arrData(6)
                FItemList(iRows).Fbuilding_no   = arrData(7)
                FItemList(iRows).Fbuilding_sub  = arrData(8)
                FItemList(iRows).Fofficial_bld  = arrData(9)
                FItemList(iRows).Fdong          = arrData(10)
                FItemList(iRows).Fri            = arrData(11)
                FItemList(iRows).Fdong_admin    = arrData(12)
                FItemList(iRows).Fissan         = arrData(13)
                FItemList(iRows).Fjibun_main    = arrData(14)
                FItemList(iRows).Fjibun_sub     = arrData(15)
                FItemList(iRows).Fold_zipcode   = arrData(16)
                
			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

	'####### 검색 그룹  카운팅  ######

		

	'####### 상품 검색 브랜드별 카운팅  ######
	PUBLIC SUB getGroupbySido()

		'// 검색 결과 출력 시나리오명
		
		Scn= "scn_dt_zipcodeGroup"
		
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// 검색시작 Row
		call getSearchQuery(SearchQuery)	'// 검색 쿼리생성

		'//그룹 범위별 지정(정렬 쿼리 생성)
		
		SortQuery = " GROUP BY zipgroup order by zipgroup asc " ''desc $RELEVANCE
		
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

			SET FItemList(iRows) = new CGroupbySidoItem
				FItemList(iRows).Fzipgroup	= arrData(0)
				FItemList(iRows).Fsido		= arrData(1)
				FItemList(iRows).Fgungu	    = arrData(2)
				FItemList(iRows).FCNT 	    = Scores(iRows)

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

End Class



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
