<%
'###########################################################
'	Description : 상품상세 상단 브랜드 클래스
'	History		: 2017.01.20 유태욱 생성
'###########################################################
%>
<%
function getBrandNoticeGubun(v)
	if v = 1 then
		getBrandNoticeGubun = "일반공지"
	elseif v = 2 then
		getBrandNoticeGubun = "배송공지"
	elseif v = 3 then 
		getBrandNoticeGubun = "기타공지"
	else
		getBrandNoticeGubun = "일반공지"
	end if
end function

class CBrandNoticeItem
	public Fidx
	public Frank
	public FReqSdate
	public FReqEdate
	public Freqgubun
	public FReqBrandid
	public FReqIsusing
	public FReqmakerid
	public FreqRegdate
	public FReqnotice_text
	public FReqnotice_title
	public Freqinfiniteregyn
end class

class CBrandNotice
	public FItemList()
	public FCurrPage
	public FPageSize
	public FTotalPage
	public FPageCount
	public FTotalCount
	public FResultCount
	public FScrollCount
	
	public Fgubun
	public FIsusing
	public FValiddate
	public Fbrandidtext
	
	'###### 브랜드 공지 리스트 ######
	public sub fnGetBrandNoticeList
		dim sqlStr,i, sqlsearch

		if Fgubun <> "" Then
			sqlsearch = sqlsearch & " AND gubun = '"& Fgubun &"'"
		end if

		if FIsusing <> "" Then
			sqlsearch = sqlsearch & " AND isusing ='"& FIsusing &"'"
		end if

		if Fbrandidtext <> "" Then
			''sqlsearch = sqlsearch & " AND brandid like '%"& Fbrandidtext &"%'"  ''2017/10/27 by eastone  업체어드민에서 다른 브랜드가 보임.
			sqlsearch = sqlsearch & " AND brandid = '"& Fbrandidtext &"'"
		end if
		
        if FValiddate<>"" then
            sqlsearch = sqlsearch + " AND (edate > getdate() or infiniteregyn='Y') "
        end if

		'글의 총 갯수 구하기
		sqlStr = "select count(*) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sqlStr = sqlStr & " from db_board.dbo.tbl_brand_notice_list with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		if FTotalCount < 1 then exit sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'DB 데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx, sdate, edate, isusing, regdate, gubun, makerid, brandid, infiniteregyn, notice_title, notice_text, Rank() over (partition by brandid,gubun,isusing order by idx desc) as rank"
		sqlStr = sqlStr & " from db_board.dbo.tbl_brand_notice_list with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx Desc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize		
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CBrandNoticeItem
					FItemList(i).Fidx					= rsget("idx")
					FItemList(i).Frank					= rsget("rank")
					FItemList(i).FReqSdate				= rsget("sdate")
					FItemList(i).FReqEdate				= rsget("edate")
					FItemList(i).Freqgubun				= rsget("gubun")
					FItemList(i).FReqIsusing			= rsget("isusing")
					FItemList(i).FreqRegdate			= rsget("regdate")
					FItemList(i).FReqmakerid			= rsget("makerid")
					FItemList(i).FReqbrandid			= rsget("brandid")
					FItemList(i).FReqnotice_title		= rsget("notice_title")
					FItemList(i).FReqnotice_text		= rsget("notice_text")
					FItemList(i).FReqinfiniteregyn	= rsget("infiniteregyn")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub	

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	
	Private Sub Class_Terminate()
	End Sub
	
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
%>






	

		