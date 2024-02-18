<%
'###########################################################
'	Description : ��ǰ�� ��� �귣�� Ŭ����
'	History		: 2017.01.20 ���¿� ����
'###########################################################
%>
<%
function getBrandNoticeGubun(v)
	if v = 1 then
		getBrandNoticeGubun = "�Ϲݰ���"
	elseif v = 2 then
		getBrandNoticeGubun = "��۰���"
	elseif v = 3 then 
		getBrandNoticeGubun = "��Ÿ����"
	else
		getBrandNoticeGubun = "�Ϲݰ���"
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
	
	'###### �귣�� ���� ����Ʈ ######
	public sub fnGetBrandNoticeList
		dim sqlStr,i, sqlsearch

		if Fgubun <> "" Then
			sqlsearch = sqlsearch & " AND gubun = '"& Fgubun &"'"
		end if

		if FIsusing <> "" Then
			sqlsearch = sqlsearch & " AND isusing ='"& FIsusing &"'"
		end if

		if Fbrandidtext <> "" Then
			''sqlsearch = sqlsearch & " AND brandid like '%"& Fbrandidtext &"%'"  ''2017/10/27 by eastone  ��ü���ο��� �ٸ� �귣�尡 ����.
			sqlsearch = sqlsearch & " AND brandid = '"& Fbrandidtext &"'"
		end if
		
        if FValiddate<>"" then
            sqlsearch = sqlsearch + " AND (edate > getdate() or infiniteregyn='Y') "
        end if

		'���� �� ���� ���ϱ�
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
		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'DB ������ ����Ʈ
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






	

		