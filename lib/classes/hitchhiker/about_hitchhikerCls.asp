<%
'###########################################################
' Description : 히치하이커 이슈영역 클래스
' Hieditor : 2014.07.24 유태욱 생성
'###########################################################
%>
<%
function getHitchhikerGubun(v)
	if v = 1 then
		getHitchhikerGubun = "발간"
	elseif v = 2 then
		getHitchhikerGubun = "에디터모집"
	else
		getHitchhikerGubun = "기타"
	end if
end function

class CHitchhikerItem
	public Fidx
	public FReqTitle
	public FReqIsusing
	public FReqSortnum
	public FReqSdate
	public FReqEdate
	public FreqRegdate
	public Freqgubun
	public Freqimghtmltext
end class

class CAbouthitchhiker
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FDevice
	public FEvt_title
	public FIsusing
	public Fgubun
	public FValiddate
	
	'###### 히치하이커 이슈영역리스트 ######
	public sub fnGetHitchhikerList
		dim sqlStr,i, sqlsearch

		if Fgubun <> "" Then
			sqlsearch = sqlsearch & " AND gubun = '"& Fgubun &"'"
		end if
				
		if FEvt_title <> "" Then
			sqlsearch = sqlsearch & " AND hic_title like'%"& FEvt_title &"%'"
		end if

		if FIsusing <> "" Then
			sqlsearch = sqlsearch & " AND isusing ='"& FIsusing &"'"
		end if
		
        if FValiddate<>"" then
            sqlsearch = sqlsearch + " AND edate > getdate()"
        end if

		'글의 총 갯수 구하기
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_list"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'DB 데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx, hic_title, isusing, sortnum, sdate, edate, regdate, gubun, imghtmltext"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_list"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by sortnum asc ,idx Desc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize		
		rsget.Open sqlStr,dbget,1

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
				set FItemList(i) = new CHitchhikerItem
					FItemList(i).Fidx = rsget("idx")
					FItemList(i).FReqTitle = rsget("hic_title")
					FItemList(i).FReqIsusing = rsget("isusing")
					FItemList(i).FReqSortnum = rsget("sortnum")
					FItemList(i).FReqSdate = rsget("sdate")
					FItemList(i).FReqEdate = rsget("edate")
					FItemList(i).FreqRegdate = rsget("regdate")
					FItemList(i).Freqgubun = rsget("gubun")
					FItemList(i).Freqimghtmltext = rsget("imghtmltext")
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






	

		