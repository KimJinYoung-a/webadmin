<%
'###########################################################
' Description : 히치하이커 메인배너 클래스
' Hieditor : 2014.07.24 유태욱 생성
'###########################################################
%>
<%
function getHitchhikerGubun(v)
	if v = 1 then
		getHitchhikerGubun = "메인상단롤링배너_링크"
	elseif v = 2 then
		getHitchhikerGubun = "메인상단롤링배너_레이어팝업"
	elseif v = 3 then
		getHitchhikerGubun = "메인상단롤링배너_OnlyView"
	else
		getHitchhikerGubun = "메인상단롤링배너_모집&발간"
	end if
end function

class CHitchhikerItem
	public Fidx
	public FReqIsusing
	public FReqSortnum
	public FReqSdate
	public FReqEdate
	public FreqRegdate
	public Freqgubun
	public Freqlinkurl
	public FReqcon_viewthumbimg
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
	
	public Fgubun
	public FIsusing
	public FValiddate
	
	'###### 히치하이커 메인배너 리스트 ######
	public sub fnGetHitchhikerList
		dim sqlStr,i, sqlsearch

		if Fgubun <> "" Then
			sqlsearch = sqlsearch & " AND gubun = '"& Fgubun &"'"
		end if

		if FIsusing <> "" Then
			sqlsearch = sqlsearch & " AND isusing ='"& FIsusing &"'"
		end if
		
        if FValiddate<>"" then
            sqlsearch = sqlsearch + " AND edate > getdate()"
        end if

		'글의 총 갯수 구하기
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_mainbanner_list"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'DB 데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx, linkurl, sdate, edate, isusing, sortnum, regdate, gubun, con_viewthumbimg"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_hitchhiker_mainbanner_list"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx Desc"
		
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
					FItemList(i).FReqSdate = rsget("sdate")
					FItemList(i).FReqEdate = rsget("edate")
					FItemList(i).Freqgubun = rsget("gubun")
					FItemList(i).FReqIsusing = rsget("isusing")
					FItemList(i).FreqRegdate = rsget("regdate")
					FItemList(i).FReqSortnum = rsget("sortnum")
					FItemList(i).FReqlinkurl = rsget("linkurl")
					FItemList(i).FReqcon_viewthumbimg = rsget("con_viewthumbimg")
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






	

		