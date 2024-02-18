<%
'###########################################################
' Description : 아카데미 모바일 메인배너 클래스
' Hieditor : 2016.07.29 유태욱 생성
'###########################################################
%>
<%
function getAcademyMobileMainBannerGubun(v)
	if v = 1 then
		getAcademyMobileMainBannerGubun = "강좌링크"

	elseif v = 2 then
		getAcademyMobileMainBannerGubun = "상품링크"

	elseif v = 3 then
		getAcademyMobileMainBannerGubun = "매거진링크"

	elseif v = 4 then
		getAcademyMobileMainBannerGubun = "강사/작가 링크"

	elseif v = 5 then
		getAcademyMobileMainBannerGubun = "기타 링크"

	else
		getAcademyMobileMainBannerGubun = "강좌링크"
	end if
end function

class CAcademyMobileMainBannerItem
	public Fidx
	public FReqSdate
	public FReqEdate
	public Freqgubun
	public FReqIsusing
	public FReqSortnum	
	public FReqlinknum
	public FReqmakerid
	public FreqRegdate
	public Freqlinkurl_etc
	public FReqcon_viewthumbimg
end class

class CAcademyMobileMainBanner
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
	public FRectSearchSDate
	public FRectSearchEDate
	
	'###### 아카데미 모바일 메인배너 리스트 ######
	public sub fnGetAcademyMobileMainBannerList
		dim sqlStr,i, sqlsearch

		if Fgubun <> "" Then
			sqlsearch = sqlsearch & " AND gubun = '"& Fgubun &"'"
		end if

		if FIsusing <> "" Then
			sqlsearch = sqlsearch & " AND isusing ='"& FIsusing &"'"
		end if
		


		if FRectSearchSDate<>"" Then
			sqlsearch = sqlsearch & "  AND sdate >= '" & FRectSearchSDate & "'" & vbcrlf
		end If
		
		if FRectSearchEDate<>"" Then
			sqlsearch = sqlsearch & "  AND edate <= '" & FRectSearchEDate & "'" & vbcrlf
		end If

		'글의 총 갯수 구하기
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_mobile_mainbanner_list"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
'response.write sqlStr
'response.end
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		'DB 데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx, linknum, linkurl_etc, sdate, edate, isusing, sortnum, regdate, gubun, con_viewthumbimg, makerid"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_academy_mobile_mainbanner_list"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by idx Desc"
		
		'response.write sqlStr &"<br>"
		rsACADEMYget.pagesize = FPageSize		
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

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
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new CAcademyMobileMainBannerItem
					FItemList(i).Fidx					= rsACADEMYget("idx")
					FItemList(i).FReqSdate				= rsACADEMYget("sdate")
					FItemList(i).FReqEdate				= rsACADEMYget("edate")
					FItemList(i).Freqgubun				= rsACADEMYget("gubun")
					FItemList(i).FReqIsusing			= rsACADEMYget("isusing")
					FItemList(i).FreqRegdate			= rsACADEMYget("regdate")
					FItemList(i).FReqSortnum			= rsACADEMYget("sortnum")
					FItemList(i).FReqlinknum			= rsACADEMYget("linknum")
					FItemList(i).FReqmakerid			= rsACADEMYget("makerid")
					FItemList(i).FReqlinkurl_etc		= rsACADEMYget("linkurl_etc")
					FItemList(i).FReqcon_viewthumbimg	= rsACADEMYget("con_viewthumbimg")
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
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






	

		