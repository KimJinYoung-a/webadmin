<%
'==========================================================================
'	Description: 매장날씨 관리 클래스
'	History: 2012.06.04 강준구 생성
'			 2012.06.12 한용민 수정(페이징수정)
'==========================================================================

Class COffShopWeatherItem
	public FCount
	public FIdx
	public FWDate
	public FShopID
	public FShopName
	public FWeather
	public FComment


	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class


class COffShopWeather
	public FItemList()
	public FCountList()
	public FPageCount
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectIdx
	public FRectShopID
	public FRectWSDate
	public FRectWEDate
	public FIdx
	public FWDate
	public FShopID
	public FShopName
	public FWeather
	public FComment
	
	'//admin/offshop/weather/index.asp
	public sub GetOffShopWeatherList()
		dim sqlStr,i ,sqlsearch

		if FRectShopID <> "" then
			sqlsearch = sqlsearch & " and w.shopid = '" & FRectShopID & "' "
		end if
		
		if FRectWSDate <> "" then
			sqlsearch = sqlsearch & " and w.wdate >= '" & FRectWSDate & "' "
		end if
		
		if FRectWEDate <> "" then
			sqlsearch = sqlsearch & " and w.wdate <= '" & FRectWEDate & "' "
		end if
		
		'총 갯수 구하기
		sqlStr = "SELECT COUNT(*) as cnt FROM [db_shop].[dbo].[tbl_shop_weather] AS w "
		sqlStr = sqlStr & "WHERE 1=1 " & sqlsearch
								
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub
			
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & "	w.idx, w.wdate, w.shopid, w.weather, w.comment, u.shopname FROM [db_shop].[dbo].[tbl_shop_weather] AS w "
		sqlStr = sqlStr & "	LEFT JOIN [db_shop].[dbo].[tbl_shop_user] AS u ON w.shopid = u.userid "
		sqlStr = sqlStr & "	WHERE 1=1 " & sqlsearch
		sqlStr = sqlStr & "	ORDER BY w.wdate DESC "

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
				set FItemList(i) = new COffShopWeatherItem
				
				FItemList(i).FIdx 		= rsget("idx")
				FItemList(i).FWDate  	= rsget("wdate")
				FItemList(i).FShopID   	= rsget("shopid")
				FItemList(i).FShopName  = rsget("shopname")
				FItemList(i).FWeather  	= "/images/weather/" & rsget("weather") & ".gif"
				FItemList(i).FComment  	= db2html(rsget("comment"))	
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	public Sub GetOffShopWeatherView
		dim i,sqlStr
		sqlStr = "SELECT w.idx, w.wdate, w.shopid, w.weather, w.comment, u.shopname FROM [db_shop].[dbo].[tbl_shop_weather] AS w "
		sqlStr = sqlStr & "		LEFT JOIN [db_shop].[dbo].[tbl_shop_user] AS u ON w.shopid = u.userid "
		sqlStr = sqlStr & "	WHERE w.idx = '" & FRectIdx & "'"
		rsget.Open sqlStr,dbget,1
		If Not rsget.Eof Then
			FIdx 		= rsget("idx")
			FWDate  	= rsget("wdate")
			FShopID   	= rsget("shopid")
			FShopName  	= rsget("shopname")
			FWeather  	= rsget("weather")
			FComment  	= db2html(rsget("comment"))
		End IF
		rsget.Close
	end Sub


	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		redim  FCountList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
end Class
%>