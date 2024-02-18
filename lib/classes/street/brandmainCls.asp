<%
Class cBrandItem
	Public FIdx
	Public FMakerid			
	Public FImagepath		
	Public FLinkpath		
	Public FIsusing		
	Public FRegdate		
	Public FGubun			
	Public FImage_order
	Public Fbrandimage
	Public Fadminid
	Public Flastupdate
	Public Flastadminid
End Class

Class cBrandMain
	Public FItemList()
	Public FOneItem
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount
	Public FRectIdx
	Public FRectGubun
	Public FRectIsUsing
	Public FRectMakerid

	Public Sub sMainTop3modify
		Dim sqlStr, i, sqlsearch
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 "
		sqlStr = sqlStr & " idx, makerid, imagepath, linkpath, isusing, regdate, gubun, image_order"
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_2013brand_image "
		sqlStr = sqlStr & " WHERE idx = '"&FRectIdx&"' "
		rsget.Open sqlStr, dbget, 1
		ftotalcount = rsget.recordcount
        SET FOneItem = new cBrandItem
	        If Not rsget.Eof then
				FOneItem.FIdx			= rsget("idx")
				FOneItem.FMakerid		= rsget("makerid")
				FOneItem.FImagepath		= rsget("imagepath")
				FOneItem.FLinkpath		= rsget("linkpath")
				FOneItem.FIsusing		= rsget("isusing")
				FOneItem.FRegdate		= rsget("regdate")
				FOneItem.FGubun			= rsget("gubun")
				FOneItem.FImage_order	= rsget("image_order")
        	End If
        rsget.Close
	End Sub
	
	Public Sub sMainTop3List
		If FRectGubun <> "" Then
			sqladd = sqladd & " and gubun = '"&FRectGubun&"' "
		End If

		Dim sqlStr, i, sqladd
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_2013brand_image "
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, imagepath, linkpath, isusing, regdate, gubun, image_order "
		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_2013brand_image "
		sqlStr = sqlStr & " WHERE 1=1 " & sqladd
		sqlStr = sqlStr & " ORDER BY image_order ASC, idx DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new cBrandItem
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FImagepath		= rsget("imagepath")
					FItemList(i).FLinkpath		= rsget("linkpath")
					FItemList(i).FIsusing		= rsget("isusing")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).FGubun			= rsget("gubun")
					FItemList(i).FImage_order	= rsget("image_order")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// 브랜드 이미지 수정
	Public Sub sBrandImageGetOne
		Dim sqlStr, i, sqlsearch
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 "
		sqlStr = sqlStr & " idx, idx , makerid , brandimage , isusing"
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_brand_image "
		sqlStr = sqlStr & " WHERE idx = '"&FRectIdx&"' "
		rsget.Open sqlStr, dbget, 1
		ftotalcount = rsget.recordcount
        SET FOneItem = new cBrandItem
	        If Not rsget.Eof then
				FOneItem.FIdx			= rsget("idx")
				FOneItem.FMakerid		= rsget("makerid")
				FOneItem.Fbrandimage	= rsget("brandimage")
				FOneItem.FIsusing		= rsget("isusing")
        	End If
        rsget.Close
	End Sub

	'// 브랜드 이미지 리스트
	Public Sub sBrandImageGetList
		Dim sqlStr, i, sqladd

		If FRectIsUsing<>"" then
			sqladd = " and isusing = " & FRectIsUsing & " "
		end if

		If FRectMakerid <> "" Then
			sqladd = sqladd & " and makerid = '"&FRectMakerid&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_brand_image "
		sqlStr = sqlStr & " WHERE 1=1" & sqladd
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, makerid, brandimage, isusing, regdate, adminid , lastupdate , lastadminid "
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_brand_image "
		sqlStr = sqlStr & " WHERE 1=1" & sqladd
		sqlStr = sqlStr & " ORDER BY idx DESC"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new cBrandItem
					FItemList(i).FIdx			= rsget("idx")
					FItemList(i).Fmakerid		= rsget("makerid")
					FItemList(i).Fbrandimage	= rsget("brandimage")
					FItemList(i).FIsusing		= rsget("isusing")
					FItemList(i).FRegdate		= rsget("regdate")
					FItemList(i).Fadminid		= rsget("adminid")
					FItemList(i).Flastupdate	= rsget("lastupdate")
					FItemList(i).Flastadminid	= rsget("lastadminid")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function	
End Class
%>