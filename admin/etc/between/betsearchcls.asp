<%
Class cSearchOneItem
	Public FIdx
	Public FRank
	Public FLikeword
	Public FIsusing
	Public FUpdateDate
	Public FRegdate
End Class

Class cSearch
	Public FItemList()
	Public FOneItem
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount

	Public FRectIdx
	Public FRectIsusing
	Public Sub getLikeWordList
		Dim sqlStr, i, addsql

		If FRectIsusing <> "" Then
			addSql = addSql & " and isusing = '"&FRectIsusing&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_search_likeWord "
		sqlStr = sqlStr & " WHERE 1 = 1  " & addSql
		rsCTget.Open sqlStr,dbCTget,1
			FTotalCount = rsCTget("cnt")
			FTotalPage = rsCTget("totPg")
		rsCTget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & " idx, rank, likeword, isusing, updateDate, regdate "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_search_likeWord "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY rank ASC "
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cSearchOneItem
					FItemList(i).FIdx			= rsCTget("idx")
					FItemList(i).FRank			= rsCTget("rank")
					FItemList(i).FLikeword		= rsCTget("likeword")
					FItemList(i).FIsusing		= rsCTget("isusing")
					FItemList(i).FUpdateDate	= rsCTget("updateDate")
					FItemList(i).FRegdate		= rsCTget("regdate")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub getOneLikeWord
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT idx, rank, likeword, isusing, updateDate "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_search_likeWord "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and idx = '"& FRectIdx &"' "
		rsCTget.Open sqlStr, dbCTget, 1
		FResultCount = rsCTget.RecordCount
		Set FOneItem = new cSearchOneItem
		If Not rsCTget.Eof Then
			FOneItem.FIdx		= rsCTget("idx")
			FOneItem.FRank		= rsCTget("rank")
			FOneItem.FLikeword	= rsCTget("likeword")
			FOneItem.FIsusing	= rsCTget("isusing")
			FOneItem.FUpdateDate = rsCTget("updateDate")
		End If
		rsCTget.close
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

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FScrollCount = 10
	End Sub

	Private Sub Class_Terminate()
    End Sub
End Class
%>