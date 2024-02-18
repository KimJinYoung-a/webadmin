<%
Class cNoticeFaqOneItem
	Public FIdx
	Public FGubun
	Public FSubject
	Public FContents
	Public FRegdate
	Public FIsusing
End Class

Class cNoticeFAQ
	Public FItemList()
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount

	Public FRectIdx
	Public FRectGubun
	Public FRectSubject
	Public FRectIsusing

	Public Sub getBoardList()
		Dim sqlStr, i, addsql
		If FRectGubun <> "" Then
			addsql = addsql & " and gubun = '"&FRectGubun&"' "
		End If
	
		If FRectSubject <> "" Then
			addsql = addsql & " and subject = '"&FRectSubject&"' "
		End If

		If FRectIsusing <> "" Then
			addsql = addsql & " and isusing = '"&FRectIsusing&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_noticefaq "
		sqlStr = sqlStr & "	WHERE 1 = 1 " & addSql
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
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, gubun, subject, contents, regdate, isusing "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_noticefaq "
		sqlStr = sqlStr & "	WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & "	ORDER BY idx DESC "
		rsCTget.pagesize = FPageSize
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsCTget.EOF Then
			rsCTget.absolutepage = FCurrPage
			Do until rsCTget.EOF
				Set FItemList(i) = new cNoticeFaqOneItem
					FItemList(i).FIdx		= rsCTget("idx")
					FItemList(i).FGubun		= rsCTget("gubun")
					FItemList(i).FSubject	= rsCTget("subject")
					FItemList(i).FContents	= rsCTget("contents")
					FItemList(i).FRegdate	= rsCTget("regdate")
					FItemList(i).FIsusing	= rsCTget("isusing")
				i = i + 1
				rsCTget.moveNext
			Loop
		End If
		rsCTget.Close
	End Sub

	Public Sub getNoticeModify()
		Dim sqlStr
		IF FRectIdx = "" THEN Exit Sub
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 idx, gubun, subject, contents, regdate, isusing "
		sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_between_noticefaq "
		sqlStr = sqlStr & " WHERE idx = " & FRectIdx
		rsCTget.Open sqlStr,dbCTget,1
		FResultCount = rsCTget.RecordCount
		If not rsCTget.EOF Then
			Set FItemList(0) = new cNoticeFaqOneItem
				FItemList(0).FIdx		= rsCTget("idx")
				FItemList(0).FGubun		= rsCTget("gubun")
				FItemList(0).FSubject	= rsCTget("subject")
				FItemList(0).FContents	= rsCTget("contents")
				FItemList(0).FRegdate	= rsCTget("regdate")
				FItemList(0).FIsusing	= rsCTget("isusing")
		End If
		rsCTget.Close
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