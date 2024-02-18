<%
Class coneline_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public fuserid
	public fcomment
	public fwinYN
	public fisusing
	public fregdate
	public ficon
	public fuserlevel
	
End Class


Class ClsOneLine

	public FItemList()
	public FOneItem
	public FGubun
	public FEvtCode
	public FUserID
	public FIdx
	public ftotalcount
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalPage
	public FScrollCount

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	
	
	public sub FOneLineList
		Dim sqlStr, i, vSubQuery
		sqlStr = "SELECT COUNT(*) " & _
				 "		FROM [db_contents].[dbo].[tbl_one_comment] AS O " & _
				 "	INNER JOIN [db_user].[dbo].[tbl_logindata] AS U ON O.userid = U.userid " & _
				 "	WHERE O.evt_code = '" & FEvtCode & "' " & _
				 "	" & vSubQuery & " "
		rsget.Open sqlStr, dbget ,1
		ftotalcount = rsget(0)
		rsget.Close
		
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " O.idx, O.userid, O.comment, O.winYN, O.isusing, O.regdate, O.icon, U.userlevel " & _
				 "		FROM [db_contents].[dbo].[tbl_one_comment] AS O " & _
				 "	INNER JOIN [db_user].[dbo].[tbl_logindata] AS U ON O.userid = U.userid " & _
				 "	WHERE O.evt_code = '" & FEvtCode & "' " & _
				 "	" & vSubQuery & " " & _
				 "	ORDER BY O.idx DESC "
		
		rsget.Open sqlStr, dbget ,1
		
		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		rsget.PageSize= FPageSize
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do Until rsget.Eof
				set FItemList(i) = new coneline_oneitem

					FItemList(i).fidx		= rsget("idx")
					FItemList(i).fuserid	= rsget("userid")
					FItemList(i).fcomment	= db2html(rsget("comment"))
					FItemList(i).fwinYN		= rsget("winYN")
					FItemList(i).fisusing	= rsget("isusing")
					FItemList(i).fregdate	= rsget("regdate")
					FItemList(i).ficon		= rsget("icon")
					FItemList(i).fuserlevel	= rsget("userlevel")
					
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end sub
	
	
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

End Class


'지정된 날짜로 그달의 주차 반환 함수
Function getWeekSerial(dt)
	dim startWeek, totalWeek
	totalWeek = DatePart("ww", dt)	'전체 주차
	startWeek = DatePart("ww", DateSerial(year(dt),month(dt),"01"))		'첫째일 주차

	'계산 및 값 반환
	getWeekSerial = totalWeek - startWeek + 1	
end Function
%>