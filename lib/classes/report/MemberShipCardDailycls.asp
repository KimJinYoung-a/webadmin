<%

Class cMemberShipCardDailyitem
	public Fregdate
	public FCardRegCnt
	public FCardSavingCnt
	public FCardSavingPoint
	public FCardUsingCnt
	public FCardUsingPoint
	public FChangeOnlineCnt
	public FChangeOnlinePoint


    Private Sub Class_Initialize()
	end sub

	Private Sub Class_Terminate()
	End Sub
End Class


Class CMemberShipCardDaily

	public FItemList()
	public FRectFromDate
	public FRectToDate
	public FRectSearchType
	Public FTotalCount


	public Sub GetMemberShipCardDailyReport

		If FRectToDate <> "" Then
			FRectToDate = DateAdd("d", 1, FRectToDate)
			FRectToDate = Left(FRectToDate, 10)
		End If

		dim sqlStr, i
		sqlStr = " Select CardRegCnt, CardSavingCnt, CardSavingPoint, CardUsingCnt, CardUsingPoint, ChangeOnlineCnt, ChangeOnlinePoint, convert(varchar(10), regdate, 120) as regdate  "
		sqlStr = sqlStr + " From db_datamart.dbo.tbl_MemberShipCardSummary "
		sqlStr = sqlStr + " Where idx is not null "
		sqlStr = sqlStr + " And regdate >= '"&FRectFromDate&"' "
		sqlStr = sqlStr + " And regdate < '"&FRectToDate&"' "
		sqlStr = sqlStr + " order by idx asc "
'		response.write sqlStr
		db3_rsget.open sqlstr,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		redim FItemList(FTotalCount)
		i = 0

		if not db3_rsget.eof Then
			Do Until db3_rsget.eof
				set FItemList(i) = new cMemberShipCardDailyitem

				FItemList(i).Fregdate = db3_rsget("regdate")
				FItemList(i).FCardRegCnt = db3_rsget("CardRegCnt")
				FItemList(i).FCardSavingCnt = db3_rsget("CardSavingCnt")
				FItemList(i).FCardSavingPoint = db3_rsget("CardSavingPoint")
				FItemList(i).FCardUsingCnt = db3_rsget("CardUsingCnt")
				FItemList(i).FCardUsingPoint = db3_rsget("CardUsingPoint")
				FItemList(i).FChangeOnlineCnt = db3_rsget("ChangeOnlineCnt")
				FItemList(i).FChangeOnlinePoint = db3_rsget("ChangeOnlinePoint")


				i=i+1
			db3_rsget.MoveNext
			Loop
		End If

		db3_rsget.close
	End Sub

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

%>