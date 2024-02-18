<%

Class cwishAppDailyitem
	public Fregdate
	public FlogWeb
	public FlogMobile
	public FlogApp
	public FfollowuCnt
	public FfollowpCnt
	public FprevDayPM
	public FwishpdAll
	public FwishpdView
	public FwishfdAll
	public FwishfdView
	public Fappios
	public Fappand

    public FAppIosNid
    public FAppAndNid
    
    Private Sub Class_Initialize()
	end sub

	Private Sub Class_Terminate()
	End Sub
End Class


Class CwishAppDaily

	public FItemList()
	public FRectFromDate
	public FRectToDate
	public FRectSearchType
	Public FTotalCount


	public Sub GetwishAppDailyReport

		If FRectToDate <> "" Then
			FRectToDate = DateAdd("d", 1, FRectToDate)
			FRectToDate = Left(FRectToDate, 10)
		End If

		dim sqlStr, i
		sqlStr = " Select logWeb, logMobile, logApp, isNULL(FollowuCnt,0) as FollowuCnt, FollowpCnt, isNULL(PrevDayPM,0) as PrevDayPM, wishpdAll, wishpdView, wishfdAll, wishfdView, Appios, Appand"
		sqlStr = sqlStr + " , isNULL(AppIosNid,0) as AppIosNid, isNULL(AppAndNid,0) as AppAndNid, convert(varchar(10), regdate, 120) as regdate  "
		sqlStr = sqlStr + " From db_datamart.dbo.tbl_WishApp_DailyCnt "
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
				set FItemList(i) = new cwishAppDailyitem

				FItemList(i).Fregdate = db3_rsget("regdate")
				FItemList(i).FlogWeb = db3_rsget("logWeb")
				FItemList(i).FlogMobile = db3_rsget("logMobile")
				FItemList(i).FlogApp = db3_rsget("logApp")
				FItemList(i).FfollowuCnt = db3_rsget("FollowuCnt")
				FItemList(i).FfollowpCnt = db3_rsget("FollowpCnt")
				FItemList(i).FprevDayPM = db3_rsget("PrevDayPM")
				FItemList(i).FwishpdAll = db3_rsget("wishpdAll")
				FItemList(i).FwishpdView = db3_rsget("wishpdView")
				FItemList(i).FwishfdAll = db3_rsget("wishfdAll")
				FItemList(i).FwishfdView = db3_rsget("wishfdView")
				FItemList(i).Fappios = db3_rsget("Appios")
				FItemList(i).Fappand = db3_rsget("Appand")
    
                FItemList(i).FAppIosNid = db3_rsget("AppIosNid")
				FItemList(i).FAppAndNid = db3_rsget("AppAndNid")
				
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