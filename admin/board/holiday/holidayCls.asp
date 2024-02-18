<%
Class CHolidayItem
	Public FNum
	Public FLunar_date
	Public FSolar_date
	Public FYun
	Public FGanji
	Public FHoliday
	Public FHoliday_name
	Public FWeek
	Public FSolar_yyyy
	Public FSolar_yyyymm
	Public FLogics_holiday
	Public FUpche_holiday
End Class

Class CHoliday
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectStartDate
	Public FRectEndDate
	Public FRectHoliday
	Public FRectLogicsHoliday
	Public FRectUpcheHoliday
	Public FRectNum

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
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

	Public Sub getHolidayItemList
		Dim sqlStr, i, addSql
		If FRectHoliday <> "" Then
			addSql = addSql & " and holiday in ('1', '2') "
		End If

		If FRectLogicsHoliday <> "" Then
			addSql = addSql & " and logics_holiday in ('2') "
		End If

		If FRectUpcheHoliday <> "" Then
			addSql = addSql & " and upche_holiday in ('2') "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM [db_sitemaster].[dbo].[LunarToSolar] with (nolock) " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and solar_date between '"& FRectStartDate &"' and '"& FRectEndDate &"' " & VBCRLF
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		If CLng(FCurrPage) > CLng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & "	[num], [lunar_date], [solar_date], [yun], [ganji], [holiday], [holiday_name], [week], [solar_yyyy], [solar_yyyymm], [logics_holiday], [upche_holiday] " & VBCRLF
		sqlStr = sqlStr & " FROM [db_sitemaster].[dbo].[LunarToSolar] with (nolock) " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1  " & VBCRLF
		sqlStr = sqlStr & " and solar_date between '"& FRectStartDate &"' and '"& FRectEndDate &"' " & VBCRLF
		sqlStr = sqlStr & addSql & VBCRLF
	    sqlStr = sqlStr & " ORDER BY num ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CHolidayItem
					FItemList(i).FNum				= rsget("num")
					FItemList(i).FLunar_date		= rsget("lunar_date")
					FItemList(i).FSolar_date		= rsget("solar_date")
					FItemList(i).FYun				= rsget("yun")
					FItemList(i).FGanji				= rsget("ganji")
					FItemList(i).FHoliday			= rsget("holiday")
					FItemList(i).FHoliday_name		= rsget("holiday_name")
					FItemList(i).FWeek				= rsget("week")
					FItemList(i).FSolar_yyyy		= rsget("solar_yyyy")
					FItemList(i).FSolar_yyyymm		= rsget("solar_yyyymm")
					FItemList(i).FLogics_holiday	= rsget("logics_holiday")
					FItemList(i).FUpche_holiday		= rsget("upche_holiday")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getHolidayOneItem
		Dim sqlStr, i, addSql

		If FRectNum <> "" Then
			addSql = addSql & " and num = '"& FRectNum &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 " & VBCRLF
		sqlStr = sqlStr & "	[num], [lunar_date], [solar_date], [yun], [ganji], [holiday], [holiday_name], [week], [solar_yyyy], [solar_yyyymm], [logics_holiday], [upche_holiday] " & VBCRLF
		sqlStr = sqlStr & " FROM [db_sitemaster].[dbo].[LunarToSolar] with (nolock) " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1  " & VBCRLF
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHolidayItem
				FOneItem.FNum				= rsget("num")
				FOneItem.FLunar_date		= rsget("lunar_date")
				FOneItem.FSolar_date		= rsget("solar_date")
				FOneItem.FYun				= rsget("yun")
				FOneItem.FGanji				= rsget("ganji")
				FOneItem.FHoliday			= rsget("holiday")
				FOneItem.FHoliday_name		= rsget("holiday_name")
				FOneItem.FWeek				= rsget("week")
				FOneItem.FSolar_yyyy		= rsget("solar_yyyy")
				FOneItem.FSolar_yyyymm		= rsget("solar_yyyymm")
				FOneItem.FLogics_holiday	= rsget("logics_holiday")
				FOneItem.FUpche_holiday		= rsget("upche_holiday")
		End If
		rsget.Close
	End Sub

End Class

Function Get_Lastday(nYear, nMonth)
    Get_Lastday = Day(DateSerial(nYear, nMonth + 1, 1 - 1))
End Function
%>