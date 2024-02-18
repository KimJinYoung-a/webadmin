<%
'####################################################
' Description :  세미나실 클래스
' History : 2012.10.24 김진영 생성
'####################################################
Class CSeminarRoomCalendarItem
	Public Fidx
	Public FCount
	Public FDate
	Public Fusestart
	Public Fusetime
	Public Fbasictime
	Public Fbasictime2
	Public Fusername
	Public Fuserphone
	Public Fusepeople
	Public Froomidx
	Public Fstart_date
	Public Fend_date
	Public Fgroupname
	Public Fusepurpose
	Public Fusercell
	Public FuseSu
	Public Fetc
	Public Flecnum
	Public Fisusing
	Public FadminID
	Public Fregdate
	Public FStartDateUse
	Public FEndDateUse

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CSeminarRoomCalendar
	Public FItemList()
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	Public FRectYYYYMM
	Public FRectRoom
	Public FIdx

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		Redim  FItemList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Sub list()
		dim sql, i

		sql = "select count(start_date) as cnt, convert(varchar(10),start_date,20) as usedate" + vbcrlf
		sql = sql + " from [db_partner].[dbo].[tbl_seminar_schedule] " + vbcrlf
		sql = sql + " where isusing = 'Y'" + vbcrlf
		sql = sql + " and convert(varchar(7),start_date,20) = '" + FRectYYYYMM + "'" + vbcrlf
		sql = sql + " group by convert(varchar(10),start_date,20)" + vbcrlf
		sql = sql + " order by convert(varchar(10),start_date,20) asc"
		
		'response.write sql &"<br>"
		rsget.Open sql, dbget, 1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		If not rsget.EOF Then
			i = 0
			Do until rsget.eof
				set FItemList(i) = new CSeminarRoomCalendarItem
					FItemList(i).FCount          = rsget("cnt")
					FItemList(i).FDate          = rsget("usedate")
				i=i+1
				rsget.moveNext
			loop
		End If
		rsget.close
	End Sub

	public Sub DailyList()
		Dim sql, i

		sql = "SELECT idx, roomidx, start_date, end_date, groupname, usepurpose, usercell, useSu, etc, lecnum, isusing, adminID, regdate " + vbcrlf
		sql = sql + " ,(select username from db_partner.dbo.tbl_user_tenbyten where userid = adminID) as username " + vbcrlf
		sql = sql + " FROM [db_partner].[dbo].[tbl_seminar_schedule] " + vbcrlf
		sql = sql + " WHERE isusing = 'Y' and convert(varchar, start_date,23) = '" + FRectYYYYMM + "' " + vbcrlf
		sql = sql + " and roomidx = '" + Cstr(FRectRoom) + "'" + vbcrlf
		sql = sql + " order by start_date asc"

		rsget.Open sql, dbget, 1
		FResultCount = rsget.RecordCount

		Redim preserve FItemList(FResultCount)
		If  not rsget.EOF  Then
			i = 0
			Do until rsget.eof
				Set FItemList(i) = new CSeminarRoomCalendarItem
					FItemList(i).Fidx			= rsget("idx")
					FItemList(i).Froomidx		= rsget("roomidx")
					FItemList(i).Fstart_date	= rsget("start_date")
					FItemList(i).FStartDateUse	= FormatDateTime(rsget("start_date"),4)
					FItemList(i).Fend_date		= rsget("end_date")
					FItemList(i).FEndDateUse	= FormatDateTime(rsget("end_date"),4)
					FItemList(i).Fbasictime		= getBaseTime(FItemList(i).FStartDateUse)
					FItemList(i).Fbasictime2	= getBaseTime(FItemList(i).FEndDateUse)
					FItemList(i).Fusetime		= FItemList(i).Fbasictime2 - FItemList(i).Fbasictime
					FItemList(i).Fgroupname		= rsget("groupname")
					FItemList(i).Fusepurpose	= rsget("usepurpose")
					FItemList(i).Fusercell		= rsget("usercell")
					FItemList(i).FuseSu			= rsget("useSu")
					FItemList(i).Fetc			= rsget("etc")
					FItemList(i).Flecnum		= rsget("lecnum")
					FItemList(i).Fisusing		= rsget("isusing")
					FItemList(i).FadminID		= rsget("adminID")
					FItemList(i).Fregdate		= rsget("regdate")
					FItemList(i).Fusername		= rsget("username")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.close
	End Sub

	Function getBaseTime(strTime)
		Select Case strTime
			Case "09:00"	getBaseTime = "6"
			Case "09:30"	getBaseTime = "7"
			Case "10:00"	getBaseTime = "8"
			Case "10:30"	getBaseTime = "9"
			Case "11:00"	getBaseTime = "10"
			Case "11:30"	getBaseTime = "11"
			Case "12:00"	getBaseTime = "12"
			Case "12:30"	getBaseTime = "13"
			Case "13:00"	getBaseTime = "14"
			Case "13:30"	getBaseTime = "15"
			Case "14:00"	getBaseTime = "16"
			Case "14:30"	getBaseTime = "17"
			Case "15:00"	getBaseTime = "18"
			Case "15:30"	getBaseTime = "19"
			Case "16:00"	getBaseTime = "20"
			Case "16:30"	getBaseTime = "21"
			Case "17:00"	getBaseTime = "22"
			Case "17:30"	getBaseTime = "23"
			Case "18:00"	getBaseTime = "24"
			Case "18:30"	getBaseTime = "25"
			Case "19:00"	getBaseTime = "26"
			Case "19:30"	getBaseTime = "27"
			Case "20:00"	getBaseTime = "28"
			Case "20:30"	getBaseTime = "29"
			Case "21:00"	getBaseTime = "30"
			Case "21:30"	getBaseTime = "31"
			Case "22:00"	getBaseTime = "32"
			Case "22:30"	getBaseTime = "33"
			Case "23:00"	getBaseTime = "34"
			Case "23:30"	getBaseTime = "35"
		End Select
	End Function

	Public Function fnGetSchedule
		Dim strSql, i
		strSql = ""
		strSql = strSql & " SELECT idx, roomidx, start_date, end_date, groupname, usepurpose, usercell, useSu, etc, lecnum, isusing " & vbcrlf
		strSql = strSql & " FROM [db_partner].[dbo].[tbl_seminar_schedule] " & vbcrlf
		strSql = strSql & " where idx = '"& FIdx &"'" & vbcrlf
		rsget.Open strSql,dbget,1

		IF not rsget.EOF THEN
			fnGetSchedule = rsget.getRows()
		End IF

		rsget.Close
	End Function

	public Sub getReservationList()
		Dim sql, i

		sql = "SELECT roomidx, start_date, end_date, groupname, usepurpose, usercell, useSu, etc, p.username " & vbcrlf
		sql = sql & " FROM [db_partner].[dbo].[tbl_seminar_schedule] WITH(NOLOCK)" & vbcrlf
		sql = sql & " CROSS APPLY (" & vbcrlf
		sql = sql & " SELECT username FROM db_partner.dbo.tbl_user_tenbyten WITH(NOLOCK) WHERE userid = adminID " & vbcrlf
		sql = sql & " ) as p " & vbcrlf
		sql = sql & " WHERE isusing = 'Y' and convert(varchar, start_date,23) > getdate()-7 " + vbcrlf
		sql = sql & " ORDER BY start_date ASC"

		rsget.Open sql, dbget, 1
		FResultCount = rsget.RecordCount

		Redim preserve FItemList(FResultCount)
		If  not rsget.EOF  Then
			i = 0
			Do until rsget.eof
				Set FItemList(i) = new CSeminarRoomCalendarItem
					FItemList(i).Froomidx		= rsget("roomidx")
					FItemList(i).Fstart_date	= rsget("start_date")
					FItemList(i).Fend_date		= rsget("end_date")
					FItemList(i).Fgroupname		= rsget("groupname")
					FItemList(i).Fusepurpose	= rsget("usepurpose")
					FItemList(i).Fusercell		= rsget("usercell")
					FItemList(i).FuseSu			= rsget("useSu")
					FItemList(i).Fetc			= rsget("etc")
					FItemList(i).Fusername		= rsget("username")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.close
	End Sub

end Class

Class CSeminarManageItem
	Public Fidx
	Public Froomname
	Public FMaxSu
	Public ForderNo
	Public Fisusing
End Class

Class CSeminarManage
	Public FItemList()
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	Public Fidx
	Public Froomname
	Public FMaxSu
	Public ForderNo
	Public Fisusing


	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Sub List()
		Dim strSql, i
		strSql = ""
		strSql = strSql & " SELECT idx, roomname, MaxSu, orderNo, isusing " & vbcrlf
		strSql = strSql & " From db_partner.dbo.tbl_seminarRoom " & vbcrlf
		strSql = strSql & " Order by orderNo ASC " & vbcrlf
		'response.write strSql &"<br>"
		rsget.Open strSql, dbget, 1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		If  not rsget.EOF  then
	        i = 0
			Do until rsget.eof
				Set FItemList(i) = new CSeminarManageItem
					FItemList(i).Fidx		= rsget("idx")
					FItemList(i).Froomname	= rsget("roomname")
					FItemList(i).FMaxSu		= rsget("MaxSu")
					FItemList(i).ForderNo	= rsget("orderNo")
					FItemList(i).Fisusing	= rsget("isusing")
				i=i+1
				rsget.moveNext
			loop
		End If
		rsget.close
	End Sub

	Public Sub Modify()
		Dim strSql, i
		strSql = ""
		strSql = strSql & " SELECT idx, roomname, MaxSu, orderNo, isusing " & vbcrlf
		strSql = strSql & " From db_partner.dbo.tbl_seminarRoom " & vbcrlf
		strSql = strSql & " where idx = '"&Fidx&"' " & vbcrlf
		'response.write strSql &"<br>"

		rsget.Open strSql, dbget, 1
		If  not rsget.EOF  then
			Fidx		= rsget("idx")
			Froomname	= rsget("roomname")
			FMaxSu		= rsget("MaxSu")
			ForderNo	= rsget("orderNo")
			Fisusing	= rsget("isusing")
		End If
		rsget.close

	End Sub
End Class
%>