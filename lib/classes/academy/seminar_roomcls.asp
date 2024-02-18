<%
'####################################################
' Description :  세미나실 클래스
' History : 2009.04.07 서동석 생성
'			2010.12.27 한용민 수정
'####################################################

class CSeminarRoomCalendarItem
	public Fidx
	public FCount
	public FDate
	public Fusestart
	public Fusetime
	public Fbasictime
	public Fusername
	public Fuserphone
	public Fusepeople

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CSeminarRoomCalendar
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectYYYYMM
	public FRectRoom

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

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

		sql = "select count(usestart) as cnt, convert(varchar(10),usestart,20) as usedate" + vbcrlf
		sql = sql + " from [db_shop].[dbo].tbl_seminar_room" + vbcrlf
		sql = sql + " where isusing = 'Y'" + vbcrlf
		sql = sql + " and convert(varchar(7),usestart,20) = '" + FRectYYYYMM + "'" + vbcrlf
		sql = sql + " group by convert(varchar(10),usestart,20)" + vbcrlf
		sql = sql + " order by convert(varchar(10),usestart,20) asc"
		
		'response.write sql &"<br>"
		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			do until rsget.eof
				set FItemList(i) = new CSeminarRoomCalendarItem

				FItemList(i).FCount          = rsget("cnt")
				FItemList(i).FDate          = rsget("usedate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub DailyList()
		dim sql, i

		sql = "select idx,usestart,usetime,basictime,username,userphone,usepeople from [db_shop].[dbo].tbl_seminar_room" + vbcrlf
		sql = sql + " where isusing = 'Y'" + vbcrlf
		sql = sql + " and convert(varchar(10),usestart,20) = '" + FRectYYYYMM + "'" + vbcrlf
		sql = sql + " and roomid='" + Cstr(FRectRoom) + "'" + vbcrlf
		sql = sql + " order by usestart asc"
		
		'response.write sql &"<br>"
		rsget.Open sql, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			do until rsget.eof
				set FItemList(i) = new CSeminarRoomCalendarItem

				FItemList(i).Fidx          = rsget("idx")
				FItemList(i).Fusestart          = rsget("usestart")
				FItemList(i).Fusetime          = rsget("usetime")
				FItemList(i).Fbasictime          = rsget("basictime")
				FItemList(i).Fusername          = rsget("username")
				FItemList(i).Fuserphone          = rsget("userphone")
				FItemList(i).Fusepeople          = rsget("usepeople")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

end Class

Class CSeminarRoomDetail
	public Fidx
	public Froomid
	public Fusestart
	public Fusetime
	public Fbasictime
	public Fgroupname
	public Fusername
	public Fuserphone
	public Fuserhp
	public Fusepeople
	public Fisusing
	public Fregdate
	public Fetc
	public flecturer_idx
	
	public function GetRoomName()
		if Froomid = "01" then
			GetRoomName = "Idea"
		elseif Froomid = "02" then
			GetRoomName = "Paper"
		elseif Froomid = "03" then
			GetRoomName = "Heart"
		elseif Froomid = "04" then
			GetRoomName = "Fingers"
		elseif Froomid = "05" then
			GetRoomName = "Television"
		elseif Froomid = "06" then
			GetRoomName = "Chocolate"
		elseif Froomid = "07" then
			GetRoomName = "Bingo"
		elseif Froomid = "08" then
			GetRoomName = "Moon"
		elseif Froomid = "09" then
			GetRoomName = "Star"
		else
			GetRoomName = "?"
		end if
	end function

	Function UseTimeName()
		if Fusetime = 12 then
			UseTimeName = "12:00"
		elseif Fusetime = 13 then
			UseTimeName = "12:30"
		elseif Fusetime = 14 then
			UseTimeName = "13:00"
		elseif Fusetime = 15 then
			UseTimeName = "13:30"
		elseif Fusetime = 16 then
			UseTimeName = "14:00"
		elseif Fusetime = 17 then
			UseTimeName = "14:30"
		elseif Fusetime = 18 then
			UseTimeName = "15:00"
		elseif Fusetime = 19 then
			UseTimeName = "15:30"
		elseif Fusetime = 20 then
			UseTimeName = "16:00"
		elseif Fusetime = 21 then
			UseTimeName = "16:30"
		elseif Fusetime = 22 then
			UseTimeName = "17:00"
		elseif Fusetime = 23 then
			UseTimeName = "17:30"
		elseif Fusetime = 24 then
			UseTimeName = "18:00"
		elseif Fusetime = 25 then
			UseTimeName = "18:30"
		elseif Fusetime = 26 then
			UseTimeName = "19:00"
		elseif Fusetime = 27 then
			UseTimeName = "19:30"
		elseif Fusetime = 28 then
			UseTimeName = "20:00"
		elseif Fusetime = 29 then
			UseTimeName = "20:30"
		elseif Fusetime = 30 then
			UseTimeName = "21:00"
		elseif Fusetime = 31 then
			UseTimeName = "21:30"
		elseif Fusetime = 32 then
			UseTimeName = "22:00"
		elseif Fusetime = 33 then
			UseTimeName = "22:30"
		elseif Fusetime = 34 then
			UseTimeName = "23:00"
		elseif Fusetime = 35 then
			UseTimeName = "23:30"
		end if
	End Function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'//admin/lecture/seminar_room_edit.asp
	public Sub read(byval idx)
		dim sql, i

		sql = "select * from [db_shop].[dbo].tbl_seminar_room" + vbcrlf
		sql = sql + " where isusing = 'Y'" + vbcrlf
		sql = sql + " and idx='" + Cstr(idx) + "'" + vbcrlf
		sql = sql + " order by idx desc"

		'response.write sql &"<br>"

		rsget.Open sql, dbget, 1

		if  not rsget.EOF  then

				Fidx          = rsget("idx")
				Froomid   = rsget("roomid")
				Fusestart   = rsget("usestart")
				Fusetime   =  rsget("usetime")
				Fbasictime   =  rsget("basictime")
				Fgroupname   =  rsget("groupname")
				Fusername   =  rsget("username")
				Fuserphone   =  rsget("userphone")
				Fuserhp   =  rsget("userhp")
				Fusepeople   =  rsget("usepeople")
				Fetc   =  rsget("etc")
				Fisusing   = rsget("isusing")
				Fregdate      = rsget("regdate")
				flecturer_idx      = rsget("lecturer_idx")

		end if
		rsget.close
	end sub

end Class
%>