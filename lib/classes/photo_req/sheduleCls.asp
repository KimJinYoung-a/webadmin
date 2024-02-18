<%
'####################################################
' Description :  세미나실 클래스
' History : 2009.04.07 서동석 생성
'			2010.12.27 한용민 수정
'####################################################
Class CSeminarRoomCalendarItem
	Public Fidx
	Public Fusestart
	Public Fusetime
	Public Fbasictime
	Public Fbasictime2
	Public Fusername
	Public Fuserphone
	Public Fusepeople

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
	Public FReq_photo

	Public FRectRoom
	Public FReqStatus
	Public FStatus
	Public FReqStatusColor
	Public FReqDepartment
	Public FPrdName
	Public FUsername
	Public FPrd_type2
	Public FStartDate
	Public FStartDateUse
	Public FEndDateUse
	Public FEndDate
	Public FReqUse
	Public FReqUseDetail
	Public FReqName
	Public Fuser_id
	Public Fuser_name
	Public Freq_category
	Public Freq_stylist
	Public FCount
	Public FReqNo
	Public FDate
	Public FRname
	Public Fbasictime
	Public Fbasictime2
	Public Fusetime
	Public FWritename
	Public FCode_nm
	Public FUse_yn
	Public FReq_comment
	Public FSchedule_no

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

	' /admin/photo_req/request_cal.asp
	Public Sub list()
		Dim sql, i, addsql

		If FReq_photo <> "" Then
			addsql = addsql + " and b.req_photo = '"& FReq_photo &"' "
		Else
			addsql = ""
		End If

		sql = "select distinct count(a.req_no)as cnt, convert(varchar(10),b.start_date,20) as usedate, a.req_name as req_name," + vbcrlf
		sql = sql + " a.req_status as req_status, a.req_department as req_department, a.prd_name as prd_name, T.username, A.prd_type2 " + vbcrlf
		sql = sql + " FROM [db_partner].[dbo].[tbl_photo_req] A" + vbcrlf
		sql = sql + " left join  [db_partner].[dbo].[tbl_photo_schedule] b " + vbcrlf
		sql = sql + " 	on a.req_no = b.req_no " + vbcrlf
		sql = sql + " Inner Join db_partner.dbo.tbl_user_tenbyten as T on A.req_name = T.userid " + vbcrlf
		sql = sql + " where convert(varchar(7),b.start_date,20) = '" + FRectYYYYMM + "' "& addsql &" " + vbcrlf
		sql = sql + " and a.use_yn = 'Y' " + vbcrlf
		sql = sql + " group by convert(varchar(10),b.start_date,20), a.req_name, a.req_status, a.req_department, a.prd_name, T.username, a.prd_type2 " + vbcrlf
		sql = sql + " order by convert(varchar(10),b.start_date,20) asc"

		'response.write sql & "<Br>"
		rsget.Open sql, dbget, 1
		FResultCount = rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		If  not rsget.EOF  then
			i = 0
			Do until rsget.eof
				set FItemList(i) = new CSeminarRoomCalendar

				FItemList(i).FCount					= rsget("cnt")
				FItemList(i).FDate					= rsget("usedate")
				FItemList(i).FRname					= rsget("req_name")
				If rsget("req_status") = "4" Then
					FItemList(i).FReqStatus = "추가 기입 요청"
					FItemList(i).FReqStatusColor = "#000000"
				ElseIf rsget("req_status") = "1" Then
					FItemList(i).FReqStatus = "촬영스케줄 지정"
					FItemList(i).FReqStatusColor = "#000000"
				ElseIf rsget("req_status") = "2" Then
					FItemList(i).FReqStatus = "촬영중"
					FItemList(i).FReqStatusColor = "#000000"
				ElseIf rsget("req_status") = "3" Then
					FItemList(i).FReqStatus = "촬영완료"
					FItemList(i).FReqStatusColor = "#FF0000"
				End If 
				FItemList(i).FReqDepartment			= rsget("req_department")
				FItemList(i).FPrdName				= rsget("prd_name")
				FItemList(i).FUsername				= rsget("username")
				FItemList(i).FPrd_type2				= rsget("prd_type2")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	End Sub

	Public Sub DailyList()
		Dim sql, i

		sql = "select distinct count(a.req_no) as cnt, a.req_no as req_no," + vbcrlf
		sql = sql + " convert(varchar(10),b.start_date,20) as aaa, " + vbcrlf
		sql = sql + " a.req_name as req_name, a.req_department as req_department, c.user_id as user_id, c.user_name as user_name, " + vbcrlf
		sql = sql + " a.req_status as req_status, a.req_use as req_use, a.req_use_detail as req_use_detail, " + vbcrlf
		sql = sql + " a.prd_name as prd_name, a.req_category as req_category, b.req_stylist as req_stylist, " + vbcrlf
		sql = sql + " b.start_date as start_date, b.end_date as end_date, " + vbcrlf
		sql = sql + " cc.user_name as writename, L.code_nm, a.use_yn, a.req_comment, b.schedule_no, b.status " + vbcrlf
		sql = sql + " FROM [db_partner].[dbo].[tbl_photo_req] a " + vbcrlf
		sql = sql + " left join [db_partner].[dbo].[tbl_photo_schedule] b" + vbcrlf
		sql = sql + " 	on a.req_no = b.req_no " + vbcrlf
		sql = sql + " left join [db_partner].[dbo].[tbl_photo_user] c" + vbcrlf
		sql = sql + " 	on b.req_photo = c.user_id " + vbcrlf
		sql = sql + " left join [db_partner].[dbo].[tbl_photo_user] cc on a.req_name = cc.user_id " + vbcrlf
		sql = sql + " Left Join db_item.dbo.tbl_Cate_large as L on a.req_category = L.code_large  " + vbcrlf
		sql = sql + " where convert(varchar(10),b.start_date,20) = '" + FRectYYYYMM + "'" + vbcrlf
		sql = sql + " and c.user_id  ='" + Cstr(FRectRoom) + "' and (a.use_yn = 'Y' OR a.use_yn = 'S') " + vbcrlf
		sql = sql + " group by convert(varchar(10),b.start_date,20), a.req_no, a.req_department, a.req_name, c.user_id, c.user_name, " + vbcrlf
		sql = sql + " a.req_status, a.req_use, a.req_use_detail, a.prd_name, a.req_category, b.req_stylist, " + vbcrlf
		sql = sql + " b.start_date, b.end_date, cc.user_name, L.code_nm, a.use_yn, a.req_comment, b.schedule_no, b.status " + vbcrlf
		sql = sql + " order by convert(varchar(10),b.start_date,20) asc"

		rsget.Open sql, dbget, 1
		FResultCount = rsget.RecordCount
		'response.write sql
		'response.write  FResultCount

		Redim preserve FItemList(FResultCount)
		If  not rsget.EOF  Then
			i = 0
			Do until rsget.eof
				Set FItemList(i) = new CSeminarRoomCalendar
				FItemList(i).FReqNo				= rsget("req_no")
				FItemList(i).FReqName			= rsget("req_name")

				If rsget("req_department") & "" = "" Then
					FItemList(i).FReqDepartment	= ""
				Else
					FItemList(i).FReqDepartment	= rsget("req_department")
				End If 
				
				FItemList(i).Fuser_id	        = rsget("user_id")
				FItemList(i).Fuser_name	        = rsget("user_name")
				FItemList(i).FStatus			= rsget("status")
				FItemList(i).FReqUse			= rsget("req_use")
				FItemList(i).FReqUseDetail		= rsget("req_use_detail")
				FItemList(i).FPrdName			= rsget("prd_name")
				FItemList(i).Freq_category		= rsget("req_category")
				FItemList(i).Freq_stylist		= rsget("req_stylist")
				FItemList(i).FStartDate			= rsget("start_date")
				FItemList(i).FStartDateUse		= FormatDateTime(rsget("start_date"),4)
				FItemList(i).FEndDate			= rsget("end_date")
				FItemList(i).FEndDateUse		= FormatDateTime(rsget("end_date"),4)
				FItemList(i).Fbasictime			= getBaseTime(FItemList(i).FStartDateUse)
				FItemList(i).Fbasictime2		= getBaseTime(FItemList(i).FEndDateUse)
				FItemList(i).Fusetime			= FItemList(i).Fbasictime2 - FItemList(i).Fbasictime
				FItemList(i).FWritename			= rsget("writename")
				FItemList(i).FCode_nm			= rsget("code_nm")
				FItemList(i).FUse_yn			= rsget("use_yn")
				FItemList(i).FReq_comment		= rsget("req_comment")
				FItemList(i).FSchedule_no		= rsget("schedule_no")
				i=i+1
				rsget.moveNext
			loop
		End If
		rsget.close
	End Sub

	Function getBaseTime(strTime)
		Select Case strTime
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
		End Select
	End Function

	Public Function fnGetSchedule
		Dim strSql, i

		strSql = ""
		strSql = strSql & " select R.req_no, s.req_photo, S.start_date, S.end_date, R.req_comment " & vbcrlf
		strSql = strSql & " from [db_partner].[dbo].[tbl_photo_schedule] as S " & vbcrlf
		strSql = strSql & " Inner Join [db_partner].[dbo].[tbl_photo_req] as R On S.req_no = R.req_no " & vbcrlf
		strSql = strSql & " where S.req_no = '"& FReqNo &"'" & vbcrlf
		rsget.Open strSql,dbget,1

		IF not rsget.EOF THEN
			fnGetSchedule = rsget.getRows()
		End IF

		rsget.Close
	End Function


End Class


Function UserCodeType(uc, doctype, wr)
Dim query1
	If uc = "user" Then
		query1 = " select user_name from [db_partner].[dbo].tbl_photo_user"
		query1 = query1 + " where user_type='2' and user_useyn = 'Y' and user_id = '"&wr&"'"
		rsget.Open query1,dbget,1
		if  not rsget.EOF  Then
			UserCodeType = rsget("user_name")
		end if
		rsget.close
	ElseIf uc = "code" Then
		query1 = " select code_name from [db_partner].[dbo].tbl_photo_code"
		query1 = query1 + " where code_type='"&doctype&"' and code_name = '"&wr&"'"
		rsget.Open query1,dbget,1
		if  not rsget.EOF  Then
			UserCodeType = rsget("code_name")
		end if
		rsget.close
	End If
End Function

Sub SelectUser(BB, CC)
	Dim query1
	query1 = " select user_no, user_id, user_name from [db_partner].[dbo].tbl_photo_user"
	query1 = query1 + " where user_type='1' and user_useyn = 'Y'"
	rsget.Open query1,dbget,1
%>
	<select class="select" name='<%=BB%>' onchange="fnSearch();">
		<option value=''>-- 포토그래퍼 선택 --</option>
<%
	If not rsget.EOF Then
		rsget.Movefirst
		Do until rsget.EOF
			response.write("<option value='"&rsget("user_id")& "' "& chkIIF(CC = rsget("user_id"),"selected","") &">" & rsget("user_name") & "" & "</option>")
			rsget.MoveNext
		Loop
	End If
	rsget.close
	response.write("</select>")
End Sub
%>