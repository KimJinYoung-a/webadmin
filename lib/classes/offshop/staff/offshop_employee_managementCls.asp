<%
class cEmployeeManagementClass_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub


	public FWorkCode
	public FStartWork
	public FEndWork
	public FUserName
	public FEmpNO
	public FPartName
	public FWorkDate
	public FWorkTime
	public FInTime
	public FOutTime
	public FPlaceName
	public FShopName
	public FdepartmentNameFull

end class
class cEmployeeManagementClass_list

	public FList
	public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FPageCount
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	public FRectWorkCode
	public FRectWorkDate1
	public FRectWorkDate2
	public FRectPartSN
	public FRectPositSN
	public FRectPartName
	public FRectPlaceName
	public FRectSearchKey
	public FRectSearchString
	public FRectShopID
	public FRectOrderBy

	public Fdepartment_id
	public Finc_subdepartment

	public sub fWorkScheduleList()
		dim sqlStr,i, addStr

		If FRectPartSN <> "" Then
			addStr = addStr & " AND w.part_sn = '" & FRectPartSN & "' "
		End If

		If FRectShopID <> "" Then
			addStr = addStr & " AND ps.shopid = '" & FRectShopID & "' "
		End If

		If FRectWorkDate1 <> "" Then
			addStr = addStr & " AND w.workdate Between '" & FRectWorkDate1 & "' AND '" & FRectWorkDate2 & "' "
		End IF

		If FRectPositSN <> "" Then
			addStr = addStr & " AND u.posit_sn = '" & FRectPositSN & "' "
		End IF

		If FRectSearchKey <> "" Then
			If FRectSearchKey = "1" Then
				addStr = addStr & " AND u.empno = '" & FRectSearchString & "' "
			ElseIf FRectSearchKey = "2" Then
				addStr = addStr & " AND u.username = '" & FRectSearchString & "' "
			End If
		End If

		if (Fdepartment_id <> "") then
			if (Finc_subdepartment = "N") then
				addStr = addStr & " AND u.department_id = '" & Fdepartment_id & "' "
			else
				addStr = addStr & " AND (IsNull(dv.cid1, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid2, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid3, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid4, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid5, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid6, -1) = '" & Fdepartment_id & "') "
			end if
		end if


		'총 갯수 구하기
		sqlStr = "SELECT count(*) AS cnt " & vbCrLf
		sqlStr = sqlStr & " FROM [db_partner].[dbo].[tbl_offshop_employee_workschedule] AS w " & vbCrLf
		sqlStr = sqlStr & " 	INNER JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS u ON w.empno = u.empno " & vbCrLf
		sqlStr = sqlStr & " 	INNER JOIN [db_partner].[dbo].[tbl_partInfo] AS p ON w.part_sn = p.part_sn " & vbCrLf
		sqlStr = sqlStr & " 	INNER JOIN [db_partner].[dbo].[tbl_offshop_employee_workcode] AS c ON w.workcode = c.workcode " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_partner].[dbo].[tbl_partner_shopuser] AS ps ON w.empno = ps.empno and ps.firstisusing = 'Y' " & vbCrLf
		sqlStr = sqlStr & " 	left join db_partner.dbo.vw_user_department dv " & vbCrLf
		sqlStr = sqlStr & " 	on " & vbCrLf
		sqlStr = sqlStr & " 		1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 		and u.department_id = dv.cid " & vbCrLf
		sqlStr = sqlStr & " 		and dv.useYN = 'Y' " & vbCrLf
		sqlStr = sqlStr & " WHERE 1=1 " & addStr & " " & vbCrLf
		'response.write sqlStr &"<br>"			
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트 
		sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage) & vbCrLf
		sqlStr = sqlStr & " u.username, w.empno, p.part_name, Convert(varchar(10),w.workdate,120) as workdate, c.startwork, c.endwork, w.workcode, " & vbCrLf
		sqlStr = sqlStr & " isNull(l.InTime,'') AS InTime, isNull(l.OutTime,'') AS OutTime, isNull(pl.placeiname,'') AS placeiname, isNull(su.shopname,'') AS shopname, isNull(dv.departmentNameFull,'') AS departmentNameFull " & vbCrLf
		sqlStr = sqlStr & " FROM [db_partner].[dbo].[tbl_offshop_employee_workschedule] AS w " & vbCrLf
		sqlStr = sqlStr & " 	INNER JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS u ON w.empno = u.empno " & vbCrLf
		sqlStr = sqlStr & " 	INNER JOIN [db_partner].[dbo].[tbl_partInfo] AS p ON w.part_sn = p.part_sn " & vbCrLf
		sqlStr = sqlStr & " 	INNER JOIN [db_partner].[dbo].[tbl_offshop_employee_workcode] AS c ON w.workcode = c.workcode " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN " & vbCrLf
		sqlStr = sqlStr & " 	( " & vbCrLf
		sqlStr = sqlStr & " 		SELECT yyyymmdd,empno,placeid,  " & vbCrLf
		sqlStr = sqlStr & " 			min(CASE WHEN inoutType=1 THEN '2901-12-31' ELSE inoutTime END) as InTime, " & vbCrLf
		sqlStr = sqlStr & " 			max(CASE WHEN inoutType=0 THEN '1900-01-01' ELSE inoutTime END) as OutTime " & vbCrLf
		sqlStr = sqlStr & " 		FROM [db_partner].[dbo].[tbl_user_inouttime_log] " & vbCrLf
		sqlStr = sqlStr & " 		WHERE yyyymmdd Between '" & FRectWorkDate1 & "' AND '" & FRectWorkDate2 & "' GROUP BY empno, yyyymmdd, placeid " & vbCrLf
		sqlStr = sqlStr & " 	) AS l ON l.empno = w.empno AND Convert(varchar(10),w.workdate,120) = l.YYYYMMDD " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_partner].[dbo].[tbl_user_inouttime_place] AS pl ON l.placeid = pl.placeid " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_partner].[dbo].[tbl_partner_shopuser] AS ps ON w.empno = ps.empno and ps.firstisusing = 'Y' " & vbCrLf
		sqlStr = sqlStr & " 	LEFT JOIN [db_shop].[dbo].[tbl_shop_user] AS su ON ps.shopid = su.userid " & vbCrLf
		sqlStr = sqlStr & " 	left join db_partner.dbo.vw_user_department dv " & vbCrLf
		sqlStr = sqlStr & " 	on " & vbCrLf
		sqlStr = sqlStr & " 		1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 		and u.department_id = dv.cid " & vbCrLf
		sqlStr = sqlStr & " 		and dv.useYN = 'Y' " & vbCrLf
		sqlStr = sqlStr & " WHERE 1=1 " & addStr & " " & vbCrLf
		sqlStr = sqlStr & " ORDER BY u." & FRectOrderBy & " ASC, Convert(varchar(10),w.workdate,120) ASC " & vbCrLf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cEmployeeManagementClass_oneitem

				FItemList(i).FUserName		= rsget("username")
				FItemList(i).FEmpNO			= rsget("empno")
				FItemList(i).FPartName		= rsget("part_name")
				FItemList(i).FPlaceName		= rsget("placeiname")
				FItemList(i).FWorkDate		= rsget("workdate")
				FItemList(i).FStartWork		= fnChangeTimeType(rsget("startwork"))
				FItemList(i).FEndWork		= fnChangeTimeType(rsget("endwork"))
				FItemList(i).FWorkTime		= fnWorkTimeCalc(rsget("startwork"), rsget("endwork"))
				FItemList(i).FWorkCode		= rsget("workcode")
				FItemList(i).FInTime		= rsget("InTime")
				FItemList(i).FOutTime		= rsget("OutTime")
				FItemList(i).FShopName		= rsget("shopname")
				FItemList(i).FdepartmentNameFull		= rsget("departmentNameFull")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub


	public function fWorkCodeList			'근무코드리스트
	dim i , sql

	sql = "SELECT "
	sql = sql & "	w.workcode, w.startwork, w.endwork "
	sql = sql & " FROM [db_partner].[dbo].[tbl_offshop_employee_workcode] AS w "
	sql = sql & " ORDER BY w.workcode ASC "

	rsget.open sql,dbget,1

	FTotalCount = rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsget.Eof Then
		Do Until rsget.Eof
			set FList(i) = new cEmployeeManagementClass_oneitem
				FList(i).FWorkCode		= rsget("workcode")
				FList(i).FStartWork		= rsget("startwork")
				FList(i).FEndWork		= rsget("endwork")

		rsget.movenext
		i = i + 1
		Loop
	End If

	rsget.close
	end function


	public function fWorkCodeView			'근무코드보기
	dim i , sql

	sql = "SELECT "
	sql = sql & "	w.workcode, w.startwork, w.endwork "
	sql = sql & " FROM [db_partner].[dbo].[tbl_offshop_employee_workcode] AS w WHERE w.workcode = '" & FRectWorkCode & "' "
	sql = sql & " ORDER BY w.workcode ASC "

	rsget.open sql,dbget,1

	If Not rsget.Eof Then
		set FOneItem = new cEmployeeManagementClass_oneitem
		FOneItem.FWorkCode		= rsget("workcode")
		FOneItem.FStartWork		= rsget("startwork")
		FOneItem.FEndWork		= rsget("endwork")
	End If

	rsget.close
	end function


	Private Sub Class_Initialize()
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

end class


Function fnChangeTimeType(t)
	If t = "" Then
		fnChangeTimeType = ""
	Else
		If IsNumeric(t) = True Then
			fnChangeTimeType = fnCalculateTime(t)
		Else
			fnChangeTimeType = t
		End If
	End If
End Function

Function fnCalculateTime(t)
	Dim vTemp, vTime, vMinute

	vTemp = t
	IF t >= 1440 Then
		vTemp = t - 1440
	End If

	vTime = Fix(vTemp/60)
	vMinute = vTemp mod 60

	fnCalculateTime = TwoNumber(vTime) & ":" & TwoNumber(vMinute)

End Function

Function fnWorkTimeCalc(s,e)
	Dim vTemp
	If e <> "" Then
		vTemp = fnChangeTimeType(e-s)
		fnWorkTimeCalc = Split(vTemp,":")(0)
		IF Split(vTemp,":")(1) <> "00" Then
			fnWorkTimeCalc = vTemp
		End IF
	Else
		fnWorkTimeCalc = ""
	End If
End Function

Function fnDatetimeToHourMinute(t)
	fnDatetimeToHourMinute = TwoNumber(Hour(t)) & ":" & TwoNumber(Minute(t))
End Function
%>
