<%
'###########################################################
' Description : 고객블랙리스트 클래스
' Hieditor : 2014.03.06 한용민 생성
'###########################################################

Class CCompanyCalendarItem
	Private Sub Class_Initialize()
		''
	End Sub

	Private Sub Class_Terminate()
		''
	End Sub

	''idx, title, contents, startDate, endDate, importantLevel, openLevel, useYN, reguserid, modiuserid, regdate, lastupdate

	public Fidx
	public Ftitle
	public Fcontents
	public FstartDate
	public FendDate
	public FimportantLevel
	public FopenLevel
	public FuseYN
	public Freguserid
	public Fmodiuserid
	public Fregdate
	public Flastupdate
	public FpName

	public function GetImportantLevelName()
		if FimportantLevel = "0" then
			GetImportantLevelName = "없음"
		elseif FimportantLevel = "10" then
			GetImportantLevelName = "낮음"
		elseif FimportantLevel = "20" then
			GetImportantLevelName = "보통"
		elseif FimportantLevel = "30" then
			GetImportantLevelName = "높음"
		else
			GetImportantLevelName = FimportantLevel
		end if
	end function

	public function GetOpenLevelName()
		if FopenLevel = "0" then
			GetOpenLevelName = "없음"
		elseif FopenLevel = "10" then
			GetOpenLevelName = "부서공지"
		elseif FopenLevel = "20" then
			GetOpenLevelName = "전체공지"
		else
			GetOpenLevelName = FimportantLevel
		end if
	end function

end class

Class CCompanyCalendarDetailItem
	Private Sub Class_Initialize()
		''
	End Sub

	Private Sub Class_Terminate()
		''
	End Sub

	''o.department_id, p.departmentNameFull, o.empno, u.username, p2.departmentname, p3.posit_name
	''idx, calIdx, departmentId, empno, useYN, regdate, lastupdate

	public Fidx
	public FcalIdx
	public Fdepartment_id
	public Fempno
	public FdepartmentNameFull
	public Fusername
	public Fdepartmentname
	public Fposit_name
	public Fregdate
	public Flastupdate

end class



''idx, calIdx, departmentId, empno, useYN, regdate, lastupdate

class CCompanyCalendar
	public FItemList()
	public FOneItem

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectIdx
	public FRectUseYN
	public FRectStartDate
	public FRectEndDate
	public FRectDepartmentID
	public FRectEmpNO
	
	public FRectYear
	public FRectMonth

	public sub getCompanyCalendarList()
		dim sqlStr, i, addSql

		if FRectUseYN <> "" then
			addSql = addSql & " and c.useYN = '"& FRectUseYN &"'" + vbCrLf
		end if

		If (FRectYear <> "") and (FRectMonth <> "") Then
			addSql = addSql & " and convert(varchar(7),c.startDate,20) <= '" & FRectYear & "-" & FRectMonth & "'" & vbcrlf
			addSql = addSql & " and convert(varchar(7),c.endDate,20) >= '" & FRectYear & "-" & FRectMonth & "'" & vbcrlf
		End If

		if FRectStartDate <> "" then
			addSql = addSql & " and c.startDate >= '"& Left(FRectStartDate,10) &"'" + vbCrLf
		end if

		if FRectEndDate <> "" then
			addSql = addSql & " and c.endDate <= '"& Left(FRectEndDate,10) &"'" + vbCrLf
		end if

		'// Count
		sqlStr = " select count(c.idx) as cnt " + vbCrLf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_compCalendar c with (nolock)" + vbCrLf
		'2015-09-24 김진영 하단 JOIN문 추가
		sqlStr = sqlStr & " JOIN (select tt.userid, u.departmentName from db_partner.dbo.tbl_user_tenbyten as tt with (nolock) join db_partner.dbo.tbl_user_department as u with (nolock) on tt.department_id = u.cid) as K on c.reguserid=K.userid "

		if FRectDepartmentID <> "" then
			sqlStr = sqlStr & " join ( " + vbCrLf
			sqlStr = sqlStr & " select o.calIdx " + vbCrLf
			sqlStr = sqlStr & " from " + vbCrLf
			sqlStr = sqlStr & " 	[db_partner].[dbo].[tbl_compCalendar_OpenList] o with (nolock)" + vbCrLf
			sqlStr = sqlStr & " 	join [db_partner].[dbo].[vw_user_department_v2] v with (nolock)" + vbCrLf
			sqlStr = sqlStr & " 	on " + vbCrLf
			sqlStr = sqlStr & " 		1 = 1 " + vbCrLf
			sqlStr = sqlStr & " 		and o.department_id is not NULL " + vbCrLf
			sqlStr = sqlStr & " 		and o.useYN = 'Y' " + vbCrLf
			sqlStr = sqlStr & " 		and v.cidArr like '%" & FRectDepartmentID & "%' " + vbCrLf
			sqlStr = sqlStr & " 		and v.cidArr like '%' + convert(varchar,o.department_id) + '%' " + vbCrLf
			sqlStr = sqlStr & " group by o.calIdx " + vbCrLf
			sqlStr = sqlStr & " ) oT " + vbCrLf
			sqlStr = sqlStr & " on " + vbCrLf
			sqlStr = sqlStr & " 	c.idx = oT.calIdx " + vbCrLf
		end if

		if FRectEmpNO <> "" then
			sqlStr = sqlStr & " join ( " + vbCrLf
			sqlStr = sqlStr & " select o.calIdx " + vbCrLf
			sqlStr = sqlStr & " from " + vbCrLf
			sqlStr = sqlStr & " 	[db_partner].[dbo].[tbl_compCalendar_OpenList] o with (nolock)" + vbCrLf
			sqlStr = sqlStr & " 	left join ( " + vbCrLf
			sqlStr = sqlStr & " 		select v.cidArr " + vbCrLf
			sqlStr = sqlStr & " 		from " + vbCrLf
			sqlStr = sqlStr & " 			[db_partner].[dbo].[vw_user_department_v2] v with (nolock)" + vbCrLf
			sqlStr = sqlStr & " 			join [db_partner].[dbo].tbl_user_tenbyten t with (nolock)" + vbCrLf
			sqlStr = sqlStr & " 			on " + vbCrLf
			sqlStr = sqlStr & " 				1 = 1 " + vbCrLf
			sqlStr = sqlStr & " 				and t.empno = '" & FRectEmpNO & "' " + vbCrLf
			sqlStr = sqlStr & " 				and v.cid = t.department_id " + vbCrLf
			sqlStr = sqlStr & " 	) V " + vbCrLf
			sqlStr = sqlStr & " 	on " + vbCrLf
			sqlStr = sqlStr & " 		1 = 1 " + vbCrLf
			sqlStr = sqlStr & " 		and v.cidArr like '%' + convert(varchar,o.department_id) + '%' " + vbCrLf
			sqlStr = sqlStr & " where " + vbCrLf
			sqlStr = sqlStr & " 	1 = 1 " + vbCrLf
			sqlStr = sqlStr & " 	and (o.empno = '" & FRectEmpNO & "' or v.cidArr is not NULL) " + vbCrLf
			sqlStr = sqlStr & " 	and o.useYN = 'Y' " + vbCrLf
			sqlStr = sqlStr & " group by o.calIdx " + vbCrLf
			sqlStr = sqlStr & " ) T " + vbCrLf
			sqlStr = sqlStr & " on " + vbCrLf
			sqlStr = sqlStr & " 	c.idx = T.calIdx " + vbCrLf
		end if

		sqlStr = sqlStr & " where 1 = 1 " & addSql + vbCrLf
		'response.write sqlStr &"<br>"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub


		'// List
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbCrLf
		sqlStr = sqlStr & " c.idx, c.title, c.contents, c.startDate, c.endDate, c.importantLevel, c.openLevel, c.useYN, c.reguserid, c.modiuserid, c.regdate, c.lastupdate,K.departmentName as pName " + vbCrLf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_compCalendar c " + vbCrLf
		'2015-09-24 김진영 하단 JOIN문 추가
		sqlStr = sqlStr & " JOIN (select tt.userid, u.departmentName from db_partner.dbo.tbl_user_tenbyten as tt with (nolock) join db_partner.dbo.tbl_user_department as u with (nolock) on tt.department_id = u.cid) as K on c.reguserid=K.userid "

		if FRectDepartmentID <> "" then
			sqlStr = sqlStr & " join ( " + vbCrLf
			sqlStr = sqlStr & " select o.calIdx " + vbCrLf
			sqlStr = sqlStr & " from " + vbCrLf
			sqlStr = sqlStr & " 	[db_partner].[dbo].[tbl_compCalendar_OpenList] o with (nolock)" + vbCrLf
			sqlStr = sqlStr & " 	join [db_partner].[dbo].[vw_user_department_v2] v with (nolock)" + vbCrLf
			sqlStr = sqlStr & " 	on " + vbCrLf
			sqlStr = sqlStr & " 		1 = 1 " + vbCrLf
			sqlStr = sqlStr & " 		and o.department_id is not NULL " + vbCrLf
			sqlStr = sqlStr & " 		and o.useYN = 'Y' " + vbCrLf
			sqlStr = sqlStr & " 		and v.cidArr like '%" & FRectDepartmentID & "%' " + vbCrLf
			sqlStr = sqlStr & " 		and v.cidArr like '%' + convert(varchar,o.department_id) + '%' " + vbCrLf
			sqlStr = sqlStr & " group by o.calIdx " + vbCrLf
			sqlStr = sqlStr & " ) oT " + vbCrLf
			sqlStr = sqlStr & " on " + vbCrLf
			sqlStr = sqlStr & " 	c.idx = oT.calIdx " + vbCrLf
		end if

		if FRectEmpNO <> "" then
			sqlStr = sqlStr & " join ( " + vbCrLf
			sqlStr = sqlStr & " select o.calIdx " + vbCrLf
			sqlStr = sqlStr & " from " + vbCrLf
			sqlStr = sqlStr & " 	[db_partner].[dbo].[tbl_compCalendar_OpenList] o with (nolock)" + vbCrLf
			sqlStr = sqlStr & " 	left join ( " + vbCrLf
			sqlStr = sqlStr & " 		select v.cidArr " + vbCrLf
			sqlStr = sqlStr & " 		from " + vbCrLf
			sqlStr = sqlStr & " 			[db_partner].[dbo].[vw_user_department_v2] v with (nolock)" + vbCrLf
			sqlStr = sqlStr & " 			join [db_partner].[dbo].tbl_user_tenbyten t with (nolock)" + vbCrLf
			sqlStr = sqlStr & " 			on " + vbCrLf
			sqlStr = sqlStr & " 				1 = 1 " + vbCrLf
			sqlStr = sqlStr & " 				and t.empno = '" & FRectEmpNO & "' " + vbCrLf
			sqlStr = sqlStr & " 				and v.cid = t.department_id " + vbCrLf
			sqlStr = sqlStr & " 	) V " + vbCrLf
			sqlStr = sqlStr & " 	on " + vbCrLf
			sqlStr = sqlStr & " 		1 = 1 " + vbCrLf
			sqlStr = sqlStr & " 		and v.cidArr like '%' + convert(varchar,o.department_id) + '%' " + vbCrLf
			sqlStr = sqlStr & " where " + vbCrLf
			sqlStr = sqlStr & " 	1 = 1 " + vbCrLf
			sqlStr = sqlStr & " 	and (o.empno = '" & FRectEmpNO & "' or v.cidArr is not NULL) " + vbCrLf
			sqlStr = sqlStr & " 	and o.useYN = 'Y' " + vbCrLf
			sqlStr = sqlStr & " group by o.calIdx " + vbCrLf
			sqlStr = sqlStr & " ) T " + vbCrLf
			sqlStr = sqlStr & " on " + vbCrLf
			sqlStr = sqlStr & " 	c.idx = T.calIdx " + vbCrLf
		end if

		sqlStr = sqlStr & " where 1 = 1 " & addSql + vbCrLf
		sqlStr = sqlStr & " order by c.idx Desc" + vbCrLf
		''response.write sqlStr &"<br>"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage + 1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CCompanyCalendarItem

				FItemList(i).Fidx = rsget("idx")
				FItemList(i).Ftitle = db2html(rsget("title"))
				FItemList(i).Fcontents = db2html(rsget("contents"))
				FItemList(i).FstartDate = rsget("startDate")
				FItemList(i).FendDate = rsget("endDate")
				FItemList(i).FimportantLevel = rsget("importantLevel")
				FItemList(i).FopenLevel = rsget("openLevel")
				FItemList(i).FuseYN = rsget("useYN")
				FItemList(i).Freguserid = rsget("reguserid")
				FItemList(i).Fmodiuserid = rsget("modiuserid")
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).Flastupdate = rsget("lastupdate")
				FItemList(i).FpName = rsget("pName")
				

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub getCompanyCalendarItem()
		dim sqlStr, addSql

		if frectidx <> "" then
			addSql = addSql & " and c.idx = "& FRectIdx & ""
		else
			set FOneItem = new CCompanyCalendarItem
			exit sub
		end if

		sqlStr = "select top 1 " + vbCrLf
		sqlStr = sqlStr & " c.idx, c.title, c.contents, c.startDate, c.endDate, c.importantLevel, c.openLevel, c.useYN, c.reguserid, c.modiuserid, c.regdate, c.lastupdate " + vbCrLf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_compCalendar c " + vbCrLf
		sqlStr = sqlStr & " where 1=1 " & addSql + vbCrLf
		'response.write sqlStr &"<br>"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount
		FResultCount = rsget.recordcount

		if not rsget.EOF  then
			set FOneItem = new CCompanyCalendarItem

			FOneItem.Fidx = rsget("idx")
			FOneItem.Ftitle = db2html(rsget("title"))
			FOneItem.Fcontents = db2html(rsget("contents"))
			FOneItem.FstartDate = rsget("startDate")
			FOneItem.FendDate = rsget("endDate")
			FOneItem.FimportantLevel = rsget("importantLevel")
			FOneItem.FopenLevel = rsget("openLevel")
			FOneItem.FuseYN = rsget("useYN")
			FOneItem.Freguserid = rsget("reguserid")
			FOneItem.Fmodiuserid = rsget("modiuserid")
			FOneItem.Fregdate = rsget("regdate")
			FOneItem.Flastupdate = rsget("lastupdate")
		end if
		rsget.Close
	end sub

	public sub getPartOrMemberList()
		dim sqlStr, i, addSql

		addSql = " and o.useYN = 'Y' " + vbCrLf
		addSql = addSql + " and o.calIdx = " & FRectIdx & " " + vbCrLf

		'// Count
		sqlStr = " select count(o.idx) as cnt " + vbCrLf
		sqlStr = sqlStr & " from "
		sqlStr = sqlStr & " 	[db_partner].[dbo].[tbl_compCalendar_OpenList] o "
		sqlStr = sqlStr & " where 1 = 1 " & addSql + vbCrLf
		''response.write sqlStr &"<br>"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub

		'// List
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbCrLf
		sqlStr = sqlStr & " o.idx, o.calIdx, o.department_id, p.departmentNameFull, o.empno, u.username, p2.departmentname, p3.posit_name, o.regdate, o.lastupdate " + vbCrLf
		sqlStr = sqlStr & " from " + vbCrLf
		sqlStr = sqlStr & " 	[db_partner].[dbo].[tbl_compCalendar_OpenList] o " + vbCrLf
		sqlStr = sqlStr & " 	left join db_partner.dbo.vw_user_department p " + vbCrLf
		sqlStr = sqlStr & " 	on " + vbCrLf
		sqlStr = sqlStr & " 		o.department_id = p.cid " + vbCrLf
		sqlStr = sqlStr & " 	left join [db_partner].[dbo].tbl_user_tenbyten u " + vbCrLf
		sqlStr = sqlStr & " 	on " + vbCrLf
		sqlStr = sqlStr & " 		o.empno = u.empno " + vbCrLf
		sqlStr = sqlStr & " 	left join db_partner.dbo.vw_user_department p2 " + vbCrLf
		sqlStr = sqlStr & " 	on " + vbCrLf
		sqlStr = sqlStr & " 		u.department_id = p2.cid " + vbCrLf
		sqlStr = sqlStr & " 	left join db_partner.dbo.tbl_positInfo p3 " + vbCrLf
		sqlStr = sqlStr & " 	on " + vbCrLf
		sqlStr = sqlStr & " 		u.posit_sn = p3.posit_sn " + vbCrLf
		sqlStr = sqlStr & " where 1 = 1 " & addSql + vbCrLf
		sqlStr = sqlStr & " order by o.lastupdate, o.idx " + vbCrLf
		''response.write sqlStr &"<br>"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage + 1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CCompanyCalendarDetailItem

				FItemList(i).Fidx = rsget("idx")
				FItemList(i).FcalIdx = rsget("calIdx")
				FItemList(i).Fdepartment_id = rsget("department_id")
				FItemList(i).Fempno = rsget("empno")
				FItemList(i).FdepartmentNameFull = rsget("departmentNameFull")
				FItemList(i).Fusername = rsget("username")
				FItemList(i).Fdepartmentname = rsget("departmentname")
				FItemList(i).Fposit_name = rsget("posit_name")
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).Flastupdate = rsget("lastupdate")

				rsget.movenext
				i=i+1
			loop
		end if
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

	Private Sub Class_Initialize()
		FCurrPage 		= 1
		FPageSize 		= 50
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

function Drawinvalidgubun1111(selectBoxName, selectedId, changeFlag)
%>
	<select name="<%= selectBoxName %>" <%= changeFlag %>>
		<option value="" <% if selectedId="" then response.write " selected" %>>전체</option>
		<option value="ONEVT" <% if selectedId="ONEVT" then response.write " selected" %>>이벤트</option>
		<option value="ETC" <% if selectedId="ETC" then response.write " selected" %>>기타</option>
	</select>
<%
end function

function getinvalidgubun1111(gubun)
	if gubun="ONEVT" then
		getinvalidgubun1111 = "이벤트"
	elseif gubun="ETC" then
		getinvalidgubun1111 = "기타"
	else
		getinvalidgubun1111 = ""
	end if
End function
%>
