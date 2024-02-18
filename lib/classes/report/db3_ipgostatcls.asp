<%

Class CActUserStatItem
	public Fyyyymmdd
	public Fempno
	public Fusername
	public Fon_ipgo_checkno
	public Fon_ipgo_rackipgono
	public Ftotpay

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
End Class

Class CActUserStat
	public FItemList()
    public FOneItem

    public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount

	public FRectActGubun
	public FRectRefCode
	public FRectSearchField
	public FRectSearchText

	public FRectStartdate
	public FRectEndDate
	public FRectDateGubun

	'// 매월 26일 아래 쿼리 실행할 것!!
	'// exec [db_datamart].[dbo].[usp_Ten_Make_LogicsActLogSum] '2016-07-26', '2016-08-25'
	public Sub GetActUserStatList
		dim sqlStr, addSqlStr, tmpStr, i

		if (FRectStartdate <> "") then
			addSqlStr = addSqlStr + " and s.yyyymmdd >= '" + CStr(FRectStartdate) + "' "
		end if

		if (FRectEndDate <> "") then
			addSqlStr = addSqlStr + " and s.yyyymmdd <= '" + CStr(FRectEndDate) + "' "
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			addSqlStr = addSqlStr + " and " + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
		end if

		if FRectDateGubun = "yyyymm" then
			tmpStr = " select s.empno "
			tmpStr = tmpStr + " from [db_datamart].[dbo].[tbl_logics_act_log_SUM] s "
			tmpStr = tmpStr + " 	left join db_partner.dbo.tbl_user_tenbyten u "
			tmpStr = tmpStr + " 	on "
			tmpStr = tmpStr + " 		s.empno = u.empno "
			tmpStr = tmpStr + " where "
			tmpStr = tmpStr + " 	1 = 1 "
			tmpStr = tmpStr + addSqlStr
			tmpStr = tmpStr + " group by s.empno "
		else
			tmpStr = " select s.empno "
			tmpStr = tmpStr + " from [db_datamart].[dbo].[tbl_logics_act_log_SUM] s "
			tmpStr = tmpStr + " 	left join db_partner.dbo.tbl_user_tenbyten u "
			tmpStr = tmpStr + " 	on "
			tmpStr = tmpStr + " 		s.empno = u.empno "
			tmpStr = tmpStr + " where "
			tmpStr = tmpStr + " 	1 = 1 "
			tmpStr = tmpStr + addSqlStr
			tmpStr = tmpStr + " group by yyyymmdd, s.empno "
		end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	( " + tmpStr + " ) T "
		''response.write sqlStr &"<br>"

		db3_rsget.Open sqlStr,db3_dbget,1
			FTotalCount = db3_rsget("cnt")
		db3_rsget.Close

		if (FTotalCount = 0) then
			exit sub
		end if


		if FRectDateGubun = "yyyymm" then
			sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " '' as yyyymmdd, s.empno, u.username, sum(s.on_ipgo_checkno) as on_ipgo_checkno, sum(s.on_ipgo_rackipgono) as on_ipgo_rackipgono, IsNull((select max(p.totpay) from db_partner.dbo.tbl_user_monthlypay p where s.empno = p.empno and p.yyyymm = '" & Left(FRectEndDate, 7) & "'),0) as totpay "
		else
			sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " s.yyyymmdd, s.empno, u.username, sum(s.on_ipgo_checkno) as on_ipgo_checkno, sum(s.on_ipgo_rackipgono) as on_ipgo_rackipgono, 0 as totpay "
		end if


		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_datamart].[dbo].[tbl_logics_act_log_SUM] s "
		sqlStr = sqlStr + " 	left join db_partner.dbo.tbl_user_tenbyten u "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		s.empno = u.empno "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr

		if FRectDateGubun = "yyyymm" then
			sqlStr = sqlStr + " group by s.empno, u.username "
			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	s.empno "
		else
			sqlStr = sqlStr + " group by s.yyyymmdd, s.empno, u.username "
			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	s.yyyymmdd desc, s.empno "
		end if


		''response.write sqlStr &"<br>"

		db3_rsget.pagesize = FPageSize
		db3_rsget.Open sqlStr,db3_dbget,1

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
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.EOF
				set FItemList(i) = new CActUserStatItem

				''s.yyyymmdd, s.empno, u.username, s.on_ipgo_checkno, s.on_ipgo_rackipgono

				FItemList(i).Fyyyymmdd = db3_rsget("yyyymmdd")
				FItemList(i).Fempno = db3_rsget("empno")
				FItemList(i).Fusername = db2html(db3_rsget("username"))
				FItemList(i).Fon_ipgo_checkno = db3_rsget("on_ipgo_checkno")
				FItemList(i).Fon_ipgo_rackipgono = db3_rsget("on_ipgo_rackipgono")
				FItemList(i).Ftotpay = db3_rsget("totpay")

				db3_rsget.movenext
				i=i+1
			loop
		end if
		db3_rsget.Close
	End Sub

    Private Sub Class_Initialize()
		ReDim FItemList(0)

		FCurrPage		= 1
		FPageSize 		= 20
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
	End Sub

	Private Sub Class_Terminate()
		'
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

end Class

%>
