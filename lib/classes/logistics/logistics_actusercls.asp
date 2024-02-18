<%

Class CActUserItem
    public Fidx
	public Fempno
	public Fusername
	public Factgubun
	public Frefcode
	public Fregdate
	public FworkSecond

	public FitemNo
	public FtotalCount
	public FCheckCount

	public function GetActGubunName()
		Select Case Factgubun
			Case "onlineOrderChulgo"
				GetActGubunName = "온라인출고"
			Case "onlineOrderChulgoForce"
				GetActGubunName = "온라인출고(재출력)"
			Case "onLineIpgoSongjangno"
				GetActGubunName = "온라인입고 송장입력"
			Case "onLineIpgoCheck"
				GetActGubunName = "온라인입고 검품"
			Case "onLineIpgoRackIpgo"
				GetActGubunName = "온라인입고 랙입고"
			Case "onlineOrderMisend"
				GetActGubunName = "온라인 미배등록"
			Case "offlineOrderChulgo"
				GetActGubunName = "오프라인샵출고"
			Case "onlineOrderPickup"
				GetActGubunName = "온라인 픽업"
			Case Else
				GetActGubunName = Factgubun
		End Select
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CActUser
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

	public Sub GetActUserList
		dim sqlStr, addSqlStr, i
		dim minidx : minidx = 0
		''idx, empno, username, actgubun, refcode, regdate

		if (FRectStartdate <> "") then
			sqlStr = " select IsNull(min(idx),0) as minidx from db_log.dbo.tbl_logics_act_log where regdate >= '" & FRectStartdate & "' "
			rsget.Open sqlStr,dbget,1
			if  not rsget.EOF  then
				minidx = rsget("minidx")
			end if
			rsget.Close
		end if


		addSqlStr = ""

		if (FRectStartdate <> "") then
			addSqlStr = addSqlStr + " and l.regdate >= '" + CStr(FRectStartdate) + "' "
			addSqlStr = addSqlStr + " and l.idx >= " & minidx & " "
		end if

		if (FRectEndDate <> "") then
			addSqlStr = addSqlStr + " and l.regdate < '" + CStr(FRectEndDate) + "' "
		end if

		if (FRectActGubun <> "") then
			addSqlStr = addSqlStr + " and l.actgubun = '" + CStr(FRectActGubun) + "' "
		end if

		if (FRectRefCode <> "") then
			addSqlStr = addSqlStr + " and l.refcode = '" + CStr(FRectRefCode) + "' "
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			addSqlStr = addSqlStr + " and " + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
		end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_log.dbo.tbl_logics_act_log l "
		sqlStr = sqlStr + " 	left join db_partner.dbo.tbl_user_tenbyten u "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		l.empno = u.empno "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr
		'response.write sqlStr &"<br>"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if (FTotalCount = 0) then
			exit sub
		end if


		sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " l.idx, l.empno, u.username, l.actgubun, l.refcode, l.regdate "

		sqlStr = sqlStr + " ,(case "
		sqlStr = sqlStr + " 	when l.actgubun in ('onLineIpgoCheck', 'onLineIpgoRackIpgo') then [db_storage].[dbo].[UF_GetOrderCheckItemNo](l.refcode) "
		sqlStr = sqlStr + " 	when l.actgubun = 'onlineOrderChulgo' then [db_order].[dbo].[UF_GetTenbeaOrderItemNo](l.refcode) "
		sqlStr = sqlStr + " 	when l.actgubun in ('offlineOrderChulgo') then [db_storage].[dbo].[UF_GetShopChulgoItemNo](l.refcode) "
		sqlStr = sqlStr + " 	else 0 end) as itemno "
		sqlStr = sqlStr + " 	, left(convert(varchar, DATEADD(second, IsNull(workSecond,0), 0), 114),8) as workSecond "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_log.dbo.tbl_logics_act_log l "
		sqlStr = sqlStr + " 	left join db_partner.dbo.tbl_user_tenbyten u "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		l.empno = u.empno "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr
		sqlStr = sqlStr + " order by l.idx desc "
		''response.write sqlStr &"<br>"

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
				set FItemList(i) = new CActUserItem

				''idx, empno, username, actgubun, refcode, regdate

				FItemList(i).Fidx = rsget("idx")
				FItemList(i).Fempno = rsget("empno")
				FItemList(i).Fusername = rsget("username")
				FItemList(i).Factgubun = rsget("actgubun")
				FItemList(i).Frefcode = rsget("refcode")
				FItemList(i).Fregdate = rsget("regdate")

				FItemList(i).FitemNo = rsget("itemno")
				FItemList(i).FworkSecond = rsget("workSecond")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
    end Sub

	public Sub GetActUserStatisticList
		dim sqlStr, addSqlStr, i
		dim tmpSql

		addSqlStr = ""

		if (FRectStartdate <> "") then
			addSqlStr = addSqlStr + " and l.regdate >= '" + CStr(FRectStartdate) + "' "
		end if

		if (FRectEndDate <> "") then
			addSqlStr = addSqlStr + " and l.regdate < '" + CStr(FRectEndDate) + "' "
		end if

		if (FRectActGubun <> "") then
			addSqlStr = addSqlStr + " and l.actgubun = '" + CStr(FRectActGubun) + "' "
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			addSqlStr = addSqlStr + " and " + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
		end if

		sqlStr = " select l.empno, u.username, l.actgubun "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_log.dbo.tbl_logics_act_log l "
		sqlStr = sqlStr + " 	left join db_partner.dbo.tbl_user_tenbyten u "
		sqlStr = sqlStr + " 	on l.empno = u.empno "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr
		sqlStr = sqlStr + " group by "
		sqlStr = sqlStr + " 	l.empno, u.username, l.actgubun "

		sqlStr = "select count(*) as cnt from (" + sqlStr + ") T "
		'response.write sqlStr &"<br>"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " l.empno, u.username, l.actgubun, count(*) as cnt, left(convert(varchar, DATEADD(second, avg(IsNull(workSecond,0)), 0), 114),8) as workSecond "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_log.dbo.tbl_logics_act_log l "
		sqlStr = sqlStr + " 	left join db_partner.dbo.tbl_user_tenbyten u "
		sqlStr = sqlStr + " 	on l.empno = u.empno "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr
		sqlStr = sqlStr + " group by "
		sqlStr = sqlStr + " 	l.empno, u.username, l.actgubun "
		sqlStr = sqlStr + " order by count(*) desc "
		''response.write sqlStr &"<br>"

		if (FRectActGubun = "onLineIpgoCheck") then
			tmpSql = " select top " + Cstr(FPageSize * FCurrPage) + " T1.empno, T1.username, T1.actgubun, T1.cnt, IsNull(T2.chk,0) as chk, T1.workSecond "
			tmpSql = tmpSql + " from ( "
			tmpSql = tmpSql + sqlStr
			tmpSql = tmpSql + " ) T1 "
			tmpSql = tmpSql + " left join ( "
			tmpSql = tmpSql + " 	select l.empno, IsNull(sum(d.checkitemno),0) as chk "
			tmpSql = tmpSql + " 	from "
			tmpSql = tmpSql + " 		db_log.dbo.tbl_logics_act_log l "
			tmpSql = tmpSql + " 		join [db_storage].[dbo].[tbl_ordersheet_master] m "
			tmpSql = tmpSql + " 		on "
			tmpSql = tmpSql + " 			l.refcode = m.baljucode and empno = checkusersn "
			tmpSql = tmpSql + " 		join [db_storage].[dbo].[tbl_ordersheet_detail] d "
			tmpSql = tmpSql + " 		on "
			tmpSql = tmpSql + " 			m.idx = d.masteridx "
			tmpSql = tmpSql + " 	where "
			tmpSql = tmpSql + " 		1 = 1 "
			tmpSql = tmpSql + " 		and m.deldt is NULL "
			tmpSql = tmpSql + " 		and d.deldt is NULL "
			tmpSql = tmpSql + " 		and m.divcode in ('301', '302') "
			tmpSql = tmpSql + " 		and l.actgubun = 'onLineIpgoCheck' "
			tmpSql = tmpSql + " 		and l.regdate >= '" & FRectStartdate & "' "
			tmpSql = tmpSql + " 		and l.regdate < '" & FRectEndDate & "' "
			tmpSql = tmpSql + " 	group by "
			tmpSql = tmpSql + " 		l.empno "
			tmpSql = tmpSql + " ) T2 "
			tmpSql = tmpSql + " on "
			tmpSql = tmpSql + " 	T1.empno = T2.empno "
			tmpSql = tmpSql + " order by "
			tmpSql = tmpSql + " 	T1.cnt desc "

			sqlStr = tmpSql
		elseif (FRectActGubun = "onLineIpgoRackIpgo") then
			tmpSql = " select top " + Cstr(FPageSize * FCurrPage) + " T1.empno, T1.username, T1.actgubun, T1.cnt, IsNull(T2.chk,0) as chk, T1.workSecond "
			tmpSql = tmpSql + " from ( "
			tmpSql = tmpSql + sqlStr
			tmpSql = tmpSql + " ) T1 "
			tmpSql = tmpSql + " left join ( "
			tmpSql = tmpSql + " 	select l.empno, IsNull(sum(d.realitemno),0) as chk "
			tmpSql = tmpSql + " 	from "
			tmpSql = tmpSql + " 		db_log.dbo.tbl_logics_act_log l "
			tmpSql = tmpSql + " 		join [db_storage].[dbo].[tbl_ordersheet_master] m "
			tmpSql = tmpSql + " 		on "
			tmpSql = tmpSql + " 			l.refcode = m.baljucode and empno = rackipgousersn "
			tmpSql = tmpSql + " 		join [db_storage].[dbo].[tbl_ordersheet_detail] d "
			tmpSql = tmpSql + " 		on "
			tmpSql = tmpSql + " 			m.idx = d.masteridx "
			tmpSql = tmpSql + " 	where "
			tmpSql = tmpSql + " 		1 = 1 "
			tmpSql = tmpSql + " 		and m.deldt is NULL "
			tmpSql = tmpSql + " 		and d.deldt is NULL "
			tmpSql = tmpSql + " 		and m.divcode in ('301', '302') "
			tmpSql = tmpSql + " 		and l.actgubun = 'onLineIpgoRackIpgo' "
			tmpSql = tmpSql + " 		and l.regdate >= '" & FRectStartdate & "' "
			tmpSql = tmpSql + " 		and l.regdate < '" & FRectEndDate & "' "
			tmpSql = tmpSql + " 	group by "
			tmpSql = tmpSql + " 		l.empno "
			tmpSql = tmpSql + " ) T2 "
			tmpSql = tmpSql + " on "
			tmpSql = tmpSql + " 	T1.empno = T2.empno "
			tmpSql = tmpSql + " order by "
			tmpSql = tmpSql + " 	T1.cnt desc "

			sqlStr = tmpSql
		end if

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
				set FItemList(i) = new CActUserItem

				''idx, empno, username, actgubun, refcode, regdate

				FItemList(i).Fempno = rsget("empno")
				FItemList(i).Fusername = rsget("username")
				FItemList(i).Factgubun = rsget("actgubun")
				FItemList(i).FtotalCount = rsget("cnt")
				FItemList(i).FworkSecond = rsget("workSecond")

				if (FRectActGubun = "onLineIpgoCheck") or (FRectActGubun = "onLineIpgoRackIpgo") then
					FItemList(i).FCheckCount = rsget("chk")
				else
					FItemList(i).FCheckCount = 0
				end if

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
    end Sub

    Private Sub Class_Initialize()
		ReDim FItemList(0)

		FCurrPage		= 1
		FPageSize 		= 20
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
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
End Class



%>
