<%
Class CVipCorner

	public Fisusing
	public Fregdate
	Public FevtCode
	Public Fpcimg
	Public Fmaing
	Public Forderby
	Public Fregname
	Public Fmodname
	Public Fmodifydate
	Public Fidx
	Public FevtName
	Public FevtStartDate
	Public FevtEndDate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CVip
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectVSN
	public FRectDiv
	public FRectUsing
	Public FRectidx

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function GetVipCornerList()
		dim sqlStr, addSql, i

		'추가 조건절
		if FRectUsing<>"" Then
			If FRectUsing="all" Then
			Else
				addSql = addSql & " and v.isUsing='" & FRectUsing & "'"
			End If
		end if

		'카운트
		sqlStr = "select count(idx) as cnt"
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_vipCorner as v" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		'목록 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " v.idx, v.evt_code, v.pcimg, v.maing, v.orderby, v.isusing, v.regname, v.modname, v.regdate, v.modifydate, e.evt_name, e.evt_startdate, e.evt_enddate " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_vipCorner v" + vbcrlf
		sqlStr = sqlStr + " inner join db_event.dbo.tbl_event e on v.evt_code = e.evt_code" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		sqlStr = sqlStr + " order by v.orderby asc, v.idx desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CVipCorner
				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).FevtCode		= rsget("evt_code")
				FItemList(i).Fpcimg		= rsget("pcimg")
				FItemList(i).Fmaing	= db2html(rsget("maing"))
				FItemList(i).Forderby	= rsget("orderby")
				FItemList(i).Fregname	= rsget("regname")
				FItemList(i).Fmodname	= rsget("modname")
				FItemList(i).Fmodifydate		= rsget("modifydate")
				FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).FevtName = rsget("evt_name")
				FItemList(i).FevtStartDate = rsget("evt_startdate")
				FItemList(i).FevtEndDate = rsget("evt_enddate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function


	public Function GetVipCornerModify()
		dim sqlStr, addSql, i

		'추가 조건절
		addSql = addSql & " and v.idx='" & FRectidx & "'"

		'목록 접수
		sqlStr = "select top 1 " + vbcrlf
		sqlStr = sqlStr + " v.idx, v.evt_code, v.pcimg, v.maing, v.orderby, v.isusing, v.regname, v.modname, v.regdate, v.modifydate, e.evt_name, e.evt_startdate, e.evt_enddate " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_vipCorner v" + vbcrlf
		sqlStr = sqlStr + " inner join db_event.dbo.tbl_event e on v.evt_code = e.evt_code" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		redim preserve FItemList(1)
		i=0
		if  not rsget.EOF  then
				set FItemList(0) = new CVipCorner
				FItemList(0).Fidx		= rsget("idx")
				FItemList(0).FevtCode		= rsget("evt_code")
				FItemList(0).Fpcimg		= rsget("pcimg")
				FItemList(0).Fmaing	= db2html(rsget("maing"))
				FItemList(0).Forderby	= rsget("orderby")
				FItemList(0).Fregname	= rsget("regname")
				FItemList(0).Fmodname	= rsget("modname")
				FItemList(0).Fmodifydate		= rsget("modifydate")
				FItemList(0).Fisusing		= rsget("isusing")
				FItemList(0).Fregdate		= rsget("regdate")
				FItemList(0).FevtName = rsget("evt_name")
				FItemList(0).FevtStartDate = rsget("evt_startdate")
				FItemList(0).FevtEndDate = rsget("evt_enddate")
		end if

		rsget.Close
	end function


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