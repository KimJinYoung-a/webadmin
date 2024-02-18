<%

Class CTSKeywordlItem

	public Fidx
	public Ftextname
	Public FsortNo
	public Fregdate
	Public Ftitle
	Public Flinkinfo
	Public FisUsing
	Public Fenddate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CSearchKeyWord
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectIdx
	public FRectUsing
	public FRectSearch

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function GetSearchKeyWord()
		dim sqlStr, addSql, i

		if FRectUsing<>"" then
			addSql = " Where isusing='" & FRectUsing & "'"
		else
			addSql = " Where isusing='Y'"
		end if

		if FRectSearch<>"" then
			addSql = addSql & " and textname like '%" & FRectSearch & "%'"
		end if

		if FRectIdx<>"" then
			addSql = addSql & " and idx='" & FRectIdx & "'"
		end if

		sqlStr = "select count(idx) as cnt from [db_sitemaster].[dbo].tbl_mainTextissue"
		sqlStr = sqlStr + addSql

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " idx,textname,sortNo,linkinfo,isUsing,regdate , enddate" + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_mainTextissue" + addSql + vbcrlf
		sqlStr = sqlStr + " order by sortNo asc, idx desc"

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
				set FItemList(i) = new CTSKeywordlItem

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Ftextname	= rsget("textname")
				FItemList(i).FsortNo	= rsget("sortNo")
				FItemList(i).Flinkinfo	= rsget("linkinfo")
				FItemList(i).FisUsing	= rsget("isUsing")
				FItemList(i).Fregdate	= rsget("regdate")
				FItemList(i).Fenddate= rsget("enddate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Function

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
