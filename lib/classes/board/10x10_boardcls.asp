<%

class CHopeBoardSubItem

	public Fidx
	public Fgubun
	public Fusername
	public Ftitle
	public Fregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function FGubunName()
		if FGubun = "01" then
			FGubunName = "건의"
		elseif FGubun = "02" then
			FGubunName = "제안"
		elseif FGubun = "03" then
			FGubunName = "회람"
		elseif FGubun = "04" then
			FGubunName = "기타"
		end if
	end Function

end Class

Class CHopeBoard

	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

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

		sql = "select count(idx) as cnt "
		sql = sql + " from [db_board].[dbo].tbl_10x10_board"
		sql = sql + " where isusing = 'Y'"

		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

		sql = " select top " + CStr(FPageSize*FCurrPage) + " idx,gubun,username,title,regdate"
		sql = sql + " from [db_board].[dbo].tbl_10x10_board"
		sql = sql + " where isusing = 'Y'"
		sql = sql + " order by regdate desc "

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		        i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CHopeBoardSubItem

				FItemList(i).Fidx     = rsget("idx")
				FItemList(i).Fgubun     = rsget("gubun")
				FItemList(i).Fusername   = rsget("username")
				FItemList(i).Ftitle   =  db2html(rsget("title"))
				FItemList(i).Fregdate    = rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
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
end Class


class CHopeBoardDetail

	public Fidx
	public Fgubun
	public Fuserid
	public Fusername
	public Ftitle
	public Fcontents
	public Fregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub read(byVal v)
		dim sql, i

		sql = " select  idx,gubun,userid,username,title,contents,regdate"
		sql = sql + " from [db_board].[dbo].tbl_10x10_board"
		sql = sql + " where idx=" + v

		rsget.Open sql, dbget, 1

		if  not rsget.EOF  then
			Fidx          = rsget("idx")
			Fgubun     = rsget("gubun")
			Fuserid   = rsget("userid")
			Fusername   = rsget("username")
			Ftitle   = db2html(rsget("title"))
			Fcontents   = nl2br(db2html(rsget("contents")))
			Fregdate      = rsget("regdate")
		end if
		rsget.close
	end sub

	Public Function modify(byval idx,gubun,userid,username,title,contents)
                dim sql, i

                sql = "update [db_board].[dbo].tbl_10x10_board " + VbCrlf
                sql = sql + " set gubun = '" + gubun + "'," + VbCrlf
				sql = sql + " userid = '" + userid + "'," + VbCrlf
				sql = sql + " username = '" + username + "'," + VbCrlf
				sql = sql + " title = '" + title + "'," + VbCrlf
                sql = sql + " contents = '" + contents + "'" + VbCrlf
                sql = sql + " where idx = " + idx
                rsget.Open sql, dbget, 1
	end Function

	Public Function write(byval gubun,userid,username,title,contents)
                dim sql, i

                sql = " insert into [db_board].[dbo].tbl_10x10_board(gubun,userid,username,title,contents) "
                sql = sql + " values('" + gubun + "','" + userid + "', '" + username + "', '" + title + "','" + contents + "') "
                rsget.Open sql, dbget, 1
	end Function

	Public Function delete(byval idx)
                dim sql, i

                sql = "update [db_board].[dbo].tbl_10x10_board " + VbCrlf
                sql = sql + " set isusing = 'N'" + VbCrlf
                sql = sql + " where idx = " + idx
                rsget.Open sql, dbget, 1
	end Function

end Class

%>