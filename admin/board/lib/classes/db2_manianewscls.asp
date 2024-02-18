<%

'id, title, contents, regdate, yuhyostart, yuhyoend, isusing
Class CBoardNoticeItem
	private Fid
	private Ftitle
	private Fcontents
	private Fregdate
	private Fyuhyostart
	private Fyuhyoend
	private Fisusing
	private Fmalltype
	private Fnoticetype
	private Ffixyn

        '==========================================================================
	Property Get id()
		id = Fid
	end Property

	Property Get title()
		title = Ftitle
	end Property

	Property Get contents()
		contents = Fcontents
	end Property

	Property Get regdate()
		regdate = Fregdate
	end Property

	Property Get isusing()
		isusing = Fisusing
	end Property

'==========================================================================
	Property Let id(byVal v)
		Fid = v
	end Property

	Property Let title(byVal v)
		Ftitle = v
	end Property

	Property Let contents(byVal v)
		Fcontents = v
	end Property

	Property Let regdate(byVal v)
		Fregdate = v
	end Property

	Property Let isusing(byVal v)
		Fisusing = v
	end Property

'==========================================================================
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CBoardNotice
        public results()

	private FCurrPage
	private FTotalPage
	private FTotalCount
	private FPageSize
	private FResultCount
	private FScrollCount
	private FIDBefore
	private FIDAfter

	Property Get CurrPage()
		CurrPage = FCurrPage
	end Property

	Property Get TotalPage()
		TotalPage = FTotalPage
	end Property

	Property Get TotalCount()
		TotalCount = FTotalCount
	end Property

	Property Get PageSize()
		PageSize = FPageSize
	end Property

	Property Get ResultCount()
		ResultCount = FResultCount
	end Property

	Property Get ScrollCount()
		ScrollCount = FScrollCount
	end Property

	Property Get IDBefore()
		IDBefore = FIDBefore
	end Property

	Property Get IDAfter()
		IDAfter = FIDAfter
	end Property


	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = TotalPage > StartScrollPage + ScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((Currpage-1)\ScrollCount)*ScrollCount +1
	end Function


	Property Let CurrPage(byVal v)
		FCurrPage = v
	end Property

	Property Let PageSize(byVal v)
		FPageSize = v
	end Property

	Property Let ScrollCount(byVal v)
		FScrollCount = v
	end Property

	Private Sub Class_Initialize()
				redim results(0)
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub


        '======================================================================
	Public Function list()
        dim sql, i
		sql = " select count(idx) as cnt from  [db_cts].[dbo].tbl_mania_news "

		db2_rsget.Open sql, db2_dbget, 1
		FTotalCount = db2_rsget("cnt")
		db2_rsget.close

        sql = " select top " + CStr(FPageSize*FCurrPage) + " idx, title, isusing, regdate from [db_cts].[dbo].tbl_mania_news"
        sql = sql + " order by idx desc "

		db2_rsget.pagesize = FPageSize
		db2_rsget.Open sql, db2_dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db2_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if (FResultCount > PageSize) then
			FResultCount = PageSize
		end if

                redim preserve results(FResultCount)

                if not db2_rsget.EOF then
                        i = 0
                        db2_rsget.absolutepage = FCurrPage
                        do until ( db2_rsget.eof or (i > FResultCount))
                                set results(i) = new CBoardNoticeItem

                                results(i).id = db2_rsget("idx")
										  results(i).title = db2html(db2_rsget("title"))
                                results(i).isusing = db2_rsget("isusing")
                                results(i).regdate = db2_rsget("regdate")

				db2_rsget.MoveNext
				i = i + 1
                        loop
                end if
                db2_rsget.close
	end Function

	Public Function read(byval idx)
                dim sql, i

                sql = " select idx, title, contents, isusing, regdate from [db_cts].[dbo].tbl_mania_news "
                sql = sql + " where (idx = " + idx + ") "
                db2_rsget.Open sql, db2_dbget, 1
                'response.write sql

                redim preserve results(db2_rsget.RecordCount)

                if not db2_rsget.EOF then
                        set results(0) = new CBoardNoticeItem

                        results(0).id = db2_rsget("idx")
                        results(0).title = db2html(db2_rsget("title"))
                        results(0).contents = db2html(db2_rsget("contents"))
                        results(0).regdate = db2_rsget("regdate")
                        results(0).isusing = db2_rsget("isusing")

                end if
                db2_rsget.close
	end Function

	Public Function modify(byval boarditem)
                dim sql, i

                sql = "update [db_cts].[dbo].tbl_mania_news " + VbCrlf
					 sql = sql + " set title = '" + boarditem.title + "'," + VbCrlf
                sql = sql + " contents = '" + boarditem.contents + "'," + VbCrlf
                sql = sql + " isusing = '" + boarditem.isusing + "'" + VbCrlf
                sql = sql + " where (idx = " + boarditem.id + ") "
                db2_rsget.Open sql, db2_dbget, 1
	end Function

	Public Function write(byval boarditem)
                dim sql, i

                sql = " insert into [db_cts].[dbo].tbl_mania_news(title, contents, isusing) "
                sql = sql + " values('" + boarditem.title + "', '" + boarditem.contents + "', '" + boarditem.isusing + "') "
                db2_rsget.Open sql, db2_dbget, 1
	end Function

end Class
%>

    