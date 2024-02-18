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

        '==========================================================================
	Property Get id()
		id = Fid
	end Property

	Property Get malltype()
		malltype = Fmalltype
	end Property

	Property Get noticetype()
		noticetype = Fnoticetype
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

	Property Get yuhyostart()
		yuhyostart = Fyuhyostart
	end Property

	Property Get yuhyoend()
		yuhyoend = Fyuhyoend
	end Property

	Property Get isusing()
		isusing = Fisusing
	end Property

        '==========================================================================
	Property Let id(byVal v)
		Fid = v
	end Property

	Property Let malltype(byVal v)
		Fmalltype = v
	end Property

	Property Let noticetype(byVal v)
		Fnoticetype = v
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

	Property Let yuhyostart(byVal v)
		Fyuhyostart = v
	end Property

	Property Let yuhyoend(byVal v)
		Fyuhyoend = v
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

                sql = " select top " + CStr(FPageSize*FCurrPage) + " id, title, regdate, yuhyostart, yuhyoend, isusing from [db_shop].[dbo].tbl_offshop_notice "
                sql = sql + " where isusing = 'Y' "
                sql = sql + " order by regdate desc "

		if FPageSize<>0 then
			rsget.pagesize = PageSize
		end if
                rsget.Open sql, dbget, 1
                'response.write sql

		FTotalCount = rsget.RecordCount
		FTotalPage = rsget.PageCount
		FResultCount = rsget.RecordCount - (CurrPage-1)*PageSize

		if (FResultCount > PageSize) then
			FResultCount = PageSize
		end if

                redim preserve results(FResultCount)

                if not rsget.EOF then
                        i = 0
                        rsget.absolutepage = FCurrPage
                        do until ( rsget.eof or (i > FResultCount))
                                set results(i) = new CBoardNoticeItem

                                results(i).id = rsget("id")
								results(i).title = rsget("title")
                                'results(i).contents = rsget("contents")
                                results(i).regdate = rsget("regdate")
                                results(i).yuhyostart = rsget("yuhyostart")
                                results(i).yuhyoend = rsget("yuhyoend")

				rsget.MoveNext
				i = i + 1
                        loop
                end if
                rsget.close
	end Function

	Public Function read(byval id)
                dim sql, i

                sql = " select id, malltype, noticetype, title, contents, regdate, yuhyostart, yuhyoend, isusing from [db_shop].[dbo].tbl_offshop_notice "
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
                'response.write sql

                redim preserve results(rsget.RecordCount)

                if not rsget.EOF then
                        set results(0) = new CBoardNoticeItem

                        results(0).id = rsget("id")
						results(0).malltype = rsget("malltype")
						results(0).noticetype = rsget("noticetype")
                        results(0).title = rsget("title")
                        results(0).contents = rsget("contents")
                        results(0).regdate = rsget("regdate")
                        results(0).yuhyostart = rsget("yuhyostart")
                        results(0).yuhyoend = rsget("yuhyoend")
                end if
                rsget.close
	end Function

	Public Function modify(byval boarditem)
                dim sql, i

                sql = "update [db_shop].[dbo].tbl_offshop_notice " + VbCrlf
                sql = sql + " set malltype = '" + boarditem.malltype + "'," + VbCrlf
				sql = sql + " noticetype = '" + boarditem.noticetype + "'," + VbCrlf
				sql = sql + " title = '" + boarditem.title + "'," + VbCrlf
                sql = sql + " contents = '" + boarditem.contents + "'," + VbCrlf
                sql = sql + " yuhyostart = '" + boarditem.yuhyostart + "'," + VbCrlf
                sql = sql + " yuhyoend = '" + boarditem.yuhyoend + "' " + VbCrlf
                sql = sql + " where (id = " + boarditem.id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function write(byval boarditem)
                dim sql, i

                sql = " insert into [db_shop].[dbo].tbl_offshop_notice(malltype, noticetype, title, contents, regdate, yuhyostart, yuhyoend, isusing) "
                sql = sql + " values('" + boarditem.malltype + "', '" + boarditem.noticetype + "', '" + boarditem.title + "', '" + boarditem.contents + "', getdate(), '" + boarditem.yuhyostart + "', '" + boarditem.yuhyoend + "', 'Y') "
                rsget.Open sql, dbget, 1
	end Function

	Public Function delete(byval id)
                dim sql, i

                sql = "update [db_shop].[dbo].tbl_offshop_notice set isusing = 'N' " + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function
end Class
%>

    