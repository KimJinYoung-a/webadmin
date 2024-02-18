<%

'id, divcd, subcd, title, contents, regdate, hitcount, isusing
Class CBoardFAQItem
	private Fid
	private Fdivcd
	private Fsubcd
	private Ftitle
	private Fcontents
	private Fregdate
	private Fhitcount
	private Fisusing

        '==========================================================================
	Property Get id()
		id = Fid
	end Property

	Property Get divcd()
		divcd = Fdivcd
	end Property

	Property Get subcd()
		subcd = Fsubcd
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

	Property Get hitcount()
		hitcount = Fhitcount
	end Property

	Property Get isusing()
		isusing = Fisusing
	end Property

        '==========================================================================
	Property Let id(byVal v)
		Fid = v
	end Property

	Property Let divcd(byVal v)
		Fdivcd = v
	end Property

	Property Let subcd(byVal v)
		Fsubcd = v
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

	Property Let hitcount(byVal v)
		Fhitcount = v
	end Property

	Property Let isusing(byVal v)
		Fisusing = v
	end Property

        '==========================================================================
	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
        '
	End Sub
end Class

Class CBoardFAQ
        public results()

	private FCurrPage
	private FTotalPage
	private FTotalCount
	private FPageSize
	private FResultCount
	private FScrollCount

	private FIDBefore
	private FIDAfter

	private FSearchString
	private FSearchDivCode
	private FSearchSort

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


	Property Let SearchString(byVal v)
		FSearchString = v
	end Property

	Property Let SearchDivCode(byVal v)
		FSearchDivCode = v
	end Property

	Property Let SearchSort(byVal v)
		FSearchSort = v
	end Property


	Private Sub Class_Initialize()
			redim results(0)
		FSearchString = ""
		FSearchDivCode = ""
		FSearchSort = ""
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub


        '======================================================================
	Public Function list()
                dim sql, i

                sql = " select top " + CStr(FPageSize*FCurrPage) + " id, divcd, subcd, title, regdate, hitcount, isusing from [db_cs].[10x10].tbl_faq "
                sql = sql + " where (1 = 1) "

                if (FSearchString <> "") then
                        sql = sql + " and (title like '%" + FSearchString + "%') "
                end if

                if (FSearchDivCode <> "") then
                        sql = sql + " and (divcd = '" + FSearchDivCode + "') "
                end if

                if (FSearchSort = "regdate") then
                        sql = sql + " order by regdate desc "
                elseif (FSearchSort = "count") then
                        sql = sql + " order by hitcount desc, regdate desc "
                else
                        sql = sql + " order by isusing desc, regdate desc "
                end if

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
                                set results(i) = new CBoardFAQItem

                                results(i).id = rsget("id")
                                results(i).divcd = rsget("divcd")
                                results(i).subcd = rsget("subcd")
                                results(i).title = rsget("title")
                                'results(i).contents = rsget("contents")
                                results(i).regdate = rsget("regdate")
                                results(i).hitcount = rsget("hitcount")
                                results(i).isusing = rsget("isusing")

				rsget.MoveNext
				i = i + 1
                        loop
                end if
                rsget.close
	end Function

	Public Function listWithContents()
                dim sql, i

                sql = " select id, divcd, subcd, title, contents, regdate, hitcount, isusing from [db_cs].[10x10].tbl_faq "
                sql = sql + " where (isusing = 'Y') "

                if (FSearchString <> "") then
                        sql = sql + " and (title like '%" + FSearchString + "%') "
                end if

                if (FSearchDivCode <> "") then
                        sql = sql + " and (divcd = '" + FSearchDivCode + "') "
                end if

                if (FSearchSort = "regdate") then
                        sql = sql + " order by regdate desc "
                elseif (FSearchSort = "count") then
                        sql = sql + " order by hitcount desc, regdate desc "
                else
                        sql = sql + " order by regdate desc "
                end if

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
                                set results(i) = new CBoardFAQItem

                                results(i).id = rsget("id")
                                results(i).divcd = rsget("divcd")
                                results(i).subcd = rsget("subcd")
                                results(i).title = rsget("title")
                                results(i).contents = rsget("contents")
                                results(i).regdate = rsget("regdate")
                                results(i).hitcount = rsget("hitcount")

				rsget.MoveNext
				i = i + 1
                        loop
                end if
                rsget.close
	end Function

	Public Function read(byval id)
                dim sql, i

                sql = " select id, divcd, subcd, title, contents, regdate, hitcount, isusing from [db_cs].[10x10].tbl_faq "
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
                'response.write sql

                redim preserve results(rsget.RecordCount)

                if not rsget.EOF then
                        set results(0) = new CBoardFAQItem

                        results(i).id = rsget("id")
                        results(i).divcd = rsget("divcd")
                        results(i).subcd = rsget("subcd")
                        results(i).title = rsget("title")
                        results(i).contents = rsget("contents")
                        results(i).regdate = rsget("regdate")
                        results(i).hitcount = rsget("hitcount")
                end if
                rsget.close
	end Function

	Public Function modify(byval boarditem)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_faq " + VbCrlf
                sql = sql + " set divcd = '" + boarditem.divcd + "'," + VbCrlf
                sql = sql + " title = '" + boarditem.title + "'," + VbCrlf
                sql = sql + " contents = '" + boarditem.contents + "' " + VbCrlf
                sql = sql + " where (id = " + boarditem.id + ") "
                'response.write sql
                'dbget.close()	:	response.End
                rsget.Open sql, dbget, 1
	end Function

	Public Function addcount(byval id)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_faq set hitcount = hitcount+1 " + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function write(byval boarditem)
                dim sql, i

                sql = " insert into [db_cs].[10x10].tbl_faq(divcd, subcd, title, contents, regdate, hitcount, isusing) "
                sql = sql + " values('" + boarditem.divcd + "', '00', '" + boarditem.title + "', '" + boarditem.contents + "', getdate(), 0, 'Y') "
                rsget.Open sql, dbget, 1
	end Function

	Public Function delete(byval id)
                dim sql, i

                sql = "update [db_cs].[10x10].tbl_faq set isusing = 'N' " + VbCrlf
                sql = sql + " where (id = " + id + ") "
                rsget.Open sql, dbget, 1
	end Function

	Public Function code2name(byval v)
                if (v = "01") then
                        code2name = "회원정보관련"
                elseif (v = "02") then
                        code2name = "상품문의"
                elseif (v = "03") then
                        code2name = "주문/결재"
                elseif (v = "04") then
                        code2name = "취소/반품"
                elseif (v = "05") then
                        code2name = "기타"
                else
                        code2name = ""
                end if
	end Function
end Class
%>

    