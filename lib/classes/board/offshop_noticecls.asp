<%

class CNoticeSubItem

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

	function GubunName()
		if Fgubun = "00" then
			GubunName = "전체"
		elseif Fgubun = "01" then
			GubunName = "zoom"
		elseif Fgubun = "02" then
			GubunName = "college"
		elseif Fgubun = "03" then
			GubunName = "cafe"
		elseif Fgubun = "50" then
			GubunName = "매장전체"
		elseif Fgubun = "51" then
			GubunName = "직영매장"
		elseif Fgubun = "52" then
			GubunName = "수수료매장"
		elseif Fgubun = "53" then
			GubunName = "프랜챠이즈"
		end if
	end function

end Class

Class CNotice

	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FrectMallType
	public FRectCDL
	public FRectGubun

	public FRectListAll

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

        public Sub offshopnoticelist()
		dim sql, i

		sql = "select count(idx) as cnt "
		sql = sql + " from [db_board].[dbo].tbl_offshop_notice"
		sql = sql + " where isusing = 'Y' "
		if FRectGubun <> "" then
		sql = sql + " and gubun in ('50','" + FRectGubun + "')"
		end if

		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

		sql = " select top " + CStr(FPageSize*FCurrPage) + " idx,userid,username, title, regdate"
		sql = sql + " from [db_board].[dbo].tbl_offshop_notice"
		sql = sql + " where isusing = 'Y' "
		if FRectGubun <> "" then
		sql = sql + " and gubun in ('50','" + FRectGubun + "')"
		end if
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
				set FItemList(i) = new CNoticeSubItem

				FItemList(i).Fidx          = rsget("idx")
				FItemList(i).Fuserid   = rsget("userid")
				FItemList(i).Fusername   = rsget("username")
				FItemList(i).Ftitle   =  db2html(rsget("title"))
				FItemList(i).Fregdate      = rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub


	public Sub list()
		dim sql, i

		sql = "select count(idx) as cnt "
		sql = sql + " from [db_board].[dbo].tbl_offshop_notice"
		sql = sql + " where isusing = 'Y' "
		if FRectListAll<>"" then

		else
			sql = sql + " and gubun < 50"
		end if

		if FRectGubun <> "" then
		sql = sql + " and gubun in ('00','" + FRectGubun + "')"
		end if

		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

		sql = " select top " + CStr(FPageSize*FCurrPage) + " idx,userid,username, title, regdate"
		sql = sql + " from [db_board].[dbo].tbl_offshop_notice"
		sql = sql + " where isusing = 'Y' "
		if FRectListAll<>"" then

		else
			sql = sql + " and gubun < 50"
		end if

		if FRectGubun <> "" then
		sql = sql + " and gubun in ('00','" + FRectGubun + "')"
		end if
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
				set FItemList(i) = new CNoticeSubItem

				FItemList(i).Fidx          = rsget("idx")
				FItemList(i).Fuserid   = rsget("userid")
				FItemList(i).Fusername   = rsget("username")
				FItemList(i).Ftitle   =  db2html(rsget("title"))
				FItemList(i).Fregdate      = rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class


class CNoticeDetail

	public Fidx
	public Fgubun
	public Fuserid
	public Fusername
	public Ftitle
	public Fcontents
	public Ffile
	public Ffilelink
	public Fregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	function GubunName()
		if Fgubun = "00" then
			GubunName = "전체"
		elseif Fgubun = "01" then
			GubunName = "zoom"
		elseif Fgubun = "02" then
			GubunName = "college"
		elseif Fgubun = "03" then
			GubunName = "cafe"
		elseif Fgubun = "50" then
			GubunName = "매장전체"
		elseif Fgubun = "51" then
			GubunName = "직영매장"
		elseif Fgubun = "52" then
			GubunName = "수수료매장"
		elseif Fgubun = "53" then
			GubunName = "프랜챠이즈"
		end if
	end function

	public Sub read(byVal v)
		dim sql, i

		sql = " select top 1 idx,gubun,userid,username,title,contents,[file],regdate"
		sql = sql + " from [db_board].[dbo].tbl_offshop_notice"
		sql = sql + " where isusing = 'Y' "
		sql = sql + " and  idx=" + Cstr(v)

		rsget.Open sql, dbget, 1

		if  not rsget.EOF  then
			Fidx          = rsget("idx")
			Fgubun          = rsget("gubun")
			Fuserid   = rsget("userid")
			Fusername   = rsget("username")
			Ftitle   = db2html(rsget("title"))
			Fcontents   = db2html(rsget("contents"))
			Ffile   = rsget("file")
			Fregdate      = rsget("regdate")
			Ffilelink   = "http://webadmin.10x10.co.kr/admin/board/noticefile/" + rsget("file")
		end if
		rsget.close
	end sub

end Class

%>