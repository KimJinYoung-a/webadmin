<%
Class CBrandSellReportItem
	public Fuserid
	public Fuserdiv
	public Fmaeipdiv
	public Fdefaultmargine
	public Fsocname_kor
	public Fisusing
	public Fmduserid
	public Fregdate
	public Fitemcount

	public Fsellttl
	public Fbuyttl

	public function GetUserDivName
		if Fuserdiv="02" then
			GetUserDivName = "디자인업체"
		elseif Fuserdiv="03" then
			GetUserDivName = "플라워업체"
		elseif Fuserdiv="04" then
			GetUserDivName = "패션업체"
		elseif Fuserdiv="05" then
			GetUserDivName = "쥬얼리업체"
		elseif Fuserdiv="06" then
			GetUserDivName = "케어업체"
		elseif Fuserdiv="07" then
			GetUserDivName = "애견업체"
		elseif Fuserdiv="08" then
			GetUserDivName = "보드게임"
		elseif Fuserdiv="13" then
			GetUserDivName = "여행몰업체"
		elseif Fuserdiv="14" then
			GetUserDivName = "강사"
		elseif Fuserdiv="20" then
			GetUserDivName = "텐바이텐소호"
		else
			GetUserDivName = Fuserdiv
		end if
	end function

	public function GetMaeipDivName
		if Fmaeipdiv="M" then
			GetMaeipDivName = "매입"
		elseif Fmaeipdiv="W" then
			GetMaeipDivName = "위탁"
		elseif Fmaeipdiv="U" then
			GetMaeipDivName = "업체"
		else

		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CNewReport
	public FItemList()
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
    public FCurrPage

	public FRectFromDate
	public FRectToDate
	public FRectSearchType

	public Sub GetNewBrandSellReport
		dim sqlStr, i
		sqlStr = "select top " + CStr(FPageSize) + " c.userid,c.userdiv,c.maeipdiv,c.defaultmargine,c.socname_kor,c.isusing,c.mduserid,c.regdate,c.itemcount"
		sqlStr = sqlStr + " ,IsNULL(T.sellttl,0) as sellttl, IsNULL(T.buyttl,0) as buyttl"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select d.makerid,	sum(d.itemcost*d.itemno) as sellttl, sum(d.buycash*d.itemno) as buyttl"
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " 	 [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " 	where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " 	and m.regdate>='" + Cstr(FRectFromDate) + "'"
		sqlStr = sqlStr + " 	and m.regdate<'" + Cstr(FRectToDate) + "'"
		sqlStr = sqlStr + " 	and m.ipkumdiv>3"
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.itemid<>0"
		sqlStr = sqlStr + " 	group by d.makerid"
		sqlStr = sqlStr + " ) T on c.userid=T.makerid"
		sqlStr = sqlStr + " where c.userdiv<21"
		if FRectSearchType="N" then
			sqlStr = sqlStr + " and datediff(d,c.regdate,getdate())<31"
		end if

		if FRectSearchType="N" then
			sqlStr = sqlStr + " order by T.sellttl/(datediff(d,c.regdate,getdate())+1) desc"
		else
			sqlStr = sqlStr + " order by T.sellttl  desc"
		end if

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)

		do until rsget.eof
				set FItemList(i) = new CBrandSellReportItem
			    FItemList(i).Fuserid        = rsget("userid")
				FItemList(i).Fuserdiv       = rsget("userdiv")
				FItemList(i).Fmaeipdiv      = rsget("maeipdiv")
				FItemList(i).Fdefaultmargine= rsget("defaultmargine")
				FItemList(i).Fsocname_kor   = db2html(rsget("socname_kor"))
				FItemList(i).Fisusing       = rsget("isusing")
				FItemList(i).Fmduserid      = rsget("mduserid")
				FItemList(i).Fregdate       = rsget("regdate")
				FItemList(i).Fitemcount		= rsget("itemcount")

				FItemList(i).Fsellttl       = rsget("sellttl")
				FItemList(i).Fbuyttl        = rsget("buyttl")


				rsget.MoveNext
				i = i + 1
		loop
		rsget.close

	end Sub

	Private Sub Class_Initialize()
		redim FItemList(0)

		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub


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

%>