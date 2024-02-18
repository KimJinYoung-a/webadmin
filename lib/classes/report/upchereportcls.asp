<%
Class CUpcheReportItem
	public FDateGubun
	public FSubCount
	public FSubItemCount
	public FSubSellTotal
	public FSubBuyTotal

	public FBestItem1
	public FBestItem2
	public FBestItem3

	public FEventNo
	public FTotalCount
	public FTotalSum

	public FRank
	public FTotalUpCheCount

	public FMWTotal
	public FManCount
	public FWoManCount
	public FAvgManyear
	public FAvgWoManyear

	public function getManTargetNai()
		if CLng(FAvgManyear)>40 then
			getManTargetNai = CLng(year(now)) - CLng("19" + Cstr(FAvgManyear))
		end if
	end function

	public function getWoManTargetNai()
		if CLng(FAvgWoManyear)>40 then
			getWoManTargetNai = CLng(year(now)) - CLng("19" + Cstr(FAvgWoManyear))
		end if
	end function

	Private Sub Class_Initialize()
		FRank = 1
		FMWTotal     =0
		FManCount    =0
		FWoManCount  =0
		FAvgManyear  =0
		FAvgWoManyear=0

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CUpcheReport
	public FItemList()

	public FCurrPage
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount

	public FRectDesigner

	public FRectFromDate
	public FRectToDate

	public FRectMonth

	public Sub GetUpcheSheet3()
		dim i,sqlStr
		dim cnt, bufsno, bufsex, bufyear, bufdate
		cnt = UBound(FItemList)-1

		sqlStr = "select sum(d.itemno) as sno, Right(Left(n.juminno,8),1) as gsex,"
		sqlStr = sqlStr + " Avg(convert(int,Left(n.juminno,2))) as gyear,"
		sqlStr = sqlStr + " convert(varchar(7),m.regdate,20) as dategubun"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d,"
		sqlStr = sqlStr + " [db_user].[dbo].tbl_user_n n"
		sqlStr = sqlStr + " where m.regdate>='" + FRectFromDate + "'"
		sqlStr = sqlStr + " and m.regdate<'" + FRectToDate + "'"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.sitename='10x10'"
		sqlStr = sqlStr + " and m.userid<>''"
		sqlStr = sqlStr + " and m.userid=n.userid"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " group by convert(varchar(7),m.regdate,20), Right(Left(n.juminno,8),1)"

		rsget.Open sqlStr,dbget,1

		do until rsget.eof
			bufsno   = rsget("sno")
			bufsex = rsget("gsex")
			bufyear = rsget("gyear")
			bufdate   = rsget("dategubun")

			for i=0 to cnt
				if (FItemList(i).FDateGubun=bufdate) then
					FItemList(i).FMWTotal = FItemList(i).FMWTotal +  bufsno

					if bufsex="1" then
						FItemList(i).FManCount =  bufsno
						FItemList(i).FAvgManyear  = bufyear
					else
						FItemList(i).FWoManCount =  bufsno
						FItemList(i).FAvgWoManyear = bufyear
					end if
					Exit for
				end if
			next

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close

	end sub

	public Sub GetUpcheSheet2()
		dim i,sqlStr
		dim cnt, bufcnt, buftotsum, bufdate, bufmakerid
		cnt = UBound(FItemList)-1

		sqlStr = "select count(m.idx) as cnt, sum(d.itemcost*d.itemno) as totsum,"
		sqlStr = sqlStr + " convert(varchar(7),m.regdate,20) as dategubun, d.makerid"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.regdate>='" + FRectFromDate + "'"
		sqlStr = sqlStr + " and m.regdate<'" + FRectToDate + "'"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " group by convert(varchar(7),m.regdate,20), makerid"

		rsget.Open sqlStr,dbget,1

		do until rsget.eof
			bufcnt   = rsget("cnt")
			buftotsum = rsget("totsum")
			bufdate   = rsget("dategubun")
			bufmakerid = rsget("makerid")

			for i=0 to cnt
				if (FItemList(i).FDateGubun=bufdate) then
					FItemList(i).FTotalUpCheCount = FItemList(i).FTotalUpCheCount +1
					FItemList(i).FTotalCount = FItemList(i).FTotalCount + bufcnt
					FItemList(i).FTotalSum = FItemList(i).FTotalSum + buftotsum

					if buftotsum>FItemList(i).FSubSellTotal then
							FItemList(i).FRank = FItemList(i).FRank + 1
					end if

					Exit for
				end if
			next

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end Sub

	public Sub GetUpcheSheet1()
		dim i,sqlStr

		sqlStr = "select count(m.idx) as cnt, sum(d.itemno) as itemcnt, sum(d.itemcost*d.itemno) as sellcash,"
		sqlStr = sqlStr + " sum(d.buycash*d.itemno) as buycashcash, convert(varchar(7),m.regdate,20) as dategubun"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.regdate>='" + FRectFromDate + "'"
		sqlStr = sqlStr + " and m.regdate<'" + FRectToDate + "'"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and m.ipkumdiv>=4"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " group by convert(varchar(7),m.regdate,20)"
		sqlStr = sqlStr + " order by dategubun"
'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new CUpcheReportItem
			FItemList(i).FDateGubun = rsget("dategubun")
			FItemList(i).FSubCount  = rsget("cnt")
			FItemList(i).FSubItemCount = rsget("itemcnt")
			FItemList(i).FSubSellTotal  = rsget("sellcash")
			FItemList(i).FSubBuyTotal  = rsget("buycashcash")

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end Sub

	public Sub GetUpcheAllMeaChul()
		dim i,sqlStr
		dim cnt
		dim bufno,bufsum,bufdate

		cnt = UBound(FItemList)-1

		sqlStr = " select count(m.idx) as cnt, sum(d.itemcost*d.itemno) as totsum,"
		sqlStr = sqlStr + " convert(varchar(7),m.regdate,20) as dategubun"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.regdate>='" + FRectFromDate + "'"
		sqlStr = sqlStr + " and m.regdate<'" + FRectToDate + "'"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " group by convert(varchar(7),m.regdate,20)"

		rsget.Open sqlStr,dbget,1

		do until rsget.eof

			bufno = rsget("cnt")
			bufsum = rsget("totsum")
			bufdate = rsget("dategubun")

			for i=0 to cnt
				if (FItemList(i).FDateGubun=bufdate) then
					FItemList(i).FTotalCount = bufno
					FItemList(i).FTotalSum = bufsum
					Exit for
				end if
			next

			rsget.MoveNext

		loop
		rsget.close
	end sub

	public Sub GetUpcheBestItemByMonth(byval ix)
		dim i,sqlStr
		sqlStr = "select top 3 sum(d.itemno) as itemcnt , d.itemname"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where convert(varchar(7),m.regdate,20)='" + FRectMonth + "'"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " group by itemname"
		sqlStr = sqlStr + " order by itemcnt desc"

		rsget.Open sqlStr,dbget,1
		do until rsget.eof
			if FItemList(ix).FBestItem1 = "" then
				FItemList(ix).FBestItem1 = rsget("itemname") + "(" +  CStr(rsget("itemcnt")) + ")"
			elseif FItemList(ix).FBestItem2 = "" then
				FItemList(ix).FBestItem2 = rsget("itemname") + "(" +  CStr(rsget("itemcnt")) + ")"
			elseif FItemList(ix).FBestItem3 = "" then
				FItemList(ix).FBestItem3 = rsget("itemname") + "(" +  CStr(rsget("itemcnt")) + ")"
			end if
			rsget.MoveNext
		loop
		rsget.close
	end Sub

	public Sub GetUpcheBestItem()
		dim i,sqlStr
		dim bufno, bufdate, bufitemname, cnt

		cnt = UBound(FItemList)-1

		sqlStr = "select top 200 sum(d.itemno) as itemcnt , d.itemname,"
		sqlStr = sqlStr + " convert(varchar(7),m.regdate,20) as dategubun"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where  m.regdate>='" + FRectFromDate + "'"
		sqlStr = sqlStr + " and m.regdate<'" + FRectToDate + "'"
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " group by convert(varchar(7),m.regdate,20),itemname"
		sqlStr = sqlStr + " order by  itemcnt desc, dategubun"

		rsget.Open sqlStr,dbget,1

		do until rsget.eof

			bufno = rsget("itemcnt")
			bufdate = rsget("dategubun")
			bufitemname = rsget("itemname")

			for i=0 to cnt
				if (FItemList(i).FDateGubun=bufdate) then
					if FItemList(i).FBestItem1 = "" then
						FItemList(i).FBestItem1 = bufitemname + "(" +  CStr(bufno) + ")"
					elseif FItemList(i).FBestItem2 = "" then
						FItemList(i).FBestItem2 = bufitemname + "(" +  CStr(bufno) + ")"
					elseif FItemList(i).FBestItem3 = "" then
						FItemList(i).FBestItem3 = bufitemname + "(" +  CStr(bufno) + ")"
					end if
					exit for
				end if
			next

			rsget.MoveNext

		loop
		rsget.close
	end Sub

	public Sub GetEventItem()
		dim i,sqlStr, cnt
		dim bufno, bufdate1, bufdate2

		cnt = UBound(FItemList)-1

		sqlStr = "select count(idx) as cnt,convert(varchar(7),startday,21) as sdt ,"
		sqlStr = sqlStr + " convert(varchar(7),endday,21) as edt"
		sqlStr = sqlStr + " from [db_contents].[dbo].tbl_event_master"
		sqlStr = sqlStr + " where designerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and ((startday>'" + FRectFromDate + "'"
		sqlStr = sqlStr + " and startday<'" + FRectToDate + "'"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " or (endday>'" + FRectFromDate + "'"
		sqlStr = sqlStr + " and endday<'" + FRectToDate + "'))"
		sqlStr = sqlStr + " group by convert(varchar(7),startday,21), convert(varchar(7),endday,21)"

		rsget.Open sqlStr,dbget,1

		do until rsget.eof

			bufno = rsget("cnt")
			bufdate1 = rsget("sdt")
			bufdate2 = rsget("edt")


			for i=0 to cnt
				if (FItemList(i).FDateGubun=bufdate1) or (FItemList(i).FDateGubun=bufdate2) then
					FItemList(i).FEventNo = bufno
					Exit for
				end if
			next

			rsget.MoveNext

		loop
		rsget.close
	end Sub

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)
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