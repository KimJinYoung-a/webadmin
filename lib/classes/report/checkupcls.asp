<%
class CReportItem
	public FItemId
	public FItemName
	public FItemOptionName
	public FCount
	public FPriceSum
	public FDesignerID
	public FRealSell

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CCheckUp
	public FItemList()

	public FCurrPage
    public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount

    public FRectDesignerID
    public FRectRegStartPre
	public FRectRegStart
	public FRectRegEnd

	public Sub getMinusItemList()
		dim sqlStr,i

		''########## ÃÑ °¹¼ö ################''
		sqlStr = "select count(d.idx) as cnt"
		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.regdate>='" + FRectRegStart + "'"
		sqlStr = sqlStr + " and m.regdate<'" + FRectRegEnd + "'"
		sqlStr = sqlStr + " and m.jumundiv='9'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.itemcost<0"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		if FRectDesignerID<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		end if

		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " d.itemid, d.itemoption, d.itemname, d.itemoptionname,d.makerid,"
		sqlStr = sqlStr + " sum(d.itemno) as totcnt, sum(d.itemcost) as totsum, IsNull(T.cnt,0) as sellcnt"

		sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, "
		sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " left join  ("
		sqlStr = sqlStr + " 	select itemid,itemoption,sum(d.itemno) as cnt"
		sqlStr = sqlStr + " 	from [db_order].[dbo].tbl_order_master m, "
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d"
		sqlStr = sqlStr + " 	where m.regdate>='" + FRectRegStartPre + "' "
		sqlStr = sqlStr + " 	and m.regdate<'" + FRectRegEnd + "' "
		sqlStr = sqlStr + " 	and m.orderserial=d.orderserial "
		sqlStr = sqlStr + " 	and d.itemid<>0"
		sqlStr = sqlStr + " 	and d.itemcost*d.itemno>0 "
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
		if FRectDesignerID<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		end if
		sqlStr = sqlStr + " group by d.itemid, d.itemoption"
		sqlStr = sqlStr + " ) as T on T.itemid=d.itemid and T.itemoption=d.itemoption"

		sqlStr = sqlStr + " where m.orderserial=d.orderserial "
		sqlStr = sqlStr + " and m.regdate>='" + FRectRegStart + "' "
		sqlStr = sqlStr + " and m.regdate<'" + FRectRegEnd + "' "
		sqlStr = sqlStr + " and m.jumundiv='9' "
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.itemcost*d.itemno<0 "
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		if FRectDesignerID<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		end if
		sqlStr = sqlStr + " group by d.itemid, d.itemoption,d.itemname, d.itemoptionname,d.makerid,IsNull(T.cnt,0)"
		sqlStr = sqlStr + " order by totcnt ,totsum desc"

		'sqlStr = " select top " + Cstr(FPageSize * FCurrPage) + " d.itemid, d.itemoption,"
		'sqlStr = sqlStr + " d.itemname, d.itemoptionname,d.makerid,"
		'sqlStr = sqlStr + " sum(d.itemno) as totcnt,  sum(d.itemcost) as totsum"
		'sqlStr = sqlStr + " from [10x10].[dbo].tbl_order_master m,"
		'sqlStr = sqlStr + " [10x10].[dbo].tbl_order_detail d"
		'sqlStr = sqlStr + " where m.jumundiv='9'"
		'sqlStr = sqlStr + " and m.regdate>='" + FRectRegStart + "'"
		'sqlStr = sqlStr + " and m.regdate<'" + FRectRegEnd + "'"
		'sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		'sqlStr = sqlStr + " and d.itemid<>0"
		'sqlStr = sqlStr + " and d.itemcost<0"
		'sqlStr = sqlStr + " and m.cancelyn='N'"
		'sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		'if FRectDesignerID<>"" then
		'	sqlStr = sqlStr + " and d.makerid='" + FRectDesignerID + "'"
		'end if
		'sqlStr = sqlStr + " group by d.itemid, d.itemoption,d.itemname, d.itemoptionname,d.makerid"
		'sqlStr = sqlStr + " order by totcnt desc,totsum"

		'response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1
		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CReportItem
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemName       = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName = db2html(rsget("itemoptionname"))
				FItemList(i).FCount          = rsget("totcnt")
				FItemList(i).FPriceSum       = rsget("totsum")
				FItemList(i).FDesignerID     = rsget("makerid")
				FItemList(i).FRealSell		 = rsget("sellcnt")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

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