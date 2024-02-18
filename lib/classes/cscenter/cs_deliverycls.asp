<%

'// idx, orderserial, songjangdiv, songjangno, isupchebeasong, checkCnt, beasongdate, findDate, realDeliveryDate, regdate, lastupdate

Class CCSDeliveryItem
	public Fidx
	public Forderserial
	public Fsongjangdiv
	public FsongjangName
	public Fsongjangno
	public Ffindurl
	public Fisupchebeasong
	public FcheckCnt
	public Fbeasongdate
	public FfindDate
	public FrealDeliveryDate
	public Fregdate
	public Flastupdate
	public Fmakerid
	public Fipkumdate
	public FDPlusNDay
	public Fdelaydiv
	public Fitemid

	public function GetDelayDivName()
		select case Fdelaydiv
			case "delay"
				GetDelayDivName = "배송지연"
			case "stockout"
				GetDelayDivName = "품절"
			case else
				GetDelayDivName = Fdelaydiv
		end select
	end function

    Private Sub Class_Initialize()
		'//
    End Sub
    Private Sub Class_Terminate()
		'//
    End Sub
end class

Class CCSDelivery
    public FItemList()
    public FOneItem
    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

	public FRectStartDate
	public FRectEndDate
	public FRectSongjangDiv
	public FRectOrderSerial
	public FRectDelayDelivOnly
	public FRectMakerid
	public FRectCheckCnt
	public FRectDelayDiv

	public Sub GetCSMemoDeliveryDelaySUM()
		dim i, sqlStr, sqlWhere

		sqlWhere = ""
		sqlWhere = sqlWhere + " 	and DPlusNDay >= '" & FRectStartDate & "' "
		sqlWhere = sqlWhere + " 	and DPlusNDay < '" & FRectEndDate & "' "
		if (FRectDelayDiv <> "") then
			sqlWhere = sqlWhere + " 	and delaydiv = '" & FRectDelayDiv & "' "
		end if

		sqlStr = sqlStr + " select delaydiv, makerid, count(*) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_DeliveryDelayList] "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1=1 "
		sqlStr = sqlStr + sqlWhere
		sqlStr = sqlStr + " group by makerid, delaydiv "
		sqlStr = sqlStr + " order by cnt desc, delaydiv, makerid "
        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCSDeliveryItem

				FItemList(i).Fdelaydiv			= rsget("delaydiv")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).FcheckCnt			= rsget("cnt")

				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.close()
	end sub

	public Sub GetCSMemoDeliveryDelayByMakerid()
		dim i, sqlStr, sqlWhere

		sqlWhere = ""
		sqlWhere = ""
		sqlWhere = sqlWhere + " 	and DPlusNDay >= '" & FRectStartDate & "' "
		sqlWhere = sqlWhere + " 	and DPlusNDay < '" & FRectEndDate & "' "
		sqlWhere = sqlWhere + " 	and makerid = '" & FRectMakerid & "' "
		if (FRectDelayDiv <> "") then
			sqlWhere = sqlWhere + " 	and delaydiv = '" & FRectDelayDiv & "' "
		end if

		'// ====================================================================
		sqlStr = " select count(l.orderserial) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_DeliveryDelayList] l "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + sqlWhere
		''response.write sqlStr

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close


		'// ====================================================================
		sqlStr = " select top " & CStr(FPageSize*FCurrPage) & " l.* "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_DeliveryDelayList] l "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + sqlWhere
		sqlStr = sqlStr + " order by DPlusNDay desc, delaydiv, makerid, orderserial, maxitemid "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCSDeliveryItem

				FItemList(i).Fdelaydiv			= rsget("delaydiv")
				FItemList(i).Forderserial		= rsget("orderserial")
				FItemList(i).Fitemid			= rsget("maxitemid")
				FItemList(i).FDPlusNDay			= rsget("DPlusNDay")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).FcheckCnt			= 1

				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.close()
	end sub

	public Sub GetCSMemoDeliveryFixSUM()
		dim i, sqlStr, sqlWhere

		sqlWhere = ""
		sqlWhere = sqlWhere + " 	and l.beasongdate >= '" & FRectStartDate & "' "
		sqlWhere = sqlWhere + " 	and l.beasongdate < '" & FRectEndDate & "' "
		if (FRectMakerid <> "") then
			sqlWhere = sqlWhere + " 	and l.makerid = '" & FRectMakerid & "' "
		end if
		sqlWhere = sqlWhere + " 	and (((l.realDeliveryDate is NULL) and (checkCnt = 5)) "
		sqlWhere = sqlWhere + " 	or "
		sqlWhere = sqlWhere + " 	(l.realDeliveryDate < Convert(varchar(10), DateAdd(day, 0, l.ipkumdate), 121)) "
		sqlWhere = sqlWhere + " 	or "
		sqlWhere = sqlWhere + " 	(l.realDeliveryDate >= DPlusNDay)) "

		sqlStr = " select l.makerid, count(*) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_DeliveryTrackingList] l "
		sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_songjang_div] dv on l.songjangdiv = dv.divcd "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere
		sqlStr = sqlStr + " group by l.makerid "
		sqlStr = sqlStr + " order by cnt desc "
        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCSDeliveryItem

				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).FcheckCnt			= rsget("cnt")

				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.close()
	end sub

	public Sub GetCSMemoDeliveryFixByMakerid()
		dim i, sqlStr, sqlWhere

		sqlWhere = ""
		sqlWhere = sqlWhere + " 	and l.beasongdate >= '" & FRectStartDate & "' "
		sqlWhere = sqlWhere + " 	and l.beasongdate < '" & FRectEndDate & "' "
		sqlWhere = sqlWhere + " 	and l.makerid = '" & FRectMakerid & "' "
		sqlWhere = sqlWhere + " 	and (((l.realDeliveryDate is NULL) and (checkCnt = 5)) "
		sqlWhere = sqlWhere + " 	or "
		sqlWhere = sqlWhere + " 	(l.realDeliveryDate < Convert(varchar(10), DateAdd(day, 0, l.ipkumdate), 121)) "
		sqlWhere = sqlWhere + " 	or "
		sqlWhere = sqlWhere + " 	(l.realDeliveryDate >= DPlusNDay)) "

		'// ====================================================================
		sqlStr = " select count(l.idx) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_DeliveryTrackingList] l "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + sqlWhere
		''response.write sqlStr

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close


		'// ====================================================================
		sqlStr = " select top " & CStr(FPageSize*FCurrPage) & " l.*, dv.divname, dv.findurl "
		sqlStr = sqlStr + " , l.makerid "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_DeliveryTrackingList] l "
		sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_songjang_div] dv on l.songjangdiv = dv.divcd "
		''sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_order_master] m on m.orderserial = l.orderserial "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + sqlWhere
		sqlStr = sqlStr + " order by l.beasongdate desc, l.idx desc "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr, dbget, 1

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCSDeliveryItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Forderserial		= rsget("orderserial")
				FItemList(i).Fsongjangdiv		= rsget("songjangdiv")
				FItemList(i).FsongjangName		= db2html(rsget("divname"))
				FItemList(i).Fsongjangno		= rsget("songjangno")
				FItemList(i).Ffindurl			= db2html(rsget("findurl"))
				FItemList(i).Fisupchebeasong	= rsget("isupchebeasong")
				FItemList(i).FcheckCnt			= rsget("checkCnt")
				FItemList(i).Fbeasongdate		= rsget("beasongdate")
				FItemList(i).FfindDate			= rsget("findDate")
				FItemList(i).FrealDeliveryDate	= rsget("realDeliveryDate")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fipkumdate			= rsget("ipkumdate")
				FItemList(i).FDPlusNDay			= rsget("DPlusNDay")

				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.close()
	end sub

	public Sub GetCSMemoDeliverySUM()
		dim i, sqlStr
		sqlStr = " select l.songjangdiv, dv.divname, count(*) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_datamart].[dbo].[tbl_DeliveryTrackingList] l "
		sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_songjang_div] dv on l.songjangdiv = dv.divcd "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + " 	and l.beasongdate >= '" & FRectStartDate & "' "
		sqlStr = sqlStr + " 	and l.beasongdate < '" & FRectEndDate & "' "
		''sqlStr = sqlStr + " 	and l.checkCnt = 5 "
		sqlStr = sqlStr + " 	and (((l.realDeliveryDate is NULL) and (checkCnt = 5)) "
		sqlStr = sqlStr + " 	or "
		sqlStr = sqlStr + " 	(l.realDeliveryDate < Convert(varchar(10), DateAdd(day, 0, l.ipkumdate), 121)) "
		sqlStr = sqlStr + " 	or "
		sqlStr = sqlStr + " 	(l.realDeliveryDate >= Convert(varchar(10), DateAdd(day, 3, l.beasongdate), 121))) "
		sqlStr = sqlStr + " group by l.songjangdiv, dv.divname "
		sqlStr = sqlStr + " order by dv.divname, l.songjangdiv "
        db3_rsget.pagesize = FPageSize
        db3_rsget.Open sqlStr, db3_dbget, 1

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CCSDeliveryItem

				FItemList(i).Fsongjangdiv		= db3_rsget("songjangdiv")
				FItemList(i).FsongjangName		= db2html(db3_rsget("divname"))
				FItemList(i).FcheckCnt			= db3_rsget("cnt")

				db3_rsget.moveNext
				i=i+1
			loop
		end if
		db3_rsget.close()
	end sub

    public Sub GetCSMemoDeliveryList()
        dim i, sqlStr, sqlWhere

		sqlWhere = ""
		sqlWhere = sqlWhere + " 	and l.beasongdate >= '" & FRectStartDate & "' "
		sqlWhere = sqlWhere + " 	and l.beasongdate < '" & FRectEndDate & "' "
		''sqlWhere = sqlWhere + " 	and m.cancelyn = 'N' "
		if (FRectDelayDelivOnly <> "") then
			sqlWhere = sqlWhere + " 	and (((l.realDeliveryDate is NULL) and (checkCnt = 5)) "
			sqlWhere = sqlWhere + " 	or "
			sqlWhere = sqlWhere + " 	(l.realDeliveryDate < Convert(varchar(10), DateAdd(day, 0, l.ipkumdate), 121)) "
			sqlWhere = sqlWhere + " 	or "
			sqlWhere = sqlWhere + " 	(l.realDeliveryDate >= Convert(varchar(10), DateAdd(day, 3, l.beasongdate), 121))) "
		end if
		if (FRectMakerid <> "") then
			sqlWhere = sqlWhere + " 	and l.makerid = '" & FRectMakerid & "' "
		end if
		if (FRectSongjangDiv <> "") then
			sqlWhere = sqlWhere + " 	and l.songjangdiv = " & FRectSongjangDiv
		end if
		if (FRectOrderSerial <> "") then
			sqlWhere = sqlWhere + " 	and l.orderserial = '" & FRectOrderSerial & "' "
		end if
		if (FRectCheckCnt <> "") then
			sqlWhere = sqlWhere + " 	and l.checkCnt >= " & FRectCheckCnt
		end if


		'// ====================================================================
		sqlStr = " select count(l.idx) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_datamart].[dbo].[tbl_DeliveryTrackingList] l "
		''sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_order_master] m on m.orderserial = l.orderserial "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + sqlWhere
		''response.write sqlStr

		db3_rsget.Open sqlStr,db3_dbget,1
			FTotalCount = db3_rsget("cnt")
		db3_rsget.Close


		'// ====================================================================
		sqlStr = " select top " & CStr(FPageSize*FCurrPage) & " l.*, dv.divname, dv.findurl "
		sqlStr = sqlStr + " , l.makerid "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_datamart].[dbo].[tbl_DeliveryTrackingList] l "
		sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_songjang_div] dv on l.songjangdiv = dv.divcd "
		''sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_order_master] m on m.orderserial = l.orderserial "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + sqlWhere
		sqlStr = sqlStr + " order by l.beasongdate desc, l.idx desc "

        db3_rsget.pagesize = FPageSize
        db3_rsget.Open sqlStr, db3_dbget, 1

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CCSDeliveryItem

				FItemList(i).Fidx				= db3_rsget("idx")
				FItemList(i).Forderserial		= db3_rsget("orderserial")
				FItemList(i).Fsongjangdiv		= db3_rsget("songjangdiv")
				FItemList(i).FsongjangName		= db2html(db3_rsget("divname"))
				FItemList(i).Fsongjangno		= db3_rsget("songjangno")
				FItemList(i).Ffindurl			= db2html(db3_rsget("findurl"))
				FItemList(i).Fisupchebeasong	= db3_rsget("isupchebeasong")
				FItemList(i).FcheckCnt			= db3_rsget("checkCnt")
				FItemList(i).Fbeasongdate		= db3_rsget("beasongdate")
				FItemList(i).FfindDate			= db3_rsget("findDate")
				FItemList(i).FrealDeliveryDate	= db3_rsget("realDeliveryDate")
				FItemList(i).Fregdate			= db3_rsget("regdate")
				FItemList(i).Flastupdate		= db3_rsget("lastupdate")
				FItemList(i).Fmakerid			= db3_rsget("makerid")
				FItemList(i).Fipkumdate			= db3_rsget("ipkumdate")

				db3_rsget.moveNext
				i=i+1
			loop
		end if
		db3_rsget.close()

	end sub

    Private Sub Class_Initialize()
        FCurrPage       = 1
        FPageSize       = 10
        FResultCount    = 0
        FScrollCount    = 10
        FTotalCount     = 0
    End Sub

    Private Sub Class_Terminate()
		'//
    End Sub

    public Function HasPreScroll()
        HasPreScroll = StartScrollPage > 1
    end Function

    public Function HasNextScroll()
        HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
    end Function

    public Function StartScrollPage()
        StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
    end Function
end class

%>
