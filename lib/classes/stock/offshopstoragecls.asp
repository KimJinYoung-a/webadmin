<%
Class COffShopStorageItem
	public Fmakerid
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fsellcash
	public FLastrealno
	public Fipno
	public Freno
	public Fsellno

	public function GetBargode()
		GetBargode = Fitemgubun + Format00(6,Fitemid) + Fitemoption
		if (FItemID >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + Format00(8,FItemId) + FItemOption
    	end if
	end function

	public function GetMayno()
		GetMayno = FLastrealno + Fipno + Freno - Fsellno
	end function

	Private Sub Class_Initialize()
		Fsellcash   = 0
		FLastrealno = 0
		Fipno       = 0
		Freno       = 0
		Fsellno     = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COffShopStorage
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectShopid
	public FRectStartDate
	public FRectEndDate
	public FRectMakerid

	public Sub getStorageNSellList()
		dim sqlStr
		sqlStr = " select top 5000 d.makerid, d.itemgubun,d.itemid,d.itemoption,"
		sqlStr = sqlStr + " d.itemname,d.itemoptionname,"
		sqlStr = sqlStr + " sum("
		sqlStr = sqlStr + " case when d.realitemno>=0 then d.realitemno"
		sqlStr = sqlStr + " else 0"
		sqlStr = sqlStr + " end"
		sqlStr = sqlStr + " ) as ipno"
		sqlStr = sqlStr + " ,"
		sqlStr = sqlStr + " sum("
		sqlStr = sqlStr + " case when d.realitemno<0 then d.realitemno"
		sqlStr = sqlStr + " else 0"
		sqlStr = sqlStr + " end"
		sqlStr = sqlStr + " ) as reno"
		sqlStr = sqlStr + " , IsNULL(S.sellno,0) as sellno"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " select d.itemgubun,d.itemid,d.itemoption,sum(d.itemno) as sellno"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
		end if
		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		end if
		sqlStr = sqlStr + " and m.shopregdate>'" + FRectStartDate + "'"
		sqlStr = sqlStr + " and m.shopregdate<'" + FRectEndDate + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn='N'"
		sqlStr = sqlStr + " group by d.itemgubun,d.itemid,d.itemoption"
		sqlStr = sqlStr + " ) S on d.itemgubun=s.itemgubun"
		sqlStr = sqlStr + " and d.itemid=s.itemid"
		sqlStr = sqlStr + " and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.baljuid='" + FRectShopid + "'"
		end if
		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		end if
		sqlStr = sqlStr + " and m.ipgodate>'" + FRectStartDate + "'"
		sqlStr = sqlStr + " and m.ipgodate<'" + FRectEndDate + "'"
		sqlStr = sqlStr + " and m.statecd='7'"
		sqlStr = sqlStr + " and m.divcode in ('501','502','503')"
		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " group by d.makerid, d.itemgubun,d.itemid,d.itemoption,d.itemname,d.itemoptionname,sellno"
		sqlStr = sqlStr + " order by d.makerid, d.itemgubun,d.itemid,d.itemoption"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopStorageItem
					FItemList(i).Fmakerid 	= rsget("makerid")
					FItemList(i).Fitemgubun      = rsget("itemgubun")
					FItemList(i).Fitemid         = rsget("itemid")
					FItemList(i).Fitemoption     = rsget("itemoption")
					FItemList(i).Fitemname       = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
					'FItemList(i).Fsellcash       = rsget("sellcash")
					''FItemList(i).FLastrealno     = rsget("Lastrealno")
					FItemList(i).Fipno           = rsget("ipno")
					FItemList(i).Freno           = rsget("reno")
					FItemList(i).Fsellno         = rsget("sellno")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class
%>