<%
Class COffShopDishCorrectItem
	public Fimgsmall
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fonlinesellcash
	public Fofflinesellcash
	public Fipchulsellcash
	public Fipchulno

	public Fipchuldiffsellprice
	public Fipchuldiffsellno

	public Fipchulsamesellprice
	public Fipchulsamesellno

	public Fstockcurrno

	public function GetMayBojungCount()
		GetMayBojungCount = Fipchulno-Fipchulsamesellno
	end function

	public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (FItemID >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,FItemId)) + CStr(Fitemoption)
    	end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COffShopDishCorrect
	public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectIdx
	public FRectMakerid
	public FRectShopID

	public FRecAvailbojung

	public sub GetDishValidList()
		dim sqlStr, i

		sqlStr = "select A.iitemgubun,A.itemid,A.itemoption,s.shopitemname,s.shopitemoptionname,i.sellcash as onlinesellcash,"
		sqlStr = sqlStr + " s.shopitemprice as offlinesellcash,A.sellcash as ipchulsellcash,A.ipchulno,"
		sqlStr = sqlStr + " IsNULL(B.sellprice,0) as ipchuldiffsellprice,IsNULL(B.sellno,0) as ipchuldiffsellno,"
		sqlStr = sqlStr + " IsNULL(C.sellprice,0) as ipchulsamesellprice,IsNULL(C.sellno,0) as ipchulsamesellno,"
		sqlStr = sqlStr + " t.currno,i.smallimage,s.offimgsmall"
		'sqlStr = sqlStr + " t.ipno, t.reno, t.upcheipno, t.upchereno, t.sellno, t.currno"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + " 	select d.iitemgubun,d.itemid,d.itemoption, d.sellcash, sum(d.itemno*-1) as ipchulno"
		sqlStr = sqlStr + " 	from [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " 	where m.code=d.mastercode"
		sqlStr = sqlStr + " 	and m.deldt is null"
		sqlStr = sqlStr + " 	and d.deldt is null"
		sqlStr = sqlStr + " 	and m.socid='" + FRectShopID + "'"
		sqlStr = sqlStr + " 	and d.imakerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " 	and d.itemno<>0"
		sqlStr = sqlStr + " 	group by d.iitemgubun,d.itemid,d.itemoption, d.sellcash"
		sqlStr = sqlStr + " ) A"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on A.iitemgubun='10' and a.itemid=i.itemid"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s on A.iitemgubun=s.itemgubun and A.itemid=s.shopitemid and A.itemoption=s.itemoption"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_day_stock t on t.shopid='" + FRectShopID + "' and A.iitemgubun=t.itemgubun and A.itemid=t.itemid and A.itemoption=t.itemoption"
		sqlStr = sqlStr + " left join "
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + " 	select T.itemgubun,T.itemid,T.itemoption, T.sellprice, sum(T.itemno) as sellno"
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	("
		sqlStr = sqlStr + " 	select d.idx, d.itemgubun,d.itemid,d.itemoption, d.sellprice, d.itemno"
		sqlStr = sqlStr + " 	from [db_shoplog].[dbo].tbl_old_shopjumun_master m,"
		sqlStr = sqlStr + " 	[db_shoplog].[dbo].tbl_old_shopjumun_detail d"
		sqlStr = sqlStr + " 	where m.idx=d.masteridx"
		sqlStr = sqlStr + " 	and m.shopid='" + FRectShopID + "' "
		sqlStr = sqlStr + " 	and d.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn='N'"
		sqlStr = sqlStr + " union "
		sqlStr = sqlStr + " 	select d.idx, d.itemgubun,d.itemid,d.itemoption, d.sellprice, d.itemno"
		sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " 	where m.idx=d.masteridx"
		sqlStr = sqlStr + " 	and m.shopid='" + FRectShopID + "' "
		sqlStr = sqlStr + " 	and d.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn='N'"
		sqlStr = sqlStr + " 	) T"
		sqlStr = sqlStr + " 	group by T.itemgubun,T.itemid,T.itemoption, T.sellprice"
		sqlStr = sqlStr + " ) B on A.iitemgubun=B.itemgubun and A.itemid=B.itemid and A.itemoption=B.itemoption and A.sellcash<>IsNULL(B.sellprice,0)"
		sqlStr = sqlStr + " left join "
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + " 	select T.itemgubun,T.itemid,T.itemoption, T.sellprice, sum(T.itemno) as sellno"
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	( "
		sqlStr = sqlStr + " 	select d.idx, d.itemgubun,d.itemid,d.itemoption, d.sellprice, d.itemno"
		sqlStr = sqlStr + " 	from [db_shoplog].[dbo].tbl_old_shopjumun_master m,"
		sqlStr = sqlStr + " 	[db_shoplog].[dbo].tbl_old_shopjumun_detail d"
		sqlStr = sqlStr + " 	where m.idx=d.masteridx"
		sqlStr = sqlStr + " 	and m.shopid='" + FRectShopID + "' "
		sqlStr = sqlStr + " 	and d.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn='N'"
		sqlStr = sqlStr + " union "
		sqlStr = sqlStr + " 	select d.idx, d.itemgubun,d.itemid,d.itemoption, d.sellprice, d.itemno"
		sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " 	where m.idx=d.masteridx"
		sqlStr = sqlStr + " 	and m.shopid='" + FRectShopID + "' "
		sqlStr = sqlStr + " 	and d.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn='N'"
		sqlStr = sqlStr + " 	) T"
		sqlStr = sqlStr + " 	group by T.itemgubun,T.itemid,T.itemoption, T.sellprice"
		sqlStr = sqlStr + " ) C on A.iitemgubun=C.itemgubun and A.itemid=C.itemid and A.itemoption=C.itemoption and  A.sellcash=IsNULL(C.sellprice,0)"
		sqlStr = sqlStr + " where A.ipchulno<>0"
		if FRecAvailbojung="on" then
		sqlStr = sqlStr + " and A.sellcash<>i.sellcash or A.sellcash<>shopitemprice"
		end if
		sqlStr = sqlStr + " order by A.iitemgubun,A.itemid,A.itemoption"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDishCorrectItem
				FItemList(i).Fitemgubun             = rsget("iitemgubun")
				FItemList(i).Fitemid                 = rsget("itemid")
				FItemList(i).Fitemoption             = rsget("itemoption")
				FItemList(i).Fitemname               = db2html(rsget("shopitemname"))
				FItemList(i).Fitemoptionname         = db2html(rsget("shopitemoptionname"))
				FItemList(i).Fonlinesellcash         = rsget("onlinesellcash")
				FItemList(i).Fofflinesellcash        = rsget("offlinesellcash")
				FItemList(i).Fipchulsellcash         = rsget("ipchulsellcash")
				FItemList(i).Fipchulno               = rsget("ipchulno")

				FItemList(i).Fipchuldiffsellprice    = rsget("ipchuldiffsellprice")
				FItemList(i).Fipchuldiffsellno       = rsget("ipchuldiffsellno")

				FItemList(i).Fipchulsamesellprice    = rsget("ipchulsamesellprice")
				FItemList(i).Fipchulsamesellno       = rsget("ipchulsamesellno")

				FItemList(i).Fstockcurrno            = rsget("currno")

				if FItemList(i).Fitemgubun="10" then
					FItemList(i).Fimgsmall        = rsget("smallimage")
				else
					FItemList(i).Fimgsmall        = rsget("offimgsmall")
				end if

				if IsNULL(FItemList(i).Fimgsmall) then

				elseif	FItemList(i).Fitemgubun="10" then
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
				else
					FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
				end if

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
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



