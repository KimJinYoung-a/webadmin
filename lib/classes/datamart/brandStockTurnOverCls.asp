<%
Class COutItemItem
    public Fshopid
    public Fmakerid
    public Fyyyymm
    public Fitemgubun
    public Fitemid
    public FitemOption
    public FItemName
    public FItemOptionName

    public Fstockno
    public FtotSellNo
    public FtotRealSellPrice

    public FOffimgSmall
    public FimageSmall

    public function IsImageExists()
        dim buf : buf = GetImageSmall
        IsImageExists = false
        if IsNULL(buf) then Exit function
        if Right(buf,1)="/" then Exit function

        IsImageExists = true
    end function

    public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = FimageSmall
		else
			GetImageSmall = FOffImgSmall
		end if
	end function

    public function getTenBarCode()
        getTenBarCode = Fitemgubun + Format00(6,Fitemid) + FitemOption
        if (Fitemid >= 1000000) then
    		getTenBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

    End Sub
End Class

Class COutItemBrandItem
    public Fshopid
    public Fmakerid
    public Fyyyymm
    public Fcomm_cd
    public Fcomm_name
    public FItemCnt
    public FitemTaragetCnt
    public FstockTaragetCnt
    public FtotSellNo
    public FtotRealSellPrice

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

    End Sub
End Class

Class CBrandStockTurnOverMasterItem
    public Fshopid
    public Fmakerid
    public Fyyyymm
    public Fcomm_cd
    public Fcomm_name
    public Fstockno
    public FStShopItemPrice
    public FStShopBuyPrice
    public FpreStockNo
    public FPreStShopItemPrice
    public FPreStShopBuyPrice
    public FrealCheckErrNo
    public FrealCheckErrShopItemPrice
    public FrealCheckErrShopBuyPrice
    public FtotSellno
    public FtotRealSellPrice
    public FtotTenBuyPrice
    public FtotShopBuyPrice
    public FStTurnOverBySell
    public FStTurnOver
    public Fregdate

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

    End Sub
End Class

Class CBrandStockTurnOver
    public FItemList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectYYYYMM
	public FRectJungsanGubun
	public FRectShopid
	public FRectMakerid
	public FRectComm_cd

	public FRectItemGubun
	public FRectItemID
	public FRectItemOption
	public FRectBarcode
	public FRectStartDate
	public FRectSearchMode

	''---월별 정리 대상 상품.------------------------------------------------------------------------------
	public Sub getOutItemList
	    Dim sqlStr, i, sqlADD

	    sqlADD = ""
	    if (FRectYYYYMM<>"") then
            sqlADD = sqlADD & " and I.yyyymm='"&FRectYYYYMM&"'"
        end if

        if (FRectShopid<>"") then
            sqlADD = sqlADD & " and I.shopid='"&FRectShopid&"'"
        end if

        if (FRectMakerid<>"") then
            sqlADD = sqlADD & " and I.makerid='"&FRectMakerid&"'"
        end if

        if (FRectSearchMode="X") then
            sqlADD = sqlADD & " and IsNULL(I.stockno,0)>1 and IsNULL(I.totSellno,0)<1"
        elseif (FRectSearchMode="S") then
            sqlADD = sqlADD & " and IsNULL(I.stockno,0)>1"
        else

        end if

	    sqlStr = " select count(*) as CNT"
	    sqlStr = sqlStr & " from db_datamart.dbo.tbl_Shop_ItemTurnOver I"
	    sqlStr = sqlStr & " where 1=1"
	    sqlStr = sqlStr & sqlADD

	    db3_rsget.Open sqlStr,db3_dbget,1
            FTotalCount = db3_rsget("cnt")
        db3_rsget.Close

        sqlStr = " select top "&(FPageSize*FCurrPage)
        sqlStr = sqlStr & " I.shopid, I.makerid, I.yyyymm, I.ItemGubun, I.ItemID, I.Itemoption"
        sqlStr = sqlStr & " , I.stockno"
        sqlStr = sqlStr & " , (I.totSellNo) as totSellNo "
        sqlStr = sqlStr & " , (I.totRealSellPrice) as totRealSellPrice "
        sqlStr = sqlStr & " , S.Shopitemname, S.Shopitemoptionname, S.offimgsmall, o.smallimage"
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_Shop_ItemTurnOver I"
        sqlStr = sqlStr & "     left join db_datamart.dbo.tbl_DataMart_shop_item s"
        sqlStr = sqlStr & "     on I.ItemGubun=s.itemgubun"
        sqlStr = sqlStr & "     and I.ItemId=s.ShopItemID"
        sqlStr = sqlStr & "     and I.ItemOption=s.Itemoption"
        sqlStr = sqlStr & "     left join db_datamart.dbo.tbl_item o"
        sqlStr = sqlStr & "     on I.ItemGubun='10'"
        sqlStr = sqlStr & "     and I.ItemId=o.itemid"
        sqlStr = sqlStr & " where 1=1"
	    sqlStr = sqlStr & sqlADD
''rw sqlStr

        db3_rsget.pagesize = FPageSize
        db3_rsget.Open sqlStr,db3_dbget,1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new COutItemItem
				FItemList(i).Fshopid                    = db3_rsget("shopid")
                FItemList(i).Fmakerid                   = db3_rsget("makerid")
                FItemList(i).Fyyyymm                    = db3_rsget("yyyymm")

                FItemList(i).Fitemgubun                 = db3_rsget("itemgubun")
                FItemList(i).Fitemid                    = db3_rsget("itemid")
                FItemList(i).FitemOption                = db3_rsget("itemOption")

                FItemList(i).Fstockno                   = db3_rsget("stockno")
                FItemList(i).Fitemname                  = db3_rsget("Shopitemname")
                FItemList(i).Fitemoptionname            = db3_rsget("Shopitemoptionname")
                FItemList(i).FtotSellNo                 = db3_rsget("totSellNo")
                FItemList(i).FtotRealSellPrice          = db3_rsget("totRealSellPrice")

                FItemList(i).FOffimgSmall	= db3_rsget("offimgsmall")
                if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall

    			FItemList(i).FimageSmall     = db3_rsget("smallimage")
    			if FItemList(i).FimageSmall<>"" then
    				FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
    			end if

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    End Sub

	public Sub getOutItemBrandList
	    Dim sqlStr, i
        sqlStr = " select I.shopid, I.makerid, I.yyyymm, B.comm_cd, J.comm_name"
        sqlStr = sqlStr & " , count(*) as ItemCnt"
        sqlStr = sqlStr & " , Sum(CASE WHEN IsNULL(I.stockno,0)>1 and IsNULL(I.totSellno,0)<1 THEN 1 ELSE 0 END) as itemTaragetCnt"
        sqlStr = sqlStr & " , Sum(CASE WHEN IsNULL(I.stockno,0)>1 and IsNULL(I.totSellno,0)<1 THEN I.stockno ELSE 0 END) as stockTaragetCnt"
        sqlStr = sqlStr & " , Sum(I.totSellNo) as totSellNo "
        sqlStr = sqlStr & " , Sum(I.totRealSellPrice) as totRealSellPrice "
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_Shop_ItemTurnOver I"
        sqlStr = sqlStr & "     left join db_datamart.dbo.tbl_Shop_BrandTurnOver B"
        sqlStr = sqlStr & "     on B.yyyymm=I.yyyymm"
        sqlStr = sqlStr & "     and B.shopid=I.shopid"
        sqlStr = sqlStr & "     and B.makerid=I.makerid"
        sqlStr = sqlStr & "     left join db_datamart.dbo.tbl_jungsan_comm_code J"
        sqlStr = sqlStr & "     on B.comm_cd=j.comm_cd"

        sqlStr = sqlStr & " where 1=1"
        if (FRectYYYYMM<>"") then
            sqlStr = sqlStr & " and I.yyyymm='"&FRectYYYYMM&"'"
        end if

        if (FRectShopid<>"") then
            sqlStr = sqlStr & " and I.shopid='"&FRectShopid&"'"
        end if

        if (FRectMakerid<>"") then
            sqlStr = sqlStr & " and I.makerid='"&FRectMakerid&"'"
        end if

        if (FRectComm_cd<>"") then
            if (FRectComm_cd="B099") then
                sqlStr = sqlStr + " and B.comm_cd in ('B031','B011')"
            elseif (FRectComm_cd="B088") then
                sqlStr = sqlStr + " and B.comm_cd in ('B012','B022')"
            else
                sqlStr = sqlStr + " and B.comm_cd='" + FRectComm_cd + "'"
            end if
        end if
        sqlStr = sqlStr & " group by I.Shopid, I.makerid, I.yyyymm, B.comm_cd, J.comm_name"
        sqlStr = sqlStr & " Having Sum(CASE WHEN IsNULL(I.stockno,0)>1 and IsNULL(I.totSellno,0)<1 THEN 1 ELSE 0 END)>0"
        sqlStr = sqlStr & " order by stockTaragetCnt desc"
''rw sqlStr
        db3_rsget.Open sqlStr,db3_dbget,1
        FResultCount = db3_rsget.RecordCount
        FTotalCount  = FResultCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new COutItemBrandItem
				FItemList(i).Fshopid                    = db3_rsget("shopid")
                FItemList(i).Fmakerid                   = db3_rsget("makerid")
                FItemList(i).Fyyyymm                    = db3_rsget("yyyymm")
                FItemList(i).Fcomm_cd                   = db3_rsget("comm_cd")
                FItemList(i).Fcomm_name                 = db3_rsget("comm_name")
                FItemList(i).FItemCnt                   = db3_rsget("ItemCnt")
                FItemList(i).FitemTaragetCnt            = db3_rsget("itemTaragetCnt")
                FItemList(i).FstockTaragetCnt           = db3_rsget("stockTaragetCnt")
                FItemList(i).FtotSellNo                 = db3_rsget("totSellNo")
                FItemList(i).FtotRealSellPrice          = db3_rsget("totRealSellPrice")


				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    End Sub

	''--과 재고 검색 (판매<0, 재고>0)
	public Sub getStockExistsAndNoSellBrandList_Crent_REAL()
	    Dim sqlStr, i
	    sqlStr = " select"
        sqlStr = sqlStr + " i.makerid"
        sqlStr = sqlStr + " ,sum(CASE WHEN s.realstockno>0 and IsNULL(T.SellCnt,0)<1 THEN 1 ELSE 0 END) as noSellItemCNT"
        sqlStr = sqlStr + " ,sum(CASE WHEN s.realstockno>0 and IsNULL(T.SellCnt,0)<1 THEN s.realstockno ELSE 0 END) as StockNO"
        sqlStr = sqlStr + " ,sum(CASE WHEN s.realstockno>0 THEN s.realstockno ELSE 0 END) as Totrealstockno"
        sqlStr = sqlStr + " ,sum(IsNULL(SellCnt,0)) as totSellCnt"
        sqlStr = sqlStr + " ,sum(IsNULL(TTLSell,0)) as totSellsum"
        sqlStr = sqlStr + " from db_summary.dbo.tbl_current_shopStock_summary s"
        sqlStr = sqlStr + " 	Join db_shop.dbo.tbl_shop_item i"
        sqlStr = sqlStr + " 	on s.itemgubun=i.itemgubun"
        sqlStr = sqlStr + " 	and s.itemid=i.shopitemid"
        sqlStr = sqlStr + " 	and s.itemoption=i.itemoption"
        sqlStr = sqlStr + " 	left join ("
        sqlStr = sqlStr + " 		select itemgubun,itemid,itemoption,sum(itemno) as SellCnt, sum(itemno*realsellprice) as TTLSell"
        sqlStr = sqlStr + " 			from db_shop.dbo.tbl_shopjumun_master m"
        sqlStr = sqlStr + " 			Join db_shop.dbo.tbl_shopjumun_detail d"
        sqlStr = sqlStr + " 			on m.orderno=d.orderno"
        sqlStr = sqlStr + " 		where m.shopid='"&FRectShopID&"'"
        sqlStr = sqlStr + " 		and m.IXyyyymmdd>='"&FRectStartDate&"'"
        '''sqlStr = sqlStr + " 		and m.IXyyyymmdd<'"&FRectEndDate&"'"
        sqlStr = sqlStr + " 		and m.cancelyn='N'"
        sqlStr = sqlStr + " 		and d.cancelyn='N'"
        sqlStr = sqlStr + " 		group by itemgubun, itemid,itemoption"
        sqlStr = sqlStr + " 	) T"
        sqlStr = sqlStr + " 	on s.itemgubun=T.itemgubun"
        sqlStr = sqlStr + " 	and s.itemid=T.itemid"
        sqlStr = sqlStr + " 	and s.itemoption=T.itemoption"
        sqlStr = sqlStr + " where s.shopid='"&FRectShopID&"'"
        sqlStr = sqlStr + " group by i.makerid"
        sqlStr = sqlStr + " Having sum(CASE WHEN s.realstockno>0 and IsNULL(T.SellCnt,0)<1 THEN 1 ELSE 0 END) >0"
        sqlStr = sqlStr + " order by StockNO desc"

    end Sub

	public Sub getItemStockTurnOverList
	    Dim sqlStr

        sqlStr = "SELECT T.*e, "
        sqlStr = sqlStr + " (CASE WHEN IsNULL(T.StShopBuyprice,0)+IsNULL(T.PreStShopBuyprice,0)=0 THEN -1"
        sqlStr = sqlStr + " ELSE IsNULL(T.totShopBuyprice,0)/((IsNULL(T.StShopBuyprice,0)+IsNULL(T.PreStShopBuyprice,0))/2)"
        sqlStr = sqlStr + " END) as StTurnOver"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + " 		select "
        sqlStr = sqlStr + " 		A.itemgubun,A.itemid,A.itemoption "
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.stockno,0)) as stockno"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.StShopitemprice,0)) as StShopitemprice"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.StShopBuyprice,0)) as StShopBuyprice"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.StTenBuyprice,0)) as StTenBuyprice"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.prestockno,0)) as prestockno"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.PreStShopitemprice,0)) as PreStShopitemprice"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.PreStShopBuyprice,0)) as PreStShopBuyprice"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.PreStTenBuyprice,0)) as PreStTenBuyprice"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.realCheckErrNo,0)) as realCheckErrNo"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.realCheckErrShopItemPrice,0)) as realCheckErrShopItemPrice"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.realCheckErrShopBuyPrice,0)) as realCheckErrShopBuyPrice"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.realCheckErrTenBuyPrice,0)) as realCheckErrTenBuyPrice"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.totSellNo,0)) as totSellNo"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.totrealsellprice,0)) as totrealsellprice"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.totShopBuyprice,0)) as totShopBuyprice"
        sqlStr = sqlStr + " 		, Sum(IsNULL(A.totTenBuyprice,0)) as totTenBuyprice"
        sqlStr = sqlStr + " 		from db_datamart.dbo.tbl_Shop_ItemTurnOver A"
        sqlStr = sqlStr + " 		where 1=1"

        if (FRectShopid<>"") then
            sqlStr = sqlStr + " 		and A.shopid='"&FRectShopid&"'"
        end if
        if (FRectYYYYMM<>"") then
            sqlStr = sqlStr + " 		and A.yyyymm='"&FRectYYYYMM&"'"
        end if
        if (FRectMakerid<>"") then
            sqlStr = sqlStr + " 		and A.makerid='"&FRectMakerid&"'"
        end if

        sqlStr = sqlStr + " 		group by A.itemgubun,A.itemid,A.itemoption"
        sqlStr = sqlStr + " 	) T"
        sqlStr = sqlStr + " 	order by StTurnOver desc"

        ''db3_rsget.pagesize = FPageSize
        db3_rsget.Open sqlStr,db3_dbget,1

        ''FtotalPage =  CInt(FTotalCount\FPageSize)
		''if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
		''	FtotalPage = FtotalPage +1
		''end if
		''FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
        ''if (FResultCount<1) then FResultCount=0
        FResultCount = db3_rsget.RecordCount
        FTotalCount  = FResultCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CBrandStockTurnOverMasterItem
				FItemList(i).Fshopid                    = db3_rsget("shopid")
                FItemList(i).Fmakerid                   = db3_rsget("makerid")
                FItemList(i).Fyyyymm                    = db3_rsget("yyyymm")
                FItemList(i).Fcomm_cd                   = db3_rsget("comm_cd")
                FItemList(i).Fcomm_name                 = db3_rsget("comm_name")
                FItemList(i).Fstockno                   = db3_rsget("stockno")
                FItemList(i).FtotSellno                 = db3_rsget("totSellno")
                FItemList(i).FStShopItemPrice           = db3_rsget("StShopItemPrice")
                FItemList(i).FStShopBuyPrice            = db3_rsget("StShopBuyPrice")
                FItemList(i).FpreStockNo                = db3_rsget("preStockNo")
                FItemList(i).FPreStShopItemPrice        = db3_rsget("PreStShopItemPrice")
                FItemList(i).FPreStShopBuyPrice         = db3_rsget("PreStShopBuyPrice")
                FItemList(i).FrealCheckErrNo            = db3_rsget("realCheckErrNo")
                FItemList(i).FrealCheckErrShopItemPrice = db3_rsget("realCheckErrShopItemPrice")
                FItemList(i).FrealCheckErrShopBuyPrice  = db3_rsget("realCheckErrShopBuyPrice")
                FItemList(i).FtotRealSellPrice          = db3_rsget("totRealSellPrice")
                FItemList(i).FtotTenBuyPrice            = db3_rsget("totTenBuyPrice")
                FItemList(i).FtotShopBuyPrice           = db3_rsget("totShopBuyPrice")
                FItemList(i).FStTurnOverBySell          = db3_rsget("StTurnOverBySell")
                FItemList(i).FStTurnOver                = db3_rsget("StTurnOver")
                FItemList(i).Fregdate                   = db3_rsget("regdate")


				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    End Sub

	public Sub getBrandStockTurnOverList
	    Dim sqlStr, sqlADD, i
	    sqlADD = ""
	    if (FRectYYYYMM<>"") then
            sqlADD = sqlADD & " and b.yyyymm='"&FRectYYYYMM&"'"
        end if

        if (FRectShopid<>"") then
            sqlADD = sqlADD & " and b.shopid='"&FRectShopid&"'"
        end if

        if (FRectMakerid<>"") then
            sqlADD = sqlADD & " and b.makerid='"&FRectMakerid&"'"
        end if

        if (FRectComm_cd<>"") then
            if (FRectComm_cd="B099") then
                sqlADD = sqlADD + " and b.comm_cd in ('B031','B011')"
            elseif (FRectComm_cd="B088") then
                sqlADD = sqlADD + " and b.comm_cd in ('B012','B022')"
            else
                sqlADD = sqlADD + " and b.comm_cd='" + FRectComm_cd + "'"
            end if
        end if

	    sqlStr = "select count(*) as CNT"
	    sqlStr = sqlStr & " from db_dataMart.dbo.tbl_Shop_BrandTurnOver b"
	    sqlStr = sqlStr & " 	left join db_datamart.dbo.tbl_jungsan_comm_code j"
        sqlStr = sqlStr & " 	on b.comm_cd=j.comm_cd"
        sqlStr = sqlStr & " where 1=1"
        sqlStr = sqlStr & sqlADD

        db3_rsget.Open sqlStr,db3_dbget,1
            FTotalCount = db3_rsget("cnt")
        db3_rsget.Close

	    sqlStr = "select b.*,j.comm_name"
        sqlStr = sqlStr & " from db_dataMart.dbo.tbl_Shop_BrandTurnOver b"
        sqlStr = sqlStr & " 	left join db_datamart.dbo.tbl_jungsan_comm_code j"
        sqlStr = sqlStr & " 	on b.comm_cd=j.comm_cd"
        sqlStr = sqlStr & " where 1=1"
        sqlStr = sqlStr & sqlADD
        sqlStr = sqlStr & " order by b.StTurnOver desc"

        db3_rsget.pagesize = FPageSize
        db3_rsget.Open sqlStr,db3_dbget,1
''rw sqlStr
        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CBrandStockTurnOverMasterItem
				FItemList(i).Fshopid                    = db3_rsget("shopid")
                FItemList(i).Fmakerid                   = db3_rsget("makerid")
                FItemList(i).Fyyyymm                    = db3_rsget("yyyymm")
                FItemList(i).Fcomm_cd                   = db3_rsget("comm_cd")
                FItemList(i).Fcomm_name                 = db3_rsget("comm_name")
                FItemList(i).Fstockno                   = db3_rsget("stockno")
                FItemList(i).FtotSellno                 = db3_rsget("totSellno")
                FItemList(i).FStShopItemPrice           = db3_rsget("StShopItemPrice")
                FItemList(i).FStShopBuyPrice            = db3_rsget("StShopBuyPrice")
                FItemList(i).FpreStockNo                = db3_rsget("preStockNo")
                FItemList(i).FPreStShopItemPrice        = db3_rsget("PreStShopItemPrice")
                FItemList(i).FPreStShopBuyPrice         = db3_rsget("PreStShopBuyPrice")
                FItemList(i).FrealCheckErrNo            = db3_rsget("realCheckErrNo")
                FItemList(i).FrealCheckErrShopItemPrice = db3_rsget("realCheckErrShopItemPrice")
                FItemList(i).FrealCheckErrShopBuyPrice  = db3_rsget("realCheckErrShopBuyPrice")
                FItemList(i).FtotRealSellPrice          = db3_rsget("totRealSellPrice")
                FItemList(i).FtotTenBuyPrice            = db3_rsget("totTenBuyPrice")
                FItemList(i).FtotShopBuyPrice           = db3_rsget("totShopBuyPrice")
                FItemList(i).FStTurnOverBySell          = db3_rsget("StTurnOverBySell")
                FItemList(i).FStTurnOver                = db3_rsget("StTurnOver")
                FItemList(i).Fregdate                   = db3_rsget("regdate")


				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

    end Sub

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

    public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

End Class
%>