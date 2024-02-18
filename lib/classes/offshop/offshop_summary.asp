<%
''한글
'###########################################################
' Description :  오프라인매장관리
' History : 2009.04.07 이상구 생성
'			2010.04.02 한용민 수정
'###########################################################

class CShopBrandCheckStockItem
    public FShopid
    public Fmakerid
    public FtotItemNo
    public FtotSellNo
    public FtotRealStockNo
    public FtotSysRealStockNo
    public FtotStockBuySum
    public FtotOwnStockBuySum   ''본사매입가
    public Ffirstipgodate
    public FlastStockdate
    public Fcomm_cd
    public Fcomm_name
    public FstTakingIdx
    public FstStatus
    public FstExistCnt
    public FstPLusStockItemCnt
    public Flastipgodate

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CShopItemSummaryItem
    public Fyyyymm
    public Fyyyymmdd
	public Fshopid
	public Fitemgubun
	public Fitemid
	public Fitemoption
    public FShopItemName
    public FShopItemOptionName
	public Fsellno
	public Fresellno
	public Flogicsipgono
	public Flogicsreipgono
	public Fbrandipgono
	public Fbrandreipgono
	public Fsysstockno
    public Ferrsampleitemno
    public Ferrbaditemno
    public Ferrrealcheckno
    public Frealstockno
    public Fregdate
    public Flastupdate
    public Fisusing
    public FImageSmall
	public FOffImgSmall
	public FCenterMwdiv
	public FMakerid
	public FBatchItemNo
	public fshopitemprice
	public fshopsuplycash   '''             // 본사매입가
	public Fshopbuyprice    ''' 2011 추가  // 매장매입가
	public forgsellprice
	public fdiscountsellprice
	public fdefaultmargin
	public fdefaultsuplymargin
	public FOnlinebuycash
	public FOnlineOptaddbuyprice

	public Fpreorderno
    public Fpreordernofix

    public FRegUserID
    public FmodiUserID

    public Fcomm_cd
    public Fcomm_name
    public FAccSysstockno

	Public Flogischulgo
	Public Flogisreturn

	''유효재고
    public function getAvailStock()
        getAvailStock = FrealstockNo + Ferrbaditemno + Ferrsampleitemno
    end Function

    public function getShopRealStock()
        getShopRealStock = FrealstockNo + Ferrbaditemno + Ferrsampleitemno + Flogischulgo + Flogisreturn
    end function

    public function getShopRealStockNoExc()
        getShopRealStockNoExc = FrealstockNo + Flogischulgo + Flogisreturn
    end function

    public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (FItemID >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + Format00(8,FItemId) + FItemOption
    	end if
	end function

	public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = FimageSmall
		else
			GetImageSmall = FOffImgSmall
		end if
	end function

	''직영점 공급시 매입가(업체로부터 매입하는가격) : 가맹점과 동일
	public function GetOfflineBuycash()
		GetOfflineBuycash = GetFranchiseBuycash
	end function

    ''직영점 공급가 : 가맹점과 동일
	public function GetOfflineSuplycash()
		GetOfflineSuplycash = GetFranchiseSuplycash
	end function

	''가맹점 공급시 매입가(업체로부터 매입하는가격) //본사매입가
	public function GetFranchiseBuycash()
		dim ibuycash
		''가맹점 매입가가 0 인경우 기본 마진으로 구한다
		if Fshopsuplycash<>0 then
			ibuycash = Fshopsuplycash
		else
		    'if (ISNULL(fdefaultmargin)) then fdefaultmargin=0

		    if Fshopitemprice <> "" and fdefaultmargin <> "" then
				ibuycash = CLng(Fshopitemprice * (100-fdefaultmargin)/100)

				''온라인 매입가보다 큰경우 온라인 매입가를 사용(Fshopsuplycash 가 지정된 경우는 제외)
				''200906 FOnlineOptaddbuyprice 추가
			    if (FOnlinebuycash<>0) and (ibuycash>FOnlinebuycash+FOnlineOptaddbuyprice) then ibuycash=FOnlinebuycash+FOnlineOptaddbuyprice
			else
				ibuycash = 0
			end if
		end if

		GetFranchiseBuycash = ibuycash
	end function

	''가맹점 공급가
	public function GetFranchiseSuplycash()
		dim ishopsupycash

		''가맹점공급가가 0 인경우 기본 마진으로 구한다
		if Fshopbuyprice<>0 then
			ishopsupycash = Fshopbuyprice
		else
		    ''마진이 설정 안되있는경우 매입마진-5%
		    if IsNULL(fdefaultsuplymargin) or (fdefaultsuplymargin=0) then
		        If isNULL(fdefaultmargin) then
		            ishopsupycash = CLng(Fshopitemprice * (100-(35-5))/100)
		        else
    		        ishopsupycash = CLng(Fshopitemprice * (100-(fdefaultmargin-5))/100)
    		    end if
		    else
			    ishopsupycash = CLng(Fshopitemprice * (100-fdefaultsuplymargin)/100)
			end if
		end if

		''공급가가 매입가보다 작은경우 공급가를 사용
		if (ishopsupycash<GetFranchiseBuycash) then ishopsupycash = GetFranchiseBuycash

		GetFranchiseSuplycash = ishopsupycash
	end function

	Private Sub Class_Initialize()
		FOnlineOptaddbuyprice = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CShopItemSummary
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectShopID
	public FRectItemGubun
	public FRectItemId
	public FRectItemOption
    public FRectMakerID
    public FRectCenterMwDiv
    public FRectIsUsing

    public FRectErrType
    public FRectStartDate
    public FRectEndDate

    public FRectBatchIdx
    public FRectComm_cd
    public FRectNoZeroStock
    public FRectGroupType
	public FRectShowMinusOnly

    ''브랜드별 배치 재고파악 list
    public function GetShopBrandBatchCheckList()
        dim sqlStr, i
        sqlStr = " select i.makerid"
        sqlStr = sqlStr + " , sum(IsNULL(s.realstockno,0)) as totSysRealStockNo"
        sqlStr = sqlStr + " , sum(i.itemno) as totRealStockNo"
        sqlStr = sqlStr + " , sum(i.itemno*i.suplyprice) as totSuplyPrice"
        sqlStr = sqlStr + " , d.firstipgodate, d.lastStockdate"
        sqlStr = sqlStr + " , d.comm_cd, j.comm_name"
        sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_tempstock_detail i"
        sqlStr = sqlStr + " 	left Join db_summary.dbo.tbl_current_shopstock_summary s"
        sqlStr = sqlStr + " 	on s.itemgubun=i.itemgubun"
        sqlStr = sqlStr + " 	and s.itemid=i.itemid"
        sqlStr = sqlStr + " 	and s.itemoption=i.itemoption"
        sqlStr = sqlStr + "     and s.shopid='"&FRectShopID&"'"
        sqlStr = sqlStr + "  	left join db_shop.dbo.tbl_shop_designer d"
        sqlStr = sqlStr + "     on d.shopid='"&FRectShopID&"'"
        sqlStr = sqlStr + "     and i.makerid=d.makerid"
        sqlStr = sqlStr + "  	left Join db_jungsan.dbo.tbl_jungsan_comm_code J"
        sqlStr = sqlStr + "  	on d.comm_cd=j.comm_cd"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and i.masteridx="&FRectBatchIdx&""
        sqlStr = sqlStr + " group by i.makerid, d.firstipgodate, d.lastStockdate , d.comm_cd, j.comm_name"
        sqlStr = sqlStr + " order by d.lastStockdate asc, totSuplyPrice desc,totRealStockNo desc"

        rsget.Open sqlStr,dbget,1
        FTotalCount  = rsget.RecordCount
		FResultCount = FTotalCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopBrandCheckStockItem

		        FItemList(i).FShopid            = FRectShopID
                FItemList(i).Fmakerid           = rsget("makerid")
                ''FItemList(i).FtotSellNo         = rsget("totSellNo")
                FItemList(i).FtotSysRealStockNo    = rsget("totSysRealStockNo")
                FItemList(i).FtotRealStockNo    = rsget("totRealStockNo")
                'FItemList(i).FtotStockBuySum    = rsget("totStockBuySum")
                FItemList(i).Ffirstipgodate     = rsget("firstipgodate")
                FItemList(i).FlastStockdate     = rsget("lastStockdate")
                FItemList(i).Fcomm_cd           = rsget("comm_cd")
                FItemList(i).Fcomm_name         = rsget("comm_name")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    ''브랜드별 재고파악 할 List
    '//common/offshop/brandRealStockCheckList.asp
    public function GetShopBrandRealCheckRequire()
        dim sqlStr, i

        sqlStr = " select"
        sqlStr = sqlStr + " T.makerid, T.totRealStockNo, T.totItemNo, T.totSellNo , T.totStockBuySum, T.totOwnStockBuySum"
        sqlStr = sqlStr + " ,T.firstipgodate, T.lastStockdate, T.comm_cd, j.comm_name, m.stTakingIdx, m.stStatus"
        sqlStr = sqlStr + " ,T.stExistCnt, T.stPLusStockItemCnt, T.lastipgodate"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "  	select i.makerid"
        sqlStr = sqlStr + " 	, sum(s.sellno) as totSellNo"
        sqlStr = sqlStr + "  	, sum(s.realstockno) as totRealStockNo"
        sqlStr = sqlStr + " 	, count(i.shopitemid) as totItemNo"
        sqlStr = sqlStr + " 	, sum(CASE WHEN s.realstockno<>0 then 1 ELSE 0 end) as stExistCnt"
        sqlStr = sqlStr + " 	, sum(CASE WHEN s.realstockno>0 then 1 ELSE 0 end) as stPLusStockItemCnt"
        sqlStr = sqlStr + " 	, sum((CASE WHEN s.realstockno<=0 then 0 WHEN i.shopbuyprice=0 then convert(int,i.shopitemprice*(100-defaultSuplymargin)/100) ELSE i.shopbuyprice end)*s.realstockno) as totStockBuySum"
        sqlStr = sqlStr + " 	, sum((CASE WHEN s.realstockno<=0 then 0 WHEN i.shopsuplycash=0 then convert(int,i.shopitemprice*(100-defaultmargin)/100) ELSE i.shopsuplycash end)*s.realstockno) as totOwnStockBuySum"
        sqlStr = sqlStr + " 	, d.firstipgodate, d.lastStockdate"
        sqlStr = sqlStr + " 	, d.comm_cd, d.lastipgodate"
        sqlStr = sqlStr + "  	from db_summary.dbo.tbl_current_shopstock_summary s"
        sqlStr = sqlStr + "  	Join db_shop.dbo.tbl_shop_item i"
        sqlStr = sqlStr + "  		on s.itemgubun=i.itemgubun"
        sqlStr = sqlStr + "  		and s.itemid=i.shopitemid"
        sqlStr = sqlStr + "  		and s.itemoption=i.itemoption"

        if (FRectComm_cd<>"") then
            sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shop_designer d"
            sqlStr = sqlStr + "  		on d.shopid='"+ FRectShopID + "'"
            sqlStr = sqlStr + "   		and i.makerid=d.makerid"

            if (FRectComm_cd="B099") then
                sqlStr = sqlStr + " and d.comm_cd in ('B031','B011')"
            elseif (FRectComm_cd="B088") then
                sqlStr = sqlStr + " and d.comm_cd in ('B012','B022')"
            elseif (FRectComm_cd="B077") then
                sqlStr = sqlStr + " and d.comm_cd in ('B031','B011','B013')"
            else
                sqlStr = sqlStr + " and d.comm_cd='" + FRectComm_cd + "'"
            end if
        else
            sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_designer d"
            sqlStr = sqlStr + "  		on d.shopid='"+ FRectShopID + "'"
            sqlStr = sqlStr + "   		and i.makerid=d.makerid"
        end if

        sqlStr = sqlStr + "  	where s.shopid='"+ FRectShopID + "'"


        if (FRectIsUsing<>"") then
            sqlStr = sqlStr + "  	and i.isusing='"&FRectIsUsing&"'"
        end if

        if (FRectMakerid<>"") then
            sqlStr = sqlStr + "  	and i.makerid='"&FRectMakerid&"'"
        end if

        sqlStr = sqlStr + "  	group by i.makerid, d.lastStockdate	, d.comm_cd, d.firstipgodate, d.lastipgodate"

		''오프상품테이블에 데이타 없는 상품
        sqlStr = sqlStr + " union all "
   		sqlStr = sqlStr + " select '10-' + convert(varchar,i.itemid) + '-' + s.itemoption  as makerid, sum(s.sellno) as totSellNo , sum(s.realstockno) as totRealStockNo , count(i.itemid) as totItemNo , sum(CASE WHEN s.realstockno<>0 then 1 ELSE 0 end) as stExistCnt , sum(CASE WHEN s.realstockno>0 then 1 ELSE 0 end) as stPLusStockItemCnt , NULL as totStockBuySum , NULL as totOwnStockBuySum , NULL as firstipgodate, NULL as lastStockdate , NULL as comm_cd, NULL as lastipgodate  "
    	sqlStr = sqlStr + " from db_summary.dbo.tbl_current_shopstock_summary s  "
   		sqlStr = sqlStr + " Join db_item.dbo.tbl_item i on s.itemgubun=i.itemgubun and s.itemid=i.itemid and s.itemgubun = '10' "
    	sqlStr = sqlStr + " Left Join db_shop.dbo.tbl_shop_item si on s.itemgubun=si.itemgubun and s.itemid=si.shopitemid and s.itemoption=si.itemoption  "
   		sqlStr = sqlStr + " where s.shopid='"+ FRectShopID + "' and si.itemgubun is NULL "

		if (FRectIsUsing<>"") then
			sqlStr = sqlStr + "  	and i.isusing='"&FRectIsUsing&"'"
		end if

		if (FRectMakerid<>"") then
			sqlStr = sqlStr + "  	and i.makerid='"&FRectMakerid&"'"
		end if

   		sqlStr = sqlStr + " group by '10-' + convert(varchar,i.itemid) + '-' + s.itemoption "

		sqlStr = sqlStr + "  ) T"
        sqlStr = sqlStr + "  left Join db_jungsan.dbo.tbl_jungsan_comm_code J"
        sqlStr = sqlStr + "  	on T.comm_cd=j.comm_cd"
        sqlStr = sqlStr + "  left Join db_shop.dbo.tbl_shop_stockTaking_master m"
        sqlStr = sqlStr + "  	on T.makerid=m.makerid"
        sqlStr = sqlStr + "  	and m.shopid='"+ FRectShopID + "'"
        sqlStr = sqlStr + "  	and m.stStatus in (0,3)"

        if (FRectNoZeroStock<>"") then
            sqlStr = sqlStr + " where T.stExistCnt>0"
        else
            sqlStr = sqlStr + " where 1=1"
        end if
        sqlStr = sqlStr + " order by T.lastStockdate asc, totStockBuySum desc,abs(T.totRealStockNo) desc, T.makerid"


        ''response.write sqlStr & "<Br>"
        rsget.Open sqlStr,dbget,1
        FTotalCount  = rsget.RecordCount
		FResultCount = FTotalCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopBrandCheckStockItem

		        FItemList(i).FShopid            = FRectShopID
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).FtotSellNo         = rsget("totSellNo")
                FItemList(i).FtotItemNo         = rsget("totItemNo")
                FItemList(i).FtotRealStockNo    = rsget("totRealStockNo")
                FItemList(i).FtotStockBuySum    = rsget("totStockBuySum")
                FItemList(i).FtotOwnStockBuySum = rsget("totOwnStockBuySum")
                FItemList(i).Ffirstipgodate     = rsget("firstipgodate")
                FItemList(i).Flastipgodate      = rsget("lastipgodate")
                FItemList(i).FlastStockdate     = rsget("lastStockdate")
                FItemList(i).Fcomm_cd           = rsget("comm_cd")
                FItemList(i).Fcomm_name         = rsget("comm_name")

                FItemList(i).FstTakingIdx       = rsget("stTakingIdx")
                FItemList(i).FstStatus          = rsget("stStatus")
                FItemList(i).FstExistCnt         = rsget("stExistCnt")
                FItemList(i).FstPLusStockItemCnt = rsget("stPLusStockItemCnt")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    ''샵별 브랜드별 재고합계
    public function GetShopItemCurrentSummaryByBrand
        dim sqlStr, i


        sqlStr = " select i.makerid,sum(s.realstockno) as realstocksum"
        sqlStr = sqlStr + " ,sum(i.shopitemprice*s.realstockno) as selltotal"
        sqlStr = sqlStr + " ,sum(CASE WHEN i.shopsuplycash=0 then i.shopitemprice*(100-IsNULL(d.defaultmargin,35))/100*s.realstockno ELSE (i.shopsuplycash*s.realstockno) END) as suplytotal"
        sqlStr = sqlStr + " ,sum(CASE WHEN i.shopbuyprice=0 then i.shopitemprice*(100-IsNULL(d.defaultsuplymargin,30))/100*s.realstockno ELSE (i.shopbuyprice*s.realstockno) END) as shopbuytotal"
        sqlStr = sqlStr + "  from "
        sqlStr = sqlStr + " 	db_summary.dbo.tbl_current_shopstock_summary s"
        sqlStr = sqlStr + " 	Join db_shop.dbo.tbl_shop_item i"
        sqlStr = sqlStr + " 	on s.shopid='"+ FRectShopID + "'"
        sqlStr = sqlStr + " 	and s.itemgubun=i.itemgubun"
        sqlStr = sqlStr + " 	and s.itemid=i.shopitemid"
        sqlStr = sqlStr + " 	and s.itemoption=i.itemoption"
        sqlStr = sqlStr + " 	left Join db_shop.dbo.tbl_shop_designer d"
        sqlStr = sqlStr + " 	on d.shopid='"+ FRectShopID + "'"
        sqlStr = sqlStr + " 	and i.makerid=d.makerid"
        sqlStr = sqlStr + " where i.isusing='Y'"
        sqlStr = sqlStr + " group by i.makerid"
        sqlStr = sqlStr + " order by realstocksum desc"

    end function

    public function GetShopCurrentStockByBatchJobByBrand()
        dim sqlStr, i

'		sqlStr = " select top 3000 '"&FRectShopID&"' as shopid, c.itemgubun, c.itemid, c.itemoption "
'		sqlStr = sqlStr + " ,IsNULL(c.sellno,0) as sellno, IsNULL(c.resellno,0) as resellno"
'		sqlStr = sqlStr + " ,IsNULL(c.logicsipgono,0) as logicsipgono, IsNULL(c.logicsreipgono,0) as logicsreipgono"
'		sqlStr = sqlStr + " ,IsNULL(c.brandipgono,0) as brandipgono, IsNULL(c.brandreipgono,0) as brandreipgono"
'		sqlStr = sqlStr + " ,IsNULL(c.sysstockno,0) as sysstockno,IsNULL(c.errsampleitemno,0) as errsampleitemno"
'		sqlStr = sqlStr + " ,IsNULL(c.errbaditemno,0) as errbaditemno,IsNULL(c.errrealcheckno,0) as errrealcheckno"
'		sqlStr = sqlStr + " ,IsNULL(c.realstockno,0) as realstockno"
'		sqlStr = sqlStr + " ,i.shopitemname, i.shopitemoptionname,i.isusing, i.centermwdiv, i.offimgsmall, o.smallimage"
'		sqlStr = sqlStr + " ,d.itemno as batchItemNo"
'		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_current_shopstock_summary c"
'		sqlStr = sqlStr + "     join db_shop.dbo.tbl_shop_item i"
'		sqlStr = sqlStr + "     on c.itemgubun=i.itemgubun"
'		sqlStr = sqlStr + "     and c.itemid=i.shopitemid"
'		sqlStr = sqlStr + "     and c.itemoption=i.itemoption"
'		sqlStr = sqlStr + "     and i.makerid='"&FRectMakerid&"'"
'		sqlStr = sqlStr + "     and c.shopid = '" + FRectShopID + "' "
'		sqlStr = sqlStr + "     left Join  [db_shop].[dbo].tbl_shop_tempstock_detail d"
'		sqlStr = sqlStr + "     on d.itemgubun=c.itemgubun"
'		sqlStr = sqlStr + "     and d.itemid=c.itemid"
'		sqlStr = sqlStr + "     and d.itemoption=c.itemoption"
'		sqlStr = sqlStr + "     and d.masteridx=" & FRectBatchIdx
'		sqlStr = sqlStr + "     left join db_item.dbo.tbl_item o"
'		sqlStr = sqlStr + "     on d.itemgubun='10'"
'		sqlStr = sqlStr + "     and d.itemid=o.itemid"
'		sqlStr = sqlStr + " where (c.realstockno<>0) or (d.itemid is Not NULL)"
'		sqlStr = sqlStr + " order by d.itemgubun, d.itemid, d.itemoption"


		sqlStr = " select top 3000 '"&FRectShopID&"' as shopid, i.itemgubun, i.shopitemid as itemid, i.itemoption "
		sqlStr = sqlStr + " ,IsNULL(c.sellno,0) as sellno, IsNULL(c.resellno,0) as resellno"
		sqlStr = sqlStr + " ,IsNULL(c.logicsipgono,0) as logicsipgono, IsNULL(c.logicsreipgono,0) as logicsreipgono"
		sqlStr = sqlStr + " ,IsNULL(c.brandipgono,0) as brandipgono, IsNULL(c.brandreipgono,0) as brandreipgono"
		sqlStr = sqlStr + " ,IsNULL(c.sysstockno,0) as sysstockno,IsNULL(c.errsampleitemno,0) as errsampleitemno"
		sqlStr = sqlStr + " ,IsNULL(c.errbaditemno,0) as errbaditemno,IsNULL(c.errrealcheckno,0) as errrealcheckno"
		sqlStr = sqlStr + " ,IsNULL(c.realstockno,0) as realstockno"
		sqlStr = sqlStr + " ,i.shopitemname, i.shopitemoptionname,i.isusing, i.centermwdiv, i.offimgsmall, o.smallimage"
		sqlStr = sqlStr + " ,d.itemno as batchItemNo"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_item i"
		sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_tempstock_detail d"
		sqlStr = sqlStr + "     on d.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=i.itemoption"
		sqlStr = sqlStr + "     and d.masteridx="&FRectBatchIdx&""
		sqlStr = sqlStr + "     left Join [db_summary].[dbo].tbl_current_shopstock_summary c"
		sqlStr = sqlStr + "     on i.itemgubun=c.itemgubun"
		sqlStr = sqlStr + "     and i.shopitemid=c.itemid"
		sqlStr = sqlStr + "     and i.itemoption=c.itemoption"
		sqlStr = sqlStr + "     and c.shopid = '" + FRectShopID + "' "
		sqlStr = sqlStr + "     left join db_item.dbo.tbl_item o"
		sqlStr = sqlStr + "     on i.itemgubun='10'"
		sqlStr = sqlStr + "     and i.shopitemid=o.itemid"
		sqlStr = sqlStr + " where i.makerid='"&FRectMakerid&"'"
		sqlStr = sqlStr + " and ((IsNULL(c.realstockno,0)<>0) or (IsNULL(d.itemno,0)<>0)) "
		sqlStr = sqlStr + " order by i.itemgubun, i.shopitemid, i.itemoption"

'''response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

        		        FItemList(i).Fshopid        = rsget("shopid")
        		        FItemList(i).Fitemgubun     = rsget("itemgubun")
        		        FItemList(i).Fitemid        = rsget("itemid")
        		        FItemList(i).Fitemoption    = rsget("itemoption")

                        FItemList(i).Fshopitemname      = db2Html(rsget("shopitemname"))
                        FItemList(i).FshopitemOptionname= db2Html(rsget("shopitemOptionname"))

        		        FItemList(i).Fsellno         = rsget("sellno")
        		        FItemList(i).Fresellno       = rsget("resellno")
        		        FItemList(i).Flogicsipgono   = rsget("logicsipgono")
        		        FItemList(i).Flogicsreipgono = rsget("logicsreipgono")
        		        FItemList(i).Fbrandipgono    = rsget("brandipgono")
        		        FItemList(i).Fbrandreipgono  = rsget("brandreipgono")

        		        FItemList(i).Fsysstockno    = rsget("sysstockno")

        		        FItemList(i).Ferrsampleitemno= rsget("errsampleitemno")
        		        FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
        		        FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
        		        FItemList(i).Frealstockno    = rsget("realstockno")

        		        FItemList(i).Fisusing        = rsget("isusing")
        		        FItemList(i).FCenterMwdiv    = rsget("centermwdiv")

        		        FItemList(i).FOffimgSmall	= rsget("offimgsmall")
        		        if FItemList(i).FOffimgSmall<>"" then
        		            FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
                        end if

            			FItemList(i).FimageSmall     = rsget("smallimage")
            			if FItemList(i).FimageSmall<>"" then
            				FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
            			end if

                        FItemList(i).FBatchItemNo = rsget("batchItemNo")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    public function GetShopCurrentStockByStockTaking()
        dim sqlStr, i

		sqlStr = " select top 3000 '"&FRectShopID&"' as shopid, i.itemgubun, i.shopitemid as itemid, i.itemoption "
		sqlStr = sqlStr + " ,IsNULL(c.sellno,0) as sellno, IsNULL(c.resellno,0) as resellno"
		sqlStr = sqlStr + " ,IsNULL(c.logicsipgono,0) as logicsipgono, IsNULL(c.logicsreipgono,0) as logicsreipgono"
		sqlStr = sqlStr + " ,IsNULL(c.brandipgono,0) as brandipgono, IsNULL(c.brandreipgono,0) as brandreipgono"
		sqlStr = sqlStr + " ,IsNULL(c.sysstockno,0) as sysstockno,IsNULL(c.errsampleitemno,0) as errsampleitemno"
		sqlStr = sqlStr + " ,IsNULL(c.errbaditemno,0) as errbaditemno,IsNULL(c.errrealcheckno,0) as errrealcheckno"
		sqlStr = sqlStr + " ,IsNULL(c.realstockno,0) as realstockno"
		sqlStr = sqlStr + " ,i.shopitemname, i.shopitemoptionname,i.isusing, i.centermwdiv, i.offimgsmall, o.smallimage"
		sqlStr = sqlStr + " ,d.stno as batchItemNo"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_item i"
		sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_stockTaking_Detail d"
		sqlStr = sqlStr + "     on d.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=i.itemoption"
		sqlStr = sqlStr + "     and d.stTakingIdx="&FRectBatchIdx&""
		sqlStr = sqlStr + "     left Join [db_summary].[dbo].tbl_current_shopstock_summary c"
		sqlStr = sqlStr + "     on i.itemgubun=c.itemgubun"
		sqlStr = sqlStr + "     and i.shopitemid=c.itemid"
		sqlStr = sqlStr + "     and i.itemoption=c.itemoption"
		sqlStr = sqlStr + "     and c.shopid = '" + FRectShopID + "' "
		sqlStr = sqlStr + "     left join db_item.dbo.tbl_item o"
		sqlStr = sqlStr + "     on i.itemgubun='10'"
		sqlStr = sqlStr + "     and i.shopitemid=o.itemid"
		sqlStr = sqlStr + " where i.makerid='"&FRectMakerid&"'"
		sqlStr = sqlStr + " and ((IsNULL(c.realstockno,0)<>0) or (IsNULL(d.stno,0)<>0)) "
		sqlStr = sqlStr + " order by i.itemgubun, i.shopitemid, i.itemoption"

'''response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

        		        FItemList(i).Fshopid        = rsget("shopid")
        		        FItemList(i).Fitemgubun     = rsget("itemgubun")
        		        FItemList(i).Fitemid        = rsget("itemid")
        		        FItemList(i).Fitemoption    = rsget("itemoption")

                        FItemList(i).Fshopitemname      = db2Html(rsget("shopitemname"))
                        FItemList(i).FshopitemOptionname= db2Html(rsget("shopitemOptionname"))

        		        FItemList(i).Fsellno         = rsget("sellno")
        		        FItemList(i).Fresellno       = rsget("resellno")
        		        FItemList(i).Flogicsipgono   = rsget("logicsipgono")
        		        FItemList(i).Flogicsreipgono = rsget("logicsreipgono")
        		        FItemList(i).Fbrandipgono    = rsget("brandipgono")
        		        FItemList(i).Fbrandreipgono  = rsget("brandreipgono")

        		        FItemList(i).Fsysstockno    = rsget("sysstockno")

        		        FItemList(i).Ferrsampleitemno= rsget("errsampleitemno")
        		        FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
        		        FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
        		        FItemList(i).Frealstockno    = rsget("realstockno")

        		        FItemList(i).Fisusing        = rsget("isusing")
        		        FItemList(i).FCenterMwdiv    = rsget("centermwdiv")

        		        FItemList(i).FOffimgSmall	= rsget("offimgsmall")
        		        if FItemList(i).FOffimgSmall<>"" then
        		            FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
                        end if

            			FItemList(i).FimageSmall     = rsget("smallimage")
            			if FItemList(i).FimageSmall<>"" then
            				FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
            			end if

                        FItemList(i).FBatchItemNo = rsget("batchItemNo")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    public function GetShopCurrentStockByBatchJob()
        dim sqlStr, i
        ''재고내역에 없을수 있음..

		sqlStr = " select top 2000 '"&FRectShopID&"' as shopid, d.itemgubun, d.itemid, d.itemoption "
		sqlStr = sqlStr + " ,IsNULL(c.sellno,0) as sellno, IsNULL(c.resellno,0) as resellno"
		sqlStr = sqlStr + " ,IsNULL(c.logicsipgono,0) as logicsipgono, IsNULL(c.logicsreipgono,0) as logicsreipgono"
		sqlStr = sqlStr + " ,IsNULL(c.brandipgono,0) as brandipgono, IsNULL(c.brandreipgono,0) as brandreipgono"
		sqlStr = sqlStr + " ,IsNULL(c.sysstockno,0) as sysstockno,IsNULL(c.errsampleitemno,0) as errsampleitemno"
		sqlStr = sqlStr + " ,IsNULL(c.errbaditemno,0) as errbaditemno,IsNULL(c.errrealcheckno,0) as errrealcheckno"
		sqlStr = sqlStr + " ,IsNULL(c.realstockno,0) as realstockno"
		sqlStr = sqlStr + " ,i.shopitemname, i.shopitemoptionname,i.isusing, i.centermwdiv, i.offimgsmall, o.smallimage"
		sqlStr = sqlStr + " ,d.itemno as batchItemNo"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_tempstock_detail d"
		sqlStr = sqlStr + "     left Join [db_summary].[dbo].tbl_current_shopstock_summary c "
		sqlStr = sqlStr + "     on d.itemgubun=c.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=c.itemid"
		sqlStr = sqlStr + "     and d.itemoption=c.itemoption"
		sqlStr = sqlStr + "     and c.shopid = '" + FRectShopID + "' "
		sqlStr = sqlStr + "     left join db_shop.dbo.tbl_shop_item i"
		sqlStr = sqlStr + "     on d.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=i.itemoption"
		sqlStr = sqlStr + "     left join db_item.dbo.tbl_item o"
		sqlStr = sqlStr + "     on d.itemgubun='10'"
		sqlStr = sqlStr + "     and d.itemid=o.itemid"

		sqlStr = sqlStr + " where d.masteridx=" & FRectBatchIdx

		if (FRectMakerID<>"") then
                sqlStr = sqlStr + " and i.makerid = '" + FRectMakerID + "' "
        end if
		if (FRectItemGubun <> "") then
		        sqlStr = sqlStr + " and c.itemgubun = '" + FRectItemGubun + "' "
		end if
		if (FRectItemId <> "") then
		        sqlStr = sqlStr + " and c.itemid = '" + CStr(FRectItemId) + "' "
		end if
		if (FRectItemOption <> "") then
		        sqlStr = sqlStr + " and c.itemoption = '" + FRectItemOption + "' "
		end if
		if (FRectCenterMwDiv<>"") then
		    if (FRectCenterMwDiv="NULL") then
		        sqlStr = sqlStr + " and i.centermwdiv is NULL"
		    elseif (FRectCenterMwDiv="MW") then
		        sqlStr = sqlStr + " and i.centermwdiv in ('M','W') "
		    else
		        sqlStr = sqlStr + " and i.centermwdiv = '" + FRectCenterMwDiv + "' "
		    end if
		end if
		if (FRectIsUsing<>"") then
		    sqlStr = sqlStr + " and i.isusing = '" + FRectIsUsing + "' "
		end if

		sqlStr = sqlStr + " order by d.itemgubun, d.itemid, d.itemoption"

''response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

        		        FItemList(i).Fshopid        = rsget("shopid")
        		        FItemList(i).Fitemgubun     = rsget("itemgubun")
        		        FItemList(i).Fitemid        = rsget("itemid")
        		        FItemList(i).Fitemoption    = rsget("itemoption")

                        FItemList(i).Fshopitemname      = db2Html(rsget("shopitemname"))
                        FItemList(i).FshopitemOptionname= db2Html(rsget("shopitemOptionname"))

        		        FItemList(i).Fsellno         = rsget("sellno")
        		        FItemList(i).Fresellno       = rsget("resellno")
        		        FItemList(i).Flogicsipgono   = rsget("logicsipgono")
        		        FItemList(i).Flogicsreipgono = rsget("logicsreipgono")
        		        FItemList(i).Fbrandipgono    = rsget("brandipgono")
        		        FItemList(i).Fbrandreipgono  = rsget("brandreipgono")

        		        FItemList(i).Fsysstockno    = rsget("sysstockno")

        		        FItemList(i).Ferrsampleitemno= rsget("errsampleitemno")
        		        FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
        		        FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
        		        FItemList(i).Frealstockno    = rsget("realstockno")

        		        FItemList(i).Fisusing        = rsget("isusing")
        		        FItemList(i).FCenterMwdiv    = rsget("centermwdiv")

        		        FItemList(i).FOffimgSmall	= rsget("offimgsmall")
        		        if FItemList(i).FOffimgSmall<>"" then
        		            FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
                        end if

            			FItemList(i).FimageSmall     = rsget("smallimage")
            			if FItemList(i).FimageSmall<>"" then
            				FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
            			end if

                        FItemList(i).FBatchItemNo = rsget("batchItemNo")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    public function GetOFFErrItemSummaryGroupByItem()
        dim sqlStr, i
        sqlStr = " select top 2000 s.shopid "
        sqlStr = sqlStr + " , s.itemgubun, s.shopitemid, s.itemoption"
        sqlStr = sqlStr + " , i.shopitemname, i.shopitemoptionname"
        sqlStr = sqlStr + " , sum(errrealcheckno) as errrealcheckno"
        sqlStr = sqlStr + " , (shopitemprice) as shopitempriceSum"
        sqlStr = sqlStr + " , ((CASE WHEN i.shopbuyprice=0 then convert(int,i.shopitemprice*(100-IsNULL(d.defaultSuplymargin,30))/100) ELSE i.shopbuyprice end)) as totStockBuySum"
        sqlStr = sqlStr + " , ((CASE WHEN i.shopsuplycash=0 then convert(int,i.shopitemprice*(100-IsNULL(d.defaultmargin,35))/100) ELSE i.shopsuplycash end)) as totOwnStockBuySum"
        sqlStr = sqlStr + " from"
        sqlStr = sqlStr + " [db_summary].[dbo].tbl_erritem_shop_summary s "
        sqlStr = sqlStr + "     left join db_shop.dbo.tbl_shop_item i "
        sqlStr = sqlStr + "     on s.itemgubun=i.itemgubun"
        sqlStr = sqlStr + "     and s.shopitemid=i.shopitemid"
        sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
        sqlStr = sqlStr + " 	left Join db_shop.dbo.tbl_shop_designer d"
        sqlStr = sqlStr + " 	on d.shopid=s.shopid"
        sqlStr = sqlStr + " 	and i.makerid=d.makerid"
        sqlStr = sqlStr + " where 1=1"
        if (FRectShopID<>"") then
            sqlStr = sqlStr + " and s.shopid='" + FRectShopID + "'"
        end if

        if (FRectStartDate<>"") then
            sqlStr = sqlStr + " and s.yyyymmdd>='" + FRectStartDate + "'"
        end if

        if (FRectEndDate<>"") then
            sqlStr = sqlStr + " and s.yyyymmdd<'" + FRectEndDate + "'"
        end if

        if (FRectMakerID<>"") then
                sqlStr = sqlStr + " and i.makerid = '" + FRectMakerID + "' "
        end if
        sqlStr = sqlStr + " group by s.shopid, s.itemgubun, s.shopitemid, s.itemoption, i.shopitemname, i.shopitemoptionname "
        sqlStr = sqlStr + " , (shopitemprice)"
        sqlStr = sqlStr + " , ((CASE WHEN i.shopbuyprice=0 then convert(int,i.shopitemprice*(100-IsNULL(d.defaultSuplymargin,30))/100) ELSE i.shopbuyprice end)) "
        sqlStr = sqlStr + " , ((CASE WHEN i.shopsuplycash=0 then convert(int,i.shopitemprice*(100-IsNULL(d.defaultmargin,35))/100) ELSE i.shopsuplycash end)) "
        sqlStr = sqlStr + " having sum(errrealcheckno)<>0"
''rw sqlStr
        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

		        FItemList(i).Fshopid        = rsget("shopid")
		        FItemList(i).FMakerid       = FRectMakerID

		        FItemList(i).Fitemgubun     = rsget("itemgubun")
		        FItemList(i).Fitemid        = rsget("shopitemid")
		        FItemList(i).Fitemoption    = rsget("itemoption")

		        FItemList(i).Fshopitemname    = rsget("shopitemname")
		        FItemList(i).Fshopitemoptionname    = rsget("shopitemoptionname")

		        'FItemList(i).Fsysstockno    = rsget("sysstockno")

		        ''FItemList(i).Ferrsampleitemno= rsget("errsampleitemno")
		        ''FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
		        FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")

		        FItemList(i).Fshopitemprice = rsget("shopitempriceSum")
		        FItemList(i).Fshopbuyprice  = rsget("totStockBuySum")
		        FItemList(i).fshopsuplycash  = rsget("totOwnStockBuySum")
		        'FItemList(i).Frealstockno    = rsget("realstockno")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	'//common/offshop/shopErrSummary.asp
    public function GetOFFErrItemSummary()
        dim sqlStr, i

        sqlStr = " select top 2000 s.shopid "

        IF (FRectGroupType="M") then
            sqlStr = sqlStr + " , convert(varchar(7),s.YYYYMMDD,21) as YYYYMMDD"
        end if

        IF (FRectShopID<>"") then
            sqlStr = sqlStr + " , i.makerid"
            sqlStr = sqlStr + " , d.defaultmargin, d.defaultsuplymargin, d.comm_cd, c1.comm_name"
        end if

        sqlStr = sqlStr + " , sum(errsampleitemno) as errsampleitemno, sum(errbaditemno) as errbaditemno, sum(errrealcheckno) as errrealcheckno"
        sqlStr = sqlStr + " , sum(errrealcheckno*shopitemprice) as shopitempriceSum"
        sqlStr = sqlStr + " 	, sum((CASE WHEN i.shopbuyprice=0 then convert(int,i.shopitemprice*(100-IsNULL(d.defaultSuplymargin,30))/100) ELSE i.shopbuyprice end)*s.errrealcheckno) as totStockBuySum"
        sqlStr = sqlStr + " 	, sum((CASE WHEN i.shopsuplycash=0 then convert(int,i.shopitemprice*(100-IsNULL(d.defaultmargin,35))/100) ELSE i.shopsuplycash end)*s.errrealcheckno) as totOwnStockBuySum"
        ''sqlStr = sqlStr + " , sum(s.sellno) as sellno, sum(s.resellno) as resellno"
        sqlStr = sqlStr + " from"
        sqlStr = sqlStr + " [db_summary].[dbo].tbl_erritem_shop_summary s "
        sqlStr = sqlStr + "     left join db_shop.dbo.tbl_shop_item i "
        sqlStr = sqlStr + "     on s.itemgubun=i.itemgubun"
        sqlStr = sqlStr + "     and s.shopitemid=i.shopitemid"
        sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
        sqlStr = sqlStr + " 	left Join db_shop.dbo.tbl_shop_designer d"
        sqlStr = sqlStr + " 	on d.shopid=s.shopid"
        sqlStr = sqlStr + " 	and i.makerid=d.makerid"
        sqlStr = sqlStr + " 	left join db_jungsan.dbo.tbl_jungsan_comm_code c1"
        sqlStr = sqlStr + " 	on d.comm_cd=c1.comm_cd"
'        sqlStr = sqlStr + " 	left Join db_summary.dbo.tbl_monthly_shop_designer sd"
'        sqlStr = sqlStr + " 	on s.shopid=sd.shopid"
'        sqlStr = sqlStr + " 	and i.makerid=sd.makerid"

        sqlStr = sqlStr + " where 1=1"
        if (FRectShopID<>"") then
            sqlStr = sqlStr + " and s.shopid='" + FRectShopID + "'"
        end if

        if (FRectStartDate<>"") then
            sqlStr = sqlStr + " and s.yyyymmdd>='" + FRectStartDate + "'"
        end if

        if (FRectEndDate<>"") then
            sqlStr = sqlStr + " and s.yyyymmdd<'" + FRectEndDate + "'"
        end if

        if (FRectItemGubun<>"") then
            sqlStr = sqlStr + " and s.itemgubun='" + FRectItemGubun + "'"
        end if

        if (FRectItemid<>"") then
            sqlStr = sqlStr + " and  s.shopitemid=" + FRectItemid + ""
        end if

        if (FRectItemOption<>"") then
            sqlStr = sqlStr + " and s.itemoption='" + FRectItemOption + "'"
        end if

        if (FRectMakerID<>"") then
			sqlStr = sqlStr + " and i.makerid = '" + FRectMakerID + "' "
		end if

        if (FRectShopID <> "") and (FRectComm_cd<>"") then
            if (FRectComm_cd="B099") then
                sqlStr = sqlStr + " and d.comm_cd in ('B031','B011')"
            elseif (FRectComm_cd="B088") then
                sqlStr = sqlStr + " and d.comm_cd in ('B012','B022')"
            elseif (FRectComm_cd="B077") then
                sqlStr = sqlStr + " and d.comm_cd in ('B031','B011','B013')"
            else
                sqlStr = sqlStr + " and d.comm_cd='" + FRectComm_cd + "'"
            end if
        end if

        sqlStr = sqlStr + " group by s.shopid "
        IF (FRectGroupType="M") then
            sqlStr = sqlStr + " , convert(varchar(7),s.YYYYMMDD,21)"
        end if
        IF (FRectShopID<>"") then
            sqlStr = sqlStr + ", i.makerid"
            sqlStr = sqlStr + ", d.defaultmargin, d.defaultsuplymargin, d.comm_cd, c1.comm_name "
        end if
        'response.write sqlStr & "<br>"
        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

				        IF (FRectGroupType="M") then
                            FItemList(i).Fyyyymmdd      = rsget("yyyymmdd")
                        ELSE
                            FItemList(i).Fyyyymmdd      = FRectStartDate&"~"&DateAdd("d",-1,FRectEndDate)
                        END IF

        		        FItemList(i).Fshopid        = rsget("shopid")

        		        IF (FRectShopID<>"") then
        		            FItemList(i).FMakerid       = rsget("makerid")
        		            FItemList(i).Fdefaultmargin         = rsget("defaultmargin")
            		        FItemList(i).Fdefaultsuplymargin    = rsget("defaultsuplymargin")
            		        FItemList(i).Fcomm_cd               = rsget("comm_cd")
            		        FItemList(i).Fcomm_name             = rsget("comm_name")
        		        ELSE
        		            FItemList(i).FMakerid       = "전체"
        		        END IF

        		        FItemList(i).Ferrsampleitemno= rsget("errsampleitemno")
        		        FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
        		        FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")

        		        FItemList(i).Fshopitemprice = rsget("shopitempriceSum")
        		        FItemList(i).Fshopbuyprice  = rsget("totStockBuySum")
        		        FItemList(i).fshopsuplycash  = rsget("totOwnStockBuySum")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	'불량 상품 목록		'//admin/stock/off_baditem_list.asp
    public function GetOFFDailyErrItemList()
        dim sqlStr, i

        if FRectShopID = "" then exit function

        sqlStr = " select top 1000 s.*, i.makerid, i.shopitemname, i.shopitemoptionname, i.isusing, i.centermwdiv,sd.comm_cd ,c1.comm_name"
        sqlStr = sqlStr + " from [db_summary].[dbo].tbl_erritem_shop_summary s"
        sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_item i"
        sqlStr = sqlStr + "     on s.itemgubun=i.itemgubun"
        sqlStr = sqlStr + "     and s.shopitemid=i.shopitemid"
        sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_designer sd"
		sqlStr = sqlStr + " 	on s.shopid=sd.shopid"
		sqlStr = sqlStr + " 	and i.makerid=sd.makerid"
        sqlStr = sqlStr + " left join db_jungsan.dbo.tbl_jungsan_comm_code c1"
        sqlStr = sqlStr + " 	on sd.comm_cd=c1.comm_cd"
        ''sqlStr = sqlStr + " where (s.errsampleitemno<>0 or s.errrealcheckno<>0)"
        sqlStr = sqlStr + " where 1=1"

        if (FRectShopID<>"") then
            sqlStr = sqlStr + " and s.shopid='" + FRectShopID + "'"
        end if

        if (FRectStartDate<>"") then
            sqlStr = sqlStr + " and s.yyyymmdd>='" + FRectStartDate + "'"
        end if

        if (FRectEndDate<>"") then
            sqlStr = sqlStr + " and s.yyyymmdd<'" + FRectEndDate + "'"
        end if

        if (FRectItemGubun<>"") then
            sqlStr = sqlStr + " and s.itemgubun='" + FRectItemGubun + "'"
        end if

        if (FRectItemid<>"") then
            sqlStr = sqlStr + " and  s.shopitemid=" + FRectItemid + ""
        end if

        if (FRectItemOption<>"") then
            sqlStr = sqlStr + " and s.itemoption='" + FRectItemOption + "'"
        end if

        if (FRectMakerID<>"") then
                sqlStr = sqlStr + " and i.makerid = '" + FRectMakerID + "' "
        end if
        rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

						FItemList(i).fcomm_name      = rsget("comm_name")
						FItemList(i).fcomm_cd      = rsget("comm_cd")
                        FItemList(i).Fyyyymmdd      = rsget("yyyymmdd")
        		        FItemList(i).Fshopid        = rsget("shopid")
        		        FItemList(i).Fitemgubun     = rsget("itemgubun")
        		        FItemList(i).Fitemid        = rsget("shopitemid")
        		        FItemList(i).Fitemoption    = rsget("itemoption")
        		        FItemList(i).FMakerid       = rsget("makerid")

                        FItemList(i).Fshopitemname      = db2Html(rsget("shopitemname"))
                        FItemList(i).FshopitemOptionname= db2Html(rsget("shopitemOptionname"))

                        FItemList(i).FRegUserID       = rsget("RegUser")
                        FItemList(i).FModiUserId      = rsget("ModiUser")

        		        'FItemList(i).Fsellno         = rsget("sellno")
        		        'FItemList(i).Fresellno       = rsget("resellno")
        		        'FItemList(i).Flogicsipgono   = rsget("logicsipgono")
        		        'FItemList(i).Flogicsreipgono = rsget("logicsreipgono")
        		        'FItemList(i).Fbrandipgono    = rsget("brandipgono")
        		        'FItemList(i).Fbrandreipgono  = rsget("brandreipgono")

        		        'FItemList(i).Fsysstockno    = rsget("sysstockno")

        		        FItemList(i).Ferrsampleitemno= rsget("errsampleitemno")
        		        FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
        		        FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
        		        'FItemList(i).Frealstockno    = rsget("realstockno")

        		        FItemList(i).Fisusing        = rsget("isusing")
        		        FItemList(i).FCenterMwdiv    = rsget("centermwdiv")

'        		        FItemList(i).FOffimgSmall	= rsget("offimgsmall")
'        		        if FItemList(i).FOffimgSmall<>"" then
'        		            FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
'                        end if
'
'            			FItemList(i).FimageSmall     = rsget("smallimage")
'            			if FItemList(i).FimageSmall<>"" then
'            				FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
'            			end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	'샵아이템현재재고
	public function GetShopItemCurrentSummary()
		dim sqlStr, i

		sqlStr = " select top 1 c.shopid, c.itemgubun, c.itemid, c.itemoption, c.sellno,"
		sqlStr = sqlStr + " c.resellno, c.logicsipgono, c.logicsreipgono, c.brandipgono, c.brandreipgono, c.sysstockno, "
		sqlStr = sqlStr + " c.errsampleitemno, c.errbaditemno, c.errrealcheckno, c.realstockno, c.regdate, c.lastupdate, "
		sqlStr = sqlStr + " c.preorderno, c.preordernofix, c.logischulgo, c.logisreturn "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_current_shopstock_summary c "
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and c.shopid = '" + FRectShopID + "' "
		sqlStr = sqlStr + " and c.itemgubun = '" + FRectItemGubun + "' "
		sqlStr = sqlStr + " and c.itemid = " + CStr(FRectItemId) + " "
		sqlStr = sqlStr + " and c.itemoption = '" + FRectItemOption + "' "
		rsget.Open sqlStr,dbget,1

		set FOneItem = new CShopItemSummaryItem

		if (not rsget.EOF) then
		        FOneItem.Fshopid        = rsget("shopid")
		        FOneItem.Fitemgubun     = rsget("itemgubun")
		        FOneItem.Fitemid        = rsget("itemid")
		        FOneItem.Fitemoption    = rsget("itemoption")

		        FOneItem.Fsellno         = rsget("sellno")
		        FOneItem.Fresellno       = rsget("resellno")
		        FOneItem.Flogicsipgono   = rsget("logicsipgono")
		        FOneItem.Flogicsreipgono = rsget("logicsreipgono")
		        FOneItem.Fbrandipgono    = rsget("brandipgono")
		        FOneItem.Fbrandreipgono  = rsget("brandreipgono")

		        FOneItem.Fsysstockno    = rsget("sysstockno")

		        FOneItem.Ferrsampleitemno= rsget("errsampleitemno")
		        FOneItem.Ferrbaditemno   = rsget("errbaditemno")
		        FOneItem.Ferrrealcheckno = rsget("errrealcheckno")
		        FOneItem.Frealstockno    = rsget("realstockno")

		        FOneItem.Fpreorderno    = rsget("preorderno")
		        FOneItem.Fpreordernofix = rsget("preordernofix")

				FOneItem.Flogischulgo = rsget("logischulgo")
				FOneItem.Flogisreturn = rsget("logisreturn")
		end if
		rsget.Close
	end function

    '샵아이템 현재재고 목록
	public function GetShopItemCurrentSummaryList()
		dim sqlStr, i

		sqlStr = " select top " & FCurrPage*FPageSize & " c.shopid, c.itemgubun, c.itemid, c.itemoption "
		sqlStr = sqlStr + " ,c.sellno, c.resellno, c.logicsipgono, c.logicsreipgono, c.brandipgono, c.brandreipgono, c.sysstockno "
		sqlStr = sqlStr + " ,c.errsampleitemno, c.errbaditemno, c.errrealcheckno, c.realstockno , IsNULL(o.buycash,0) as onlinebuycash"
		sqlStr = sqlStr + " ,i.shopitemname, i.shopitemoptionname,i.isusing, i.centermwdiv, i.offimgsmall"
		sqlStr = sqlStr + " ,i.shopitemprice, i.shopsuplycash,i.shopbuyprice,  i.orgsellprice , o.smallimage,d.defaultmargin,d.defaultsuplymargin, c.logischulgo, c.logisreturn"
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_current_shopstock_summary c "
		sqlStr = sqlStr + "     left join db_shop.dbo.tbl_shop_item i"
		sqlStr = sqlStr + "     on c.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and c.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and c.itemoption=i.itemoption"
		sqlStr = sqlStr + "     left join db_item.dbo.tbl_item o"
		sqlStr = sqlStr + "     on c.itemgubun='10'"
		sqlStr = sqlStr + "     and c.itemid=o.itemid"
		sqlStr = sqlStr + "		left join [db_shop].[dbo].tbl_shop_designer d "
		sqlStr = sqlStr + "		on d.shopid='" + FRectShopid + "' and i.makerid=d.makerid"

		sqlStr = sqlStr + " where 1 = 1 "

		if (FRectShopID <> "") then
		    sqlStr = sqlStr + " and c.shopid = '" + FRectShopID + "' "
		end if
		if (FRectMakerID<>"") then
            sqlStr = sqlStr + " and i.makerid = '" + FRectMakerID + "' "
        end if
		if (FRectItemGubun <> "") then
		    sqlStr = sqlStr + " and c.itemgubun = '" + FRectItemGubun + "' "
		end if
		if (FRectItemId <> "") then
		    sqlStr = sqlStr + " and c.itemid = '" + CStr(FRectItemId) + "' "
		end if
		if (FRectItemOption <> "") then
		    sqlStr = sqlStr + " and c.itemoption = '" + FRectItemOption + "' "
		end if
		if (FRectCenterMwDiv<>"") then
		    if (FRectCenterMwDiv="NULL") then
		        sqlStr = sqlStr + " and i.centermwdiv is NULL"
		    elseif (FRectCenterMwDiv="MW") then
		        sqlStr = sqlStr + " and i.centermwdiv in ('M','W') "
		    else
		        sqlStr = sqlStr + " and i.centermwdiv = '" + FRectCenterMwDiv + "' "
		    end if
		end if
		if (FRectIsUsing<>"") then
		    sqlStr = sqlStr + " and i.isusing = '" + FRectIsUsing + "' "
		end if
		if (FRectNoZeroStock<>"") then
		    sqlStr = sqlStr + " and c.realstockno<>0"
		end if
		if (FRectShowMinusOnly<>"") then
		    sqlStr = sqlStr + " and c.realstockno<0"
		end if
		sqlStr = sqlStr + " order by c.itemgubun, c.itemid, c.itemoption"

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

				FItemList(i).fshopitemprice = rsget("shopitemprice")
				FItemList(i).FOnlinebuycash = rsget("onlinebuycash")
				FItemList(i).fdefaultmargin = rsget("defaultmargin")
				FItemList(i).fdefaultsuplymargin = rsget("defaultsuplymargin")
        		FItemList(i).Fshopid        = rsget("shopid")
        		FItemList(i).Fitemgubun     = rsget("itemgubun")
        		FItemList(i).Fitemid        = rsget("itemid")
        		FItemList(i).Fitemoption    = rsget("itemoption")
        		FItemList(i).fshopsuplycash    = rsget("shopsuplycash")
        		FItemList(i).Fshopbuyprice     = rsget("shopbuyprice")
        		FItemList(i).fshopitemprice    = rsget("shopitemprice")
                FItemList(i).Fshopitemname      = db2Html(rsget("shopitemname"))
                FItemList(i).FshopitemOptionname= db2Html(rsget("shopitemOptionname"))
        		FItemList(i).Fsellno         = rsget("sellno")
        		FItemList(i).Fresellno       = rsget("resellno")
        		FItemList(i).Flogicsipgono   = rsget("logicsipgono")
        		FItemList(i).Flogicsreipgono = rsget("logicsreipgono")
        		FItemList(i).Fbrandipgono    = rsget("brandipgono")
        		FItemList(i).Fbrandreipgono  = rsget("brandreipgono")
        		FItemList(i).Fsysstockno    = rsget("sysstockno")
        		FItemList(i).Ferrsampleitemno= rsget("errsampleitemno")
        		FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
        		FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
        		FItemList(i).Frealstockno    = rsget("realstockno")
        		FItemList(i).Fisusing        = rsget("isusing")
        		FItemList(i).FCenterMwdiv    = rsget("centermwdiv")
        		FItemList(i).FOffimgSmall	= rsget("offimgsmall")
        		if FItemList(i).FOffimgSmall<>"" then
        		    FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
                end if

            	FItemList(i).FimageSmall     = rsget("smallimage")
            	if FItemList(i).FimageSmall<>"" then
            		FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
            	end if

				FItemList(i).Flogischulgo = rsget("logischulgo")
				FItemList(i).Flogisreturn = rsget("logisreturn")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	Public function GetDirectShopList()
		dim sqlStr, i

		sqlStr = " EXEC [db_summary].[dbo].[usp_Ten_brandStockByShop_GetShopList] '" & FRectShopID & "', '" & FRectMakerID & "' "
		rsget.CursorLocation = adUseClient
		rsget.open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		''response.write sqlStr
		''dbget.close : response.end
		FTotalCount = rsget.RecordCount
		FresultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

				FItemList(i).Fshopid        = rsget("shopid")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	Public function GetDirectShopBrandList()
		dim sqlStr, i

		sqlStr = " EXEC [db_summary].[dbo].[usp_Ten_brandStockByShop_GetList] '" & FRectShopID & "', '" & FRectMakerID & "' "
		''response.write sqlStr
		''dbget.close : response.end
		rsget.CursorLocation = adUseClient
		rsget.open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FresultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		''response.write FResultCount
		''dbget.close : response.end


		If Not rsget.EOF Then
			GetDirectShopBrandList = rsget.getRows()
		End If
		rsget.close
	end function

    '샵아이템월별재고목록
	public function GetShopItemMonthlySummaryList()
		dim sqlStr, i
		dim month_pre_2

		month_pre_2 = Left(dateadd("m", -2, now()), 7)


		sqlStr = " select top 1000 c.shopid, c.itemgubun, c.itemid, c.itemoption"
		sqlStr = sqlStr + " ,c.yyyymm, c.sellno, c.resellno, c.logicsipgono, c.logicsreipgono, c.brandipgono, c.brandreipgono, c.sysstockno "
		sqlStr = sqlStr + " ,c.errsampleitemno, c.errbaditemno, c.errrealcheckno, c.realstockno"
		sqlStr = sqlStr + " ,ss.lstComm_cd,ss.lstCenterMwdiv,ss.sysstockno as lstsysstockno"
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_shopstock_summary c "
		sqlStr = sqlStr + "     left join db_summary.dbo.tbl_monthly_accumulated_shopstock_summary ss"
		sqlStr = sqlStr + "     on c.yyyymm=ss.yyyymm"
		sqlStr = sqlStr + "     and c.shopid=ss.shopid"
		sqlStr = sqlStr + "     and c.itemgubun=ss.itemgubun"
		sqlStr = sqlStr + "     and c.itemid=ss.itemid"
		sqlStr = sqlStr + "     and c.itemoption=ss.itemoption"
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and c.yyyymm <= '" + month_pre_2 + "' "


		if (FRectShopID <> "") then
		        sqlStr = sqlStr + " and c.shopid = '" + FRectShopID + "' "
		end if
		if (FRectItemGubun <> "") then
		        sqlStr = sqlStr + " and c.itemgubun = '" + FRectItemGubun + "' "
		end if
		if (FRectItemId <> "") then
		        sqlStr = sqlStr + " and c.itemid = " + CStr(FRectItemId) + " "
		end if
		if (FRectItemOption <> "") then
		        sqlStr = sqlStr + " and c.itemoption = '" + FRectItemOption + "' "
		end if
		sqlStr = sqlStr + " order by c.yyyymm, c.shopid, c.itemgubun, c.itemid, c.itemoption "
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly


		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

                        FItemList(i).Fyyyymm        = rsget("yyyymm")

        		        FItemList(i).Fshopid        = rsget("shopid")
        		        FItemList(i).Fitemgubun     = rsget("itemgubun")
        		        FItemList(i).Fitemid        = rsget("itemid")
        		        FItemList(i).Fitemoption    = rsget("itemoption")

        		        FItemList(i).Fsellno         = rsget("sellno")
        		        FItemList(i).Fresellno       = rsget("resellno")
        		        FItemList(i).Flogicsipgono   = rsget("logicsipgono")
        		        FItemList(i).Flogicsreipgono = rsget("logicsreipgono")
        		        FItemList(i).Fbrandipgono    = rsget("brandipgono")
        		        FItemList(i).Fbrandreipgono  = rsget("brandreipgono")

        		        FItemList(i).Fsysstockno    = rsget("sysstockno")

                        FItemList(i).Ferrsampleitemno= rsget("errsampleitemno")
        		        FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
        		        FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
        		        FItemList(i).Frealstockno    = rsget("realstockno")

                        FItemList(i).Fcomm_cd       = rsget("lstComm_cd")
                        FItemList(i).FCenterMwdiv   = rsget("lstCenterMwdiv")
                        FItemList(i).FAccSysstockno   = rsget("lstsysstockno")


				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

        '샵아이템현재재고(last month)
	public function GetShopItemLastMonthSummary()
		dim sqlStr, i

		sqlStr = " select top 1 c.shopid, c.itemgubun, c.itemid, c.itemoption"
		sqlStr = sqlStr + " ,c.sellno, c.resellno, c.logicsipgono, c.logicsreipgono, c.brandipgono, c.brandreipgono, c.sysstockno "
		sqlStr = sqlStr + " ,c.errsampleitemno, c.errbaditemno, c.errrealcheckno, c.realstockno"
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_last_monthly_shopstock c "
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " and c.shopid = '" + FRectShopID + "' "
		sqlStr = sqlStr + " and c.itemgubun = '" + FRectItemGubun + "' "
		sqlStr = sqlStr + " and c.itemid = " + CStr(FRectItemId) + " "
		sqlStr = sqlStr + " and c.itemoption = '" + FRectItemOption + "' "
''rw sqlStr
		rsget.Open sqlStr,dbget,1

		set FOneItem = new CShopItemSummaryItem

		if (not rsget.EOF) then
		        FOneItem.Fshopid        = rsget("shopid")
		        FOneItem.Fitemgubun     = rsget("itemgubun")
		        FOneItem.Fitemid        = rsget("itemid")
		        FOneItem.Fitemoption    = rsget("itemoption")

		        FOneItem.Fsellno         = rsget("sellno")
		        FOneItem.Fresellno       = rsget("resellno")
		        FOneItem.Flogicsipgono   = rsget("logicsipgono")
		        FOneItem.Flogicsreipgono = rsget("logicsreipgono")
		        FOneItem.Fbrandipgono    = rsget("brandipgono")
		        FOneItem.Fbrandreipgono  = rsget("brandreipgono")

		        FOneItem.Fsysstockno    = rsget("sysstockno")

		        FOneItem.Ferrsampleitemno= rsget("errsampleitemno")
		        FOneItem.Ferrbaditemno   = rsget("errbaditemno")
		        FOneItem.Ferrrealcheckno = rsget("errrealcheckno")
		        FOneItem.Frealstockno    = rsget("realstockno")
		end if
		rsget.Close
	end function

        '샵별아이템일별재고목록
	public function GetShopItemDailySummaryList()
		dim sqlStr, i

		sqlStr = " select top 1000 c.shopid, c.itemgubun, c.itemid, c.itemoption"
		sqlStr = sqlStr + " ,c.yyyymmdd, c.sellno, c.resellno, c.logicsipgono, c.logicsreipgono, c.brandipgono, c.brandreipgono, c.sysstockno "
		sqlStr = sqlStr + " ,c.errsampleitemno, c.errbaditemno, c.errrealcheckno, c.realstockno"
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_daily_shopstock_summary c "
		sqlStr = sqlStr + " where 1 = 1 "

		if (FRectStartDate<>"") then
		        sqlStr = sqlStr + " and c.yyyymmdd >= '" + FRectStartDate + "' "
		end if

		if (FRectShopID <> "") then
		        sqlStr = sqlStr + " and c.shopid = '" + FRectShopID + "' "
		end if
		if (FRectItemGubun <> "") then
		        sqlStr = sqlStr + " and c.itemgubun = '" + FRectItemGubun + "' "
		end if
		if (FRectItemId <> "") then
		        sqlStr = sqlStr + " and c.itemid = " + CStr(FRectItemId) + " "
		end if
		if (FRectItemOption <> "") then
		        sqlStr = sqlStr + " and c.itemoption = '" + FRectItemOption + "' "
		end if
		sqlStr = sqlStr + " order by c.yyyymmdd, c.shopid, c.itemgubun, c.itemid, c.itemoption "
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CShopItemSummaryItem

                        FItemList(i).Fyyyymmdd      = rsget("yyyymmdd")

        		        FItemList(i).Fshopid        = rsget("shopid")
        		        FItemList(i).Fitemgubun     = rsget("itemgubun")
        		        FItemList(i).Fitemid        = rsget("itemid")
        		        FItemList(i).Fitemoption    = rsget("itemoption")

        		        FItemList(i).Fsellno         = rsget("sellno")
        		        FItemList(i).Fresellno       = rsget("resellno")
        		        FItemList(i).Flogicsipgono   = rsget("logicsipgono")
        		        FItemList(i).Flogicsreipgono = rsget("logicsreipgono")
        		        FItemList(i).Fbrandipgono    = rsget("brandipgono")
        		        FItemList(i).Fbrandreipgono  = rsget("brandreipgono")

        		        FItemList(i).Fsysstockno    = rsget("sysstockno")

        		        FItemList(i).Ferrsampleitemno= rsget("errsampleitemno")
        		        FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
        		        FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
        		        FItemList(i).Frealstockno    = rsget("realstockno")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage       = 1
		FPageSize       = 100
		FResultCount    = 0
		FScrollCount    = 10
		FTotalCount     = 0
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
