<%

Class CDanpumBaljuOfflineItem
    public Fcompanyid
    public Fcompany_name
    public Fprdcode
    public Fprdname
    public Flocationidmaker
    public Fmainimageurl
    public FDivCD

	public function GetDivCDString()
		if FDivCD="O" then
			'단품
			GetDivCDString = "단품상품"
		elseif FDivCD="E" then
			'제외
			GetDivCDString = "제외상품"
		elseif FDivCD="I" then
			'포함
			GetDivCDString = "포함상품"
		end if
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COrderSheetItem
    public Fbaljucode
	public Fbaljuid
	public Fmakerid
	public Fstatecd
	public Fscheduledate
	public Fprdcode
	public Fbaljuitemno
	public Ftargetid
	public Fbeasongdate
	public Frealitemno
	public Fsongjangdiv
	public Fsongjangno

	public function GetStateName()
		if Fstatecd="0" then
			GetStateName = "주문접수"
		elseif Fstatecd="1" then
			GetStateName = "주문확인"
		elseif Fstatecd="2" then
			GetStateName = "입금대기"
		elseif Fstatecd="5" then
			GetStateName = "배송준비"
		elseif Fstatecd="6" then
			GetStateName = "출고대기"
		elseif Fstatecd="7" then
			GetStateName = "출고완료"
		elseif Fstatecd="8" then
			GetStateName = "검품완료<br>(입고대기)"
		elseif Fstatecd="9" then
			GetStateName = "입고완료"
		end if
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CDanpumBaljuBrandOffline

	public Fcompanyid
    public Flocationidmaker
    public Flocation_name
    public Fcompany_name
    public FDivCD

	public function GetDivCDString()
		if FDivCD="O" then
			'단품
			GetDivCDString = "단품브랜드"
		elseif FDivCD="E" then
			'제외
			GetDivCDString = "제외브랜드"
		elseif FDivCD="I" then
			'포함
			GetDivCDString = "포함브랜드"
		end if
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CBaljuOfflineItem
	public Fmasteridx
	public Fdetailidx
	public Fcompanyid
	public Fordercode
	public FcountryCode
	public Fdivcd
	public Flocationidto
	public Flocationnameto
	public Frequestdt

	public Fprdcode
	public Fitemgubun
	public Fitemid
	public Fitemoption

	public Fmwdiv

	public Fprdname
	public Fbrandid
	public Fbrandname
	public Fitemoptionname
	public Fcustomerprice
	public Frequestedno
	public Frealstockno
	public Fmoveoutdiv5
	public Fmoveoutdiv7
	public Fselldiv2
	public Fselldiv4
	public Fselldiv5
	public Fchulgodiv5

	public Fsellcountday
	public Ftotalsellcountday
	public Fonsellcountbyday
	public Fstockneedday

	public Fdanjongyn
	public FreipgoMayDate
	public Fpreorderno
	public Fpreordernofix
	public Foffjupno

	public function GetMWDivName()

		if ((Fmwdiv = "M") or (Fmwdiv = "W")) then
			GetMWDivName = "텐배"
		elseif (Fmwdiv = "U") then
			GetMWDivName = "업배"
		elseif (Fmwdiv = "O") then
			GetMWDivName = "오프"
		end if

	end function

    'stockneedday 일간 온라인 필요수량
    public function GetOnlineRequireNo()
    	dim result

    	if ((Fsellcountday <> "") and (Ftotalsellcountday <> "")) then

	    	'sellcountday 동안의 판매량
	    	result = (Fonsellcountbyday)

	    	'1 일 판매량(totalsellcountday : 판매일수)
	    	result = result / Ftotalsellcountday

	    	'stockneedday 동안 필요 재고량
	    	result = result * Fstockneedday

	    	'무조건 올림
	    	if (result > Fix(result)) then
	    		GetOnlineRequireNo = Fix(result) + 1
	    	elseif (result < Fix(result)) then
	    		GetOnlineRequireNo = Fix(result) - 1
	    	else
	    		GetOnlineRequireNo = result
	    	end if

		else

			GetOnlineRequireNo = 0

		end if
	end function

	'출고 가능수량
    public function GetChulgoAvailableNo()
    	dim result

		result = Frealstockno														'실사

		result = result + (Fmoveoutdiv5 + Fmoveoutdiv7 + Fselldiv5 + Fchulgodiv5)	'기발주

		result = result + Fselldiv4													'온라인결재

		GetChulgoAvailableNo = result
	end function

    public function GetOffChulgoAvailableNo()
    	dim result

		result = (Frealstockno + (Fselldiv5 + Fchulgodiv5 + Fselldiv4 + Fselldiv2))

		GetOffChulgoAvailableNo = result
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CTenBaljuOffline
	public FItemList()

	public FLastQuery

	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FTotalCount

	public FRectCompanyId
	public FRectLocationid3PL
	public FRectLocationidTo
	public FRectMakerid

	public FRectRegStart
	public FRectNotitemlist
	public FRectItemlist

	public FRectNotIncludeItem
	public FRectIncludeItem

	public FRectIncludeMinus			'마이너스주문
	public FRectIncludeZeroStock		'전제재고없는주문

	public FRectNotIncludebrand
	public FRectIncludebrand

    public FRectUpbeaInclude
    public FRectTenbeaOnly
    public FRectDeliveryArea
    public FRectOnlyManyItem

    public FRectOnlyOneJumun
    public FRectOnlyOneJumunType
    public FRectOnlyOneJumunCount
    public FRectOnlyOneJumunCompare

	public FRectItemDivCD
	public FRectBrandDivCD


    public FRectOnlySagawaDeliverArea

	public FRectdcnt

	public FSubTotalsum
	public FAvgTotalsum

	Public FRectMWDiv

	public FRectItemGubun
	public FRectItemId
	public FRectItemOption

    public Sub GetBaljuItemListNewOffline()

		dim sqlStr,i,tmp

		'======================================================================
		''총 갯수. 총금액
		sqlStr = "select count(m.idx) as cnt, sum(m.totalsellcash) as subtotal , avg(m.totalsellcash) as avgtotal " & vbcrlf
		sqlStr = sqlStr & "from [db_storage].[dbo].tbl_ordersheet_master m " & vbcrlf
		sqlStr = sqlStr & "where 1 = 1 " & vbcrlf
		sqlStr = sqlStr & "	and m.deldt is Null " & vbcrlf
		sqlStr = sqlStr & "	and m.statecd = '0' " & vbcrlf
		sqlStr = sqlStr & "	and m.divcode in ('501','502','503') " & vbcrlf
		sqlStr = sqlStr & "	and m.targetid = '" & FRectLocationid3PL & "' " & vbcrlf

		sqlStr = sqlStr & "	and m.regdate>'" & FRectRegStart & "' " & vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")

			FSubtotalsum = rsget("subtotal")
			FAvgTotalsum = rsget("avgtotal")

			if IsNull(FSubtotalsum) then FSubtotalsum=0
			if IsNull(FAvgTotalsum) then FAvgTotalsum=0
		rsget.Close



		'======================================================================
		'데이타
		sqlStr = " select top " & CStr(FPageSize) & " " & vbcrlf
		sqlStr = sqlStr & " 	d.masteridx " & vbcrlf
		sqlStr = sqlStr & " 	, d.idx as detailidx " & vbcrlf
		sqlStr = sqlStr & " 	, m.baljucode as ordercode " & vbcrlf
		sqlStr = sqlStr & " 	, 'CH' as divcd " & vbcrlf
		sqlStr = sqlStr & " 	, IsNull(u.shopCountryCode, 'KR') as countryCode " & vbcrlf
		sqlStr = sqlStr & " 	, m.baljuid as locationidto " & vbcrlf
		sqlStr = sqlStr & " 	, m.baljuname as locationnameto " & vbcrlf
		sqlStr = sqlStr & " 	, m.scheduledate as requestdt " & vbcrlf
		sqlStr = sqlStr & " 	, [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun,d.itemid,d.itemoption) as prdcode " & vbcrlf
		sqlStr = sqlStr & " 	, d.itemname as prdname " & vbcrlf
		sqlStr = sqlStr & " 	, d.makerid as brandid " & vbcrlf
		sqlStr = sqlStr & " 	, '' as brandname " & vbcrlf
		sqlStr = sqlStr & " 	, d.itemoptionname " & vbcrlf
		sqlStr = sqlStr & " 	, d.sellcash as customerprice " & vbcrlf
		sqlStr = sqlStr & " 	, d.baljuitemno as requestedno " & vbcrlf
		sqlStr = sqlStr & " 	, c.realstock as realstockno " & vbcrlf
		sqlStr = sqlStr & " 	, 0 as moveoutdiv5 " & vbcrlf
		sqlStr = sqlStr & " 	, 0 as moveoutdiv7 " & vbcrlf
		sqlStr = sqlStr & " 	, c.ipkumdiv2 as selldiv2 " & vbcrlf
		sqlStr = sqlStr & " 	, c.ipkumdiv4 as selldiv4 " & vbcrlf
		sqlStr = sqlStr & " 	, c.ipkumdiv5 as selldiv5 " & vbcrlf
		sqlStr = sqlStr & " 	, c.offconfirmno as chulgodiv5 " & vbcrlf
		sqlStr = sqlStr & " 	, 7 as sellcountday " & vbcrlf
		sqlStr = sqlStr & " 	, c.maxsellday as totalsellcountday " & vbcrlf
		sqlStr = sqlStr & " 	, c.sell7days as onsellcountbyday " & vbcrlf
		sqlStr = sqlStr & " 	, 7 as stockneedday " & vbcrlf
		sqlStr = sqlStr & " 	, (case when i.itemid is null then 'O' else i.mwdiv end) as mwdiv " & vbcrlf
		sqlStr = sqlStr & " 	, i.danjongyn, T2.StockReipgoDate as reipgoMayDate, IsNull(c.preorderno,0) as preorderno, IsNull(c.preordernofix,0) as preordernofix " & vbcrlf
		sqlStr = sqlStr & " 	, IsNull(c.offjupno,0) as offjupno " & vbcrlf
		sqlStr = sqlStr & " 	, d.itemgubun,d.itemid,d.itemoption " & vbcrlf

		sqlStr = sqlStr & GetFromWhereOfflineItem

		sqlStr = sqlStr & "order by m.idx, prdcode " & vbcrlf

		''response.write "<!-- " & sqlStr & " -->"
		'response.end

		FLastQuery = sqlStr


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new CBaljuOfflineItem

			FItemList(i).Fmasteridx 		= rsget("masteridx")
			FItemList(i).Fdetailidx 		= rsget("detailidx")
			FItemList(i).Fordercode 		= rsget("ordercode")
			FItemList(i).FcountryCode 		= rsget("countryCode")
			FItemList(i).Fdivcd 			= rsget("divcd")
			FItemList(i).Flocationidto 		= rsget("locationidto")
			FItemList(i).Flocationnameto 	= rsget("locationnameto")
			FItemList(i).Frequestdt 		= rsget("requestdt")
			FItemList(i).Fprdcode 			= rsget("prdcode")
			FItemList(i).Fprdname 			= rsget("prdname")
			FItemList(i).Fbrandid 			= rsget("brandid")
			FItemList(i).Fbrandname 		= rsget("brandname")
			FItemList(i).Fitemoptionname 	= rsget("itemoptionname")
			FItemList(i).Fcustomerprice 	= rsget("customerprice")
			FItemList(i).Frequestedno 		= rsget("requestedno")
			FItemList(i).Frealstockno 		= rsget("realstockno")
			FItemList(i).Fmoveoutdiv5 		= rsget("moveoutdiv5")
			FItemList(i).Fmoveoutdiv7 		= rsget("moveoutdiv7")
			FItemList(i).Fselldiv2 			= rsget("selldiv2")
			FItemList(i).Fselldiv4 			= rsget("selldiv4")
			FItemList(i).Fselldiv5 			= rsget("selldiv5")
			FItemList(i).Fchulgodiv5 		= rsget("chulgodiv5")
			FItemList(i).Fsellcountday 		= rsget("sellcountday")
			FItemList(i).Ftotalsellcountday = rsget("totalsellcountday")
			FItemList(i).Fonsellcountbyday 	= rsget("onsellcountbyday")
			FItemList(i).Fstockneedday 		= rsget("stockneedday")
			FItemList(i).Fmwdiv 			= rsget("mwdiv")

			FItemList(i).Fdanjongyn 		= rsget("danjongyn")
			FItemList(i).FreipgoMayDate 	= rsget("reipgoMayDate")
			FItemList(i).Fpreorderno 		= rsget("preorderno")
			FItemList(i).Fpreordernofix 	= rsget("preordernofix")
			FItemList(i).Foffjupno 			= rsget("offjupno")

			FItemList(i).Fitemgubun 		= rsget("itemgubun")
			FItemList(i).Fitemid 			= rsget("itemid")
			FItemList(i).Fitemoption 		= rsget("itemoption")

			rsget.movenext
			i=i+1
		loop
		rsget.Close

    end Sub

	'제외/포함 상품
    public Sub GetDanpumBaljuItemListOffline()
        dim sqlStr,i

        sqlStr = "select count(b.prdcode) as cnt from [db_threepl].[dbo].tbl_balju_control_item b, [db_threepl].[dbo].tbl_item i"
        sqlStr = sqlStr + " where b.prdcode=i.prdcode"

        if (FRectCompanyId <> "") then
        	'sqlStr = sqlStr + " and b.companyid= '" & FRectCompanyId & "' "
        end if

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

        sqlStr = "select top " + CStr(FPageSize*FCurrpage) + " i.companyid, i.prdcode, i.prdname, i.locationid as locationidmaker, i.mainimageurl, IsNull(b.divcd, 'O') as divcd, c.company_name "

        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_threepl].[dbo].tbl_balju_control_item b "
        sqlStr = sqlStr + " 	LEFT JOIN [db_threepl].[dbo].tbl_item i "
        sqlStr = sqlStr + " 	ON "
        sqlStr = sqlStr + " 		1 = 1 "
        'sqlStr = sqlStr + " 		and b.companyid=i.companyid "
        sqlStr = sqlStr + " 		and b.prdcode=i.prdcode "
        sqlStr = sqlStr + " 	LEFT JOIN [db_threepl].[dbo].tbl_company c "
       sqlStr = sqlStr + " 	ON "
        'sqlStr = sqlStr + " 		b.companyid=c.companyid "

        sqlStr = sqlStr + " where 1 = 1 "

        if (FRectCompanyId <> "") then
        	'sqlStr = sqlStr + " and b.companyid= '" & FRectCompanyId & "' "
        end if

        if (FRectItemDivCD <> "") then
        	sqlStr = sqlStr + " and b.divcd= '" & FRectItemDivCD & "' "
        end if

        sqlStr = sqlStr + " order by b.regdate desc"

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1
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

    			set FItemList(i) = new CDanpumBaljuItem

    			'FItemList(i).Fcompanyid    = rsget("companyid")
    			FItemList(i).Fcompany_name    	= rsget("company_name")
    			FItemList(i).Fprdcode    = rsget("prdcode")
                FItemList(i).Fprdname  = db2html(rsget("prdname"))
                FItemList(i).Flocationidmaker   = rsget("locationidmaker")
                FItemList(i).Fmainimageurl= rsget("mainimageurl")

                FItemList(i).FDivCD     = rsget("divcd")

    			rsget.movenext
    			i=i+1
    		loop
    	end if
		rsget.Close

    end Sub

	'제외/포함 브랜드
    public Sub GetDanpumBaljuBrandListOffline()
        dim sqlStr,i

        sqlStr = " select count(b.locationidmaker) as cnt "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_threepl].[dbo].tbl_balju_control_brand b "
        sqlStr = sqlStr + " 	left join [db_threepl].[dbo].tbl_location l "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        'sqlStr = sqlStr + " 		and b.companyid = l.companyid "
        sqlStr = sqlStr + " 		and b.locationidmaker = l.locationid "

        if (FRectCompanyId <> "") then
        	'sqlStr = sqlStr + " and b.companyid= '" & FRectCompanyId & "' "
        end if

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

        sqlStr = "select top " + CStr(FPageSize*FCurrpage) + " b.locationidmaker, IsNull(b.divcd, 'O') as divcd, l.location_name, c.companyid, c.company_name "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_threepl].[dbo].tbl_balju_control_brand b "
        sqlStr = sqlStr + " 	join [db_threepl].[dbo].tbl_location l "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        'sqlStr = sqlStr + " 		and b.companyid = l.companyid "
        sqlStr = sqlStr + " 		and b.locationidmaker = l.locationid "
        sqlStr = sqlStr + " 	join [db_threepl].[dbo].tbl_company c "
        sqlStr = sqlStr + " 	on "
        'sqlStr = sqlStr + " 		b.companyid=c.companyid "

        if (FRectCompanyId <> "") then
        	'sqlStr = sqlStr + " and b.companyid= '" & FRectCompanyId & "' "
        end if

        if (FRectBrandDivCD <> "") then
        	sqlStr = sqlStr + " and b.divcd= '" & FRectBrandDivCD & "' "
        end if

        sqlStr = sqlStr + " order by b.regdate desc"
        'response.write sqlStr

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1
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

    			set FItemList(i) = new CDanpumBaljuBrand

    			'FItemList(i).Fcompanyid    		= rsget("companyid")
    			FItemList(i).Flocationidmaker   = rsget("locationidmaker")
    			FItemList(i).Flocation_name    	= rsget("location_name")
    			FItemList(i).Fcompany_name    	= rsget("company_name")
    			FItemList(i).FDivCD    			= rsget("DivCD")

    			rsget.movenext
    			i=i+1
    		loop
    	end if
	rsget.Close

    end Sub

    public Sub GetShopOrderList()
        dim sqlStr, i
		dim addSql

		addSql = " from "
		addSql = addSql & " 	[db_storage].[dbo].tbl_ordersheet_master m "
		addSql = addSql & " 	join [db_storage].dbo.tbl_ordersheet_detail d "
		addSql = addSql & " 	on "
		addSql = addSql & " 		m.idx = d.masteridx "
		addSql = addSql & " 	left join db_shop.dbo.tbl_shop_user u "
		addSql = addSql & " 	on "
		addSql = addSql & " 		m.baljuid = u.userid "
		addSql = addSql & " where "
		addSql = addSql & " 	1 = 1 "
		addSql = addSql & " 	and m.deldt is null "
		addSql = addSql & " 	and d.deldt is null "
		addSql = addSql & " 	and d.itemgubun = '" & FRectItemGubun & "' "
		addSql = addSql & " 	and d.itemid = " & FRectItemId & " "
		addSql = addSql & " 	and d.itemoption = '" & FRectItemOption & "' "
		addSql = addSql & " 	and m.statecd = '0' "
		addSql = addSql & " 	and m.divcode in ('501','502','503') "
		addSql = addSql & " 	and m.targetid = '10x10' "
		addSql = addSql & " 	and m.regdate > DateAdd(m, -2, getdate()) "
		addSql = addSql & " 	and (IsNULL(u.shopCountryCode, 'KR') = 'KR') "

		sqlStr = " select count(m.baljucode) as cnt "
		sqlStr = sqlStr & addSql

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrpage) + " m.baljucode, m.baljuid, d.makerid, m.statecd, m.scheduledate, [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun, d.itemid, d.itemoption) AS prdcode, d.baljuitemno "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " order by "
		sqlStr = sqlStr & "		m.idx desc "

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1
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

    			set FItemList(i) = new COrderSheetItem

    			FItemList(i).Fbaljucode    = rsget("baljucode")
				FItemList(i).Fbaljuid    = rsget("baljuid")
				FItemList(i).Fmakerid    = rsget("makerid")
				FItemList(i).Fstatecd    = rsget("statecd")
				FItemList(i).Fscheduledate    = rsget("scheduledate")
				FItemList(i).Fprdcode    = rsget("prdcode")
				FItemList(i).Fbaljuitemno    = rsget("baljuitemno")

    			rsget.movenext
    			i=i+1
    		loop
    	end if
		rsget.Close
    end Sub

    public Sub GetUpcheOrderList()
        dim sqlStr, i
		dim addSql

		addSql = " from "
		addSql = addSql & " 	[db_storage].[dbo].tbl_ordersheet_master m "
		addSql = addSql & " 	join [db_storage].dbo.tbl_ordersheet_detail d "
		addSql = addSql & " 	on "
		addSql = addSql & " 		m.idx = d.masteridx "
		addSql = addSql & " where "
		addSql = addSql & " 	1 = 1 "
		addSql = addSql & " 	and m.deldt is null "
		addSql = addSql & " 	and d.deldt is null "
		addSql = addSql & " 	and d.itemgubun = '" & FRectItemGubun & "' "
		addSql = addSql & " 	and d.itemid = " & FRectItemId & " "
		addSql = addSql & " 	and d.itemoption = '" & FRectItemOption & "' "
		addSql = addSql & " 	and m.statecd >= '0' "
		addSql = addSql & " 	and m.statecd < '9' "
		addSql = addSql & " 	and m.divcode < '300' "
		addSql = addSql & " 	and m.targetid = '" & FRectMakerID & "' "
		addSql = addSql & " 	and m.regdate > DateAdd(m, -2, getdate()) "

		sqlStr = " select count(m.baljucode) as cnt "
		sqlStr = sqlStr & addSql

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close


		sqlStr = " select top " + CStr(FPageSize*FCurrpage) + " m.baljucode, m.targetid, m.statecd, m.scheduledate, m.beasongdate, [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun, d.itemid, d.itemoption) AS prdcode, d.baljuitemno, d.realitemno, m.songjangdiv, m.songjangno "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " order by "
		sqlStr = sqlStr & "		m.idx desc "
		''rw sqlStr

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1
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

    			set FItemList(i) = new COrderSheetItem

    			FItemList(i).Fbaljucode  	= rsget("baljucode")
				FItemList(i).Fstatecd    	= rsget("statecd")
				FItemList(i).Fscheduledate  = rsget("scheduledate")
				FItemList(i).Fprdcode    	= rsget("prdcode")
				FItemList(i).Fbaljuitemno   = rsget("baljuitemno")

				FItemList(i).Ftargetid   = rsget("targetid")
				FItemList(i).Fbeasongdate   = rsget("beasongdate")
				FItemList(i).Frealitemno   = rsget("realitemno")
				FItemList(i).Fsongjangdiv   = rsget("songjangdiv")
				FItemList(i).Fsongjangno   = rsget("songjangno")

    			rsget.movenext
    			i=i+1
    		loop
    	end if
		rsget.Close
    end Sub

	public Function GetFromWhereOfflineItem
		dim tmpsql

        tmpsql = " from  " & vbCrLf
        tmpsql = tmpsql + " 	[db_storage].[dbo].tbl_ordersheet_master m " & vbCrLf
        tmpsql = tmpsql + " 	join [db_storage].dbo.tbl_ordersheet_detail d " & vbCrLf
        tmpsql = tmpsql + " 	on " & vbCrLf
        tmpsql = tmpsql + " 		m.idx = d.masteridx " & vbCrLf
        tmpsql = tmpsql + " 	left join db_summary.dbo.tbl_current_logisstock_summary c " & vbCrLf
        tmpsql = tmpsql + " 	on " & vbCrLf
        tmpsql = tmpsql + " 		1 = 1 " & vbCrLf
        tmpsql = tmpsql + " 		and d.itemgubun = c.itemgubun " & vbCrLf
        tmpsql = tmpsql + " 		and d.itemid = c.itemid " & vbCrLf
        tmpsql = tmpsql + " 		and d.itemoption = c.itemoption " & vbCrLf
        tmpsql = tmpsql + " 	left join db_shop.dbo.tbl_shop_user u " & vbCrLf
        tmpsql = tmpsql + " 	on " & vbCrLf
        tmpsql = tmpsql + " 		m.baljuid = u.userid " & vbCrLf

		tmpsql = tmpsql + " 	left join [db_item].[dbo].tbl_item i "
		tmpsql = tmpsql + " 	on " & vbCrLf
        tmpsql = tmpsql + " 		1 = 1 " & vbCrLf
        tmpsql = tmpsql + " 		and d.itemgubun = '10' " & vbCrLf
        tmpsql = tmpsql + " 		and d.itemid = i.itemid " & vbCrLf
		tmpsql = tmpsql + " 	left join [db_item].[dbo].tbl_item_option_Stock T2 "
		tmpsql = tmpsql + " 	on "
		tmpsql = tmpsql + " 		1 = 1 "
		tmpsql = tmpsql + " 		and T2.itemgubun = '10' "
		tmpsql = tmpsql + " 		and d.itemgubun = T2.itemgubun "
		tmpsql = tmpsql + " 		and d.itemid=T2.itemid "
		tmpsql = tmpsql + " 		and d.itemoption=T2.itemoption "

        tmpsql = tmpsql + " 	left join ( " & vbCrLf
        tmpsql = tmpsql + " 		select " & vbCrLf
        tmpsql = tmpsql + " 			m.idx " & vbCrLf
        tmpsql = tmpsql + " 			, sum(case when d.baljuitemno < 0 then 1 else 0 end) as minusordercnt " & vbCrLf    '''1->0 수정
        tmpsql = tmpsql + " 			, sum(case when IsNull(c.realstock, 0) > 0 then 1 else 0 end) as notzerostockordercnt " & vbCrLf
        tmpsql = tmpsql + " 		from " & vbCrLf
        tmpsql = tmpsql + " 			[db_storage].[dbo].tbl_ordersheet_master m " & vbCrLf
        tmpsql = tmpsql + " 			join [db_storage].dbo.tbl_ordersheet_detail d " & vbCrLf
        tmpsql = tmpsql + " 			on " & vbCrLf
        tmpsql = tmpsql + " 				m.idx = d.masteridx " & vbCrLf
        tmpsql = tmpsql + " 			left join db_summary.dbo.tbl_current_logisstock_summary c " & vbCrLf
        tmpsql = tmpsql + " 			on " & vbCrLf
        tmpsql = tmpsql + " 				1 = 1 " & vbCrLf
        tmpsql = tmpsql + " 				and d.itemgubun = c.itemgubun " & vbCrLf
        tmpsql = tmpsql + " 				and d.itemid = c.itemid " & vbCrLf
        tmpsql = tmpsql + " 				and d.itemoption = c.itemoption " & vbCrLf
        tmpsql = tmpsql + " 		where " & vbCrLf
        tmpsql = tmpsql + " 			1 = 1 " & vbCrLf
        tmpsql = tmpsql + " 			and m.deldt is Null " & vbCrLf
        tmpsql = tmpsql & "				and m.statecd = '0' " & vbcrlf
        tmpsql = tmpsql + " 			and m.divcode in ('501','502','503') " & vbCrLf
        tmpsql = tmpsql + "			 	and m.targetid = '" & FRectLocationid3PL & "' " & vbCrLf
        tmpsql = tmpsql + " 		 	and m.regdate > '" & FRectRegStart & "' " & vbCrLf
        tmpsql = tmpsql + " 		group by " & vbCrLf
        tmpsql = tmpsql + " 			m.idx " & vbCrLf
        tmpsql = tmpsql + " 	) T " & vbCrLf
        tmpsql = tmpsql + " 	on " & vbCrLf
        tmpsql = tmpsql + " 		T.idx = m.idx " & vbCrLf

        tmpsql = tmpsql + " where " & vbCrLf
        tmpsql = tmpsql + " 	1 = 1 " & vbCrLf
        tmpsql = tmpsql + " 	and m.deldt is Null " & vbCrLf
        tmpsql = tmpsql + " 	and d.deldt is Null " & vbCrLf
        tmpsql = tmpsql & "		and m.statecd = '0' " & vbcrlf
        tmpsql = tmpsql + " 	and m.divcode in ('501','502','503') " & vbCrLf
        tmpsql = tmpsql + " 	and m.targetid = '" & FRectLocationid3PL & "' " & vbCrLf
        tmpsql = tmpsql + "  	and m.regdate > '" & FRectRegStart & "' " & vbCrLf

		if (FRectLocationidTo <> "") then
			tmpsql = tmpsql + "  	and m.baljuid = '" & FRectLocationidTo & "' " & vbCrLf
		end if

		if (FRectDeliveryArea <> "") then
			if (FRectDeliveryArea = "ZZ") then
				'군부대배송
				tmpsql = tmpsql & " and (IsNULL(u.shopCountryCode,'KR')='ZZ') "
			elseif (FRectDeliveryArea = "EMS") then
				'해외배송
				tmpsql = tmpsql & " and ((IsNULL(u.shopCountryCode,'KR')<>'KR') and (IsNULL(u.shopCountryCode,'KR')<>'ZZ')) "
			else
				'국내배송
				tmpsql = tmpsql & " and (IsNULL(u.shopCountryCode,'KR')='KR') "
			end if
		end if

		if (FRectIncludeMinus = "N") then
			tmpsql = tmpsql + "  	and T.minusordercnt = 0 " & vbCrLf
		end if

		if (FRectIncludeZeroStock = "N") then
			tmpsql = tmpsql + "  	and T.notzerostockordercnt > 0 " & vbCrLf
		end If

		if (FRectMakerid <> "") then
			tmpsql = tmpsql + "  	and d.makerid = '" & FRectMakerid & "' " & vbCrLf
		end If

		If (FRectMWDiv <> "") Then
			tmpsql = tmpsql + "  	And m.baljucode in ( " & vbCrLf
			tmpsql = tmpsql + "  	select m.baljucode " & vbCrLf
			tmpsql = tmpsql + "  	FROM [db_storage].[dbo].tbl_ordersheet_master m " & vbCrLf
			tmpsql = tmpsql + "  	INNER JOIN [db_storage].dbo.tbl_ordersheet_detail d ON m.idx = d.masteridx " & vbCrLf
			tmpsql = tmpsql + "  	LEFT JOIN [db_item].[dbo].tbl_item i ON 1 = 1 " & vbCrLf
			tmpsql = tmpsql + "  		AND d.itemgubun = '10' " & vbCrLf
			tmpsql = tmpsql + "  		AND d.itemid = i.itemid " & vbCrLf
			tmpsql = tmpsql + "  	where " & vbCrLf
			tmpsql = tmpsql + "  		1 = 1 " & vbCrLf
			tmpsql = tmpsql + "  			AND m.targetid = '10x10' " & vbCrLf
			tmpsql = tmpsql + "  			AND m.regdate > '" & FRectRegStart & "' " & vbCrLf
			tmpsql = tmpsql + "  			AND m.statecd = '0' " & vbCrLf
			tmpsql = tmpsql + "  	group by " & vbCrLf
			tmpsql = tmpsql + "  		m.baljucode " & vbCrLf
			tmpsql = tmpsql + "  	having count(distinct (CASE  WHEN i.itemid IS NULL THEN 'O' WHEN i.mwdiv in ('M', 'W') THEN 'T' ELSE i.mwdiv END)) = 1 " & vbCrLf
			tmpsql = tmpsql + "  	) " & vbCrLf
			tmpsql = tmpsql + "  	And (CASE  WHEN i.itemid IS NULL THEN 'O' WHEN i.mwdiv in ('M', 'W') THEN 'T' ELSE i.mwdiv END) = '" & FRectMWDiv & "' " & vbCrLf
		End If

		GetFromWhereOfflineItem = tmpsql

	end Function

	Private Sub Class_Initialize()
		redim FItemList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
end Class
%>
