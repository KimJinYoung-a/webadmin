<%
'###########################################################
' Description :  cs 메모
' History : 2007.10.26 한용민 수정
'###########################################################

function drawSelectBoxShopDiv(selectBoxName,selectedId,isNotProtoShop)
   dim tmp_str,sqlStr
   %>
     <select name="<%=selectBoxName%>" >
     <option value='' <%if selectedId="" then response.write " selected" %> >선택</option>
   <%
       sqlStr = " select shopdiv,divName from db_shop.dbo.tbl_shopDiv"
       sqlStr = sqlStr &" where 1=1"
       if (isNotProtoShop) then
        sqlStr = sqlStr &" and shopdiv in (1,3,5,7,9,11)"
       end if
       sqlStr = sqlStr &" order by shopdiv"

       rsget.Open sqlStr,dbget,1

       if  not rsget.EOF  then
           do until rsget.EOF
               if LCase(selectedId) = LCase(rsget("shopdiv")) then
                   tmp_str = " selected"
               end if
               response.write("<option value='" & rsget("shopdiv") & "' " & tmp_str & ">" + db2html(rsget("divName")) + " </option>")
               tmp_str = ""
               rsget.MoveNext
           loop
       end if
       rsget.close
   %>
       </select>
   <%
End function

Class COffShopCostItem
    public FShopID
    public FShopName

    public FPrev_Comm_cd
    public FPrev_Comm_name

    public FComm_cd
    public FComm_name

    public FShopDiv
    public FDivName

    public FttlSell
    public FttlBuy
	public FetcBuy
    public FttlChulSum
    public FttlSuplySum
    public FinnerMOrd
    public FinnerFOrd
    public Fmakerid

    public FstockPricePrevMonth
    public FstockPriceThisMonth

    public function getShopGainSum()
        getShopGainSum = FttlSell - getCostPrice
    end function

    public function getShopGainPro()
        dim igainSum : igainSum = getShopGainSum
        if (FttlSell<>0) then
            on Error resume Next
            getShopGainPro = CLNG(igainSum/FttlSell*100*100)/100

            if err then
                getShopGainPro = CLNG(igainSum/FttlSell*100)
            end if
            on Error Goto 0
        end if
    end function

    public function getTurnoverPro()
        dim abgStock
        abgStock = (FstockPricePrevMonth+FstockPriceThisMonth)/2

        if (abgStock<>0) then
            getTurnoverPro = CLng(FttlSell/abgStock*100*100)/100

        else
            getTurnoverPro = 0
        end if
    end function

    public function getTurnoverProByCost()
        dim abgStock
        abgStock = (FstockPricePrevMonth+FstockPriceThisMonth)/2

        if (abgStock<>0) then
            getTurnoverProByCost = CLng(getCostPrice/abgStock*100*100)/100

        else
            getTurnoverProByCost = 0
        end if
    end function



    public function getCostPrice()
        getCostPrice = FstockPricePrevMonth + getCostPriceThisMonth - FstockPriceThisMonth
    end function

    public function getCostPriceThisMonth()
        getCostPriceThisMonth = FttlBuy + FinnerMOrd + FinnerFOrd
    end function

    public function isCheckChulSumDiff()
        if (FComm_cd="B031") and (FttlChulSum<>(FttlBuy+FinnerMOrd+FinnerFOrd)) then
            isCheckChulSumDiff = true
        end if
    end function

    public function getChulSumDiffValue()
        getChulSumDiffValue = FttlChulSum-(FttlBuy+FinnerMOrd+FinnerFOrd)
    end function


    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub

end Class

Class COffShopCostPerMeachulItem

        public Fmakerid
        public Fdefaultmargin
        public Fmwdiv
        public Fmwname

        public Fpredefaultmargin
        public Fpremwdiv
        public Fpremwname

        public Fvatyn
        public Fshopbuysumprevmonth		'// 기초재고
        public Fshopbuysumthismonth		'// 기말재고
        public Fshopmeachul				'// 소비자매출
        public Fshopmeaip				'// 매장매입

        public Fshopminusprevmonth		'// 전달마이너스재고
        public Fshopminusthismonth		'// 당월마이너스재고

        public Fshoperrorthismonth		'// 당월오차
        public Fshopminuserrorthismonth
        public Fshoppluserrorthismonth

	    public function GetDivcdColor()

	        if (Fmwdiv="B031") then
	        	GetDivcdColor = "green"
	        elseif (Fmwdiv="B022") then
	        	GetDivcdColor = "blue"
	        else
	            GetDivcdColor = "#000000"
	        end if

	    end function

	    public function GetPreDivcdColor()

	        if (FPremwdiv="B031") then
	        	GetPreDivcdColor = "green"
	        elseif (FPremwdiv="B022") then
	        	GetPreDivcdColor = "blue"
	        else
	            GetPreDivcdColor = "#000000"
	        end if

	    end function

	    public function GetDivcdName()

	        if (Fdivcd="101") then
	        	GetDivcdName = "출고매입"		'// B031
	        elseif (Fdivcd="OW") then
	        	GetDivcdName = "오프<br>매입"	'// B021
	        elseif (Fdivcd="102") then
	        	GetDivcdName = "매장매입"		'// B022
	        elseif (Fdivcd="103") then
	        	GetDivcdName = "업체위탁"		'// B012
	        elseif (Fdivcd="201") then
	        	GetDivcdName = "온매입"
	        elseif (Fdivcd="301") then
	        	GetDivcdName = "오프매입"
	        else
	            GetDivcdName = Fdivcd
	        end if

	    end function



        Private Sub Class_Initialize()

        End Sub

        Private Sub Class_Terminate()

        End Sub
end Class

Class COffShopMakerMonthlyMaeipItem

		public Fshopid
		public Fipchulcode
		public Fmakerid
		public Fitemgubun
		public Fitemid
		public Fitemoption
		public Fitemname
		public Fitemoptionname
		public Fbuycash
		public Fsuplycash
		public Fitemno
		public Fvatyn

	    public function GetBarcode()

			if (Fitemid>=1000000) then
                GetBarcode = Fitemgubun & Right((100000000 + Fitemid), 8) & Fitemoption
            else
			    GetBarcode = Fitemgubun & Right((1000000 + Fitemid), 6) & Fitemoption
		    end if

	    end function

        Private Sub Class_Initialize()

        End Sub

        Private Sub Class_Terminate()

        End Sub
end Class




Class COffShopCostPerMeachul
        public FItemList()
        public FOneItem

        public FCurrPage
        public FTotalPage
        public FPageSize
        public FResultCount
        public FScrollCount
        public FTotalCount

		public FRectShopID
		public FRectMakerID
		public FRectYYYYMM
		public FRectGubun
		public FRectJungsanGubun
		public FRectMWDiv

        public FRectShopDiv

        public FRectMinusStockExclude

        public Sub GetOffShopCostSumByShopNew()
			dim i,sqlStr
			'response.write "'" & FRectShopID & "', '" & FRectShopDiv & "', '" & FRectYYYYMM & "', '" & FRectJungsanGubun & "', '" & FRectMakerID & "', '" & FRectMinusStockExclude & "'"
	        sqlStr = " exec [db_summary].[dbo].[usp_Ten_Shop_CostPerMeachulByShop] '" & FRectShopID & "', '" & FRectShopDiv & "', '" & FRectYYYYMM & "', '" & FRectJungsanGubun & "', '" & FRectMakerID & "', '" & FRectMinusStockExclude & "' "

	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
            rsget.Open sqlStr, dbget, 1
            FResultCount = rsget.RecordCount
            FTotalCount = FResultCount

            if FResultCount<1 then FResultCount=0

            redim preserve FItemList(FResultCount)

    		if  not rsget.EOF  then
    		    i = 0
    		    do until rsget.eof
        			set FItemList(i) = new COffShopCostItem
                    FItemList(i).FShopID      = rsget("shopid")
                    FItemList(i).FShopName    = rsget("shopname")
                    'FItemList(i).FComm_cd     = rsget("Comm_cd")
                    'FItemList(i).FComm_name   = rsget("Comm_name")
                    FItemList(i).FShopDiv     = rsget("ShopDiv")
                    FItemList(i).FDivName     = rsget("DivName")
                    FItemList(i).FttlSell     = rsget("ttlSell")
                    FItemList(i).FttlBuy      = rsget("ttlBuy")
					FItemList(i).FetcBuy      = rsget("etcBuy")
                    FItemList(i).FttlChulSum  = rsget("ttlChulSum")
                    FItemList(i).FttlSuplySum = rsget("ttlSuplySum")
                    FItemList(i).FinnerMOrd   = rsget("innerMOrd")
                    FItemList(i).FinnerFOrd   = rsget("innerFOrd")

                    FItemList(i).FstockPricePrevMonth   = rsget("stockPricePrevMonth")
                    FItemList(i).FstockPriceThisMonth   = rsget("stockPriceThisMonth")

        			rsget.MoveNext
        			i = i + 1
        		loop
    	    end if

            rsget.Close
        end Sub

        public Sub GetOffMainCostSumByShop()
			dim i,sqlStr

	        sqlStr = " exec [db_summary].[dbo].[usp_Ten_OffMain_CostPerMeachulByShop] '" & FRectShopID & "', '" & FRectShopDiv & "', '" & FRectYYYYMM & "', '" & FRectMakerID & "', '" & FRectMinusStockExclude & "' "

	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
            rsget.Open sqlStr, dbget, 1
            FResultCount = rsget.RecordCount
            FTotalCount = FResultCount

            if FResultCount<1 then FResultCount=0

            redim preserve FItemList(FResultCount)

    		if  not rsget.EOF  then
    		    i = 0
    		    do until rsget.eof
        			set FItemList(i) = new COffShopCostItem
                    FItemList(i).FShopID      = rsget("shopid")
                    FItemList(i).FShopName    = rsget("shopname")
                    'FItemList(i).FComm_cd     = rsget("Comm_cd")
                    'FItemList(i).FComm_name   = rsget("Comm_name")
                    FItemList(i).FShopDiv     = rsget("ShopDiv")
                    FItemList(i).FDivName     = rsget("DivName")
                    FItemList(i).FttlSell     = rsget("ttlSell")
                    FItemList(i).FttlBuy      = rsget("ttlBuy")
					FItemList(i).FetcBuy      = rsget("etcBuy")
                    FItemList(i).FttlChulSum  = rsget("ttlChulSum")
                    FItemList(i).FttlSuplySum = rsget("ttlSuplySum")
                    FItemList(i).FinnerMOrd   = rsget("innerMOrd")
                    FItemList(i).FinnerFOrd   = rsget("innerFOrd")

                    FItemList(i).FstockPricePrevMonth   = rsget("stockPricePrevMonth")
                    FItemList(i).FstockPriceThisMonth   = rsget("stockPriceThisMonth")

        			rsget.MoveNext
        			i = i + 1
        		loop
    	    end if

            rsget.Close
        end Sub

        public Sub GetOffShopCostSumByShop()
            dim i,sqlStr

            sqlStr = " select T.*, S.shopname, S.shopDiv, D.DivName"
            sqlStr = sqlStr & " from ("
            sqlStr = sqlStr & " 	select shopid,sum(CASE WHEN sumgbn='S' THEN tSellSum ELSE 0 END) as ttlSell"
            sqlStr = sqlStr & " 	,sum(CASE WHEN sumgbn='J' THEN tBuySum ELSE 0 END) as ttlBuy"
            sqlStr = sqlStr & " 	,sum(CASE WHEN sumgbn='C' THEN tBuySum ELSE 0 END) as ttlChulSum"
            sqlStr = sqlStr & " 	,sum(CASE WHEN sumgbn='C' THEN tSuplySum ELSE 0 END) as ttlSuplySum"
            sqlStr = sqlStr & " 	,sum(CASE WHEN sumgbn='M' THEN tBuySum ELSE 0 END) as innerMOrd"
            sqlStr = sqlStr & " 	,sum(CASE WHEN sumgbn='F' THEN tBuySum ELSE 0 END) as innerFOrd"
            sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shop_brand_monthly_Sum A"
            sqlStr = sqlStr & " 	where A.yyyymm='"&FRectYYYYMM&"'"
            if (FRectShopID<>"") then
                sqlStr = sqlStr & " 	and A.shopid='"&FRectShopID&"'"
            end if
            sqlStr = sqlStr & " 	group by A.shopid"
            sqlStr = sqlStr & " ) T"
            sqlStr = sqlStr & "     left join db_shop.dbo.tbl_shop_user S"
            sqlStr = sqlStr & "     on T.shopid=S.userid"
            sqlStr = sqlStr & "     left join db_shop.dbo.tbl_shopDiv D"
            sqlStr = sqlStr & "     on S.shopdiv=D.shopdiv"
            sqlStr = sqlStr & "     where 1=1"
            if (FRectShopDiv<>"") then
                sqlStr = sqlStr & "     and ((S.shopdiv='"&FRectShopDiv&"') or (T.shopid=''))"
            end if
            sqlStr = sqlStr & " order by S.shopdiv, T.shopid"
            ''rw sqlStr

            rsget.Open sqlStr, dbget, 1
            FResultCount = rsget.RecordCount
            FTotalCount = FResultCount

            if FResultCount<1 then FResultCount=0

            redim preserve FItemList(FResultCount)

    		if  not rsget.EOF  then
    		    i = 0
    		    do until rsget.eof
        			set FItemList(i) = new COffShopCostItem
                    FItemList(i).FShopID      = rsget("shopid")
                    FItemList(i).FShopName    = rsget("shopname")
                    'FItemList(i).FComm_cd     = rsget("Comm_cd")
                    'FItemList(i).FComm_name   = rsget("Comm_name")
                    FItemList(i).FShopDiv     = rsget("ShopDiv")
                    FItemList(i).FDivName     = rsget("DivName")
                    FItemList(i).FttlSell     = rsget("ttlSell")
                    FItemList(i).FttlBuy      = rsget("ttlBuy")
                    FItemList(i).FttlChulSum  = rsget("ttlChulSum")
                    FItemList(i).FttlSuplySum = rsget("ttlSuplySum")
                    FItemList(i).FinnerMOrd   = rsget("innerMOrd")
                    FItemList(i).FinnerFOrd   = rsget("innerFOrd")

        			rsget.MoveNext
        			i = i + 1
        		loop
    	    end if

            rsget.Close
        end Sub

        public Sub GetOffShopCostSumByJungsanNew()
            dim i,sqlStr

	        sqlStr = " exec [db_summary].[dbo].[usp_Ten_Shop_CostPerMeachulByJungsan] '" & FRectShopID & "', '" & FRectShopDiv & "', '" & FRectYYYYMM & "', '" & FRectJungsanGubun & "', '" & FRectMakerID & "', '" & FRectMinusStockExclude & "' "
			''rw sqlStr

	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
            rsget.Open sqlStr, dbget, 1
            FResultCount = rsget.RecordCount
            FTotalCount = FResultCount

            if FResultCount<1 then FResultCount=0

            redim preserve FItemList(FResultCount)

    		if  not rsget.EOF  then
    		    i = 0
    		    do until rsget.eof
        			set FItemList(i) = new COffShopCostItem
                    ''FItemList(i).FShopID      = rsget("shopid")
                    ''FItemList(i).FShopName    = rsget("shopname")
                    FItemList(i).FComm_cd     = rsget("Comm_cd")
                    FItemList(i).FComm_name   = rsget("Comm_name")
                    ''FItemList(i).FShopDiv     = rsget("ShopDiv")
                    ''FItemList(i).FDivName     = rsget("DivName")
                    FItemList(i).FttlSell     = rsget("ttlSell")
                    FItemList(i).FttlBuy      = rsget("ttlBuy")
					FItemList(i).FetcBuy      = rsget("etcBuy")
                    FItemList(i).FttlChulSum  = rsget("ttlChulSum")
                    FItemList(i).FttlSuplySum = rsget("ttlSuplySum")
                    FItemList(i).FinnerMOrd   = rsget("innerMOrd")
                    FItemList(i).FinnerFOrd   = rsget("innerFOrd")

                    FItemList(i).FstockPricePrevMonth   = rsget("stockPricePrevMonth")
                    FItemList(i).FstockPriceThisMonth   = rsget("stockPriceThisMonth")

        			rsget.MoveNext
        			i = i + 1
        		loop
    	    end if

            rsget.Close
        end sub

        public Sub GetOffMainCostSumByJungsan()
            dim i,sqlStr

	        sqlStr = " exec [db_summary].[dbo].[usp_Ten_OffMain_CostPerMeachulByJungsan] '" & FRectShopID & "', '" & FRectShopDiv & "', '" & FRectYYYYMM & "', '" & FRectJungsanGubun & "', '" & FRectMakerID & "', '" & FRectMinusStockExclude & "' "
			''rw sqlStr

	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
            rsget.Open sqlStr, dbget, 1
            FResultCount = rsget.RecordCount
            FTotalCount = FResultCount

            if FResultCount<1 then FResultCount=0

            redim preserve FItemList(FResultCount)

    		if  not rsget.EOF  then
    		    i = 0
    		    do until rsget.eof
        			set FItemList(i) = new COffShopCostItem
                    ''FItemList(i).FShopID      = rsget("shopid")
                    ''FItemList(i).FShopName    = rsget("shopname")
                    FItemList(i).FComm_cd     = rsget("Comm_cd")
                    FItemList(i).FComm_name   = rsget("Comm_name")
                    ''FItemList(i).FShopDiv     = rsget("ShopDiv")
                    ''FItemList(i).FDivName     = rsget("DivName")
                    FItemList(i).FttlSell     = rsget("ttlSell")
                    FItemList(i).FttlBuy      = rsget("ttlBuy")
					FItemList(i).FetcBuy      = rsget("etcBuy")
                    FItemList(i).FttlChulSum  = rsget("ttlChulSum")
                    FItemList(i).FttlSuplySum = rsget("ttlSuplySum")
                    FItemList(i).FinnerMOrd   = rsget("innerMOrd")
                    FItemList(i).FinnerFOrd   = rsget("innerFOrd")

                    FItemList(i).FstockPricePrevMonth   = rsget("stockPricePrevMonth")
                    FItemList(i).FstockPriceThisMonth   = rsget("stockPriceThisMonth")

        			rsget.MoveNext
        			i = i + 1
        		loop
    	    end if

            rsget.Close
        end sub

        public Sub GetOffShopCostSumByJungsan()
            dim i,sqlStr

            sqlStr = " select A.comm_cd, sum(CASE WHEN sumgbn='S' THEN tSellSum ELSE 0 END) as ttlSell"
            sqlStr = sqlStr & " ,sum(CASE WHEN sumgbn='J' THEN tBuySum ELSE 0 END) as ttlBuy"
            sqlStr = sqlStr & " ,sum(CASE WHEN sumgbn='C' THEN tBuySum ELSE 0 END) as ttlChulSum"
            sqlStr = sqlStr & " ,sum(CASE WHEN sumgbn='C' THEN tSuplySum ELSE 0 END) as ttlSuplySum"
            sqlStr = sqlStr & " ,sum(CASE WHEN sumgbn='M' THEN tBuySum ELSE 0 END) as innerMOrd"
            sqlStr = sqlStr & " ,sum(CASE WHEN sumgbn='F' THEN tBuySum ELSE 0 END) as innerFOrd"
            sqlStr = sqlStr & " ,IsNULL(C.Comm_name,'미지정') as Comm_name"
            sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_brand_monthly_Sum A"
            sqlStr = sqlStr & " 	left join db_shop.dbo.tbl_shop_user S"
            sqlStr = sqlStr & " 	on A.shopid=S.userid"
            sqlStr = sqlStr & " 	left join db_shop.dbo.tbl_shopDiv D"
            sqlStr = sqlStr & " 	on S.shopdiv=D.shopdiv"
            sqlStr = sqlStr & " 	left join db_jungsan.dbo.tbl_jungsan_comm_code C"
            sqlStr = sqlStr & " 	on A.comm_cd=C.comm_cd"
            sqlStr = sqlStr & " where A.yyyymm='"&FRectYYYYMM&"'"
            if (FRectShopID<>"") then
                sqlStr = sqlStr & " 	and A.shopid='"&FRectShopID&"'"
            end if
            if (FRectShopDiv<>"") then
                sqlStr = sqlStr & "     and ((S.shopdiv='"&FRectShopDiv&"') or (A.shopid=''))"
            end if
            sqlStr = sqlStr & " group by A.comm_cd, IsNULL(C.Comm_name,'미지정')"
            sqlStr = sqlStr & " order by A.comm_cd"

            rsget.Open sqlStr, dbget, 1
            FResultCount = rsget.RecordCount
            FTotalCount = FResultCount

            if FResultCount<1 then FResultCount=0

            redim preserve FItemList(FResultCount)

    		if  not rsget.EOF  then
    		    i = 0
    		    do until rsget.eof
        			set FItemList(i) = new COffShopCostItem
                    ''FItemList(i).FShopID      = rsget("shopid")
                    ''FItemList(i).FShopName    = rsget("shopname")
                    FItemList(i).FComm_cd     = rsget("Comm_cd")
                    FItemList(i).FComm_name   = rsget("Comm_name")
                    ''FItemList(i).FShopDiv     = rsget("ShopDiv")
                    ''FItemList(i).FDivName     = rsget("DivName")
                    FItemList(i).FttlSell     = rsget("ttlSell")
                    FItemList(i).FttlBuy      = rsget("ttlBuy")
                    FItemList(i).FttlChulSum  = rsget("ttlChulSum")
                    FItemList(i).FttlSuplySum = rsget("ttlSuplySum")
                    FItemList(i).FinnerMOrd   = rsget("innerMOrd")
                    FItemList(i).FinnerFOrd   = rsget("innerFOrd")

        			rsget.MoveNext
        			i = i + 1
        		loop
    	    end if

            rsget.Close
        end sub

		public Sub GetOffShopCostSumDetailNew()

	        dim sqlStr
	        sqlStr = " exec [db_summary].[dbo].[usp_Ten_Shop_CostPerMeachulNew] '" & FRectShopID & "', '" & FRectYYYYMM & "', '" & FRectJungsanGubun & "', '" & FRectMakerID & "', '" & FRectMinusStockExclude & "' "

	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
            rsget.Open sqlStr, dbget, 1
            FResultCount = rsget.RecordCount
            FTotalCount = FResultCount

            if FResultCount<1 then FResultCount=0

            redim preserve FItemList(FResultCount)

    		if  not rsget.EOF  then
    		    i = 0
    		    do until rsget.eof
        			set FItemList(i) = new COffShopCostItem
                    FItemList(i).FShopID      = rsget("shopid")
                    FItemList(i).FShopName    = rsget("shopname")
                    FItemList(i).FPrev_Comm_cd     = rsget("Prev_Comm_cd")
                    FItemList(i).FPrev_Comm_name   = rsget("Prev_Comm_name")
                    FItemList(i).FComm_cd     = rsget("Comm_cd")
                    FItemList(i).FComm_name   = rsget("Comm_name")
                    FItemList(i).Fmakerid     = rsget("makerid")
                    FItemList(i).FttlSell     = rsget("ttlSell")
                    FItemList(i).FttlBuy      = rsget("ttlBuy")
					FItemList(i).FetcBuy      = rsget("etcBuy")
                    FItemList(i).FttlChulSum  = rsget("ttlChulSum")
                    FItemList(i).FttlSuplySum = rsget("ttlSuplySum")
                    FItemList(i).FinnerMOrd   = rsget("innerMOrd")
                    FItemList(i).FinnerFOrd   = rsget("innerFOrd")

                    FItemList(i).FstockPricePrevMonth   = rsget("stockPricePrevMonth")
                    FItemList(i).FstockPriceThisMonth   = rsget("stockPriceThisMonth")

        			rsget.MoveNext
        			i = i + 1
        		loop
    	    end if

            rsget.Close

		end sub

		public Sub GetOffMainCostSumDetail()

	        dim sqlStr
			sqlStr = " exec [db_summary].[dbo].[usp_Ten_OffMain_CostPerMeachulByBrand] '" & FRectShopID & "', '" & FRectShopDiv & "', '" & FRectYYYYMM & "', '" & FRectMakerID & "', '" & FRectMinusStockExclude & "' "

	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
            rsget.Open sqlStr, dbget, 1
            FResultCount = rsget.RecordCount
            FTotalCount = FResultCount

            if FResultCount<1 then FResultCount=0

            redim preserve FItemList(FResultCount)

    		if  not rsget.EOF  then
    		    i = 0
    		    do until rsget.eof
        			set FItemList(i) = new COffShopCostItem
                    FItemList(i).FShopID      = rsget("shopid")
                    FItemList(i).FShopName    = rsget("shopname")
                    FItemList(i).FShopDiv     = rsget("ShopDiv")
                    FItemList(i).FDivName     = rsget("DivName")
                    ''FItemList(i).FPrev_Comm_cd     = rsget("Prev_Comm_cd")
                    ''FItemList(i).FPrev_Comm_name   = rsget("Prev_Comm_name")
                    ''FItemList(i).FComm_cd     = rsget("Comm_cd")
                    ''FItemList(i).FComm_name   = rsget("Comm_name")
                    FItemList(i).Fmakerid     = rsget("makerid")
                    FItemList(i).FttlSell     = rsget("ttlSell")
                    FItemList(i).FttlBuy      = rsget("ttlBuy")
					FItemList(i).FetcBuy      = rsget("etcBuy")
                    FItemList(i).FttlChulSum  = rsget("ttlChulSum")
                    FItemList(i).FttlSuplySum = rsget("ttlSuplySum")
                    FItemList(i).FinnerMOrd   = rsget("innerMOrd")
                    FItemList(i).FinnerFOrd   = rsget("innerFOrd")

                    FItemList(i).FstockPricePrevMonth   = rsget("stockPricePrevMonth")
                    FItemList(i).FstockPriceThisMonth   = rsget("stockPriceThisMonth")

        			rsget.MoveNext
        			i = i + 1
        		loop
    	    end if

            rsget.Close

		end sub

        public Sub GetOffShopCostSumDetail()
            dim i,sqlStr
            dim PrevYYYYMM

            if (FRectYYYYMM <> "") then
            	PrevYYYYMM = FRectYYYYMM + "-01"
            	PrevYYYYMM = Left(DateAdd("m", -1, PrevYYYYMM), 7)
            end if


            sqlStr = " select A.shopid, S.shopName, A.makerid, A.comm_cd"
            sqlStr = sqlStr & " ,sum(CASE WHEN sumgbn='S' THEN tSellSum ELSE 0 END) as ttlSell"
            sqlStr = sqlStr & " ,sum(CASE WHEN sumgbn='J' THEN tBuySum ELSE 0 END) as ttlBuy"
            sqlStr = sqlStr & " ,sum(CASE WHEN sumgbn='C' THEN tBuySum ELSE 0 END) as ttlChulSum"
            sqlStr = sqlStr & " ,sum(CASE WHEN sumgbn='C' THEN tSuplySum ELSE 0 END) as ttlSuplySum"
            sqlStr = sqlStr & " ,sum(CASE WHEN sumgbn='M' THEN tBuySum ELSE 0 END) as innerMOrd"
            sqlStr = sqlStr & " ,sum(CASE WHEN sumgbn='F' THEN tBuySum ELSE 0 END) as innerFOrd"
            sqlStr = sqlStr & " ,IsNULL(C.Comm_name,'미지정') as Comm_name"
            sqlStr = sqlStr & " ,IsNULL(T.psum,0) as stockPricePrevMonth"
            sqlStr = sqlStr & " ,IsNULL(T.csum,0) as stockPriceThisMonth"
            sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_brand_monthly_Sum A"
            sqlStr = sqlStr & " 	left join db_shop.dbo.tbl_shop_user S"
            sqlStr = sqlStr & " 	on A.shopid=S.userid"
            sqlStr = sqlStr & " 	left join db_shop.dbo.tbl_shopDiv D"
            sqlStr = sqlStr & " 	on S.shopdiv=D.shopdiv"
            sqlStr = sqlStr & " 	left join db_jungsan.dbo.tbl_jungsan_comm_code C"
            sqlStr = sqlStr & " 	on A.comm_cd=C.comm_cd"
            sqlStr = sqlStr & " 	left join ( "
            sqlStr = sqlStr & " 		select IsNull(p.shopid,c.shopid) as shopid, i.makerid, IsNull(IsNull(cd.comm_cd, pd.comm_cd), 'B000') as comm_cd "
            sqlStr = sqlStr & " 			, sum( "
	        sqlStr = sqlStr & "     				(CASE "
			sqlStr = sqlStr & "             			WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(pd.defaultmargin,35))/100*i.shopitemprice)  "
			sqlStr = sqlStr & "             			ELSE i.shopsuplycash "
			sqlStr = sqlStr & "             		END) * IsNull(p.sysstockno, 0) "
			sqlStr = sqlStr & "             ) as psum "
            sqlStr = sqlStr & " 			, sum((CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(cd.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END) * IsNull(c.sysstockno, 0)) as csum "
            sqlStr = sqlStr & " 		from "
            sqlStr = sqlStr & " 			[db_shop].[dbo].tbl_shop_item i "
            sqlStr = sqlStr & " 			left JOIN db_summary.dbo.tbl_monthly_accumulated_shopstock_summary p "
            sqlStr = sqlStr & " 			on "
            sqlStr = sqlStr & " 				1 = 1 "
            sqlStr = sqlStr & " 				and p.yyyymm = '"&PrevYYYYMM&"' "
            sqlStr = sqlStr & " 				and p.shopid = '"&FRectShopID&"' "
            sqlStr = sqlStr & " 				and p.itemgubun = i.itemgubun "
            sqlStr = sqlStr & " 				and p.itemid = i.shopitemid "
            sqlStr = sqlStr & " 				and p.itemoption = i.itemoption "

            if (FRectMinusStockExclude = "Y") then
            	sqlStr = sqlStr & " 				and p.sysstockno > 0 "
            end if

            sqlStr = sqlStr & " 			left Join db_summary.dbo.tbl_monthly_shop_designer pd "
            sqlStr = sqlStr & " 				on p.shopid=pd.shopid "
            sqlStr = sqlStr & " 				and i.makerid=pd.makerid "
            sqlStr = sqlStr & " 				and p.yyyymm=pd.yyyymm "
            sqlStr = sqlStr & " 			left JOIN db_summary.dbo.tbl_monthly_accumulated_shopstock_summary c "
            sqlStr = sqlStr & " 			on "
            sqlStr = sqlStr & " 				1 = 1 "
            sqlStr = sqlStr & " 				and c.yyyymm = '"&FRectYYYYMM&"' "
            sqlStr = sqlStr & " 				and c.shopid = '"&FRectShopID&"' "
            sqlStr = sqlStr & " 				and c.itemgubun = i.itemgubun "
            sqlStr = sqlStr & " 				and c.itemid = i.shopitemid "
            sqlStr = sqlStr & " 				and c.itemoption = i.itemoption "

            if (FRectMinusStockExclude = "Y") then
            	sqlStr = sqlStr & " 				and c.sysstockno > 0 "
            end if

            sqlStr = sqlStr & " 			left Join db_summary.dbo.tbl_monthly_shop_designer cd "
            sqlStr = sqlStr & " 				on c.shopid=cd.shopid "
            sqlStr = sqlStr & " 				and i.makerid=cd.makerid "
            sqlStr = sqlStr & " 				and c.yyyymm=cd.yyyymm "
            sqlStr = sqlStr & " 		where "
            sqlStr = sqlStr & " 			1 = 1 "
            sqlStr = sqlStr & " 			and (p.shopid is not null or c.shopid is not null) "
            sqlStr = sqlStr & " 			and IsNull(IsNull(cd.comm_cd, pd.comm_cd), 'B000') not in ('B011', 'B012', 'B013') "
            sqlStr = sqlStr & " 		group by "
            sqlStr = sqlStr & " 			IsNull(p.shopid,c.shopid), i.makerid, IsNull(IsNull(cd.comm_cd, pd.comm_cd), 'B000') "
            sqlStr = sqlStr & " 	) T "
            sqlStr = sqlStr & " 	on A.shopid=T.shopid and A.makerid=T.makerid and A.comm_cd=T.comm_cd "

            sqlStr = sqlStr & " where A.yyyymm='"&FRectYYYYMM&"'"

            if (FRectShopID<>"") then
                sqlStr = sqlStr & " 	and A.shopid='"&FRectShopID&"'"
            end if

            if (FRectMakerID<>"") then
                sqlStr = sqlStr & " 	and A.makerid='"&FRectMakerID&"'"
            end if

            if (FRectShopDiv<>"") then
                sqlStr = sqlStr & "     and ((S.shopdiv='"&FRectShopDiv&"') or (A.shopid=''))"
            end if
            if (FRectJungsanGubun<>"") then
                sqlStr = sqlStr & " and A.comm_cd='"&FRectJungsanGubun&"'"
            end if
            sqlStr = sqlStr & " and (tSellSum<>0 or tBuySum<>0 or tSuplySum<>0)"
            sqlStr = sqlStr & " group by A.shopid, S.shopName, A.makerid, A.comm_cd, IsNULL(C.Comm_name,'미지정'),IsNULL(T.psum,0),IsNULL(T.csum,0)"
            sqlStr = sqlStr & " order by A.comm_cd, ttlSell desc"
			''response.write sqlStr

            rsget.Open sqlStr, dbget, 1
            FResultCount = rsget.RecordCount
            FTotalCount = FResultCount

            if FResultCount<1 then FResultCount=0

            redim preserve FItemList(FResultCount)

    		if  not rsget.EOF  then
    		    i = 0
    		    do until rsget.eof
        			set FItemList(i) = new COffShopCostItem
                    FItemList(i).FShopID      = rsget("shopid")
                    FItemList(i).FShopName    = rsget("shopname")
                    FItemList(i).FComm_cd     = rsget("Comm_cd")
                    FItemList(i).FComm_name   = rsget("Comm_name")
                    FItemList(i).Fmakerid     = rsget("makerid")
                    FItemList(i).FttlSell     = rsget("ttlSell")
                    FItemList(i).FttlBuy      = rsget("ttlBuy")
                    FItemList(i).FttlChulSum  = rsget("ttlChulSum")
                    FItemList(i).FttlSuplySum = rsget("ttlSuplySum")
                    FItemList(i).FinnerMOrd   = rsget("innerMOrd")
                    FItemList(i).FinnerFOrd   = rsget("innerFOrd")

                    FItemList(i).FstockPricePrevMonth   = rsget("stockPricePrevMonth")
                    FItemList(i).FstockPriceThisMonth   = rsget("stockPriceThisMonth")

        			rsget.MoveNext
        			i = i + 1
        		loop
    	    end if

            rsget.Close

        end sub

        public Sub GetOffShopCostPerMeachulByBrandList()
            dim i,sqlStr
            dim ArrList

			'// ===============================================================
			IF (FRectGubun="real") then
			    sqlStr = "[db_summary].[dbo].usp_Ten_Shop_CostPerMeachulReal('" + CStr(FRectShopID) + "', '" + CStr(FRectYYYYMM) + "', '" + CStr(FRectMWDiv) + "')"
			ELSE
			    sqlStr = "[db_summary].[dbo].usp_Ten_Shop_CostPerMeachulSYS('" + CStr(FRectShopID) + "', '" + CStr(FRectYYYYMM) + "', '" + CStr(FRectMWDiv) + "')"
		    END IF
	        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
	        'response.write sqlStr
			'response.end

			FTotalCount = 0
			FResultCount = 0
    		IF Not (rsget.EOF OR rsget.BOF) THEN
    			ArrList = rsget.getRows()
    			FResultCount = UBound(ArrList,2)+1
    			FTotalCount = FResultCount
    		END IF
    		rsget.close

	        FTotalPage =  CLng(FTotalCount\FPageSize)
			if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
				FTotalPage = FtotalPage + 1
			end if

	        if FResultCount<1 then FResultCount=0

			redim preserve FItemList(FResultCount)
			i=0
			If IsArray(ArrList) then
				For i=0 to FResultCount-1
					set FItemList(i) = new COffShopCostPerMeachulItem

					FItemList(i).Fmakerid         		= ArrList(0,i)
					FItemList(i).Fmwdiv         		= ArrList(1,i)
					FItemList(i).Fvatyn         		= ArrList(2,i)
					FItemList(i).Fshopbuysumprevmonth   = ArrList(3,i)
					FItemList(i).Fshopbuysumthismonth   = ArrList(4,i)
					FItemList(i).Fshopmeachul         	= ArrList(5,i)
					FItemList(i).Fshopmeaip         	= ArrList(6,i)
					FItemList(i).Fmwname         		= ArrList(7,i)

					FItemList(i).Fshoperrorthismonth	= ArrList(8,i)
					FItemList(i).Fshopminusthismonth	= ArrList(9,i)

					FItemList(i).Fshopminuserrorthismonth	= ArrList(10,i)
					FItemList(i).Fshoppluserrorthismonth	= ArrList(11,i)
					FItemList(i).Fdefaultmargin				= ArrList(12,i)
					FItemList(i).Fshopminusprevmonth		= ArrList(13,i)

					FItemList(i).Fpredefaultmargin      = ArrList(14,i)
					FItemList(i).Fpremwdiv         		= ArrList(15,i)
                    FItemList(i).Fpremwname             = ArrList(16,i)
				next
			end if

        end sub

        public Sub GetOffShopMakerMonthlyMaeip()
            dim i,sqlStr
            dim ArrList

			'// ===============================================================
			sqlStr = "[db_partner].[dbo].usp_Ten_Shop_MakerMonthlyMaeip('" + CStr(FRectShopID) + "', '" + CStr(FRectYYYYMM) + "', '" + CStr(FRectMakerID) + "', '" + CStr(FRectJungsanGubun) + "')"
	        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
	        'response.write sqlStr
			'response.end

			FTotalCount = 0
			FResultCount = 0
    		IF Not (rsget.EOF OR rsget.BOF) THEN
    			ArrList = rsget.getRows()
    			FResultCount = UBound(ArrList,2)+1
    			FTotalCount = FResultCount
    		END IF
    		rsget.close

	        FTotalPage =  CLng(FTotalCount\FPageSize)
			if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
				FTotalPage = FtotalPage + 1
			end if

	        if FResultCount<1 then FResultCount=0

			redim preserve FItemList(FResultCount)
			i=0
			If IsArray(ArrList) then
				For i=0 to FResultCount-1
					set FItemList(i) = new COffShopMakerMonthlyMaeipItem

					FItemList(i).Fshopid         		= ArrList(0,i)
					FItemList(i).Fipchulcode         	= ArrList(1,i)
					FItemList(i).Fmakerid         		= ArrList(2,i)
					FItemList(i).Fitemgubun   			= ArrList(3,i)
					FItemList(i).Fitemid   				= ArrList(4,i)
					FItemList(i).Fitemoption         	= ArrList(5,i)
					FItemList(i).Fitemname	         	= db2html(ArrList(6,i))
					FItemList(i).Fitemoptionname        = db2html(ArrList(7,i))
					FItemList(i).Fbuycash         		= ArrList(8,i)
					FItemList(i).Fsuplycash         	= ArrList(9,i)
					FItemList(i).Fitemno         		= ArrList(10,i)
					FItemList(i).Fvatyn         		= ArrList(11,i)

				next
			end if

        end sub

        Private Sub Class_Initialize()
                FCurrPage       = 1
                FPageSize       = 1000
                FResultCount    = 0
                FScrollCount    = 10
                FTotalCount     = 0
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

end Class

%>
