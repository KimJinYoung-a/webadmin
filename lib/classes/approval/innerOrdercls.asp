<%
'###########################################################
' Description :  cs 메모
' History : 2007.10.26 한용민 수정
'###########################################################

Class CInnerOrderItem
        public Fidx
        public Fdetailidx

        public Fdivcd
        public Facc_cd
        public Facc_nm

        public FappDate
        public FSELLBIZSECTION_CD
        public FBUYBIZSECTION_CD

        public FSELLBIZSECTION_NM
        public FBUYBIZSECTION_NM

        public FtotalSum
        public FsupplySum
        public FtaxSum
        public Fregdate
        public Freguserid

		public Fselluserid
		public Fbuyuserid

		public Fshopid
		public Fshopname
		public Fmakerid
        public FmakerTotalSum
        public FmakerSupplySum
        public FmakerTaxSum

        public Fsitename
        public Fmeachulgubun
        public Fsitefee
        public Ftotalsellcash
        public Ftotalchulgocash
        public Ftotalbuycash
        public Finnerorderpercentage

        public Fitemdiv
        public Fpricediv
        public Fdealdiv

	    public function GetDivcdColor()

	        if (Fdivcd="101") then
	        	GetDivcdColor = "green"
	        elseif (Fdivcd="102") then
	        	GetDivcdColor = "#FF0000"
	        elseif (Fdivcd="103") then
	        	GetDivcdColor = "#000000"
	        elseif (Fdivcd="201") then
	        	GetDivcdColor = "blue"
	        elseif (Fdivcd="202") then
	        	GetDivcdColor = "blue"
	        elseif (Fdivcd="301") then
	        	GetDivcdColor = "green"
	        elseif (Fdivcd="302") then
	        	GetDivcdColor = "green"
	        elseif (Fdivcd="303") then
	        	GetDivcdColor = "green"
	        elseif (Fdivcd="304") then
	        	GetDivcdColor = "green"
	        elseif (Fdivcd="305") then
	        	GetDivcdColor = "green"
	        elseif (Fdivcd="306") then
	        	GetDivcdColor = "green"
	        elseif (Fdivcd="307") then
	        	GetDivcdColor = "green"
	        elseif (Fdivcd="501") then
	        	GetDivcdColor = "green"
	        elseif (Fdivcd="502") then
	        	GetDivcdColor = "green"
	        else
	            GetDivcdColor = "#000000"
	        end if

	    end function

	    public function GetDivcdName()

	        if (Fdivcd="101") then
	        	GetDivcdName = "매장매입"
	        elseif (Fdivcd="102") then
	        	GetDivcdName = "업체특정"
	        elseif (Fdivcd="103") then
	        	GetDivcdName = "기타정산"
	        elseif (Fdivcd="201") then
	        	GetDivcdName = "아이띵소매입(ON)"
	        elseif (Fdivcd="202") then
	        	GetDivcdName = "아이띵소매출(ON)"
	        elseif (Fdivcd="301") then
	        	GetDivcdName = "출고매입(ON상품)"
	        elseif (Fdivcd="302") then
	        	GetDivcdName = "출고매입(OFF상품)"
	        elseif (Fdivcd="303") then
	        	GetDivcdName = "기타매입(ON상품)"
	        elseif (Fdivcd="304") then
	        	GetDivcdName = "기타매입(OFF상품)"
	        elseif (Fdivcd="305") then
	        	GetDivcdName = "출고매입(띵소상품)"
	        elseif (Fdivcd="306") then
	        	GetDivcdName = "기타매입(띵소상품)"
	        elseif (Fdivcd="307") then
	        	GetDivcdName = "출고매입(특정상품)"
	        elseif (Fdivcd="501") then
	        	GetDivcdName = "매장판매(띵소상품)"
	        elseif (Fdivcd="502") then
	        	GetDivcdName = "기타판매(띵소상품)"
	        else
	            GetDivcdName = Fdivcd
	        end if

	    end function

	    public function GetMeachulGubunName()

	        if (Fmeachulgubun="1") then
	        	GetMeachulGubunName = "소비자매출"
	        elseif (Fmeachulgubun="2") then
	        	GetMeachulGubunName = "출고가매출"
	        elseif (Fmeachulgubun="9") then
	        	GetMeachulGubunName = "매출없음"
	        else
	            GetMeachulGubunName = Fmeachulgubun
	        end if

	    end function

	    public function GetItemDivName()

	        if (Fitemdiv="1") then
	        	GetItemDivName = "온매입"
	        elseif (Fitemdiv="2") then
	        	GetItemDivName = "오프매입"
	        elseif (Fitemdiv="3") then
	        	GetItemDivName = "띵소매입"
	        elseif (Fitemdiv="4") then
	        	GetItemDivName = "특정상품"
	        else
	            GetItemDivName = Fitemdiv
	        end if

	    end function

	    public function GetPriceDivName()

	        if (Fpricediv="1") then
	        	GetPriceDivName = "출고액"
	        elseif (Fpricediv="2") then
	        	GetPriceDivName = "정산액"
	        elseif (Fpricediv="3") then
	        	GetPriceDivName = "매출액"
	        else
	            GetPriceDivName = Fpricediv
	        end if

	    end function

	    public function GetDealDivName()

	        if (Fdealdiv="1") then
	        	GetDealDivName = "매입액이전"
	        elseif (Fdealdiv="2") then
	        	GetDealDivName = "매출분배"
	        else
	            GetDealDivName = Fdealdiv
	        end if

	    end function

        Private Sub Class_Initialize()

        End Sub

        Private Sub Class_Terminate()

        End Sub
end Class

Class CInnerOrder
        public FItemList()
        public FOneItem

        public FCurrPage
        public FTotalPage
        public FPageSize
        public FResultCount
        public FScrollCount
        public FTotalCount

		public FRectIdx
		public FRectStartYYYYMMDD
		public FRectEndYYYYMMDD
		public FRectBizSection_CD
		''public FRectGroupingYN
        'public FRectUserID

        public FRectGroupBy

		public function GetFromWhere
			dim tmpSql

			GetFromWhere = ""

			tmpSql = " from "
			tmpSql = tmpSql + "	db_partner.dbo.tbl_InternalOrder o "
			tmpSql = tmpSql + "	join db_partner.dbo.tbl_TMS_BA_BIZSECTION bsell "
			tmpSql = tmpSql + "	on "
			tmpSql = tmpSql + "		o.SELLBIZSECTION_CD = bsell.BIZSECTION_CD "
			tmpSql = tmpSql + "	join db_partner.dbo.tbl_TMS_BA_BIZSECTION bbuy "
			tmpSql = tmpSql + "	on "
			tmpSql = tmpSql + "		o.BUYBIZSECTION_CD = bbuy.BIZSECTION_CD "
			tmpSql = tmpSql + "	left join db_partner.dbo.tbl_TMS_SL_ACC_CD a "
			tmpSql = tmpSql + "	on "
			tmpSql = tmpSql + "		o.acc_cd = a.acc_cd "
			tmpSql = tmpSql + " where "
			tmpSql = tmpSql + " 	1 = 1 "
			tmpSql = tmpSql + " 	and o.useyn = 'Y' "

			if (FRectIdx <> "") then
				tmpSql = tmpSql + " 	and o.idx = " + CStr(FRectIdx) + " "
			end if

			if (FRectStartYYYYMMDD <> "") then
				tmpSql = tmpSql + " 	and o.appDate >= '" + CStr(FRectStartYYYYMMDD) + "' "
			end if

			if (FRectEndYYYYMMDD <> "") then
				tmpSql = tmpSql + " 	and o.appDate < '" + CStr(FRectEndYYYYMMDD) + "' "
			end if

			if (FRectBizSection_CD <> "") then
				tmpSql = tmpSql + " 	and ((o.SELLBIZSECTION_CD = '" + CStr(FRectBizSection_CD) + "') or (o.BUYBIZSECTION_CD = '" + CStr(FRectBizSection_CD) + "')) "
			end if

			GetFromWhere = tmpSql

		end function

		public function GetFromWhereDetail
			dim tmpSql

			GetFromWhereDetail = ""

			tmpSql = " from "
			tmpSql = tmpSql + "	db_partner.dbo.tbl_InternalOrder o "
			tmpSql = tmpSql + "	join db_partner.dbo.tbl_InternalOrderDetail d "
			tmpSql = tmpSql + "	on "
			tmpSql = tmpSql + "		o.idx = d.masteridx "
			tmpSql = tmpSql + "	join db_partner.dbo.tbl_TMS_BA_BIZSECTION bsell "
			tmpSql = tmpSql + "	on "
			tmpSql = tmpSql + "		o.SELLBIZSECTION_CD = bsell.BIZSECTION_CD "
			tmpSql = tmpSql + "	join db_partner.dbo.tbl_TMS_BA_BIZSECTION bbuy "
			tmpSql = tmpSql + "	on "
			tmpSql = tmpSql + "		o.BUYBIZSECTION_CD = bbuy.BIZSECTION_CD "
			tmpSql = tmpSql + "	left join db_partner.dbo.tbl_TMS_SL_ACC_CD a "
			tmpSql = tmpSql + "	on "
			tmpSql = tmpSql + "		o.acc_cd = a.acc_cd "
			tmpSql = tmpSql + "	left join [db_user].[dbo].tbl_user_c c "
			tmpSql = tmpSql + "	on "
			tmpSql = tmpSql + "		c.userid = d.shopid "
			tmpSql = tmpSql + " where "
			tmpSql = tmpSql + " 	1 = 1 "
			tmpSql = tmpSql + " 	and o.useyn = 'Y' "
			tmpSql = tmpSql + " 	and o.idx = " + CStr(FRectIdx) + " "

			GetFromWhereDetail = tmpSql

		end function

        public Sub GetInnerOrderList()
            dim i,sqlStr

			'// ===============================================================
			sqlStr = " select count(o.idx) as cnt "

			sqlStr = sqlStr + GetFromWhere

			'response.write sqlStr
	        rsget.Open sqlStr, dbget, 1
	            FTotalCount = rsget("cnt")
	        rsget.Close

			'// ===============================================================
			sqlStr = " select top " + CStr(FPageSize*FCurrPage)
			sqlStr = sqlStr + " o.idx, o.divcd, o.appDate, o.SELLBIZSECTION_CD, o.BUYBIZSECTION_CD, o.totalSum, o.supplySum, o.taxSum, o.regdate, o.reguserid "
			sqlStr = sqlStr + " , bsell.BIZSECTION_NM as SELLBIZSECTION_NM, bbuy.BIZSECTION_NM as BUYBIZSECTION_NM "
			sqlStr = sqlStr + " , o.acc_cd, a.acc_nm "

			sqlStr = sqlStr + GetFromWhere

			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	o.appDate desc, idx desc "

	        rsget.pagesize = FPageSize
	        rsget.Open sqlStr, dbget, 1
	        'response.write sqlStr

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
					set FItemList(i) = new CInnerOrderItem

					FItemList(i).Fidx         		= rsget("idx")

					FItemList(i).Fdivcd         	= rsget("divcd")
					FItemList(i).Facc_cd         	= rsget("acc_cd")

					if (FItemList(i).Facc_cd = "1") then
						FItemList(i).Facc_nm = "상품매출원가"
					elseif (FItemList(i).Facc_cd = "2") then
						FItemList(i).Facc_nm = "내부거래매출"
					else
						FItemList(i).Facc_nm = FItemList(i).Facc_cd
					end if

					''FItemList(i).Facc_nm         	= rsget("acc_nm")

					FItemList(i).FappDate         	= rsget("appDate")
					FItemList(i).FSELLBIZSECTION_CD = rsget("SELLBIZSECTION_CD")
					FItemList(i).FBUYBIZSECTION_CD  = rsget("BUYBIZSECTION_CD")

					FItemList(i).FSELLBIZSECTION_NM  = rsget("SELLBIZSECTION_NM")
					FItemList(i).FBUYBIZSECTION_NM  	= rsget("BUYBIZSECTION_NM")

					FItemList(i).FtotalSum         	= rsget("totalSum")
					FItemList(i).FsupplySum         = rsget("supplySum")
					FItemList(i).FtaxSum         	= rsget("taxSum")
					FItemList(i).Fregdate         	= rsget("regdate")
					FItemList(i).Freguserid         = rsget("reguserid")

					rsget.moveNext
					i=i+1
				loop
			end if

			rsget.Close
        end sub

        public Sub GetInnerOrderOne()
            dim i,sqlStr

			'// ===============================================================
			sqlStr = " select top 1 "
			sqlStr = sqlStr + " o.idx, o.divcd, o.appDate, o.SELLBIZSECTION_CD, o.BUYBIZSECTION_CD, o.totalSum, o.supplySum, o.taxSum, o.regdate, o.reguserid "
			sqlStr = sqlStr + " , bsell.BIZSECTION_NM as SELLBIZSECTION_NM, bbuy.BIZSECTION_NM as BUYBIZSECTION_NM "
			sqlStr = sqlStr + " , a.acc_cd, a.acc_nm "
			sqlStr = sqlStr + " , o.selluserid, o.buyuserid "

			sqlStr = sqlStr + GetFromWhere

	        rsget.pagesize = FPageSize
	        rsget.Open sqlStr, dbget, 1
	        'response.write sqlStr

			i=0
			if  not rsget.EOF  then
				set FOneItem = new CInnerOrderItem

				FOneItem.Fidx         		= rsget("idx")

				FOneItem.FappDate         	= rsget("appDate")

				FOneItem.Fdivcd         	= rsget("divcd")
				FOneItem.Facc_cd         	= rsget("acc_cd")
				FOneItem.Facc_nm         	= rsget("acc_nm")

				FOneItem.FappDate         	= rsget("appDate")
				FOneItem.FSELLBIZSECTION_CD = rsget("SELLBIZSECTION_CD")
				FOneItem.FBUYBIZSECTION_CD  = rsget("BUYBIZSECTION_CD")

				FOneItem.FSELLBIZSECTION_NM  = rsget("SELLBIZSECTION_NM")
				FOneItem.FBUYBIZSECTION_NM  	= rsget("BUYBIZSECTION_NM")

				FOneItem.FtotalSum         	= rsget("totalSum")
				FOneItem.FsupplySum         = rsget("supplySum")
				FOneItem.FtaxSum         	= rsget("taxSum")
				FOneItem.Fregdate         	= rsget("regdate")
				FOneItem.Freguserid         = rsget("reguserid")

				FOneItem.Fselluserid        = rsget("selluserid")
				FOneItem.Fbuyuserid         = rsget("buyuserid")
			else
				Call GetInnerOrderBlankOne()
			end if

			rsget.Close
        end sub

        public Sub GetInnerOrderSummaryList()
            dim i,sqlStr

			'// ===============================================================
			sqlStr = " select count(o.appDate) as cnt "

			sqlStr = sqlStr + GetFromWhere

			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	o.appDate, o.SELLBIZSECTION_CD, o.BUYBIZSECTION_CD, bsell.BIZSECTION_NM, bbuy.BIZSECTION_NM "

			'response.write sqlStr
	        rsget.Open sqlStr, dbget, 1
	            FTotalCount = rsget("cnt")
	        rsget.Close

			'// ===============================================================
			sqlStr = " select top " + CStr(FPageSize*FCurrPage)
			sqlStr = sqlStr + " o.appDate, o.SELLBIZSECTION_CD, o.BUYBIZSECTION_CD, sum(o.totalSum) as totalSum, sum(o.supplySum) as supplySum, sum(o.taxSum) as taxSum "
			sqlStr = sqlStr + " , bsell.BIZSECTION_NM as SELLBIZSECTION_NM, bbuy.BIZSECTION_NM as BUYBIZSECTION_NM "

			sqlStr = sqlStr + GetFromWhere

			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	o.appDate, o.SELLBIZSECTION_CD, o.BUYBIZSECTION_CD, bsell.BIZSECTION_NM, bbuy.BIZSECTION_NM "

			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	o.appDate desc "

	        rsget.pagesize = FPageSize
	        rsget.Open sqlStr, dbget, 1
	        'response.write sqlStr

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
					set FItemList(i) = new CInnerOrderItem

					FItemList(i).Fidx         		= -1
					'FItemList(i).Fdivcd         	= rsget("divcd")
					FItemList(i).FappDate         	= rsget("appDate")
					FItemList(i).FSELLBIZSECTION_CD = rsget("SELLBIZSECTION_CD")
					FItemList(i).FBUYBIZSECTION_CD  = rsget("BUYBIZSECTION_CD")

					FItemList(i).FSELLBIZSECTION_NM  = rsget("SELLBIZSECTION_NM")
					FItemList(i).FBUYBIZSECTION_NM  	= rsget("BUYBIZSECTION_NM")

					FItemList(i).FtotalSum         	= rsget("totalSum")
					FItemList(i).FsupplySum         = rsget("supplySum")
					FItemList(i).FtaxSum         	= rsget("taxSum")
					'FItemList(i).Fregdate         	= rsget("regdate")
					'FItemList(i).Freguserid         = rsget("reguserid")

					rsget.moveNext
					i=i+1
				loop
			end if

			rsget.Close
        end sub

        public Sub GetInnerOrderDetail()
            dim i,sqlStr

			sqlStr = " select top " + CStr(FPageSize*FCurrPage)
			sqlStr = sqlStr + " o.idx, o.divcd, o.appDate, o.SELLBIZSECTION_CD, o.BUYBIZSECTION_CD, o.totalSum, o.supplySum, o.taxSum, o.regdate, o.reguserid "
			sqlStr = sqlStr + " , bsell.BIZSECTION_NM as SELLBIZSECTION_NM, bbuy.BIZSECTION_NM as BUYBIZSECTION_NM "
			sqlStr = sqlStr + " , d.shopid, d.makerid, d.totalSum as makerTotalSum, d.supplySum as makerSupplySum, d.taxSum as makerTaxSum "

            sqlStr = sqlStr + GetFromWhereDetail

			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	d.supplySum desc "

	        rsget.pagesize = FPageSize
	        rsget.Open sqlStr, dbget, 1
	        'response.write sqlStr

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
					set FItemList(i) = new CInnerOrderItem

					FItemList(i).Fidx         		= rsget("idx")
					FItemList(i).Fdivcd         	= rsget("divcd")
					FItemList(i).FappDate         	= rsget("appDate")
					FItemList(i).FSELLBIZSECTION_CD = rsget("SELLBIZSECTION_CD")
					FItemList(i).FBUYBIZSECTION_CD  = rsget("BUYBIZSECTION_CD")

					FItemList(i).FSELLBIZSECTION_NM = rsget("SELLBIZSECTION_NM")
					FItemList(i).FBUYBIZSECTION_NM  = rsget("BUYBIZSECTION_NM")

					FItemList(i).FtotalSum         	= rsget("totalSum")
					FItemList(i).FsupplySum         = rsget("supplySum")
					FItemList(i).FtaxSum         	= rsget("taxSum")
					FItemList(i).Fregdate         	= rsget("regdate")
					FItemList(i).Freguserid         = rsget("reguserid")

					FItemList(i).Fshopid         	= rsget("shopid")
					FItemList(i).Fmakerid         	= rsget("makerid")
					FItemList(i).FmakerTotalSum     = rsget("makerTotalSum")
					FItemList(i).FmakerSupplySum    = rsget("makerSupplySum")
					FItemList(i).FmakerTaxSum     	= rsget("makerTaxSum")

					rsget.moveNext
					i=i+1
				loop
			end if

            rsget.close
        end sub

        public Sub GetInnerOrderDetailNew()
            dim i,sqlStr

			sqlStr = " select top " + CStr(FPageSize*FCurrPage)
			sqlStr = sqlStr + " o.idx, o.divcd, o.appDate, o.SELLBIZSECTION_CD, o.BUYBIZSECTION_CD, o.totalSum, o.supplySum, o.taxSum, o.regdate, o.reguserid "
			sqlStr = sqlStr + " , bsell.BIZSECTION_NM as SELLBIZSECTION_NM, bbuy.BIZSECTION_NM as BUYBIZSECTION_NM "
			sqlStr = sqlStr + " , sum(d.totalSum) as makerTotalSum, sum(d.supplySum) as makerSupplySum, sum(d.taxSum) as makerTaxSum "

			if (FRectGroupBy = "shopid") then
				sqlStr = sqlStr + " , d.shopid, '' as makerid "
			else
				sqlStr = sqlStr + " , '' as shopid, d.makerid "
			end if

            sqlStr = sqlStr + GetFromWhereDetail

			if (FRectGroupBy = "shopid") then
				sqlStr = sqlStr + " group by "
				sqlStr = sqlStr + " 	o.idx, o.divcd, o.appDate, o.SELLBIZSECTION_CD, o.BUYBIZSECTION_CD, o.totalSum, o.supplySum, o.taxSum, o.regdate, o.reguserid, bsell.BIZSECTION_NM, bbuy.BIZSECTION_NM, d.shopid "
			else
				sqlStr = sqlStr + " group by "
				sqlStr = sqlStr + " 	o.idx, o.divcd, o.appDate, o.SELLBIZSECTION_CD, o.BUYBIZSECTION_CD, o.totalSum, o.supplySum, o.taxSum, o.regdate, o.reguserid, bsell.BIZSECTION_NM, bbuy.BIZSECTION_NM, d.makerid "
			end if

			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	sum(d.supplySum) desc "

	        rsget.pagesize = FPageSize
	        rsget.Open sqlStr, dbget, 1
	        'response.write sqlStr

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
					set FItemList(i) = new CInnerOrderItem

					FItemList(i).Fidx         		= rsget("idx")
					FItemList(i).Fdivcd         	= rsget("divcd")
					FItemList(i).FappDate         	= rsget("appDate")
					FItemList(i).FSELLBIZSECTION_CD = rsget("SELLBIZSECTION_CD")
					FItemList(i).FBUYBIZSECTION_CD  = rsget("BUYBIZSECTION_CD")

					FItemList(i).FSELLBIZSECTION_NM = rsget("SELLBIZSECTION_NM")
					FItemList(i).FBUYBIZSECTION_NM  = rsget("BUYBIZSECTION_NM")

					FItemList(i).FtotalSum         	= rsget("totalSum")
					FItemList(i).FsupplySum         = rsget("supplySum")
					FItemList(i).FtaxSum         	= rsget("taxSum")
					FItemList(i).Fregdate         	= rsget("regdate")
					FItemList(i).Freguserid         = rsget("reguserid")

					FItemList(i).Fshopid         	= rsget("shopid")
					FItemList(i).Fmakerid         	= rsget("makerid")
					FItemList(i).FmakerTotalSum     = rsget("makerTotalSum")
					FItemList(i).FmakerSupplySum    = rsget("makerSupplySum")
					FItemList(i).FmakerTaxSum     	= rsget("makerTaxSum")

					rsget.moveNext
					i=i+1
				loop
			end if

            rsget.close
        end sub

        public Sub GetOnlineInnerOrderDetail()
            dim i,sqlStr

			sqlStr = " select top " + CStr(FPageSize*FCurrPage)
			sqlStr = sqlStr + " o.idx, o.divcd, o.appDate, o.SELLBIZSECTION_CD, o.BUYBIZSECTION_CD, o.totalSum, o.supplySum, o.taxSum, o.regdate, o.reguserid "
			sqlStr = sqlStr + " , bsell.BIZSECTION_NM as SELLBIZSECTION_NM, bbuy.BIZSECTION_NM as BUYBIZSECTION_NM "
			sqlStr = sqlStr + " , d.shopid, d.makerid, d.totalSum as makerTotalSum, d.supplySum as makerSupplySum, d.taxSum as makerTaxSum, d.sitename, d.meachulgubun, d.sitefee, d.totalsellcash, d.innerorderpercentage, d.idx as detailidx "

            sqlStr = sqlStr + GetFromWhereDetail

			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	d.supplySum desc "

	        rsget.pagesize = FPageSize
	        rsget.Open sqlStr, dbget, 1
	        'response.write sqlStr

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
					set FItemList(i) = new CInnerOrderItem

					FItemList(i).Fidx         		= rsget("idx")
					FItemList(i).Fdetailidx         = rsget("detailidx")

					FItemList(i).Fdivcd         	= rsget("divcd")
					FItemList(i).FappDate         	= rsget("appDate")
					FItemList(i).FSELLBIZSECTION_CD = rsget("SELLBIZSECTION_CD")
					FItemList(i).FBUYBIZSECTION_CD  = rsget("BUYBIZSECTION_CD")

					FItemList(i).FSELLBIZSECTION_NM = rsget("SELLBIZSECTION_NM")
					FItemList(i).FBUYBIZSECTION_NM  = rsget("BUYBIZSECTION_NM")

					FItemList(i).FtotalSum         	= rsget("totalSum")
					FItemList(i).FsupplySum         = rsget("supplySum")
					FItemList(i).FtaxSum         	= rsget("taxSum")
					FItemList(i).Fregdate         	= rsget("regdate")
					FItemList(i).Freguserid         = rsget("reguserid")

					FItemList(i).Fshopid         	= rsget("shopid")
					FItemList(i).Fmakerid         	= rsget("makerid")
					FItemList(i).FmakerTotalSum     = rsget("makerTotalSum")
					FItemList(i).FmakerSupplySum    = rsget("makerSupplySum")
					FItemList(i).FmakerTaxSum     	= rsget("makerTaxSum")

					FItemList(i).Fsitename			= rsget("sitename")
					FItemList(i).Fmeachulgubun		= rsget("meachulgubun")
					FItemList(i).Fsitefee			= rsget("sitefee")
					FItemList(i).Ftotalsellcash		= rsget("totalsellcash")
					FItemList(i).Finnerorderpercentage	= rsget("innerorderpercentage")

					rsget.moveNext
					i=i+1
				loop
			end if

            rsget.close
        end sub

        public Sub GetOnOffInnerOrderDetailNew()
            dim i,sqlStr

			sqlStr = " select top " + CStr(FPageSize*FCurrPage)
			sqlStr = sqlStr + " o.idx, o.divcd, o.appDate, o.SELLBIZSECTION_CD, o.BUYBIZSECTION_CD, o.totalSum, o.supplySum, o.taxSum, o.regdate, o.reguserid "
			sqlStr = sqlStr + " , bsell.BIZSECTION_NM as SELLBIZSECTION_NM, bbuy.BIZSECTION_NM as BUYBIZSECTION_NM "
			sqlStr = sqlStr + " , d.shopid, d.makerid, d.totalSum as makerTotalSum, d.supplySum as makerSupplySum, d.taxSum as makerTaxSum, d.sitename, d.meachulgubun, d.sitefee, d.totalsellcash, d.totalchulgocash, d.totalbuycash, d.innerorderpercentage, d.itemdiv, d.pricediv, d.dealdiv, d.idx as detailidx, c.socname_kor "

            sqlStr = sqlStr + GetFromWhereDetail

			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	d.shopid, d.dealdiv, d.supplySum desc "

	        rsget.pagesize = FPageSize
	        rsget.Open sqlStr, dbget, 1
	        'response.write sqlStr

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
					set FItemList(i) = new CInnerOrderItem

					FItemList(i).Fidx         		= rsget("idx")
					FItemList(i).Fdetailidx         = rsget("detailidx")

					FItemList(i).Fdivcd         	= rsget("divcd")
					FItemList(i).FappDate         	= rsget("appDate")
					FItemList(i).FSELLBIZSECTION_CD = rsget("SELLBIZSECTION_CD")
					FItemList(i).FBUYBIZSECTION_CD  = rsget("BUYBIZSECTION_CD")

					FItemList(i).FSELLBIZSECTION_NM = rsget("SELLBIZSECTION_NM")
					FItemList(i).FBUYBIZSECTION_NM  = rsget("BUYBIZSECTION_NM")

					FItemList(i).FtotalSum         	= rsget("totalSum")
					FItemList(i).FsupplySum         = rsget("supplySum")
					FItemList(i).FtaxSum         	= rsget("taxSum")
					FItemList(i).Fregdate         	= rsget("regdate")
					FItemList(i).Freguserid         = rsget("reguserid")

					FItemList(i).Fshopid         	= rsget("shopid")
					FItemList(i).Fmakerid         	= rsget("makerid")
					FItemList(i).FmakerTotalSum     = rsget("makerTotalSum")
					FItemList(i).FmakerSupplySum    = rsget("makerSupplySum")
					FItemList(i).FmakerTaxSum     	= rsget("makerTaxSum")

					FItemList(i).Fsitename			= rsget("sitename")
					FItemList(i).Fmeachulgubun		= rsget("meachulgubun")
					FItemList(i).Fsitefee			= rsget("sitefee")
					FItemList(i).Ftotalsellcash		= rsget("totalsellcash")
					FItemList(i).Ftotalchulgocash	= rsget("totalchulgocash")
					FItemList(i).Ftotalbuycash		= rsget("totalbuycash")

					FItemList(i).Fitemdiv			= rsget("itemdiv")
					FItemList(i).Fpricediv			= rsget("pricediv")
					FItemList(i).Fdealdiv			= rsget("dealdiv")

					FItemList(i).Finnerorderpercentage	= rsget("innerorderpercentage")

					FItemList(i).Fshopname			= rsget("socname_kor")

					if IsNull(FItemList(i).Ftotalchulgocash) then
						FItemList(i).Ftotalchulgocash = 0
					end if

					if IsNull(FItemList(i).Ftotalbuycash) then
						FItemList(i).Ftotalbuycash = 0
					end if

					rsget.moveNext
					i=i+1
				loop
			end if

            rsget.close
        end sub

        public Sub GetInnerOrderBlankDetail()
            dim i,sqlStr

            set FOneItem = new CInnerOrderItem
        end sub

        public Sub GetInnerOrderBlankOne()
            dim i,sqlStr

            set FOneItem = new CInnerOrderItem
        end sub

        Private Sub Class_Initialize()
                FCurrPage       = 1
                FPageSize       = 20
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
