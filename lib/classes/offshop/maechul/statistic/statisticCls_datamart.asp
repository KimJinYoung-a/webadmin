<%
'###########################################################
' Description :  오프라인 매출통계 클래스
' History : 2012.10.04 한용민 생성
'###########################################################

class cStaticdatamart_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public FRegdate
	public FMaechulPlus
	public FMaechulMinus
	public FCountPlus
	public FCountMinus
	public FSubtotalprice
	public FMiletotalprice
	public FTotalcheckprice
	public FMinDate
	public FMaxDate
	public FWeek
	public FMonth
	public fspendmile
	public fspendmilecnt
	public fTenGiftCardPaySum
	public fTenGiftCardPaycount
	public fcardsum
	public fcardcnt
	public fcashsum
	public fcashcnt
	public fgiftcardPaysum
	public fgiftcardPaycnt
	public fselltotal
	public fordercnt
	public fshopid
	public fshopname
	public ftotalsum
	public frealsum
	public fgainmile
	public FMaechul
	public fbonuscouponprice
	public fitemno
	public fIorgsellprice
	public fsellprice
	public frealsellprice
	public fsuplyprice
	public fshopbuyprice
	public FMaechulProfit
	public FMaechulProfitPer
	public FMakerid
	public FCount
	public FSum
	public fprofit
	public FaddTaxChargeSum
	public FChargeDiv
	public FPurchasetype
	public fIXyyyymmdd
	public FCateCDL
	public FCateName
	public Fsellsum
	public Fsellcnt
	public FCateCDM
	public FCateCDN
	public Fitemgubun
	public FItemID
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	Public Fjcomm_cd
	public fpurchasetypename

	public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

		public function getChargeDivName()
		if FChargeDiv="2" then
			getChargeDivName = "출고위탁" '10x10 위탁
		elseif FChargeDiv="4" then
			getChargeDivName = "출고매입" ''"10x10 매입"
		elseif FChargeDiv="5" then
			getChargeDivName = "출고매입" '출고분정산
		elseif FChargeDiv="6" then
			getChargeDivName = "업체위탁"
		elseif FChargeDiv="8" then
			getChargeDivName = "업체매입"
		elseif FChargeDiv="9" then
			getChargeDivName = "가맹점"
		elseif FChargeDiv="0" then
			getChargeDivName = "통합"
		else
			getChargeDivName = FChargeDiv
		end if
	end function
	
	' 사용중지. 디비에서 일괄로 쿼리해서 가져 오세요.
	public Function getPurchasetypeName()
    	IF FPurchasetype = "1" then
    	    getPurchasetypeName = "일반유통" 
    	ELSEIF FPurchasetype = "4" then
    	    getPurchasetypeName = "사입" 
    	ELSEIF FPurchasetype = "5" then
    	    getPurchasetypeName = "OFF사입" 
    	ELSEIF FPurchasetype = "6" then
    	    getPurchasetypeName = "수입" 
    	ELSEIF FPurchasetype = "7" then
    	    getPurchasetypeName = "브랜드수입"
        ELSEIF FPurchasetype = "8" then
    	    getPurchasetypeName = "제작" 
        ELSEIF FPurchasetype = "9" then
    	    getPurchasetypeName = "해외직구" 
        ELSEIF FPurchasetype = "10" then
    	    getPurchasetypeName = "B2B" 
    	END IF
    end Function
end class

class cStaticdatamart_list
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		IF application("Svr_Info")="Dev" THEN
			FTENDB="[TENDB]."
		end if
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

	public FItemList()
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	public FRectdatefg
	public FRectStartdate
	public FRectEndDate
	public FRectshopid
	public FRectBanPum
	public FRectPurchasetype
	public frectmakerid
	public FRectOffgubun
	public FRectOldData
	public FRectNormalOnly
	public FRectStartDay
	public FRectEndDay
	public frectdategubun
	public frectoffcatecode
	public frectoffmduserid
	public FRectBrandPurchaseType
	public FRectOrdertype
	public FRectOnlyShop
	public FRectCDL
	public FRectCDM
	public FRectCDN
	public frectcatecdnull
	public frectweekdate
	public maxt
	public maxc
	public FTotalSum
	public FTENDB
	public FRectInc3pl
	''public FRectJungSanGubun
	public FRectCommCd
	public FRectOnlyTenShop
	Public FRectChkShowGubun

	'/common/offshop/maechul/statistic/statistic_category_datamart.asp
	public sub SearchCategorySellrePort3_datamart()
		Dim sql, i ,sqlsearch

	    maxt = -1
	    maxc = -1

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')=''"
	    end if
		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum >= 0"
		end if

		if FRectShopID<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopID + "'"
		end if
		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.IXyyyymmdd) = "&frectweekdate&""
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid = '"&FRectmakerid&"'"
		end if

		sql = "select"
		sql = sql + " isNull(sum(d.itemno),0) as sellcnt"
		sql = sql + " ,isNull(sum( (d.realsellprice) *d.itemno),0) as sumtotal"
		sql = sql + " ,isNull(sum( (d.suplyprice) *d.itemno),0) as suplyprice"
		sql = sql & " ,isNull(sum( (d.Iorgsellprice) *d.itemno),0) AS Iorgsellprice"
		sql = sql & " ,isNull(sum( (d.sellprice) *d.itemno),0) AS sellprice"
		sql = sql & " ,isNull(sum( (d.shopbuyprice) *d.itemno),0) AS shopbuyprice"
		sql = sql + " , s.catecdn,cs.code_nm"

		if FRectOldData="on" then
			sql = sql + " from "&FTENDB&"[db_shoplog].[dbo].tbl_old_shopjumun_master m with (nolock)"
			sql = sql + " join "&FTENDB&"[db_shoplog].[dbo].tbl_old_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		else
			sql = sql + " from "&FTENDB&"[db_shop].[dbo].tbl_shopjumun_master m with (nolock)"
			sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		end if

		if (FRectBrandPurchaseType<>"") then
		    sql = sql + " Join "&FTENDB&"db_partner.dbo.tbl_partner p with (nolock)"
		    sql = sql + " 	on d.makerid=p.id"

			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				sql = sql + " 	and p.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				sql = sql & " 	and p.purchasetype in ('3','5','6')"
			else
				sql = sql + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shop_item s with (nolock)"
		sql = sql + " 	on d.itemgubun=s.itemgubun"
		sql = sql + " 	and d.itemid=s.shopitemid"
		sql = sql + " 	and d.itemoption=s.itemoption"
		sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shop_user u with (nolock)"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join "&FTENDB&"db_partner.dbo.tbl_partner pp with (nolock)"
	    sql = sql & "       on m.shopid=pp.id "
		sql = sql + " left join "&FTENDB&"[db_item].[dbo].tbl_cate_small cs with (nolock)"
		sql = sql + " 	on s.catecdl=cs.code_large and s.catecdm=cs.code_mid and s.catecdn=cs.code_small"
		sql = sql + " where 1=1 "
		sql = sql + " and s.catecdl='" + FRectCDL + "'"
		sql = sql + " and s.catecdm='" + FRectCDM + "' " & sqlsearch
		sql = sql + " group by s.catecdn,cs.code_nm"
		sql = sql + " order by s.catecdn"

		'response.write sql &"<br>"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sql, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
	    redim preserve FItemList(FResultCount)
		do until db3_rsget.eof
			set FItemList(i) = new cStaticdatamart_oneitem

			FItemList(i).fIorgsellprice	= db3_rsget("Iorgsellprice")
			FItemList(i).fsellprice	= db3_rsget("sellprice")
			FItemList(i).fsuplyprice	= db3_rsget("suplyprice")
			FItemList(i).fshopbuyprice	= db3_rsget("shopbuyprice")
			FItemList(i).FCateCDL	= FRectCDL
			FItemList(i).FCateCDM	= FRectCDM
			FItemList(i).FCateCDN	= db3_rsget("catecdn")
		    FItemList(i).FCateName 	= db2html(db3_rsget("code_nm"))
			FItemList(i).Fsellsum	= db3_rsget("sumtotal")
			FItemList(i).Fsellcnt 	= db3_rsget("sellcnt")

			if IsNULL(FItemList(i).FCateName) then FItemList(i).FCateName = "미지정"

			if Not IsNull(FItemList(i).Fsellsum) then
				maxt = maxvalreturn(maxt,FItemList(i).Fsellsum)
				maxc = maxvalreturn(maxc,FItemList(i).Fsellcnt)
			end if

			if Not IsNull(FItemList(i).fsellsum) then
				FTotalSum = FTotalSum + db3_rsget("sumtotal")
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close
	end sub

	'/common/offshop/maechul/statistic/statistic_category_datamart.asp
	public sub SearchCategorySellrePort2_datamart()
		Dim sql, i ,sqlsearch

	    maxt = -1
	    maxc = -1

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')=''"
	    end if
		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum >= 0"
		end if

		if FRectShopID<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopID + "'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.IXyyyymmdd) = "&frectweekdate&""
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if
		end if
		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid = '"&FRectmakerid&"'"
		end if

		sql = "select"
		sql = sql + " isNull(sum(d.itemno),0) as sellcnt"
		sql = sql + " ,isNull(sum( (d.realsellprice) *d.itemno),0) as sumtotal"
		sql = sql + " ,isNull(sum( (d.suplyprice) *d.itemno),0) as suplyprice"
		sql = sql & " ,isNull(sum( (d.Iorgsellprice) *d.itemno),0) AS Iorgsellprice"
		sql = sql & " ,isNull(sum( (d.sellprice) *d.itemno),0) AS sellprice"
		sql = sql & " ,isNull(sum( (d.shopbuyprice) *d.itemno),0) AS shopbuyprice"
		sql = sql + " , s.catecdm,cm.code_nm"

		if FRectOldData="on" then
			sql = sql + " from "&FTENDB&"[db_shoplog].[dbo].tbl_old_shopjumun_master m with (nolock)"
			sql = sql + " join "&FTENDB&"[db_shoplog].[dbo].tbl_old_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		else
			sql = sql + " from "&FTENDB&"[db_shop].[dbo].tbl_shopjumun_master m with (nolock)"
			sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		end if

		if (FRectBrandPurchaseType<>"") then
		    sql = sql + " Join "&FTENDB&"db_partner.dbo.tbl_partner p with (nolock)"
		    sql = sql + " 	on d.makerid=p.id"

			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				sql = sql + " 	and p.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				sql = sql & " 	and p.purchasetype in ('3','5','6')"
			else
				sql = sql + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shop_item s with (nolock)"
		sql = sql + " 	on d.itemgubun=s.itemgubun"
		sql = sql + " 	and d.itemid=s.shopitemid"
		sql = sql + " 	and d.itemoption=s.itemoption"
		sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shop_user u with (nolock)"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join "&FTENDB&"db_partner.dbo.tbl_partner pp with (nolock)"
	    sql = sql & "       on m.shopid=pp.id "
		sql = sql + " left join "&FTENDB&"[db_item].[dbo].tbl_cate_mid cm with (nolock)"
		sql = sql + " 	on s.catecdl=cm.code_large and s.catecdm=cm.code_mid"
		sql = sql + " where 1=1 "
		sql = sql + " and s.catecdl='" + FRectCDL + "' " & sqlsearch
		sql = sql + " group by s.catecdm,cm.code_nm"
		sql = sql + " order by s.catecdm"

		'response.write sql &"<br>"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sql, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
	    redim preserve FItemList(FResultCount)
		do until db3_rsget.eof
			set FItemList(i) = new cStaticdatamart_oneitem

			FItemList(i).fIorgsellprice	= db3_rsget("Iorgsellprice")
			FItemList(i).fsellprice	= db3_rsget("sellprice")
			FItemList(i).fsuplyprice	= db3_rsget("suplyprice")
			FItemList(i).fshopbuyprice	= db3_rsget("shopbuyprice")
			FItemList(i).FCateCDL	= FRectCDL
			FItemList(i).FCateCDM	= db3_rsget("catecdm")
		    FItemList(i).FCateName 	= db2html(db3_rsget("code_nm"))
			FItemList(i).Fsellsum	= db3_rsget("sumtotal")
			FItemList(i).Fsellcnt 	= db3_rsget("sellcnt")

			if IsNULL(FItemList(i).FCateName) then
				 FItemList(i).FCateName = "미지정"
			end if

			if Not IsNull(FItemList(i).Fsellsum) then
				maxt = maxvalreturn(maxt,FItemList(i).Fsellsum)
				maxc = maxvalreturn(maxc,FItemList(i).Fsellcnt)
			end if

			if Not IsNull(FItemList(i).fsellsum) then
				FTotalSum = FTotalSum + db3_rsget("sumtotal")
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close
	end sub

	'/common/offshop/maechul/statistic/statistic_category_datamart.asp
	public sub Searchcategorysellreport_Datamart()
		Dim sql, i ,sqlsearch

	    maxt = -1
	    maxc = -1

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')=''"
	    end if
		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum >= 0"
		end if

		if FRectShopID<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopID + "'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.IXyyyymmdd) = "&frectweekdate&""
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if
		end if
		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid = '"&FRectmakerid&"'"
		end if

		sql = "select"
		sql = sql + " isNull(sum(d.itemno),0) as sellcnt"
		sql = sql + " ,isNull(sum( (d.realsellprice) *d.itemno),0) as sumtotal"
		sql = sql + " ,isNull(sum( (d.suplyprice) *d.itemno),0) as suplyprice"
		sql = sql & " ,isNull(sum( (d.Iorgsellprice) *d.itemno),0) AS Iorgsellprice"
		sql = sql & " ,isNull(sum( (d.sellprice) *d.itemno),0) AS sellprice"
		sql = sql & " ,isNull(sum( (d.shopbuyprice) *d.itemno),0) AS shopbuyprice"
		sql = sql + " , s.catecdl,cl.code_nm"

		if FRectOldData="on" then
			sql = sql + " from "&FTENDB&"[db_shoplog].[dbo].tbl_old_shopjumun_master m with (nolock)"
			sql = sql + " join "&FTENDB&"[db_shoplog].[dbo].tbl_old_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		else
			sql = sql + " from "&FTENDB&"[db_shop].[dbo].tbl_shopjumun_master m with (nolock)"
			sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		end if

		if (FRectBrandPurchaseType<>"") then
		    sql = sql + " Join "&FTENDB&"db_partner.dbo.tbl_partner p with (nolock)"
		    sql = sql + " 	on d.makerid=p.id"

			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				sql = sql + " 	and p.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				sql = sql & " 	and p.purchasetype in ('3','5','6')"
			else
				sql = sql + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shop_item s with (nolock)"
		sql = sql + " 	on d.itemgubun=s.itemgubun"
		sql = sql + " 	and d.itemid=s.shopitemid"
		sql = sql + " 	and d.itemoption=s.itemoption"
		sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shop_user u with (nolock)"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join "&FTENDB&"db_partner.dbo.tbl_partner pp with (nolock)"
	    sql = sql & "       on m.shopid=pp.id "
		sql = sql + " left join "&FTENDB&"[db_item].[dbo].tbl_Cate_large cl with (nolock)"
		sql = sql + " 	on s.catecdl=cl.code_large"
		sql = sql + " where 1=1 " & sqlsearch
		sql = sql + " group by s.catecdl,cl.code_nm"
		sql = sql + " order by s.catecdl"

		'response.write sql &"<br>"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sql, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof

			set FItemList(i) = new cStaticdatamart_oneitem

			FItemList(i).fIorgsellprice	= db3_rsget("Iorgsellprice")
			FItemList(i).fsellprice	= db3_rsget("sellprice")
			FItemList(i).fsuplyprice	= db3_rsget("suplyprice")
			FItemList(i).fshopbuyprice	= db3_rsget("shopbuyprice")
			FItemList(i).FCateCDL	= db3_rsget("catecdl")
		    FItemList(i).FCateName 	= db2html(db3_rsget("code_nm"))
			FItemList(i).Fsellsum	= db3_rsget("sumtotal")
			FItemList(i).Fsellcnt 	= db3_rsget("sellcnt")

			if IsNULL(FItemList(i).FCateName) then
				if IsNULL(FItemList(i).FCateCDL) then FItemList(i).FCateCDL=""
				FItemList(i).FCateName = "미지정"
			end if

			if Not IsNull(FItemList(i).Fsellsum) then
				maxt = maxvalreturn(maxt,FItemList(i).Fsellsum)
				maxc = maxvalreturn(maxc,FItemList(i).Fsellcnt)
			end if

			if Not IsNull(FItemList(i).fsellsum) then
				FTotalSum = FTotalSum + db3_rsget("sumtotal")
			end if

			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close
	end sub

	'/common/offshop/maechul/statistic/statistic_category_datamart.asp
	public sub SearchCategorySellItems_datamart()
		Dim sql, i ,sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')=''"
	    end if
		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum >= 0"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.IXyyyymmdd) = "&frectweekdate&""
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if
		end if

		if FRectShopID<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopID + "'"
		end if
		if FRectCDL<>"" then
			sqlsearch = sqlsearch + " and s.catecdl='" + FRectCDL + "'"
		end if
		if FRectCDM<>"" then
			sqlsearch = sqlsearch + " and s.catecdm='" + FRectCDM + "'"
		end if
		if FRectCDN<>"" then
			sqlsearch = sqlsearch + " and s.catecdn='" + FRectCDN + "'"
		end if
		if frectcatecdnull = "ON" then
			sqlsearch = sqlsearch + " and s.catecdl is NULL"
		end if
		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid = '"&FRectmakerid&"'"
		end if

		sql = "select top " + CStr(FCurrPage*FPageSize)
		sql = sql + " d.makerid, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.realsellprice"
		sql = sql + " ,isNull(sum(d.itemno),0) as sellcnt"
		sql = sql + " ,isNull(sum( (d.realsellprice) *d.itemno),0) as sellsum"
		sql = sql + " ,isNull(sum( (d.suplyprice) *d.itemno),0) as suplyprice"
		sql = sql & " ,isNull(sum( (d.Iorgsellprice) *d.itemno),0) AS Iorgsellprice"
		sql = sql & " ,isNull(sum( (d.sellprice) *d.itemno),0) AS sellprice"
		sql = sql & " ,isNull(sum( (d.shopbuyprice) *d.itemno),0) AS shopbuyprice"

		if FRectOldData="on" then
			sql = sql + " from "&FTENDB&"[db_shoplog].[dbo].tbl_old_shopjumun_master m with (nolock)"
			sql = sql + " Join "&FTENDB&"[db_shoplog].[dbo].tbl_old_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.idx = d.masteridx"
		else
			sql = sql + " from "&FTENDB&"[db_shop].[dbo].tbl_shopjumun_master m with (nolock)"
			sql = sql + " Join "&FTENDB&"[db_shop].[dbo].tbl_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.idx = d.masteridx"
		end if

		if (FRectBrandPurchaseType<>"") then
		    sql = sql + " Join "&FTENDB&"db_partner.dbo.tbl_partner p with (nolock)"
		    sql = sql + " 	on d.makerid=p.id"

			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				sql = sql + " 	and p.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				sql = sql & " 	and p.purchasetype in ('3','5','6')"
			else
				sql = sql + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shop_user u with (nolock)"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join "&FTENDB&"db_partner.dbo.tbl_partner pp with (nolock)"
	    sql = sql & "       on m.shopid=pp.id "
		sql = sql + " left Join "&FTENDB&"[db_shop].[dbo].tbl_shop_item s with (nolock)"
		sql = sql + " 	on d.itemgubun=s.itemgubun"
		sql = sql + " 	and d.itemid=s.shopitemid"
		sql = sql + " 	and d.itemoption=s.itemoption"

		sql = sql + " where"
		sql = sql + " m.cancelyn='N'"
		sql = sql + " and d.cancelyn='N' " & sqlsearch
		sql = sql + " group by d.makerid, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.realsellprice"
		sql = sql + " order by sellcnt desc"

		'response.write  sql &"<br>"
		db3_rsget.pagesize = FPageSize
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sql, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<0 then FResultCount=0
		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof

				set FItemList(i) = new cStaticdatamart_oneitem

				FItemList(i).fIorgsellprice	= db3_rsget("Iorgsellprice")
				FItemList(i).fsellsum	= db3_rsget("sellprice")
				FItemList(i).fsuplyprice	= db3_rsget("suplyprice")
				FItemList(i).fshopbuyprice	= db3_rsget("shopbuyprice")
				FItemList(i).FMakerid = db3_rsget("makerid")
				FItemList(i).Fitemgubun	= db3_rsget("itemgubun")
				FItemList(i).Fitemid	= db3_rsget("itemid")
				FItemList(i).Fitemoption= db3_rsget("itemoption")
			    FItemList(i).Fitemname 	= db2html(db3_rsget("itemname"))
				FItemList(i).Fitemoptionname	= db2html(db3_rsget("itemoptionname"))
				FItemList(i).Frealsellprice		= db3_rsget("realsellprice")
				FItemList(i).Fsellcnt 	= db3_rsget("sellcnt")

				if Not IsNull(FItemList(i).fsellsum) then
					FTotalSum = FTotalSum + db3_rsget("sellsum")
				end if

				db3_rsget.MoveNext
				i = i + 1
			loop
		end if
		db3_rsget.close
	end sub

	'/common/offshop/maechul/statistic/statistic_brand_datamart.asp
	public Sub GetBrandSellSumList_datamart()
		dim i,sqlStr ,sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')=''"
	    end if
		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum >= 0"
		end if

		if FRectNormalOnly="on" then
			sqlsearch = sqlsearch + " and m.cancelyn='N'"
			sqlsearch = sqlsearch + " and d.cancelyn='N'"
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		if FRectOnlyShop<>"" then
			sqlsearch = sqlsearch + " and Left(m.shopid,4)<>'cafe'"
		end if

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid = '"&FRectmakerid&"'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if

		If frectoffcatecode <> "" Then
			sqlsearch = sqlsearch + " and p.offcatecode = '" + CStr(frectoffcatecode) + "' "
		End IF

		If frectoffmduserid <> "" Then
			sqlsearch = sqlsearch + " and p.offmduserid = '" + CStr(frectoffmduserid) + "' "
		End IF

		'if FRectJungSanGubun <> "" and FRectShopid<>"" then
		'	sqlsearch = sqlsearch + " and s.chargediv = " + CStr(FRectJungSanGubun)
		'end if
        
        if FRectCommCd <> "" and FRectShopid<>"" then
			sqlsearch = sqlsearch + " and d.jcomm_cd = '" + CStr(FRectCommCd) + "'"
		end if

		if (FRectBrandPurchaseType<>"") then
			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				sqlsearch = sqlsearch + " 	and p.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				sqlsearch = sqlsearch & " 	and p.purchasetype in ('3','5','6')"
			else
				sqlsearch = sqlsearch + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		sqlStr = " SELECT top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " isnull(sum(d.itemno * (d.realsellprice) ),0) as subtotal"
		sqlStr = sqlStr + " , isnull(sum(d.itemno * d.addTaxCharge),0) as addTaxChargeSum"
		sqlStr = sqlStr + " , isnull(sum(d.itemno),0) as cnt"
		sqlStr = sqlStr + " , d.makerid"
		sqlStr = sqlStr + " , isnull(sum(d.itemno * (d.suplyprice) ),0) as suplyprice"
		sqlStr = sqlStr + " , isnull(sum(d.itemno * (d.Iorgsellprice) ),0) as Iorgsellprice"
		sqlStr = sqlStr + " , isnull(sum( (d.itemno * (d.realsellprice)) - (d.itemno * (d.suplyprice)) ),0) as profit"
        sqlStr = sqlStr + " , p.purchasetype, pc.pcomm_name as purchasetypename"
		if frectdategubun = "M" then
			sqlStr = sqlStr & " ,convert(varchar(7),m.IXyyyymmdd) as IXyyyymmdd"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " ,d.jcomm_cd"
		end if

		if FRectOldData="on" then
			sqlStr = sqlStr + " from "&FTENDB&"[db_shoplog].[dbo].tbl_old_shopjumun_master m with (nolock)"
			sqlStr = sqlStr + " join "&FTENDB&"[db_shoplog].[dbo].tbl_old_shopjumun_detail d with (nolock)"
			sqlStr = sqlStr + " 	on m.orderno = d.orderno"
		else
			sqlStr = sqlStr + " from "&FTENDB&"[db_shop].[dbo].tbl_shopjumun_master m with (nolock)"
			sqlStr = sqlStr + " join "&FTENDB&"[db_shop].[dbo].tbl_shopjumun_detail d with (nolock)"
			sqlStr = sqlStr + " 	on m.orderno = d.orderno"
		end if

		sqlStr = sqlStr + " join "&FTENDB&"[db_shop].[dbo].tbl_shop_user u with (nolock)"
		sqlStr = sqlStr + " 	on m.shopid = u.userid"
		sqlStr = sqlStr + " join "&FTENDB&"[db_partner].[dbo].tbl_partner p with (nolock) on d.makerid = p.id "
		sqlStr = sqlStr & " LEFT JOIN "&FTENDB&"[db_partner].[dbo].tbl_partner_comm_code as pc with (nolock)"
		sqlStr = sqlStr & " 	on pc.pcomm_group='purchasetype' and pc.pcomm_isusing='Y' and p.purchasetype=pc.pcomm_cd"
		sqlStr = sqlStr & " left join "&FTENDB&"db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr & "       on m.shopid=pp.id "

		if FRectShopid<>"" then
			sqlStr = sqlStr + " left join "&FTENDB&"[db_shop].[dbo].tbl_shop_designer s with (nolock)"
			sqlStr = sqlStr + " 	on s.shopid='" + FRectShopid + "'"
			sqlStr = sqlStr + " 	and d.makerid=s.makerid"
		end if

		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " group by d.makerid , p.purchaseType, pc.pcomm_name"

		if frectdategubun = "M" then
			sqlStr = sqlStr & " ,convert(varchar(7),m.IXyyyymmdd)"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " ,d.jcomm_cd"
		end if

		sqlStr = sqlStr + " order by"

		if frectdategubun = "M" then
			sqlStr = sqlStr & " IXyyyymmdd desc,"
		end if

		Select Case FRectOrdertype
			Case "totalprice"
				'매출순
				sqlStr = sqlStr & " subtotal Desc"
	    	Case "gain"
	    		'수익순
	            sqlStr = sqlStr & " profit Desc"
			Case "ea"
				'수량순
				sqlStr = sqlStr & " cnt Desc, subtotal desc"
			case else
				sqlStr = sqlStr + " subtotal desc"
		end Select

		'response.write sqlStr & "<br>"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new cStaticdatamart_oneitem
				FItemList(i).fpurchasetypename  = db3_rsget("purchasetypename")
				FItemList(i).fIorgsellprice  = db3_rsget("Iorgsellprice")
				FItemList(i).FMakerid  = db3_rsget("makerid")
				FItemList(i).FCount = db3_rsget("cnt")
				FItemList(i).FSum   = db3_rsget("subtotal")
				FItemList(i).fsuplyprice  = db3_rsget("suplyprice")
				FItemList(i).fprofit  = db3_rsget("profit")
				FItemList(i).FaddTaxChargeSum  = db3_rsget("addTaxChargeSum")

				if frectdategubun = "M" then
					FItemList(i).fIXyyyymmdd = db3_rsget("IXyyyymmdd")
				end if

				if FRectShopid<>"" then
					''FItemList(i).FChargeDiv = db3_rsget("chargediv")
					FItemList(i).FjComm_cd = db3_rsget("jcomm_cd")
				end if

                    FItemList(i).FPurchasetype = db3_rsget("purchasetype")
				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
	end Sub

	'//common/offshop/maechul/statistic/statistic_daily_item_datamart.asp
	public function fStatistic_daily_item_datamart
		dim i , sql, sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')=''"
	    end if
		if frectmakerid <> "" then
			sqlsearch = sqlsearch & " and d.makerid = '"& frectmakerid &"'"
		end if

		if FRectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"& FRectshopid &"'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDate) + "'"
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDate) + "'"
			end if
		end if

		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum >= 0"
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		sql = "SELECT"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " Convert(varchar(10),m.shopregdate,121) AS regdate"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " m.IXyyyymmdd AS regdate"
		end if

		sql = sql & " ,isNull(sum(d.itemno),0) AS itemno"
		sql = sql & " ,isNull(sum( ((CASE WHEN isNULL(d.iorgsellprice,0)=0 THEN d.sellprice WHEN d.sellprice/d.iorgsellprice>1 THEN d.sellprice ELSE d.iorgsellprice END)+isNULL(d.addtaxcharge,0)) *d.itemno),0) AS Iorgsellprice"
		sql = sql & " ,sum( (isNull(d.sellprice,0)+isNULL(d.addtaxcharge,0)) *d.itemno) AS sellprice"
		sql = sql & " ,sum( (isNull(d.realsellprice,0)+isNULL(d.addtaxcharge,0)) *d.itemno) AS realsellprice"
		sql = sql & " ,sum( (isNull(d.suplyprice,0)+isNULL(d.addtaxcharge,0)) *d.itemno) AS suplyprice"
		sql = sql & " ,sum( (isNull(d.shopbuyprice,0)+isNULL(d.addtaxcharge,0)) *d.itemno) AS shopbuyprice"

		If (FRectChkShowGubun = "Y") Then
			sql = sql & " ,d.jcomm_cd "
		End If

		if FRectOldData="on" then
			sql = sql & " from "&FTENDB&"db_shoplog.[dbo].tbl_old_shopjumun_master m"
			sql = sql & " join "&FTENDB&"db_shoplog.[dbo].tbl_old_shopjumun_detail d"
			sql = sql & " 	on m.idx=d.masteridx"
		else
			sql = sql & " from "&FTENDB&"db_shop.dbo.tbl_shopjumun_master m"
			sql = sql & " join "&FTENDB&"db_shop.dbo.tbl_shopjumun_detail d"
			sql = sql & " 	on m.idx=d.masteridx"
		end if

		if (FRectPurchasetype<>"") then
		    sql = sql + " Join "&FTENDB&"db_partner.dbo.tbl_partner p"
		    sql = sql + " on d.makerid=p.id"
		    sql = sql + " and p.purchaseType="&FRectPurchasetype&""
		end if

		sql = sql & " left join "&FTENDB&"db_partner.dbo.tbl_partner pp"
	    sql = sql & "       on m.shopid=pp.id "
		sql = sql & " left join "&FTENDB&"db_shop.dbo.tbl_shop_user u"
		sql = sql & " 		on m.shopid = u.userid"

		sql = sql & " where m.cancelyn='N' and d.cancelyn='N' " & sqlsearch

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " group by Convert(varchar(10),m.shopregdate,121)"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " group by m.IXyyyymmdd"
		end If

		If (FRectChkShowGubun = "Y") Then
			sql = sql & " ,d.jcomm_cd "
		End If

		sql = sql & " order by regdate desc"
		If (FRectChkShowGubun = "Y") Then
			sql = sql & " ,d.jcomm_cd "
		End If

		'response.write sql &"<Br>"
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
				set FItemList(i) = new cStaticdatamart_oneitem

				FItemList(i).fregdate					= db3_rsget("regdate")
				FItemList(i).fitemno					= db3_rsget("itemno")
				FItemList(i).fIorgsellprice				= db3_rsget("Iorgsellprice")
				FItemList(i).fsellprice					= db3_rsget("sellprice")
				FItemList(i).frealsellprice				= db3_rsget("realsellprice")
				FItemList(i).fsuplyprice				= db3_rsget("suplyprice")
				FItemList(i).fshopbuyprice				= db3_rsget("shopbuyprice")
				FItemList(i).FMaechulProfit				= db3_rsget("realsellprice") - db3_rsget("suplyprice")
				FItemList(i).FMaechulProfitPer			= Round(((db3_rsget("realsellprice") - db3_rsget("suplyprice"))/CHKIIF(db3_rsget("realsellprice")=0,1,db3_rsget("realsellprice")))*100,2)

				If (FRectChkShowGubun = "Y") Then
					FItemList(i).Fjcomm_cd					= db3_rsget("jcomm_cd")
				End If

			db3_rsget.movenext
			i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

	'//common/offshop/maechul/statistic/statistic_shop_datamart.asp
	public function fStatistic_shop_datamart
		dim i , sql , sqlsearch, sqlsearch1

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch1 = sqlsearch1 & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch1 = sqlsearch1 & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"& FRectshopid &"'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDate) + "'"
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDate) + "'"
			end if
		end if

		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum >= 0"
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		sql = "SELECT "
		sql = sql & " count(*) as ordercnt, m.shopid, u.shopname"
		sql = sql & " ,isNull(SUM(m.totalsum),0) as totalsum"
		sql = sql & " ,isNull(SUM(m.realsum),0) as realsum"
		sql = sql & " ,isNull(SUM(m.spendmile),0) as spendmile"
		sql = sql & " ,isNull(SUM(t.iorgsellpriceSum),0) as iorgsellpriceSum"
		sql = sql & "  ,isNull(SUM(t.profit),0) as profit "
		''sql = sql & " ,isNull(SUM(t.bonuscouponprice),0) as bonuscouponprice"

		if FRectOldData="on" then
			sql = sql & " from "&FTENDB&"db_shoplog.[dbo].tbl_old_shopjumun_master m"
		else
			sql = sql & " from "&FTENDB&"db_shop.dbo.tbl_shopjumun_master m"
		end if

		sql = sql & " left join "&FTENDB&"db_shop.dbo.tbl_shop_user u"
		sql = sql & " 		on m.shopid = u.userid"
		sql = sql & " left join "&FTENDB&"db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "
		sql = sql & " left join ("
		sql = sql & " 		select m.orderno"
		sql = sql & " 		, isnull(sum((CASE WHEN isNULL(d.iorgsellprice,0)=0 THEN d.sellprice WHEN d.sellprice/d.iorgsellprice>1 THEN d.sellprice ELSE d.iorgsellprice END) *d.itemno),0) as iorgsellpriceSum" ''
        sql = sql & "       , isnull(sum( (d.itemno * (d.realsellprice)) - (d.itemno * (d.suplyprice)) ),0) as profit "
		if FRectOldData="on" then
			sql = sql & " 	from "&FTENDB&"db_shoplog.[dbo].tbl_old_shopjumun_master m"
			sql = sql & " 	join "&FTENDB&"db_shoplog.[dbo].tbl_old_shopjumun_detail d"
			sql = sql & " 		on m.idx=d.masteridx"
		else
			sql = sql & " 	from "&FTENDB&"db_shop.dbo.tbl_shopjumun_master m"
			sql = sql & " 	join "&FTENDB&"db_shop.dbo.tbl_shopjumun_detail d"
			sql = sql & " 		on m.idx=d.masteridx"
		end if

		sql = sql & " 		left join "&FTENDB&"db_shop.dbo.tbl_shop_user u"
		sql = sql & " 			on m.shopid = u.userid"
		sql = sql & " 		where m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sql = sql & " 		group by m.orderno"
		sql = sql & " ) as t"
		sql = sql & " 		on m.orderno = t.orderno"
'		sql = sql & " left join ("
'		sql = sql & " 		select m.orderno"
'		sql = sql & " 		, isnull(sum(d.realsellprice*d.itemno),0) as bonuscouponprice"
'		sql = sql & " 		FROM "&FTENDB&"[db_shop].[dbo].[tbl_shopjumun_master] as m"
'		sql = sql & " 		join "&FTENDB&"[db_shop].[dbo].[tbl_shopjumun_detail] as d"
'		sql = sql & " 			on m.idx=d.masteridx"
'		sql = sql & " 		left join "&FTENDB&"db_shop.dbo.tbl_shop_user u"
'		sql = sql & " 			on m.shopid = u.userid"
'		sql = sql & " 		where m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
'		sql = sql & " 		and ((d.itemgubun='60') or (itemgubun='90' and itemid in (32681, 34978, 35215))"
'		sql = sql & " 		group by m.orderno"
'		sql = sql & " ) as t"
'		sql = sql & " 		on m.orderno = t.orderno"
		sql = sql & " where m.cancelyn='N' " & sqlsearch & sqlsearch1
		sql = sql & " GROUP BY m.shopid, u.shopname, u.shopdiv"
		sql = sql & " ORDER BY convert(int,u.shopdiv)+10 asc, m.shopid asc"

		'response.write sql &"<Br>"
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
				set FItemList(i) = new cStaticdatamart_oneitem

				'FItemList(i).fbonuscouponprice	= 0 ''db3_rsget("bonuscouponprice")
				FItemList(i).fordercnt			= db3_rsget("ordercnt")
				FItemList(i).fshopid			= db3_rsget("shopid")
				FItemList(i).fshopname			= db3_rsget("shopname")
				FItemList(i).ftotalsum			= db3_rsget("totalsum")
				FItemList(i).frealsum			= db3_rsget("realsum")
				FItemList(i).fspendmile			= db3_rsget("spendmile")
				FItemList(i).FMaechul			= FItemList(i).frealsum + FItemList(i).fspendmile
                FItemList(i).fIorgsellprice     = db3_rsget("iorgsellpriceSum")
                FItemList(i).fprofit            = db3_rsget("profit")
			db3_rsget.movenext
			i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

	'//common/offshop/maechul/statistic/statistic_checkmethod_datamart.asp
	public function fStatistic_checkmethod_datamart
		dim i , sql , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"& FRectshopid &"'"
		end if

		if (FRectOnlyTenShop = "Y") then
			sqlsearch = sqlsearch & " and m.shopid in ('streetshop011', 'streetshop014', 'streetshop018') "
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDate) + "'"
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDate) + "'"
			end if
		end if

		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum >= 0"
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		sql = "select"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " Convert(varchar(10),m.shopregdate,121) AS regdate"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " m.IXyyyymmdd AS regdate"
		end if

		sql = sql & " ,sum(spendmile) as spendmile"
		sql = sql & " ,sum(case when spendmile<>0 then 1 else 0 end) as 'spendmilecnt'"
		sql = sql & " ,sum(TenGiftCardPaySum) as TenGiftCardPaySum"
		sql = sql & " ,sum(case when isnull(TenGiftCardMatchCode,'')<>'' then 1 else 0 end) as 'TenGiftCardPaycount'"
		sql = sql & " ,sum(cardsum) as cardsum"
		sql = sql & " ,sum(case when jumunmethod='02' then 1 else 0 end) as 'cardcnt'"
		sql = sql & " ,sum(cashsum) as cashsum"
		sql = sql & " ,sum(case when jumunmethod='01' then 1 else 0 end) as 'cashcnt'"
		sql = sql & " ,sum(giftcardPaysum) as giftcardPaysum"
		sql = sql & " ,sum(case when giftcardPaysum<>0 then 1 else 0 end) as 'giftcardPaycnt'"
		sql = sql & " ,(sum(spendmile) + sum(TenGiftCardPaySum) + sum(cardsum) + sum(cashsum) + sum(giftcardPaysum)) as selltotal"

		if FRectOldData="on" then
			sql = sql & " from "&FTENDB&"db_shoplog.[dbo].tbl_old_shopjumun_master m"
		else
			sql = sql & " from "&FTENDB&"db_shop.dbo.tbl_shopjumun_master m"
		end if

		sql = sql & " join "&FTENDB&"[db_shop].[dbo].tbl_shop_user u"
		sql = sql & " 	on m.shopid = u.userid"
		sql = sql & " left join "&FTENDB&"db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "
		sql = sql & " where m.cancelyn='N' " & sqlsearch
		sql = sql & " group by"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " 	Convert(varchar(10),m.shopregdate,121)"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " 	m.IXyyyymmdd"
		end if

		sql = sql & " ORDER BY"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " 	Convert(varchar(10),m.shopregdate,121) DESC"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " 	m.IXyyyymmdd DESC"
		end if

		'response.write sql & "<Br>"
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof

			set FItemList(i) = new cStaticdatamart_oneitem

			FItemList(i).FRegdate			= db3_rsget("regdate")
			FItemList(i).fspendmile			= db3_rsget("spendmile")
			FItemList(i).fspendmilecnt			= db3_rsget("spendmilecnt")
			FItemList(i).fTenGiftCardPaySum			= db3_rsget("TenGiftCardPaySum")
			FItemList(i).fTenGiftCardPaycount			= db3_rsget("TenGiftCardPaycount")
			FItemList(i).fcardsum			= db3_rsget("cardsum")
			FItemList(i).fcardcnt			= db3_rsget("cardcnt")
			FItemList(i).fcashsum			= db3_rsget("cashsum")
			FItemList(i).fcashcnt			= db3_rsget("cashcnt")
			FItemList(i).fgiftcardPaysum			= db3_rsget("giftcardPaysum")
			FItemList(i).fgiftcardPaycnt			= db3_rsget("giftcardPaycnt")
			FItemList(i).fselltotal			= db3_rsget("selltotal")

			db3_rsget.movenext
			i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

	'/common/offshop/maechul/statistic/statistic_month_datamart.asp
	public function fStatistic_monthlist_datamart
		dim i , sql , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"& FRectshopid &"'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDate) + "'"
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDate) + "'"
			end if
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		sql = "SELECT "

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " Convert(varchar(10),min(m.shopregdate),121) AS mindate"
			sql = sql & " , Convert(varchar(10),max(m.shopregdate),121) AS maxdate"
			sql = sql & " , Convert(varchar(7),m.shopregdate,121) AS regmonth"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " min(m.IXyyyymmdd) AS mindate"
			sql = sql & " ,max(m.IXyyyymmdd) AS maxdate"
			sql = sql & " , Convert(varchar(7),m.IXyyyymmdd,121) AS regmonth"

		end if

		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum >= 0 then 1"
		sql = sql & " end),0) as countplus"
		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum >= 0 then m.realsum + m.spendmile"
		sql = sql & " end),0) as maechulplus"
		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum < 0 then 1"
		sql = sql & " end),0) as countminus"
		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum < 0 then m.realsum + m.spendmile"
		sql = sql & " end),0) as maechulminus"
		sql = sql & " ,isnull(sum(m.realsum),0) as subtotalprice"
		sql = sql & " ,isnull(sum(m.spendmile),0) as miletotalprice"

		if FRectOldData="on" then
			sql = sql & " from "&FTENDB&"db_shoplog.[dbo].tbl_old_shopjumun_master m"
		else
			sql = sql & " from "&FTENDB&"db_shop.dbo.tbl_shopjumun_master m"
		end if

		sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shop_user u"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join "&FTENDB&"db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "
		sql = sql & " where m.cancelyn='N' " & sqlsearch
		sql = sql & " group by"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " 	Convert(varchar(7),m.shopregdate,121)"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " 	Convert(varchar(7),m.IXyyyymmdd,121)"

		end if

		sql = sql & " ORDER BY"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " 	Convert(varchar(7),m.shopregdate,121) DESC"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " 	Convert(varchar(7),m.IXyyyymmdd,121) DESC"

		end if

		'response.write sql & "<Br>"
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
				set FItemList(i) = new cStaticdatamart_oneitem

				FItemList(i).FMinDate			= db3_rsget("mindate")
				FItemList(i).FMaxDate			= db3_rsget("maxdate")
				FItemList(i).FMonth				= db3_rsget("regmonth")
				FItemList(i).FCountPlus 		= db3_rsget("countplus")
				FItemList(i).FCountMinus      	= db3_rsget("countminus")
				FItemList(i).FMaechulPlus 		= db3_rsget("maechulplus")
				FItemList(i).FMaechulMinus     	= db3_rsget("maechulminus")
				FItemList(i).FSubtotalprice     = db3_rsget("subtotalprice")
				FItemList(i).FMiletotalprice	= db3_rsget("miletotalprice")

			db3_rsget.movenext
			i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

	'/common/offshop/maechul/statistic/statistic_week_datamart.asp
	public function fStatistic_weeklist_datamart
		dim i , sql , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"& FRectshopid &"'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDate) + "'"
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDate) + "'"
			end if
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		sql = "SELECT "

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " Convert(varchar(10),min(m.shopregdate),121) AS mindate"
			sql = sql & " , Convert(varchar(10),max(m.shopregdate),121) AS maxdate"
			sql = sql & " , DATEPART(ww,m.shopregdate) as weekdt"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " min(m.IXyyyymmdd) AS mindate"
			sql = sql & " ,max(m.IXyyyymmdd) AS maxdate"
			sql = sql & " ,DATEPART(ww,m.IXyyyymmdd) AS weekdt"

		end if

		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum >= 0 then 1"
		sql = sql & " end),0) as countplus"
		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum >= 0 then m.realsum + m.spendmile"
		sql = sql & " end),0) as maechulplus"
		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum < 0 then 1"
		sql = sql & " end),0) as countminus"
		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum < 0 then m.realsum + m.spendmile"
		sql = sql & " end),0) as maechulminus"
		sql = sql & " ,isnull(sum(m.realsum),0) as subtotalprice"
		sql = sql & " ,isnull(sum(m.spendmile),0) as miletotalprice"

		if FRectOldData="on" then
			sql = sql & " from "&FTENDB&"db_shoplog.[dbo].tbl_old_shopjumun_master m"
		else
			sql = sql & " from "&FTENDB&"db_shop.dbo.tbl_shopjumun_master m"
		end if

		sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shop_user u"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join "&FTENDB&"db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "
		sql = sql & " where m.cancelyn='N' " & sqlsearch
		sql = sql & " group by"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " 	DATEPART(ww,m.shopregdate)"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " 	DATEPART(ww,m.IXyyyymmdd)"

		end if

		sql = sql & " ORDER BY"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " 	Convert(varchar(10),max(m.shopregdate),121) DESC"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " 	max(m.IXyyyymmdd) DESC"

		end if

		'response.write sql & "<Br>"
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof
			set FItemList(i) = new cStaticdatamart_oneitem

				FItemList(i).FMinDate			= db3_rsget("mindate")
				FItemList(i).FMaxDate			= db3_rsget("maxdate")
				FItemList(i).FWeek				= db3_rsget("weekdt")
				FItemList(i).FCountPlus 		= db3_rsget("countplus")
				FItemList(i).FCountMinus      	= db3_rsget("countminus")
				FItemList(i).FMaechulPlus 		= db3_rsget("maechulplus")
				FItemList(i).FMaechulMinus     	= db3_rsget("maechulminus")
				FItemList(i).FSubtotalprice     = db3_rsget("subtotalprice")
				FItemList(i).FMiletotalprice	= db3_rsget("miletotalprice")

			db3_rsget.movenext
			i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

	'//common/offshop/maechul/statistic/statistic_daily_datamart.asp
	public function fStatistic_dailylist_datamart
		dim i , sql , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"& FRectshopid &"'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDate) + "'"
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDate) + "'"
			end if
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		sql = "select"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " Convert(varchar(10),m.shopregdate,121) AS regdate"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " m.IXyyyymmdd AS regdate"
		end if

		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum >= 0 then 1"
		sql = sql & " end),0) as countplus"
		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum >= 0 then m.realsum + m.spendmile"
		sql = sql & " end),0) as maechulplus"
		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum < 0 then 1"
		sql = sql & " end),0) as countminus"
		sql = sql & " ,isnull(sum(case"
		sql = sql & " 		when m.totalsum < 0 then m.realsum + m.spendmile"
		sql = sql & " end),0) as maechulminus"
		sql = sql & " ,isnull(sum(m.realsum),0) as subtotalprice"
		sql = sql & " ,isnull(sum(m.spendmile),0) as miletotalprice"

		if FRectOldData="on" then
			sql = sql & " from "&FTENDB&"db_shoplog.[dbo].tbl_old_shopjumun_master m"
		else
			sql = sql & " from "&FTENDB&"db_shop.dbo.tbl_shopjumun_master m"
		end if

		sql = sql + " join "&FTENDB&"[db_shop].[dbo].tbl_shop_user u"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join "&FTENDB&"db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "
		sql = sql & " where m.cancelyn='N' " & sqlsearch
		sql = sql & " group by"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " 	Convert(varchar(10),m.shopregdate,121)"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " 	m.IXyyyymmdd"
		end if

		sql = sql & " ORDER BY"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " 	Convert(varchar(10),m.shopregdate,121) DESC"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " 	m.IXyyyymmdd DESC"
		end if

		'response.write sql & "<Br>"
		db3_rsget.open sql,db3_dbget,1

		FTotalCount = db3_rsget.recordcount
		FresultCount = db3_rsget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not db3_rsget.Eof Then
			Do Until db3_rsget.Eof

			set FItemList(i) = new cStaticdatamart_oneitem

			FItemList(i).FRegdate			= db3_rsget("regdate")
			FItemList(i).FCountPlus 		= db3_rsget("countplus")
			FItemList(i).FCountMinus      	= db3_rsget("countminus")
			FItemList(i).FMaechulPlus 		= db3_rsget("maechulplus")
			FItemList(i).FMaechulMinus     	= db3_rsget("maechulminus")
			FItemList(i).FSubtotalprice     = db3_rsget("subtotalprice")
			FItemList(i).FMiletotalprice	= db3_rsget("miletotalprice")

			db3_rsget.movenext
			i = i + 1
			Loop
		End If

		db3_rsget.close
	end function

end class
%>
