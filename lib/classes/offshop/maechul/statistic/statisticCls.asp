<%
'###########################################################
' Description :  오프라인 매출통계 클래스
' History : 2012.10.04 한용민 생성
'###########################################################

class cStatic_oneitem
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
	public fextPaysum
	public fextPaycnt
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
end class

class cStatic_list
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
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
	public FRectInc3pl
	
	'//common/offshop/maechul/statistic/statistic_daily_item.asp
	public function fStatistic_daily_item
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
		sql = sql & " ,isNull(sum( (d.sellprice+isNULL(d.addtaxcharge,0)) *d.itemno),0) AS sellprice"
		sql = sql & " ,isNull(sum( (d.realsellprice+isNULL(d.addtaxcharge,0)) *d.itemno),0) AS realsellprice"
		sql = sql & " ,isNull(sum( (d.suplyprice+isNULL(d.addtaxcharge,0)) *d.itemno),0) AS suplyprice"
		sql = sql & " ,isNull(sum( (d.shopbuyprice+isNULL(d.addtaxcharge,0)) *d.itemno),0) AS shopbuyprice"
		sql = sql & " FROM db_shop.dbo.tbl_shopjumun_master as m"
		sql = sql & " join db_shop.dbo.tbl_shopjumun_detail as d"
		sql = sql & " 		on m.idx=d.masteridx"

		if (FRectPurchasetype<>"") then
		    sql = sql + " Join db_partner.dbo.tbl_partner p"
		    sql = sql + " on d.makerid=p.id"
		    sql = sql + " and p.purchaseType="&FRectPurchasetype&""
		end if 

		sql = sql & " left join db_partner.dbo.tbl_partner pp"
	    sql = sql & "       on m.shopid=pp.id "
		sql = sql & " left join db_shop.dbo.tbl_shop_user u"
		sql = sql & " 		on m.shopid = u.userid"

		sql = sql & " where m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		
		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " group by Convert(varchar(10),m.shopregdate,121)"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " group by m.IXyyyymmdd"
		end if
		
		sql = sql & " order by regdate desc"
		
		'response.write sql &"<Br>"
		rsget.open sql,dbget,1
	
		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount
		
		redim FItemList(FTotalCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cStatic_oneitem

				FItemList(i).fregdate					= rsget("regdate")
				FItemList(i).fitemno					= rsget("itemno")
				FItemList(i).fIorgsellprice				= rsget("Iorgsellprice")
				FItemList(i).fsellprice					= rsget("sellprice")
				FItemList(i).frealsellprice				= rsget("realsellprice")
				FItemList(i).fsuplyprice				= rsget("suplyprice")
				FItemList(i).fshopbuyprice				= rsget("shopbuyprice")
				FItemList(i).FMaechulProfit				= rsget("realsellprice") - rsget("suplyprice")
				FItemList(i).FMaechulProfitPer			= Round(((rsget("realsellprice") - rsget("suplyprice"))/CHKIIF(rsget("realsellprice")=0,1,rsget("realsellprice")))*100,2)
	
			rsget.movenext
			i = i + 1
			Loop
		End If
			
		rsget.close
	end function

	'//common/offshop/maechul/statistic/statistic_shop.asp
	public function fStatistic_shop
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
		''sql = sql & " ,isNull(SUM(t.bonuscouponprice),0) as bonuscouponprice"
		sql = sql & " FROM [db_shop].[dbo].[tbl_shopjumun_master] as m"
		sql = sql & " left join db_shop.dbo.tbl_shop_user u"
		sql = sql & " 		on m.shopid = u.userid"
		sql = sql & " left join db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "		
		sql = sql & " left join ("
		sql = sql & " 		select m.orderno"
		sql = sql & " 		, isnull(sum((CASE WHEN isNULL(d.iorgsellprice,0)=0 THEN d.sellprice WHEN d.sellprice/d.iorgsellprice>1 THEN d.sellprice ELSE d.iorgsellprice END) *d.itemno),0) as iorgsellpriceSum" ''
		sql = sql & " 		FROM [db_shop].[dbo].[tbl_shopjumun_master] as m"
		sql = sql & " 		join [db_shop].[dbo].[tbl_shopjumun_detail] as d"
		sql = sql & " 			on m.idx=d.masteridx"
		sql = sql & " 		left join db_shop.dbo.tbl_shop_user u"
		sql = sql & " 			on m.shopid = u.userid"				
		sql = sql & " 		where m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sql = sql & " 		group by m.orderno"
		sql = sql & " ) as t"
		sql = sql & " 		on m.orderno = t.orderno"	
'		sql = sql & " left join ("
'		sql = sql & " 		select m.orderno"
'		sql = sql & " 		, isnull(sum(d.realsellprice*d.itemno),0) as bonuscouponprice"
'		sql = sql & " 		FROM [db_shop].[dbo].[tbl_shopjumun_master] as m"
'		sql = sql & " 		join [db_shop].[dbo].[tbl_shopjumun_detail] as d"
'		sql = sql & " 			on m.idx=d.masteridx"
'		sql = sql & " 		left join db_shop.dbo.tbl_shop_user u"
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
		rsget.open sql,dbget,1
		
		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount
		
		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cStatic_oneitem
				
				'FItemList(i).fbonuscouponprice	= 0 ''rsget("bonuscouponprice")
				FItemList(i).fordercnt			= rsget("ordercnt")
				FItemList(i).fshopid			= rsget("shopid")
				FItemList(i).fshopname			= rsget("shopname")
				FItemList(i).ftotalsum			= rsget("totalsum")
				FItemList(i).frealsum			= rsget("realsum")
				FItemList(i).fspendmile			= rsget("spendmile")
				FItemList(i).FMaechul			= FItemList(i).frealsum + FItemList(i).fspendmile
                FItemList(i).fIorgsellprice     = rsget("iorgsellpriceSum")
                
			rsget.movenext
			i = i + 1
			Loop
		End If
			
		rsget.close
	end function
	

	'//common/offshop/maechul/statistic/statistic_checkmethod.asp
	public function fStatistic_checkmethod
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
		sql = sql & " ,isNull(sum(extPaysum),0) as extPaysum"
		sql = sql & " ,sum(case when extPaysum<>0 then 1 else 0 end) as 'extPaycnt'"
		sql = sql & " ,(sum(spendmile) + sum(TenGiftCardPaySum) + sum(cardsum) + sum(cashsum) + sum(giftcardPaysum) + isNull(sum(extPaysum),0)) as selltotal"
		sql = sql & " from db_shop.dbo.tbl_shopjumun_master m"
		sql = sql & " join [db_shop].[dbo].tbl_shop_user u"
		sql = sql & " 	on m.shopid = u.userid"
		sql = sql & " left join db_partner.dbo.tbl_partner p"
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
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount
		
		redim FItemList(FTotalCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				
			set FItemList(i) = new cStatic_oneitem

			FItemList(i).FRegdate			= rsget("regdate")
			FItemList(i).fspendmile			= rsget("spendmile")
			FItemList(i).fspendmilecnt			= rsget("spendmilecnt")
			FItemList(i).fTenGiftCardPaySum			= rsget("TenGiftCardPaySum")
			FItemList(i).fTenGiftCardPaycount			= rsget("TenGiftCardPaycount")
			FItemList(i).fcardsum			= rsget("cardsum")
			FItemList(i).fcardcnt			= rsget("cardcnt")
			FItemList(i).fcashsum			= rsget("cashsum")
			FItemList(i).fcashcnt			= rsget("cashcnt")
			FItemList(i).fgiftcardPaysum			= rsget("giftcardPaysum")
			FItemList(i).fgiftcardPaycnt			= rsget("giftcardPaycnt")
			FItemList(i).fextPaysum			= rsget("extPaysum")
			FItemList(i).fextPaycnt			= rsget("extPaycnt")
			FItemList(i).fselltotal			= rsget("selltotal")
		
			rsget.movenext
			i = i + 1
			Loop
		End If
			
		rsget.close
	end function
	
	'/common/offshop/maechul/statistic/statistic_month.asp
	public function fStatistic_monthlist
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
		sql = sql & " from db_shop.dbo.tbl_shopjumun_master m"
		sql = sql + " join [db_shop].[dbo].tbl_shop_user u"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join db_partner.dbo.tbl_partner p"
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
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount
		
		redim FItemList(FTotalCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cStatic_oneitem
			
				FItemList(i).FMinDate			= rsget("mindate")
				FItemList(i).FMaxDate			= rsget("maxdate")
				FItemList(i).FMonth				= rsget("regmonth")
				FItemList(i).FCountPlus 		= rsget("countplus")
				FItemList(i).FCountMinus      	= rsget("countminus")
				FItemList(i).FMaechulPlus 		= rsget("maechulplus")
				FItemList(i).FMaechulMinus     	= rsget("maechulminus")
				FItemList(i).FSubtotalprice     = rsget("subtotalprice")
				FItemList(i).FMiletotalprice	= rsget("miletotalprice")

			rsget.movenext
			i = i + 1
			Loop
		End If
			
		rsget.close
	end function
	
	'/common/offshop/maechul/statistic/statistic_week.asp
	public function fStatistic_weeklist
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
		sql = sql & " from db_shop.dbo.tbl_shopjumun_master m"
		sql = sql & " join [db_shop].[dbo].tbl_shop_user u"
		sql = sql & " 	on m.shopid = u.userid"
		sql = sql & " left join db_partner.dbo.tbl_partner p"
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
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount
		
		redim FItemList(FTotalCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
			set FItemList(i) = new cStatic_oneitem
			
				FItemList(i).FMinDate			= rsget("mindate")
				FItemList(i).FMaxDate			= rsget("maxdate")
				FItemList(i).FWeek				= rsget("weekdt")
				FItemList(i).FCountPlus 		= rsget("countplus")
				FItemList(i).FCountMinus      	= rsget("countminus")
				FItemList(i).FMaechulPlus 		= rsget("maechulplus")
				FItemList(i).FMaechulMinus     	= rsget("maechulminus")
				FItemList(i).FSubtotalprice     = rsget("subtotalprice")
				FItemList(i).FMiletotalprice	= rsget("miletotalprice")

			rsget.movenext
			i = i + 1
			Loop
		End If
			
		rsget.close
	end function
	
	'//common/offshop/maechul/statistic/statistic_daily.asp
	public function fStatistic_dailylist
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
		sql = sql & " from db_shop.dbo.tbl_shopjumun_master m"
		sql = sql + " join [db_shop].[dbo].tbl_shop_user u"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join db_partner.dbo.tbl_partner p"
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
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount
		
		redim FItemList(FTotalCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				
			set FItemList(i) = new cStatic_oneitem
			
			FItemList(i).FRegdate			= rsget("regdate")
			FItemList(i).FCountPlus 		= rsget("countplus")
			FItemList(i).FCountMinus      	= rsget("countminus")
			FItemList(i).FMaechulPlus 		= rsget("maechulplus")
			FItemList(i).FMaechulMinus     	= rsget("maechulminus")
			FItemList(i).FSubtotalprice     = rsget("subtotalprice")
			FItemList(i).FMiletotalprice	= rsget("miletotalprice")
	
			rsget.movenext
			i = i + 1
			Loop
		End If
			
		rsget.close
	end function
	
end class
%>