<%
'####################################################
' Description :  오프라인 매출 클래스
' History : 2009.04.07 서동석 생성
'			2010.05.12 한용민 수정
'####################################################

Class COffshopReportItem
	public FCateCDL
	public FCateCDM
	public FCateCDN
	public FCateName
	public FSellSum
	public FSellCnt
	public fsuplyprice
	public FMakerid
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public FItemName
	public FItemOptionName
	public Frealsellprice
	public Fselltotal
	public Fbuytotal
	public FItemNo
	public fsellprice
	public fipgosum
	public ftotsellsum
	public ftotipgocnt
	public ftotsellcnt
	public fpro
	public fremainSum
	public fcomm_cd
	public fcomm_name
	public fshopid
	public fstsum
	public fstno
	public fextbarcode
	public fIorgsellprice
	public fshopbuyprice

	public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COffshopReport
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public maxt
	public maxc
	public FRectShopID
	public FRectStartDay
	public FRectEndDay
	public FRectOldData
	public FRectCDL
	public FRectCDM
	public FRectCDN
	public FRectFromDate
	public FRectToDate
	public FRectCD1
	public FRectCD2
	public FRectDispY
	public FRectSellY
	public FRectOrdertype
	public FRectOldJumun
	public FRectOffgubun
	public FRectCD3
	public FRectyyyymm
	public frectcommcd
	public frectdatefg
	public frectmakerid
	public frectsearchgubun
	public frectcatecdnull
	public frectweekdate
	public FTotalSum
	public FRectBrandPurchaseType
	public frectBanPum
	public frectbuyergubun
	public FRectInc3pl

	'//2012-01-04 김진영 Clng에서 Ccur로 변경
	function MaxVal(a,b)
		if a = "" then a = 0
		if b = "" then b = 0

		if (CCur(a)> CCur(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

	'//admin/offshop/maechul/brandipgomaechul.asp
	public Sub getbrandipgomaechul()
		dim sqlStr , i ,sqlsearch, sqlsearch1

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectShopID <> "" then
			sqlsearch1 = sqlsearch1 & " and d.shopid='"&FRectShopID&"'"
		end if

		if FRectyyyymm <> "" then
			sqlsearch1 = sqlsearch1 & " and d.yyyymm='"&FRectyyyymm&"'"
		end if

		if frectcommcd <> "" then
			sqlsearch1 = sqlsearch1 & " and j.comm_cd='"&frectcommcd&"'"
		end if

		sqlStr = "select top " + CStr(FCurrPage*FPageSize)
		sqlStr = sqlStr & " d.shopid, i.makerid"
		sqlStr = sqlStr & " ,isnull(sum((s.logicsipgono+s.brandipgono)*i.shopitemprice),0) as ipgosum"	'총입고액
		sqlStr = sqlStr & " ,isnull(sum(s.sellno*-1*i.shopitemprice),0) as totsellsum"			'총판매액
		sqlStr = sqlStr & " ,isnull(sum(s.logicsipgono+s.brandipgono),0) as totipgocnt"			'그달 입고수량
		sqlStr = sqlStr & " ,isnull(sum(s.sellno*-1),0) as totsellcnt"					'그달 판매수량
		sqlStr = sqlStr & " ,isnull((CASE"
		sqlStr = sqlStr & " 	WHEN sum((s.logicsipgono+s.brandipgono)*i.shopitemprice)=0 THEN 1"
		sqlStr = sqlStr & " 	ELSE sum(s.sellno*-1*i.shopitemprice)/sum((s.logicsipgono+s.brandipgono)*i.shopitemprice)"
		sqlStr = sqlStr & " 	END)*100-100,0) as pro"  '판매액/입고액 100이 기준임.
		sqlStr = sqlStr & " ,isnull(sum(s.sellno*-1*i.shopitemprice)-sum((s.logicsipgono+s.brandipgono)*i.shopitemprice),0) as remainSum" '판매액-입고액
		sqlStr = sqlStr & " ,isnull(sum(m.realstockno*i.shopitemprice),0) as stsum" '월말재고액(소비가)
		sqlStr = sqlStr & " ,isnull(sum(m.realstockno),0) as stno"	'월말재고수량
		sqlStr = sqlStr & " ,d.comm_cd, j.comm_name"
		sqlStr = sqlStr & " from db_jungsan.dbo.tbl_jungsan_comm_code j"
		sqlStr = sqlStr & " join db_summary.dbo.tbl_monthly_shop_designer d"
		sqlStr = sqlStr & " 	on j.comm_cd=d.comm_cd " & sqlsearch1
		sqlStr = sqlStr & " Join db_shop.dbo.tbl_shop_item i"
		sqlStr = sqlStr & " 	on d.makerid=i.makerid"
		sqlStr = sqlStr & " join db_summary.dbo.tbl_monthly_shopstock_summary s"
		sqlStr = sqlStr & " 	on d.yyyymm=s.yyyymm"
		sqlStr = sqlStr & " 	and d.shopid=s.shopid"
		sqlStr = sqlStr & " 	and i.makerid=d.makerid"
		sqlStr = sqlStr & " 	and i.itemgubun=s.itemgubun"
		sqlStr = sqlStr & " 	and i.shopitemid=s.itemid"
		sqlStr = sqlStr & " 	and i.itemoption=s.itemoption"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "		on d.shopid=p.id "
		sqlStr = sqlStr & " left join db_summary.dbo.tbl_monthly_accumulated_shopstock_summary m" '월말재고
		sqlStr = sqlStr & " 	on s.itemgubun=m.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid=m.itemid"
		sqlStr = sqlStr & " 	and s.itemoption=m.itemoption"
		sqlStr = sqlStr & " 	and s.yyyymm=m.yyyymm"
		sqlStr = sqlStr & " 	and s.shopid=m.shopid"
		sqlStr = sqlStr & " 	and m.realstockno>0"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " group by d.shopid, i.makerid, d.comm_cd, j.comm_name"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		ftotalcount = rsget.recordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new COffshopReportItem

			FItemList(i).fshopid       = rsget("shopid")
			FItemList(i).fmakerid       = rsget("makerid")
			FItemList(i).fipgosum       = rsget("ipgosum")
			FItemList(i).ftotsellsum       = rsget("totsellsum")
			FItemList(i).ftotipgocnt       = rsget("totipgocnt")
			FItemList(i).ftotsellcnt       = rsget("totsellcnt")
			FItemList(i).fpro       = rsget("pro")
			FItemList(i).fremainSum       = rsget("remainSum")
			FItemList(i).fcomm_cd       = rsget("comm_cd")
			FItemList(i).fcomm_name       = rsget("comm_name")
			FItemList(i).fstsum       = rsget("stsum")
			FItemList(i).fstno       = rsget("stno")

			rsget.movenext
			i=i+1
			loop
		rsget.Close
	end Sub

	'//admin/offshop/offshop_categorybestseller.asp
	public Sub SearchCategoryBestseller()
		dim sqlStr , i ,sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and s.makerid='" + CStr(FRectmakerid) + "'"
		end if
		if (FRectOffgubun<>"") then
		    if (FRectOffgubun="1") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('1','2')"
		    elseif (FRectOffgubun="3") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('3','4')"
		    elseif (FRectOffgubun="5") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('5','6')"
		    elseif (FRectOffgubun="7") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('7','8')"
		    elseif (FRectOffgubun="9") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('9')"
		    end if
		end if
		if FRectShopid<>"" then
			sqlsearch = sqlsearch & " and m.shopid='" + CStr(FRectShopid) + "'" + vbcrlf
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectToDate) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectToDate) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.IXyyyymmdd) = "&frectweekdate&""
			end if
		else
			if FRectFromDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectFromDate) + "'"
			end if
			if FRectToDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectToDate) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if
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

		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr & " sum(d.itemno) as itemno ,sum(d.itemno*d.realsellprice) as sellsum"
		sqlStr = sqlStr & " ,sum(d.itemno*d.suplyprice) as buysum, d.makerid"
		sqlStr = sqlStr & " ,sum(d.itemno*d.sellprice) as totsellsum"

		'/상품기준
		if searchgubun = "I" then
			sqlStr = sqlStr & " , d.itemid, d.sellprice , d.itemname, d.itemoptionname , d.itemoption,d.itemgubun , s.extbarcode"
		end if

		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m"
			sqlStr = sqlStr + " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d"
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
			sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d"
		end if

		sqlStr = sqlStr & " 	on m.idx = d.masteridx "
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr + " 	on m.shopid = u.userid"
		sqlStr = sqlStr & " left Join [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr & " 	on d.itemgubun=s.itemgubun "
		sqlStr = sqlStr & " 	and d.itemid=s.shopitemid "
		sqlStr = sqlStr & " 	and d.itemoption=s.itemoption"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on m.shopid=p.id "
		sqlStr = sqlStr & " where "
		sqlStr = sqlStr & " m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " group by d.makerid"

		'/상품기준
		if searchgubun = "I" then
			sqlStr = sqlStr & " , d.itemid, d.sellprice, d.itemname, d.itemoptionname, d.itemoption,d.itemgubun, s.extbarcode"
		end if

		'/상품기준
		if searchgubun = "I" then
			Select Case FRectOrdertype
				Case "totalprice"
					'매출순
					sqlStr = sqlStr & " order by sellsum Desc"
		    	Case "gain"
		    		'수익순
		            sqlStr = sqlStr & " order by sum(d.itemno*(d.sellprice-d.suplyprice)) Desc"
				Case "unitCost"
					'객단가순
					sqlStr = sqlStr & " order by d.sellprice Desc"
				Case Else
					'수량순
					sqlStr = sqlStr & " order by itemno Desc, sellsum desc"
			end Select

		'/브랜드기준
		elseif searchgubun = "B" then
			Select Case FRectOrdertype
				Case "totalprice"
					'매출순
					sqlStr = sqlStr & " order by sum(d.itemno*d.realsellprice) Desc"
		    	Case "gain"
		    		'수익순
		            sqlStr = sqlStr & " order by sum(d.itemno*d.realsellprice)-sum(d.itemno*d.suplyprice) Desc"
				Case Else
					'수량순
					sqlStr = sqlStr & " order by sum(d.itemno) Desc, sum(d.itemno*d.realsellprice) desc"
			end Select
		end if

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		ftotalcount = rsget.recordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
			set FItemList(i) = new COffshopReportItem

			FItemList(i).ftotsellsum       = rsget("totsellsum")
			FItemList(i).Fselltotal       = rsget("sellsum")
			FItemList(i).Fbuytotal       = rsget("buysum")
			FItemList(i).FItemNo       = rsget("itemno")
			FItemList(i).FMakerid		= rsget("makerid")

			'/상품기준
			if searchgubun = "I" then
				FItemList(i).fextbarcode		 = rsget("extbarcode")
				FItemList(i).fitemoption	= rsget("itemoption")
				FItemList(i).fitemgubun	= rsget("itemgubun")
				FItemList(i).FItemID       = rsget("itemid")
				FItemList(i).fsellprice       = rsget("sellprice")
				FItemList(i).FItemName     = db2html(rsget("itemname"))
				FItemList(i).fitemoptionname= db2html(rsget("itemoptionname"))
			end if

			rsget.movenext
			i=i+1
			loop
		rsget.Close
	end Sub

	'//admin/offshop/offshop_categorysellsum.asp
	public sub SearchCategorySellrePort()
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
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum > 0"
		end if

		if FRectShopID<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopID + "'" + vbcrlf
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
		sql = sql + " ,isNull(sum( (d.realsellprice+d.addtaxcharge) *d.itemno),0) as sumtotal"
		sql = sql + " ,isNull(sum( (d.suplyprice+d.addtaxcharge) *d.itemno),0) as suplyprice"
		sql = sql & " ,isNull(sum( (d.Iorgsellprice+d.addtaxcharge) *d.itemno),0) AS Iorgsellprice"
		sql = sql & " ,isNull(sum( (d.sellprice+d.addtaxcharge) *d.itemno),0) AS sellprice"
		sql = sql & " ,isNull(sum( (d.shopbuyprice+d.addtaxcharge) *d.itemno),0) AS shopbuyprice"
		sql = sql + " , s.catecdl,cl.code_nm"

		if FRectOldData="on" then
			sql = sql + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m with (nolock)"
			sql = sql + " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		else
			sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m with (nolock)"
			sql = sql + " join [db_shop].[dbo].tbl_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		end if

		if (FRectBrandPurchaseType<>"") then
		    sql = sql + " Join db_partner.dbo.tbl_partner p with (nolock)"
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

		sql = sql + " join [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sql = sql + " 	on d.itemgubun=s.itemgubun"
		sql = sql + " 	and d.itemid=s.shopitemid"
		sql = sql + " 	and d.itemoption=s.itemoption"
		sql = sql + " join [db_shop].[dbo].tbl_shop_user u with (nolock)"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql + " left join [db_item].[dbo].tbl_Cate_large cl with (nolock)" + vbcrlf
		sql = sql + " 	on s.catecdl=cl.code_large" + vbcrlf
		sql = sql & " left join db_partner.dbo.tbl_partner pp with (nolock)"
	    sql = sql & "       on m.shopid=pp.id "
		sql = sql + " where 1=1 " & sqlsearch
		sql = sql + " group by s.catecdl,cl.code_nm" + vbcrlf
		sql = sql + " order by s.catecdl"

		'response.write sql &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)

		do until rsget.eof

			set FItemList(i) = new COffshopReportItem

			FItemList(i).fIorgsellprice	= rsget("Iorgsellprice")
			FItemList(i).fsellprice	= rsget("sellprice")
			FItemList(i).fsuplyprice	= rsget("suplyprice")
			FItemList(i).fshopbuyprice	= rsget("shopbuyprice")
			FItemList(i).FCateCDL	= rsget("catecdl")
		    FItemList(i).FCateName 	= db2html(rsget("code_nm"))
			FItemList(i).Fsellsum	= rsget("sumtotal")
			FItemList(i).Fsellcnt 	= rsget("sellcnt")

			if IsNULL(FItemList(i).FCateName) then
				if IsNULL(FItemList(i).FCateCDL) then FItemList(i).FCateCDL=""
				FItemList(i).FCateName = "미지정"
			end if

			if Not IsNull(FItemList(i).Fsellsum) then
				maxt = MaxVal(maxt,FItemList(i).Fsellsum)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
			end if

			if Not IsNull(FItemList(i).fsellsum) then
				FTotalSum = FTotalSum + rsget("sumtotal")
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	'//admin/offshop/offshop_categorysellsum.asp
	public sub SearchCategorySellrePort2()
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
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum > 0"
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
		sql = sql + " ,isNull(sum( (d.realsellprice+d.addtaxcharge) *d.itemno),0) as sumtotal"
		sql = sql + " ,isNull(sum( (d.suplyprice+d.addtaxcharge) *d.itemno),0) as suplyprice"
		sql = sql & " ,isNull(sum( (d.Iorgsellprice+d.addtaxcharge) *d.itemno),0) AS Iorgsellprice"
		sql = sql & " ,isNull(sum( (d.sellprice+d.addtaxcharge) *d.itemno),0) AS sellprice"
		sql = sql & " ,isNull(sum( (d.shopbuyprice+d.addtaxcharge) *d.itemno),0) AS shopbuyprice"
		sql = sql + " , s.catecdm,cm.code_nm"

		if FRectOldData="on" then
			sql = sql + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m with (nolock)"
			sql = sql + " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		else
			sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m with (nolock)"
			sql = sql + " join [db_shop].[dbo].tbl_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		end if

		if (FRectBrandPurchaseType<>"") then
		    sql = sql + " Join db_partner.dbo.tbl_partner p with (nolock)"
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

		sql = sql + " join [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sql = sql + " 	on d.itemgubun=s.itemgubun"
		sql = sql + " 	and d.itemid=s.shopitemid"
		sql = sql + " 	and d.itemoption=s.itemoption"
		sql = sql + " join [db_shop].[dbo].tbl_shop_user u with (nolock)"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql + " left join [db_item].[dbo].tbl_cate_mid cm with (nolock)"
		sql = sql + " 	on s.catecdl=cm.code_large and s.catecdm=cm.code_mid"
		sql = sql & " left join db_partner.dbo.tbl_partner pp with (nolock)"
	    sql = sql & "       on m.shopid=pp.id "
		sql = sql + " where 1=1 "
		sql = sql + " and s.catecdl='" + FRectCDL + "' " & sqlsearch
		sql = sql + " group by s.catecdm,cm.code_nm"
		sql = sql + " order by s.catecdm"

		'response.write sql &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)
		do until rsget.eof
			set FItemList(i) = new COffshopReportItem

			FItemList(i).fIorgsellprice	= rsget("Iorgsellprice")
			FItemList(i).fsellprice	= rsget("sellprice")
			FItemList(i).fsuplyprice	= rsget("suplyprice")
			FItemList(i).fshopbuyprice	= rsget("shopbuyprice")
			FItemList(i).FCateCDL	= FRectCDL
			FItemList(i).FCateCDM	= rsget("catecdm")
		    FItemList(i).FCateName 	= db2html(rsget("code_nm"))
			FItemList(i).Fsellsum	= rsget("sumtotal")
			FItemList(i).Fsellcnt 	= rsget("sellcnt")

			if IsNULL(FItemList(i).FCateName) then
				 FItemList(i).FCateName = "미지정"
			end if

			if Not IsNull(FItemList(i).Fsellsum) then
				maxt = MaxVal(maxt,FItemList(i).Fsellsum)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
			end if

			if Not IsNull(FItemList(i).fsellsum) then
				FTotalSum = FTotalSum + rsget("sumtotal")
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	'/admin/offshop/offshop_categorysellsum.asp
	public sub SearchCategorySellrePort3()
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
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum > 0"
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
		sql = sql + " ,isNull(sum( (d.realsellprice+d.addtaxcharge) *d.itemno),0) as sumtotal"
		sql = sql + " ,isNull(sum( (d.suplyprice+d.addtaxcharge) *d.itemno),0) as suplyprice"
		sql = sql & " ,isNull(sum( (d.Iorgsellprice+d.addtaxcharge) *d.itemno),0) AS Iorgsellprice"
		sql = sql & " ,isNull(sum( (d.sellprice+d.addtaxcharge) *d.itemno),0) AS sellprice"
		sql = sql & " ,isNull(sum( (d.shopbuyprice+d.addtaxcharge) *d.itemno),0) AS shopbuyprice"
		sql = sql + " , s.catecdn,cs.code_nm"

		if FRectOldData="on" then
			sql = sql + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m with (nolock)"
			sql = sql + " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		else
			sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m with (nolock)"
			sql = sql + " join [db_shop].[dbo].tbl_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.orderno = d.orderno"
			sql = sql + " 	and m.cancelyn='N'"
			sql = sql + " 	and d.cancelyn='N'"
		end if

		if (FRectBrandPurchaseType<>"") then
		    sql = sql + " Join db_partner.dbo.tbl_partner p with (nolock)"
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

		sql = sql + " join [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sql = sql + " 	on d.itemgubun=s.itemgubun"
		sql = sql + " 	and d.itemid=s.shopitemid"
		sql = sql + " 	and d.itemoption=s.itemoption"
		sql = sql + " join [db_shop].[dbo].tbl_shop_user u with (nolock)"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql + " left join [db_item].[dbo].tbl_cate_small cs with (nolock)"
		sql = sql + " 	on s.catecdl=cs.code_large and s.catecdm=cs.code_mid and s.catecdn=cs.code_small"
		sql = sql & " left join db_partner.dbo.tbl_partner pp with (nolock)"
	    sql = sql & "       on m.shopid=pp.id "
		sql = sql + " where 1=1 "
		sql = sql + " and s.catecdl='" + FRectCDL + "'"
		sql = sql + " and s.catecdm='" + FRectCDM + "' " & sqlsearch
		sql = sql + " group by s.catecdn,cs.code_nm"
		sql = sql + " order by s.catecdn"

		'response.write sql &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)
		do until rsget.eof
			set FItemList(i) = new COffshopReportItem

			FItemList(i).fIorgsellprice	= rsget("Iorgsellprice")
			FItemList(i).fsellprice	= rsget("sellprice")
			FItemList(i).fsuplyprice	= rsget("suplyprice")
			FItemList(i).fshopbuyprice	= rsget("shopbuyprice")
			FItemList(i).FCateCDL	= FRectCDL
			FItemList(i).FCateCDM	= FRectCDM
			FItemList(i).FCateCDN	= rsget("catecdn")
		    FItemList(i).FCateName 	= db2html(rsget("code_nm"))
			FItemList(i).Fsellsum	= rsget("sumtotal")
			FItemList(i).Fsellcnt 	= rsget("sellcnt")

			if IsNULL(FItemList(i).FCateName) then FItemList(i).FCateName = "미지정"

			if Not IsNull(FItemList(i).Fsellsum) then
				maxt = MaxVal(maxt,FItemList(i).Fsellsum)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
			end if

			if Not IsNull(FItemList(i).fsellsum) then
				FTotalSum = FTotalSum + rsget("sumtotal")
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	'//admin/offshop/offshop_categorysellsum.asp
	public sub SearchCategorySellItems()
		Dim sql, i ,sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')=''"
	    end if
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum > 0"
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
		sql = sql + " ,isNull(sum( (d.realsellprice+d.addtaxcharge) *d.itemno),0) as sellsum"
		sql = sql + " ,isNull(sum( (d.suplyprice+d.addtaxcharge) *d.itemno),0) as suplyprice"
		sql = sql & " ,isNull(sum( (d.Iorgsellprice+d.addtaxcharge) *d.itemno),0) AS Iorgsellprice"
		sql = sql & " ,isNull(sum( (d.sellprice+d.addtaxcharge) *d.itemno),0) AS sellprice"
		sql = sql & " ,isNull(sum( (d.shopbuyprice+d.addtaxcharge) *d.itemno),0) AS shopbuyprice"

		if FRectOldData="on" then
			sql = sql + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m with (nolock)"
			sql = sql + " Join [db_shoplog].[dbo].tbl_old_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.idx = d.masteridx"
		else
			sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m with (nolock)"
			sql = sql + " Join [db_shop].[dbo].tbl_shopjumun_detail d with (nolock)"
			sql = sql + " 	on m.idx = d.masteridx"
		end if

		if (FRectBrandPurchaseType<>"") then
		    sql = sql + " Join db_partner.dbo.tbl_partner p with (nolock)"
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

		sql = sql + " join [db_shop].[dbo].tbl_shop_user u with (nolock)"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql + " left Join  [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sql = sql + " 	on d.itemgubun=s.itemgubun"
		sql = sql + " 	and d.itemid=s.shopitemid"
		sql = sql + " 	and d.itemoption=s.itemoption"
		sql = sql & " left join db_partner.dbo.tbl_partner pp with (nolock)"
	    sql = sql & "       on m.shopid=pp.id "
		sql = sql + " where"
		sql = sql + " m.cancelyn='N'"
		sql = sql + " and d.cancelyn='N' " & sqlsearch
		sql = sql + " group by d.makerid, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.realsellprice"
		sql = sql + " order by sellcnt desc"
 
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<0 then FResultCount=0
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof

				set FItemList(i) = new COffshopReportItem

				FItemList(i).fIorgsellprice	= rsget("Iorgsellprice")
				FItemList(i).fsellsum	= rsget("sellsum")
				FItemList(i).fsuplyprice	= rsget("suplyprice")
				FItemList(i).fshopbuyprice	= rsget("shopbuyprice")
				FItemList(i).FMakerid = rsget("makerid")
				FItemList(i).Fitemgubun	= rsget("itemgubun")
				FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).Fitemoption= rsget("itemoption")
			    FItemList(i).Fitemname 	= db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname	= db2html(rsget("itemoptionname"))
				FItemList(i).Frealsellprice		= rsget("realsellprice")
				FItemList(i).Fsellcnt 	= rsget("sellcnt")

				if Not IsNull(FItemList(i).fsellsum) then
					FTotalSum = FTotalSum + rsget("sellsum")
				end if

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		redim  FCountList(0)
		FCurrPage =1
		FPageSize = 20
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