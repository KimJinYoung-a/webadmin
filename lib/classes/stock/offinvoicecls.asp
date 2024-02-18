<%

Sub drawSelectBoxPriceUnit(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value=''>선택</option><%
   query1 = " select currencyunit, currencychar from db_shop.dbo.tbl_shop_exchangeRate "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("currencyunit")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("currencyunit")&"' "&tmp_str&">" + rsget("currencyunit") + " [" + rsget("currencychar") + "]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Class COffInvoiceMasterItem
	public Fidx
	public Fshopid
	public Fshopname
	public Fjungsanidx
	public Fworkidx
	public Finvoiceno
	public Finvoicedate
	public Fdelivermethod		'운송방법(EMS/항공/해운)
	public Fexportmethod		'수출방법(FOB:업체운임부담/C&F:텐텐운임부담)
	public Fjungsantype			'정산시기(TT:선정산/LC:후정산)
	public Fpriceunit			'작성화폐
	public Fpriceunitstring
	public Fexchangerate		'환율
	public Ftotalboxno
	public Ftotalboxprice			'운임(원)			'// deprecated
	public Ftotalgoodsprice			'상품금액(원)		'// deprecated
	public Ftotalprice				'총액(환산)			'// deprecated
	public Fexportdeclarefilename	'수출허가증
	public Fexportdeclarefilename2	'수출허가증
	public Fexportdeclarefilename3	'수출허가증
	public Frealfilename
	public Frealfilename2
	public Frealfilename3
	public Fexporteraddr			'수출업자주소
	public Friskmesseraddr			'수입업자
	public Fnotifyaddr				'통지주소
	public Fportname				'출발지항구
	public Fdestinationname			'도착국가명
	public Fcarriername				'선박이름
	public Fcarrierdate				'선적일
	public Fgoodscomment1
	public Fgoodscomment2
	public Flccomment				'신용장
	public Flcbank
	public Fcomment

	public Fstatecd

	public Freguserid
	public Flastupdate
	public Fregdate

	public Freportno
	public Freportno2
	public Freportno3

	public Freportdate
	public Freportpriceunit
	public Freportexchangerate
	public Freportforeigntotalprice
	public Freporttotalprice

	public FtotalGoodsPriceWon			'// 원화
	public FtotalDeliverPriceWon
	public FtotalPriceWon

	public FtotalGoodsPriceForeign		'// 외화
	public FtotalDeliverPriceForeign
	public FtotalPriceForeign

	public Floginsite
	public Fcurrencyunit


	public function GetDeliverMethodName()
		if Fdelivermethod = "E" then
			GetDeliverMethodName = "EMS"
		elseif Fdelivermethod = "D" then
			GetDeliverMethodName = "DHL"
		elseif Fdelivermethod = "F" then
			GetDeliverMethodName = "항공"
		elseif Fdelivermethod = "S" then
			GetDeliverMethodName = "해운"
		elseif Fdelivermethod = "P" then
			GetDeliverMethodName = "국제소포(선편)"
		elseif Fdelivermethod = "T" then
			GetDeliverMethodName = "국내택배"
		else
			GetDeliverMethodName = Fdelivermethod
		end if
	end function

	public function GetStateCDName()
		if Fstatecd = "1" then
			GetStateCDName = "작성중"
		elseif Fstatecd = "7" then
			GetStateCDName = "작성완료"
		else
			GetStateCDName = Fstatecd
		end if
	end function

	public function GetExportMethodName()
		Select Case Fexportmethod
			Case "F"
				GetExportMethodName = "FOB"
			Case "C"
				GetExportMethodName = "C&F"
			Case "W"
				GetExportMethodName = "EXW"
			Case "A"
				GetExportMethodName = "FCA"
			Case Else
				GetExportMethodName = Fexportmethod
		End Select
	end function

	public function GetJungsanTypeName()
		if Fjungsantype = "T" then
			GetJungsanTypeName = "TT"
		elseif Fjungsantype = "L" then
			GetJungsanTypeName = "LC"
		else
			GetJungsanTypeName = Fjungsantype
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COffInvoiceDetailItem
	public Fidx
	public Fmasteridx
	public Fcartonboxno
	public Fgoodscomment
	public Fnweight
	public Fgweight
	public FemsPrice
	public Flastupdate
	public Fregdate

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COffInvoiceProductDetailItem
	public Fidx
	public Fmasteridx
	public Forderno
	public Fgoodscomment
	public Ftotalboxno
	public Fpriceperbox
	public Ftotalprice
	public Flastupdate
	public Fregdate

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COffInvoice
	public FItemList()
	public FOneItem

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectShopid
	public FRectStateCD

	public FRectMasterIdx
	public FRectDetailIdx
	public FRectLoginsite

	public FRectJungsanidx
	public FRectWorkidx

	public FRectReportDate
	public FRectReportNo
	public FRectExcNoReport

	public FRectDateFlag
	public FRectFromDate
	public FRectToDate

	public FRectbaljucode

	public FinvoiceNo
	public Finvoicedate
	public Fbaljucode
	public Fbaljuid
	public Fbeasongdate
	public Fregdate
	public Fpriceunit
	public FtotalGoodsPriceWon
	public FtotalDeliverPriceWon
 	public FtotalPriceWon
	public FtotalGoodsPriceForeign
	public FtotalDeliverPriceForeign
	public FtotalPriceForeign
	public FfreightTerm
	public FopenState
	public FshippingAddress
	public FinvoiceAddress
	public Fcurrencychar
	public Fcurrencyunit
	public FcountryLangCD
	public Fcomment
	public FtplGubun

	public Sub GetMasterList()

		dim i, sqlStr, sqlFromWhere
		dim tmpstr

		'======================================================================
		sqlFromWhere = " from "

		sqlFromWhere = sqlFromWhere + " 	[db_storage].[dbo].tbl_offline_invoice_master m "
		sqlFromWhere = sqlFromWhere + " 	left join [db_shop].[dbo].tbl_shop_user s "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.shopid = s.userid "
		sqlFromWhere = sqlFromWhere + " 	left join [db_storage].[dbo].tbl_cartoonbox_master cm "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.workidx = cm.idx "
		sqlFromWhere = sqlFromWhere + " where "
		sqlFromWhere = sqlFromWhere + " 	1 = 1 "

		if FRectShopid<>"" then
			sqlFromWhere = sqlFromWhere + " 	and m.shopid='" + FRectShopid + "' "
		end if

		if FRectStateCD<>"" then
			sqlFromWhere = sqlFromWhere + " 	and m.statecd='" + FRectStateCD + "' "
		end if

		if FRectReportDate<>"" then
			sqlFromWhere = sqlFromWhere + " 	and m.reportdate='" + FRectReportDate + "' "
		end if

		if FRectReportNo<>"" then
			sqlFromWhere = sqlFromWhere + " 	and m.reportno='" + FRectReportNo + "' "
		end if

		if FRectMasterIDX<>"" then
			sqlFromWhere = sqlFromWhere + " 	and m.idx='" + FRectMasterIDX + "' "
		end if

		if FRectExcNoReport = "Y" then
			sqlFromWhere = sqlFromWhere + " 	and IsNull(m.exportdeclarefilename, '') <> '' "
		end if

		if FRectDateFlag <> "" then
			if (FRectDateFlag = "regdate") then
				sqlFromWhere = sqlFromWhere + " 	and m.regdate >= '" + FRectFromDate + "' "
				sqlFromWhere = sqlFromWhere + " 	and m.regdate < '" + FRectToDate + "' "
			elseif (FRectDateFlag = "reportdate") then
				sqlFromWhere = sqlFromWhere + " 	and m.reportdate >= '" + FRectFromDate + "' "
				sqlFromWhere = sqlFromWhere + " 	and m.reportdate < '" + FRectToDate + "' "
			end if
		end if

		if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				sqlFromWhere = sqlFromWhere + " 	and m.shopid not in (select id from db_partner.dbo.tbl_partner where IsNull(tplcompanyid, '') <> '') "
			else
				sqlFromWhere = sqlFromWhere + " 	and m.shopid in (select id from db_partner.dbo.tbl_partner where IsNull(tplcompanyid, '') = '" + CStr(FtplGubun) + "') "
			end if
		end if

		'======================================================================
		sqlStr = " select count(m.idx) as cnt "

		sqlStr = sqlStr + sqlFromWhere

		rsget.Open sqlStr, dbget, 1
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		'======================================================================
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	m.idx "
		sqlStr = sqlStr + " 	, m.shopid "
		sqlStr = sqlStr + " 	, s.shopname "
		sqlStr = sqlStr + " 	, cm.jungsanidx "
		sqlStr = sqlStr + " 	, m.workidx "
		sqlStr = sqlStr + " 	, m.invoiceno "
		sqlStr = sqlStr + " 	, m.invoicedate "
		sqlStr = sqlStr + " 	, m.delivermethod "
		sqlStr = sqlStr + " 	, m.exportmethod "
		sqlStr = sqlStr + " 	, m.jungsantype "
		sqlStr = sqlStr + " 	, m.priceunit "
		sqlStr = sqlStr + " 	, m.exchangerate "
		sqlStr = sqlStr + " 	, m.totalboxno "
		sqlStr = sqlStr + " 	, m.totalboxprice "
		sqlStr = sqlStr + " 	, m.totalgoodsprice "
		sqlStr = sqlStr + " 	, m.totalprice "
		sqlStr = sqlStr + " 	, IsNull(m.exportdeclarefilename, '') as exportdeclarefilename "
		sqlStr = sqlStr + " 	, m.realfilename "
		sqlStr = sqlStr + " 	, m.exporteraddr "
		sqlStr = sqlStr + " 	, m.riskmesseraddr "
		sqlStr = sqlStr + " 	, m.notifyaddr "
		sqlStr = sqlStr + " 	, m.portname "
		sqlStr = sqlStr + " 	, m.destinationname "
		sqlStr = sqlStr + " 	, m.carriername "
		sqlStr = sqlStr + " 	, m.carrierdate "
		sqlStr = sqlStr + " 	, m.goodscomment1 "
		sqlStr = sqlStr + " 	, m.goodscomment2 "
		sqlStr = sqlStr + " 	, m.lccomment "
		sqlStr = sqlStr + " 	, m.lcbank "
		sqlStr = sqlStr + " 	, m.comment "
		sqlStr = sqlStr + " 	, m.statecd "
		sqlStr = sqlStr + " 	, m.reguserid "
		sqlStr = sqlStr + " 	, m.lastupdate "
		sqlStr = sqlStr + " 	, m.regdate "

		sqlStr = sqlStr + " 	, m.reportno "
		sqlStr = sqlStr + " 	, m.reportdate "
		sqlStr = sqlStr + " 	, m.reportpriceunit "
		sqlStr = sqlStr + " 	, IsNull(m.reportexchangerate, 0) as reportexchangerate "
		sqlStr = sqlStr + " 	, IsNull(m.reportforeigntotalprice, 0) as reportforeigntotalprice "
		sqlStr = sqlStr + " 	, IsNull(m.reporttotalprice, 0) as reporttotalprice "

		sqlStr = sqlStr + " 	, m.totalGoodsPriceWon, m.totalDeliverPriceWon, m.totalPriceWon, m.totalGoodsPriceForeign, m.totalDeliverPriceForeign, m.totalPriceForeign "
		sqlStr = sqlStr + " 	, IsNull(m.exportdeclarefilename2, '') as exportdeclarefilename2 "
		sqlStr = sqlStr + " 	, IsNull(m.exportdeclarefilename3, '') as exportdeclarefilename3 "
		sqlStr = sqlStr + " 	, m.realfilename2, m.realfilename3 "

		sqlStr = sqlStr + sqlFromWhere

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	m.idx desc "

		'======================================================================
		''response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

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
				set FItemList(i) = new COffInvoiceMasterItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fshopid			= rsget("shopid")
				FItemList(i).Fshopname			= db2html(rsget("shopname"))
				FItemList(i).Fjungsanidx		= rsget("jungsanidx")
				FItemList(i).Fworkidx			= rsget("workidx")
				FItemList(i).Finvoiceno			= db2html(rsget("invoiceno"))
				FItemList(i).Finvoicedate		= rsget("invoicedate")
				FItemList(i).Fdelivermethod		= rsget("delivermethod")
				FItemList(i).Fexportmethod		= rsget("exportmethod")
				FItemList(i).Fjungsantype		= rsget("jungsantype")
				FItemList(i).Fpriceunit			= rsget("priceunit")
				FItemList(i).Fexchangerate		= rsget("exchangerate")
				FItemList(i).Ftotalboxno		= rsget("totalboxno")
				FItemList(i).Ftotalboxprice		= rsget("totalboxprice")					'// deprecated
				FItemList(i).Ftotalgoodsprice	= rsget("totalgoodsprice")					'// deprecated
				FItemList(i).Ftotalprice		= rsget("totalprice")						'// deprecated
				FItemList(i).Fexportdeclarefilename	= rsget("exportdeclarefilename")
				FItemList(i).Fexportdeclarefilename2	= rsget("exportdeclarefilename2")
				FItemList(i).Fexportdeclarefilename3	= rsget("exportdeclarefilename3")
				FItemList(i).Frealfilename			= db2html(rsget("realfilename"))
				FItemList(i).Frealfilename2			= db2html(rsget("realfilename2"))
				FItemList(i).Frealfilename3			= db2html(rsget("realfilename3"))
				FItemList(i).Fexporteraddr			= db2html(rsget("exporteraddr"))
				FItemList(i).Friskmesseraddr		= db2html(rsget("riskmesseraddr"))
				FItemList(i).Fnotifyaddr			= db2html(rsget("notifyaddr"))
				FItemList(i).Fportname				= db2html(rsget("portname"))
				FItemList(i).Fdestinationname		= db2html(rsget("destinationname"))
				FItemList(i).Fcarriername			= db2html(rsget("carriername"))
				FItemList(i).Fcarrierdate			= rsget("carrierdate")
				FItemList(i).Fgoodscomment1			= db2html(rsget("goodscomment1"))
				FItemList(i).Fgoodscomment2			= db2html(rsget("goodscomment2"))
				FItemList(i).Flccomment				= db2html(rsget("lccomment"))
				FItemList(i).Flcbank				= db2html(rsget("lcbank"))
				FItemList(i).Fcomment				= db2html(rsget("comment"))

				FItemList(i).Fstatecd				= rsget("statecd")

				FItemList(i).Freguserid				= rsget("reguserid")
				FItemList(i).Flastupdate			= rsget("lastupdate")
				FItemList(i).Fregdate				= rsget("regdate")

				FItemList(i).Freportno					= rsget("reportno")
				FItemList(i).Freportdate				= rsget("reportdate")
				FItemList(i).Freportpriceunit			= rsget("reportpriceunit")
				FItemList(i).Freportexchangerate		= rsget("reportexchangerate")
				FItemList(i).Freportforeigntotalprice	= rsget("reportforeigntotalprice")
				FItemList(i).Freporttotalprice			= rsget("reporttotalprice")

				FItemList(i).FtotalGoodsPriceWon		= rsget("totalGoodsPriceWon")
				FItemList(i).FtotalDeliverPriceWon		= rsget("totalDeliverPriceWon")
				FItemList(i).FtotalPriceWon				= rsget("totalPriceWon")

				FItemList(i).FtotalGoodsPriceForeign	= rsget("totalGoodsPriceForeign")
				FItemList(i).FtotalDeliverPriceForeign	= rsget("totalDeliverPriceForeign")
				FItemList(i).FtotalPriceForeign			= rsget("totalPriceForeign")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

	public Sub GetMasterOne()

		dim i, sqlStr, sqlFromWhere
		dim tmpstr

		'======================================================================
		sqlFromWhere = " from "

		sqlFromWhere = sqlFromWhere + " 	[db_storage].[dbo].tbl_offline_invoice_master m "
		sqlFromWhere = sqlFromWhere + " 	left join [db_shop].[dbo].tbl_shop_user s "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.shopid = s.userid "
		sqlFromWhere = sqlFromWhere + " 	left join [db_storage].[dbo].tbl_cartoonbox_master cm "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.workidx = cm.idx "
		sqlFromWhere = sqlFromWhere + " 	left join db_shop.dbo.tbl_shop_exchangeRate ex "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.priceunit = ex.currencyunit "
		sqlFromWhere = sqlFromWhere + " where "
		sqlFromWhere = sqlFromWhere + " 	m.idx = " + CStr(FRectMasterIdx) + + " "

		if FRectShopid<>"" then
			sqlFromWhere = sqlFromWhere + " 	and m.shopid='" + FRectShopid + "' "
		end if

		if FRectStateCD<>"" then
			sqlFromWhere = sqlFromWhere + " 	and m.statecd='" + FRectStateCD + "' "
		end if

		'======================================================================
		sqlStr = " select top 1 "
		sqlStr = sqlStr + " 	m.idx "
		sqlStr = sqlStr + " 	, m.shopid "
		sqlStr = sqlStr + " 	, s.shopname "
		sqlStr = sqlStr + " 	, cm.jungsanidx "
		sqlStr = sqlStr + " 	, m.workidx "
		sqlStr = sqlStr + " 	, m.invoiceno "
		sqlStr = sqlStr + " 	, m.invoicedate "
		sqlStr = sqlStr + " 	, m.delivermethod "
		sqlStr = sqlStr + " 	, m.exportmethod "
		sqlStr = sqlStr + " 	, m.jungsantype "
		sqlStr = sqlStr + " 	, m.priceunit "
		' ==== US$, SG$, WON
		sqlStr = sqlStr & " 	, (case when IsNull(ex.currencyunit, '') = 'WON' or IsNull(ex.currencyunit, '') = 'KRW' then 'KRW' else Left(IsNull(ex.currencyunit, ''), 2) end)"
		sqlStr = sqlStr & " 		+ (case when IsNull(ex.currencyunit, '') = 'WON' or IsNull(ex.currencyunit, '') = 'KRW' then '' else IsNull(ex.currencychar, '') end) as priceunitstring"
		sqlStr = sqlStr + " 	, m.exchangerate "
		sqlStr = sqlStr + " 	, m.totalboxno "
		sqlStr = sqlStr + " 	, m.totalboxprice "			'// deprecated
		sqlStr = sqlStr + " 	, m.totalgoodsprice "		'// deprecated
		sqlStr = sqlStr + " 	, m.totalprice "			'// deprecated
		sqlStr = sqlStr + " 	, m.exportdeclarefilename "
		sqlStr = sqlStr + " 	, m.realfilename "
		sqlStr = sqlStr + " 	, m.exporteraddr "
		sqlStr = sqlStr + " 	, m.riskmesseraddr "
		sqlStr = sqlStr + " 	, m.notifyaddr "
		sqlStr = sqlStr + " 	, m.portname "
		sqlStr = sqlStr + " 	, m.destinationname "
		sqlStr = sqlStr + " 	, m.carriername "
		sqlStr = sqlStr + " 	, m.carrierdate "
		sqlStr = sqlStr + " 	, m.goodscomment1 "
		sqlStr = sqlStr + " 	, m.goodscomment2 "
		sqlStr = sqlStr + " 	, m.lccomment "
		sqlStr = sqlStr + " 	, m.lcbank "
		sqlStr = sqlStr + " 	, m.comment "
		sqlStr = sqlStr + " 	, m.statecd "
		sqlStr = sqlStr + " 	, m.reguserid "
		sqlStr = sqlStr + " 	, m.lastupdate "
		sqlStr = sqlStr + " 	, m.regdate "

		sqlStr = sqlStr + " 	, m.reportno "
		sqlStr = sqlStr + " 	, m.reportdate "
		sqlStr = sqlStr + " 	, m.reportpriceunit "
		sqlStr = sqlStr + " 	, IsNull(m.reportexchangerate, 0) as reportexchangerate "
		sqlStr = sqlStr + " 	, IsNull(m.reportforeigntotalprice, 0) as reportforeigntotalprice "
		sqlStr = sqlStr + " 	, IsNull(m.reporttotalprice, 0) as reporttotalprice "

		sqlStr = sqlStr + " 	, m.totalGoodsPriceWon, m.totalDeliverPriceWon, m.totalPriceWon, m.totalGoodsPriceForeign, m.totalDeliverPriceForeign, m.totalPriceForeign "
		sqlStr = sqlStr + "     , s.loginsite, s.currencyunit"
		sqlStr = sqlStr + "		, m.reportno2, m.reportno3 "
		sqlStr = sqlStr + "		, m.exportdeclarefilename2, m.exportdeclarefilename3 "
		sqlStr = sqlStr + " 	, m.realfilename2,  m.realfilename3"

		sqlStr = sqlStr + sqlFromWhere

		'=======================================================================
		''response.write sqlStr
		''response.end
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		set FOneItem = new COffInvoiceMasterItem
		if Not rsget.Eof then

			FOneItem.Fidx       	= rsget("idx")

			FOneItem.Fidx				= rsget("idx")
			FOneItem.Fshopid			= rsget("shopid")
			FOneItem.Fshopname			= db2html(rsget("shopname"))
			FOneItem.Fjungsanidx		= rsget("jungsanidx")
			FOneItem.Fworkidx			= rsget("workidx")
			FOneItem.Finvoiceno			= db2html(rsget("invoiceno"))
			FOneItem.Finvoicedate		= rsget("invoicedate")
			FOneItem.Fdelivermethod		= rsget("delivermethod")
			FOneItem.Fexportmethod		= rsget("exportmethod")
			FOneItem.Fjungsantype		= rsget("jungsantype")
			FOneItem.Fpriceunit			= rsget("priceunit")
			FOneItem.Fpriceunitstring	= rsget("priceunitstring")
			FOneItem.Fexchangerate		= rsget("exchangerate")
			FOneItem.Ftotalboxno		= rsget("totalboxno")
			FOneItem.Ftotalboxprice		= rsget("totalboxprice")				'// deprecated
			FOneItem.Ftotalgoodsprice	= rsget("totalgoodsprice")				'// deprecated
			FOneItem.Ftotalprice		= rsget("totalprice")					'// deprecated
			FOneItem.Fexportdeclarefilename	= rsget("exportdeclarefilename")
			FOneItem.Fexportdeclarefilename2	= rsget("exportdeclarefilename2")
			FOneItem.Fexportdeclarefilename3	= rsget("exportdeclarefilename3")

			FOneItem.Frealfilename			= db2html(rsget("realfilename"))
			FOneItem.Frealfilename2			= db2html(rsget("realfilename2"))
			FOneItem.Frealfilename3			= db2html(rsget("realfilename3"))

			FOneItem.Fexporteraddr			= db2html(rsget("exporteraddr"))
			FOneItem.Friskmesseraddr		= db2html(rsget("riskmesseraddr"))
			FOneItem.Fnotifyaddr			= db2html(rsget("notifyaddr"))
			FOneItem.Fportname				= db2html(rsget("portname"))
			FOneItem.Fdestinationname		= db2html(rsget("destinationname"))
			FOneItem.Fcarriername			= db2html(rsget("carriername"))
			FOneItem.Fcarrierdate			= rsget("carrierdate")
			FOneItem.Fgoodscomment1			= db2html(rsget("goodscomment1"))
			FOneItem.Fgoodscomment2			= db2html(rsget("goodscomment2"))
			FOneItem.Flccomment				= db2html(rsget("lccomment"))
			FOneItem.Flcbank				= db2html(rsget("lcbank"))
			FOneItem.Fcomment				= db2html(rsget("comment"))

			FOneItem.Fstatecd				= rsget("statecd")

			FOneItem.Freguserid				= rsget("reguserid")
			FOneItem.Flastupdate			= rsget("lastupdate")
			FOneItem.Fregdate				= rsget("regdate")

			FOneItem.Freportno					= rsget("reportno")
			FOneItem.Freportno2					= rsget("reportno2")
			FOneItem.Freportno3					= rsget("reportno3")
			FOneItem.Freportdate				= rsget("reportdate")
			FOneItem.Freportpriceunit			= rsget("reportpriceunit")
			FOneItem.Freportexchangerate		= rsget("reportexchangerate")
			FOneItem.Freportforeigntotalprice	= rsget("reportforeigntotalprice")
			FOneItem.Freporttotalprice			= rsget("reporttotalprice")

			FOneItem.FtotalGoodsPriceWon		= rsget("totalGoodsPriceWon")
			FOneItem.FtotalDeliverPriceWon		= rsget("totalDeliverPriceWon")
			FOneItem.FtotalPriceWon				= rsget("totalPriceWon")

			FOneItem.FtotalGoodsPriceForeign	= rsget("totalGoodsPriceForeign")
			FOneItem.FtotalDeliverPriceForeign	= rsget("totalDeliverPriceForeign")
			FOneItem.FtotalPriceForeign			= rsget("totalPriceForeign")
			FOneItem.Floginsite					= rsget("loginsite")
			FOneItem.Fcurrencyunit				= rsget("currencyunit")

		end if
		rsget.close

	end sub

	public Sub GetDetailList()

		dim i, sqlStr, sqlFromWhere
		dim tmpstr

		'======================================================================
		sqlFromWhere = " from "

		sqlFromWhere = sqlFromWhere + " 	[db_storage].[dbo].tbl_offline_invoice_detail d "
		sqlFromWhere = sqlFromWhere + " 	left join [db_storage].[dbo].tbl_offline_invoice_master m "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.idx = d.masteridx "
		sqlFromWhere = sqlFromWhere + " where "
		sqlFromWhere = sqlFromWhere + " 	1 = 1 "

		if (FRectMasterIdx <> "") then
			sqlFromWhere = sqlFromWhere + " 	and m.idx = " + CStr(FRectMasterIdx) + + " "
		end if

		if FRectShopid<>"" then
			if (FRectShopid <> "ALL") then
				sqlFromWhere = sqlFromWhere + " 	and m.shopid='" + FRectShopid + "' "
			end if
		end if

		'======================================================================
		sqlStr = " select count(d.idx) as cnt "

		sqlStr = sqlStr + sqlFromWhere

		rsget.Open sqlStr, dbget, 1
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		'======================================================================
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	d.idx "
		sqlStr = sqlStr + " 	, d.masteridx "
		sqlStr = sqlStr + " 	, d.cartonboxno "
		sqlStr = sqlStr + " 	, d.goodscomment "
		sqlStr = sqlStr + " 	, d.nweight "
		sqlStr = sqlStr + " 	, d.gweight "
		sqlStr = sqlStr + " 	, IsNull(d.emsPrice, 0) as emsPrice "
		sqlStr = sqlStr + " 	, d.lastupdate "
		sqlStr = sqlStr + " 	, d.regdate "

		sqlStr = sqlStr + sqlFromWhere

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	d.cartonboxno  "

		'======================================================================
		'response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

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
				set FItemList(i) = new COffInvoiceDetailItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fmasteridx			= rsget("masteridx")
				FItemList(i).Fcartonboxno		= rsget("cartonboxno")
				FItemList(i).Fgoodscomment		= rsget("goodscomment")
				FItemList(i).Fnweight			= rsget("nweight")
				FItemList(i).Fgweight			= rsget("gweight")
				FItemList(i).FemsPrice			= rsget("emsPrice")
				FItemList(i).Flastupdate		= rsget("lastupdate")
				FItemList(i).Fregdate			= rsget("regdate")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

	public Sub GetDetailOne()

		dim i,sqlStr, sqlFromWhere
		dim tmpstr

	end sub

	public Sub GetProductDetailList()

		dim i, sqlStr, sqlFromWhere
		dim tmpstr

		'======================================================================
		sqlFromWhere = " from "

		sqlFromWhere = sqlFromWhere + " 	[db_storage].[dbo].tbl_offline_invoice_product_detail d "
		sqlFromWhere = sqlFromWhere + " 	left join [db_storage].[dbo].tbl_offline_invoice_master m "
		sqlFromWhere = sqlFromWhere + " 	on "
		sqlFromWhere = sqlFromWhere + " 		m.idx = d.masteridx "
		sqlFromWhere = sqlFromWhere + " where "
		sqlFromWhere = sqlFromWhere + " 	1 = 1 "

		if (FRectMasterIdx <> "") then
			sqlFromWhere = sqlFromWhere + " 	and m.idx = " + CStr(FRectMasterIdx) + + " "
		end if

		if FRectShopid<>"" then
			if (FRectShopid <> "ALL") then
				sqlFromWhere = sqlFromWhere + " 	and m.shopid='" + FRectShopid + "' "
			end if
		end if

		'======================================================================
		sqlStr = " select count(d.idx) as cnt "

		sqlStr = sqlStr + sqlFromWhere

		rsget.Open sqlStr, dbget, 1
		if  not rsget.EOF  then
			FTotalCount = rsget("cnt")
		end if
		rsget.close

		'======================================================================
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	d.idx "
		sqlStr = sqlStr + " 	, d.masteridx "
		sqlStr = sqlStr + " 	, d.orderno "
		sqlStr = sqlStr + " 	, d.goodscomment "
		sqlStr = sqlStr + " 	, d.totalboxno "
		sqlStr = sqlStr + " 	, d.priceperbox "
		sqlStr = sqlStr + " 	, d.totalprice "
		sqlStr = sqlStr + " 	, d.lastupdate "
		sqlStr = sqlStr + " 	, d.regdate "

		sqlStr = sqlStr + sqlFromWhere

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	d.orderno, d.idx  "

		'======================================================================
		'response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

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
				set FItemList(i) = new COffInvoiceProductDetailItem

				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fmasteridx			= rsget("masteridx")
				FItemList(i).Forderno			= rsget("orderno")
				FItemList(i).Fgoodscomment		= db2html(rsget("goodscomment"))
				FItemList(i).Ftotalboxno		= rsget("totalboxno")
				FItemList(i).Fpriceperbox		= rsget("priceperbox")
				FItemList(i).Ftotalprice		= rsget("totalprice")
				FItemList(i).Flastupdate		= rsget("lastupdate")
				FItemList(i).Fregdate			= rsget("regdate")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

	end sub

	public Sub GetProductDetailOne()

		dim i,sqlStr, sqlFromWhere
		dim tmpstr

	end sub

	'--해외출고 견적서
	public Function fnGetQuotationSheet
		Dim strSql

		If (FRectJungsanidx = "") And (FRectWorkidx <> "") then
			strSql ="[db_storage].[dbo].sp_Ten_fran_QuotationSheetByWork('"&FRectMasterIdx&"','"&FRectWorkidx&"','"&FRectLoginsite&"')"
		Else
			strSql ="[db_storage].[dbo].sp_Ten_fran_QuotationSheet('"&FRectMasterIdx&"','"&FRectLoginsite&"')"
		End If
		''response.Write strSql
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			  fnGetQuotationSheet = rsget.getRows()
		END IF
		rsget.close
	End Function

	'견적서 상품리스트
	public Function fnGetQuotationSheetItemList
		Dim strSql
		If (FRectJungsanidx = "") And (FRectWorkidx <> "") then
			strSql ="[db_storage].[dbo].sp_Ten_fran_QuotationitemListByWork('"&FRectMasterIdx&"','"&FRectWorkidx&"','"&FRectLoginsite&"')"
		Else
			strSql ="[db_storage].[dbo].sp_Ten_fran_QuotationitemList('"&FRectMasterIdx&"','"&FRectLoginsite&"')"
		End If
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetQuotationSheetItemList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'견적서 상품리스트
	public Function fnGetFranItemList
		Dim strSql
		If (FRectJungsanidx = "") And (FRectWorkidx <> "") then
			strSql ="[db_storage].[dbo].sp_Ten_fran_itemListByWork('"&FRectMasterIdx&"','"&FRectWorkidx&"','"&FRectLoginsite&"')"
		Else
			strSql ="[db_storage].[dbo].sp_Ten_fran_itemList('"&FRectMasterIdx&"','"&FRectLoginsite&"')"
		End If
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetFranItemList = rsget.getRows()
		END IF
		rsget.close
	End Function


  	'
  	public Function fnGetFranInvoice
		Dim strSql
		If (FRectJungsanidx = "") And (FRectWorkidx <> "") then
			strSql ="[db_storage].[dbo].sp_Ten_fran_InvoiceByWork('"&FRectMasterIdx&"','"&FRectWorkidx&"','"&FRectLoginsite&"')"
		Else
			strSql ="[db_storage].[dbo].sp_Ten_fran_Invoice('"&FRectMasterIdx&"','"&FRectLoginsite&"')"
		End If
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			  FinvoiceNo				= rsget("invoiceno")
			  Finvoicedate				= rsget("invoicedate")
			  Fbaljucode               	= rsget("baljucode")
			  Fbaljuid                	= rsget("baljuid")
			  Fbeasongdate             	= rsget("beasongdate")
			  Fregdate					= rsget("regdate")
			  Fpriceunit				= rsget("priceunit")
			  FtotalGoodsPriceWon		= rsget("totalGoodsPriceWon")
			  FtotalDeliverPriceWon 	= rsget("totalDeliverPriceWon")
			  FtotalPriceWon  			= rsget("totalPriceWon")
			  FtotalGoodsPriceForeign	= rsget("totalGoodsPriceForeign")
			  FtotalDeliverPriceForeign	= rsget("totalDeliverPriceForeign")
			  FtotalPriceForeign		= rsget("totalPriceForeign")
			  FfreightTerm             	= rsget("freightTerm")
			  FopenState               	= rsget("openState")
			  FshippingAddress         	= rsget("shippingAddress")
			  FinvoiceAddress          	= rsget("invoiceAddress")
			  FcurrencyChar				= rsget("currencyChar")
			  Fcurrencyunit				= rsget("currencyunit")
			  Fcomment					= rsget("comment")
		END IF
		rsget.close
	End Function

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 200
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
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
