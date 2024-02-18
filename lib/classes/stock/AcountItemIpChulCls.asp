<%
Class CAcountItemIpChulItem
    public FIpChulidx
	public FIpChulCode
	public Fdivcode
	public FItemId
	public FItemOption
	public FItemName
	public FItemOptionName
	public FSellCash
	public FSuplycash
	public FBuycash
	public FItemNo
	public FSocID
	public FExecutedt
	public fpurchasetypename
	public FItemgubun
	public Fipchulflag
	public Fimakerid

    public FmasterDeldt
    public FdetailDeldt
    public FIpChulMwgubun
    public Fdetailid
    public FonlineMwDiv
    public Fcentermwdiv

    public FStockMwDiv
    public FStockBuyPrice
    public FStockShopComm_cd
    public FStockShopBuyCash
	public Fpurchasetype

    public FavgipgoPrice
    public FavgShopIpgoPrice

    public Fcatename

    public function isDeleted
        isDeleted = (Not isNULL(FmasterDeldt)) or (Not isNULL(FdetailDeldt))
    end function

	public function GetIpchulColor()
		if Fipchulflag="I" then
			GetIpchulColor = "#3333EE"
		elseif Fipchulflag="S" then
			GetIpchulColor = "#EE3333"
		elseif Fipchulflag="E" then
			GetIpchulColor = "#EE33EE"
		end if

	end function

    public function GetDivCodeColor()
		if Fdivcode="002" then
			GetDivCodeColor = "#000000"
		elseif Fdivcode="001" then
			GetDivCodeColor = "#DD5555"
		elseif Fdivcode="801" then
			GetDivCodeColor = "#DD5555"
		elseif Fdivcode="802" then
			GetDivCodeColor = "#5555DD"
		end if
	end function

	public function GetDivCodeName()
		if Fdivcode="002" then
			GetDivCodeName = "위탁"
		elseif Fdivcode="001" then
			GetDivCodeName = "매입"
		elseif Fdivcode="003" then
			GetDivCodeName = "판촉출고"
		elseif Fdivcode="004" then
			GetDivCodeName = "외부출고"
		elseif Fdivcode="005" then
			GetDivCodeName = "협찬출고"
		elseif Fdivcode="006" then
			GetDivCodeName = "B2B출고"
		elseif Fdivcode="007" then
			GetDivCodeName = "기타출고"
		elseif Fdivcode="101" then
			GetDivCodeName = "위탁출고"
		elseif Fdivcode="999" then
			GetDivCodeName = "기타(정산않함)"
		elseif Fdivcode="801" then
			GetDivCodeName = "Off매입"
		elseif Fdivcode="802" then
			GetDivCodeName = "Off위탁"

		end if
	end function

	public function GetBarCode()
		GetBarCode = Fitemgubun & Format00(6,Fitemid) & Fitemoption
		if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CAcountItemIpChul
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public fArrLIst

	public FRectStartDay
	public FRectEndDay
	public FRectGubun
	public FRectDesigner
	public FRectItemId
	public FRectItemOption
	public FRectShopid

	public FRectItemGubun
    public FRectDeletInclude
    public FRectIpChulCode
	public FtplGubun
	public FRectIpChulMwgubun
	public FRectOnlineMwDiv
	public FRectCentermwdiv
	public FRectStockMwDiv
	public FRectBrandPurchaseType

	public Sub getIpChulListByItemByShop()
		dim sqlStr,i

		sqlStr = " select top 1000 m.code, m.executedt, d.iitemgubun, d.itemid, d.itemoption, "
		sqlStr = sqlStr + " d.sellcash, d.suplycash, d.buycash, d.itemno,"
		sqlStr = sqlStr + " m.socid, d.iitemname, d.iitemoptionname,m.ipchulflag,m.divcode"
		sqlStr = sqlStr + " ,d.imakerid"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " where m.code=d.mastercode"

        if FRectStartDay<>"" then
		 	sqlStr = sqlStr + " and m.executedt>='" + FRectStartDay + "'"
		end if

        if FRectEndDay<>"" then
		   	sqlStr = sqlStr + " and m.executedt<'" + FRectEndDay + "'"
		end if

		if FRectGubun<>"" then
			sqlStr = sqlStr + " and m.ipchulflag = '" + FRectGubun + "'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.imakerid='" + FRectDesigner + "'"
		end if

		if FRectItemGubun<>"" then
			sqlStr = sqlStr + " and d.iitemgubun='" + FRectItemGubun + "'"
		end if

		if FRectItemId<>"" then
			sqlStr = sqlStr + " and d.itemid=" + FRectItemId + ""
		end if

		if FRectItemOption<>"" then
			sqlStr = sqlStr + " and d.itemoption='" + FRectItemOption + "' "
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.socid='" + CStr(FRectShopid) + "'"
		end if

		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"

		sqlStr = sqlStr + " order by m.code , m.socid, d.itemid"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CAcountItemIpChulItem

				FItemList(i).FIpChulCode	 = rsget("code")
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemOption     = rsget("itemoption")
				FItemList(i).FItemName       = db2html(rsget("iitemname"))
				FItemList(i).FItemOptionName = db2html(rsget("iitemoptionname"))
				FItemList(i).FSellCash       = rsget("sellcash")
				FItemList(i).FSuplycash      = rsget("suplycash")
				FItemList(i).FBuycash		 = rsget("buycash")
				FItemList(i).FItemNo         = rsget("itemno")
				FItemList(i).FSocID         = rsget("socid")
				FItemList(i).Fexecutedt		= rsget("executedt")
				FItemList(i).FItemgubun		= rsget("iitemgubun")
				FItemList(i).Fipchulflag	= rsget("ipchulflag")
				FItemList(i).Fdivcode       = rsget("divcode")

				FItemList(i).Fimakerid      = rsget("imakerid")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	' 밑에 함수를 수정할경우 getIpChulListByItemNotPaging 함수도 똑같이 수정해야 한다.
	public Sub getIpChulListByItem()
		dim sqlStr,i, AddSql

		AddSql=""
		if FRectStartDay<>"" then
		    AddSql = AddSql + " and m.executedt>='" + FRectStartDay + "'"
		end if

        if FRectEndDay<>"" then
		    AddSql = AddSql + " and m.executedt<'" + FRectEndDay + "'"
		end if

		if FRectGubun<>"" Then
			Select Case FRectGubun
				Case "SM"
					AddSql = AddSql + " and m.ipchulflag = 'S' and p.userdiv = 501 and d.mwgubun <> 'C' "
				Case "SW"
					AddSql = AddSql + " and m.ipchulflag = 'S' and p.userdiv = 501 and d.mwgubun = 'C' "
				Case "S"
					AddSql = AddSql + " and m.ipchulflag = 'S' "
				Case "SE"
					AddSql = AddSql + " and m.ipchulflag = 'S' and p.userdiv <> 501 "
				Case Else
					AddSql = AddSql + " and m.ipchulflag = '" + FRectGubun + "'"
			End Select
		end if

		if FRectDesigner<>"" then
			AddSql = AddSql + " and d.imakerid='" + FRectDesigner + "'"
		end if

		if FRectItemGubun<>"" then
			AddSql = AddSql + " and d.iitemgubun='" + FRectItemGubun + "'"
		end if

		if FRectItemId<>"" then
			AddSql = AddSql + " and d.itemid=" + FRectItemId + ""
		end if

		if FRectItemOption<>"" then
			AddSql = AddSql + " and d.itemoption='" + FRectItemOption + "' "
		end if

		if FRectShopid<>"" then
			AddSql = AddSql + " and m.socid='" + CStr(FRectShopid) + "'"
		end if

        if (FRectDeletInclude<>"on") then
    		AddSql = AddSql + " and m.deldt is NULL"
    		AddSql = AddSql + " and d.deldt is NULL"
        end if

        if (FRectIpChulCode<>"") then
            AddSql = AddSql + " and m.code='"&FRectIpChulCode&"'"
        end if

		if (FRectIpChulMwgubun <> "") then
			if (FRectIpChulMwgubun = "X") then
				AddSql = AddSql + " and isnull(d.mwgubun,'') not in ('M', 'F', 'C', 'W') "
			else
				AddSql = AddSql + " and isnull(d.mwgubun,'') = '" & CStr(FRectIpChulMwgubun) & "' "
			end if
		end if

		if (FRectOnlineMwDiv <> "") then
			if (FRectOnlineMwDiv = "X") then
				AddSql = AddSql + " and isnull(i.mwdiv,'') not in ('M', 'W', 'U') "
			else
				AddSql = AddSql + " and isnull(i.mwdiv,'') = '" & CStr(FRectOnlineMwDiv) & "' "
			end if
		end if

		if (FRectCentermwdiv <> "") then
			if (FRectCentermwdiv = "X") then
				AddSql = AddSql + " and isnull(s.centermwdiv,'') not in ('M', 'W') "
			else
				AddSql = AddSql + " and isnull(s.centermwdiv,'') = '" & CStr(FRectCentermwdiv) & "' "
			end if
		end if

		if (FRectStockMwDiv <> "") then
			if (FRectStockMwDiv = "X") then
				AddSql = AddSql + " and IsNull(L.LastMwdiv, '') not in ('M', 'W') "
			else
				AddSql = AddSql + " and isnull(L.LastMwdiv,'') = '" & CStr(FRectStockMwDiv) & "' "
			end if
		end if
		if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				AddSql = AddSql + " 	and IsNull(p.tplcompanyid, '') = '' "
			else
				AddSql = AddSql + " 	and IsNull(p.tplcompanyid, '') = '" + CStr(FtplGubun) + "' "
			end if
		end if

		if (FRectBrandPurchaseType <> "") then
			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				AddSql = AddSql + " 	and pp.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				AddSql = AddSql & " 	and pp.purchasetype in ('3','5','6')"
			else
				AddSql = AddSql + " 	and pp.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		sqlStr = " select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m with (nolock)"
		sqlStr = sqlStr + " Join [db_storage].[dbo].tbl_acount_storage_detail d with (nolock)"
		sqlStr = sqlStr + "     on m.code=d.mastercode"
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + "     on d.iitemgubun='10'"
		sqlStr = sqlStr + "     and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + "     on d.iitemgubun=s.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=s.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " left Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary L with (nolock)"
		sqlStr = sqlStr + "     on L.yyyymm=convert(Varchar(7),m.executeDt,21)"
		sqlStr = sqlStr + "     and  L.itemgubun=D.iitemgubun"
		sqlStr = sqlStr + "     and L.itemid=D.itemid"
		sqlStr = sqlStr + "     and L.itemoption=D.itemoption"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner p with (nolock)"
	    sqlStr = sqlStr + " 	on m.socid=p.id"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr + " 	on d.imakerid=pp.id"
		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].tbl_partner_comm_code as pc with (nolock)"
		sqlStr = sqlStr & " 	on pc.pcomm_group='purchasetype' and pc.pcomm_isusing='Y' and pp.purchasetype=pc.pcomm_cd"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + AddSql

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select top "&FPageSize*FCurrPage&" m.id, m.code, m.divcode, m.executedt, d.iitemgubun, d.itemid, d.itemoption, "
		sqlStr = sqlStr + " d.sellcash, d.suplycash, d.buycash, d.itemno,"
		sqlStr = sqlStr + " m.socid, d.iitemname, d.iitemoptionname,m.ipchulflag, d.imakerid"
		sqlStr = sqlStr + " ,m.deldt as masterDeldt, d.deldt as detailDeldt"
		sqlStr = sqlStr + " ,d.mwgubun as ipchulMwGubun, d.id as detailid"
		sqlStr = sqlStr + " ,i.mwdiv as onlineMwDiv"
		sqlStr = sqlStr + " ,s.centermwdiv, pc.pcomm_name as purchasetypename"
		sqlStr = sqlStr + " ,L.LastMwdiv, L.lastBuyPrice, SL.LstComm_cd, SL.LstBuyCash, pp.purchasetype, L.avgipgoPrice, SL.avgShopIpgoPrice "
        sqlStr = sqlStr + " , (case when dc.catename is not NULL then dc.catename else '미지정' end) as catename "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m with (nolock)"
		sqlStr = sqlStr + "     Join [db_storage].[dbo].tbl_acount_storage_detail d with (nolock)"
		sqlStr = sqlStr + "     on m.code=d.mastercode"
		sqlStr = sqlStr + " left Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary L with (nolock)"
		sqlStr = sqlStr + "     on L.yyyymm=convert(Varchar(7),m.executeDt,21)"
		sqlStr = sqlStr + "     and  L.itemgubun=D.iitemgubun"
		sqlStr = sqlStr + "     and L.itemid=D.itemid"
		sqlStr = sqlStr + "     and L.itemoption=D.itemoption"
		sqlStr = sqlStr + " left Join db_summary.dbo.tbl_monthly_accumulated_Shopstock_summary SL with (nolock)"
		sqlStr = sqlStr + "     on SL.yyyymm=convert(Varchar(7),m.executeDt,21)"
		sqlStr = sqlStr + "     and SL.shopid=m.socid"
		sqlStr = sqlStr + "     and SL.itemgubun=D.iitemgubun"
		sqlStr = sqlStr + "     and SL.itemid=D.itemid"
		sqlStr = sqlStr + "     and SL.itemoption=D.itemoption"
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + "     on d.iitemgubun='10'"
		sqlStr = sqlStr + "     and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + "     on d.iitemgubun=s.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=s.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=s.itemoption"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner p with (nolock)"
	    sqlStr = sqlStr + " 	on m.socid=p.id"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr + " 	on d.imakerid=pp.id"
		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].tbl_partner_comm_code as pc with (nolock)"
		sqlStr = sqlStr & " 	on pc.pcomm_group='purchasetype' and pc.pcomm_isusing='Y' and pp.purchasetype=pc.pcomm_cd"
		sqlStr = sqlStr & " left join db_item.[dbo].[tbl_display_cate] dc with (nolock) on i.dispcate1 = dc.catecode "
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + AddSql

		sqlStr = sqlStr + " order by m.id desc"

		''response.write sqlStr
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/06

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CAcountItemIpChulItem
				FItemList(i).fpurchasetypename      = rsget("purchasetypename")
                FItemList(i).FIpChulidx      = rsget("id")
				FItemList(i).FIpChulCode	 = rsget("code")
				FItemList(i).Fdivcode        = rsget("divcode")
				FItemList(i).FItemId         = rsget("itemid")
				FItemList(i).FItemOption     = rsget("itemoption")
				FItemList(i).FItemName       = db2html(rsget("iitemname"))
				FItemList(i).FItemOptionName = db2html(rsget("iitemoptionname"))
				FItemList(i).FSellCash       = rsget("sellcash")
				FItemList(i).FSuplycash      = rsget("suplycash")
				FItemList(i).FBuycash		 = rsget("buycash")
				FItemList(i).FItemNo         = rsget("itemno")
				FItemList(i).FSocID         = rsget("socid")
				FItemList(i).Fexecutedt		= rsget("executedt")
				FItemList(i).FItemgubun		= rsget("iitemgubun")
				FItemList(i).Fipchulflag	= rsget("ipchulflag")
				FItemList(i).Fimakerid		= rsget("imakerid")

                FItemList(i).FmasterDeldt 	= rsget("masterDeldt")
                FItemList(i).FdetailDeldt 	= rsget("detailDeldt")

                FItemList(i).FIpChulMwgubun = rsget("ipchulMwGubun")
                FItemList(i).Fdetailid      = rsget("detailid")
                FItemList(i).FonlineMwDiv   = rsget("onlineMwDiv")
                FItemList(i).Fcentermwdiv   = rsget("centermwdiv")

                FItemList(i).FStockMwDiv    = rsget("LastMwdiv")
                FItemList(i).FStockBuyPrice = rsget("lastBuyPrice")

                FItemList(i).FStockShopComm_cd  = rsget("LstComm_cd")
                FItemList(i).FStockShopBuyCash  = rsget("LstBuyCash")
				FItemList(i).Fpurchasetype    	= rsget("purchasetype")

                FItemList(i).FavgipgoPrice    	= rsget("avgipgoPrice")
                FItemList(i).FavgShopIpgoPrice  = rsget("avgShopIpgoPrice")

                FItemList(i).Fcatename  	= db2html(rsget("catename"))

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	' 밑에 함수를 수정할경우 getIpChulListByItem 함수도 똑같이 수정해야 한다.
	public Sub getIpChulListByItemNotPaging()
		dim sqlStr,i, AddSql

		AddSql=""
		if FRectStartDay<>"" then
		    AddSql = AddSql + " and m.executedt>='" + FRectStartDay + "'"
		end if

        if FRectEndDay<>"" then
		    AddSql = AddSql + " and m.executedt<'" + FRectEndDay + "'"
		end if

		if FRectGubun<>"" Then
			Select Case FRectGubun
				Case "SM"
					AddSql = AddSql + " and m.ipchulflag = 'S' and p.userdiv = 501 and d.mwgubun <> 'C' "
				Case "SW"
					AddSql = AddSql + " and m.ipchulflag = 'S' and p.userdiv = 501 and d.mwgubun = 'C' "
				Case "S"
					AddSql = AddSql + " and m.ipchulflag = 'S' "
				Case "SE"
					AddSql = AddSql + " and m.ipchulflag = 'S' and p.userdiv <> 501 "
				Case Else
					AddSql = AddSql + " and m.ipchulflag = '" + FRectGubun + "'"
			End Select
		end if

		if FRectDesigner<>"" then
			AddSql = AddSql + " and d.imakerid='" + FRectDesigner + "'"
		end if

		if FRectItemGubun<>"" then
			AddSql = AddSql + " and d.iitemgubun='" + FRectItemGubun + "'"
		end if

		if FRectItemId<>"" then
			AddSql = AddSql + " and d.itemid=" + FRectItemId + ""
		end if

		if FRectItemOption<>"" then
			AddSql = AddSql + " and d.itemoption='" + FRectItemOption + "' "
		end if

		if FRectShopid<>"" then
			AddSql = AddSql + " and m.socid='" + CStr(FRectShopid) + "'"
		end if

        if (FRectDeletInclude<>"on") then
    		AddSql = AddSql + " and m.deldt is NULL"
    		AddSql = AddSql + " and d.deldt is NULL"
        end if

        if (FRectIpChulCode<>"") then
            AddSql = AddSql + " and m.code='"&FRectIpChulCode&"'"
        end if

		if (FRectIpChulMwgubun <> "") then
			if (FRectIpChulMwgubun = "X") then
				AddSql = AddSql + " and isnull(d.mwgubun,'') not in ('M', 'F', 'C', 'W') "
			else
				AddSql = AddSql + " and isnull(d.mwgubun,'') = '" & CStr(FRectIpChulMwgubun) & "' "
			end if
		end if

		if (FRectOnlineMwDiv <> "") then
			if (FRectOnlineMwDiv = "X") then
				AddSql = AddSql + " and isnull(i.mwdiv,'') not in ('M', 'W', 'U') "
			else
				AddSql = AddSql + " and isnull(i.mwdiv,'') = '" & CStr(FRectOnlineMwDiv) & "' "
			end if
		end if

		if (FRectCentermwdiv <> "") then
			if (FRectCentermwdiv = "X") then
				AddSql = AddSql + " and isnull(s.centermwdiv,'') not in ('M', 'W') "
			else
				AddSql = AddSql + " and isnull(s.centermwdiv,'') = '" & CStr(FRectCentermwdiv) & "' "
			end if
		end if

		if (FRectStockMwDiv <> "") then
			if (FRectStockMwDiv = "X") then
				AddSql = AddSql + " and IsNull(L.LastMwdiv, '') not in ('M', 'W') "
			else
				AddSql = AddSql + " and isnull(L.LastMwdiv,'') = '" & CStr(FRectStockMwDiv) & "' "
			end if
		end if
		if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				AddSql = AddSql + " 	and IsNull(p.tplcompanyid, '') = '' "
			else
				AddSql = AddSql + " 	and IsNull(p.tplcompanyid, '') = '" + CStr(FtplGubun) + "' "
			end if
		end if

		if (FRectBrandPurchaseType <> "") then
			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				AddSql = AddSql + " 	and pp.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				AddSql = AddSql & " 	and pp.purchasetype in ('3','5','6')"
			else
				AddSql = AddSql + " 	and pp.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		sqlStr = " select top "&FPageSize*FCurrPage&" m.id, m.code, m.divcode, m.executedt, d.iitemgubun, d.itemid, d.itemoption, "
		sqlStr = sqlStr + " d.sellcash, d.suplycash, d.buycash, d.itemno,"
		sqlStr = sqlStr + " m.socid"
		sqlStr = sqlStr & " ,replace(replace(replace(replace(replace(d.iitemname,char(9),''),char(10),''),char(13),''),'""',''),'''','') as iitemname"
		sqlStr = sqlStr & " ,replace(replace(replace(replace(replace(d.iitemoptionname,char(9),''),char(10),''),char(13),''),'""',''),'''','') as iitemoptionname"
		sqlStr = sqlStr + " ,m.ipchulflag, d.imakerid"
		sqlStr = sqlStr + " ,m.deldt as masterDeldt, d.deldt as detailDeldt"
		sqlStr = sqlStr + " ,d.mwgubun as ipchulMwGubun, d.id as detailid"
		sqlStr = sqlStr + " ,i.mwdiv as onlineMwDiv"
		sqlStr = sqlStr + " ,s.centermwdiv, pc.pcomm_name as purchasetypename"
		sqlStr = sqlStr + " ,L.LastMwdiv, L.lastBuyPrice, SL.LstComm_cd, SL.LstBuyCash, pp.purchasetype, L.avgipgoPrice, SL.avgShopIpgoPrice "
        sqlStr = sqlStr + " , (case when dc.catename is not NULL then dc.catename else '미지정' end) as catename "
		sqlStr = sqlStr & " , (db_item.[dbo].[uf_getTenBarCodeType](d.iitemgubun,d.itemid,d.itemoption)) as tenbarcode"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m with (nolock)"
		sqlStr = sqlStr + "     Join [db_storage].[dbo].tbl_acount_storage_detail d with (nolock)"
		sqlStr = sqlStr + "     on m.code=d.mastercode"
		sqlStr = sqlStr + " left Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary L with (nolock)"
		sqlStr = sqlStr + "     on L.yyyymm=convert(Varchar(7),m.executeDt,21)"
		sqlStr = sqlStr + "     and  L.itemgubun=D.iitemgubun"
		sqlStr = sqlStr + "     and L.itemid=D.itemid"
		sqlStr = sqlStr + "     and L.itemoption=D.itemoption"
		sqlStr = sqlStr + " left Join db_summary.dbo.tbl_monthly_accumulated_Shopstock_summary SL with (nolock)"
		sqlStr = sqlStr + "     on SL.yyyymm=convert(Varchar(7),m.executeDt,21)"
		sqlStr = sqlStr + "     and SL.shopid=m.socid"
		sqlStr = sqlStr + "     and SL.itemgubun=D.iitemgubun"
		sqlStr = sqlStr + "     and SL.itemid=D.itemid"
		sqlStr = sqlStr + "     and SL.itemoption=D.itemoption"
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + "     on d.iitemgubun='10'"
		sqlStr = sqlStr + "     and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + "     on d.iitemgubun=s.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=s.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=s.itemoption"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner p with (nolock)"
	    sqlStr = sqlStr + " 	on m.socid=p.id"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr + " 	on d.imakerid=pp.id"
		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].tbl_partner_comm_code as pc with (nolock)"
		sqlStr = sqlStr & " 	on pc.pcomm_group='purchasetype' and pc.pcomm_isusing='Y' and pp.purchasetype=pc.pcomm_cd"
		sqlStr = sqlStr & " left join db_item.[dbo].[tbl_display_cate] dc with (nolock) on i.dispcate1 = dc.catecode "
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + AddSql

		sqlStr = sqlStr + " order by m.id desc"

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		dbget.CommandTimeout = 120  ''2016/01/06 (기본 30초)
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/06

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		i=0
		if  not rsget.EOF  then
			fArrLIst = rsget.getrows()
		end if

		rsget.Close
	end sub

	public Sub getIpChulListByBrand()
		dim sqlStr,i, AddSql
		dim tmpSql

		AddSql=""
		if FRectStartDay<>"" then
		    AddSql = AddSql + " and m.executedt>='" + FRectStartDay + "'"
		end if

        if FRectEndDay<>"" then
		    AddSql = AddSql + " and m.executedt<'" + FRectEndDay + "'"
		end if

		if FRectGubun<>"" Then
			Select Case FRectGubun
				Case "SM"
					AddSql = AddSql + " and m.ipchulflag = 'S' and p.userdiv = 501 and d.mwgubun <> 'C' "
				Case "SW"
					AddSql = AddSql + " and m.ipchulflag = 'S' and p.userdiv = 501 and d.mwgubun = 'C' "
				Case "S"
					AddSql = AddSql + " and m.ipchulflag = 'S' and p.userdiv <> 501 "
				Case Else
					AddSql = AddSql + " and m.ipchulflag = '" + FRectGubun + "'"
			End Select
		end if

		if FRectDesigner<>"" then
			AddSql = AddSql + " and d.imakerid='" + FRectDesigner + "'"
		end if

		if FRectItemGubun<>"" then
			AddSql = AddSql + " and d.iitemgubun='" + FRectItemGubun + "'"
		end if

		if FRectItemId<>"" then
			AddSql = AddSql + " and d.itemid=" + FRectItemId + ""
		end if

		if FRectItemOption<>"" then
			AddSql = AddSql + " and d.itemoption='" + FRectItemOption + "' "
		end if

		if FRectShopid<>"" then
			AddSql = AddSql + " and m.socid='" + CStr(FRectShopid) + "'"
		end if

        if (FRectDeletInclude<>"on") then
    		AddSql = AddSql + " and m.deldt is NULL"
    		AddSql = AddSql + " and d.deldt is NULL"
        end if

        if (FRectIpChulCode<>"") then
            AddSql = AddSql + " and m.code='"&FRectIpChulCode&"'"
        end if

		if (FRectIpChulMwgubun <> "") then
			if (FRectIpChulMwgubun = "X") then
				AddSql = AddSql + " and isnull(d.mwgubun,'') not in ('M', 'F', 'C', 'W') "
			else
				AddSql = AddSql + " and isnull(d.mwgubun,'') = '" & CStr(FRectIpChulMwgubun) & "' "
			end if
		end if

		if (FRectOnlineMwDiv <> "") then
			if (FRectOnlineMwDiv = "X") then
				AddSql = AddSql + " and isnull(i.mwdiv,'') not in ('M', 'W', 'U') "
			else
				AddSql = AddSql + " and isnull(i.mwdiv,'') = '" & CStr(FRectOnlineMwDiv) & "' "
			end if
		end if

		if (FRectCentermwdiv <> "") then
			if (FRectCentermwdiv = "X") then
				AddSql = AddSql + " and isnull(s.centermwdiv,'') not in ('M', 'W') "
			else
				AddSql = AddSql + " and isnull(s.centermwdiv,'') = '" & CStr(FRectCentermwdiv) & "' "
			end if
		end if

		if (FRectStockMwDiv <> "") then
			if (FRectStockMwDiv = "X") then
				AddSql = AddSql + " and IsNull(L.LastMwdiv, '') not in ('M', 'W') "
			else
				AddSql = AddSql + " and isnull(L.LastMwdiv,'') = '" & CStr(FRectStockMwDiv) & "' "
			end if
		end if
		if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				AddSql = AddSql + " 	and IsNull(p.tplcompanyid, '') = '' "
			else
				AddSql = AddSql + " 	and IsNull(p.tplcompanyid, '') = '" + CStr(FtplGubun) + "' "
			end if
		end if

		if (FRectBrandPurchaseType <> "") then
			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				AddSql = AddSql + " 	and pp.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				AddSql = AddSql & " 	and pp.purchasetype in ('3','5','6')"
			else
				AddSql = AddSql + " 	and pp.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		sqlStr = " select d.imakerid "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m with (nolock)"
		sqlStr = sqlStr + " Join [db_storage].[dbo].tbl_acount_storage_detail d with (nolock)"
		sqlStr = sqlStr + "     on m.code=d.mastercode"
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + "     on d.iitemgubun='10'"
		sqlStr = sqlStr + "     and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + "     on d.iitemgubun=s.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=s.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " left Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary L with (nolock)"
		sqlStr = sqlStr + "     on L.yyyymm=convert(Varchar(7),m.executeDt,21)"
		sqlStr = sqlStr + "     and  L.itemgubun=D.iitemgubun"
		sqlStr = sqlStr + "     and L.itemid=D.itemid"
		sqlStr = sqlStr + "     and L.itemoption=D.itemoption"
		sqlStr = sqlStr + " left Join db_summary.dbo.tbl_monthly_accumulated_Shopstock_summary SL with (nolock)"
		sqlStr = sqlStr + "     on SL.yyyymm=convert(Varchar(7),m.executeDt,21)"
		sqlStr = sqlStr + "     and SL.shopid=m.socid"
		sqlStr = sqlStr + "     and SL.itemgubun=D.iitemgubun"
		sqlStr = sqlStr + "     and SL.itemid=D.itemid"
		sqlStr = sqlStr + "     and SL.itemoption=D.itemoption"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner p with (nolock)"
	    sqlStr = sqlStr + " 	on m.socid=p.id"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr + " 	on d.imakerid=pp.id"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + AddSql
        sqlStr = sqlStr + " group by "
        sqlStr = sqlStr + " 	m.divcode "
        sqlStr = sqlStr + " 	,d.iitemgubun "
        sqlStr = sqlStr + " 	,m.socid "
        sqlStr = sqlStr + " 	,m.ipchulflag "
        sqlStr = sqlStr + " 	,d.imakerid "
        sqlStr = sqlStr + " 	,d.mwgubun "
        sqlStr = sqlStr + " 	,i.mwdiv "
        sqlStr = sqlStr + " 	,s.centermwdiv "
        sqlStr = sqlStr + " 	,L.LastMwdiv "
        sqlStr = sqlStr + " 	,SL.LstComm_cd "
        sqlStr = sqlStr + " 	,pp.purchasetype "

		tmpSql = sqlStr

		sqlStr = " select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sqlStr = sqlStr + " from ( "
		sqlStr = sqlStr + tmpSql
		sqlStr = sqlStr + " ) T "

		'response.write sqlStr & "<br>"
		''rsget.Open sqlStr,dbget,1
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = " select top "&FPageSize*FCurrPage&" "

		sqlStr = sqlStr + " m.divcode "
		sqlStr = sqlStr + " 	,d.iitemgubun "
		sqlStr = sqlStr + " 	,sum(d.sellcash*d.itemno) as sellcash "
		sqlStr = sqlStr + " 	,sum(d.suplycash*d.itemno) as suplycash "
		sqlStr = sqlStr + " 	,sum(d.buycash*d.itemno) as buycash "
		sqlStr = sqlStr + " 	,sum(d.itemno) as itemno "
		sqlStr = sqlStr + " 	,m.socid "
		sqlStr = sqlStr + " 	,m.ipchulflag "
		sqlStr = sqlStr + " 	,d.imakerid "
		sqlStr = sqlStr + " 	,d.mwgubun AS ipchulMwGubun "
		sqlStr = sqlStr + " 	,i.mwdiv AS onlineMwDiv "
		sqlStr = sqlStr + " 	,s.centermwdiv "
		sqlStr = sqlStr + " 	,L.LastMwdiv "
		sqlStr = sqlStr + " 	,SL.LstComm_cd "
		sqlStr = sqlStr + " 	,pp.purchasetype "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m with (nolock)"
		sqlStr = sqlStr + "     Join [db_storage].[dbo].tbl_acount_storage_detail d with (nolock)"
		sqlStr = sqlStr + "     on m.code=d.mastercode"
		sqlStr = sqlStr + " left Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary L with (nolock)"
		sqlStr = sqlStr + "     on L.yyyymm=convert(Varchar(7),m.executeDt,21)"
		sqlStr = sqlStr + "     and  L.itemgubun=D.iitemgubun"
		sqlStr = sqlStr + "     and L.itemid=D.itemid"
		sqlStr = sqlStr + "     and L.itemoption=D.itemoption"
		sqlStr = sqlStr + " left Join db_summary.dbo.tbl_monthly_accumulated_Shopstock_summary SL with (nolock)"
		sqlStr = sqlStr + "     on SL.yyyymm=convert(Varchar(7),m.executeDt,21)"
		sqlStr = sqlStr + "     and SL.shopid=m.socid"
		sqlStr = sqlStr + "     and SL.itemgubun=D.iitemgubun"
		sqlStr = sqlStr + "     and SL.itemid=D.itemid"
		sqlStr = sqlStr + "     and SL.itemoption=D.itemoption"
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + "     on d.iitemgubun='10'"
		sqlStr = sqlStr + "     and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + "     on d.iitemgubun=s.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=s.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=s.itemoption"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner p with (nolock)"
	    sqlStr = sqlStr + " 	on m.socid=p.id"
	    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr + " 	on d.imakerid=pp.id"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + AddSql
        sqlStr = sqlStr + " group by "
        sqlStr = sqlStr + " 	m.divcode "
        sqlStr = sqlStr + " 	,d.iitemgubun "
        sqlStr = sqlStr + " 	,m.socid "
        sqlStr = sqlStr + " 	,m.ipchulflag "
        sqlStr = sqlStr + " 	,d.imakerid "
        sqlStr = sqlStr + " 	,d.mwgubun "
        sqlStr = sqlStr + " 	,i.mwdiv "
        sqlStr = sqlStr + " 	,s.centermwdiv "
        sqlStr = sqlStr + " 	,L.LastMwdiv "
        sqlStr = sqlStr + " 	,SL.LstComm_cd "
        sqlStr = sqlStr + " 	,pp.purchasetype "
        sqlStr = sqlStr + " order by "
        sqlStr = sqlStr + " 	m.divcode "
        sqlStr = sqlStr + " 	,d.iitemgubun "
        sqlStr = sqlStr + " 	,m.socid "
        sqlStr = sqlStr + " 	,m.ipchulflag "
        sqlStr = sqlStr + " 	,d.imakerid "
        sqlStr = sqlStr + " 	,d.mwgubun "
        sqlStr = sqlStr + " 	,i.mwdiv "
        sqlStr = sqlStr + " 	,s.centermwdiv "
        sqlStr = sqlStr + " 	,L.LastMwdiv "
        sqlStr = sqlStr + " 	,SL.LstComm_cd "
        sqlStr = sqlStr + " 	,pp.purchasetype "

		''response.write sqlStr
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/06

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CAcountItemIpChulItem
                FItemList(i).Fdivcode        = rsget("divcode")
				FItemList(i).FSellCash       = rsget("sellcash")
				FItemList(i).FSuplycash      = rsget("suplycash")
				FItemList(i).FBuycash		 = rsget("buycash")
				FItemList(i).FItemNo         = rsget("itemno")
				FItemList(i).FSocID         = rsget("socid")
				FItemList(i).FItemgubun		= rsget("iitemgubun")
				FItemList(i).Fipchulflag	= rsget("ipchulflag")
				FItemList(i).Fimakerid	= rsget("imakerid")
                FItemList(i).FIpChulMwgubun = rsget("ipchulMwGubun")
                FItemList(i).FonlineMwDiv   = rsget("onlineMwDiv")
                FItemList(i).Fcentermwdiv   = rsget("centermwdiv")
                FItemList(i).FStockMwDiv    = rsget("LastMwdiv")
                FItemList(i).FStockShopComm_cd  = rsget("LstComm_cd")
				FItemList(i).Fpurchasetype    	= rsget("purchasetype")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	Private Sub Class_Initialize()
'		redim preserve FItemList(0)
		redim FItemList(0)
		FCurrPage = 1
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

function GetDivCodeName(divcode)
	dim resultDivCodeName

	if divcode="002" then
		resultDivCodeName = "위탁"
	elseif divcode="001" then
		resultDivCodeName = "매입"
	elseif divcode="003" then
		resultDivCodeName = "판촉출고"
	elseif divcode="004" then
		resultDivCodeName = "외부출고"
	elseif divcode="005" then
		resultDivCodeName = "협찬출고"
	elseif divcode="006" then
		resultDivCodeName = "B2B출고"
	elseif divcode="007" then
		resultDivCodeName = "기타출고"
	elseif divcode="101" then
		resultDivCodeName = "위탁출고"
	elseif divcode="999" then
		resultDivCodeName = "기타(정산않함)"
	elseif divcode="801" then
		resultDivCodeName = "Off매입"
	elseif divcode="802" then
		resultDivCodeName = "Off위탁"
	else
		resultDivCodeName = divcode
	end if

	GetDivCodeName=resultDivCodeName
end function

%>
