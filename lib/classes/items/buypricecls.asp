<%
response.write "buypricecls 사용중지 -> itemcls_2008.asp 로 변경요망"
	response.End

class CBuyPriceItem
	public FItemID
	public FItemName
	public FMakerID
	public FSellPrice
	public FSellVat
	public FMarginrate
	public FBuyPrice
	public FBuyvat
	public FVatInclude
	public FDisplayYn
	public FSellYn
	public FBaesongGB
	public FMarginDiv
	public FPojangYn
	public FLimitYn
	public FLimitDiv
	public FLimitNo
	public FLimitSold

	public FImageSmall
	public FMwDiv

	public Fsailyn
	public Fsailprice
	public Fsailsuplycash

	public Fisusing
	public FOptionName
	public FCurrNo
	public FSellCnt

	public Frealstock
	public Fipkumdiv5
	public Foffconfirmno
	public Fipkumdiv4
	public Fipkumdiv2
	public Foffjupno
	Public Fitemdiv


	public function GetCalcuMarginRate
		GetCalcuMarginRate = 0
		if FSellPrice<>0 then
			GetCalcuMarginRate = 100-CLng(FBuyPrice/FSellPrice*100*100)/100
		end if
	end function

	public function GetMwDivName
		if Fmwdiv="M" then
			GetMwDivName = "매입"
		elseif Fmwdiv="W" then
			GetMwDivName = "위탁"
		elseif Fmwdiv="U" then
			GetMwDivName = "업체"
		end if
	end function

	public function GetMwDivColor
		if Fmwdiv="M" then
			GetMwDivColor = "#FF0000"
		elseif Fmwdiv="W" then
			GetMwDivColor = "#0000FF"
		elseif Fmwdiv="U" then
			GetMwDivColor = "#000000"
		end if
	end function

	public function IsSoldOut()
		IsSoldOut = (FSellYn<>"Y") or (FDisplayYn<>"Y") or ((FLimitYn<>"N") and (FLimitNo-FLimitSold<1))
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CBuyPrice
	public FItemList()

	public FSearchItemid
	public FSearchItemName
	public FSearchDesigner
	'public FSearchDispYn
	public FSearchSellYn
	public FSearchLimitYn
	''public FSearchLimitDiv
	public FSearchBaedalDiv
	public FSearchusingyn
	public FSearchSailYn

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectCD1
	public FRectCD2
	public FRectCD3
	public FRectGubun

	public FRectOnlyTenBeasong
	public FRectMwDiv

	public FRectItemIDArr

	Private Sub Class_Initialize()
	redim  FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public function GetImageFolerName(byval i)
	    GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).FItemID)
		''GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public sub UpdateItemLimit(byval iitem)
		dim sqlStr
		sqlStr = "update [db_item].[dbo].tbl_item"
		sqlStr = sqlStr + " set limityn='" + CStr(iitem.FLimitYn) + "',"
		sqlStr = sqlStr + " limitno='" + CStr(iitem.FLimitNo) + "',"
		sqlStr = sqlStr + " limitsold='" + CStr(iitem.FLimitSold) + "'"

		sqlStr = sqlStr + " where itemid=" + CStr(iitem.FItemID)

		''response.write sqlStr

		rsget.Open sqlStr, dbget, 1
	end sub

	public sub UpdateItemView(byval iitem)
		dim sqlStr
		sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
		sqlStr = sqlStr + " set sellyn='" + CStr(iitem.FSellYn) + "'" + VbCrlf
		sqlStr = sqlStr + " ,pojangok='" + CStr(iitem.FPojangYn) + "'" + VbCrlf
		sqlStr = sqlStr + " ,isusing='" + CStr(iitem.Fisusing) + "'" + VbCrlf

		sqlStr = sqlStr + " where itemid=" + CStr(iitem.FItemID)

		''response.write sqlStr

		rsget.Open sqlStr, dbget, 1
	end sub

	public sub UpdateOneItem(byval iitem)
		dim sqlStr
		sqlStr = "update [db_item].[dbo].tbl_item"
		sqlStr = sqlStr + " set sellcash=" + CStr(iitem.FSellprice) + ","
		sqlStr = sqlStr + " sellvat=" + CStr(iitem.FSellvat) + ","
		sqlStr = sqlStr + " buycash=" + CStr(iitem.FBuyPrice) + ","
		sqlStr = sqlStr + " buyvat=" + CStr(iitem.FBuyVat) + ","
		sqlStr = sqlStr + " margin=" + CStr(iitem.FMarginrate) + ","
		sqlStr = sqlStr + " vatinclude='" + CStr(iitem.FVatInclude) + "',"
		sqlStr = sqlStr + " margindiv='" + CStr(iitem.FMarginDiv) + "'"

		sqlStr = sqlStr + " where itemid=" + CStr(iitem.FItemID)

		''response.write sqlStr

		rsget.Open sqlStr, dbget, 1
	end sub

	public sub getPrcList()
		dim sqlStr, sqlrect, i

		'// 추가 조건 쿼리
		if (FSearchItemid<>"") then
			sqlrect = sqlrect + " and i.itemid in (" + CStr(FSearchItemid) + ") "
		end if

		if (FSearchItemName<>"") then
			sqlrect = sqlrect + " and i.itemname like '%" + CStr(FSearchItemname) + "%'"
		end if

		if (FSearchDesigner<>"") then
			sqlrect = sqlrect + " and i.makerid = '" + CStr(FSearchDesigner) + "'"
		end if


		if (FSearchSellYn<>"") then
			sqlrect = sqlrect + " and i.sellyn = '" + CStr(FSearchSellYn) + "'"
		end if

		if (FSearchusingyn<>"") then
			sqlrect = sqlrect + " and i.isusing = '" + CStr(FSearchusingyn) + "'"
		end if
		

		if (FSearchsailyn <> "") then
			sqlrect = sqlrect + " and i.sailyn = '" + CStr(FSearchsailyn) + "'"
		end if

		if (FSearchLimitYn<>"") then
			sqlrect = sqlrect + " and i.limityn = '" + CStr(FSearchLimitYn) + "'"
		end if

		if (FSearchBaedalDiv<>"") then
			sqlrect = sqlrect + " and i.deliverytype = '" + CStr(FSearchBaedalDiv) + "'"
		end if

		if (FRectCD1 <> "") then
			sqlrect = sqlrect + " and i.itemserial_large = '" + CStr(FRectCD1) + "'"
		end if

		if (FRectCD2 <> "") then
			sqlrect = sqlrect + " and i.itemserial_mid = '" + CStr(FRectCD2) + "'"
		end if

		if (FRectCD3 <> "") then
			sqlrect = sqlrect + " and i.itemserial_small = '" + CStr(FRectCD3) + "'"
		end if

		if (FRectItemIDArr <> "") then
			sqlrect = sqlrect + " and i.itemid in (" + CStr(FRectItemIDArr) + ")"
		end if

		if (FRectGubun <> "") then
			if FRectGubun = "01" then
				sqlrect = sqlrect + " and i.sellcash < 2000"
			elseif FRectGubun = "02" then
				sqlrect = sqlrect + " and i.sellcash >= 2000"
				sqlrect = sqlrect + " and i.sellcash < 4000"
			elseif FRectGubun = "03" then
				sqlrect = sqlrect + " and i.sellcash >= 4000"
				sqlrect = sqlrect + " and i.sellcash < 6000"
			end if
		end if


		'// 결과 카운트
		sqlStr = "select count(i.itemid) as cnt from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.itemid<>0"

		rsget.Open sqlStr + sqlrect,dbget,1
		FTotalCount = rsget("cnt")
		rsget.close

		
		'// 본문 목록 접수
		sqlrect = sqlrect + " order by i.itemid desc"

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.itemid, i.itemdiv, i.itemname, i.makerid, i.buycash,"
		sqlStr = sqlStr + " i.sellcash, "
		sqlStr = sqlStr + " i.sellyn, i.deliverytype , i.vatinclude, "
		sqlStr = sqlStr + " i.pojangok, i.limityn,  i.limitno, i.limitsold, i.smallimage as imgsmall,"
		sqlStr = sqlStr + " i.mwdiv ,i.sailyn, i.isusing,"
		sqlStr = sqlStr + " i.sailprice, i.sailsuplycash "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.itemid<>0"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr + sqlrect,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CBuyPriceItem
				FItemList(i).FItemID    = rsget("itemid")
				FItemList(i).Fitemdiv    = rsget("itemdiv")
				FItemList(i).FItemName  = db2html(rsget("itemname"))
				FItemList(i).FMakerID   = rsget("makerid")
				FItemList(i).FSellPrice = rsget("sellcash")
				FItemList(i).FBuyPrice  = rsget("buycash")
				FItemList(i).FVatInclude= rsget("vatinclude")
				FItemList(i).FSellYn    = rsget("sellyn")
				FItemList(i).FBaesongGB = rsget("deliverytype")
				FItemList(i).FPojangYn = rsget("pojangok")

				FItemList(i).FLimitYn = rsget("limityn")
				FItemList(i).FLimitNo = rsget("limitno")
				FItemList(i).FLimitSold = rsget("limitsold")
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget("imgsmall")
				FItemList(i).FMWDiv		= rsget("mwdiv")
				FItemList(i).Fsailyn	= rsget("sailyn")
				FItemList(i).Fisusing	= rsget("isusing")

				FItemList(i).Fsailprice	= rsget("sailprice")
				FItemList(i).Fsailsuplycash	= rsget("sailsuplycash")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.close
	end Sub

	public sub GetLimitSoldOut()
		dim sqlStr
		dim wheredetail
		dim i

		sqlStr = "select count(i.itemid) as cnt" + vbcrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_const_day_stock s on i.itemid=s.itemid" + vbcrlf

		sqlStr = sqlStr + " where i.limityn='Y'" + vbcrlf
		if FRectCD1<>"" then
			sqlStr = sqlStr + " and i.itemserial_large='" + CStr(FRectCD1) + "'"
		end if

		if FRectMwDiv<>"" then
			sqlStr = sqlStr + " and i.mwdiv='" + CStr(FRectMwDiv) + "'"
		end if

		sqlStr = sqlStr + " and i.limitno-i.limitsold<1" + vbcrlf
		sqlStr = sqlStr + " and (i.sellyn='Y')" + vbcrlf
		sqlStr = sqlStr + " and IsNULL(s.sell7days,0)<1" + vbcrlf
		if (FSearchItemid<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FSearchItemid)
		end if

		if (FSearchItemName<>"") then
			sqlStr = sqlStr + " and i.itemname like '%" + CStr(FSearchItemname) + "%'"
		end if

		if (FSearchDesigner<>"") then
			sqlStr = sqlStr + " and i.makerid = '" + CStr(FSearchDesigner) + "'"
		end if


		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.close


		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.makerid,i.itemid,i.itemname,mwdiv,IsNULL(s.sell7days,0) as sellcnt,i.sellyn,"
		sqlStr = sqlStr + " i.limitno,i.limitsold,s.currno, i.smallimage,vw.optionname as codeview, " + vbcrlf
		sqlStr = sqlStr + " c.realstock, c.ipkumdiv5, c.offconfirmno, c.ipkumdiv4, c.ipkumdiv2, c.offjupno  " + vbcrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_const_day_stock s on i.itemid=s.itemid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option vw on s.itemoption = vw.itemoption and s.itemid=vw.itemid" + vbcrlf
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary c on s.itemid = c.itemid and s.itemoption = c.itemoption and c.itemgubun = '10' " + vbcrlf

		sqlStr = sqlStr + " where i.limityn='Y'" + vbcrlf
		if FRectCD1<>"" then
			sqlStr = sqlStr + " and i.itemserial_large='" + CStr(FRectCD1) + "'"
		end if

		if FRectMwDiv<>"" then
			sqlStr = sqlStr + " and i.mwdiv='" + CStr(FRectMwDiv) + "'"
		end if

		sqlStr = sqlStr + " and i.limitno-i.limitsold<1" + vbcrlf
		sqlStr = sqlStr + " and (i.sellyn='Y')" + vbcrlf
		sqlStr = sqlStr + " and IsNULL(s.sell7days,0)<1" + vbcrlf
		if (FSearchItemid<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FSearchItemid)
		end if

		if (FSearchItemName<>"") then
			sqlStr = sqlStr + " and i.itemname like '%" + CStr(FSearchItemname) + "%'"
		end if

		if (FSearchDesigner<>"") then
			sqlStr = sqlStr + " and i.makerid = '" + CStr(FSearchDesigner) + "'"
		end if
		sqlStr = sqlStr + " order by i.itemid"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0

                if not rsget.EOF then
                        rsget.absolutepage = FCurrPage

        		do until rsget.eof
        			set FItemList(i) = new CBuyPriceItem

        			FItemList(i).FItemID    = rsget("itemid")
        			FItemList(i).FItemName  = db2html(rsget("itemname"))
        			FItemList(i).FMakerID   = rsget("makerid")
        			FItemList(i).Fmwdiv = rsget("mwdiv")
        			FItemList(i).FSellCnt   = rsget("sellcnt")
        			FItemList(i).FSellYn    = rsget("sellyn")
        			FItemList(i).FLimitNo    = rsget("limitno")
        			FItemList(i).FLimitSold    = rsget("limitsold")
        			FItemList(i).FOptionName = rsget("codeview")
        			FItemList(i).FCurrNo = rsget("currno")
        			FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("smallimage")

        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
                                FItemList(i).Foffjupno      = rsget("offjupno")

        			rsget.movenext
        			i=i+1
        		loop
		end if
		rsget.close
	end Sub

	public sub GetDispYNSet()
		dim sqlStr
		dim wheredetail
		dim i
        
        response.write "MayBe  Not Using.."
        dbget.close()	:	response.End
        
		sqlStr = "select count(i.itemid) as cnt" + vbcrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + vbcrlf

		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_const_day_stock s"
		sqlStr = sqlStr + " on i.itemid=s.itemid "

		sqlStr = sqlStr + " where i.dispyn='Y'" + vbcrlf
		sqlStr = sqlStr + " and i.sellyn='N'" + vbcrlf
		sqlStr = sqlStr + " and i.mwdiv<>'U'" + vbcrlf
		sqlStr = sqlStr + " and IsNULL(s.sell7days,0)<1" + vbcrlf
		if (FSearchItemid<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FSearchItemid)
		end if

		if (FSearchItemName<>"") then
			sqlStr = sqlStr + " and i.itemname like '%" + CStr(FSearchItemname) + "%'"
		end if

		if (FSearchDesigner<>"") then
			sqlStr = sqlStr + " and i.makerid = '" + CStr(FSearchDesigner) + "'"
		end if

		if FRectCD1<>"" then
			sqlStr = sqlStr + " and i.itemserial_large='" + CStr(FRectCD1) + "'"
		end if

		if FRectMwDiv<>"" then
			sqlStr = sqlStr + " and i.mwdiv='" + CStr(FRectMwDiv) + "'"
		end if
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.close


		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " i.makerid,i.itemid,i.itemname,i.mwdiv,IsNULL(s.sell7days,0) as monthlycnt, s.currno, i.sellyn,"
		sqlStr = sqlStr + " i.limitno,i.limitsold, i.smallimage as imgsmall," + vbcrlf
		sqlStr = sqlStr + " v.optionname as codeview, "
		sqlStr = sqlStr + " c.realstock, c.ipkumdiv5, c.offconfirmno, c.ipkumdiv4, c.ipkumdiv2, c.offjupno  " + vbcrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + vbcrlf

		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_const_day_stock s"
		sqlStr = sqlStr + " on i.itemid=s.itemid "
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v"
		sqlStr = sqlStr + " on s.itemoption=v.itemoption and s.itemid=v.itemid"
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary c on s.itemid = c.itemid and s.itemoption = c.itemoption and c.itemgubun = '10' " + vbcrlf

		sqlStr = sqlStr + " where i.dispyn='Y'" + vbcrlf
		sqlStr = sqlStr + " and i.sellyn='N'" + vbcrlf
		sqlStr = sqlStr + " and i.mwdiv<>'U'" + vbcrlf
		sqlStr = sqlStr + " and IsNULL(s.sell7days,0)<1" + vbcrlf
		if (FSearchItemid<>"") then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FSearchItemid)
		end if

		if (FSearchItemName<>"") then
			sqlStr = sqlStr + " and i.itemname like '%" + CStr(FSearchItemname) + "%'"
		end if

		if FRectCD1<>"" then
			sqlStr = sqlStr + " and i.itemserial_large='" + CStr(FRectCD1) + "'"
		end if

		if FRectMwDiv<>"" then
			sqlStr = sqlStr + " and i.mwdiv='" + CStr(FRectMwDiv) + "'"
		end if

		if (FSearchDesigner<>"") then
			sqlStr = sqlStr + " and i.makerid = '" + CStr(FSearchDesigner) + "'"
		end if
		sqlStr = sqlStr + " order by i.itemid"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0

                if not rsget.EOF then
                        rsget.absolutepage = FCurrPage

        		do until rsget.eof
        			set FItemList(i) = new CBuyPriceItem

        			FItemList(i).FItemID    = rsget("itemid")
        			FItemList(i).FItemName  = db2html(rsget("itemname"))
        			FItemList(i).FMakerID   = rsget("makerid")
        			FItemList(i).Fmwdiv = rsget("mwdiv")
        			FItemList(i).FSellCnt   = rsget("monthlycnt")
        			'FItemList(i).FDisplayYn = rsget("dispyn")
        			FItemList(i).FSellYn    = rsget("sellyn")
        			FItemList(i).FLimitNo    = rsget("limitno")
        			FItemList(i).FLimitSold    = rsget("limitsold")
        			FItemList(i).FOptionName = rsget("codeview")
        			FItemList(i).FCurrNo = rsget("currno")
        			FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("imgsmall")

        			FItemList(i).Frealstock     = rsget("realstock")
        			FItemList(i).Fipkumdiv5     = rsget("ipkumdiv5")
        			FItemList(i).Foffconfirmno  = rsget("offconfirmno")
        			FItemList(i).Fipkumdiv4     = rsget("ipkumdiv4")
        			FItemList(i).Fipkumdiv2     = rsget("ipkumdiv2")
                                FItemList(i).Foffjupno      = rsget("offjupno")

        			rsget.movenext
        			i=i+1
        		loop
		end if
		rsget.close
	end Sub

	

	public sub UpdateItemSellDispYN(byval iitem)
		dim sqlStr
		sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
		sqlStr = sqlStr + " set sellyn='" + CStr(iitem.FSellYn) + "'" + VbCrlf
		sqlStr = sqlStr + " where itemid=" + CStr(iitem.FItemID)

		''response.write sqlStr

		rsget.Open sqlStr, dbget, 1
	end sub

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