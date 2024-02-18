<%

class CUpcheOrderOneItem

	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fipgono
	public Freipgono
	public Ftotipgono
	public Foffchulgono
	public Foffrechulgono
	public Fetcchulgono
	public Fetcrechulgono
	public Ftotchulgono
	public Fsellno
	public Fresellno
	public Ftotsellno
	public Ferrcsno
	public Ferrbaditemno
	public Ferrrealcheckno
	public Ferretcno
	public Ftoterrno
	public Ftotsysstock
	public Favailsysstock
	public Frealstock
	public Fsell7days
	public Foffchulgo7days
	public Fipkumdiv5
	public Fipkumdiv4
	public Fipkumdiv2
	public Foffconfirmno
	public Foffjupno
	public Frequireno
	public Fshortageno
	public Fpreorderno
	public Fpreordernofix
	public Foffsellno
	public Fmaxsellday
	public Fregdate
	public Flastupdate

	public FMakerid
	public Fitemname
	public Fitemoptionname
	public Fisusing
	public FImageSmall
	public FImageList

	public Fsellyn
	public Flimityn
	public Flimitno
	public Flimitsold

	public Fdanjongyn

	public Foptioncount
	public Foptionlimitno
	public Foptionlimitsold

	public FSellcash
	public Fbuycash
	public FmwDiv
	public Foptusing


	public FOffImgMain
	public FOffImgList
	public FOffImgSmall

	public Fextbarcode


	public FOFFdefaultmargin
	public FOFFdefaultsuplymargin
	public FOFFchargediv

	public FOffMwMargin

	'public FSocName
	'public FSocNameKor
	'public Fextbarcode
	'public FmakerMargin
	'public FshopMargin


	'public Flastrealdate
	'public Flastrealno
	'public Fipno
	'public Fchulno
	'public Fsellno
	'public Fcurrno

	public function IsOffContractExist()
		IsOffContractExist = Not (FOffMwMargin = "")
	end function

	public function GetOffContractMWDiv()
		dim tmpArr

		GetOffContractMWDiv = ""

		tmpArr = Split(FOffMwMargin, "_")
		if UBound(tmpArr) = 2 then
			GetOffContractMWDiv = tmpArr(0)
		end if
	end function

	public function GetOffContractMargin()
		dim tmpArr

		GetOffContractMargin = ""

		'// M_45_0
		tmpArr = Split(FOffMwMargin, "_")
		if UBound(tmpArr) = 2 then
			GetOffContractMargin = tmpArr(1)
		end if
	end function

	public function GetOffContractBuycash()
		dim tmpArr

		GetOffContractBuycash = FBuycash

		'// M_45_0
		tmpArr = Split(FOffMwMargin, "_")
		if UBound(tmpArr) = 2 then
			if tmpArr(1) <> 0 and tmpArr(2) = 0 then
				'// 마진적용
				GetOffContractBuycash = CLng(Fsellcash * (100 - tmpArr(1)) / 100)
			elseif tmpArr(2) <> 0 then
				'//상품매입가
				GetOffContractBuycash = tmpArr(2)
			end if
		end if
	end function

	public function GetOffContractCenterMW()
		dim tmpArr

		GetOffContractCenterMW = "U"

		'// M_45_0
		tmpArr = Split(FOffMwMargin, "_")
		if UBound(tmpArr) = 2 then
			GetOffContractCenterMW = tmpArr(0)
		end if
	end function

    ''재고파악재고
	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function

    ''금일 상품준비수량
	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function


    ''출고이전 필요수량(접수,결제완료..)
    public function GetReqNotChulgoNo()
		GetReqNotChulgoNo = Fipkumdiv5 + Foffconfirmno + Fipkumdiv4 + Fipkumdiv2 + Foffjupno
	end function

	''N일치 부족수량
    public function GetNdayShortageNo(nday)
		GetNdayShortageNo = Fshortageno + CLng(Frequireno*(nday-7)/7)
	end function

	public function getMwDivName()
		if Fitemgubun="10" then
			if FmwDiv="M" then
				getMwDivName = "매입"
			elseif FmwDiv="W" then
				getMwDivName = "위탁"
			elseif FmwDiv="U" then
				getMwDivName = "업체"
			end if
		else
			if FOFFchargediv="2" then
				getMwDivName = "텐위"
			elseif FOFFchargediv="4" then
				getMwDivName = "텐매"
			elseif FOFFchargediv="6" then
				getMwDivName = "업위"
			elseif FOFFchargediv="8" then
				getMwDivName = "업매"
			end if
		end if
	end function

	public function getMwDivColor()
		if Fitemgubun="10" then
			if FmwDiv="M" then
				getMwDivColor = "#CC2222"
			elseif FmwDiv="W" then
				getMwDivColor = "#2222CC"
			end if
		else
			if FOFFchargediv="2" then
				getMwDivColor = "#FF4444"
			elseif FOFFchargediv="4" then
				getMwDivColor = "#44FF44"
			elseif FOFFchargediv="6" then
				getMwDivColor = "#4444FF"
			elseif FOFFchargediv="8" then
				getMwDivColor = "#FF44FF"
			end if
		end if
	end function

	public function getSoldOutColor()
		if (Foptusing="N") or (IsSoldOut) then
			getSoldOutColor = "#AAAAAA"
		else
			getSoldOutColor = "#000000"
		end if
	end function

	public function IsSoldOut()
		IsSoldOut = (Fsellyn="N") or ((Flimityn="Y") and (Flimitno-Flimitsold<1))
	end function

	public function IsUpchebeasongItem()
		if Fitemgubun="90" or  Fitemgubun="80" then
			IsUpchebeasongItem = true
		else
			if FdeliveryType="2" or FdeliveryType="5" then
				IsUpchebeasongItem = true
			else
				IsUpchebeasongItem = false
			end if
		end if
	end function

	public function getLimitNo()
		getLimitNo =0
		if (Flimityn="Y") then
			getLimitNo = Flimitno-Flimitsold
		end if

		if getLimitNo<1 then getLimitNo=0
	end function

	public function getOptionLimitNo()
		getOptionLimitNo =0
		if ((Flimityn="Y") and (Foptusing="Y")) then
			getOptionLimitNo = Foptionlimitno-Foptionlimitsold
		end if

		if getOptionLimitNo<1 then getOptionLimitNo=0

		if (Foptioncount < 1) then
		    getOptionLimitNo = getLimitNo
		end if
	end function

	public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = FimageSmall
		else
			GetImageSmall = FOffImgSmall
		end if
	end function

	public function GetFranchiseSuplycash()
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetFranchiseSuplycash = Fshopsuplycash
		else
			GetFranchiseSuplycash = CLng(Fshopitemprice * (100-FshopMargin)/100)
		end if
	end function

	public function GetOfflineSuplycash()
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetOfflineSuplycash = Fshopsuplycash
		else
			GetOfflineSuplycash = CLng(Fshopitemprice * (100-FshopMargin)/100)
		end if
	end function

	public function GetFranchiseBuycash()
		dim ibuycash
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetFranchiseBuycash = Fshopsuplycash
		else
			ibuycash = CLng(Fshopitemprice * (100-FmakerMargin)/100)

			if (FOnlinebuycash<>0) and (ibuycash>FOnlinebuycash) then
				GetFranchiseBuycash = FOnlinebuycash
			else
				GetFranchiseBuycash = ibuycash
			end if
		end if
	end function

	public function GetOfflineBuycash()
		dim ibuycash
		if (Fitemgubun="70") or (Fitemgubun="80") then
			GetOfflineBuycash = Fshopsuplycash
		else
			ibuycash = CLng(Fshopitemprice * (100-FmakerMargin)/100)

			if (FOnlinebuycash<>0) and (ibuycash>FOnlinebuycash) then
				GetOfflineBuycash = FOnlinebuycash
			else
				GetOfflineBuycash = ibuycash
			end if
		end if
	end function

	public function GetChargeMaySuplycash()
			if Fshopsuplycash<>0 then
				GetChargeMaySuplycash = Fshopsuplycash
			else
				GetChargeMaySuplycash = CLng(Fshopitemprice * 0.65)
			end if

		'if Fchargeid="10x10" then
		'	if Fshopsuplycash<>0 then
		'		GetChargeMaySuplycash = Fshopsuplycash
		'	else
		'		GetChargeMaySuplycash = CLng(Fshopitemprice * 0.7)
		'	end if
		'else
		'	if Fshopsuplycash<>0 then
		'		GetChargeMaySuplycash = Fshopsuplycash
		'	else
		'		GetChargeMaySuplycash = Fshopitemprice
		'	end if
		'end if
	end function

	public function GetBarCode()
		GetBarCode = Fitemgubun & Format00(6,Fitemid) & Fitemoption
		if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

	public function GetBarCodeBoldStr()
		GetBarCodeBoldStr = Fitemgubun & "-" & Format00(6,Fitemid) & "-" & Fitemoption
		if (Fitemid >= 1000000) then
    		GetBarCodeBoldStr = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CUpcheOrderItem

	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectDesigner
	public FRectItemgubun
	public FRectItemId
	public FRectItemOption

	public FRectBarCode
	public FRectOnlyOffLine
	public FRectOnlyNotIpGo

	public FRectIpChulId
	public FRectIpGoOnly

	public FRectOnly90
	public FRectOnly10
	public FRectOnlyNew
	public FRectOnOffGubun
	public FRectOnlyTenbeasong
	public FRectArraryBarCode

	public FRectAllPriceDiff
	public FRectOnly10beasong

	public FRectOrder
	public FRectDesignerjungsangubun
	public FRectShopid

	public FRectNoSearchUpcheBeasong
	public FRectNoSearchNotusingItem
	public FRectNoSearchNotusingItemOption
	public FRectNoSearchDanjong
	public FRectNoSearchSoldoutover7days
	public FRectBarCodeArr

	public FRectImageView
	public FRectDispShowAll
	public FRectAdminView

	public FRectUpchebeasongInclude
	public FRectShortage7days
	public FRectShortage14days
    public FRectShortageRealStock

	public function GetImageFolerName(byval i)
		GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function


	public function GetOffShopItemList()
		dim sqlStr,i, sqlsearch

		if FRectItemgubun<>"" then
			sqlsearch = sqlsearch & " and i.itemgubun = '"& FRectItemgubun &"'"
		end if
		if FRectItemid<>"" then
			sqlsearch = sqlsearch & " and i.shopitemid = '"& FRectItemid &"'"
		end if
		if FRectItemOption<>"" then
			sqlsearch = sqlsearch & " and i.itemoption = '"& FRectItemOption &"'"
		end if
		if FRectDesigner<>"" then
			sqlsearch = sqlsearch & " and i.makerid = '"& FRectDesigner &"'"
		end if
		if FRectNoSearchNotusingItem="on" then
			sqlsearch = sqlsearch + " and i.isusing='Y'"
		end if

		sqlStr = " select count(i.shopitemid) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " where i.itemgubun<>'10' " & sqlsearch

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " i.itemgubun, i.shopitemid, i.itemoption, i.offimgmain, i.offimglist, i.offimgsmall"
		sqlStr = sqlStr + " , i.makerid, i.shopitemname, i.shopitemoptionname,"
		sqlStr = sqlStr + " i.shopitemprice, i.shopsuplycash, i.isusing, i.regdate, i.extbarcode,"
		sqlStr = sqlStr + " d.defaultmargin, d.defaultsuplymargin, d.chargediv,"

		sqlStr = sqlStr + " IsNull(s.ipgono,0) as ipgono, IsNULL(s.reipgono,0) as reipgono, "
		sqlStr = sqlStr + " IsNull(s.totipgono,0) as totipgono, IsNull(s.offchulgono,0) as offchulgono, "
        sqlStr = sqlStr + " IsNull(s.offrechulgono,0) as offrechulgono, IsNull(s.etcchulgono,0) as etcchulgono, "
        sqlStr = sqlStr + " IsNull(s.etcrechulgono,0) as etcrechulgono, IsNull(s.totchulgono,0) as totchulgono, "
        sqlStr = sqlStr + " IsNull(s.sellno,0) as sellno, IsNull(s.resellno,0) as resellno, "
   		sqlStr = sqlStr + " IsNull(s.totsellno,0) as totsellno, IsNull(s.errcsno,0) as errcsno, "
        sqlStr = sqlStr + " IsNull(s.errbaditemno,0) as errbaditemno, IsNull(s.errrealcheckno,0) as errrealcheckno, "
        sqlStr = sqlStr + " IsNull(s.erretcno,0) as erretcno, IsNull(s.toterrno,0) as toterrno, "
    	sqlStr = sqlStr + " IsNull(s.totsysstock,0) as totsysstock, IsNull(s.availsysstock,0) as availsysstock, "
        sqlStr = sqlStr + " IsNull(s.realstock,0) as realstock, IsNull(s.sell7days,0) as sell7days, "
     	sqlStr = sqlStr + " IsNull(s.offchulgo7days,0) as offchulgo7days, IsNull(s.ipkumdiv5,0) as ipkumdiv5, "
        sqlStr = sqlStr + " IsNull(s.ipkumdiv4,0) as ipkumdiv4, IsNull(s.ipkumdiv2,0) as ipkumdiv2, "
        sqlStr = sqlStr + " IsNull(s.offconfirmno,0) as offconfirmno, IsNull(s.offjupno,0) as offjupno, "
  		sqlStr = sqlStr + " IsNull(s.requireno,0) as requireno, IsNull(s.shortageno,0) as  shortageno, "
        sqlStr = sqlStr + " IsNull(s.preorderno,0) as preorderno, IsNull(s.offsellno,0) as  offsellno, "
        sqlStr = sqlStr + " IsNull(s.maxsellday,1) as maxsellday, s.regdate, s.lastupdate"

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " 	on i.makerid=d.makerid"
		sqlStr = sqlStr + " 	and d.shopid='streetshop000'"
		sqlStr = sqlStr + " 	left join [db_summary].[dbo].tbl_current_logisstock_summary s"
		sqlStr = sqlStr + " 	on i.itemgubun=s.itemgubun"
		sqlStr = sqlStr + " 	and i.shopitemid=s.itemid"
		sqlStr = sqlStr + " 	and i.itemoption=s.itemoption"

		sqlStr = sqlStr + " where i.itemgubun<>'10' " & sqlsearch
		sqlStr = sqlStr + " order by i.itemgubun desc, i.shopitemid desc"

		'response.write sqlStr & "<Br>"
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
				set FItemList(i) = new CUpcheOrderOneItem
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fsellcash     = rsget("shopitemprice")
				FItemList(i).Fbuycash     = rsget("shopsuplycash")
				FItemList(i).Fisusing           = rsget("isusing")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).Fextbarcode 		= rsget("extbarcode")

				FItemList(i).FOFFdefaultmargin		= rsget("defaultmargin")
				FItemList(i).FOFFdefaultsuplymargin	= rsget("defaultsuplymargin")

				if (FItemList(i).Fbuycash=0) and (FItemList(i).FOFFdefaultmargin<>0) then
					FItemList(i).Fbuycash = CLng(FItemList(i).Fsellcash * (100-FItemList(i).FOFFdefaultmargin) /100)
				end if

				FItemList(i).FOFFchargediv	= rsget("chargediv")


				FItemList(i).Fipgono         = rsget("ipgono")
				FItemList(i).Freipgono       = rsget("reipgono")
				FItemList(i).Ftotipgono      = rsget("totipgono")
				FItemList(i).Foffchulgono    = rsget("offchulgono")
				FItemList(i).Foffrechulgono  = rsget("offrechulgono")
				FItemList(i).Fetcchulgono    = rsget("etcchulgono")
				FItemList(i).Fetcrechulgono  = rsget("etcrechulgono")
				FItemList(i).Ftotchulgono    = rsget("totchulgono")
				FItemList(i).Fsellno         = rsget("sellno") * -1
				FItemList(i).Fresellno       = rsget("resellno") * -1
				FItemList(i).Ftotsellno      = rsget("totsellno") * -1
				FItemList(i).Ferrcsno        = rsget("errcsno")
				FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
				FItemList(i).Ferretcno       = rsget("erretcno")
				FItemList(i).Ftoterrno       = rsget("toterrno")
				FItemList(i).Ftotsysstock    = rsget("totsysstock")
				FItemList(i).Favailsysstock  = rsget("availsysstock")
				FItemList(i).Frealstock      = rsget("realstock")
				FItemList(i).Fsell7days      = rsget("sell7days")
				FItemList(i).Foffchulgo7days = rsget("offchulgo7days")
				FItemList(i).Fipkumdiv5      = rsget("ipkumdiv5")
				FItemList(i).Fipkumdiv4      = rsget("ipkumdiv4")
				FItemList(i).Fipkumdiv2      = rsget("ipkumdiv2")
				FItemList(i).Foffconfirmno   = rsget("offconfirmno")
				FItemList(i).Foffjupno       = rsget("offjupno")
				FItemList(i).Frequireno      = rsget("requireno")
				FItemList(i).Fshortageno     = rsget("shortageno")
				FItemList(i).Fpreorderno     = rsget("preorderno")
				FItemList(i).Foffsellno      = rsget("offsellno")
				FItemList(i).Fmaxsellday     = rsget("maxsellday")
				FItemList(i).Fregdate        = rsget("regdate")
				FItemList(i).Flastupdate     = rsget("lastupdate")

				FItemList(i).FOffimgMain	= rsget("offimgmain")
					if isnull(FItemList(i).FOffimgMain) then FItemList(i).FOffimgMain=""
				FItemList(i).FOffimgList	= rsget("offimglist")
					if isnull(FItemList(i).FOffimgList) then FItemList(i).FOffimgList=""
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
					if isnull(FItemList(i).FOffimgSmall) then FItemList(i).FOffimgSmall=""

				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgSmall

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function


	public sub GetOnLineJumunByBrand()
		dim sqlStr,i

		sqlStr = " select count(i.itemid) as cnt "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option v "
		sqlStr = sqlStr + " 	on i.itemid=v.itemid"
		sqlStr = sqlStr + " 	left join [db_summary].[dbo].tbl_current_logisstock_summary s "
		sqlStr = sqlStr + " 	on s.itemgubun = '10' and i.itemid=s.itemid  and IsNULL(v.itemoption,'0000')=s.itemoption"
		sqlStr = sqlStr + " where i.itemid<>0"

		''7일 초과부족수량
		if FRectShortage7days="on" then
		    sqlStr = sqlstr + " and s.shortageno<0"
		end if

		''14일 초과부족 수량
        if FRectShortage14days="on" then
		    sqlStr = sqlstr + " and (s.shortageno + s.requireno)<0"
		end if

        ''현재고 N개 이하
        if FRectShortageRealStock="on" then
		    sqlStr = sqlstr + " and s.realstock<=5"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if

		if FRectItemgubun<>"" then
			sqlStr = sqlStr + " and i.itemgubun='" + CStr(FRectItemgubun) + "'"
		end if

		if FRectItemid<>"" then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemid) + ""
		end if

		if ((FRectItemOption <> "") and (FRectItemOption <> "0000")) then
			sqlStr = sqlStr + " and v.itemoption='" + CStr(FRectItemOption) + "'"
		end if

		if FRectNoSearchUpcheBeasong="on" then
			sqlStr = sqlStr + " and i.mwdiv<>'U'"
		end if

		if FRectNoSearchNotusingItem="on" then
			sqlStr = sqlStr + " and i.isusing ='Y'"
		end if

		if FRectNoSearchNotusingItemOption="on" then
			sqlStr = sqlStr + " and ((v.isusing ='Y') or (v.isusing is NULL))"
		end if

		if (FRectNoSearchDanjong="on" and FRectNoSearchSoldoutover7days="on") then
			sqlStr = sqlStr + " and i.danjongyn = 'N'"
		elseif (FRectNoSearchDanjong="on" and FRectNoSearchSoldoutover7days="") then
			sqlStr = sqlStr + " and i.danjongyn <> 'Y'"
		elseif (FRectNoSearchDanjong="" and FRectNoSearchSoldoutover7days="on") then
			sqlStr = sqlStr + " and i.danjongyn <> 'S'"
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " '10' as itemgubun, i.itemid, IsNULL(v.itemoption,'0000') as itemoption, "
		sqlStr = sqlStr + " i.makerid, i.itemname, IsNULL(v.optionname,'') as itemoptionname, IsNULL(v.isusing,'Y') as optusing,"
		sqlStr = sqlStr + " i.sellcash + IsNULL(v.optaddprice,0) as sellcash, i.buycash + IsNULL(v.optaddbuyprice,0) as buycash, i.smallimage as imgsmall, "
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, i.mwdiv, i.isusing, i.danjongyn, i.optioncnt, v.optlimitno, v.optlimitsold, "

		sqlStr = sqlStr + " IsNull(s.ipgono,0) as ipgono, IsNULL(s.reipgono,0) as reipgono, "
		sqlStr = sqlStr + " IsNull(s.totipgono,0) as totipgono, IsNull(s.offchulgono,0) as offchulgono, "
        sqlStr = sqlStr + " IsNull(s.offrechulgono,0) as offrechulgono, IsNull(s.etcchulgono,0) as etcchulgono, "
        sqlStr = sqlStr + " IsNull(s.etcrechulgono,0) as etcrechulgono, IsNull(s.totchulgono,0) as totchulgono, "
        sqlStr = sqlStr + " IsNull(s.sellno,0) as sellno, IsNull(s.resellno,0) as resellno, "
   		sqlStr = sqlStr + " IsNull(s.totsellno,0) as totsellno, IsNull(s.errcsno,0) as errcsno, "
        sqlStr = sqlStr + " IsNull(s.errbaditemno,0) as errbaditemno, IsNull(s.errrealcheckno,0) as errrealcheckno, "
        sqlStr = sqlStr + " IsNull(s.erretcno,0) as erretcno, IsNull(s.toterrno,0) as toterrno, "
    	sqlStr = sqlStr + " IsNull(s.totsysstock,0) as totsysstock, IsNull(s.availsysstock,0) as availsysstock, "
        sqlStr = sqlStr + " IsNull(s.realstock,0) as realstock, IsNull(s.sell7days,0) as sell7days, "
     	sqlStr = sqlStr + " IsNull(s.offchulgo7days,0) as offchulgo7days, IsNull(s.ipkumdiv5,0) as ipkumdiv5, "
        sqlStr = sqlStr + " IsNull(s.ipkumdiv4,0) as ipkumdiv4, IsNull(s.ipkumdiv2,0) as ipkumdiv2, "
        sqlStr = sqlStr + " IsNull(s.offconfirmno,0) as offconfirmno, IsNull(s.offjupno,0) as offjupno, "
  		sqlStr = sqlStr + " IsNull(s.requireno,0) as requireno, IsNull(s.shortageno,0) as  shortageno, "
        sqlStr = sqlStr + " IsNull(s.preorderno,0) as preorderno, IsNull(s.preordernofix,0) as  preordernofix, "
        sqlStr = sqlStr + " IsNull(s.offsellno,0) as  offsellno, "
        sqlStr = sqlStr + " IsNull(s.maxsellday,1) as maxsellday, s.regdate, s.lastupdate,"
		sqlStr = sqlStr + " (case when i.mwdiv = 'U' then [db_shop].[dbo].[uf_GetCenterMWDivMargin](i.itemid) else '' end) as offmwmargin"
'11111111111111111111111111
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option v "
		sqlStr = sqlStr + " 	on i.itemid=v.itemid "
		sqlStr = sqlStr + " 	left join [db_summary].[dbo].tbl_current_logisstock_summary s"
		sqlStr = sqlStr + " 	on s.itemgubun = '10' and i.itemid=s.itemid and IsNULL(v.itemoption,'0000')=s.itemoption"


		sqlStr = sqlStr + " where i.itemid<>0"

		if FRectShortage7days="on" then
		    sqlStr = sqlstr + " and s.shortageno<0"
		end if

		if FRectShortage14days="on" then
		    sqlStr = sqlstr + " and (s.shortageno + s.requireno)<0"
		end if

		''현재고 N개 이하
        if FRectShortageRealStock="on" then
		    sqlStr = sqlstr + " and s.realstock<=5"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if

		if FRectItemgubun<>"" then
			sqlStr = sqlStr + " and i.itemgubun=" + CStr(FRectItemgubun) + ""
		end if

		if FRectItemid<>"" then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemid) + ""
		end if

		if ((FRectItemOption <> "") and (FRectItemOption <> "0000")) then
			sqlStr = sqlStr + " and v.itemoption='" + CStr(FRectItemOption) + "'"
		end if

		if FRectNoSearchUpcheBeasong="on" then
			sqlStr = sqlStr + " and i.mwdiv<>'U'"
		end if

		if FRectNoSearchNotusingItem="on" then
			sqlStr = sqlStr + " and i.isusing ='Y'"
		end if

		if FRectNoSearchNotusingItemOption="on" then
			sqlStr = sqlStr + " and ((v.isusing ='Y') or (v.isusing is NULL))"
		end if

		if (FRectNoSearchDanjong="on" and FRectNoSearchSoldoutover7days="on") then
			sqlStr = sqlStr + " and i.danjongyn = 'N'"
		elseif (FRectNoSearchDanjong="on" and FRectNoSearchSoldoutover7days="") then
			sqlStr = sqlStr + " and i.danjongyn <> 'Y'"
		elseif (FRectNoSearchDanjong="" and FRectNoSearchSoldoutover7days="on") then
			sqlStr = sqlStr + " and i.danjongyn <> 'S'"
		end if

		sqlStr = sqlStr + " order by i.itemid desc, s.itemoption"
		'response.write sqlStr

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1


		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CUpcheOrderOneItem
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fitemid            = rsget("itemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fitemname          = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
				FItemList(i).Fsellcash          = rsget("sellcash")
				FItemList(i).Fbuycash           = rsget("buycash")

				FItemList(i).FimageSmall     = rsget("imgsmall")
				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
				end if

                FItemList(i).Fsellyn               = rsget("sellyn")
				FItemList(i).Flimityn              = rsget("limityn")
				FItemList(i).Flimitno              = rsget("limitno")
				FItemList(i).Flimitsold            = rsget("limitsold")
				FItemList(i).Foptioncount          = rsget("optioncnt")
				FItemList(i).Foptionlimitno        = rsget("optlimitno")
				FItemList(i).Foptionlimitsold      = rsget("optlimitsold")
				FItemList(i).FmwDiv				   = rsget("mwdiv")
				FItemList(i).Foptusing		    	= rsget("optusing")
				FItemList(i).Fisusing				= rsget("isusing")
				FItemList(i).Fdanjongyn				= rsget("danjongyn")


				FItemList(i).Fipgono         = rsget("ipgono")
				FItemList(i).Freipgono       = rsget("reipgono")
				FItemList(i).Ftotipgono      = rsget("totipgono")
				FItemList(i).Foffchulgono    = rsget("offchulgono")
				FItemList(i).Foffrechulgono  = rsget("offrechulgono")
				FItemList(i).Fetcchulgono    = rsget("etcchulgono")
				FItemList(i).Fetcrechulgono  = rsget("etcrechulgono")
				FItemList(i).Ftotchulgono    = rsget("totchulgono")
				FItemList(i).Fsellno         = rsget("sellno") * -1
				FItemList(i).Fresellno       = rsget("resellno") * -1
				FItemList(i).Ftotsellno      = rsget("totsellno") * -1
				FItemList(i).Ferrcsno        = rsget("errcsno")
				FItemList(i).Ferrbaditemno   = rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
				FItemList(i).Ferretcno       = rsget("erretcno")
				FItemList(i).Ftoterrno       = rsget("toterrno")
				FItemList(i).Ftotsysstock    = rsget("totsysstock")
				FItemList(i).Favailsysstock  = rsget("availsysstock")
				FItemList(i).Frealstock      = rsget("realstock")
				FItemList(i).Fsell7days      = rsget("sell7days")
				FItemList(i).Foffchulgo7days = rsget("offchulgo7days")
				FItemList(i).Fipkumdiv5      = rsget("ipkumdiv5")
				FItemList(i).Fipkumdiv4      = rsget("ipkumdiv4")
				FItemList(i).Fipkumdiv2      = rsget("ipkumdiv2")
				FItemList(i).Foffconfirmno   = rsget("offconfirmno")
				FItemList(i).Foffjupno       = rsget("offjupno")
				FItemList(i).Frequireno      = rsget("requireno")
				FItemList(i).Fshortageno     = rsget("shortageno")
				FItemList(i).Fpreorderno     = rsget("preorderno")
				FItemList(i).Fpreordernofix  = rsget("preordernofix")
				FItemList(i).Foffsellno      = rsget("offsellno")
				FItemList(i).Fmaxsellday     = rsget("maxsellday")
				FItemList(i).Fregdate        = rsget("regdate")
				FItemList(i).Flastupdate     = rsget("lastupdate")

				FItemList(i).FOffMwMargin    = rsget("offmwmargin")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub



	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
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
