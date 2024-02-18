<%
'####################################################
' Description :  오프라인 상품 클래스
' History : 2009.04.07 서동석 생성
'			2010.06.01 한용민 수정
'####################################################

class CIpgoWaitItem
	public Fid
	public FCode
	public Fsocid
	public Fdivcode
	public Fscheduledt
	public Fcompany_name
	public FlinkState
	public Ftotalsellcash
	public FTotalSuplyCash

	public function getStateName()
		getStateName = ""
		if IsNull(FlinkState) then Exit Function

		if FlinkState="0" then
			getStateName = "입고대기"
		elseif FlinkState="7" then
			getStateName = "입고완료"
		end if

	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COffShopOneItem
	public fitemcnt
	public fsellsum
	public fsuplyprice
	public Fitemgubun
	public Fshopitemid
	public Fitemoption
	Public Fitemno
	public Fmakerid
	public Fshopitemname
	public Fshopitemoptionname
	public fitemcopy
    public FShopItemOrgprice   '' 소비자가.
	public Fshopitemprice
	public Fshopsuplycash
	public Fisusing
	public Fregdate
	public Fipgodate
	public Fupdt
	public FOnLineItemprice     '' 온라인 판매가
    public FOnlineitemorgprice  '' 온라인 소비자가
	public fonofflinkyn
    public Foptaddprice
    public Foptaddbuyprice
	public ftermsale
	public FItemCouponYn
	public FCurrItemCouponIdx
	public FItemCouponType
	public FItemCouponValue
	public Fcouponbuyprice

	public FImageSmall
	public FImageList
	public FOffImgMain
	public FOffImgList
	public FOffImgSmall
	public FSocName
	public FSocName_Kor
	public FSocNameKor
	public Fextbarcode
	public FmakerMargin
	public FshopMargin
	public Fsellyn
	public Flimityn
	public Flimitno
	public Flimitsold
	public Flastrealdate
	public Flastrealno
	public Fipno
	public Fchulno
	public Fsellno
	public Fcurrno
	public FmwDiv
	public Foptusing
	public Fsell7days
	public Fjupsu7days
	public Foffchulgo7days
	public Foffconfirmno
	public Foffjupno
	public Frequireno
	public Fshortageno
	public FPreOrderNo
	public fpreordernofix
	public Fmaxsellday
	public FOnlineSellcash
	public FOnlineBuycash
	public FOnlineOptaddprice
    public FOnlineOptaddbuyprice
	public FOnlineOrgprice
	public FOnlineSailYn
	public FdeliveryType
	public Fvatinclude
	public Fvatinclude10
	public Fdiscountsellprice
	public FShopbuyprice
	public Fdefaultmargin
	public Fdefaultsuplymargin
	public Fdefaultmargine_fran
	public Fdefaultsuplymargine_fran
	public Fonlineeventprice
	public Fipkumdiv4
	public Fipkumdiv2
	public Fonlineitemname
	public Fonlineitemoptionname
	public Fonlinemakerid
	public Fchargeid
	public FCateCDL
	public FCateCDM
	public FCateCDS
	public FCateCDLName
	public FCateCDMName
	public FCateCDSName
	public Fchargediv
    public Fcomm_cd
    public Fforeign_sellprice
    public Fforeign_suplyprice

    public FCenterMwdiv
    public FonLineDanjongyn
    public FOffsell7days
	' 상품별주문 추가 속성
	Public Fidx
	Public FmasterIdx
	Public FjumunItemNo		' 상품별 주문수량
	Public FstateCd			' 상품별 주문상태
	Public FmasterStateCd	' 마스터 상태
	Public FconfirmDate
	Public Fcomment
	Public FsellCash
	Public FsuplyCash
	Public FbuyCash
	Public Frealstock
	Public Fipkumdiv5
	public fsell3days
	Public Frealstockno
	Public Ffirstipgodate
	Public FShopID
	public fstockregdate
	public Fstockitemid

	public flogicsipgono
	public flogicsreipgono
	public fbrandipgono
	public fbrandreipgono
	public fresellno
	public ferrsampleitemno
	public ferrbaditemno
	public ferrrealcheckno
	public fsysstockno
	public frequiredStock
	public frequire3daystock
	public frequire7daystock
	public frequire14daystock
	Public FLogicsRealStock

    public Ftnbarcode  ''2017/05/23
	public Fgeneralbarcode

    public FvolX
    public FvolY
    public FvolZ
    public FitemWeight

	''유효재고
    public function getAvailStock()
        getAvailStock = FrealstockNo + Ferrsampleitemno + Ferrbaditemno
    end function

	public function getJungsanDivName()
		if FComm_cd="B011" then
			getJungsanDivName = "텐바이텐위탁"
		elseif FComm_cd="B031" then
			getJungsanDivName = "매입출고정산"
		elseif FComm_cd="B012" then
			getJungsanDivName = "업체위탁"
		elseif FComm_cd="B022" then
			getJungsanDivName = "업체매입"
		elseif FComm_cd="B021" then
			getJungsanDivName = "오프매입"
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

    public function IsShopContractExists()
        IsShopContractExists = Not ((IsNULL(Fcomm_cd)) or (Fcomm_cd=""))
    end function

    public function IsOffSaleItem()
        IsOffSaleItem = (FShopItemOrgprice>Fshopitemprice)
    end function

	public function GetShortageNo()
		GetShortageNo = Fshortageno - Fipkumdiv4 - Fipkumdiv2
	end function

	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "위탁"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function

	public function getMwDivColor()
		if FmwDiv="M" then
			getMwDivColor = "#CC2222"
		elseif FmwDiv="W" then
			getMwDivColor = "#2222CC"
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
	end Function

	public function IsTempSoldOut()
		IsTempSoldOut = Not IsSoldOut And (Fsellyn="S")
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

	public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = FimageSmall
		else
			GetImageSmall = FOffImgSmall
		end if
	end function

    public function GetImageList()
		if Fitemgubun="10" then
			GetImageList = FimageList
		else
			GetImageList = FOffImgList
		end if
	end function

	''가맹점 공급가
	public function GetFranchiseSuplycash()
		dim ishopsupycash

		''가맹점공급가가 0 인경우 기본 마진으로 구한다
		if Fshopbuyprice<>0 then
			ishopsupycash = Fshopbuyprice
		else
		    ''마진이 설정 안되있는경우 매입마진-5%
		    if IsNULL(FshopMargin) or (FshopMargin=0) then
		        IF (FmakerMargin=0) then FmakerMargin=35 '''기본값;
		        ishopsupycash = CLng(Fshopitemprice * (100-(FmakerMargin-5))/100)
		    else
			    ishopsupycash = CLng(Fshopitemprice * (100-FshopMargin)/100)
			end if
		end if

		''공급가가 매입가보다 작은경우 공급가를 사용
		if (ishopsupycash<GetFranchiseBuycash) then ishopsupycash = GetFranchiseBuycash

		GetFranchiseSuplycash = ishopsupycash
	end function

	''직영점 공급가 : 가맹점과 동일
	public function GetOfflineSuplycash()
		GetOfflineSuplycash = GetFranchiseSuplycash
	end function

	''가맹점 공급시 매입가(업체로부터 매입하는가격)
	public function GetFranchiseBuycash()
		dim ibuycash
		''가맹점 매입가가 0 인경우 기본 마진으로 구한다
		if Fshopsuplycash<>0 then
			ibuycash = Fshopsuplycash
		else
		    IF (FmakerMargin=0) then FmakerMargin=35 '''기본값;
			ibuycash = CLng(Fshopitemprice * (100-FmakerMargin)/100)

			''온라인 매입가보다 큰경우 온라인 매입가를 사용(Fshopsuplycash 가 지정된 경우는 제외)
			''200906 FOnlineOptaddbuyprice 추가
			''온라인만 세일 하는 경우등 // 위탁->매입출고 인 경우 // 이 조건 제외 (2012-02-16)
		    ''if (FOnlinebuycash<>0) and (ibuycash>FOnlinebuycash+FOnlineOptaddbuyprice) then ibuycash=FOnlinebuycash+FOnlineOptaddbuyprice
		end if

        ''(온라인) 매입상품인경우 - 원 매입가 사용. 20121029
		if (FItemGubun="10") and (FMwdiv="M") then
            if (FOnlinebuycash<>0) then
                if (Fcomm_cd<>"B012") and (Fcomm_cd<>"B022") then       ''업체위탁은 제외
                    ibuycash = FOnlinebuycash+FOnlineOptaddbuyprice
                end if
            end if
        end if

		GetFranchiseBuycash = ibuycash
	end function

    '// 수익률 분석용 출고내역 매입가
    '// !!! 정산내역에 올라가면 안됨 !!!
	public function GetFranchiseBuycashByItemInfo()
		dim ibuycash

		''가맹점 매입가가 0 인경우 기본 마진으로 구한다
		if Fshopsuplycash<>0 then
			ibuycash = Fshopsuplycash
		else
		    IF (FmakerMargin=0) then FmakerMargin=35 '''기본값;
			ibuycash = CLng(Fshopitemprice * (100-FmakerMargin)/100)
		end if

		GetFranchiseBuycashByItemInfo = ibuycash
	end function

	''직영점 공급시 매입가(업체로부터 매입하는가격) : 가맹점과 동일
	public function GetOfflineBuycash()
		GetOfflineBuycash = GetFranchiseBuycash
	end function

	''업체위탁 공급가격
	public function GetChargeMaySuplycash()
		GetChargeMaySuplycash = GetFranchiseBuycash
	end function

	public function GetBarCode()
		GetBarCode = Fitemgubun & Format00(6,Fshopitemid) & Fitemoption
		if (Fshopitemid >= 1000000) then
			getBarCode = CStr(Fitemgubun) + Format00(8,Fshopitemid) + Fitemoption
		end if
	end function

	public function GetBarCodeBoldStr()
		GetBarCodeBoldStr = Fitemgubun & "-" & Format00(6,Fshopitemid) & "-" & Fitemoption
		if (Fshopitemid >= 1000000) then
			GetBarCodeBoldStr = CStr(Fitemgubun) & "-" & Format00(8,Fshopitemid) & "-" & Fitemoption
		end if
	end function

	Private Sub Class_Initialize()
		FmakerMargin=0

		FOnlineOptaddbuyprice = 0
		FOnlineOptaddprice    = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CPrtOneItem
	public Fitemid
	public Fitemname
	public Fmakerid
	public Fitemprice
	public Fisusing
	public FImageSmall
	public FImageList

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COffShopItem
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
	public FRectPrdCode						'물류코드(102222220000)
	public FRectGeneralBarCode				'범용바코드
	public FRectOnlyOffLine
	public FRectOnlyNotIpGo
	public FRectOnlyUsing
	public FRectIpChulId
	public FRectIpGoOnly
    public FRectLogicsIpGoOnly
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
	public FRectBarCodeArr
	public FRectItemName
	public FRectShopItemName
	public FRectImageView
	public FRectAdminView
	public FRectUpchebeasongInclude
    public FRectCDL
    public FRectCDM
    public FRectCDS
    public FRectOnlineExpiredItem
    public FRectOnlyActive
    public FRectPriceRow
    public FRectSellYN
    public frectumwdiv
	public frectsaleflg
	public FRectDateSearch
	public FRectSDate
	public FRectEDate
	public FRectIsusing
	public FRectSorting
	public frectpublicbarcode
	public FRectcomm_cd
	public FPageCount
	public FRectStartDay
	public FRectEndDay
	public frectdatefg
	public FRectInc3pl
    public FRectMakerid
    public FRectCenterMwdiv
    public frectgetrows
    public Fcurrencyunit
    public Floginsite
    public fcountrylangcd
    public fcurrencyChar
	public FRectSell7days
	public FRectIncludePreOrder
	public FRectShortageType
	Public FRectOnlineMWdiv
	Public FRectCopyIdx
	Public flinkPriceType
	public fmultiplerate
    public FRectContractYN

    public FRectSizeYn
    public FRectIsWeight

	public function GetImageFolerName(byval i)
		GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public function GetOnOffDiffItemPriceList()
		dim sqlStr,i

		sqlStr = " select distinct top " + CStr(FPageSize) + " s.shopitemid , s.shopitemname,"
		sqlStr = sqlStr + " s.makerid, s.shopitemprice, IsNull(i.sellcash,0) as sellcash, i.sailyn, i.orgprice"
		''옵션 추가금액
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i  "
		sqlStr = sqlStr + " on s.itemgubun='10' and s.shopitemid=i.itemid"
        ''옵션 추가금액
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on "
		sqlStr = sqlStr + "		s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " where s.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and s.makerid<>i.makerid"
		else
			sqlStr = sqlStr + " where s.shopitemprice<>i.sellcash"
		end if
		sqlStr = sqlStr + " order by s.shopitemid desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopOneItem
				FItemList(i).Fshopitemid     = rsget("shopitemid")
				FItemList(i).FShopItemName     = db2html(rsget("shopitemname"))
				FItemList(i).Fmakerid    = rsget("makerid")
				FItemList(i).Fshopitemprice  = rsget("shopitemprice")
				FItemList(i).Fonlinesellcash  = rsget("sellcash")
				FItemList(i).Fonlinesailyn  = rsget("sailyn")
				FItemList(i).Fonlineorgprice  = rsget("orgprice")
                ''옵션 추가금액
    			FOneItem.FOnlineOptaddprice = rsget("optaddprice")
    			FOneItem.FOnlineOptaddbuyprice = rsget("optaddbuyprice")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function

	public function GetOnOffDiffItemBrandList()
		dim sqlStr,i

		sqlStr = " select distinct top " + CStr(FPageSize) + " s.shopitemid , s.shopitemname,"
		sqlStr = sqlStr + " s.makerid, IsNULL(i.makerid,'') as onlinemakerid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i  "
		sqlStr = sqlStr + " on s.itemgubun='10' and s.shopitemid=i.itemid"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " where s.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and s.makerid<>i.makerid"
		else
			sqlStr = sqlStr + " where s.makerid<>i.makerid"
		end if
		sqlStr = sqlStr + " order by s.shopitemid desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopOneItem
				FItemList(i).Fshopitemid     = rsget("shopitemid")
				FItemList(i).FShopItemName     = db2html(rsget("shopitemname"))
				FItemList(i).Fmakerid    = rsget("makerid")
				FItemList(i).Fonlinemakerid  = rsget("onlinemakerid")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function

	public function GetOnOffDiffItemNameList()
		dim sqlStr,i

		sqlStr = " select distinct top " + CStr(FPageSize) + " s.shopitemid , s.shopitemname,"
		sqlStr = sqlStr + " IsNULL(i.itemname,'') as itemname, s.makerid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i  "
		sqlStr = sqlStr + " on s.itemgubun='10' and s.shopitemid=i.itemid"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " where s.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and s.itemname<>i.itemname"
		else
			sqlStr = sqlStr + " where s.shopitemname<>i.itemname"
		end if
		sqlStr = sqlStr + " order by s.shopitemid desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopOneItem
				FItemList(i).Fshopitemid     = rsget("shopitemid")
				FItemList(i).Fonlineitemname     = db2html(rsget("itemname"))
				FItemList(i).FShopItemName     = db2html(rsget("shopitemname"))
				FItemList(i).Fmakerid    = rsget("makerid")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function

	public function GetOffOneItem()
		dim sqlStr,i

		sqlStr = " select top 1 s.*"
		sqlStr = sqlStr + " ,IsNULL(i.sellcash,0) as sellcash, IsNULL(i.buycash,0) as buycash"
		sqlStr = sqlStr + " ,IsNULL(i.orgprice,0) as orgprice, i.sailyn, i.smallimage, i.listimage"
		sqlStr = sqlStr + " ,c.nmlarge, c.nmmid, c.nmsmall, IsNULL(i.mwdiv,'') as mwdiv, IsNull(i.vatinclude,'Y') as vatinclude10 "
		''옵션 추가금액
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"

		if (FRectShopid<>"") then
		    sqlStr = sqlStr + " ,d.chargediv, d.comm_cd, IsNULL(d.defaultmargin,0) as defaultmargin ,IsNULL(d.defaultsuplymargin,0) as defaultsuplymargin"
		end if

		sqlStr = sqlStr + "  ,isNull(a.itemid,0) as stockitemid "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on "
		sqlStr = sqlStr + "		s.itemgubun='10' and s.shopitemid=i.itemid"
		''옵션 추가금액
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on "
		sqlStr = sqlStr + "		s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].vw_category c on"
		sqlStr = sqlStr + "		s.catecdl=c.cdlarge and s.catecdm=c.cdmid and s.catecdn=c.cdsmall"

		if (FRectShopid<>"") then
		    sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d "
		    sqlStr = sqlStr + " on d.shopid='" + FRectShopid + "' and s.makerid=d.makerid"
		end if

		sqlStr = sqlStr + " left outer join db_summary.dbo.tbl_current_logisstock_summary as a "
		sqlStr = sqlStr + "  on s.shopitemid = a.itemid and s.itemoption = a.itemoption and s.itemgubun= a.itemgubun "
		sqlStr = sqlStr + " where s.itemgubun='" + FRectItemgubun + "'"
		sqlStr = sqlStr + " and s.shopitemid=" + CStr(FRectItemId) + ""
		sqlStr = sqlStr + " and s.itemoption='" + FRectItemOption + "'"

		if (FRectMakerid<>"") then
		    sqlStr = sqlStr + " and s.makerid='" + FRectMakerid + "'"
		end if

		rsget.Open sqlStr,dbget,1
        FResultCount = rsget.RecordCount

		if not rsget.Eof then
			set FOneItem = new COffShopOneItem
			FOneItem.Fitemgubun         = rsget("itemgubun")
			FOneItem.Fshopitemid        = rsget("shopitemid")
			FOneItem.Fitemoption     	= rsget("itemoption")
			FOneItem.Fmakerid           = rsget("makerid")
			FOneItem.Fshopitemname      = db2html(rsget("shopitemname"))
			FOneItem.Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
			FOneItem.fitemcopy= db2html(rsget("itemcopy"))
            FOneItem.FShopItemOrgprice  = rsget("orgsellprice")
			FOneItem.Fshopitemprice     = rsget("shopitemprice")
			FOneItem.Fshopsuplycash     = rsget("shopsuplycash")
			FOneItem.Fisusing           = rsget("isusing")
			FOneItem.Fregdate           = rsget("regdate")
			FOneItem.Fupdt           	= rsget("updt")
			FOneItem.Fextbarcode 		= rsget("extbarcode")
			FOneItem.Fvatinclude		= rsget("vatinclude")
			FOneItem.Fvatinclude10		= rsget("vatinclude10")
			FOneItem.FOnlineSellcash	= rsget("sellcash")
			FOneItem.FOnlineBuycash		= rsget("buycash")
			''옵션 추가금액
			FOneItem.FOnlineOptaddprice = rsget("optaddprice")
			FOneItem.FOnlineOptaddbuyprice = rsget("optaddbuyprice")
			FOneItem.FOnlineOrgprice	= rsget("orgprice")
			FOneItem.FOnlineSailYn		= rsget("sailyn")
			FOneItem.Fdiscountsellprice = rsget("discountsellprice")
			FOneItem.FOffimgMain	= rsget("offimgmain")
			FOneItem.FOffimgList	= rsget("offimglist")
			FOneItem.FOffimgSmall	= rsget("offimgsmall")
			FOneItem.FShopbuyprice	= rsget("shopbuyprice")

			if FOneItem.FOffimgMain<>"" then FOneItem.FOffimgMain = webImgUrl + "/offimage/offmain/i" + FOneItem.Fitemgubun + "/" + GetImageSubFolderByItemid(FOneItem.Fshopitemid) + "/" + FOneItem.FOffimgMain
			if FOneItem.FOffimgList<>"" then FOneItem.FOffimgList = webImgUrl + "/offimage/offlist/i" + FOneItem.Fitemgubun + "/" + GetImageSubFolderByItemid(FOneItem.Fshopitemid) + "/" + FOneItem.FOffimgList
			if FOneItem.FOffimgSmall<>"" then FOneItem.FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FOneItem.Fitemgubun + "/" + GetImageSubFolderByItemid(FOneItem.Fshopitemid) + "/" + FOneItem.FOffimgSmall

			FOneItem.FimageSmall     = rsget("smallimage")
			if FOneItem.FimageSmall<>"" then
				FOneItem.FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FOneItem.Fshopitemid) + "/" + FOneItem.FimageSmall
			end if

			FOneItem.FimageList     = rsget("listimage")
			if FOneItem.FimageList<>"" then
				FOneItem.FimageList     = webImgUrl + "/image/list/" + GetImageSubFolderByItemid(FOneItem.Fshopitemid) + "/" + FOneItem.FimageList
			end if

			FOneItem.FCateCDL		= rsget("catecdl")
			FOneItem.FCateCDM		= rsget("catecdm")
			FOneItem.FCateCDS		= rsget("catecdn")
			FOneItem.FCateCDLName	= db2html(rsget("nmlarge"))
			FOneItem.FCateCDMName	= db2html(rsget("nmmid"))
			FOneItem.FCateCDSName	= db2html(rsget("nmsmall"))
			FOneItem.FCenterMwdiv   = rsget("centermwdiv")
			FOneItem.FmwDiv	        = rsget("mwdiv")    ''온라인 매입구분(tbl_item)

			if (FRectShopid<>"") then
			    FOneItem.Fcomm_cd       = rsget("comm_cd")
    			FOneItem.FMakerMargin    = rsget("defaultmargin")
    			FOneItem.FShopMargin    = rsget("defaultsuplymargin")
    		end if

    		FOneItem.Fstockitemid	= rsget("stockitemid")
		end if

		rsget.Close
	end function

	public function GetItemPrintList()
		dim sqlStr,i

		sqlStr = " select i.itemid,i.makerid,"
		sqlStr = sqlStr + " i.itemname,i.sellcash,"
		sqlStr = sqlStr + " IsNull(m.imgsmall,'') as smallimage,IsNull(m.imglist,'') as listimage "
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		if FRectOnly10beasong="on" then
			sqlStr = sqlStr + " and i.deliverytype in ('1','3','4')"
		end if

		sqlStr = sqlStr + " order by i.itemid desc"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  rsget.RecordCount
		FResultCount =  rsget.RecordCount

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CPrtOneItem
					FItemList(i).Fitemid     = rsget("itemid")
					FItemList(i).Fitemname     = db2html(rsget("itemname"))
					FItemList(i).Fmakerid    = rsget("makerid")
					FItemList(i).Fitemprice  = rsget("sellcash")
					FItemList(i).FImageSmall = rsget("smallimage")
					FItemList(i).FImageList  = rsget("listimage")

					if FItemList(i).FImageSmall<>"" then
						FItemList(i).FImageSmall = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageSmall
					end if
					if FItemList(i).FImageList<>"" then
						FItemList(i).FImageList = webImgUrl + "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageList
					end if

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end function

'	'//admin/offshop/shopprint.asp		'//admin/offshop/shopprint_pop.asp		'//offshop/items/shopprint.asp		'/사용안함
'	public function GetShopPrintList()
'		dim sqlStr,i
'
'		if FRectOnOffGubun="on" then
'			sqlStr = " select top " + CStr(FPageSize) + " '10' as itemgubun,i.itemid as shopitemid, ISNULL(v.itemoption,'0000') as itemoption,i.makerid,"
'			sqlStr = sqlStr + " i.itemname as shopitemname, v.optionname as shopitemoptionname, i.sellcash as shopitemprice,i.isusing,"
'			sqlStr = sqlStr + " IsNull(i.smallimage,'') as imgsmall, "
'			sqlStr = sqlStr + " IsNull(i.listimage,'') as imglist "
'			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
'			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v"
'			sqlStr = sqlStr + " on i.itemid=v.itemid"
'			sqlStr = sqlStr + " where i.itemid<>0"
'
'			if FRectOnlyTenbeasong="on" then
'				sqlStr = sqlStr + " and i.deliverytype in ('1','3','4')"
'			end if
'			if FRectDesigner<>"" then
'				sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
'			end if
'			if FRectArraryBarCode<>"" then
'				sqlStr = sqlStr + " and '10' + Right('000' + convert(varchar(6),i.itemid),6) + ISNULL(v.itemoption,'0000') in (" + FRectArraryBarCode + ")"
'			end if
'
'			sqlStr = sqlStr + " order by i.itemid desc"
'		else
'			sqlStr = " select top " + CStr(FPageSize) + " s.itemgubun,s.shopitemid,s.itemoption,s.makerid,"
'			sqlStr = sqlStr + " s.shopitemname,s.shopitemoptionname,s.shopitemprice,s.isusing,"
'			sqlStr = sqlStr + " IsNull(m.smallimage,'') as imgsmall, "
'			sqlStr = sqlStr + " IsNull(m.listimage,'') as imglist "
'			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
'			sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item m"
'			sqlStr = sqlStr + "     on s.itemgubun='10' and s.shopitemid=m.itemid"
'			sqlStr = sqlStr + " where s.makerid<>''"
'
'			if FRectDesigner<>"" then
'				sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
'			end if
'
'			if FRectArraryBarCode<>"" then
'				sqlStr = sqlStr + " and s.itemgubun + Right('000' + convert(varchar(6),s.shopitemid),6) + s.itemoption in (" + FRectArraryBarCode + ")"
'			end if
'
'			sqlStr = sqlStr + " order by s.itemgubun desc, s.shopitemid desc"
'		end if
'		rsget.Open sqlStr,dbget,1
'
'		'response.write sqlStr
'		FtotalPage =  rsget.RecordCount
'		FResultCount =  rsget.RecordCount
'
'			redim preserve FItemList(FResultCount)
'			i=0
'			if  not rsget.EOF  then
'				rsget.absolutepage = FCurrPage
'				do until rsget.eof
'					set FItemList(i) = new COffShopOneItem
'					FItemList(i).Fitemgubun         = rsget("itemgubun")
'					FItemList(i).Fshopitemid        = rsget("shopitemid")
'					FItemList(i).Fitemoption     	= rsget("itemoption")
'					FItemList(i).Fmakerid           = rsget("makerid")
'					FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
'					FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
'					FItemList(i).Fshopitemprice     = rsget("shopitemprice")
'					FItemList(i).Fisusing			= rsget("isusing")
'					FItemList(i).FImageSmall   		= rsget("imgsmall")
'
'					if FItemList(i).FImageSmall<>"" then
'						FItemList(i).FImageSmall = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FImageSmall
'					end if
'
'					FItemList(i).FImageList   		= rsget("imglist")
'
'					if FItemList(i).FImageList<>"" then
'						FItemList(i).FImageList = webImgUrl + "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FImageList
'					end if
'
'					i=i+1
'					rsget.moveNext
'				loop
'			end if
'		rsget.Close
'	end function

	public function GetNotIpChulList()
		dim sqlStr, i , defaultmargin, defaultsuplymargin

		'' ===== 오프 기본마진 getDefaultmargin : 추가 ==========
		sqlStr = " select top 1 * "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " where makerid='" + FRectDesigner + "'"
		if (FRectShopid<>"") then
			sqlStr = sqlStr + " and shopid='" + FRectShopid + "'"
		end if
		sqlStr = sqlStr + " order by shopid"

		rsget.Open sqlStr,dbget,1
		if  not rsget.EOF  then
			defaultmargin = rsget("defaultmargin")
			defaultsuplymargin = rsget("defaultsuplymargin")
		end if
		rsget.Close
		'' =====================================================

		sqlStr = " select count(s.shopitemid) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_ipchul_detail d"
		sqlStr = sqlStr + " on d.masteridx=" + CStr(FRectIpChulId)
		sqlStr = sqlStr + " and d.deleteyn='N'"
		sqlStr = sqlStr + " and s.itemgubun=d.itemgubun"
		sqlStr = sqlStr + " and s.shopitemid=d.shopitemid"
		sqlStr = sqlStr + " and s.itemoption=d.itemoption"
		sqlStr = sqlStr + " where s.makerid='" + FRectDesigner + "'"
		if FRectItemName<>"" then
			sqlStr = sqlStr + " and shopitemname like '%" + FRectItemName + "%'"
		end if
		sqlStr = sqlStr + " and d.idx is Null"
		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " s.*, d.idx, i.smallimage "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_ipchul_detail d"
		sqlStr = sqlStr + " on d.masteridx=" + CStr(FRectIpChulId)
		sqlStr = sqlStr + " and d.deleteyn='N'"
		sqlStr = sqlStr + " and s.itemgubun=d.itemgubun"
		sqlStr = sqlStr + " and s.shopitemid=d.shopitemid"
		sqlStr = sqlStr + " and s.itemoption=d.itemoption"
		sqlStr = sqlStr + " where s.makerid='" + FRectDesigner + "'"

		if FRectItemName<>"" then
			sqlStr = sqlStr + " and shopitemname like '%" + FRectItemName + "%'"
		end if

		sqlStr = sqlStr + " and d.idx is Null"
		sqlStr = sqlStr + " order by s.shopitemid desc"

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
					set FItemList(i) = new COffShopOneItem
					FItemList(i).Fitemgubun         = rsget("itemgubun")
					FItemList(i).Fshopitemid        = rsget("shopitemid")
					FItemList(i).Fitemoption     	= rsget("itemoption")
					FItemList(i).Fmakerid           = rsget("makerid")
					FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
					FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
					FItemList(i).Fshopitemprice     = rsget("shopitemprice")
					FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")

					FItemList(i).FimageSmall     = rsget("smallimage")
					if FItemList(i).FimageSmall<>"" then
						FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
					end if

					FItemList(i).FOffimgSmall	= rsget("offimgsmall")
					if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
					if FItemList(i).Fitemgubun<>"10" then FItemList(i).FimageSmall = FItemList(i).FOffimgSmall
					FItemList(i).FMakerMargin = defaultmargin
					FItemList(i).FShopMargin = defaultsuplymargin

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end function

	public function GetTen2ShopList()
		dim sqlStr,i

		sqlStr = " select top 300 m.id, m.code, m.socid, m.divcode, m.scheduledt,"
		sqlStr = sqlStr + " m.totalsellcash, m.totalsuplycash, p.company_name, s.statecd"
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p, [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_ipchul_master s on m.id=s.linkidx"
		sqlStr = sqlStr + " where p.id=m.chargeid"
		sqlStr = sqlStr + " and Left(m.socid,10)='streetshop'"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and m.executedt>'2004-05-01'"

		if FRectOnlyNotIpGo="on"  then
			sqlStr = sqlStr + " and s.linkidx is NUll"
		end if
		sqlStr = sqlStr + " order by m.id desc"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CIpgoWaitItem
					FItemList(i).Fid           = rsget("id")
					FItemList(i).FCode         = rsget("code")
					FItemList(i).Fsocid        = rsget("socid")
					FItemList(i).Fdivcode      = rsget("divcode")
					FItemList(i).Fscheduledt   = rsget("scheduledt")
					FItemList(i).Ftotalsellcash = rsget("totalsellcash")
					FItemList(i).Ftotalsuplycash = rsget("totalsuplycash")
					FItemList(i).Fcompany_name = rsget("company_name")
					FItemList(i).FlinkState      = rsget("statecd")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end function

	public sub GetLinkNotRegList2()
		dim sqlStr,i

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + ""
		sqlStr = sqlStr + " i.itemid, IsNULL(o.itemoption,'0000') as itemoption, i.itemname,s.shopitemid, i.makerid,"
		sqlStr = sqlStr + " i.sellcash, IsNull(o.optionname,'') as opt2name"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlStr + " on i.itemid=o.itemid"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " on s.itemgubun='10' and i.itemid=s.shopitemid and s.itemoption= IsNULL(o.itemoption,'0000')"
		sqlStr = sqlStr + " ,  ("
		sqlStr = sqlStr + " select distinct makerid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " where s.shopitemid is NULL"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.itemid<>0"
		sqlStr = sqlStr + " and i.makerid=T.makerid"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if

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
					set FItemList(i) = new COffShopOneItem
					FItemList(i).Fitemgubun         = "10"
					FItemList(i).Fshopitemid        = rsget("itemid")
					FItemList(i).Fitemoption     	= rsget("itemoption")
					FItemList(i).Fmakerid           = rsget("makerid")
					FItemList(i).Fshopitemname      = db2html(rsget("itemname"))
					FItemList(i).Fshopitemoptionname= db2html(rsget("opt2name"))
					FItemList(i).Fshopitemprice     = rsget("sellcash")
					FItemList(i).Fshopsuplycash     = 0
					FItemList(i).Fdiscountsellprice = 0
					FItemList(i).FShopbuyprice		= 0

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	'//admin/offshop/shopitemreg.asp '//common/offshop/pop_itemaddinfo_onofflink_off.asp
	public sub GetLinkNotRegList3()
		dim sqlStr,i

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + ""
		sqlStr = sqlStr + " i.itemid, IsNULL(o.itemoption,'0000') as itemoption, i.itemname,s.shopitemid, i.makerid,"
		sqlStr = sqlStr + " (i.sellcash + IsNULL(o.optaddprice,0)) as sellcash, IsNULL(i.mwdiv,'') as mwdiv, IsNull(o.optionname,'') as opt2name,"
        sqlStr = sqlStr + " (i.orgprice + IsNULL(o.optaddprice,0)) as onlineitemorgprice"
		'''sqlStr = sqlStr + " ,IsNULL(d1.defaultmargin,0) as defaultmargin, IsNULL(d1.defaultsuplymargin,0) as defaultsuplymargin"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlStr + " 	on i.itemid=o.itemid and o.isusing='Y'"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " 	on s.itemgubun='10' and i.itemid=s.shopitemid and s.itemoption= IsNULL(o.itemoption,'0000')"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d1 "
		sqlStr = sqlStr + " 	on i.makerid=d1.makerid and d1.shopid='streetshop000'"
		sqlStr = sqlStr + " where s.shopitemid is NULL"

		if FRectUpchebeasongInclude="on" then
			if frectumwdiv = "Y" then
				sqlStr = sqlStr + " and i.mwdiv='U'"
			elseif frectumwdiv = "N" then
				sqlStr = sqlStr + " and i.sellyn='N'"
				sqlStr = sqlStr + " and i.mwdiv<>'U'"
			end if
		else
			sqlStr = sqlStr + " and i.mwdiv<>'U'"
			sqlStr = sqlStr + " and i.sellyn<>'N'"
			sqlStr = sqlStr + " and i.isusing='Y'"
		end if

		sqlStr = sqlStr + " and i.itemid<>0"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if

		sqlStr = sqlStr + " and d1.defaultmargin is not null"       '' streetshop000 마진설정된 브랜드만
		sqlStr = sqlStr + " order by i.makerid, i.itemid desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		FTotalCount = rsget.RecordCount

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopOneItem
					FItemList(i).Fitemgubun         = "10"
					FItemList(i).Fshopitemid        = rsget("itemid")
					FItemList(i).Fitemoption     	= rsget("itemoption")
					FItemList(i).Fmakerid           = rsget("makerid")
					FItemList(i).Fshopitemname      = db2html(rsget("itemname"))
					FItemList(i).Fshopitemoptionname= db2html(rsget("opt2name"))
					FItemList(i).Fshopitemprice     = rsget("sellcash")
					FItemList(i).FmwDiv	= rsget("mwdiv")    ''온라인 매입구분
					FItemList(i).FOnlineitemorgprice= rsget("onlineitemorgprice")
					FItemList(i).Fshopsuplycash     = CLng(FItemList(i).Fshopitemprice * (100-FItemList(i).Fdefaultmargin)/100)

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	'// 핑거스→OFF 상품등록
	'// /admin/offshop/shopitemlink_ACA.asp
	public sub GetAcaLinkReqList()
		dim sqlStr, i
		Dim startNum, endNum

		''FRectitemid
		''FRectDesigner
		If (FRectitemid = "") Then
			FRectitemid = "NULL"
		End If

		sqlStr = " exec [db_shop].[dbo].[sp_Ten_ACA_Get_AcademyDiyItem_CNT] " & FRectitemid & ", '" & FRectDesigner & "', '" & FRectOnlineMWdiv & "', '" & FRectSellYN & "', '" & FRectIsusing & "' "
		''rw sqlStr
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		rsget.Close


		startNum = ((FCurrPage - 1) * FPageSize) + 1
		endNum = startNum + FPageSize

		sqlStr = " exec [db_shop].[dbo].[sp_Ten_ACA_Get_AcademyDiyItem_LIST] " & startNum & ", " & endNum & ", " & FRectitemid & ", '" & FRectDesigner & "', '" & FRectOnlineMWdiv & "', '" & FRectSellYN & "', '" & FRectIsusing & "' "
		''rw sqlStr
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COffShopOneItem

				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("itemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("itemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("sellcash")
				FItemList(i).FOnlineitemorgprice= rsget("orgprice")
				FItemList(i).FmwDiv				= rsget("mwdiv")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

    '/admin/offshop/shopitemlink.asp ''2016/04/17
    public sub GetLinkReqList()
        dim sqlStr,i ,sqlSearch
        sqlSearch = ""

		if FRectPrdCode<>"" then
			if right(FRectPrdCode,4)="0000" then
				if (Len(FRectPrdCode) = 12) then
					'sqlsearch = sqlsearch & " 	and o.itemgubun = '" & LEFT(CStr(FRectPrdCode), 2) & "' " & VbCrLf
					sqlsearch = sqlsearch & " 	and i.itemid = " & RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) & " " & VbCrLf
					'sqlsearch = sqlsearch & " 	and IsNULL(o.itemoption,'0000') = '" & RIGHT(CStr(FRectPrdCode), 4) & "' " & VbCrLf
				else
					'sqlsearch = sqlsearch & " 	and o.itemgubun = '" & LEFT(CStr(FRectPrdCode), 2) & "' " & VbCrLf
					sqlsearch = sqlsearch & " 	and i.itemid = " & RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) & " " & VbCrLf
					'sqlsearch = sqlsearch & " 	and IsNULL(o.itemoption,'0000') = '" & RIGHT(CStr(FRectPrdCode), 4) & "' " & VbCrLf
				end if
			else
				if (Len(FRectPrdCode) = 12) then
					'sqlsearch = sqlsearch & " 	and o.itemgubun = '" & LEFT(CStr(FRectPrdCode), 2) & "' " & VbCrLf
					sqlsearch = sqlsearch & " 	and o.itemid = " & RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) & " " & VbCrLf
					sqlsearch = sqlsearch & " 	and IsNULL(o.itemoption,'0000') = '" & RIGHT(CStr(FRectPrdCode), 4) & "' " & VbCrLf
				else
					'sqlsearch = sqlsearch & " 	and o.itemgubun = '" & LEFT(CStr(FRectPrdCode), 2) & "' " & VbCrLf
					sqlsearch = sqlsearch & " 	and o.itemid = " & RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) & " " & VbCrLf
					sqlsearch = sqlsearch & " 	and IsNULL(o.itemoption,'0000') = '" & RIGHT(CStr(FRectPrdCode), 4) & "' " & VbCrLf
				end if
			end if
		end if
		if frectitemoption <> "" then
			sqlsearch = sqlsearch & " and IsNULL(o.itemoption,'0000')='"& frectitemoption &"'"
		end if
        if (FRectIsusing<>"") then
            sqlsearch = sqlsearch + " and i.isusing='"&FRectIsusing&"'"
        end if
        if (FRectSellYN<>"") then
            if (FRectSellYN="YS") then
                sqlsearch = sqlsearch + " and i.sellyn in ('Y','S')"
            else
                sqlsearch = sqlsearch + " and i.sellyn='"&FRectSellYN&"'"
            end if
        end if
        if (FRectDesigner<>"") then
		    sqlSearch = sqlSearch + " and i.makerid='" + FRectDesigner + "'"
		end if
		if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if
		if (FRectOnlineMWdiv<>"") then
		    if (FRectOnlineMWdiv="MW") then
		        sqlSearch = sqlSearch + " and i.mwdiv<>'U'"
		    else
		        sqlSearch = sqlSearch + " and i.mwdiv='" + FRectOnlineMWdiv + "'"
		    end if
		end if
		if FRectitemname <> "" then
			sqlSearch = sqlSearch + " and i.itemname like'%" + FRectitemname + "%'"
		end if
        if FRectcdl<>"" then
            sqlsearch = sqlsearch + " and i.cate_large='"&FRectcdl&"'"
        end if
        if FRectcdm<>"" then
            sqlsearch = sqlsearch + " and i.cate_mid='"&FRectcdm&"'"
        end if
        if FRectcds<>"" then
            sqlsearch = sqlsearch + " and i.cate_small='"&FRectcds&"'"
        end if

        if FRectContractYN = "Y" then
            sqlsearch = sqlsearch + " and d1.defaultmargin is not null"       '' streetshop000 마진설정된 브랜드만
        end if

		'/itemdiv : 구매제한상품:07 , 티켓상품:08, Present상품:09 , 여행상품:18, 딜상품:21, 마일리지샵:82
		sqlStr = " select count(*) as cnt"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o with (nolock)"
		sqlStr = sqlStr + " 	on i.itemid=o.itemid "		'--and o.isusing='Y'
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " 	on s.itemgubun='10' and i.itemid=s.shopitemid and s.itemoption= IsNULL(o.itemoption,'0000')"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d1 with (nolock)"
		sqlStr = sqlStr + " 	on i.makerid=d1.makerid and d1.shopid='streetshop000'"
		sqlStr = sqlStr + " where s.shopitemid is NULL"
        sqlStr = sqlStr + " and i.itemdiv not in ('08','09','18','21')"		'07,'82'
        ''sqlStr = sqlStr + " and i.makerid not in ('10x10between')"
		sqlStr = sqlStr + " " & sqlsearch  ''and i.itemid<>0
		''sqlStr = sqlStr + " and d1.defaultmargin is not null"       '' streetshop000 마진설정된 브랜드만

		'response.write sqlStr & "<br>"
		'response.end
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + ""
		sqlStr = sqlStr + " i.itemid, IsNULL(o.itemoption,'0000') as itemoption, i.itemname,s.shopitemid, i.makerid,"
		sqlStr = sqlStr + " (i.sellcash + IsNULL(o.optaddprice,0)) as sellcash, IsNULL(i.mwdiv,'') as mwdiv, IsNull(o.optionname,'') as opt2name,"
        sqlStr = sqlStr + " (i.orgprice + IsNULL(o.optaddprice,0)) as onlineitemorgprice"
        sqlStr = sqlStr + " ,'X' as termsale"
''        sqlStr = sqlStr + " ,(case when t.itemid is not null then 'Y' else 'N' end) as termsale"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o with (nolock)"
		sqlStr = sqlStr + " 	on i.itemid=o.itemid "		'--and o.isusing='Y'
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " 	on s.itemgubun='10' and i.itemid=s.shopitemid and s.itemoption= IsNULL(o.itemoption,'0000')"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d1 with (nolock)"
		sqlStr = sqlStr + " 	on i.makerid=d1.makerid and d1.shopid='streetshop000'"

'' moon 요청 소비자가 default
''		sqlStr = sqlStr + " left join ("
''		sqlStr = sqlStr + " 	select si.itemid"                           ''distinct => group by
''		sqlStr = sqlStr + " 	from db_event.dbo.tbl_saleItem si"
''		sqlStr = sqlStr + " 	join db_event.dbo.tbl_sale s"
''		sqlStr = sqlStr + " 		on si.sale_code = s.sale_code"
''		sqlStr = sqlStr + " 	where "
''		''sqlStr = sqlStr + " 	getdate() between s.sale_startdate and s.sale_enddate"
''		sqlStr = sqlStr + " 	s.sale_startdate<'"&LEFT(now(),10)&"' and s.sale_enddate>='"&LEFT(now(),10)&"'"
''		sqlStr = sqlStr + " 	and s.sale_status=6"
''		sqlStr = sqlStr + " 	group by si.itemid"
''		sqlStr = sqlStr + " ) as t"
''		sqlStr = sqlStr + " 	on i.itemid = t.itemid"

		sqlStr = sqlStr + " where s.shopitemid is NULL"
        sqlStr = sqlStr + " and i.itemdiv not in ('08','09','18','21')"		'07,'82'
        ''sqlStr = sqlStr + " and i.makerid not in ('10x10between')"
		sqlStr = sqlStr + " and i.itemid<>0 " & sqlsearch
		''sqlStr = sqlStr + " and d1.defaultmargin is not null"       '' streetshop000 마진설정된 브랜드만
		sqlStr = sqlStr + " order by i.makerid, i.itemid desc"

		''rw sqlStr & "<br>"
		'response.end
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		''FTotalCount = rsget.RecordCount

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopOneItem

					FItemList(i).ftermsale        = rsget("termsale")
					FItemList(i).Fitemgubun         = "10"
					FItemList(i).Fshopitemid        = rsget("itemid")
					FItemList(i).Fitemoption     	= rsget("itemoption")
					FItemList(i).Fmakerid           = rsget("makerid")
					FItemList(i).Fshopitemname      = db2html(rsget("itemname"))
					FItemList(i).Fshopitemoptionname= db2html(rsget("opt2name"))
					FItemList(i).Fshopitemprice     = rsget("sellcash")
					FItemList(i).FmwDiv	= rsget("mwdiv")    ''온라인 매입구분
					FItemList(i).FOnlineitemorgprice= rsget("onlineitemorgprice")
					FItemList(i).Fshopsuplycash     = CLng(FItemList(i).Fshopitemprice * (100-FItemList(i).Fdefaultmargin)/100)

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

    end Sub

	'/admin/offshop/shopitemlink.asp
	public sub GetLinkNotRegsalelist()
		dim sqlStr,i ,sqlsearch

        ''' isusing은 Default Y
		if (FRectDesigner<>"") then
		    sqlsearch = sqlsearch + " and i.makerid='" + FRectDesigner + "'"

			if frectumwdiv = "Y" then
				sqlsearch = sqlsearch + " and i.mwdiv='U'"
				sqlsearch = sqlsearch + " and i.sellyn='Y'"
				sqlsearch = sqlsearch + " and i.isusing='Y'"
			elseif frectumwdiv = "N" then '''판매중지상품검색
				sqlsearch = sqlsearch + " and i.sellyn<>'Y'"
				sqlsearch = sqlsearch + " and i.isusing='Y'"
				''sqlsearch = sqlsearch + " and i.mwdiv<>'U'"
			else  ''ALL
			    sqlsearch = sqlsearch + " and i.sellyn='Y'"
				sqlsearch = sqlsearch + " and i.isusing='Y'"
			end if
		else
			sqlsearch = sqlsearch + " and i.mwdiv<>'U'"
			sqlsearch = sqlsearch + " and i.sellyn<>'N'"
			sqlsearch = sqlsearch + " and i.isusing='Y'"
		end if

		if FRectitemid<>"" then
			sqlsearch = sqlsearch + " and i.itemid=" + FRectitemid + ""
		end if

		if FRectitemname<>"" then
			sqlsearch = sqlsearch + " and i.itemname like '%" + FRectitemname + "%'"
		end If


		sqlStr = " select count(*) as cnt"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlStr + " 	on i.itemid=o.itemid "		'--and o.isusing='Y'
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " 	on s.itemgubun='10' and i.itemid=s.shopitemid and s.itemoption= IsNULL(o.itemoption,'0000')"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d1 "
		sqlStr = sqlStr + " 	on i.makerid=d1.makerid and d1.shopid='streetshop000'"
''불필요.
'		sqlStr = sqlStr + " left join ("
'		sqlStr = sqlStr + " 	select distinct si.itemid"
'		sqlStr = sqlStr + " 	from db_event.dbo.tbl_saleItem si"
'		sqlStr = sqlStr + " 	join db_event.dbo.tbl_sale s"
'		sqlStr = sqlStr + " 		on si.sale_code = s.sale_code"
'		sqlStr = sqlStr + " 	where "
'		sqlStr = sqlStr + " 	getdate() between s.sale_startdate and s.sale_enddate"
'		sqlStr = sqlStr + " 	and s.sale_status=6"
'		sqlStr = sqlStr + " ) as t"
'		sqlStr = sqlStr + " 	on i.itemid = t.itemid"
		sqlStr = sqlStr + " where s.shopitemid is NULL"
		sqlStr = sqlStr + " and i.itemid<>0 " & sqlsearch
		sqlStr = sqlStr + " and d1.defaultmargin is not null"       '' streetshop000 마진설정된 브랜드만

		''rw sqlStr
		''rsget.Open sqlStr,dbget,1
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close



		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + ""
		sqlStr = sqlStr + " i.itemid, IsNULL(o.itemoption,'0000') as itemoption, i.itemname,s.shopitemid, i.makerid,"
		sqlStr = sqlStr + " (i.sellcash + IsNULL(o.optaddprice,0)) as sellcash, IsNULL(i.mwdiv,'') as mwdiv, IsNull(o.optionname,'') as opt2name,"
        sqlStr = sqlStr + " (i.orgprice + IsNULL(o.optaddprice,0)) as onlineitemorgprice"
        sqlStr = sqlStr + " ,(case when t.itemid is not null then 'Y' else 'N' end) as termsale"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlStr + " 	on i.itemid=o.itemid "		'--and o.isusing='Y'
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " 	on s.itemgubun='10' and i.itemid=s.shopitemid and s.itemoption= IsNULL(o.itemoption,'0000')"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d1 "
		sqlStr = sqlStr + " 	on i.makerid=d1.makerid and d1.shopid='streetshop000'"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select si.itemid"                           ''distinct => group by
		sqlStr = sqlStr + " 	from db_event.dbo.tbl_saleItem si"
		sqlStr = sqlStr + " 	join db_event.dbo.tbl_sale s"
		sqlStr = sqlStr + " 		on si.sale_code = s.sale_code"
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 	getdate() between s.sale_startdate and s.sale_enddate"
		sqlStr = sqlStr + " 	and s.sale_status=6"
		sqlStr = sqlStr + " 	group by si.itemid"
		sqlStr = sqlStr + " ) as t"
		sqlStr = sqlStr + " 	on i.itemid = t.itemid"
		sqlStr = sqlStr + " where s.shopitemid is NULL"
		sqlStr = sqlStr + " and i.itemid<>0 " & sqlsearch
		sqlStr = sqlStr + " and d1.defaultmargin is not null"       '' streetshop000 마진설정된 브랜드만
		sqlStr = sqlStr + " order by i.makerid, i.itemid desc"

		''rw sqlStr
	    ''response.end

		rsget.pagesize = FPageSize
		''rsget.Open sqlStr,dbget,1
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		''FTotalCount = rsget.RecordCount

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopOneItem

					FItemList(i).ftermsale        = rsget("termsale")
					FItemList(i).Fitemgubun         = "10"
					FItemList(i).Fshopitemid        = rsget("itemid")
					FItemList(i).Fitemoption     	= rsget("itemoption")
					FItemList(i).Fmakerid           = rsget("makerid")
					FItemList(i).Fshopitemname      = db2html(rsget("itemname"))
					FItemList(i).Fshopitemoptionname= db2html(rsget("opt2name"))
					FItemList(i).Fshopitemprice     = rsget("sellcash")
					FItemList(i).FmwDiv	= rsget("mwdiv")    ''온라인 매입구분
					FItemList(i).FOnlineitemorgprice= rsget("onlineitemorgprice")
					FItemList(i).Fshopsuplycash     = CLng(FItemList(i).Fshopitemprice * (100-FItemList(i).Fdefaultmargin)/100)

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub


	''미등록상품
	public function GetLinkNotRegList()
		dim sqlStr,i

		sqlStr = " select  count(T.itemid) as cnt"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + "("
		sqlStr = sqlStr + " select o.itemid,o.itemoption,o.opt2name,s.shopitemid"
		sqlStr = sqlStr + " from (select  i.itemid,IsNull(v.itemoption,'0000') as itemoption ,"
		sqlStr = sqlStr + " IsNull(v.optionname,'') as opt2name"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid and v.isusing='Y'"
		sqlStr = sqlStr + " where i.itemid<>0) as o"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s on s.itemgubun='10' and o.itemid=s.shopitemid and o.itemoption=s.itemoption"
		sqlStr = sqlStr + " where s.shopitemid is NULL"
		sqlStr = sqlStr + ") as T ,"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " , (select distinct makerid from [db_shop].[dbo].tbl_shop_designer) d"
		'sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on i.makerid=p.id"
		sqlStr = sqlStr + " where i.itemid<>0"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.itemid=T.itemid"
		sqlStr = sqlStr + " and i.makerid=d.makerid"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select  top " + CStr(FPageSize*FCurrPage) + " T.itemid, "
		sqlStr = sqlStr + "i.itemname, T.itemoption, T.opt2name, i.makerid, i.sellcash"
		''sqlStr = sqlStr + ",d.adminopen"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + "("
		sqlStr = sqlStr + " select o.itemid,o.itemoption,o.opt2name,s.shopitemid"
		sqlStr = sqlStr + " from (select  i.itemid,IsNull(v.itemoption,'0000') as itemoption ,"
		sqlStr = sqlStr + " IsNull(v.optionname,'') as opt2name"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid and v.isusing='Y'"
		sqlStr = sqlStr + " where i.itemid<>0) as o"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s on s.itemgubun='10' and o.itemid=s.shopitemid and o.itemoption=s.itemoption"
		sqlStr = sqlStr + " where s.shopitemid is NULL"
		sqlStr = sqlStr + ") as T ,"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " , (select distinct makerid from [db_shop].[dbo].tbl_shop_designer) d"
		'sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on i.makerid=p.id"
		sqlStr = sqlStr + " where i.itemid<>0"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.itemid=T.itemid"
		sqlStr = sqlStr + " and i.makerid=d.makerid"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if

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
					set FItemList(i) = new COffShopOneItem
					FItemList(i).Fitemgubun         = "10"
					FItemList(i).Fshopitemid        = rsget("itemid")
					FItemList(i).Fitemoption     	= rsget("itemoption")
					FItemList(i).Fmakerid           = rsget("makerid")
					FItemList(i).Fshopitemname      = db2html(rsget("itemname"))
					FItemList(i).Fshopitemoptionname= db2html(rsget("opt2name"))
					FItemList(i).Fshopitemprice     = rsget("sellcash")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end function

	''온라인과 가격이 다른 리스트 '/admin/offshop/localeItem/pop_localeItem_input.asp
	'//admin/offshop/shopitemlist.asp
	public function GetOffShopPriceDiffItemList()
		dim sqlStr,i , sqlsearch

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and s.shopitemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and s.shopitemid in (" + FRectItemid + ")"
            end if
        end if
		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and s.makerid='" + FRectDesigner + "'"
		end if
		sqlsearch = sqlsearch + " and s.shopitemprice<>i.sellcash+IsNULL(o.optaddprice,0)"
        if (FRectPriceRow<>"") then
            sqlsearch = sqlsearch + " and s.shopitemprice<i.sellcash+IsNULL(o.optaddprice,0)"
        end if
        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and s.catecdl='" + FRectCDL + "'"
        end if
        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and s.catecdm='" + FRectCDM + "'"
        end if
        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and s.catecdn='" + FRectCDS + "'"
        end if
		if FRectpublicbarcode <> "" then
            sqlsearch = sqlsearch + " and s.extbarcode = '"&FRectpublicbarcode&"'"
		end if
        if (FRectOnlineExpiredItem<>"") then
            sqlsearch = sqlsearch + " and i.sellyn='N'"
            sqlsearch = sqlsearch + " and i.isusing='N'"
            sqlsearch = sqlsearch + " and i.danjongyn in ('Y','M')"
            sqlsearch = sqlsearch + " and datediff(d,i.regdate,getdate())>91"
        end if
        if (FRectOnlyUsing<>"") then
            sqlsearch = sqlsearch + " and s.isusing='" & FRectOnlyUsing & "'"
        end if
        if (FRectCenterMwdiv<>"") then
            if (FRectCenterMwdiv="X") then
                sqlsearch = sqlsearch + " and s.centermwdiv is NULL"
            else
                sqlsearch = sqlsearch + " and s.centermwdiv='" + FRectCenterMwdiv + "'"
            end if
        end if
		if (FRectOnlineMwDiv <> "") then
			if (FRectOnlineMwDiv = "X") then
				sqlsearch = sqlsearch + " and i.mwdiv not in ('M', 'W', 'U') "
			else
				sqlsearch = sqlsearch + " and i.mwdiv = '" & CStr(FRectOnlineMwDiv) & "' "
			end if
		end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + " 	on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o with (nolock)"
		sqlStr = sqlStr + " 	on s.itemgubun='10' and s.shopitemid=o.itemid"
		sqlStr = sqlStr + " 	and s.itemoption=IsNULL(o.itemoption,'0000')"
		sqlStr = sqlStr + " where s.shopitemid=i.itemid " ''& sqlsearch
		sqlStr = sqlStr + sqlsearch

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr ,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption,"
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname, s.orgsellprice, s.discountsellprice, s.shopbuyprice, "
		sqlStr = sqlStr + " s.shopitemprice, s.shopsuplycash, s.isusing, s.regdate, s.onofflinkyn"
		sqlStr = sqlStr + " ,s.centermwdiv,i.sellyn, i.danjongyn,"
		sqlStr = sqlStr + " IsNull(i.orgprice,0) as onlineitemorgprice, IsNull(i.sellcash,0) as onlineitemprice,"
		sqlStr = sqlStr + " IsNULL(i.smallimage,'') as imgsmall, IsNULL(s.offimgsmall,'') as offimgsmall"
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice, isNull(a.itemid,0) as stockitemid"
		sqlStr = sqlStr & " ,c.nmlarge, c.nmmid, c.nmsmall" & vbcrlf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + " 	on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o with (nolock)"
		sqlStr = sqlStr + " 	on s.itemgubun='10' and s.shopitemid=o.itemid"
		sqlStr = sqlStr + " 	and s.itemoption=IsNULL(o.itemoption,'0000')"
		sqlStr = sqlStr + " left outer join db_summary.dbo.tbl_current_logisstock_summary as a with (nolock)"
		sqlStr = sqlStr + "  on s.shopitemid = a.itemid and s.itemoption = a.itemoption and s.itemgubun= a.itemgubun "
		sqlStr = sqlStr & " left join [db_item].[dbo].vw_category c with (nolock)" & vbcrlf
		sqlStr = sqlStr & "		on s.catecdl=c.cdlarge and s.catecdm=c.cdmid and s.catecdn=c.cdsmall" & vbcrlf
		sqlStr = sqlStr + " where s.shopitemid=i.itemid " ''& sqlsearch
		sqlStr = sqlStr + sqlsearch
		sqlStr = sqlStr + " order by s.itemgubun desc, s.shopitemid desc"

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr ,dbget,1

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
				set FItemList(i) = new COffShopOneItem

				FItemList(i).fstockitemid         = rsget("stockitemid")
				FItemList(i).fonofflinkyn         = rsget("onofflinkyn")
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).Fisusing           = rsget("isusing")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).FOnLineItemprice	= rsget("onlineitemprice")
                FItemList(i).FOnLineItemOrgprice= rsget("onlineitemorgprice")
    			FItemList(i).FOnlineOptaddprice = rsget("optaddprice")
    			FItemList(i).FOnlineOptaddbuyprice = rsget("optaddbuyprice")
				FItemList(i).Fdiscountsellprice = rsget("discountsellprice")
				FItemList(i).Fshopbuyprice		= rsget("shopbuyprice")
                FItemList(i).FShopItemOrgprice  = rsget("orgsellprice")
				FItemList(i).FimageSmall     = rsget("imgsmall")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
                FItemList(i).Fcentermwdiv   = rsget("centermwdiv")
                FItemList(i).Fsellyn        = rsget("sellyn")
                FItemList(i).FOnlinedanjongyn     = rsget("danjongyn")
				FItemList(i).FCateCDLName	= db2html(rsget("nmlarge"))
				FItemList(i).FCateCDMName	= db2html(rsget("nmmid"))
				FItemList(i).FCateCDSName	= db2html(rsget("nmsmall"))

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	''온라인과 가격이 다른 리스트
	'/admin/offshop/localeItem/pop_localeItem_input.asp
	'//common/offshop/pop_itemAddInfo_off.asp		'/common/offshop/pop_itemAddInfo2_off.asp
	public function GetcontractOffShopPriceDiffItemList()
		dim sqlStr,i , sqlsearch

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and s.makerid='" + FRectDesigner + "'"
		end if
		sqlStr = sqlStr + " and s.shopitemprice<>i.sellcash+IsNULL(o.optaddprice,0)"

        if (FRectPriceRow<>"") then
            sqlsearch = sqlsearch + " and s.shopitemprice<i.sellcash+IsNULL(o.optaddprice,0)"
        end if

        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and s.catecdl='" + FRectCDL + "'"
        end if

        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and s.catecdm='" + FRectCDM + "'"
        end if

        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and s.catecdn='" + FRectCDS + "'"
        end if

        if (FRectOnlineExpiredItem<>"") then
            sqlsearch = sqlsearch + " and i.sellyn='N'"
            sqlsearch = sqlsearch + " and i.isusing='N'"
            sqlsearch = sqlsearch + " and i.danjongyn in ('Y','M')"
            sqlsearch = sqlsearch + " and datediff(d,i.regdate,getdate())>91"
        end if

        if (FRectOnlyUsing<>"") then
            sqlsearch = sqlsearch + " and s.isusing='" & FRectOnlyUsing & "'"
        end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_designer sd" & VbCRLF
		sqlStr = sqlStr + " 	on sd.shopid='"&frectshopid&"' and s.makerid=sd.makerid" & VbCRLF
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o "
		sqlStr = sqlStr + " 	on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr + " where s.shopitemid=i.itemid " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption,"
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname, s.orgsellprice, s.discountsellprice"
		sqlStr = sqlStr + " ,s.shopitemprice, s.isusing, s.regdate"
		sqlStr = sqlStr + " ,s.centermwdiv,i.sellyn, i.danjongyn, sd.defaultmargin,sd.defaultsuplymargin,sd.comm_cd"
		sqlStr = sqlStr + " ,IsNull(i.orgprice,0) as onlineitemorgprice, IsNull(i.sellcash,0) as onlineitemprice,"
		sqlStr = sqlStr + " IsNULL(i.smallimage,'') as imgsmall, IsNULL(s.offimgsmall,'') as offimgsmall"
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
		sqlStr = sqlStr + " ,(CASE"
		sqlStr = sqlStr + " 	when s.shopsuplycash = 0 and sd.comm_cd in ('B011','B012','B013')"		'/매입가가 0 ,텐텐위탁, 업체위탁 ,출고위탁
		sqlStr = sqlStr + " 		then convert(int,s.shopitemprice*(100-IsNULL(sd.defaultmargin,100))/100)"
		sqlStr = sqlStr + " 	else s.shopsuplycash"
		sqlStr = sqlStr + "	end) as shopsuplycash"
		'sqlStr = sqlStr + " , s.shopsuplycash"
		sqlStr = sqlStr + " ,(CASE" & VbCRLF
		sqlStr = sqlStr + " 	when s.shopbuyprice = 0 and sd.comm_cd in ('B011','B012','B013')"		'/매장출고가 0 ,텐텐위탁, 업체위탁 ,출고위탁
		sqlStr = sqlStr + " 		then convert(int,s.shopitemprice*(100-IsNULL(sd.defaultsuplymargin,100))/100)"
		sqlStr = sqlStr + "		else s.shopbuyprice"
		sqlStr = sqlStr + "	end) as shopbuyprice"
		'sqlStr = sqlStr + " , s.shopbuyprice"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_designer sd" & VbCRLF
		sqlStr = sqlStr + " 	on sd.shopid='"&frectshopid&"' and s.makerid=sd.makerid" & VbCRLF
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o  "
		sqlStr = sqlStr + " 	on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr + " where s.shopitemid=i.itemid " & sqlsearch
		sqlStr = sqlStr + " order by s.itemgubun desc, s.shopitemid desc"

		'response.write sqlStr &"<br>"
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
				set FItemList(i) = new COffShopOneItem

				FItemList(i).fdefaultmargin = rsget("defaultmargin")
				FItemList(i).fdefaultsuplymargin = rsget("defaultsuplymargin")
				FItemList(i).fcomm_cd         = rsget("comm_cd")
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).Fisusing           = rsget("isusing")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).FOnLineItemprice	= rsget("onlineitemprice")
                FItemList(i).FOnLineItemOrgprice= rsget("onlineitemorgprice")
    			FItemList(i).FOnlineOptaddprice = rsget("optaddprice")
    			FItemList(i).FOnlineOptaddbuyprice = rsget("optaddbuyprice")
				FItemList(i).Fdiscountsellprice = rsget("discountsellprice")
				FItemList(i).Fshopbuyprice		= rsget("shopbuyprice")
                FItemList(i).FShopItemOrgprice  = rsget("orgsellprice")
				FItemList(i).FimageSmall     = rsget("imgsmall")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
                FItemList(i).Fcentermwdiv   = rsget("centermwdiv")
                FItemList(i).Fsellyn        = rsget("sellyn")
                FItemList(i).FOnlinedanjongyn     = rsget("danjongyn")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public function GetOffLineItemList()
		dim sqlStr,i
		sqlStr = " select count(s.shopitemid) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " where makerid='" + FRectDesigner + "'"

		if FRectOnlyOffLine="on" then
			sqlStr = sqlStr + " and itemgubun<>'10'"
		end if

		if FRectOnlyUsing="on" then
			sqlStr = sqlStr + " and isusing='Y'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption, "
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname,"
		sqlStr = sqlStr + " s.shopitemprice, s.shopsuplycash, s.isusing, s.regdate, s.extbarcode,"
		sqlStr = sqlStr + " IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist, IsNULL(s.offimgsmall,'') as offimgsmall"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " where s.makerid='" + FRectDesigner + "'"

		if FRectOnlyOffLine="on" then
			sqlStr = sqlStr + " and s.itemgubun<>'10'"
		end if

		if FRectOnlyUsing="on" then
			sqlStr = sqlStr + " and s.isusing='Y'"
		end if

		sqlStr = sqlStr + " order by s.itemgubun desc, s.shopitemid desc"

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
				set FItemList(i) = new COffShopOneItem
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).Fisusing           = rsget("isusing")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).Fextbarcode 		= rsget("extbarcode")
				FItemList(i).FOffimgMain	= rsget("offimgmain")
				FItemList(i).FOffimgList	= rsget("offimglist")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	'특정 샵과 계약된 내역만 가져 옴
	'//common/offshop/pop_itemAddInfo_off.asp	'/common/offshop/pop_itemAddInfo2_off.asp
	public function GetcontractShopItemList()
		dim sqlStr,i , sqlsearch

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and shopitemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and shopitemid in (" + FRectItemid + ")"
            End If
        End If

		if FRectItemName<>"" then
			sqlsearch = sqlsearch + " and shopitemname like '%" + FRectItemName + "%'"
		end if

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and s.makerid='" + FRectDesigner + "'"
		end if

		if FRectItemgubun<>"" then
			sqlsearch = sqlsearch + " and s.itemgubun='" + FRectItemgubun + "'"
		end if

		if FRectOnlyUsing<>"" then
			sqlsearch = sqlsearch + " and s.isusing='" + FRectOnlyUsing + "'"
		end if

        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and s.catecdl='" + FRectCDL + "'"
        end if

        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and s.catecdm='" + FRectCDM + "'"
        end if

        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and s.catecdn='" + FRectCDS + "'"
        end if

        if (FRectOnlineExpiredItem<>"") then
            sqlsearch = sqlsearch + " and i.sellyn='N'"
            sqlsearch = sqlsearch + " and i.danjongyn in ('Y','M')"
            sqlsearch = sqlsearch + " and datediff(d,i.regdate,getdate())>91"
        end if

		if frectsaleflg = "on" then
		else
		    sqlsearch = sqlsearch + " and s.orgsellprice = s.shopitemprice"
		end if

		sqlStr = " select count(*) as cnt"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_designer sd" & VbCRLF
		sqlStr = sqlStr + " 	on sd.shopid='"&frectshopid&"'"
		sqlStr = sqlStr + " 	and s.makerid=sd.makerid" & VbCRLF
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	on s.shopitemid=i.itemid and s.itemgubun='10'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlStr + "		on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"

		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit function

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption,s.makerid, s.shopitemname, s.shopitemoptionname"
		sqlStr = sqlStr + " ,s.orgsellprice,s.shopitemprice, s.isusing, s.regdate, s.extbarcode"
		sqlStr = sqlStr + " ,s.discountsellprice ,s.centermwdiv, IsNULL(s.offimgsmall,'') as offimgsmall"
		sqlStr = sqlStr + " ,i.sellyn, i.danjongyn,IsNull(i.orgprice,0) as onlineitemorgprice"
		sqlStr = sqlStr + " ,IsNull(i.sellcash,0) as onlineitemprice,IsNULL(i.smallimage,'') as imgsmall"
		sqlStr = sqlStr + " ,sd.defaultmargin ,sd.defaultsuplymargin , sd.Comm_cd"
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
		sqlStr = sqlStr + " ,(CASE"
		sqlStr = sqlStr + " 	when s.shopsuplycash = 0 and sd.comm_cd in ('B011','B012','B013')"		'/매입가가 0 ,텐텐위탁, 업체위탁 ,출고위탁
		sqlStr = sqlStr + " 		then convert(int,s.shopitemprice*(100-IsNULL(sd.defaultmargin,100))/100)"
		sqlStr = sqlStr + " 	else s.shopsuplycash"
		sqlStr = sqlStr + "	end) as shopsuplycash"
		sqlStr = sqlStr + " ,(CASE" & VbCRLF
		sqlStr = sqlStr + " 	when s.shopbuyprice = 0 and sd.comm_cd in ('B011','B012','B013')"		'/매장출고가 0 ,텐텐위탁, 업체위탁 ,출고위탁
		sqlStr = sqlStr + " 		then convert(int,s.shopitemprice*(100-IsNULL(sd.defaultsuplymargin,100))/100)"
		sqlStr = sqlStr + "		else s.shopbuyprice"
		sqlStr = sqlStr + "	end) as shopbuyprice"
		sqlStr = sqlStr + " , isNULL(i.mwdiv,'') as mwdiv"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " join db_shop.dbo.tbl_shop_designer sd" & VbCRLF
		sqlStr = sqlStr + " 	on sd.shopid='"&frectshopid&"'"
		sqlStr = sqlStr + " 	and s.makerid=sd.makerid" & VbCRLF
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	on s.shopitemid=i.itemid and s.itemgubun='10'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o"
		sqlStr = sqlStr + "		on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by s.itemgubun desc, s.shopitemid desc"

		'response.write sqlStr & "<br>"
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
				set FItemList(i) = new COffShopOneItem

				FItemList(i).fComm_cd = rsget("Comm_cd")
				FItemList(i).fdefaultmargin = rsget("defaultmargin")
				FItemList(i).fdefaultsuplymargin = rsget("defaultsuplymargin")
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).Fisusing           = rsget("isusing")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).Fextbarcode 		= rsget("extbarcode")
				FItemList(i).FOnLineItemprice	= rsget("onlineitemprice")
                FItemList(i).FOnlineitemorgprice= rsget("onlineitemorgprice")
			    FItemList(i).FOnlineOptaddprice = rsget("optaddprice")
			    FItemList(i).FOnlineOptaddbuyprice = rsget("optaddbuyprice")
				FItemList(i).Fdiscountsellprice = rsget("discountsellprice")
				FItemList(i).Fshopbuyprice		= rsget("shopbuyprice")
                FItemList(i).FShopItemOrgprice  = rsget("orgsellprice")
				FItemList(i).FimageSmall     = rsget("imgsmall")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
                FItemList(i).Fcentermwdiv  = rsget("centermwdiv")
                FItemList(i).Fsellyn        = rsget("sellyn")
                FItemList(i).FOnlinedanjongyn     = rsget("danjongyn")
                FItemList(i).Fmwdiv			= rsget("mwdiv")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	'//admin/offshop/localeItem/pop_localeItem_input.asp 	'//common/offshop/shopitemlist_Etc.asp
	'/admin/offshop/shopitemlist.asp
	public function GetOffNOnLineShopItemList()
		dim sqlStr,i , sqlsearch

		if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and s.shopitemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and s.shopitemid in (" + FRectItemid + ")"
            end if
        end if
		if FRectItemName<>"" then
			sqlsearch = sqlsearch + " and shopitemname like '%" + FRectItemName + "%'"
		end if
		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and s.makerid='" + FRectDesigner + "'"
		end if
		if FRectItemgubun<>"" then
			sqlsearch = sqlsearch + " and s.itemgubun='" + FRectItemgubun + "'"
		end if
		if FRectOnlyUsing<>"" then
			sqlsearch = sqlsearch + " and s.isusing='" + FRectOnlyUsing + "'"
		end if
        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and s.catecdl='" + FRectCDL + "'"
        end if
        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and s.catecdm='" + FRectCDM + "'"
        end if
        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and s.catecdn='" + FRectCDS + "'"
        end if
		if frectpublicbarcode <> "" then
			sqlsearch = sqlsearch + " and s.extbarcode='"&frectpublicbarcode&"'"
		end if
        if (FRectOnlineExpiredItem<>"") then
            sqlsearch = sqlsearch + " and i.sellyn='N'"
            sqlsearch = sqlsearch + " and i.danjongyn in ('Y','M')"
            sqlsearch = sqlsearch + " and datediff(d,i.regdate,getdate())>91"
        end if
        if (FRectCenterMwdiv<>"") then
            if (FRectCenterMwdiv="X") then
                sqlsearch = sqlsearch + " and (s.centermwdiv is NULL or s.centermwdiv = '')"
            else
                sqlsearch = sqlsearch + " and s.centermwdiv='" + FRectCenterMwdiv + "'"
            end if
        end if
		if (FRectOnlineMwDiv <> "") then
			if (FRectOnlineMwDiv = "X") then
				sqlsearch = sqlsearch + " and i.mwdiv not in ('M', 'W', 'U') "
			else
				sqlsearch = sqlsearch + " and i.mwdiv = '" & CStr(FRectOnlineMwDiv) & "' "
			end if
		end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + " 	on s.shopitemid=i.itemid and s.itemgubun='10'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o with (nolock)"
		sqlStr = sqlStr + "		on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr & " left join [db_item].[dbo].vw_category c with (nolock)" & vbcrlf
		sqlStr = sqlStr & "		on s.catecdl=c.cdlarge and s.catecdm=c.cdmid and s.catecdn=c.cdsmall" & vbcrlf
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption,"
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname, s.orgsellprice, s.onofflinkyn"
		sqlStr = sqlStr + " ,s.shopitemprice, s.shopsuplycash, s.isusing, s.regdate, s.extbarcode, s.discountsellprice , s.shopbuyprice,"
		sqlStr = sqlStr + " s.centermwdiv, i.sellyn, i.danjongyn,"
		sqlStr = sqlStr + " IsNull(i.orgprice,0) as onlineitemorgprice, IsNull(i.sellcash,0) as onlineitemprice ,"
		sqlStr = sqlStr + " IsNULL(i.smallimage,'') as imgsmall, IsNULL(s.offimgsmall,'') as offimgsmall"
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
		sqlStr = sqlStr + " , isNULL(i.mwdiv,'') as mwdiv, isNull(a.itemid,0) as stockitemid,c.nmlarge, c.nmmid, c.nmsmall"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + " 	on s.shopitemid=i.itemid and s.itemgubun='10'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o with (nolock)"
		sqlStr = sqlStr + "		on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr + " left outer join db_summary.dbo.tbl_current_logisstock_summary as a with (nolock)"
		sqlStr = sqlStr + "  on s.shopitemid = a.itemid and s.itemoption = a.itemoption and s.itemgubun= a.itemgubun "
		sqlStr = sqlStr & " left join [db_item].[dbo].vw_category c with (nolock)" & vbcrlf
		sqlStr = sqlStr & "		on s.catecdl=c.cdlarge and s.catecdm=c.cdmid and s.catecdn=c.cdsmall" & vbcrlf
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by s.itemgubun, s.shopitemid desc, s.itemoption" ''2015/05/22 수정 이문재이사요청

		'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new COffShopOneItem

				FItemList(i).fstockitemid         = rsget("stockitemid")
				FItemList(i).fonofflinkyn         = rsget("onofflinkyn")
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).Fisusing           = rsget("isusing")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).Fextbarcode 		= rsget("extbarcode")
				FItemList(i).FOnLineItemprice	= rsget("onlineitemprice")
                FItemList(i).FOnlineitemorgprice= rsget("onlineitemorgprice")
			    FItemList(i).FOnlineOptaddprice = rsget("optaddprice")
			    FItemList(i).FOnlineOptaddbuyprice = rsget("optaddbuyprice")
				FItemList(i).Fdiscountsellprice = rsget("discountsellprice")
				FItemList(i).Fshopbuyprice		= rsget("shopbuyprice")
                FItemList(i).FShopItemOrgprice  = rsget("orgsellprice")
				FItemList(i).FimageSmall     = rsget("imgsmall")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl & "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
                FItemList(i).Fcentermwdiv  = rsget("centermwdiv")
                FItemList(i).Fsellyn        = rsget("sellyn")
                FItemList(i).FOnlinedanjongyn     = rsget("danjongyn")
                FItemList(i).Fmwdiv			= rsget("mwdiv")
				FItemList(i).FCateCDLName	= db2html(rsget("nmlarge"))
				FItemList(i).FCateCDMName	= db2html(rsget("nmmid"))
				FItemList(i).FCateCDSName	= db2html(rsget("nmsmall"))

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	public function GetOffNOnLineGiftItemList()
		dim sqlStr,i , sqlsearch

		'80 : OFF사은품, 85 : ON사은품
		sqlsearch = sqlsearch + " and s.itemgubun in ('80', '85') "

		if FRectItemId<>"" then
			sqlsearch = sqlsearch + " and shopitemid=" + CStr(FRectItemId)
		end if

		if FRectItemName<>"" then
			sqlsearch = sqlsearch + " and shopitemname like '%" + FRectItemName + "%'"
		end if

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and s.makerid='" + FRectDesigner + "'"
		end if

		if FRectItemgubun<>"" then
			sqlsearch = sqlsearch + " and s.itemgubun='" + FRectItemgubun + "'"
		end if

		if FRectOnlyUsing<>"" then
			sqlsearch = sqlsearch + " and s.isusing='" + FRectOnlyUsing + "'"
		end if

        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and s.catecdl='" + FRectCDL + "'"
        end if

        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and s.catecdm='" + FRectCDM + "'"
        end if

        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and s.catecdn='" + FRectCDS + "'"
        end if

		if frectpublicbarcode <> "" then
			sqlsearch = sqlsearch + " and s.extbarcode='"&frectpublicbarcode&"'"
		end if

		if FRectSizeYn <> "" then
        	if FRectSizeYn="Y" then
	        	sqlsearch = sqlsearch + " and IsNull(s.volX,0) > 0 "
	        else
				sqlsearch = sqlsearch + " and IsNull(s.volX,0) <= 0 "
			end if
		end if

        if FRectIsWeight<>"" then
        	if FRectIsWeight="Y" then
	        	sqlsearch = sqlsearch + " and IsNull(s.itemWeight,0) > 0 "
	        else
				sqlsearch = sqlsearch + " and IsNull(s.itemWeight,0) <= 0 "
			end if
        end if

        if (FRectOnlineExpiredItem<>"") then
            sqlsearch = sqlsearch + " and i.sellyn='N'"
            sqlsearch = sqlsearch + " and i.danjongyn in ('Y','M')"
            sqlsearch = sqlsearch + " and datediff(d,i.regdate,getdate())>91"
        end if

		sqlStr = " select count(*) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"

		if (FRectOnlineExpiredItem<>"") then
		    sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item i"
		    sqlStr = sqlStr + " on s.shopitemid=i.itemid and s.itemgubun='10'"
		end if

		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption,"
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname, s.orgsellprice, s.onofflinkyn"
		sqlStr = sqlStr + " ,s.shopitemprice, s.shopsuplycash, s.isusing, s.regdate, s.extbarcode, s.discountsellprice , s.shopbuyprice,"
		sqlStr = sqlStr + " s.centermwdiv, i.sellyn, i.danjongyn,"
		sqlStr = sqlStr + " IsNull(i.orgprice,0) as onlineitemorgprice, IsNull(i.sellcash,0) as onlineitemprice ,"
		sqlStr = sqlStr + " IsNULL(i.smallimage,'') as imgsmall, IsNULL(s.offimgsmall,'') as offimgsmall"
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
		sqlStr = sqlStr + " , isNULL(i.mwdiv,'') as mwdiv, IsNull(s.volX,0) as volX, IsNull(s.volY,0) as volY, IsNull(s.volZ,0) as volZ, IsNull(s.itemWeight,0) as itemWeight"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on (s.shopitemid=i.itemid) and s.itemgubun='10'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on "
		sqlStr = sqlStr + "	s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by s.regdate desc"
		'response.write sqlStr &"<Br>"
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
				set FItemList(i) = new COffShopOneItem

				FItemList(i).fonofflinkyn         = rsget("onofflinkyn")
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).Fisusing           = rsget("isusing")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).Fextbarcode 		= rsget("extbarcode")
				FItemList(i).FOnLineItemprice	= rsget("onlineitemprice")
                FItemList(i).FOnlineitemorgprice= rsget("onlineitemorgprice")
			    FItemList(i).FOnlineOptaddprice = rsget("optaddprice")
			    FItemList(i).FOnlineOptaddbuyprice = rsget("optaddbuyprice")
				FItemList(i).Fdiscountsellprice = rsget("discountsellprice")
				FItemList(i).Fshopbuyprice		= rsget("shopbuyprice")
                FItemList(i).FShopItemOrgprice  = rsget("orgsellprice")
				FItemList(i).FimageSmall     = rsget("imgsmall")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl & "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
                FItemList(i).Fcentermwdiv  = rsget("centermwdiv")
                FItemList(i).Fsellyn        = rsget("sellyn")
                FItemList(i).FOnlinedanjongyn     = rsget("danjongyn")
                FItemList(i).Fmwdiv			= rsget("mwdiv")
			    FItemList(i).FvolX        	  = rsget("volX")
			    FItemList(i).FvolY        	  = rsget("volY")
			    FItemList(i).FvolZ        	  = rsget("volZ")
                FItemList(i).FitemWeight   	  = rsget("itemWeight")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	'//common/offshop/popoffitemreg_Etc.asp
	public function GetOffNOnLineShoponeItem()
		dim sqlStr,i , sqlsearch

		if frectitemid <> "" then
			sqlsearch = sqlsearch & " and shopitemid = "&frectitemid&""
		end if
		if frectitemgubun <> "" then
			sqlsearch = sqlsearch & " and itemgubun = '"&frectitemgubun&"'"
		end if
		if frectitemoption <> "" then
			sqlsearch = sqlsearch & " and itemoption = '"&frectitemoption&"'"
		end if

		sqlStr = "select top 1"
		sqlStr = sqlStr & " itemgubun,shopitemid,itemoption,makerid,shopitemname,shopitemoptionname"
		sqlStr = sqlStr & " ,orgsellprice,shopitemprice ,shopsuplycash,shopbuyprice, discountsellprice"
		sqlStr = sqlStr & " ,centermwdiv,extbarcode ,vatinclude ,isusing,shopsuplycash"
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item"
		sqlStr = sqlStr & " where shopitemid <>0 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1

        ftotalcount = rsget.RecordCount

		if not rsget.Eof then
			set FOneItem = new COffShopOneItem

			FOneItem.fshopsuplycash         = rsget("shopsuplycash")
			FOneItem.fisusing         = rsget("isusing")
			FOneItem.Fitemgubun         = rsget("itemgubun")
			FOneItem.Fshopitemid        = rsget("shopitemid")
			FOneItem.Fitemoption     	= rsget("itemoption")
			FOneItem.Fmakerid           = rsget("makerid")
			FOneItem.Fshopitemname      = db2html(rsget("shopitemname"))
			FOneItem.Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
            FOneItem.FShopItemOrgprice  = rsget("orgsellprice")
			FOneItem.Fshopitemprice     = rsget("shopitemprice")
			FOneItem.fshopbuyprice     = rsget("shopbuyprice")
			FOneItem.fdiscountsellprice     = rsget("discountsellprice")
			FOneItem.fcentermwdiv     = rsget("centermwdiv")
			FOneItem.fextbarcode     = rsget("extbarcode")
			FOneItem.fvatinclude     = rsget("vatinclude")

		end if

		rsget.Close
	end function

	public Sub GetOffLineJumunByItemID()
		dim sqlStr,i

		sqlStr = " select "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption, "
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname,"
		sqlStr = sqlStr + " s.shopitemprice, s.shopsuplycash, s.shopbuyprice, s.isusing, i.smallimage as imgsmall,"
		sqlStr = sqlStr + " IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist, IsNULL(s.offimgsmall,'') as offimgsmall,"
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, IsNULL(i.buycash,0) as onlinebuycash, i.deliverytype, d.defaultmargin,d.defaultsuplymargin "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " ,[db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " where s.shopitemid <>0"
		sqlStr = sqlStr + " and s.shopitemid='" + FRectItemID + "'"
		sqlStr = sqlStr + " and s.isusing='Y'"
		sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
		sqlStr = sqlStr + " and d.makerid=s.makerid"

		if FRectDesignerjungsangubun<>"" then
			sqlStr = sqlStr + " and d.chargediv in (" + FRectDesignerjungsangubun + ")"
		end if

		sqlStr = sqlStr + " order by s.itemgubun asc, s.shopitemid desc, s.itemoption"

		'response.write sqlStr
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
				set FItemList(i) = new COffShopOneItem
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).FimageSmall     = rsget("imgsmall")

				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

				FItemList(i).Fsellyn               = rsget("sellyn")
				FItemList(i).Flimityn              = rsget("limityn")
				FItemList(i).Flimitno              = rsget("limitno")
				FItemList(i).Flimitsold            = rsget("limitsold")
				FItemList(i).FMakerMargin    = rsget("defaultmargin")
				FItemList(i).FShopMargin    = rsget("defaultsuplymargin")
				FItemList(i).FOffimgMain	= rsget("offimgmain")
				FItemList(i).FOffimgList	= rsget("offimglist")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				FItemList(i).FdeliveryType  = rsget("deliverytype")
                FItemList(i).Fshopbuyprice  = rsget("shopbuyprice")
                FItemList(i).Fonlinebuycash = rsget("onlinebuycash")
				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	public Sub GetOffLineJumunByOneItemCode()
		dim sqlStr,i

		'/2016.09.13 한용민 추가
		sqlStr = "select top 1 " & vbcrlf
		sqlStr = sqlStr & " u.countrylangcd, u.currencyunit, u.loginsite" & vbcrlf
		sqlStr = sqlStr & " , e.currencyChar, e.exchangeRate, e.multiplerate, e.linkPriceType, e.basedate, e.regdate" & vbcrlf
		sqlStr = sqlStr & " , e.lastupdate, e.reguserid, e.lastuserid" & vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & " join db_item.dbo.tbl_exchangeRate e" & vbcrlf
		sqlStr = sqlStr & " 	on u.countrylangcd = e.countrylangcd" & vbcrlf
		sqlStr = sqlStr & " 	and u.currencyunit = e.currencyunit" & vbcrlf
		'sqlStr = sqlStr & " 	and u.loginsite = e.sitename" & vbcrlf
		sqlStr = sqlStr & " 	and e.sitename='WSLWEB'" & vbcrlf
		sqlStr = sqlStr & " where userid = '" + FRectShopid + "'" & vbcrlf

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
		if  not rsget.EOF  then
    		Fcurrencyunit = rsget("currencyunit")
    		Floginsite = rsget("loginsite")
    		fcountrylangcd = rsget("countrylangcd")
    		fcurrencyChar = rsget("currencyChar")
			flinkPriceType = rsget("linkPriceType")
			fmultiplerate = rsget("multiplerate")
    	end if
    	rsget.Close

		sqlStr = " select top 1 " + VbCrlf
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption, "
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname,"
		sqlStr = sqlStr + " s.shopitemprice, s.shopsuplycash, s.shopbuyprice, s.isusing, i.smallimage as imgsmall,"
		sqlStr = sqlStr + " IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist, IsNULL(s.offimgsmall,'') as offimgsmall,"
		sqlStr = sqlStr + " i.sellcash, i.buycash, i.sellyn, i.limityn, i.limitno, i.limitsold, i.deliverytype, d.defaultmargin,d.defaultsuplymargin, d.chargediv, d.comm_cd "
		sqlStr = sqlStr + " ,isNULL(i.mwdiv,'') as mwdiv"
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"

		'/홀쎄일의 경우 대표화폐가 한화 일경우 초기값 셋팅
		if Floginsite="WSLWEB" then
			if (fcurrencyUnit = "KRW" or fcurrencyUnit = "WON") Then
				'해외판매가가 입력이 안되었을경우, 국내가격을 해외 가격에 넣는다.
				if flinkPriceType="1" then
					sqlStr = sqlStr & " ,(case when isNull(mlp.orgprice,0)=0 then isNull(s.shopitemprice,0) else isNull(mlp.orgprice,0) end) as orgprice"
				else
					sqlStr = sqlStr & " ,(case when isNull(mlp.orgprice,0)=0 then isNull(s.orgsellprice,0) else isNull(mlp.orgprice,0) end) as orgprice"
				end if
				sqlStr = sqlStr & " ,(case"
				sqlStr = sqlStr & " 	when isNull(mlp.orgprice,0)=0 then isNull(s.shopsuplycash,0)"
				sqlStr = sqlStr & " 	else round( (isnull(mlp.orgprice,0)*(100-IsNULL(d.defaultsuplymargin,100))/100) ,1) end) as suplyprice"
			else
				sqlStr = sqlStr & " ,isNull(mlp.orgprice,0) as orgprice"
				sqlStr = sqlStr & " ,round( (isnull(mlp.orgprice,0)*(100-IsNULL(d.defaultsuplymargin,100))/100) ,1) as suplyprice"
			end if
		else
			sqlStr = sqlStr & " ,0 as orgprice"
			sqlStr = sqlStr & " ,0 as suplyprice"
		end if

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	on s.itemgubun='" + FRectItemgubun + "' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on "
		sqlStr = sqlStr + "	s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " 	on d.shopid='" + FRectShopid + "' and s.makerid=d.makerid"

 		'해외 상품금액
	    sqlStr = sqlStr + " left outer join db_item.dbo.tbl_item_multiLang_price as mlp " & VbCrLf
	    sqlStr = sqlStr + "		on mlp.sitename ='"&Floginsite&"' and mlp.currencyunit = '"&Fcurrencyunit&"' and mlp.itemid = i.itemid "

		sqlStr = sqlStr + " where 1=1"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
		end if
        if (frectisusing<>"") then
            sqlStr = sqlStr + " and s.isusing='" + frectisusing + "'"
        end if

		sqlStr = sqlStr + " and s.itemgubun='" + FRectItemgubun + "'"
		sqlStr = sqlStr + " and s.shopitemid=" + CStr(FRectItemId) + ""
		sqlStr = sqlStr + " and s.itemoption='" + FRectItemOption + "'"

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		i=0
		if  not rsget.EOF  then
			set FOneItem = new COffShopOneItem
			FOneItem.Fitemgubun         = rsget("itemgubun")
			FOneItem.Fshopitemid        = rsget("shopitemid")
			FOneItem.Fitemoption     	= rsget("itemoption")
			FOneItem.Fmakerid           = rsget("makerid")
			FOneItem.Fshopitemname      = db2html(rsget("shopitemname"))
			FOneItem.Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
			FOneItem.Fshopitemprice     = rsget("shopitemprice")
			FOneItem.Fshopsuplycash     = rsget("shopsuplycash")
			FOneItem.FShopbuyprice      = rsget("Shopbuyprice")
			FOneItem.Fisusing			= rsget("isusing")
			FOneItem.FimageSmall     = rsget("imgsmall")

			if FOneItem.FimageSmall<>"" then
				FOneItem.FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FOneItem.Fshopitemid) + "/" + FOneItem.FimageSmall
			end if

			FOneItem.Fsellyn               = rsget("sellyn")
			FOneItem.Flimityn              = rsget("limityn")
			FOneItem.Flimitno              = rsget("limitno")
			FOneItem.Flimitsold            = rsget("limitsold")
			FOneItem.FMakerMargin    = rsget("defaultmargin")
			FOneItem.FShopMargin    = rsget("defaultsuplymargin")
			FOneItem.Fchargediv		= rsget("chargediv")
			FOneItem.FOffimgMain	= rsget("offimgmain")
			FOneItem.FOffimgList	= rsget("offimglist")
			FOneItem.FOffimgSmall	= rsget("offimgsmall")
			FOneItem.FdeliveryType  = rsget("deliverytype")
			if FOneItem.FOffimgMain<>"" then FOneItem.FOffimgMain = webImgUrl + "/offimage/offmain/i" + FOneItem.Fitemgubun + "/" + GetImageSubFolderByItemid(FOneItem.Fshopitemid) + "/" + FOneItem.FOffimgMain
			if FOneItem.FOffimgList<>"" then FOneItem.FOffimgList = webImgUrl + "/offimage/offlist/i" + FOneItem.Fitemgubun + "/" + GetImageSubFolderByItemid(FOneItem.Fshopitemid) + "/" + FOneItem.FOffimgList
			if FOneItem.FOffimgSmall<>"" then FOneItem.FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FOneItem.Fitemgubun + "/" + GetImageSubFolderByItemid(FOneItem.Fshopitemid) + "/" + FOneItem.FOffimgSmall
            ''마진 설정이 안되어 있으면 온라인 마진으로
            if IsNULL(FOneItem.FMakerMargin) then
                FOneItem.FMakerMargin = CLng((rsget("sellcash")-rsget("buycash"))/rsget("sellcash")*100)
            end if

            FOneItem.FOnlineSellcash	= rsget("sellcash")
			FOneItem.FOnlineBuycash		= rsget("buycash")
            FOneItem.FOnlineOptaddprice = rsget("optaddprice")
			FOneItem.FOnlineOptaddbuyprice = rsget("optaddbuyprice")
            FOneItem.Fmwdiv			= rsget("mwdiv")
            FOneItem.Fcomm_cd       = rsget("comm_cd")
			FOneItem.Fforeign_sellprice	 = rsget("orgprice")
			FOneItem.Fforeign_suplyprice = rsget("suplyprice")

		end if
		rsget.Close
	end Sub

	'//admin/fran/popoffjumunbycsv.asp
	public sub GetOffLineJumunByArr()
		dim sqlStr,i

		'/2016.09.13 한용민 추가
		sqlStr = "select top 1 " & vbcrlf
		sqlStr = sqlStr & " u.countrylangcd, u.currencyunit, u.loginsite" & vbcrlf
		sqlStr = sqlStr & " , e.currencyChar, e.exchangeRate, e.multiplerate, e.linkPriceType, e.basedate, e.regdate" & vbcrlf
		sqlStr = sqlStr & " , e.lastupdate, e.reguserid, e.lastuserid" & vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & " join db_item.dbo.tbl_exchangeRate e" & vbcrlf
		sqlStr = sqlStr & " 	on u.countrylangcd = e.countrylangcd" & vbcrlf
		sqlStr = sqlStr & " 	and u.currencyunit = e.currencyunit" & vbcrlf
		'sqlStr = sqlStr & " 	and u.loginsite = e.sitename" & vbcrlf
		sqlStr = sqlStr & " 	and e.sitename='WSLWEB'" & vbcrlf
		sqlStr = sqlStr & " where userid = '" + FRectShopid + "'" & vbcrlf

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
		if  not rsget.EOF  then
    		Fcurrencyunit = rsget("currencyunit")
    		Floginsite = rsget("loginsite")
    		fcountrylangcd = rsget("countrylangcd")
    		fcurrencyChar = rsget("currencyChar")
			flinkPriceType = rsget("linkPriceType")
			fmultiplerate = rsget("multiplerate")
    	end if
    	rsget.Close

		sqlStr = " select top 200 "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption, "
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname,"
		sqlStr = sqlStr + " s.shopitemprice, s.shopsuplycash, s.shopbuyprice, "
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, i.sellcash , i.buycash,"
		sqlStr = sqlStr + " d.defaultmargin, d.defaultsuplymargin, d.comm_cd "
		sqlStr = sqlStr + " ,i.smallimage"
		sqlStr = sqlStr + " ,IsNULL(o.isusing,'Y') as optusing"
		sqlStr = sqlStr + " ,isNULL(i.mwdiv,'') as mwdiv"
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
        sqlStr = sqlStr + " ,IsNULL(s.extbarcode,'') as extbarcode,isNULL(s.tnbarcode,'') as tnbarcode"
        sqlStr = sqlStr + " ,s.isusing"

		'/홀쎄일의 경우 대표화폐가 한화 일경우 초기값 셋팅
		if Floginsite="WSLWEB" then
			if (fcurrencyUnit = "KRW" or fcurrencyUnit = "WON") Then
				'해외판매가가 입력이 안되었을경우, 국내가격을 해외 가격에 넣는다.
				if flinkPriceType="1" then
					sqlStr = sqlStr & " ,(case when isNull(mlp.orgprice,0)=0 then isNull(s.shopitemprice,0) else isNull(mlp.orgprice,0) end) as orgprice"
				else
					sqlStr = sqlStr & " ,(case when isNull(mlp.orgprice,0)=0 then isNull(s.orgsellprice,0) else isNull(mlp.orgprice,0) end) as orgprice"
				end if
				sqlStr = sqlStr & " ,(case"
				sqlStr = sqlStr & " 	when isNull(mlp.orgprice,0)=0 then isNull(s.shopsuplycash,0)"
				sqlStr = sqlStr & " 	else round( (isnull(mlp.orgprice,0)*(100-IsNULL(d.defaultsuplymargin,100))/100) ,1) end) as suplyprice"
			else
				sqlStr = sqlStr & " ,isNull(mlp.orgprice,0) as orgprice"
				sqlStr = sqlStr & " ,round( (isnull(mlp.orgprice,0)*(100-IsNULL(d.defaultsuplymargin,100))/100) ,1) as suplyprice"
			end if
		else
			sqlStr = sqlStr & " ,0 as orgprice"
			sqlStr = sqlStr & " ,0 as suplyprice"
		end if

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + "     JOin [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + "     on d.shopid='" + FRectShopid + "'" ''필수.
		sqlStr = sqlStr + "     and d.makerid=s.makerid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on s.itemgubun='10' and s.shopitemid=i.itemid and i.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
	    sqlStr = sqlStr & " left outer join [db_item].dbo.tbl_item_multiLang_price as mlp " & VbCrLf
	    sqlStr = sqlStr & "		on mlp.sitename ='"&Floginsite&"' and mlp.currencyunit = '"&Fcurrencyunit&"' and mlp.itemid = s.shopitemid "
		sqlStr = sqlStr + " where 1=1"
		if (FRectDesigner<>"") then  ''2017/05/23
		    sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		    sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
	    end if
        sqlStr = sqlStr + " and d.comm_cd in ('B013','B031')"
		sqlStr = sqlStr + " and (s.tnbarcode "  ''2017/05/23 [tnbarcode]
		sqlStr = sqlStr + "     in ("
		sqlStr = sqlStr + "     " + FRectBarCodeArr &vbCRLF
		sqlStr = sqlStr + "     )"
		sqlStr = sqlStr + " or s.extbarcode "  ''2017/05/23 [extbarcode]
		sqlStr = sqlStr + "     in ("
		sqlStr = sqlStr + "     " + FRectBarCodeArr &vbCRLF
		sqlStr = sqlStr + "     )"
		sqlStr = sqlStr + " ) "
		sqlStr = sqlStr + " order by s.makerid, s.itemgubun, s.shopitemid, s.itemoption"

		''response.write sqlStr
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
				set FItemList(i) = new COffShopOneItem

				FItemList(i).Fforeign_sellprice	 = rsget("orgprice")
				FItemList(i).Fforeign_suplyprice = rsget("suplyprice")
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")

				FItemList(i).FimageSmall     = rsget("smallimage")
				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

                FItemList(i).Fsellyn               = rsget("sellyn")
				FItemList(i).Flimityn              = rsget("limityn")
				FItemList(i).Flimitno              = rsget("limitno")
				FItemList(i).Flimitsold            = rsget("limitsold")
				FItemList(i).FMakerMargin    = rsget("defaultmargin")
				FItemList(i).FShopMargin    = rsget("defaultsuplymargin")
				FItemList(i).Foptusing      = rsget("optusing")
				FItemList(i).Fshopbuyprice  = rsget("shopbuyprice")

				FItemList(i).FOnlineSellcash	= rsget("sellcash")
    			FItemList(i).Fonlinebuycash     = rsget("buycash")
                FItemList(i).FOnlineOptaddprice = rsget("optaddprice")
    			FItemList(i).FOnlineOptaddbuyprice = rsget("optaddbuyprice")
                FItemList(i).Fmwdiv			    = rsget("mwdiv")
                FItemList(i).Fcomm_cd       = rsget("comm_cd")

                FItemList(i).Ftnbarcode     = rsget("tnbarcode")
                FItemList(i).Fextbarcode    = rsget("extbarcode")
                FItemList(i).Fisusing       = rsget("isusing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub

	' 팝업 오프라인 상품 리스트
	'//common/offshop/localeitem/popshopjumunitem_locale.asp
    public sub GetOffLineJumunItemWithStock_locale()
		dim sqlStr,i , sqlsearch , iStartDate

        if (FRectOrder="byrecent") then
            sqlStr = "select convert(varchar(10),dateadd(d,-14,getdate()),21) as stdt "
            rsget.Open sqlStr,dbget,1
    		iStartDate = rsget("stdt")
    		rsget.Close
        end if

		'/2016.09.13 한용민 추가
		sqlStr = "select top 1 " & vbcrlf
		sqlStr = sqlStr & " u.countrylangcd, u.currencyunit, u.loginsite" & vbcrlf
		sqlStr = sqlStr & " , e.currencyChar, e.exchangeRate, e.multiplerate, e.linkPriceType, e.basedate, e.regdate" & vbcrlf
		sqlStr = sqlStr & " , e.lastupdate, e.reguserid, e.lastuserid" & vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & " join db_item.dbo.tbl_exchangeRate e" & vbcrlf
		sqlStr = sqlStr & " 	on u.countrylangcd = e.countrylangcd" & vbcrlf
		sqlStr = sqlStr & " 	and u.currencyunit = e.currencyunit" & vbcrlf
		'sqlStr = sqlStr & " 	and u.loginsite = e.sitename" & vbcrlf
		sqlStr = sqlStr & " 	and e.sitename='WSLWEB'" & vbcrlf
		sqlStr = sqlStr & " where userid = '" + FRectShopid + "'" & vbcrlf

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
		if  not rsget.EOF  then
    		Fcurrencyunit = rsget("currencyunit")
    		Floginsite = rsget("loginsite")
    		fcountrylangcd = rsget("countrylangcd")
    		fcurrencyChar = rsget("currencyChar")
			flinkPriceType = rsget("linkPriceType")
			fmultiplerate = rsget("multiplerate")
    	end if
    	rsget.Close

		if FRectOrder="byetc" then
			sqlsearch = sqlsearch + " and s.itemgubun ='70'"
		elseif FRectOrder="byevent" then
			sqlsearch = sqlsearch + " and s.itemgubun ='80'"
		elseif FRectOrder="by7sell" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    if FRectDesigner<>"" then
		        sqlsearch = sqlsearch + " and st.sell7days>0"
		    else
		        sqlsearch = sqlsearch + " and st.sell7days>1"
		    end if
		elseif FRectOrder="byrecent" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and s.regdate>'" & iStartDate & "'"
		elseif FRectOrder="byonbest" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and i.itemscore>0"
		elseif FRectOrder="byoffbest" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and st.sell7days>0"
		elseif (FRectOrder="byoffbestAll") then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and st.sellcnt>0"
		else
			sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		end if

		if (FRectItemid<>"") then
		    sqlsearch = sqlsearch + " and s.shopitemid=" + CStr(FRectItemid) + ""
		end if

		if (FRectItemName<>"") then
		    sqlsearch = sqlsearch + " and s.shopitemname like '%" + CStr(FRectItemName) + "%'"
		end if

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and s.makerid='" + FRectDesigner + "'"
		end if

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and d.shopid='" + FRectShopid + "'"
		end if

		if FRectDesignerjungsangubun<>"" then
			sqlsearch = sqlsearch + " and d.chargediv in (" + FRectDesignerjungsangubun + ")"
		end if

		if FRectcomm_cd <> "" then
			sqlsearch = sqlsearch + " and d.comm_cd in ("&FRectcomm_cd&")"
		end if

        if (FRectOnlyActive<>"") then
		    sqlsearch = sqlsearch + " and (IsNULL(i.sellyn,'')<>'N') and s.isusing='Y'"
		end if

        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and s.catecdl='" + FRectCDL + "'"
        end if

        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and s.catecdm='" + FRectCDM + "'"
        end if

        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and s.catecdn='" + FRectCDS + "'"
        end if

        if (frectisusing<>"") then
            sqlsearch = sqlsearch + " and s.isusing='" + frectisusing + "'"
        end if

		sqlStr = " select count(s.shopitemid) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " on d.makerid=s.makerid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " on s.itemgubun='10' and s.shopitemid=i.itemid"

		if (FRectOrder="byoffbestAll") then
		    sqlStr = sqlStr + " left join db_const.dbo.tbl_const_offshop st"
            sqlStr = sqlStr + " on s.itemgubun=st.itemgubun and s.shopitemid=st.shopitemid and s.itemoption=st.itemoption"
		else
		    sqlStr = sqlStr + " left join db_summary.dbo.tbl_current_shopstock_summary st"
		    sqlStr = sqlStr + " on st.shopid='" + FRectShopid + "' and s.itemgubun=st.itemgubun and s.shopitemid=st.itemid and s.itemoption=st.itemoption"
	    end if
 		'해외 상품금액
	    if Fcurrencyunit <> "WON" and Fcurrencyunit <> "KRW" THEN
	    sqlStr = sqlStr + " left outer join dbo.tbl_item_multiLang_price as mlp " & VbCrLf
	    sqlStr = sqlStr + "		on mlp.sitename ='"&Floginsite&"' and mlp.currencyunit = '"&Fcurrencyunit&"' and mlp.itemid = i.itemid "
		end if

		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount>=3000 then FTotalCount=3000

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption, "
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname,"
		sqlStr = sqlStr + " s.shopitemprice, s.shopsuplycash, s.shopbuyprice, i.smallimage,i.listimage ,"
		sqlStr = sqlStr + " IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist, IsNULL(s.offimgsmall,'') as offimgsmall,"
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, IsNULL(i.buycash,0) as onlinebuycash, i.deliverytype, "
		sqlStr = sqlStr + " d.defaultmargin,d.defaultsuplymargin, "
		sqlStr = sqlStr + " IsNULL(o.isusing,'Y') as optusing, "
		sqlStr = sqlStr + " IsNULL(o.optlimitno,0) as optlimitno, IsNULL(o.optlimitsold,0) as optlimitsold, IsNULL(o.optsellyn,'Y') optsellyn,"

		if (FRectOrder="byoffbestAll") then
		    sqlStr = sqlStr + " IsNULL(st.sellcnt,0) as sell7days"
		else
		    sqlStr = sqlStr + " IsNULL(st.sell7days,0) as sell7days , IsNULL(st.sell3days,0) as sell3days"
		end if

		'/홀쎄일의 경우 대표화폐가 한화 일경우 초기값 셋팅
		if Floginsite="WSLWEB" then
			if (fcurrencyUnit = "KRW" or fcurrencyUnit = "WON") Then
				'해외판매가가 입력이 안되었을경우, 국내가격을 해외 가격에 넣는다.
				if flinkPriceType="1" then
					sqlStr = sqlStr & " ,(case when isNull(mlp.orgprice,0)=0 then isNull(s.shopitemprice,0) else isNull(mlp.orgprice,0) end) as orgprice"
				else
					sqlStr = sqlStr & " ,(case when isNull(mlp.orgprice,0)=0 then isNull(s.orgsellprice,0) else isNull(mlp.orgprice,0) end) as orgprice"
				end if
				sqlStr = sqlStr & " ,(case"
				sqlStr = sqlStr & " 	when isNull(mlp.orgprice,0)=0 then isNull(s.shopsuplycash,0)"
				sqlStr = sqlStr & " 	else round( (isnull(mlp.orgprice,0)*(100-IsNULL(d.defaultsuplymargin,100))/100) ,1) end) as suplyprice"
			else
				sqlStr = sqlStr & " ,isNull(mlp.orgprice,0) as orgprice"
				sqlStr = sqlStr & " ,round( (isnull(mlp.orgprice,0)*(100-IsNULL(d.defaultsuplymargin,100))/100) ,1) as suplyprice"
			end if
		else
			sqlStr = sqlStr & " ,0 as orgprice"
			sqlStr = sqlStr & " ,0 as suplyprice"
		end if

		''옵션 추가금액
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
		sqlStr = sqlStr + " ,IsNULL(i.mwdiv,'') as mwdiv"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " on d.makerid=s.makerid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o "
		sqlStr = sqlStr + " on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"

		if (FRectOrder="byoffbestAll") then
		    sqlStr = sqlStr + " left join db_const.dbo.tbl_const_offshop st"
            sqlStr = sqlStr + " on s.itemgubun=st.itemgubun and s.shopitemid=st.shopitemid and s.itemoption=st.itemoption"
		else
		    sqlStr = sqlStr + " left join db_summary.dbo.tbl_current_shopstock_summary st"
		    sqlStr = sqlStr + " on st.shopid='" + FRectShopid + "' and s.itemgubun=st.itemgubun and s.shopitemid=st.itemid and s.itemoption=st.itemoption"
		end if

		 '해외 상품금액
	    sqlStr = sqlStr + " left outer join dbo.tbl_item_multiLang_price as mlp " & VbCrLf
	    sqlStr = sqlStr + "		on mlp.sitename ='"&Floginsite&"' and mlp.currencyunit = '"&Fcurrencyunit&"' and mlp.itemid = i.itemid "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

        if FRectOrder="by7sell" then
            sqlStr = sqlStr + " order by st.sell7days desc,s.itemgubun asc, s.shopitemid desc, s.itemoption"
		elseif FRectOrder="byrecent" then
		    sqlStr = sqlStr + " order by s.regdate desc"
		elseif FRectOrder="byonbest" then
		    sqlStr = sqlStr + " order by i.itemscore desc,s.itemgubun asc, s.shopitemid desc, s.itemoption"
		elseif (FRectOrder="byoffbest") or (FRectOrder="byoffbestAll") then
		    sqlStr = sqlStr + " order by st.sell7days desc,s.itemgubun asc, s.shopitemid desc, s.itemoption"
		else
		    sqlStr = sqlStr + " order by s.itemgubun asc, s.shopitemid desc, s.itemoption"
		end if

		'response.write sqlStr &"<br>"
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
				set FItemList(i) = new COffShopOneItem

				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).FimageSmall     = rsget("smallimage")

				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if
				FItemList(i).FimageList     = rsget("listimage")
				if FItemList(i).FimageList<>"" then
					FItemList(i).FimageList     = webImgUrl + "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageList
				end if

				FItemList(i).Fsellyn               = rsget("sellyn")
                FItemList(i).Flimityn              = rsget("limityn")

                if FItemList(i).Fitemoption="0000" then
				    FItemList(i).Flimitno              = rsget("limitno")
				    FItemList(i).Flimitsold            = rsget("limitsold")
				else
				    if rsget("optsellyn")="N" then FItemList(i).Fsellyn="N"
				    FItemList(i).Flimitno              = rsget("optlimitno")
				    FItemList(i).Flimitsold            = rsget("optlimitsold")
				end if

				FItemList(i).FMakerMargin    = rsget("defaultmargin")
				FItemList(i).FShopMargin    = rsget("defaultsuplymargin")
				FItemList(i).Foptusing      = rsget("optusing")
				FItemList(i).Fonlinebuycash = rsget("onlinebuycash")
                ''옵션 추가금액
			    FItemList(i).FOnlineOptaddprice = rsget("optaddprice")
			    FItemList(i).FOnlineOptaddbuyprice = rsget("optaddbuyprice")
				FItemList(i).FOffimgMain	= rsget("offimgmain")
				FItemList(i).FOffimgList	= rsget("offimglist")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				FItemList(i).FdeliveryType  = rsget("deliverytype")
				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
				FItemList(i).Fshopbuyprice  = rsget("shopbuyprice")
				FItemList(i).FOffsell7days     = rsget("sell7days")
				if (FRectOrder<>"byoffbestAll") then FItemList(i).fsell3days     = rsget("sell3days")
                FItemList(i).Fmwdiv			= rsget("mwdiv")
				FItemList(i).Fforeign_sellprice	 = rsget("orgprice")
				FItemList(i).Fforeign_suplyprice = rsget("suplyprice")

				'/홀쎄일의 경우 대표화폐가 한화 일경우 초기값 셋팅
				if Floginsite="WSLWEB" then
					if (fcurrencyUnit = "KRW" or fcurrencyUnit = "WON") Then
						' 홀쎄일인데 국내에서 도매로 사용하는 케이스가 있음.
						FItemList(i).Fshopitemprice     = FItemList(i).Fforeign_sellprice
						FItemList(i).Fshopsuplycash     = FItemList(i).Fforeign_suplyprice
					end if
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end sub

	' 팝업 오프라인 상품 리스트 '//common/offshop/popshopjumunitem.asp
    public sub GetOffLineJumunItemWithStock()
		dim sqlStr,i , sqlsearch , iStartDate

        if (FRectOrder="byrecent") then
            sqlStr = "select convert(varchar(10),dateadd(d,-14,getdate()),21) as stdt "
            rsget.Open sqlStr,dbget,1
    		iStartDate = rsget("stdt")
    		rsget.Close
        end if

		'/2016.09.13 한용민 추가
		sqlStr = "select top 1 " & vbcrlf
		sqlStr = sqlStr & " u.countrylangcd, u.currencyunit, u.loginsite" & vbcrlf
		sqlStr = sqlStr & " , e.currencyChar, e.exchangeRate, e.multiplerate, e.linkPriceType, e.basedate, e.regdate" & vbcrlf
		sqlStr = sqlStr & " , e.lastupdate, e.reguserid, e.lastuserid" & vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_user u with (nolock)" & vbcrlf
		sqlStr = sqlStr & " join db_item.dbo.tbl_exchangeRate e with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on u.countrylangcd = e.countrylangcd" & vbcrlf
		sqlStr = sqlStr & " 	and u.currencyunit = e.currencyunit" & vbcrlf
		'sqlStr = sqlStr & " 	and u.loginsite = e.sitename" & vbcrlf
		sqlStr = sqlStr & " 	and e.sitename='WSLWEB'" & vbcrlf
		sqlStr = sqlStr & " where userid = '" + FRectShopid + "'" & vbcrlf

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
		if  not rsget.EOF  then
    		Fcurrencyunit = rsget("currencyunit")
    		Floginsite = rsget("loginsite")
    		fcountrylangcd = rsget("countrylangcd")
    		fcurrencyChar = rsget("currencyChar")
			flinkPriceType = rsget("linkPriceType")
			fmultiplerate = rsget("multiplerate")
    	end if
    	rsget.Close

		if (FRectOrder="all") then
		elseif FRectOrder="byetc" then
			sqlsearch = sqlsearch + " and s.itemgubun ='70'"
		elseif FRectOrder="byevent" then
			sqlsearch = sqlsearch + " and s.itemgubun ='80'"
		elseif FRectOrder="by7sell" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    if FRectDesigner<>"" then
		        sqlsearch = sqlsearch + " and stc.sell7days>0"
		    else
		        sqlsearch = sqlsearch + " and stc.sell7days>1"
		    end if
		elseif FRectOrder="byrecent" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and s.regdate>'" & iStartDate & "'"
		elseif FRectOrder="byonbest" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and i.itemscore>0"
		elseif FRectOrder="byoffbest" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and stc.sell7days>0"
		elseif (FRectOrder="byoffbestAll") then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and st.sellcnt>0"
		else
			sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		end If

		'// 55 상품 제외, skyer9, 2017-06-13
		sqlsearch = sqlsearch + " 	and s.itemgubun <>'55' "

		if (FRectItemid<>"") then
		    sqlsearch = sqlsearch + " and s.shopitemid in ("& FRectItemID&")"
		end if

		if (FRectItemName<>"") then
		    sqlsearch = sqlsearch + " and s.shopitemname like '%" + CStr(FRectItemName) + "%'"
		end if

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and s.makerid='" + FRectDesigner + "'"
		end if

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and d.shopid='" + FRectShopid + "'"
		end if

		if FRectDesignerjungsangubun<>"" then
			sqlsearch = sqlsearch + " and d.chargediv in (" + FRectDesignerjungsangubun + ")"
		end if

		if FRectcomm_cd <> "" then
			sqlsearch = sqlsearch + " and d.comm_cd in ("&FRectcomm_cd&")"
		end if

        if (FRectOnlyActive<>"") then
		    sqlsearch = sqlsearch + " and (IsNULL(i.sellyn,'')<>'N') and s.isusing='Y'"
		end if

		if FRectPrdCode<>"" then
		    if (Len(FRectPrdCode) = 12) then
				sqlsearch = sqlsearch + " 	and s.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			else
				sqlsearch = sqlsearch + " 	and s.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			end if
		end if

		if FRectGeneralBarcode<>"" then
			sqlsearch = sqlsearch + " 	and ba.barcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if

        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and s.catecdl='" + FRectCDL + "'"
        end if

        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and s.catecdm='" + FRectCDM + "'"
        end if

        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and s.catecdn='" + FRectCDS + "'"
        end if

        if (frectisusing<>"") then
            sqlsearch = sqlsearch + " and s.isusing='" + frectisusing + "'"
        end if

        ''입고된 내역만..
		if FRectIpGoOnly = "on" then
			sqlsearch = sqlsearch + " and (stc.logicsipgono+stc.logicsreipgono > 0 or stc.brandipgono+stc.brandreipgono > 0)"
		end if

		'/최근7일 판매건만
		if FRectSell7days = "on" then
			sqlsearch = sqlsearch + " and stc.sell7days > 0"
		end if

		'기주문포함부족상품
		if FRectIncludePreOrder = "on" then
	        if FRectShortageType="3" then
	    		sqlsearch = sqlsearch + " and (stc.sell3days*1) - (db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno)+stc.preordernofix) > 0"
	    	elseif FRectShortageType="7" then
	    		sqlsearch = sqlsearch + " and (stc.sell7days*1) - (db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno)+stc.preordernofix) > 0"
	    	elseif FRectShortageType="14" then
	    		sqlsearch = sqlsearch + " and (stc.sell7days*2) - (db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno)+stc.preordernofix) > 0"
	    	elseif FRectShortageType="28" then
	    		sqlsearch = sqlsearch + " and (stc.sell7days*4) - (db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno)+stc.preordernofix) > 0"
	    	else
	    		sqlsearch = sqlsearch + " and stc.preordernofix + db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno) < 1"
	    		'sqlsearch = sqlsearch + " and (stc.sell7days*1) - (db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno)+stc.preordernofix) > 0"
	    	end if
		else
	        if FRectShortageType="3" then
	    		sqlsearch = sqlsearch + " and (stc.sell3days*1) - db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno) > 0"
	    	elseif FRectShortageType="7" then
	    		sqlsearch = sqlsearch + " and (stc.sell7days*1) - db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno) > 0"
	    	elseif FRectShortageType="14" then
	    		sqlsearch = sqlsearch + " and (stc.sell7days*2) - db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno) > 0"
	    	elseif FRectShortageType="28" then
	    		sqlsearch = sqlsearch + " and (stc.sell7days*4) - db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno) > 0"
	    	else

	    	end if
		end If

		if FRectOnlineMWdiv <> "" then
			sqlsearch = sqlsearch + " and i.mwdiv = '" & FRectOnlineMWdiv & "' "
		end if
        if (FRectSellYN="YS") then
            sqlsearch = sqlsearch & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            sqlsearch = sqlsearch & " and i.sellyn='" + FRectSellYN + "'"
        end if

		sqlStr = " select count(s.shopitemid) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d with (nolock)"
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " 	on d.makerid=s.makerid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + "     on s.itemgubun='10' and s.shopitemid=i.itemid"

		if FRectGeneralBarcode<>"" then
			'범용바코드 검색
			sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_stock ba with (nolock)" & VbCrLf
			sqlStr = sqlStr + " 	on " & VbCrLf
			sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
			sqlStr = sqlStr + " 		and s.itemgubun = ba.itemgubun " & VbCrLf
			sqlStr = sqlStr + " 		and s.shopitemid = ba.itemid " & VbCrLf
			sqlStr = sqlStr + " 		and s.itemoption = ba.itemoption " & VbCrLf
		end If

		If FRectCopyIdx <> "" Then
			sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail dd with (nolock)" & VbCrLf
			sqlStr = sqlStr + " 	on " & VbCrLf
			sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
			sqlStr = sqlStr + " 		and dd.masteridx = " & FRectCopyIdx & VbCrLf
			sqlStr = sqlStr + " 		and s.itemgubun = dd.itemgubun " & VbCrLf
			sqlStr = sqlStr + " 		and s.shopitemid = dd.itemid " & VbCrLf
			sqlStr = sqlStr + " 		and s.itemoption = dd.itemoption " & VbCrLf
		End If

		'카테고리 검색
		''sqlStr = sqlStr + " left join [db_item].[dbo].vw_category c on"
		''sqlStr = sqlStr + "		s.catecdl=c.cdlarge and s.catecdm=c.cdmid and s.catecdn=c.cdsmall"

		'물류센터 재고
        if FRectLogicsIpGoOnly = "on" then
		    sqlStr = sqlStr + " join db_summary.dbo.tbl_current_logisstock_summary lc with (nolock)"
    	    sqlStr = sqlStr + "     on s.itemgubun=lc.itemgubun and s.shopitemid=lc.itemid and s.itemoption=lc.itemoption"
        end if
		'매장 재고
		if FRectIpGoOnly = "on" then
		    sqlStr = sqlStr + " join db_summary.dbo.tbl_current_shopstock_summary stc with (nolock)"
    	    sqlStr = sqlStr + "     on stc.shopid='" & FRectShopid & "' and s.itemgubun=stc.itemgubun and s.shopitemid=stc.itemid and s.itemoption=stc.itemoption"
		else
    	    sqlStr = sqlStr + " left join db_summary.dbo.tbl_current_shopstock_summary stc with (nolock)"
    	    sqlStr = sqlStr + "     on stc.shopid='" & FRectShopid & "' and s.itemgubun=stc.itemgubun and s.shopitemid=stc.itemid and s.itemoption=stc.itemoption"
        end if

		if (FRectOrder="byoffbestAll") then
		    sqlStr = sqlStr + " left join db_const.dbo.tbl_const_offshop st with (nolock)"
            sqlStr = sqlStr + "     on s.itemgubun=st.itemgubun and s.shopitemid=st.shopitemid and s.itemoption=st.itemoption"
	    end if

	    '해외 상품금액
	    sqlStr = sqlStr + " left outer join dbo.tbl_item_multiLang_price as mlp with (nolock)" & VbCrLf
	    sqlStr = sqlStr + "		on mlp.sitename ='"&Floginsite&"' and mlp.currencyunit = '"&Fcurrencyunit&"' and mlp.itemid = i.itemid "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		''response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount>=3000 then FTotalCount=3000

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption, "
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname,"
		sqlStr = sqlStr + " s.shopitemprice, s.shopsuplycash, s.shopbuyprice, i.smallimage,"
		sqlStr = sqlStr + " IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist, IsNULL(s.offimgsmall,'') as offimgsmall,"
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, IsNULL(i.buycash,0) as onlinebuycash, i.deliverytype, "
		sqlStr = sqlStr + " d.defaultmargin,d.defaultsuplymargin, d.comm_cd,"
		sqlStr = sqlStr + " IsNULL(o.isusing,'Y') as optusing, "
		sqlStr = sqlStr + " IsNULL(o.optlimitno,0) as optlimitno, IsNULL(o.optlimitsold,0) as optlimitsold, IsNULL(o.optsellyn,'Y') optsellyn,"

		'// 매장재고
		sqlStr = sqlStr & " stc.shopid, isnull(stc.logicsipgono,0) as logicsipgono, isnull(stc.logicsreipgono,0) as logicsreipgono"
		sqlStr = sqlStr & " , isnull(stc.brandipgono,0) as brandipgono, isnull(stc.brandreipgono,0) as brandreipgono"
		sqlStr = sqlStr & " , isnull(stc.sellno,0) as sellno, isnull(stc.resellno,0) as resellno,isnull(stc.errsampleitemno,0) as errsampleitemno"
		sqlStr = sqlStr & " , isnull(stc.errbaditemno,0) as errbaditemno, isnull(stc.errrealcheckno,0) as errrealcheckno"
		sqlStr = sqlStr & " , isnull(stc.sysstockno,0) as sysstockno, isnull(stc.realstockno,0) as realstockno, isnull(stc.requiredStock,0) as requiredStock"
		sqlStr = sqlStr & " , isnull(stc.sell7days,0) as sell7days, isnull(stc.sell3days,0) as sell3days ,stc.lastupdate"
		sqlStr = sqlStr & " , stc.preorderno ,stc.preordernofix"

		if FRectIncludePreOrder = "on" Then
			sqlStr = sqlStr & " ,( (stc.sell3days*1) - (db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno)+stc.preordernofix) ) as require3daystock"
			sqlStr = sqlStr & " ,( (stc.sell7days*1) - (db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno)+stc.preordernofix) ) as require7daystock"
			sqlStr = sqlStr & " ,( (stc.sell7days*2) - (db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno)+stc.preordernofix) ) as require14daystock"
		Else
			sqlStr = sqlStr & " ,( (stc.sell3days*1) - (db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno)) ) as require3daystock"
			sqlStr = sqlStr & " ,( (stc.sell7days*1) - (db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno)) ) as require7daystock"
			sqlStr = sqlStr & " ,( (stc.sell7days*2) - (db_summary.[dbo].[uf_replacezero](stc.realstockNo+stc.errsampleitemno+stc.errbaditemno)) ) as require14daystock"
		End If

		if (FRectOrder="byoffbestAll") then
			sqlStr = sqlStr & " , IsNULL(st.sellcnt,0) as sell7daysoffALL "
		else
			sqlStr = sqlStr & " , 0 as sell7daysoffALL "
		end if

		'/홀쎄일의 경우 대표화폐가 한화 일경우 초기값 셋팅
		if Floginsite="WSLWEB" then
			if (fcurrencyUnit = "KRW" or fcurrencyUnit = "WON") Then
				'해외판매가가 입력이 안되었을경우, 국내가격을 해외 가격에 넣는다.
				if flinkPriceType="1" then
					sqlStr = sqlStr & " ,(case when isNull(mlp.orgprice,0)=0 then isNull(s.shopitemprice,0) else isNull(mlp.orgprice,0) end) as orgprice"
				else
					sqlStr = sqlStr & " ,(case when isNull(mlp.orgprice,0)=0 then isNull(s.orgsellprice,0) else isNull(mlp.orgprice,0) end) as orgprice"
				end if
				sqlStr = sqlStr & " ,(case"
				sqlStr = sqlStr & " 	when isNull(mlp.orgprice,0)=0 then isNull(s.shopsuplycash,0)"
				sqlStr = sqlStr & " 	else round( (isnull(mlp.orgprice,0)*(100-IsNULL(d.defaultsuplymargin,100))/100) ,1) end) as suplyprice"
			else
				sqlStr = sqlStr & " ,isNull(mlp.orgprice,0) as orgprice"
				sqlStr = sqlStr & " ,round( (isnull(mlp.orgprice,0)*(100-IsNULL(d.defaultsuplymargin,100))/100) ,1) as suplyprice"
			end if
		else
			sqlStr = sqlStr & " ,0 as orgprice"
			sqlStr = sqlStr & " ,0 as suplyprice"
		end if

		''옵션 추가금액
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
		sqlStr = sqlStr + " ,IsNULL(i.mwdiv,'') as mwdiv, s.centermwdiv "

		If FRectCopyIdx <> "" Then
			sqlStr = sqlStr + " ,dd.baljuitemno as itemno "
		Else
			sqlStr = sqlStr + " ,0 as itemno "
		End If
		sqlStr = sqlStr + " , s.orgsellprice "
		sqlStr = sqlStr + " , lc.realstock as logicsRealStock "

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d with (nolock)"
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " 	on d.makerid=s.makerid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + "     on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o with (nolock)"
		sqlStr = sqlStr + "     on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"

		if FRectGeneralBarcode<>"" then
			'범용바코드 검색
			sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_stock ba with (nolock)" & VbCrLf
			sqlStr = sqlStr + " 	on " & VbCrLf
			sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
			sqlStr = sqlStr + " 		and s.itemgubun = ba.itemgubun " & VbCrLf
			sqlStr = sqlStr + " 		and s.shopitemid = ba.itemid " & VbCrLf
			sqlStr = sqlStr + " 		and s.itemoption = ba.itemoption " & VbCrLf
		end If

		If FRectCopyIdx <> "" Then
			sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail dd with (nolock)" & VbCrLf
			sqlStr = sqlStr + " 	on " & VbCrLf
			sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
			sqlStr = sqlStr + " 		and dd.masteridx = " & FRectCopyIdx & VbCrLf
			sqlStr = sqlStr + " 		and s.itemgubun = dd.itemgubun " & VbCrLf
			sqlStr = sqlStr + " 		and s.shopitemid = dd.itemid " & VbCrLf
			sqlStr = sqlStr + " 		and s.itemoption = dd.itemoption " & VbCrLf
		End If

		'카테고리 검색
		''sqlStr = sqlStr + " left join [db_item].[dbo].vw_category c on"
		''sqlStr = sqlStr + "		s.catecdl=c.cdlarge and s.catecdm=c.cdmid and s.catecdn=c.cdsmall"

		'물류센터 재고
        if FRectLogicsIpGoOnly = "on" then
		    sqlStr = sqlStr + " join db_summary.dbo.tbl_current_logisstock_summary lc with (nolock)"
    	    sqlStr = sqlStr + "     on s.itemgubun=lc.itemgubun and s.shopitemid=lc.itemid and s.itemoption=lc.itemoption"
		else
		    sqlStr = sqlStr + " left join db_summary.dbo.tbl_current_logisstock_summary lc with (nolock)"
    	    sqlStr = sqlStr + "     on s.itemgubun=lc.itemgubun and s.shopitemid=lc.itemid and s.itemoption=lc.itemoption"
        end if
		'매장 재고
		if FRectIpGoOnly = "on" then
		    sqlStr = sqlStr + " join db_summary.dbo.tbl_current_shopstock_summary stc with (nolock)"
	        sqlStr = sqlStr + "     on stc.shopid='" + FRectShopid + "' and s.itemgubun=stc.itemgubun and s.shopitemid=stc.itemid and s.itemoption=stc.itemoption"
		else
	        sqlStr = sqlStr + " left join db_summary.dbo.tbl_current_shopstock_summary stc with (nolock)"
	        sqlStr = sqlStr + "     on stc.shopid='" + FRectShopid + "' and s.itemgubun=stc.itemgubun and s.shopitemid=stc.itemid and s.itemoption=stc.itemoption"
        end if

		if (FRectOrder="byoffbestAll") then
		    sqlStr = sqlStr + " left join db_const.dbo.tbl_const_offshop st with (nolock)"
            sqlStr = sqlStr + "     on s.itemgubun=st.itemgubun and s.shopitemid=st.shopitemid and s.itemoption=st.itemoption"
		end if

		'해외 상품금액
	    sqlStr = sqlStr + " left outer join [db_item].dbo.tbl_item_multiLang_price as mlp with (nolock)" & VbCrLf
	    sqlStr = sqlStr + "		on mlp.sitename ='"&Floginsite&"' and mlp.currencyunit = '"&Fcurrencyunit&"' and mlp.itemid = i.itemid "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

        if FRectOrder="by7sell" then
            sqlStr = sqlStr + " order by stc.sell7days desc,s.itemgubun asc, s.shopitemid desc, s.itemoption"
		elseif FRectOrder="byrecent" then
		    sqlStr = sqlStr + " order by s.regdate desc"
		elseif FRectOrder="byonbest" then
		    sqlStr = sqlStr + " order by i.itemscore desc,s.itemgubun asc, s.shopitemid desc, s.itemoption"
		elseif (FRectOrder="byoffbest") or (FRectOrder="byoffbestAll") then
		    sqlStr = sqlStr + " order by stc.sell7days desc,s.itemgubun asc, s.shopitemid desc, s.itemoption"
		elseif FRectOrder="BI" Then
		   sqlStr = sqlStr + " order by s.makerid, s.itemgubun asc, s.shopitemid desc, s.itemoption"
		elseif FRectOrder="I" Then
		   sqlStr = sqlStr + " order by s.itemgubun asc, s.shopitemid desc, s.itemoption"
		else
		    sqlStr = sqlStr + " order by s.itemgubun asc, s.shopitemid desc, s.itemoption"
		end if

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new COffShopOneItem

				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).Fshopbuyprice  = rsget("shopbuyprice")
				FItemList(i).FimageSmall     = rsget("smallimage")

				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

				FItemList(i).Fsellyn               = rsget("sellyn")
                FItemList(i).Flimityn              = rsget("limityn")

                if FItemList(i).Fitemoption="0000" then
				    FItemList(i).Flimitno              = rsget("limitno")
				    FItemList(i).Flimitsold            = rsget("limitsold")
				else
				    if rsget("optsellyn")="N" then FItemList(i).Fsellyn="N"
				    FItemList(i).Flimitno              = rsget("optlimitno")
				    FItemList(i).Flimitsold            = rsget("optlimitsold")
				end if

				FItemList(i).flogicsipgono = rsget("logicsipgono")
				FItemList(i).flogicsreipgono = rsget("logicsreipgono")
				FItemList(i).fbrandipgono = rsget("brandipgono")
				FItemList(i).fbrandreipgono = rsget("brandreipgono")
				FItemList(i).fsellno = rsget("sellno")
				FItemList(i).fresellno = rsget("resellno")
				FItemList(i).ferrsampleitemno = rsget("errsampleitemno")
				FItemList(i).ferrbaditemno = rsget("errbaditemno")
				FItemList(i).ferrrealcheckno = rsget("errrealcheckno")
				FItemList(i).fsysstockno = rsget("sysstockno")
				FItemList(i).frealstockno = rsget("realstockno")
				FItemList(i).frequiredStock = rsget("requiredStock")
				FItemList(i).fsell7days = rsget("sell7days")
				FItemList(i).fsell3days = rsget("sell3days")
				FItemList(i).frequire3daystock = rsget("require3daystock")
				FItemList(i).frequire7daystock = rsget("require7daystock")
				FItemList(i).frequire14daystock = rsget("require14daystock")
				FItemList(i).fpreorderno = rsget("preorderno")
				FItemList(i).fpreordernofix = rsget("preordernofix")

				FItemList(i).FMakerMargin    = rsget("defaultmargin")
				FItemList(i).FShopMargin    = rsget("defaultsuplymargin")
				FItemList(i).Foptusing      = rsget("optusing")
				FItemList(i).Fonlinebuycash = rsget("onlinebuycash")
                ''옵션 추가금액
			    FItemList(i).FOnlineOptaddprice = rsget("optaddprice")
			    FItemList(i).FOnlineOptaddbuyprice = rsget("optaddbuyprice")
				FItemList(i).FOffimgMain	= rsget("offimgmain")
				FItemList(i).FOffimgList	= rsget("offimglist")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				FItemList(i).FdeliveryType  = rsget("deliverytype")
				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
				FItemList(i).FOffsell7days     = rsget("sell7days")
				if (FRectOrder<>"byoffbestAll") then FItemList(i).fsell3days     = rsget("sell3days")

				FItemList(i).Frealstockno = rsget("realstockno")

				FItemList(i).Fcentermwdiv			= rsget("centermwdiv")
				FItemList(i).Fmwdiv			= rsget("mwdiv")

				FItemList(i).Fcomm_cd       = rsget("comm_cd")

				FItemList(i).Fforeign_sellprice	 = rsget("orgprice")
				FItemList(i).Fforeign_suplyprice = rsget("suplyprice")

				FItemList(i).Fitemno			= rsget("itemno")

				FItemList(i).FShopItemOrgprice  = rsget("orgsellprice")
				FItemList(i).FLogicsRealStock  = rsget("logicsRealStock")

				'/홀쎄일의 경우 대표화폐가 한화 일경우 초기값 셋팅
				if Floginsite="WSLWEB" then
					if (fcurrencyUnit = "KRW" or fcurrencyUnit = "WON") Then
						' 홀쎄일인데 국내에서 도매로 사용하는 케이스가 있음.
						FItemList(i).Fshopitemprice     = FItemList(i).Fforeign_sellprice
						FItemList(i).Fshopsuplycash     = FItemList(i).Fforeign_suplyprice
					end if
				end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end sub

	public sub GetOffLineJumunItem()
		dim sqlStr,i

		sqlStr = " select count(s.shopitemid) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " , [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on s.itemgubun='10' and s.shopitemid=i.itemid"

		if FRectOrder="byetc" then
			sqlStr = sqlStr + " where s.itemgubun ='70'"
		elseif FRectOrder="byevent" then
			sqlStr = sqlStr + " where s.itemgubun ='80'"
		else
			sqlStr = sqlStr + " where s.itemgubun <>'70'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
		end if

		sqlStr = sqlStr + " and d.makerid=s.makerid"

		if FRectDesignerjungsangubun<>"" then
			sqlStr = sqlStr + " and d.chargediv in (" + FRectDesignerjungsangubun + ")"
		end if

		if FRectAdminView<>"on" then
			sqlStr = sqlStr + " and s.isusing='Y'"
			sqlStr = sqlStr + " and (i.isusing='Y' or i.isusing is null)"
		end if

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount>=3000 then FTotalCount=3000

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption, "
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname,"
		sqlStr = sqlStr + " s.shopitemprice, s.shopsuplycash, s.shopbuyprice, i.smallimage,"
		sqlStr = sqlStr + " IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist, IsNULL(s.offimgsmall,'') as offimgsmall,"
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, IsNULL(i.buycash,0) as onlinebuycash, i.deliverytype, "
		sqlStr = sqlStr + " d.defaultmargin,d.defaultsuplymargin, "
		sqlStr = sqlStr + " IsNULL(o.isusing,'Y') as optusing, "
		sqlStr = sqlStr + " IsNULL(o.optlimitno,0) as optlimitno, IsNULL(o.optlimitsold,0) as optlimitsold, IsNULL(o.optsellyn,'Y') optsellyn"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " ,[db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"

		if FRectOrder="byetc" then
			sqlStr = sqlStr + " where s.itemgubun ='70'"
		elseif FRectOrder="byevent" then
			sqlStr = sqlStr + " where s.itemgubun ='80'"
		else
			sqlStr = sqlStr + " where s.itemgubun <>'70'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
		end if

		sqlStr = sqlStr + " and d.makerid=s.makerid"

		if FRectDesignerjungsangubun<>"" then
			sqlStr = sqlStr + " and d.chargediv in (" + FRectDesignerjungsangubun + ")"
		end if

		if FRectAdminView<>"on" then
			sqlStr = sqlStr + " and s.isusing='Y'"
			sqlStr = sqlStr + " and (i.isusing='Y' or i.isusing is null)"
		end if

		sqlStr = sqlStr + " order by s.itemgubun asc, s.shopitemid desc, s.itemoption"

		''response.write sqlStr
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
				set FItemList(i) = new COffShopOneItem
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).FimageSmall     = rsget("smallimage")

				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

				FItemList(i).Fsellyn               = rsget("sellyn")
                FItemList(i).Flimityn              = rsget("limityn")

                if FItemList(i).Fitemoption="0000" then
				    FItemList(i).Flimitno              = rsget("limitno")
				    FItemList(i).Flimitsold            = rsget("limitsold")
				else
				    if rsget("optsellyn")="N" then FItemList(i).Fsellyn="N"
				    FItemList(i).Flimitno              = rsget("optlimitno")
				    FItemList(i).Flimitsold            = rsget("optlimitsold")
				end if

				FItemList(i).FMakerMargin    = rsget("defaultmargin")
				FItemList(i).FShopMargin    = rsget("defaultsuplymargin")
				FItemList(i).Foptusing      = rsget("optusing")
				FItemList(i).Fonlinebuycash = rsget("onlinebuycash")
				FItemList(i).FOffimgMain	= rsget("offimgmain")
				FItemList(i).FOffimgList	= rsget("offimglist")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				FItemList(i).FdeliveryType  = rsget("deliverytype")
				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
				FItemList(i).Fshopbuyprice  = rsget("shopbuyprice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub

	public sub GetOffLineBestItem() '' where using?
		dim sqlStr,i
		sqlStr = " select count(s.shopitemid) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " , [db_const].[dbo].tbl_const_offshop c"
		sqlStr = sqlStr + " , [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " where s.itemgubun<>'00'"
		sqlStr = sqlStr + " and s.itemgubun=c.itemgubun"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
		end if
		sqlStr = sqlStr + " and s.shopitemid=c.shopitemid"
		sqlStr = sqlStr + " and s.itemoption=c.itemoption"
		sqlStr = sqlStr + " and s.isusing='Y'"
		if FRectShopid<>"" then
			sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
		end if
		sqlStr = sqlStr + " and d.makerid=s.makerid"
		if FRectDesignerjungsangubun<>"" then
			sqlStr = sqlStr + " and d.chargediv in (" + FRectDesignerjungsangubun + ")"
		end if
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		'response.write sqlStr
		if FTotalCount>=3000 then FTotalCount=3000

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption, "
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname,"
		sqlStr = sqlStr + " s.shopitemprice, s.shopsuplycash, i.smallimage,"
		sqlStr = sqlStr + " IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist, IsNULL(s.offimgsmall,'') as offimgsmall,"
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold,"
		sqlStr = sqlStr + " d.defaultmargin,d.defaultsuplymargin, "
        sqlStr = sqlStr + " IsNULL(o.optlimitno,0) as optlimitno, IsNULL(o.optlimitsold,0) as optlimitsold, IsNULL(o.optsellyn,'Y') optsellyn"
		sqlStr = sqlStr + " from [db_const].[dbo].tbl_const_offshop c"
		sqlStr = sqlStr + " , [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " , [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr + " where s.itemgubun<>'00'"
		sqlStr = sqlStr + " and s.itemgubun=c.itemgubun"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
		end if

		sqlStr = sqlStr + " and s.shopitemid=c.shopitemid"
		sqlStr = sqlStr + " and s.itemoption=c.itemoption"
		sqlStr = sqlStr + " and s.isusing='Y'"

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
		end if

		sqlStr = sqlStr + " and d.makerid=s.makerid"

		if FRectDesignerjungsangubun<>"" then
			sqlStr = sqlStr + " and d.chargediv in (" + FRectDesignerjungsangubun + ")"
		end if

		sqlStr = sqlStr + " order by c.sellcnt desc, s.itemgubun desc, s.shopitemid desc"

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
				set FItemList(i) = new COffShopOneItem
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")

				FItemList(i).FimageSmall     = rsget("smallimage")
				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

				FItemList(i).Fsellyn               = rsget("sellyn")
				FItemList(i).Flimityn              = rsget("limityn")

				if FItemList(i).Fitemoption="0000" then
                    FItemList(i).Flimitno              = rsget("limitno")
                    FItemList(i).Flimitsold            = rsget("limitsold")
                else
                    if rsget("optsellyn")="N" then FItemList(i).Fsellyn="N"
                    FItemList(i).Flimitno              = rsget("optlimitno")
                    FItemList(i).Flimitsold            = rsget("optlimitsold")
                end if

				FItemList(i).FMakerMargin    = rsget("defaultmargin")
				FItemList(i).FShopMargin    = rsget("defaultsuplymargin")
				FItemList(i).FOffimgMain	= rsget("offimgmain")
				FItemList(i).FOffimgList	= rsget("offimglist")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub

	public sub GetOnlineBestItem() '' where using?
		dim sqlStr,i

		sqlStr = " select count(s.shopitemid) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " , [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " join [db_item].[dbo].tbl_item i on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " where s.itemgubun='10'"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
		end if

		sqlStr = sqlStr + " and s.isusing='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.sellyn<>'N'"

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"

		end if
		sqlStr = sqlStr + " and d.makerid=s.makerid"

		if FRectDesignerjungsangubun<>"" then
			sqlStr = sqlStr + " and d.chargediv in (" + FRectDesignerjungsangubun + ")"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount>=1000 then FTotalCount=1000

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption, "
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname,"
		sqlStr = sqlStr + " s.shopitemprice, s.shopsuplycash, s.shopbuyprice, i.smallimage,"
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, i.deliverytype,"
		sqlStr = sqlStr + " d.defaultmargin,d.defaultsuplymargin, "
		sqlStr = sqlStr + " IsNULL(o.isusing,'Y') as optusing,"
        sqlStr = sqlStr + " IsNULL(o.optlimitno,0) as optlimitno, IsNULL(o.optlimitsold,0) as optlimitsold, IsNULL(o.optsellyn,'Y') optsellyn"

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " , [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " join [db_item].[dbo].tbl_item i on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_Contents c on s.itemgubun='10' and s.shopitemid=c.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr + " where s.itemgubun='10'"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
		end if
		sqlStr = sqlStr + " and s.isusing='Y'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and i.sellyn<>'N'"
		if FRectShopid<>"" then
			sqlStr = sqlStr + " and d.shopid='" + FRectShopid + "'"
		end if
		sqlStr = sqlStr + " and d.makerid=s.makerid"
		if FRectDesignerjungsangubun<>"" then
			sqlStr = sqlStr + " and d.chargediv in (" + FRectDesignerjungsangubun + ")"
		end if
		if FRectOrder="byonfav" then
			sqlStr = sqlStr + " order by c.recentfavcount desc, i.itemscore desc, s.itemgubun desc, s.shopitemid desc"
		else
			sqlStr = sqlStr + " order by c.recentsellcount desc, i.itemscore desc, s.itemgubun desc, s.shopitemid desc"
		end if

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
				set FItemList(i) = new COffShopOneItem
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")

				FItemList(i).FimageSmall     = rsget("smallimage")
				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

				FItemList(i).Fsellyn               = rsget("sellyn")
				FItemList(i).Flimityn              = rsget("limityn")

				if FItemList(i).Fitemoption="0000" then
                    FItemList(i).Flimitno              = rsget("limitno")
                    FItemList(i).Flimitsold            = rsget("limitsold")
                else
                    if rsget("optsellyn")="N" then FItemList(i).Fsellyn="N"
                    FItemList(i).Flimitno              = rsget("optlimitno")
                    FItemList(i).Flimitsold            = rsget("optlimitsold")
                end if

				FItemList(i).FMakerMargin    = rsget("defaultmargin")
				FItemList(i).FShopMargin    = rsget("defaultsuplymargin")
				FItemList(i).Foptusing = rsget("optusing")
				FItemList(i).FdeliveryType  = rsget("deliverytype")
				FItemList(i).Fshopbuyprice  = rsget("shopbuyprice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub

	'///common/offshop/popshopitem2.asp  :: 업체위탁주문쪽.
	public function GetOffShopItemList()
		dim sqlStr,i ,defaultmargin, defaultsuplymargin , iStartDate , sqlsearch
        dim comm_cd

        if (FRectOrder="byrecent") then
            sqlStr = "select convert(varchar(10),dateadd(d,-14,getdate()),21) as stdt "
            rsget.Open sqlStr,dbget,1
    		iStartDate = rsget("stdt")
    		rsget.Close
        end if

		'''===오프 기본마진 getDefaultmargin ===============================
		sqlStr = " select top 1 * "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d with (nolock)"
		sqlStr = sqlStr + " where makerid='" + FRectDesigner + "'"
		if (FRectShopid<>"") then
			sqlStr = sqlStr + " and shopid='" + FRectShopid + "'"
		end if
		sqlStr = sqlStr + " order by shopid"
		rsget.Open sqlStr,dbget,1
		if  not rsget.EOF  then
			defaultmargin = rsget("defaultmargin")
			defaultsuplymargin = rsget("defaultsuplymargin")
			comm_cd = rsget("comm_cd")
		end if
		rsget.Close

		'''=== 검색어 처리===============================
		if FRectItemgubun<>"" then
			sqlsearch = sqlsearch + " and s.itemgubun='" + FRectItemgubun + "'"
		end if

		if FRectItemName<>"" then
			sqlsearch = sqlsearch + " and shopitemname like '%" + FRectItemName + "%'"
		end if

		if FRectOnlyOffLine<>"" then
			sqlsearch = sqlsearch + " and s.itemgubun='" + FRectOnlyOffLine + "'"
		end if

		if FRectOnlyUsing<>"" then
            sqlsearch = sqlsearch & " and s.isusing='" + FRectOnlyUsing + "'"
		end if

        if (FRectOnlyActive<>"") then
		    sqlsearch = sqlsearch + " and (IsNULL(i.sellyn,'')<>'N') and s.isusing='Y'"
		end if

		if (FRectOrder="all") then
			''전체
		elseif FRectOrder="byetc" then
			sqlsearch = sqlsearch + " and s.itemgubun ='70'"
		elseif FRectOrder="byevent" then
			sqlsearch = sqlsearch + " and s.itemgubun ='80'"
		elseif FRectOrder="by7sell" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    if FRectDesigner<>"" then
		        sqlsearch = sqlsearch + " and st.sell7days>0"
		    else
		        sqlsearch = sqlsearch + " and st.sell7days>1"
		    end if
		elseif FRectOrder="byrecent" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and s.regdate>'" & iStartDate & "'"
		elseif FRectOrder="byonbest" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and i.itemscore>0"
		elseif FRectOrder="byoffbest" then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and st.sell7days>0"
		elseif (FRectOrder="byoffbestAll") then
		    sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		    sqlsearch = sqlsearch + " and stb.sellcnt>0"
		else
			sqlsearch = sqlsearch + " and s.itemgubun <>'70'"
		end If

		if FRectIpGoOnly = "on" then
			sqlsearch = sqlsearch + " and (st.logicsipgono+st.logicsreipgono > 0 or st.brandipgono+st.brandreipgono > 0)"
		end if

		'/최근7일 판매건만
		if FRectSell7days = "on" then
			sqlsearch = sqlsearch + " and st.sell7days > 0"
		end if

		'기주문포함부족상품
		if FRectIncludePreOrder = "on" then
	        if FRectShortageType="3" then
	    		sqlsearch = sqlsearch + " and (st.sell3days*1) - (db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno)+st.preordernofix) > 0"
	    	elseif FRectShortageType="7" then
	    		sqlsearch = sqlsearch + " and (st.sell7days*1) - (db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno)+st.preordernofix) > 0"
	    	elseif FRectShortageType="14" then
	    		sqlsearch = sqlsearch + " and (st.sell7days*2) - (db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno)+st.preordernofix) > 0"
	    	elseif FRectShortageType="28" then
	    		sqlsearch = sqlsearch + " and (st.sell7days*4) - (db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno)+st.preordernofix) > 0"
	    	else
	    		sqlsearch = sqlsearch + " and st.preordernofix + db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno) < 1"
	    		'sqlsearch = sqlsearch + " and (st.sell7days*1) - (db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno)+st.preordernofix) > 0"
	    	end if
		else
	        if FRectShortageType="3" then
	    		sqlsearch = sqlsearch + " and (st.sell3days*1) - db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno) > 0"
	    	elseif FRectShortageType="7" then
	    		sqlsearch = sqlsearch + " and (st.sell7days*1) - db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno) > 0"
	    	elseif FRectShortageType="14" then
	    		sqlsearch = sqlsearch + " and (st.sell7days*2) - db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno) > 0"
	    	elseif FRectShortageType="28" then
	    		sqlsearch = sqlsearch + " and (st.sell7days*4) - db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno) > 0"
	    	else

	    	end if
		end If

		if FRectOnlineMWdiv <> "" then
			sqlsearch = sqlsearch + " and i.mwdiv = '" & FRectOnlineMWdiv & "' "
		end if
		IF FRectItemid <> "" Then
			sqlsearch = sqlsearch & " and s.shopitemid in ("& FRectItemID&")"
		END IF

		sqlStr = " select count(s.shopitemid) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + " on s.itemgubun='10' and s.shopitemid=i.itemid"

		'물류센터 재고
        if FRectLogicsIpGoOnly = "on" then
		    sqlStr = sqlStr + " join db_summary.dbo.tbl_current_logisstock_summary lc with (nolock)"
    	    sqlStr = sqlStr + "     on s.itemgubun=lc.itemgubun and s.shopitemid=lc.itemid and s.itemoption=lc.itemoption"
		else
		    sqlStr = sqlStr + " left join db_summary.dbo.tbl_current_logisstock_summary lc with (nolock)"
    	    sqlStr = sqlStr + "     on s.itemgubun=lc.itemgubun and s.shopitemid=lc.itemid and s.itemoption=lc.itemoption"
        end if
		'매장 재고
		if FRectIpGoOnly = "on" then
			sqlStr = sqlStr + " join db_summary.dbo.tbl_current_shopstock_summary st with (nolock)"
			sqlStr = sqlStr + " on st.shopid='" + FRectShopid + "' and s.itemgubun=st.itemgubun and s.shopitemid=st.itemid and s.itemoption=st.itemoption"
		else
			sqlStr = sqlStr + " left join db_summary.dbo.tbl_current_shopstock_summary st with (nolock)"
			sqlStr = sqlStr + " on st.shopid='" + FRectShopid + "' and s.itemgubun=st.itemgubun and s.shopitemid=st.itemid and s.itemoption=st.itemoption"
        end if

		If FRectCopyIdx <> "" Then
			sqlStr = sqlStr + " 	join [db_shop].[dbo].tbl_shop_ipchul_detail dd with (nolock)" & VbCrLf
			sqlStr = sqlStr + " 	on " & VbCrLf
			sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
			sqlStr = sqlStr + " 		and dd.masteridx = " & FRectCopyIdx & VbCrLf
			sqlStr = sqlStr + " 		and s.itemgubun = dd.itemgubun " & VbCrLf
			sqlStr = sqlStr + " 		and s.shopitemid = dd.shopitemid " & VbCrLf
			sqlStr = sqlStr + " 		and s.itemoption = dd.itemoption " & VbCrLf
		End If

		if (FRectOrder="byoffbestAll") then
		    sqlStr = sqlStr + " left join db_const.dbo.tbl_const_offshop stb with (nolock)"
            sqlStr = sqlStr + " on s.itemgubun=stb.itemgubun and s.shopitemid=stb.shopitemid and s.itemoption=stb.itemoption"
		end if

		sqlStr = sqlStr + " where s.makerid='" + FRectDesigner + "' " & sqlsearch

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption,"
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname,"
		sqlStr = sqlStr + " s.orgsellprice, s.shopitemprice, s.shopsuplycash, s.shopbuyprice, s.isusing, s.regdate, s.extbarcode, s.offimgsmall, i.smallimage"
		''sqlStr = sqlStr + " ,IsNULL(st.sell7days,0) as sell7days , IsNULL(st.sell3days,0) as sell3days"

		if (FRectOrder="byoffbestAll") then
		    sqlStr = sqlStr + " ,IsNULL(stb.sellcnt,0) as sell7daysALL "
		end If

		'// 매장재고
		sqlStr = sqlStr & " , st.shopid, isnull(st.logicsipgono,0) as logicsipgono, isnull(st.logicsreipgono,0) as logicsreipgono"
		sqlStr = sqlStr & " , isnull(st.brandipgono,0) as brandipgono, isnull(st.brandreipgono,0) as brandreipgono"
		sqlStr = sqlStr & " , isnull(st.sellno,0) as sellno, isnull(st.resellno,0) as resellno,isnull(st.errsampleitemno,0) as errsampleitemno"
		sqlStr = sqlStr & " , isnull(st.errbaditemno,0) as errbaditemno, isnull(st.errrealcheckno,0) as errrealcheckno"
		sqlStr = sqlStr & " , isnull(st.sysstockno,0) as sysstockno, isnull(st.realstockno,0) as realstockno, isnull(st.requiredStock,0) as requiredStock"
		sqlStr = sqlStr & " , isnull(st.sell7days,0) as sell7days, isnull(st.sell3days,0) as sell3days ,st.lastupdate"
		sqlStr = sqlStr & " , st.preorderno ,st.preordernofix"

		if FRectIncludePreOrder = "on" then
			sqlStr = sqlStr & " ,( (st.sell3days*1) - (db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno)+st.preordernofix) ) as require3daystock"
			sqlStr = sqlStr & " ,( (st.sell7days*1) - (db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno)+st.preordernofix) ) as require7daystock"
			sqlStr = sqlStr & " ,( (st.sell7days*2) - (db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno)+st.preordernofix) ) as require14daystock"
		Else
			sqlStr = sqlStr & " ,( (st.sell3days*1) - (db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno)) ) as require3daystock"
			sqlStr = sqlStr & " ,( (st.sell7days*1) - (db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno)) ) as require7daystock"
			sqlStr = sqlStr & " ,( (st.sell7days*2) - (db_summary.[dbo].[uf_replacezero](st.realstockNo+st.errsampleitemno+st.errbaditemno)) ) as require14daystock"
		End If

		sqlStr = sqlStr + " ,IsNULL(i.mwdiv,'') as mwdiv, s.centermwdiv "

		If FRectCopyIdx <> "" Then
			sqlStr = sqlStr + " ,dd.itemno "
		Else
			sqlStr = sqlStr + " ,0 as itemno "
		End If

		sqlStr = sqlStr + " , IsNULL(i.buycash,0) as buycash, IsNULL(o.optaddbuyprice,0) as optaddbuyprice "

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + " on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o with (nolock) on i.itemid = o.itemid and s.itemoption = o.itemoption"

		'물류센터 재고
        if FRectLogicsIpGoOnly = "on" then
		    sqlStr = sqlStr + " join db_summary.dbo.tbl_current_logisstock_summary lc with (nolock)"
    	    sqlStr = sqlStr + "     on s.itemgubun=lc.itemgubun and s.shopitemid=lc.itemid and s.itemoption=lc.itemoption"
		else
		    sqlStr = sqlStr + " left join db_summary.dbo.tbl_current_logisstock_summary lc with (nolock)"
    	    sqlStr = sqlStr + "     on s.itemgubun=lc.itemgubun and s.shopitemid=lc.itemid and s.itemoption=lc.itemoption"
        end if
		'매장 재고
		if FRectIpGoOnly = "on" then
			sqlStr = sqlStr + " join db_summary.dbo.tbl_current_shopstock_summary st with (nolock)"
			sqlStr = sqlStr + " on st.shopid='" + FRectShopid + "' and s.itemgubun=st.itemgubun and s.shopitemid=st.itemid and s.itemoption=st.itemoption"
		else
			sqlStr = sqlStr + " left join db_summary.dbo.tbl_current_shopstock_summary st with (nolock)"
			sqlStr = sqlStr + " on st.shopid='" + FRectShopid + "' and s.itemgubun=st.itemgubun and s.shopitemid=st.itemid and s.itemoption=st.itemoption"
        end if

		if (FRectOrder="byoffbestAll") then
		    sqlStr = sqlStr + " left join db_const.dbo.tbl_const_offshop stb with (nolock)"
            sqlStr = sqlStr + " on s.itemgubun=stb.itemgubun and s.shopitemid=stb.shopitemid and s.itemoption=stb.itemoption"
		end If

		If FRectCopyIdx <> "" Then
			sqlStr = sqlStr + " 	join [db_shop].[dbo].tbl_shop_ipchul_detail dd with (nolock)" & VbCrLf
			sqlStr = sqlStr + " 	on " & VbCrLf
			sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
			sqlStr = sqlStr + " 		and dd.masteridx = " & FRectCopyIdx & VbCrLf
			sqlStr = sqlStr + " 		and s.itemgubun = dd.itemgubun " & VbCrLf
			sqlStr = sqlStr + " 		and s.shopitemid = dd.shopitemid " & VbCrLf
			sqlStr = sqlStr + " 		and s.itemoption = dd.itemoption " & VbCrLf
		End If

		sqlStr = sqlStr + " where s.makerid='" + FRectDesigner + "' " & sqlsearch

        if FRectOrder="by7sell" then
            sqlStr = sqlStr + " order by sell7days desc,s.itemgubun asc, s.shopitemid desc, s.itemoption"
		elseif FRectOrder="byrecent" then
		    sqlStr = sqlStr + " order by s.regdate desc"
		elseif FRectOrder="byonbest" then
		    sqlStr = sqlStr + " order by i.itemscore desc,s.itemgubun asc, s.shopitemid desc, s.itemoption"
		elseif (FRectOrder="byoffbest") or (FRectOrder="byoffbestAll") then
		   sqlStr = sqlStr + " order by sell7days desc,s.itemgubun asc, s.shopitemid desc, s.itemoption"
		elseif FRectOrder="BI" Then
		   sqlStr = sqlStr + " order by s.makerid, s.itemgubun asc, s.shopitemid desc, s.itemoption"
		elseif FRectOrder="I" Then
		   sqlStr = sqlStr + " order by s.itemgubun asc, s.shopitemid desc, s.itemoption"
		else
		    sqlStr = sqlStr + " order by s.itemgubun asc, s.shopitemid desc, s.itemoption"
		end if

		''response.write sqlStr & "<br>"
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
				set FItemList(i) = new COffShopOneItem
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).FShopItemOrgprice  = rsget("orgsellprice")
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).Fisusing           = rsget("isusing")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).Fextbarcode 		= rsget("extbarcode")
                FItemList(i).Fshopbuyprice      = rsget("shopbuyprice")
				FItemList(i).FimageSmall     = rsget("smallimage")

				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
				if FItemList(i).Fitemgubun<>"10" then FItemList(i).FimageSmall = FItemList(i).FOffimgSmall
				FItemList(i).FMakerMargin   = defaultmargin
				FItemList(i).FShopMargin    = defaultsuplymargin
                FItemList(i).Fcomm_cd       = comm_cd

				FItemList(i).Flogicsipgono = rsget("logicsipgono")
				FItemList(i).Flogicsreipgono = rsget("logicsreipgono")
				FItemList(i).Fbrandipgono = rsget("brandipgono")
				FItemList(i).Fbrandreipgono = rsget("brandreipgono")
				FItemList(i).Fsellno = rsget("sellno")
				FItemList(i).Fresellno = rsget("resellno")
				FItemList(i).Ferrsampleitemno = rsget("errsampleitemno")
				FItemList(i).Ferrbaditemno = rsget("errbaditemno")
				FItemList(i).Ferrrealcheckno = rsget("errrealcheckno")
				FItemList(i).Fsysstockno = rsget("sysstockno")
				FItemList(i).Frealstockno = rsget("realstockno")
				FItemList(i).FrequiredStock = rsget("requiredStock")
				FItemList(i).Fsell7days = rsget("sell7days")
				FItemList(i).Fsell3days = rsget("sell3days")
				FItemList(i).Frequire3daystock = rsget("require3daystock")
				FItemList(i).Frequire7daystock = rsget("require7daystock")
				FItemList(i).Frequire14daystock = rsget("require14daystock")
				FItemList(i).Fpreorderno = rsget("preorderno")
				FItemList(i).Fpreordernofix = rsget("preordernofix")

				FItemList(i).Fcentermwdiv			= rsget("centermwdiv")
				FItemList(i).Fmwdiv					= rsget("mwdiv")
				FItemList(i).Fitemno				= rsget("itemno")
				FItemList(i).FOnlinebuycash			= rsget("buycash")					'// 2018-02-12, skyer9, 온라인매입가 추가
				FItemList(i).FOnlineOptaddbuyprice	= rsget("optaddbuyprice")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	'//admin/offshop/barcodeprint.asp		'/2013.03.06 한용민 생성
	public function GetBarCodeList_paging()
		dim ibargubun, ibaritemid, ibaritemoption, sqlStr, sqlsearch

		if (FRectBarCode <> "") then
			FRectPrdCode = FRectBarCode
		end if

        if (FRectItemgubun<>"") then
            sqlsearch = sqlsearch + " and s.itemgubun='" & FRectItemgubun & "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and s.shopitemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and s.shopitemid in (" + FRectItemid + ")"
            end if
        end if

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and s.makerid='" + FRectDesigner + "'"
		end if

		if (FRectOnlyUsing="on") then
		    sqlsearch = sqlsearch + " and s.isusing='Y'"
        end if

		if FRectPrdCode<>"" then
		    if (Len(FRectPrdCode) = 12) then
				sqlsearch = sqlsearch + " 	and s.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			else
				sqlsearch = sqlsearch + " 	and s.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and s.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			end if
		end if

		if FRectGeneralBarcode<>"" then
		sqlsearch = sqlsearch + " 	and ba.barcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if

        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " and s.catecdl='" + FRectCDL + "'"
        end if

        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " and s.catecdm='" + FRectCDM + "'"
        end if

        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " and s.catecdn='" + FRectCDS + "'"
        end if

		if (FRectItemName<>"") then
		    sqlsearch = sqlsearch + " and s.shopitemname like '%" + CStr(FRectItemName) + "%'"
		end if

		if FRectShopItemName<>"" then
			sqlsearch = sqlsearch + " and l.lcitemname like '%" + FRectShopItemName + "%'"
		end if

		'총 갯수 구하기
		sqlStr = "select"
		sqlStr = sqlStr + " count(*) as cnt"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " join [db_user].[dbo].tbl_user_c c "
		sqlStr = sqlStr + " 	on s.makerid=c.userid "
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock ba "
		sqlStr = sqlStr + " 	on s.itemgubun = ba.itemgubun "
		sqlStr = sqlStr + " 	and s.shopitemid = ba.itemid "
		sqlStr = sqlStr + " 	and s.itemoption = ba.itemoption "
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_locale_item l "
		sqlStr = sqlStr + " 	on l.shopid = '" & CStr(FRectShopid) & "' "
		sqlStr = sqlStr + " 	and s.shopitemid = l.shopitemid "
		sqlStr = sqlStr + " 	and s.itemoption = l.itemoption "
		sqlStr = sqlStr + " 	and s.itemgubun = l.itemgubun "
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item i 	"
		sqlStr = sqlStr & " 	on s.itemgubun='10' "
		sqlStr = sqlStr & " 	and s.shopitemid = i.itemid "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit function

		'데이터 리스트
		sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " s.itemgubun, s.shopitemid, s.itemoption, s.makerid, s.shopitemname, s.shopitemoptionname, s.shopitemprice, s.shopsuplycash"
		sqlStr = sqlStr + " , s.isusing, s.regdate, c.socname, c.socname_kor, IsNULL(s.offimgsmall,'') as offimgsmall, IsNULL(i.smallimage,'') as smallimage"
		sqlStr = sqlStr + " , s.extbarcode as generalbarcode"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " join [db_user].[dbo].tbl_user_c c "
		sqlStr = sqlStr + " 	on s.makerid=c.userid "
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock ba "
		sqlStr = sqlStr + " 	on s.itemgubun = ba.itemgubun "
		sqlStr = sqlStr + " 	and s.shopitemid = ba.itemid "
		sqlStr = sqlStr + " 	and s.itemoption = ba.itemoption "
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_locale_item l "
		sqlStr = sqlStr + " 	on l.shopid = '" & CStr(FRectShopid) & "' "
		sqlStr = sqlStr + " 	and s.shopitemid = l.shopitemid "
		sqlStr = sqlStr + " 	and s.itemoption = l.itemoption "
		sqlStr = sqlStr + " 	and s.itemgubun = l.itemgubun "
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item i 	"
		sqlStr = sqlStr & " 	on s.itemgubun='10' "
		sqlStr = sqlStr & " 	and s.shopitemid = i.itemid "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by s.itemgubun desc, s.shopitemid desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF

				set FItemList(i) = new COffShopOneItem

				FItemList(i).Fitemgubun          = rsget("itemgubun")
				FItemList(i).Fshopitemid         = rsget("shopitemid")
				FItemList(i).Fitemoption         = rsget("itemoption")
				FItemList(i).Fmakerid            = rsget("makerid")
				FItemList(i).Fshopitemname       = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname = db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice      = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash      = rsget("shopsuplycash")
				FItemList(i).Fisusing            = rsget("isusing")
				FItemList(i).Fregdate            = rsget("regdate")
				FItemList(i).FSocName			 = db2html(rsget("socname"))
				FItemList(i).FSocName_Kor		 = db2html(rsget("socname_kor"))
				FItemList(i).Fgeneralbarcode    = db2html(rsget("generalbarcode"))
				FItemList(i).FOffimgSmall		= rsget("offimgsmall")
				FItemList(i).FimageSmall     	= rsget("smallimage")

				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

				if FItemList(i).FOffimgSmall<>"" then
					FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
				end if

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end function

	public function GetBarCodeList()
	    '' 페이징 필요
		dim sqlStr
		dim ibargubun, ibaritemid, ibaritemoption

		if (FRectBarCode <> "") then
			FRectPrdCode = FRectBarCode
		end if

		sqlStr = " select top 500 s.itemgubun, s.shopitemid, s.itemoption, s.makerid," & VbCrLf
		sqlStr = sqlStr + " s.shopitemname, s.shopitemoptionname, s.shopitemprice," & VbCrLf
		sqlStr = sqlStr + " s.shopsuplycash, s.isusing, s.regdate, c.socname, c.socname_kor, IsNULL(s.offimgsmall,'') as offimgsmall, IsNULL(i.smallimage,'') as smallimage " & VbCrLf
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item s " & VbCrLf

		sqlStr = sqlStr + " 	join [db_user].[dbo].tbl_user_c c " & VbCrLf
		sqlStr = sqlStr + " 	on " & VbCrLf
		sqlStr = sqlStr + " 		s.makerid=c.userid " & VbCrLf

		'범용바코드 검색
		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_stock ba " & VbCrLf
		sqlStr = sqlStr + " 	on " & VbCrLf
		sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
		sqlStr = sqlStr + " 		and s.itemgubun = ba.itemgubun " & VbCrLf
		sqlStr = sqlStr + " 		and s.shopitemid = ba.itemid " & VbCrLf
		sqlStr = sqlStr + " 		and s.itemoption = ba.itemoption " & VbCrLf

		'샵별 상품명
		sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_locale_item l " & VbCrLf
		sqlStr = sqlStr + " 	on " & VbCrLf
		sqlStr = sqlStr + " 		1 = 1 " & VbCrLf
		sqlStr = sqlStr + " 		and l.shopid = '" & CStr(FRectShopid) & "' " & VbCrLf
		sqlStr = sqlStr + " 		and s.shopitemid = l.shopitemid " & VbCrLf
		sqlStr = sqlStr + " 		and s.itemoption = l.itemoption " & VbCrLf
		sqlStr = sqlStr + " 		and s.itemgubun = l.itemgubun " & VbCrLf

		'상품이미지
		sqlStr = sqlStr & " 	left join db_item.dbo.tbl_item i 	" & vbCrLf
		sqlStr = sqlStr & " 	on " & vbCrLf
		sqlStr = sqlStr & " 		1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemgubun='10' " & vbCrLf
		sqlStr = sqlStr & " 		and s.shopitemid = i.itemid " & vbCrLf

		sqlStr = sqlStr + " where 1 = 1 "
        if (FRectItemgubun<>"") then
            sqlStr = sqlStr + " and s.itemgubun='" & FRectItemgubun & "'" & VbCrLf
        end if

		if FRectItemId<>"" then
			sqlStr = sqlStr + " and s.shopitemid=" + CStr(FRectItemId) & VbCrLf
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'" & VbCrLf
		end if

		if (FRectOnlyUsing="on") then
		    sqlStr = sqlStr + " and s.isusing='Y'" & VbCrLf
        end if

		if FRectPrdCode<>"" then
		    if (Len(FRectPrdCode) = 12) then
				sqlStr = sqlStr + " 	and s.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlStr = sqlStr + " 	and s.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) + " " & VbCrLf
				sqlStr = sqlStr + " 	and s.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			else
				sqlStr = sqlStr + " 	and s.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlStr = sqlStr + " 	and s.shopitemid = " + RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) + " " & VbCrLf
				sqlStr = sqlStr + " 	and s.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			end if
		end if

		if FRectGeneralBarcode<>"" then
		sqlStr = sqlStr + " 	and ba.barcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if

        if (FRectCDL<>"") then
            sqlStr = sqlStr + " and s.catecdl='" + FRectCDL + "'" & VbCrLf
        end if

        if (FRectCDM<>"") then
            sqlStr = sqlStr + " and s.catecdm='" + FRectCDM + "'" & VbCrLf
        end if

        if (FRectCDS<>"") then
            sqlStr = sqlStr + " and s.catecdn='" + FRectCDS + "'" & VbCrLf
        end if

		if (FRectItemName<>"") then
		    sqlStr = sqlStr + " and s.shopitemname like '%" + CStr(FRectItemName) + "%'" & VbCrLf
		end if

		if FRectShopItemName<>"" then
			sqlStr = sqlStr + " and l.lcitemname like '%" + FRectShopItemName + "%'" & VbCrLf
		end if

		sqlStr = sqlStr + " order by s.itemgubun desc, s.shopitemid desc" & VbCrLf

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
        FTotalCount = FResultCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COffShopOneItem
				FItemList(i).Fitemgubun          = rsget("itemgubun")
				FItemList(i).Fshopitemid         = rsget("shopitemid")
				FItemList(i).Fitemoption         = rsget("itemoption")
				FItemList(i).Fmakerid            = rsget("makerid")
				FItemList(i).Fshopitemname       = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname = db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice      = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash      = rsget("shopsuplycash")
				FItemList(i).Fisusing            = rsget("isusing")
				FItemList(i).Fregdate            = rsget("regdate")
				FItemList(i).FSocName			 = db2html(rsget("socname"))
				FItemList(i).FSocNameKor		 = db2html(rsget("socname_kor"))

				FItemList(i).FOffimgSmall		= rsget("offimgsmall")
				FItemList(i).FimageSmall     	= rsget("smallimage")

				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

				if FItemList(i).FOffimgSmall<>"" then
					FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
				end if

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
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v "
			sqlStr = sqlStr + " on i.itemid=v.itemid"
			sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_const_day_stock s "
			sqlStr = sqlStr + " on i.itemid=s.itemid  and IsNULL(v.itemoption,'0000')=s.itemoption"

		sqlStr = sqlStr + " where i.itemid<>0"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if

		if FRectPriceRow<>"" then
			sqlStr = sqlStr + " and (i.sellcash + IsNull(v.optaddprice,0)) = " + CStr(FRectPriceRow) + " "
		end if

		if FRectItemid<>"" then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemid) + ""
		end if

		if FRectItemName<>"" then
			sqlStr = sqlStr + " and i.itemname like '%" + html2db(FRectItemName) + "%' "
		end if

		if FRectItemgubun<>"" then
			sqlStr = sqlStr + " and i.itemgubun='" + FRectItemgubun + "'"
		end if

		if FRectNoSearchUpcheBeasong="on" then
			sqlStr = sqlStr + " and i.deliverytype in ('1','3','4')"
		end if

		if FRectNoSearchNotusingItem="on" then
			sqlStr = sqlStr + " and i.isusing ='Y'"
		end if

		if FRectNoSearchNotusingItemOption="on" then
			sqlStr = sqlStr + " and ((v.isusing ='Y') or (v.isusing is NULL))"
		end if

		if FRectSellYN = "Y" then
			sqlStr = sqlStr + " and i.sellyn <> 'N' "
			sqlStr = sqlStr + " and i.sellyn <> 'S' "		'// 일시품절, skyer9, 2016-05-10
			sqlStr = sqlStr + " and not (i.limityn = 'Y' and (IsNULL(v.optlimitno,i.limitno) <= IsNULL(v.optlimitsold,i.limitsold))) "
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " '10' as itemgubun, i.itemid, IsNULL(v.itemoption,'0000') as itemoption, "
		sqlStr = sqlStr + " i.makerid, i.itemname, IsNULL(v.optionname,'') as itemoptionname, IsNULL(v.isusing,'Y') as optusing,"
		sqlStr = sqlStr + " i.orgprice, IsNull(v.optaddprice,0) as optaddprice, IsNull(v.optaddbuyprice,0) as optaddbuyprice, i.sellcash, i.buycash, i.smallimage as imgsmall, s.lastrealdate,"
		sqlStr = sqlStr + " IsNull(s.lastrealno,0) as lastrealno, IsNull(s.ipno,0) as ipno, IsNull(s.chulno,0) as chulno,IsNull(s.sellno,0) as sellno, IsNull(s.currno,0) as currno,"
		sqlStr = sqlStr + " IsNull(s.sell7days,0) as sell7days, IsNull(s.jupsu7days,0) as jupsu7days, IsNull(s.offchulgo7days,0) as offchulgo7days, IsNull(s.offconfirmno,0) as offconfirmno,"
		sqlStr = sqlstr + " IsNull(s.offjupno,0) as offjupno, IsNull(s.requireno,0) as requireno, IsNull(s.shortageno,0) as shortageno, IsNull(s.preorderno,0) as preorderno, s.maxsellday, "
		sqlStr = sqlStr + " IsNull(s.ipkumdiv4,0) as ipkumdiv4, IsNull(s.ipkumdiv2,0) as ipkumdiv2,"
		sqlStr = sqlStr + " i.sellyn, i.limityn, IsNULL(v.optlimitno,i.limitno) as limitno, IsNULL(v.optlimitsold,i.limitsold) as limitsold, IsNULL(i.mwdiv,'') as mwdiv"

		sqlStr = sqlStr + " , IsNull(i.ItemCouponYn, 'N') as ItemCouponYn, IsNull(i.CurrItemCouponIdx, '0') as CurrItemCouponIdx, IsNull(i.ItemCouponType, '1') as ItemCouponType, IsNull(i.ItemCouponValue, 0) as ItemCouponValue, IsNull(d.couponbuyprice, 0) as couponbuyprice "

		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v "
			sqlStr = sqlStr + " on i.itemid=v.itemid "
			sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_const_day_stock s"
			sqlStr = sqlStr + " on i.itemid=s.itemid and IsNULL(v.itemoption,'0000')=s.itemoption"
			sqlStr = sqlStr + " left join db_item.dbo.tbl_item_coupon_master m "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and m.itemcouponidx = i.CurrItemCouponIdx "
			sqlStr = sqlStr + " 	and m.itemcouponstartdate <= getdate() "
			sqlStr = sqlStr + " 	and m.itemcouponexpiredate > getdate() "
			sqlStr = sqlStr + " left join db_item.dbo.tbl_item_coupon_detail d "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and m.itemcouponidx = d.itemcouponidx "
			sqlStr = sqlStr + " 	and d.itemid = i.itemid "

		sqlStr = sqlStr + " where i.itemid<>0"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if

		if FRectPriceRow<>"" then
			sqlStr = sqlStr + " and (i.sellcash + IsNull(v.optaddprice,0)) = " + CStr(FRectPriceRow) + " "
		end if

		if FRectItemid<>"" then
			sqlStr = sqlStr + " and i.itemid=" + CStr(FRectItemid) + ""
		end if

		if FRectItemName<>"" then
			sqlStr = sqlStr + " and i.itemname like '%" + html2db(FRectItemName) + "%' "
		end if

		if FRectItemgubun<>"" then
			sqlStr = sqlStr + " and i.itemgubun='" + FRectItemgubun + "'"
		end if

		if FRectNoSearchUpcheBeasong="on" then
			sqlStr = sqlStr + " and i.deliverytype in ('1','3','4')"
		end if

		if FRectNoSearchNotusingItem="on" then
			sqlStr = sqlStr + " and i.isusing ='Y'"
		end if

		if FRectNoSearchNotusingItemOption="on" then
			sqlStr = sqlStr + " and ((v.isusing ='Y') or (v.isusing is NULL))"
		end if

		if FRectSellYN = "Y" then
			sqlStr = sqlStr + " and i.sellyn <> 'N' "
			sqlStr = sqlStr + " and i.sellyn <> 'S' "		'// 일시품절, skyer9, 2016-05-10
			sqlStr = sqlStr + " and not (i.limityn = 'Y' and (IsNULL(v.optlimitno,i.limitno) <= IsNULL(v.optlimitsold,i.limitsold))) "
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
				set FItemList(i) = new COffShopOneItem
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("itemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("itemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("sellcash")
				FItemList(i).Fshopsuplycash     = rsget("buycash")
				FItemList(i).FimageSmall     = rsget("imgsmall")

				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

                FItemList(i).Fsellyn               = rsget("sellyn")
				FItemList(i).Flimityn              = rsget("limityn")
				FItemList(i).Flimitno              = rsget("limitno")
				FItemList(i).Flimitsold            = rsget("limitsold")
				FItemList(i).Flastrealdate	= rsget("lastrealdate")
				FItemList(i).Flastrealno	= rsget("lastrealno")
				FItemList(i).Fipno			= rsget("ipno")
				FItemList(i).Fchulno		= rsget("chulno")
				FItemList(i).Fsellno		= rsget("sellno")
				FItemList(i).Fcurrno		= rsget("currno")
				FItemList(i).Fmwdiv			= rsget("mwdiv")
				FItemList(i).Foptusing = rsget("optusing")
				FItemList(i).Fsell7days     = rsget("sell7days")
				FItemList(i).Fjupsu7days    = rsget("jupsu7days")
				FItemList(i).Foffchulgo7days= rsget("offchulgo7days")
				FItemList(i).Foffconfirmno  = rsget("offconfirmno")
				FItemList(i).Foffjupno      = rsget("offjupno")
				FItemList(i).Frequireno     = rsget("requireno")
				FItemList(i).Fshortageno    = rsget("shortageno")
				FItemList(i).Fpreorderno    = rsget("preorderno")
				FItemList(i).Fmaxsellday	= rsget("maxsellday")
				FItemList(i).Fipkumdiv4		= rsget("ipkumdiv4")
				FItemList(i).Fipkumdiv2		= rsget("ipkumdiv2")

				FItemList(i).FShopItemOrgprice		= rsget("orgprice")
				FItemList(i).Foptaddprice			= rsget("optaddprice")
				FItemList(i).Foptaddbuyprice		= rsget("optaddbuyprice")

				FItemList(i).FItemCouponYn			= rsget("ItemCouponYn")
				FItemList(i).FCurrItemCouponIdx		= rsget("CurrItemCouponIdx")
				FItemList(i).FItemCouponType		= rsget("ItemCouponType")
				FItemList(i).FItemCouponValue		= rsget("ItemCouponValue")
				FItemList(i).Fcouponbuyprice		= rsget("couponbuyprice")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub

	' 오프샵 상품주문 리스트
	public Sub GetOrderItems()
		dim i,sqlStr

		sqlStr = " SELECT a.*	" & vbCrLf
		sqlStr = sqlStr & " , m.statecd masterStateCd 	" & vbCrLf
		sqlStr = sqlStr & " , s.shopitemprice, s.shopsuplycash, s.shopbuyprice 	" & vbCrLf
		sqlStr = sqlStr & " , d.defaultmargin,d.defaultsuplymargin, IsNULL(st.sell7days,0) as sell7days 	" & vbCrLf
		sqlStr = sqlStr & " , IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist, IsNULL(s.offimgsmall,'') as offimgsmall	" & vbCrLf
		sqlStr = sqlStr & " , i.smallimage, i.sellyn, i.limityn, i.limitno, i.limitsold, IsNULL(i.buycash,0) as onlinebuycash, i.deliverytype	" & vbCrLf
		sqlStr = sqlStr & " , IsNULL(o.isusing,'Y') as optusing, IsNULL(o.optlimitno,0) as optlimitno	" & vbCrLf
		sqlStr = sqlStr & " , IsNULL(o.optlimitsold,0) as optlimitsold, IsNULL(o.optsellyn,'Y') optsellyn	" & vbCrLf
		sqlStr = sqlStr & " FROM db_storage.dbo.tbl_order_items a	" & vbCrLf
		sqlStr = sqlStr & " LEFT OUTER JOIN db_summary.dbo.tbl_current_shopstock_summary st 	" & vbCrLf
		sqlStr = sqlStr & " ON a.shopID = st.shopID AND a.itemGubun = st.itemGubun AND a.itemid=st.itemid and a.itemoption=st.itemoption 	" & vbCrLf
		sqlStr = sqlStr & " LEFT OUTER JOIN db_shop.dbo.tbl_shop_designer d	" & vbCrLf
		sqlStr = sqlStr & " ON a.shopID = d.shopID AND a.makerID = d.makerID	" & vbCrLf
		sqlStr = sqlStr & " LEFT OUTER JOIN db_shop.dbo.tbl_shop_item s	" & vbCrLf
		sqlStr = sqlStr & " ON a.itemGubun = s.itemGubun AND a.itemid=s.shopitemid and a.itemoption=s.itemoption 	" & vbCrLf
		sqlStr = sqlStr & " LEFT OUTER JOIN db_item.dbo.tbl_item i 	" & vbCrLf
		sqlStr = sqlStr & " ON a.itemgubun='10' AND a.itemid=i.itemid 	" & vbCrLf
		sqlStr = sqlStr & " LEFT OUTER JOIN db_item.dbo.tbl_item_option o 	" & vbCrLf
		sqlStr = sqlStr & " ON a.itemgubun='10' AND a.itemid=o.itemid AND a.itemoption=o.itemoption 	" & vbCrLf
		sqlStr = sqlStr & " LEFT OUTER JOIN db_storage.dbo.tbl_ordersheet_master m	" & vbCrLf
		sqlStr = sqlStr & " ON a.masterIdx = m.idx	" & vbCrLf
		sqlStr = sqlStr & " WHERE 1=1	" & vbCrLf
		sqlStr = sqlStr & " AND a.shopID = '" & FRectShopid & "'	" & vbCrLf
		sqlStr = sqlStr & " AND a.stateCd IN ('0','1','2')	" & vbCrLf
		sqlStr = sqlStr & " AND isNull(m.stateCd,' ') < '7'	" & vbCrLf
		sqlStr = sqlStr & " ORDER BY a.makerID, a.itemgubun, a.itemid desc, a.itemoption, a.stateCd, a.idx	" & vbCrLf

'		response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopOneItem

				' 상품별주문 속성
				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fcomment			= rsget("comment")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).FconfirmDate		= rsget("confirmDate")
				FItemList(i).FstateCd			= rsget("stateCd")
				FItemList(i).FjumunItemNo		= rsget("jumunItemNo")
				FItemList(i).FmasterStateCd		= rsget("masterStateCd")
				FItemList(i).FsellCash			= rsget("sellCash")
				FItemList(i).FsuplyCash			= rsget("suplyCash")
				FItemList(i).FbuyCash			= rsget("buyCash")
				' 공통
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("itemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("itemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).Fshopbuyprice		= rsget("shopbuyprice")
				FItemList(i).FimageSmall     = rsget("smallimage")

				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

				FItemList(i).Fsellyn               = rsget("sellyn")
                FItemList(i).Flimityn              = rsget("limityn")

                if FItemList(i).Fitemoption="0000" then
				    FItemList(i).Flimitno              = rsget("limitno")
				    FItemList(i).Flimitsold            = rsget("limitsold")
				else
				    if rsget("optsellyn")="N" then FItemList(i).Fsellyn="N"
				    FItemList(i).Flimitno              = rsget("optlimitno")
				    FItemList(i).Flimitsold            = rsget("optlimitsold")
				end if

				FItemList(i).FMakerMargin    = rsget("defaultmargin")
				FItemList(i).FShopMargin    = rsget("defaultsuplymargin")
				FItemList(i).Foptusing      = rsget("optusing")
				FItemList(i).Fonlinebuycash = rsget("onlinebuycash")
				FItemList(i).FOffimgMain	= rsget("offimgmain")
				FItemList(i).FOffimgList	= rsget("offimglist")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				FItemList(i).FdeliveryType  = rsget("deliverytype")
				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub


	' 팝업 가맹점 상품별주문 리스트
    public sub GetShopJumunItemList()
		dim sqlStr,i

		sqlStr = " SELECT a.*	" & vbCrLf
		sqlStr = sqlStr & " , s.CenterMwdiv , IsNULL(i.mwdiv,'') as mwDiv	" & vbCrLf
		sqlStr = sqlStr & " , ls.realstock , ls.ipkumdiv5 , ls.offconfirmno , ls.ipkumdiv4 , ls.ipkumdiv2 , ls.offjupno	" & vbCrLf
		sqlStr = sqlStr & " , IsNULL(st.sell7days,0) as sell7days 	" & vbCrLf
		sqlStr = sqlStr & " , IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist, IsNULL(s.offimgsmall,'') as offimgsmall	" & vbCrLf
		sqlStr = sqlStr & " , i.smallimage, i.sellyn, i.limityn, i.limitno, i.limitsold, IsNULL(i.buycash,0) as onlinebuycash, i.deliverytype	" & vbCrLf
		sqlStr = sqlStr & " , IsNULL(o.isusing,'Y') as optusing, IsNULL(o.optlimitno,0) as optlimitno	" & vbCrLf
		sqlStr = sqlStr & " , IsNULL(o.optlimitsold,0) as optlimitsold, IsNULL(o.optsellyn,'Y') optsellyn	" & vbCrLf
		sqlStr = sqlStr & " FROM db_storage.dbo.tbl_order_items a	" & vbCrLf
		sqlStr = sqlStr & " LEFT OUTER JOIN db_summary.dbo.tbl_current_shopstock_summary st 	" & vbCrLf
		sqlStr = sqlStr & " ON a.shopID = st.shopID AND a.itemGubun = st.itemGubun AND a.itemid=st.itemid and a.itemoption=st.itemoption 	" & vbCrLf
		sqlStr = sqlStr & " LEFT OUTER JOIN db_shop.dbo.tbl_shop_item s	" & vbCrLf
		sqlStr = sqlStr & " ON a.itemGubun = s.itemGubun AND a.itemid=s.shopitemid and a.itemoption=s.itemoption 	" & vbCrLf
        sqlStr = sqlStr + "     left join db_summary.dbo.tbl_current_logisstock_summary ls "
        sqlStr = sqlStr + "     on a.itemgubun=ls.itemgubun and a.itemid=ls.itemid and a.itemoption=ls.itemoption"
		sqlStr = sqlStr & " LEFT OUTER JOIN db_item.dbo.tbl_item i 	" & vbCrLf
		sqlStr = sqlStr & " ON a.itemgubun='10' AND a.itemid=i.itemid 	" & vbCrLf
		sqlStr = sqlStr & " LEFT OUTER JOIN db_item.dbo.tbl_item_option o 	" & vbCrLf
		sqlStr = sqlStr & " ON a.itemgubun='10' AND a.itemid=o.itemid AND a.itemoption=o.itemoption 	" & vbCrLf
		sqlStr = sqlStr & " WHERE 1=1	" & vbCrLf
		sqlStr = sqlStr & " AND a.stateCd = '1'	" & vbCrLf
		sqlStr = sqlStr & " AND a.shopID = '" & FRectShopid & "'	" & vbCrLf

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " AND a.makerid='" & FRectDesigner & "'"
		end if
		sqlStr = sqlStr & " ORDER BY a.itemgubun asc, a.itemid desc, a.itemoption	" & vbCrLf

'		rw sqlStr

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		i=0

		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopOneItem

				FItemList(i).FCenterMwdiv				= rsget("CenterMwdiv")
				FItemList(i).FmwDiv				= rsget("mwDiv")
				' 재고 수량
				FItemList(i).Frealstock				= rsget("realstock")
				FItemList(i).Fipkumdiv5				= rsget("ipkumdiv5")
				FItemList(i).Foffconfirmno			= rsget("offconfirmno")
				FItemList(i).Fipkumdiv4				= rsget("ipkumdiv4")
				FItemList(i).Fipkumdiv2				= rsget("ipkumdiv2")
				FItemList(i).Foffjupno				= rsget("offjupno")
				' 상품별주문 속성
				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fcomment			= rsget("comment")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).FconfirmDate		= rsget("confirmDate")
				FItemList(i).FstateCd			= rsget("stateCd")
				FItemList(i).FjumunItemNo		= rsget("jumunItemNo")
				FItemList(i).FsellCash			= rsget("sellCash")
				FItemList(i).FsuplyCash			= rsget("suplyCash")
				FItemList(i).FbuyCash			= rsget("buyCash")
				' 상품 공통
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("itemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("itemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("sellcash")
				FItemList(i).Fshopsuplycash     = rsget("suplycash")

				FItemList(i).FimageSmall     = rsget("smallimage")
				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if
				FItemList(i).Fsellyn               = rsget("sellyn")
                FItemList(i).Flimityn              = rsget("limityn")

                if FItemList(i).Fitemoption="0000" then
				    FItemList(i).Flimitno              = rsget("limitno")
				    FItemList(i).Flimitsold            = rsget("limitsold")
				else
				    if rsget("optsellyn")="N" then FItemList(i).Fsellyn="N"
				    FItemList(i).Flimitno              = rsget("optlimitno")
				    FItemList(i).Flimitsold            = rsget("optlimitsold")
				end if

				FItemList(i).Foptusing      = rsget("optusing")
				FItemList(i).Fonlinebuycash = rsget("onlinebuycash")
				FItemList(i).FOffimgMain	= rsget("offimgmain")
				FItemList(i).FOffimgList	= rsget("offimglist")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				FItemList(i).FdeliveryType  = rsget("deliverytype")
				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
				FItemList(i).Fshopbuyprice  = rsget("sellcash")
				FItemList(i).FOffsell7days     = rsget("sell7days")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

    end sub

    '//admin/offshop/offitemlist.asp
	public sub GetOffLineNewItemList()
		dim sqlStr,i, sqlSub, sqlSort, sqlsearch2

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch2 = sqlsearch2 & " and isNULL(tp.tplcompanyid,'')<>''"
	            sqlSub = sqlSub & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch2 = sqlsearch2 & " and isNULL(tp.tplcompanyid,'')=''"
	        sqlSub = sqlSub & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		If FRectSorting <> "" Then
			If FRectSorting = "itemregdate" Then
				sqlSort = " ORDER BY s.regdate DESC, s.shopitemid DESC "
			ElseIf FRectSorting = "ipgodate" Then
				sqlSort = " ORDER BY d.firstipgodate DESC, s.shopitemid DESC "
			End If
		Else
			sqlSort = " ORDER BY s.regdate DESC "
		End If

		If FRectShopid <> "" Then
			sqlSub = sqlSub & " AND d.shopid = '" & FRectShopid & "' "
		End If

		If FRectDesigner <> "" Then
			sqlSub = sqlSub & " AND s.makerid = '" & FRectDesigner & "' "
		End If

		If FRectDateSearch <> "" Then
			If FRectDateSearch = "itemregdate" Then
				If FRectSDate <> "" Then
					sqlSub = sqlSub & " AND s.regdate >= '" & FRectSDate & "' "
				End If
				If FRectEDate <> "" Then
					sqlSub = sqlSub & " AND s.regdate <= '" & FRectEDate & "' "
				End If
			ElseIf FRectDateSearch = "ipgodate" Then
				If FRectSDate <> "" Then
					sqlSub = sqlSub & " AND d.firstipgodate >= '" & FRectSDate & "' "
				End If
				If FRectEDate <> "" Then
					sqlSub = sqlSub & " AND d.firstipgodate <= '" & FRectEDate & "' "
				End If
			ElseIf FRectDateSearch = "stockipgodate" Then
				If FRectShopid <> "" Then
					If FRectSDate <> "" Then
						sqlSub = sqlSub & " AND convert(varchar(10),ss.regdate,121) >= '" & FRectSDate & "' "
					End If
					If FRectEDate <> "" Then
						sqlSub = sqlSub & " AND  convert(varchar(10),ss.regdate,121) <= '" & FRectEDate & "' "
					End If
				end if
			else ''2015/04/08추가
			    If FRectSDate <> "" Then
					sqlSub = sqlSub & " AND s.regdate >= '" & FRectSDate & "' "
				End If
				If FRectEDate <> "" Then
					sqlSub = sqlSub & " AND s.regdate <= '" & FRectEDate & "' "
				End If
			End If
		End If

		If FRectItemgubun <> "" Then
			sqlSub = sqlSub & " AND s.itemgubun IN (" & FRectItemgubun & ") "
		End If

		If FRectItemName <> "" Then
			sqlSub = sqlSub & " AND s.shopitemname like '%" & FRectItemName & "%' "
		End If

		If FRectItemId <> "" Then
			sqlSub = sqlSub & " AND s.shopitemid = '" & FRectItemId & "' "
		End If

		If FRectIsusing <> "" Then
			sqlSub = sqlSub & " AND s.isusing = '" & FRectIsusing & "' "
		End If

		if FRectCDL<>"" then
			sqlSub = sqlSub + " and s.catecdl='" + FRectCDL + "'"
		end if
		if FRectCDM<>"" then
			sqlSub = sqlSub + " and s.catecdm='" + FRectCDM + "'"
		end if
		if FRectCDS<>"" then
			sqlSub = sqlSub + " and s.catecdn='" + FRectCDS + "'"
		end if

        '//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

		else
			if FRectStartDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if
		If FRectShopid <> "" Then
			sqlsearch2 = sqlsearch2 & " AND mm.shopid = '" & FRectShopid & "' "
		End If

		sqlStr = " select"
		sqlStr = sqlStr & " count(s.shopitemid) as cnt"
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item as s"
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shop_designer as d"
		sqlStr = sqlStr & " 	on s.makerid = d.makerid"

        sqlStr = sqlStr & "	left join ("
        sqlStr = sqlStr & "		select"
        sqlStr = sqlStr & "		mm.shopid ,dd.itemgubun, dd.itemid, dd.itemoption, sum(dd.itemno) as itemcnt"
        sqlStr = sqlStr & "		,isnull(sum((dd.realsellprice+isnull(dd.addtaxcharge,0))*dd.itemno),0) as sellsum"
        sqlStr = sqlStr & "		,sum(dd.suplyprice*dd.itemno) as suplyprice"
        sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master mm"
        sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail dd"
        sqlStr = sqlStr & "			on mm.idx=dd.masteridx and mm.cancelyn='N' and dd.cancelyn='N'"
		sqlStr = sqlStr & " 	left join db_partner.dbo.tbl_partner tp"
	    sqlStr = sqlStr & "       	on mm.shopid=tp.id "
        sqlStr = sqlStr & "		where 1=1 " &sqlsearch2
        sqlStr = sqlStr & "		group by mm.shopid, dd.itemgubun, dd.itemid, dd.itemoption"
        sqlStr = sqlStr & "	) t"
        sqlStr = sqlStr & "		on d.shopid=t.shopid"
		sqlStr = sqlStr & " 	and s.itemgubun=t.itemgubun"
		sqlStr = sqlStr & " 	and s.shopitemid=t.itemid"
		sqlStr = sqlStr & " 	and s.itemoption=t.itemoption"

		If FRectShopid <> "" Then
			sqlStr = sqlStr & " left join db_summary.dbo.tbl_current_shopstock_summary ss"
			sqlStr = sqlStr & " 	on d.shopid=ss.shopid"
			sqlStr = sqlStr & " 	and s.itemgubun=ss.itemgubun"
			sqlStr = sqlStr & " 	and s.shopitemid=ss.itemid"
			sqlStr = sqlStr & " 	and s.itemoption=ss.itemoption"
		end if

		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " 	on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & " 	on d.shopid=p.id "
		sqlStr = sqlStr & " where 1=1 " & sqlSub

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
		rsget.Close

		if FTotalCount < 1 then exit sub

		sqlStr = " select top " & CStr(FPageSize*FCurrPage) & " "
		sqlStr = sqlStr & " s.itemgubun, s.shopitemid, s.itemoption, s.makerid, s.shopitemname, s.shopitemoptionname"
		sqlStr = sqlStr & " , s.isusing ,i.smallimage, IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist"
		sqlStr = sqlStr & " , IsNULL(s.offimgsmall,'') as offimgsmall, s.regdate, s.updt, d.firstipgodate, d.shopid "

		If FRectShopid <> "" Then
			sqlStr = sqlStr & " ,ss.regdate as stockregdate"
		end if

		sqlStr = sqlStr & " ,IsNULL(t.itemcnt,0) as itemcnt, IsNULL(t.sellsum,0) as sellsum, IsNULL(t.suplyprice,0) as suplyprice"
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item as s"
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shop_designer as d"
		sqlStr = sqlStr & " 	on s.makerid = d.makerid"

        sqlStr = sqlStr & "	left join ("
        sqlStr = sqlStr & "		select"
        sqlStr = sqlStr & "		mm.shopid ,dd.itemgubun, dd.itemid, dd.itemoption, sum(dd.itemno) as itemcnt"
        sqlStr = sqlStr & "		,isnull(sum((dd.realsellprice+isnull(dd.addtaxcharge,0))*dd.itemno),0) as sellsum"
        sqlStr = sqlStr & "		,sum(dd.suplyprice*dd.itemno) as suplyprice"
        sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master mm"
        sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail dd"
        sqlStr = sqlStr & "			on mm.idx=dd.masteridx and mm.cancelyn='N' and dd.cancelyn='N'"
		sqlStr = sqlStr & " 	left join db_partner.dbo.tbl_partner tp"
	    sqlStr = sqlStr & "       	on mm.shopid=tp.id "
        sqlStr = sqlStr & "		where 1=1 " &sqlsearch2
        sqlStr = sqlStr & "		group by mm.shopid, dd.itemgubun, dd.itemid, dd.itemoption"
        sqlStr = sqlStr & "	) t"
        sqlStr = sqlStr & "		on d.shopid=t.shopid"
		sqlStr = sqlStr & " 	and s.itemgubun=t.itemgubun"
		sqlStr = sqlStr & " 	and s.shopitemid=t.itemid"
		sqlStr = sqlStr & " 	and s.itemoption=t.itemoption"

		If FRectShopid <> "" Then
			sqlStr = sqlStr & " left join db_summary.dbo.tbl_current_shopstock_summary ss"
			sqlStr = sqlStr & " 	on d.shopid=ss.shopid"
			sqlStr = sqlStr & " 	and s.itemgubun=ss.itemgubun"
			sqlStr = sqlStr & " 	and s.shopitemid=ss.itemid"
			sqlStr = sqlStr & " 	and s.itemoption=ss.itemoption"
		end if

		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " 	on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & " 	on d.shopid=p.id "
		sqlStr = sqlStr & " where 1=1 " & sqlSub
		sqlStr = sqlStr & sqlSort

	'response.write sqlStr &"<Br>"
  ' response.end

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CDbl(FTotalCount\FPageSize)
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
				set FItemList(i) = new COffShopOneItem

				FItemList(i).fitemcnt			= rsget("itemcnt")
				FItemList(i).fsellsum			= rsget("sellsum")
				FItemList(i).fsuplyprice			= rsget("suplyprice")

				If FRectShopid <> "" Then
					FItemList(i).fstockregdate			= rsget("stockregdate")
				end if

				FItemList(i).FShopID			= rsget("shopid")
				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).Fshopitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				'FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				'FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")

				FItemList(i).FimageSmall     = rsget("smallimage")
				if FItemList(i).FimageSmall<>"" then
					FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				end if

				'FItemList(i).Fsellyn			= rsget("sellyn")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Ffirstipgodate		= rsget("firstipgodate")
				FItemList(i).Fupdt				= rsget("updt")
				FItemList(i).FOffimgMain		= rsget("offimgmain")
				FItemList(i).FOffimgList		= rsget("offimglist")
				FItemList(i).FOffimgSmall		= rsget("offimgsmall")

				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end sub

    '//admin/offshop/offitemlist_xls.asp
	public sub GetOffLineNewItemList_xls()
		dim sqlStr,i, sqlSub, sqlSort, sqlsearch2

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch2 = sqlsearch2 & " and isNULL(tp.tplcompanyid,'')<>''"
	            sqlSub = sqlSub & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch2 = sqlsearch2 & " and isNULL(tp.tplcompanyid,'')=''"
	        sqlSub = sqlSub & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		If FRectSorting <> "" Then
			If FRectSorting = "itemregdate" Then
				sqlSort = " ORDER BY s.regdate DESC, s.shopitemid DESC "
			ElseIf FRectSorting = "ipgodate" Then
				sqlSort = " ORDER BY d.firstipgodate DESC, s.shopitemid DESC "
			End If
		Else
			sqlSort = " ORDER BY s.regdate DESC "
		End If

		If FRectShopid <> "" Then
			sqlSub = sqlSub & " AND d.shopid = '" & FRectShopid & "' "
		End If

		If FRectDesigner <> "" Then
			sqlSub = sqlSub & " AND s.makerid = '" & FRectDesigner & "' "
		End If

		If FRectDateSearch <> "" Then
			If FRectDateSearch = "itemregdate" Then
				If FRectSDate <> "" Then
					sqlSub = sqlSub & " AND s.regdate >= '" & FRectSDate & "' "
				End If
				If FRectEDate <> "" Then
					sqlSub = sqlSub & " AND s.regdate <= '" & FRectEDate & "' "
				End If
			ElseIf FRectDateSearch = "ipgodate" Then
				If FRectSDate <> "" Then
					sqlSub = sqlSub & " AND d.firstipgodate >= '" & FRectSDate & "' "
				End If
				If FRectEDate <> "" Then
					sqlSub = sqlSub & " AND d.firstipgodate <= '" & FRectEDate & "' "
				End If
			ElseIf FRectDateSearch = "stockipgodate" Then
				If FRectShopid <> "" Then
					If FRectSDate <> "" Then
						sqlSub = sqlSub & " AND convert(varchar(10),ss.regdate,121) >= '" & FRectSDate & "' "
					End If
					If FRectEDate <> "" Then
						sqlSub = sqlSub & " AND  convert(varchar(10),ss.regdate,121) <= '" & FRectEDate & "' "
					End If
				end if
			else ''2015/04/08추가
			    If FRectSDate <> "" Then
					sqlSub = sqlSub & " AND s.regdate >= '" & FRectSDate & "' "
				End If
				If FRectEDate <> "" Then
					sqlSub = sqlSub & " AND s.regdate <= '" & FRectEDate & "' "
				End If
			End If
		End If

		If FRectItemgubun <> "" Then
			sqlSub = sqlSub & " AND s.itemgubun IN (" & FRectItemgubun & ") "
		End If

		If FRectItemName <> "" Then
			sqlSub = sqlSub & " AND s.shopitemname like '%" & FRectItemName & "%' "
		End If

		If FRectItemId <> "" Then
			sqlSub = sqlSub & " AND s.shopitemid = '" & FRectItemId & "' "
		End If

		If FRectIsusing <> "" Then
			sqlSub = sqlSub & " AND s.isusing = '" & FRectIsusing & "' "
		End If

		if FRectCDL<>"" then
			sqlSub = sqlSub + " and s.catecdl='" + FRectCDL + "'"
		end if
		if FRectCDM<>"" then
			sqlSub = sqlSub + " and s.catecdm='" + FRectCDM + "'"
		end if
		if FRectCDS<>"" then
			sqlSub = sqlSub + " and s.catecdn='" + FRectCDS + "'"
		end if

        '//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

		else
			if FRectStartDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch2 = sqlsearch2 & " 	and mm.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if
		If FRectShopid <> "" Then
			sqlsearch2 = sqlsearch2 & " AND mm.shopid = '" & FRectShopid & "' "
		End If

		sqlStr = " select"
		sqlStr = sqlStr & " count(s.shopitemid) as cnt"
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item as s"
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shop_designer as d"
		sqlStr = sqlStr & " 	on s.makerid = d.makerid"

        sqlStr = sqlStr & "	left join ("
        sqlStr = sqlStr & "		select"
        sqlStr = sqlStr & "		mm.shopid ,dd.itemgubun, dd.itemid, dd.itemoption, sum(dd.itemno) as itemcnt"
        sqlStr = sqlStr & "		,isnull(sum((dd.realsellprice+isnull(dd.addtaxcharge,0))*dd.itemno),0) as sellsum"
        sqlStr = sqlStr & "		,sum(dd.suplyprice*dd.itemno) as suplyprice"
        sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master mm"
        sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail dd"
        sqlStr = sqlStr & "			on mm.idx=dd.masteridx and mm.cancelyn='N' and dd.cancelyn='N'"
		sqlStr = sqlStr & " 	left join db_partner.dbo.tbl_partner tp"
	    sqlStr = sqlStr & "       	on mm.shopid=tp.id "
        sqlStr = sqlStr & "		where 1=1 " &sqlsearch2
        sqlStr = sqlStr & "		group by mm.shopid, dd.itemgubun, dd.itemid, dd.itemoption"
        sqlStr = sqlStr & "	) t"
        sqlStr = sqlStr & "		on d.shopid=t.shopid"
		sqlStr = sqlStr & " 	and s.itemgubun=t.itemgubun"
		sqlStr = sqlStr & " 	and s.shopitemid=t.itemid"
		sqlStr = sqlStr & " 	and s.itemoption=t.itemoption"

		If FRectShopid <> "" Then
			sqlStr = sqlStr & " left join db_summary.dbo.tbl_current_shopstock_summary ss"
			sqlStr = sqlStr & " 	on d.shopid=ss.shopid"
			sqlStr = sqlStr & " 	and s.itemgubun=ss.itemgubun"
			sqlStr = sqlStr & " 	and s.shopitemid=ss.itemid"
			sqlStr = sqlStr & " 	and s.itemoption=ss.itemoption"
		end if

		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " 	on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & " 	on d.shopid=p.id "
		sqlStr = sqlStr & " where 1=1 " & sqlSub

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
		rsget.Close

		if FTotalCount < 1 then exit sub

		sqlStr = " select top " & CStr(FPageSize*FCurrPage) & " "
		sqlStr = sqlStr & " s.itemgubun, s.shopitemid, s.itemoption, s.makerid, s.shopitemname, s.shopitemoptionname"
		sqlStr = sqlStr & " , s.isusing ,i.smallimage, IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist"
		sqlStr = sqlStr & " , IsNULL(s.offimgsmall,'') as offimgsmall, s.regdate, s.updt, d.firstipgodate, d.shopid "

		If FRectShopid <> "" Then
			sqlStr = sqlStr & " ,ss.regdate as stockregdate"
		end if

		sqlStr = sqlStr & " ,IsNULL(t.itemcnt,0) as itemcnt, IsNULL(t.sellsum,0) as sellsum, IsNULL(t.suplyprice,0) as suplyprice"
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item as s"
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shop_designer as d"
		sqlStr = sqlStr & " 	on s.makerid = d.makerid"

        sqlStr = sqlStr & "	left join ("
        sqlStr = sqlStr & "		select"
        sqlStr = sqlStr & "		mm.shopid ,dd.itemgubun, dd.itemid, dd.itemoption, sum(dd.itemno) as itemcnt"
        sqlStr = sqlStr & "		,isnull(sum((dd.realsellprice+isnull(dd.addtaxcharge,0))*dd.itemno),0) as sellsum"
        sqlStr = sqlStr & "		,sum(dd.suplyprice*dd.itemno) as suplyprice"
        sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master mm"
        sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail dd"
        sqlStr = sqlStr & "			on mm.idx=dd.masteridx and mm.cancelyn='N' and dd.cancelyn='N'"
		sqlStr = sqlStr & " 	left join db_partner.dbo.tbl_partner tp"
	    sqlStr = sqlStr & "       	on mm.shopid=tp.id "
        sqlStr = sqlStr & "		where 1=1 " &sqlsearch2
        sqlStr = sqlStr & "		group by mm.shopid, dd.itemgubun, dd.itemid, dd.itemoption"
        sqlStr = sqlStr & "	) t"
        sqlStr = sqlStr & "		on d.shopid=t.shopid"
		sqlStr = sqlStr & " 	and s.itemgubun=t.itemgubun"
		sqlStr = sqlStr & " 	and s.shopitemid=t.itemid"
		sqlStr = sqlStr & " 	and s.itemoption=t.itemoption"

		If FRectShopid <> "" Then
			sqlStr = sqlStr & " left join db_summary.dbo.tbl_current_shopstock_summary ss"
			sqlStr = sqlStr & " 	on d.shopid=ss.shopid"
			sqlStr = sqlStr & " 	and s.itemgubun=ss.itemgubun"
			sqlStr = sqlStr & " 	and s.shopitemid=ss.itemid"
			sqlStr = sqlStr & " 	and s.itemoption=ss.itemoption"
		end if

		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " 	on s.itemgubun='10' and s.shopitemid=i.itemid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & " 	on d.shopid=p.id "
		sqlStr = sqlStr & " where 1=1 " & sqlSub
		sqlStr = sqlStr & sqlSort

		'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount

		if  not rsget.EOF  then
			frectgetrows=rsget.getrows()
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
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

''상품 기본 센터 매입 구분
function GetDefaultItemMwdivByBrand(imakerid)
    dim sqlStr
    GetDefaultItemMwdivByBrand = "W"

    sqlStr = "select top 1 IsNULL(defaultCenterMwdiv,'') as defaultCenterMwdiv from db_shop.dbo.tbl_shop_designer"
    sqlStr = sqlStr & " where shopid='streetshop000'"
    sqlStr = sqlStr & " and makerid='" & imakerid & "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        if (rsget("defaultCenterMwdiv")="M") then
            GetDefaultItemMwdivByBrand = "M"
        end if
    end if
    rsget.close
end function

''브랜드 계약구분.
function GetShopBrandContract(shopid,makerid)
    dim sqlStr
    GetShopBrandContract =""

    sqlStr = "select top 1 IsNULL(comm_cd,'') as comm_cd from db_shop.dbo.tbl_shop_designer"
    sqlStr = sqlStr + " where shopid='" + shopid + "'"
    sqlStr = sqlStr + " and makerid='" + makerid + "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        GetShopBrandContract = rsget("comm_cd")
    end if
    rsget.close
end function

''업체위탁, 매장매입 계약이 있는지
function fnIsDirectIpchulContractExistsBrand(makerid)
    dim sqlStr
    fnIsDirectIpchulContractExistsBrand = FALSE

    sqlStr = "select count(*) as CNT from db_shop.dbo.tbl_shop_designer"
    sqlStr = sqlStr + " where makerid='" + makerid + "'"
    sqlStr = sqlStr + " and comm_cd in ('B012','B022')"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        fnIsDirectIpchulContractExistsBrand = rsget("CNT")>0
    end if
    rsget.close

end function
%>
