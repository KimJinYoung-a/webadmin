<%
Const CMAXMARGIN = 15
Const CLIMIT_SOLDOUT_NO = 5
'############################## XML 생성 파일 관련 Class ##############################
class cKaffaOneItem
	public FItemID
	public FItemName
    public FMakerid

	public Fcate_large
	public Fcate_mid
	public Fcate_small
	public Fnmlarge
	public FnmMid
	public FnmSmall
	public FKaffacate1
	public FKaffacate2
	public FKaffacate3

	public Fsourcearea
	public Fitemweight
	public FMakerName
    public FBrandName

	public FSellCash
	public Forgsellcash
	public FSuplyCash
	public Fkeywords

	public FListImage
	public FSmallImage
	public FBasicImage
	public Fmainimage
	public Fmainimage2
	public Ficon1Image
	public Ficon2Image

    public FInfoImage

	public FSellyn
	public FDispyn

	public FDesigner

	public FRegdate

	public FLinkCode
	public FOptionTypeName
	public FItemOption
	public FItemOptionName
	public FItemOptionGubunName

	public FItemContent
	public Fordercomment

	public FUpDate

	public Flimityn
	public Flimitno
	public Flimitsold
	public Fstockqty

	public FSailDispNo
    public Fvatinclude

	public FTTLCode
    public FKaffacategory


    public Fitemsize
    public Fitemsource

    public Foptsellyn
    public Foptlimityn
    public Foptlimitno
    public Foptlimitsold
    public Foptaddprice

    public FLastUpdate
    public FSellEndDate

    public FInfoImage1
    public FInfoImage2
    public FInfoImage3
    public FInfoImage4

    public FDetailImage
    public FDetailImage1
    public FDetailImage2
    public FDetailImage3
    public FDetailImage4

    public Fitemdiv
    public Fkaffamakerid
    public Faddimage
    public Fkaffa
	public Fitemgubun
	public Fbuycash
	public Forgprice
    public Forgsuplycash
    public Fsailprice
    public Fsailsuplycash
    public Fmileage
    public Fdanjongyn
    public Fsailyn
    public Fisusing
    public Fisextusing
    public Fmwdiv
    public Fspecialuseritem
    public Fdeliverytype
    public Fdeliverarea
    public Fdeliverfixday
    public Fismobileitem
    public Fpojangok
    public Fevalcnt
    public Foptioncnt
    public Fitemrackcode
    public Fupchemanagecode
    public FReIpgodate
    public Flistimage120
    public Fitemcouponyn
    public Fcurritemcouponidx
    public Fitemcoupontype
    public Fitemcouponvalue
	public FdeliverOverseas
	''public Fcouponbuyprice
	public FRegUser
	public FKaffaUseYN
	public FKaffaregdate
	public FKaffalastupdate
	public FKaffaprice
	public FKaffasellyn
	public FKaffagoodno
	public FregedOptCnt
	public FrctSellCNT
	Public FaccFailCnt
	Public FlastErrStr
	public FkaffaWeight
	public FkaffaIsDisplay
    Public FdefaultfreeBeasongLimit

    Public FKaffaDiscountPrice
    Public FkaffaDiscount_Begin_dateTime
    Public FkaffaDiscount_End_dateTime
    Public FmultiPrice
    Public FmaydiscountPrice

'    ''우리쪽 기준 할인 기간
'    public function getMayDiscountDateStr()
'        getMayDiscountDateStr = ""
'    end function

    public function getForeignMultipleStr()
        if (Forgprice<>0) then
            if not IsNULL(FmultiPrice) then
                getForeignMultipleStr = CLNG(FmultiPrice/Forgprice*100)/100
            end if
        end if
    end function

    ''kaffa 기준 할인 기간
    public function getDiscountDateStr()
        getDiscountDateStr =""
        if isNULL(FkaffaDiscount_Begin_dateTime) then Exit function

        getDiscountDateStr = FkaffaDiscount_Begin_dateTime&" ~ "&FkaffaDiscount_End_dateTime
    end function

    public function IsKaffaSiteDiscountSale
        IsKaffaSiteDiscountSale = false
        if isNULL(FkaffaDiscount_Begin_dateTime) then Exit function

        IsKaffaSiteDiscountSale = FKaffaprice>FKaffaDiscountPrice and FKaffaDiscountPrice>0
        IsKaffaSiteDiscountSale = IsKaffaSiteDiscountSale and (Cdate(FkaffaDiscount_Begin_dateTime)<=date() and Cdate(FkaffaDiscount_End_dateTime)>=date())
    end function

    public function getdeliCode()
        if (Fitemdiv<>"01") then getdeliCode="7"
    end function

    public function GetSellEndDateStr()
        GetSellEndDateStr = "99991231"

        if IsNULL(FSellEndDate) then Exit function

        FSellEndDate = Replace(Left(CStr(FSellEndDate),10),"-","")
    end function

    public function GetRealSellprice()
        if (Foptaddprice>0) then
            GetRealSellprice = FSellcash + Foptaddprice
        else
            GetRealSellprice = FSellcash
        end if
    end function

    public function IsOptionSoldOut()

        IsOptionSoldOut = false
        if (FItemOption="0000") then Exit function

        ''옵션추가 금액이 있는것은 뺌
        IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno-Foptlimitsold<CLIMIT_SOLDOUT_NO)) or (Foptaddprice>0)


    end function

    public function IsSoldOut()

        IsSoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<=CLIMIT_SOLDOUT_NO))
    end function

    public function getLimitDispStr()
        getLimitDispStr = ""
        if (FLimitYn<>"Y") then Exit function

        if (FLimitNo-FLimitSold)<1 then
            getLimitDispStr = "한정 0"
        elseif (FLimitNo-FLimitSold-CLIMIT_SOLDOUT_NO)<1 then
            getLimitDispStr = "한정 "&FLimitNo-FLimitSold&"-"&CLIMIT_SOLDOUT_NO
        else
            getLimitDispStr = "한정 "&FLimitNo-FLimitSold&""
        end if

        getLimitDispStr="<br><font color='blue'>"&getLimitDispStr&"</font>"
    end function

	public function getSellStrNo()
		if (FDispyn="N") or (FSellyn="N") then
			getSellStrNo = "3"
		elseif ((FLimitYn="Y") and (FLimitNo-FLimitSold<1)) then
			getSellStrNo = "2"
		else
			getSellStrNo = "1"
		end if
	end function

	public function getkeywords()
		getkeywords = Fkeywords
	end function

	public function get400Image()
		get400Image = ""

		if IsNULL(FBasicImage) or (FBasicImage="") then Exit function

		get400Image = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemID) + "/" + FBasicImage
	end function

    public function getItemPreInfodataHTML()
        dim reStr

        if Not IsNULL(Fordercomment) then
            if Fordercomment<>"" then
                reStr = "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
            end if
        end if

        if Fitemsize<>"" then
            reStr = reStr & "- 사이즈 : " & Fitemsize & "<br>"
        end if

        if Fitemsource<>"" then
            reStr = reStr & "- 재료 : " &  Fitemsource & "<br>"
        end if

        getItemPreInfodataHTML = reStr
    end function

    public function getItemInfoImageHTML()
        dim splited, i, cnt, oneimageName

        getItemInfoImageHTML = ""

        if FDetailImage <> "" then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(FItemID) + "/" + FDetailImage + ">"
        end if

        if FDetailImage1 <> "" then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FDetailImage1 + ">"
        end if

        if FDetailImage2 <> "" then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FDetailImage2 + ">"
        end if

        if FDetailImage3 <> "" then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FDetailImage3 + ">"
        end if

        if FDetailImage4 <> "" then
            getItemInfoImageHTML = getItemInfoImageHTML + "<br><img src=http://webimage.10x10.co.kr/item/contentsimage/" + GetImageSubFolderByItemid(FItemID) + "/" + FDetailImage4 + ">"
        end if


        exit function

    end function

	public function get160Image()
		get160Image = ""

		if IsNULL(Ficon1Image) or (Ficon1Image = "") then Exit function

		get160Image = "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(FItemID) + "/" + Ficon1Image

	end function

	public function get85Image()
		get85Image = ""

		if IsNULL(FListImage) or (FListImage = "") then Exit function

		get85Image = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemID) + "/" + FListImage
	end function

	public function get60Image()
		get60Image = ""

		if IsNULL(FSmallImage) or (FSmallImage = "") then Exit function

		get60Image = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemID) + "/" + FSmallImage
	end function

	public function getAsDeliverInfo()
		getAsDeliverInfo = Fordercomment
	end function

	public function getItemContent()
		getItemContent = FItemContent
	end function

	'// 쿠폰 적용가
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

	'// 쿠폰 할인가
	public Function GetCouponDiscountPrice() '?
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

    end Function

	'// 상품 쿠폰 여부
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// 세일포함 실제가격
	public Function getRealPrice() '!

		getRealPrice = FSellCash


		'if (IsSpecialUserItem()) then
		'	getRealPrice = getSpecialShopItemPrice(FSellCash)
		'end if
	end Function

    public function getDeliverytypeName
        if (Fdeliverytype="9") then
            getDeliverytypeName = "<font color='blue'>[조건]</font>" ''"&FormatNumber(FdefaultfreeBeasongLimit,0)&"
        elseif (Fdeliverytype="7") then
            getDeliverytypeName = "<font color='red'>[업체착불]</font>"
        elseif (Fdeliverytype="2") then
            getDeliverytypeName = "<font color='blue'>[업체]</font>"
        else
            getDeliverytypeName = ""
        end if
    end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class cKaffaItem
	Public FItemList()
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	Public FRectDesigner
	Public FRectNoRegNate
	Public FBufArr
	Public FBufOptArr
	Public FBufSellcashArr
	Public FRectStartItemID
	Public FRectNotMatchCategory
	Public FRectCate_large
	Public FCateCount
	Public FRectCate1
	Public FRectCate2
	Public FRectCate3
	Public FRectUseYN
	Public FRectKaffaUseYN
	Public FRectItemID

	Public FRectMakerid
	Public FRectKAFFAPrdNo
	Public FRectItemName
	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FRectEventid
	Public FRectOrdType
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectTenSailyn
	Public FRectKaffaBaseSailyn
	Public FRectKaffaSailyn
	Public FRectMinusMargin
	Public FRectonlyValidMargin
	Public FRectFailCntExists
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptExists
	Public FRectKAFFASell10x10Soldout
	Public FRectexpensive10x10
	Public FRectdiffPrc
	Public FRectdiffMultiPrc
	Public FRectdiffWeight
	Public FRectExtSellYn
	Public FRectExtDispYn
    Public FRectMWDiv

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub GetKaffaCategoryMachingList()
        dim sqlStr,i

        sqlStr = "select  "
        sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, c.nmlarge, c.nmmid, c.nmsmall, p.kaffacate1, p.kaffacate2, p.kaffacate3 "
        sqlStr = sqlStr + "  from [db_item].[dbo].tbl_kaffa_reg_item as k "
        sqlStr = sqlStr + "     inner join [db_item].[dbo].tbl_item as i on k.itemid = i.itemid "
        sqlStr = sqlStr + "     left join [db_item].[dbo].vw_category as c on i.cate_large=c.cdlarge and i.cate_mid=c.cdmid and i.cate_small=c.cdsmall"
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_kaffa_category_mapping as p on i.cate_large = p.tencdl and i.cate_mid = p.tencdm and i.cate_small = p.tencds"
        'sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_kaffa_category as c on i.cate_large = p.tencdl and i.cate_mid = p.tencdm and i.cate_small = p.tencds"
        sqlStr = sqlStr + " where 1=1"
        if FRectCate_large <> "" Then
        	sqlStr = sqlStr + " and i.cate_large = '" & FRectCate_large & "'"
        end if
        if (FRectNotMatchCategory = "on") then
            sqlStr = sqlStr + " and p.kaffacate1 is NULL"
        end if
        sqlStr = sqlStr + " group by i.cate_large, i.cate_mid, i.cate_small, c.nmlarge, c.nmmid, c.nmsmall, p.kaffacate1, p.kaffacate2, p.kaffacate3"
        sqlStr = sqlStr + " order by i.cate_large, i.cate_mid, i.cate_small"

        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

        FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new cKaffaOneItem
				FItemList(i).FCate_Large             = rsget("Cate_Large")
                FItemList(i).FCate_Mid               = rsget("Cate_Mid")
                FItemList(i).FCate_Small             = rsget("Cate_Small")
                'FItemList(i).FItemCnt                = rsget("ItemCnt")
                FItemList(i).Fnmlarge                = db2Html(rsget("nmlarge"))
                FItemList(i).FnmMid                  = db2Html(rsget("nmMid"))
                FItemList(i).FnmSmall               = db2Html(rsget("nmSmall"))
                FItemList(i).FKaffacate1			= rsget("kaffacate1")
                FItemList(i).FKaffacate2			= rsget("kaffacate2")
                FItemList(i).FKaffacate3			= rsget("kaffacate3")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
    end Sub

	public function GetKaffaCate2List()
		dim sqlStr,i
		sqlStr = "select cate1, cate2, cate_nm from [db_item].[dbo].[tbl_kaffa_category]"
		sqlStr = sqlStr + " where cate2 <> 0 and cate3 = 0"
		If FRectCate1 <> "" Then
			sqlStr = sqlStr + " and cate1 = '" & FRectCate1 & "' "
		End IF
		sqlStr = sqlStr + " order by cate2 asc"
		rsget.Open sqlStr, dbget, 1
		if not rsget.eof then
			FTotalCount = rsget.RecordCount
			GetKaffaCate2List = rsget.getRows()
		else
			FTotalCount = 0
		end if
		rsget.Close
	end function

	public function GetKaffaCate3List()
		dim sqlStr,i
		sqlStr = "select cate1, cate2, cate3, cate_nm from [db_item].[dbo].[tbl_kaffa_category]"
		sqlStr = sqlStr + " where cate2 <> 0 and cate3 <> 0"
		If FRectCate1 <> "" Then
			sqlStr = sqlStr + " and cate1 = '" & FRectCate1 & "' "
		End IF
		If FRectCate2 <> "" Then
			sqlStr = sqlStr + " and cate2 = '" & FRectCate2 & "' "
		End IF
		sqlStr = sqlStr + " order by cate3 asc"
		rsget.Open sqlStr, dbget, 1
		if not rsget.eof then
			FTotalCount = rsget.RecordCount
			GetKaffaCate3List = rsget.getRows()
		else
			FTotalCount = 0
		end if
		rsget.Close
	end function

	public sub GetAllKaffaItemList()
		dim sqlStr,i
		sqlStr = "select  i.itemid, mi.itemname, i.makerid, i.sellcash, i.orgprice, IsNULL(c.sourcearea,'') as sourcearea," + vbcrlf
		sqlStr = sqlStr + " i.cate_large, i.cate_mid, i.cate_small, mi.makername, i.brandname," + vbcrlf
		sqlStr = sqlStr + " i.sellyn, i.limityn, (i.limitno-"&CLIMIT_SOLDOUT_NO&") as limitno, i.limitsold, c.keywords, c.ordercomment, mi.itemcontent, "
		sqlStr = sqlStr + " i.listimage, i.icon1image, i.basicimage, i.mainimage, i.mainimage2, mi.sourcearea, i.vatinclude, c.keywords, i.sellenddate, i.itemdiv, "
		sqlStr = sqlStr + " convert(varchar(19),s.regdate,120) as regdate, IsNULL(mi.itemsize,'') as itemsize, IsNULL(mi.itemsource,'') as itemsource," + vbcrlf
		sqlStr = sqlStr + " convert(varchar(19),s.lastupdate,120) as lastupdate,  c.usinghtml, i.itemWeight, s.kaffamakerid, m.kaffacate1, m.kaffacate2, m.kaffacate3, "
		sqlStr = sqlStr + " STUFF(( " + vbcrlf
		sqlStr = sqlStr + " 	SELECT ',' + 'http://webimage.10x10.co.kr/image/add' + convert(char(1),ai.GUBUN) + '/' + cast(" & GetImageSubFolderByItemid(FRectItemID) & " as varchar(10)) + '/' + ai.ADDIMAGE_400 " + vbcrlf
		sqlStr = sqlStr + " 	FROM db_item.dbo.tbl_item_addImage as ai " + vbcrlf
		sqlStr = sqlStr + " 	WHERE itemid = s.itemid AND ai.IMGTYPE = 0 " + vbcrlf
		sqlStr = sqlStr + " FOR XML PATH('') " + vbcrlf
		sqlStr = sqlStr + " ), 1, 1, '') AS addimage, " + vbcrlf
		sqlStr = sqlStr + " STUFF(( " + vbcrlf
		sqlStr = sqlStr + " 	SELECT '|option|' + cast(o.itemoption as varchar(4)) + '|||' + o.optionname + '|||' + cast(o.optaddprice as varchar(10)) + '|||' + " + vbcrlf
		sqlStr = sqlStr + " 		case when o.optionTypeName = '' then '선택' else o.optionTypeName + ' 선택' end " + vbcrlf
		sqlStr = sqlStr + " 		+ '|||' + o.optlimityn + '|||' + cast((o.optlimitno-"&CLIMIT_SOLDOUT_NO&") as varchar(10)) + '|||' + cast(o.optlimitsold as varchar(10)) " + vbcrlf
		sqlStr = sqlStr + " 	FROM [db_item].[dbo].tbl_item_option as o " + vbcrlf
		sqlStr = sqlStr + " 	WHERE o.itemid = s.itemid and o.isusing='Y' " + vbcrlf
		sqlStr = sqlStr + " 	order by o.itemoption asc " + vbcrlf
		sqlStr = sqlStr + " FOR XML PATH('') " + vbcrlf
		sqlStr = sqlStr + " ), 1, 1, '') AS itemoption " + vbcrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_Kaffa_reg_item s " + vbcrlf
		sqlStr = sqlStr + " 	inner join [db_item].[dbo].tbl_item i on s.itemid=i.itemid " + vbcrlf
		sqlStr = sqlStr + "     inner join [db_item].[dbo].tbl_item_multiLang as mi on s.itemid = mi.itemid and mi.countryCd = 'kr' "
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_contents c on i.itemid=c.itemid"
		sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_Kaffa_category_mapping m " + vbcrlf
	    sqlStr = sqlStr + "     on i.cate_large=m.tencdl " + vbcrlf
	    sqlStr = sqlStr + "     and i.cate_mid=m.tencdm " + vbcrlf
	    sqlStr = sqlStr + " 	and i.cate_small=m.tencds " + vbcrlf
		sqlStr = sqlStr + " where 1=1 "

		If FRectItemID <> "" Then
			sqlStr = sqlStr + " and s.itemid = '" & FRectItemID & "' "
		End If

		sqlStr = sqlStr + " and i.basicimage is not null and i.deliverytype in(1,4) and i.itemid not in(179229, 179233) and i.makerid <> 'urbanshop' "
		sqlStr = sqlStr + " and i.itemdiv<50"
		sqlStr = sqlStr + " and i.cate_large<>''"
		sqlStr = sqlStr + " and i.cate_large<>'999'"
		sqlStr = sqlStr + " and i.sellcash>0"
		'sqlStr = sqlStr + " and m.category is Not NULL"

		sqlStr = sqlStr + " order by i.itemid desc"
''rw sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new cKaffaOneItem
				FItemList(i).Fitemid      = rsget("itemid")
				FItemList(i).Fitemname 	  = db2html(rsget("itemname"))
				FItemList(i).FMakerid     = rsget("makerid")
				FItemList(i).Fsellcash    = rsget("sellcash")
				FItemList(i).Forgsellcash = rsget("orgprice")
				FItemList(i).Fsourcearea  = db2html(rsget("sourcearea"))
				FItemList(i).FRegdate     = rsget("regdate")
				FItemList(i).Flimityn  = rsget("limityn")
				FItemList(i).Flimitno  = rsget("limitno")
				FItemList(i).Flimitsold  = rsget("limitsold")

				If rsget("limityn") = "Y" Then
					FItemList(i).Fstockqty = rsget("limitno") - rsget("limitsold")
					If FItemList(i).Fstockqty < 1 Then
						FItemList(i).Fsellyn  = "N"
					Else
						FItemList(i).Fsellyn  = rsget("sellyn")
					End If
				Else
					FItemList(i).Fstockqty = "999"
					FItemList(i).Fsellyn  = rsget("sellyn")
				End If

				FItemList(i).Fcate_large = rsget("cate_large")
				FItemList(i).Fcate_mid = rsget("cate_mid")
				FItemList(i).Fcate_small = rsget("cate_small")

				FItemList(i).FMakerName = db2html(rsget("makername"))
				FItemList(i).FBrandName = db2html(rsget("brandname"))

				if (IsNULL(FItemList(i).FMakerName) or (FItemList(i).FMakerName="")) then
				    FItemList(i).FMakerName = FItemList(i).FBrandName
				end if

				'FItemList(i).Fkeywords = db2html(rsget("keywords"))

				If isNull(rsget("itemoption")) Then
					FItemList(i).Fitemoption  = "0000"
				Else
					FItemList(i).Fitemoption  = rsget("itemoption")
				End If

				FItemList(i).Fbasicimage  = "http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("basicimage")
				'FItemList(i).Fmainimage   = rsget("mainimage")
				FItemList(i).Fmainimage 	= "http://webimage.10x10.co.kr/image/main/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("mainimage")
				FItemList(i).Fmainimage2	= "http://webimage.10x10.co.kr/image/main2/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("mainimage2")
				FItemList(i).FDetailImage   = rsget("mainimage")

				FItemList(i).Flistimage  = "http://webimage.10x10.co.kr/image/list/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("listimage")
				'FItemList(i).Fsmallimage  = rsget("smallimage")
				FItemList(i).Ficon1image  = "http://webimage.10x10.co.kr/image/icon1/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("icon1image")
				'FItemList(i).Ficon2image  = rsget("icon2image")

				If isNull(rsget("addimage")) Then
					FItemList(i).Faddimage = ""
				Else
					FItemList(i).Faddimage = rsget("addimage")
				End If

				FItemList(i).Fsourcearea  = db2html(rsget("sourcearea"))
				FItemList(i).Fitemweight  = rsget("itemWeight")
				FItemList(i).Fkeywords  = Replace(db2html(rsget("keywords")), "&", "＆")
				FItemList(i).FItemContent = db2html(rsget("itemcontent"))


				'FItemList(i).FKaffacategory     = rsget("category")
				FItemList(i).Fitemsize      = db2html(rsget("itemsize"))
				FItemList(i).Fitemsource    = db2html(rsget("itemsource"))
				FItemList(i).FLastUpdate    = rsget("LastUpdate")
				FItemList(i).Fkaffamakerid = rsget("kaffamakerid")
				FItemList(i).Fkaffacate1 = rsget("kaffacate1")
				FItemList(i).Fkaffacate2 = rsget("kaffacate2")
				FItemList(i).Fkaffacate3 = rsget("kaffacate3")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub


	public function GetMakeProducIndexItemList()
		dim sqlStr,i
        ''-- 옵션이 다품절되는경우.. 89745
		sqlStr = "select k.itemid " + vbcrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_Kaffa_reg_item as k " + vbcrlf
		sqlStr = sqlStr + " 	inner join [db_item].[dbo].tbl_item as i on k.itemid = i.itemid " + vbcrlf
		sqlStr = sqlStr + " where 1=1" ''k.kaffamakerid is not null
		sqlStr = sqlStr + " and i.basicimage is not null"
		sqlStr = sqlStr + " and i.itemdiv<50"
		sqlStr = sqlStr + " and i.cate_large<>''"
		sqlStr = sqlStr + " and i.cate_large<>'999'"
		sqlStr = sqlStr + " and i.sellcash>0"
		If FRectUseYN <> "" Then
			sqlStr = sqlStr + " and k.useyn = '" & FRectUseYN & "' "
		End If
		sqlStr = sqlStr + " order by i.itemid desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		if  not rsget.EOF  then
			GetMakeProducIndexItemList = rsget.getRows()
		end if
		rsget.Close
	end function

    ''품절처리 요망 상품
    public Sub getKaffaReqExpireItemList()
        Dim sqlStr, i, addsql
        addSql = ""
        addSql = addSql&" and (i.mwdiv='U' "                                            '' 업배
        addSql = addSql&"       or abs(i.itemweight-isNULL(s.kaffaWeight,0))>100"       '' 무게 100g 이상차이
        addSql = addSql&"       or isNULL(s.kaffaWeight,0)=0"                          '' 무게 없는것
        addSql = addSql&"       or ((i.sellcash<>0)"
		addSql = addSql&"             and ((i.sellcash-i.buycash)/i.sellcash)*100<"&CMAXMARGIN & VbCrlf
		addSql = addSql&"             and i.mwdiv<>'M')"        ''역마진 // 매입상품은 제외.
		addSql = addSql&"		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='kaffa') "
		addSql = addSql& ")"
        addSql = addSql&" and kaffaGoodNo is Not NULL"


        '브랜드검색
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'KAFFA상품번호 검색
		If FRectKAFFAPrdNo <> "" Then
			addSql = addSql & " and s.kaffaGoodNo='" & FRectKAFFAPrdNo & "'"
		End If

		'상품번호 검색
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
		End If

		'KAFFA등록여부 검색
		If FRectKaffaUseYN <> "" Then
			addSql = addSql & " and s.useyn = '" & FRectKaffaUseYN & "' "
		End If

        '판매여부 검색
		Select Case FRectSellYn
			Case "Y"	'판매
				addSql = addSql & " and i.sellYn='Y'"
			Case "N"	'품절
				addSql = addSql & " and i.sellYn in ('S','N')"
		End Select

		'한정여부 검색
		If FRectLimitYn<>"" then
			addSql = addSql & " and i.limitYn='" & FRectLimitYn & "'"
		End if

        '세일여부 검색
        If FRectTenSailyn<>"" then
			addSql = addSql & " and i.sailyn='" & FRectTenSailyn & "'"
		End if

        '해외기준가 세일여부검색
        If FRectKaffaBaseSailyn<>"" then
			addSql = addSql & " and i.itemid "&CHKIIF(FRectKaffaBaseSailyn="Y","","not")&" in ("
			addSql = addSql & " select p.itemid from db_item.dbo.tbl_item_multiLang_price P"
            addSql = addSql & " where P.sitename='CHNWEB'"
            addSql = addSql & " and P.currencyunit='WON'"
            addSql = addSql & " and P.orgPrice>P.maydiscountPrice"
			addSql = addSql & " )"
		End if

        'Kaffa 연동내역 할인여부
        If FRectKaffaSailyn<>"" then
            addSql = addSql & " and "&CHKIIF(FRectKaffaSailyn="Y","","not")&" (s.kaffaPrice>s.kaffadiscountPrice"
            addSql = addSql & " and s.kaffaDiscount_BEGIN_dateTime<=getdate()"
            addSql = addSql & " and s.kaffaDiscount_END_dateTime>=getdate() )"
        end if

		'매입구분
		if (FRectMWDiv<>"") then
            if (FRectMWDiv="MW") then
                addSql = addSql & " and i.mwdiv<>'U'"
            else
                addSql = addSql & " and i.mwdiv='"&FRectMWDiv&"'"
            end if
        end if

        if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        addSql = addSql & " and s.kaffaSellyn<>'X'"
		    else
		        addSql = addSql & " and s.kaffaSellyn='" & FRectExtSellYn & "'"
		    end if
		end if

        if (FRectExtDispYn<>"") then
            if (FRectExtDispYn="Y") then
                addSql = addSql & " and s.kaffaIsDisplay=1"
            else
                addSql = addSql & " and isNULL(s.kaffaIsDisplay,0)<>1"
            end if
        end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(s.itemid) as cnt, CEILING(CAST(Count(s.itemid) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_Kaffa_reg_item s " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN [db_item].[dbo].tbl_item i on s.itemid = i.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_contents c on i.itemid = c.itemid" & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_Kaffa_category_mapping m " & VBCRLF
'		sqlStr = sqlStr & "     inner join [db_user].[dbo].tbl_user_c as uc on i.makerid=uc.userid"
	    sqlStr = sqlStr & "     on i.cate_large=m.tencdl " & VBCRLF
	    sqlStr = sqlStr & "     and i.cate_mid=m.tencdm " & VBCRLF
	    sqlStr = sqlStr & " 	and i.cate_small=m.tencds " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + Cstr(FPageSize * FCurrPage) & VBCRLF
        sqlStr = sqlStr & " i.* " & VBCRLF
        sqlStr = sqlStr & " ,(SELECT top 1 isNull(username,'') as username FROM [db_partner].[dbo].tbl_user_tenbyten WHERE userid = s.reguserid) as reguser, s.useyn, s.kaffaRegDateTime as kaffaregdate, s.lastupdate as kaffalastupdate " & VBCRLF
        sqlStr = sqlStr & " ,s.kaffaprice, s.kaffasellyn, s.kaffagoodno, s.regedOptCnt, s.rctSellCNT, s.accFailCnt, s.lastErrStr, isNULL(s.kaffaWeight,0) as kaffaWeight " & VBCRLF
        sqlStr = sqlStr & " ,isNULL(s.kaffaIsDisplay,0) as kaffaIsDisplay"
        sqlStr = sqlStr & " ,isNULL(s.KaffaDiscountPrice, s.kaffaprice) as KaffaDiscountPrice, convert(varchar(19),s.kaffaDiscount_Begin_dateTime,21) as kaffaDiscount_Begin_dateTime,convert(varchar(19),s.kaffaDiscount_End_dateTime,21) as kaffaDiscount_End_dateTime"
        sqlStr = sqlStr & " ,isNULL(P.orgprice,-999) as multiPrice, isNULL(P.maydiscountPrice,0) as maydiscountPrice"
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_Kaffa_reg_item s " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN [db_item].[dbo].tbl_item i on s.itemid = i.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_contents c on i.itemid = c.itemid"
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_Kaffa_category_mapping m " & VBCRLF
	    sqlStr = sqlStr & "     on i.cate_large=m.tencdl " & VBCRLF
	    sqlStr = sqlStr & "     and i.cate_mid=m.tencdm " & VBCRLF
	    sqlStr = sqlStr & " 	and i.cate_small=m.tencds " & VBCRLF
	    sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_multiLang_price P on s.itemid=p.itemid and P.sitename='CHNWEB' and currencyunit='WON'"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & addSql
		IF (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " ORDER BY i.itemscore DESC, i.itemid DESC"
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " ORDER BY s.rctSellCNT DESC, i.itemscore DESC, s.itemid DESC"
		Else
		    sqlStr = sqlStr & " ORDER BY i.itemid DESC"
	    End If
'rw sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new cKaffaOneItem
					FItemList(i).Fitemid            = rsget("itemid")
					FItemList(i).Fmakerid           = rsget("makerid")
					FItemList(i).Fcate_large        = rsget("cate_large")
					FItemList(i).Fcate_mid          = rsget("cate_mid")
					FItemList(i).Fcate_small        = rsget("cate_small")
					FItemList(i).Fitemdiv           = rsget("itemdiv")
					FItemList(i).Fitemgubun         = rsget("itemgubun")
					FItemList(i).Fitemname          = db2html(rsget("itemname"))
					FItemList(i).Fsellcash          = rsget("sellcash")
					FItemList(i).Fbuycash           = rsget("buycash")
					FItemList(i).Forgprice          = rsget("orgprice")
					FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
					FItemList(i).Fsailprice         = rsget("sailprice")
					FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
					FItemList(i).Fmileage           = rsget("mileage")
					FItemList(i).Fregdate           = rsget("regdate")
					FItemList(i).Flastupdate        = rsget("lastupdate")
					FItemList(i).FsellEndDate       = rsget("sellEndDate")
					FItemList(i).Fsellyn            = rsget("sellyn")
					FItemList(i).Flimityn           = rsget("limityn")
					FItemList(i).Fdanjongyn         = rsget("danjongyn")
					FItemList(i).Fsailyn            = rsget("sailyn")
					FItemList(i).Fisusing           = rsget("isusing")
					FItemList(i).Fisextusing        = rsget("isextusing")
					FItemList(i).Fmwdiv             = rsget("mwdiv")
					FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
					FItemList(i).Fvatinclude        = rsget("vatinclude")
					FItemList(i).Fdeliverytype      = rsget("deliverytype")
					FItemList(i).Fdeliverarea       = rsget("deliverarea")
					FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
					FItemList(i).Fismobileitem      = rsget("ismobileitem")
					FItemList(i).Fpojangok          = rsget("pojangok")
					FItemList(i).Flimitno           = rsget("limitno")
					FItemList(i).Flimitsold         = rsget("limitsold")
					FItemList(i).Fevalcnt           = rsget("evalcnt")
					FItemList(i).Foptioncnt         = rsget("optioncnt")
					FItemList(i).Fitemrackcode      = rsget("itemrackcode")
					FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
					FItemList(i).Fbrandname         = db2html(rsget("brandname"))
					FItemList(i).FitemWeight		= rsget("itemWeight")
					FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
					FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
					FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
					FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
					FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
					FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
					FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
					FItemList(i).FdeliverOverseas	= rsget("deliverOverseas")
'					FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가
					FItemList(i).FRegUser			= rsget("reguser")
					FItemList(i).FKaffaUseYN		= rsget("useyn")
					FItemList(i).FKaffaregdate		= rsget("kaffaregdate")
					FItemList(i).FKaffalastupdate	= rsget("kaffalastupdate")
					FItemList(i).FKaffaprice		= rsget("kaffaprice")
					FItemList(i).FKaffasellyn		= rsget("kaffasellyn")
					FItemList(i).FKaffagoodno		= rsget("kaffagoodno")
					FItemList(i).FregedOptCnt		= rsget("regedOptCnt")
					FItemList(i).FrctSellCNT		= rsget("rctSellCNT")
					FItemList(i).FaccFailCnt		= rsget("accFailCnt")
					FItemList(i).FlastErrStr		= rsget("lastErrStr")
                    FItemList(i).FkaffaWeight       = rsget("kaffaWeight")
                    FItemList(i).FkaffaIsDisplay    = rsget("kaffaIsDisplay")

                    FItemList(i).FKaffaDiscountPrice = rsget("KaffaDiscountPrice")
                    FItemList(i).FkaffaDiscount_Begin_dateTime = rsget("kaffaDiscount_Begin_dateTime")
                    FItemList(i).FkaffaDiscount_End_dateTime   = rsget("kaffaDiscount_End_dateTime")
                    FItemList(i).FmultiPrice         = rsget("multiPrice")
                    FItemList(i).FmaydiscountPrice   = rsget("maydiscountPrice")



					'if (rsget("infoimageExists")>0) then
					'    FItemList(i).FinfoimageExists   = true
					'else
					'    FItemList(i).FinfoimageExists   = false
					'end if

					''//기본 배송비 정책 관련 추가
					'FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
					'FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")
					'FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
    end Sub

	Public Sub GetAllKaffaItemList_USESCM()
		Dim sqlStr, i, addsql

		'브랜드검색
		If FRectMakerid <> "" Then
			addSql = addSql & " and i.makerid='" & FRectMakerid & "'"
		End If

		'KAFFA상품번호 검색
		If FRectKAFFAPrdNo <> "" Then
			addSql = addSql & " and s.kaffaGoodNo='" & FRectKAFFAPrdNo & "'"
		End If

		'상품명 검색
		If FRectItemName <> "" Then
			addSql = addSql & " and i.itemname like '%" & FRectItemName & "%'"
		End if

		'카테고리 검색
		If FRectCDL <> "" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if
		If FRectCDM <> "" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if
		If FRectCDS <> "" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End If

		'상품번호 검색
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
		End If

		'KAFFA등록여부 검색
		If FRectKaffaUseYN <> "" Then
			If FRectKaffaUseYN = "m" Then
				addSql = addSql & " and s.lastupdate < i.lastupdate "
			ElseIf FRectKaffaUseYN = "w" Then
				addSql = addSql & " and s.kaffaGoodNo is NULL "
			Else
				addSql = addSql & " and s.useyn = '" & FRectKaffaUseYN & "' "
			End If
		End If

		'이벤트번호 검색
		If FRectEventid <> "" Then
			addSql = addSql & " and i.itemid in (Select itemid From [db_event].[dbo].tbl_eventitem Where evt_code='" & FRectEventid & "')" & VbCrlf
		End If

		'판매여부 검색
		Select Case FRectSellYn
			Case "Y"	'판매
				addSql = addSql & " and i.sellYn='Y'"
			Case "N"	'품절
				addSql = addSql & " and i.sellYn in ('S','N')"
		End Select

		'한정여부 검색
		If FRectLimitYn<>"" then
			addSql = addSql & " and i.limitYn='" & FRectLimitYn & "'"
		End if

        '세일여부 검색
        If FRectTenSailyn<>"" then
			addSql = addSql & " and i.sailyn='" & FRectTenSailyn & "'"
		End if

        '해외기준가 세일여부검색
        If FRectKaffaBaseSailyn<>"" then
			addSql = addSql & " and i.itemid "&CHKIIF(FRectKaffaBaseSailyn="Y","","not")&" in ("
			addSql = addSql & " select p.itemid from db_item.dbo.tbl_item_multiLang_price P"
            addSql = addSql & " where P.sitename='CHNWEB'"
            addSql = addSql & " and P.currencyunit='WON'"
            addSql = addSql & " and P.orgPrice>P.maydiscountPrice"
			addSql = addSql & " )"
		End if

		'Kaffa 연동내역 할인여부
        If FRectKaffaSailyn<>"" then
            addSql = addSql & " and "&CHKIIF(FRectKaffaSailyn="Y","","not")&" (s.kaffaPrice>s.kaffadiscountPrice"
            addSql = addSql & " and s.kaffaDiscount_BEGIN_dateTime<=getdate()"
            addSql = addSql & " and s.kaffaDiscount_END_dateTime>=getdate() )"
        end if

		'매입구분
		if (FRectMWDiv<>"") then
            if (FRectMWDiv="MW") then
                addSql = addSql & " and i.mwdiv<>'U'"
            else
                addSql = addSql & " and i.mwdiv='"&FRectMWDiv&"'"
            end if
        end if

		'역마진 및 마진 CMAXMARGIN 이상 검색
		If (FRectMinusMargin<>"") then
		   addSql = addSql & " and i.sellcash<>0"
		   addSql = addSql & " and ((i.sellcash-i.buycash)/i.sellcash)*100<"&CMAXMARGIN & VbCrlf
		   addSql = addSql & " and s.useyn= 'y' " '''  조건 추가.
		Else
		   IF (FRectonlyValidMargin<>"") then
		        addSql = addSql & " and i.sellcash<>0"
		        addSql = addSql & " and ((i.sellcash-i.buycash)/i.sellcash)*100>="&CMAXMARGIN & VbCrlf
		   END IF
		End If

		If (FRectFailCntExists <> "") Then
			addSql = addSql & " and s.accFailCNT > 0"
		End If

		''옵션추가금액 존재상품.
		If (FRectoptAddprcExists <> "") and (FRectKaffaUseYN <> "N") Then
			addSql = addSql & " and i.itemid in ("
			addSql = addSql & "     select distinct ii.itemid "
			addSql = addSql & "     from db_item.dbo.tbl_item ii "
			addSql = addSql & "     Join db_item.dbo.tbl_item_option o "
			addSql = addSql & "     on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
			addSql = addSql & " )"
		End If

		''옵션추가금액 존재상품 제외
		If (FRectoptAddprcExistsExcept <> "") Then
			addSql = addSql & " and i.itemid Not in ("
			addSql = addSql & "     select distinct ii.itemid "
			addSql = addSql & "     from db_item.dbo.tbl_item ii "
			addSql = addSql & "     Join db_item.dbo.tbl_item_option o "
			addSql = addSql & "     on ii.itemid=o.itemid and o.optaddprice>0 and o.isusing='Y'"
			addSql = addSql & " )"
		End If

		if (FRectoptExists<>"") then
            addSql = addSql & " and i.optioncnt>0"
        end if

        ''cj판매 10x10 품절
        IF (FRectKAFFASell10x10Soldout<>"") then
            addSql = addSql & " and i.sellyn<>'Y'"
            addSql = addSql & " and s.kaffaSellyn='Y'"
        end if

        ''cj가격 <10x10 가격
        IF (FRectexpensive10x10<>"") then
            ''addSql = addSql & " and s.kaffaPrice<(CASE WHEN i.makerid in ('ithinkso','antennashop') THEN i.orgprice*1.5 ELSE i.orgprice END)" ''sellcash=>orgprice 2013/07/04
            addSql = addSql & " and s.kaffaPrice<P.orgprice" '' 2013/07/24
        end if

        ''가격상이(on판매가<>kaffa)
        if FRectdiffPrc <> "" then
		   addSql = addSql & " and s.kaffaPrice is Not Null and (CASE WHEN i.makerid in ('ithinkso','antennashop') THEN i.orgprice*1.5 ELSE i.orgprice END) <> s.kaffaPrice " ''sellcash=>orgprice 2013/07/04
		end if

        if FRectdiffMultiPrc<>"" then
           'addSql = addSql & " and isNULL(P.orgprice,-999)<>s.kaffaPrice "
           addSql = addSql & " and ((isNULL(P.orgprice,-999)<>s.kaffaPrice) or (isNULL(P.maydiscountPrice,P.orgprice)<>s.kaffaDiscountPrice)) "
        end if

        if FRectdiffWeight <> "" then
		   addSql = addSql & " and i.itemWeight <> isNULL(s.kaffaWeight,0) "
		end if

		if (FRectExtSellYn<>"") then
		    if (FRectExtSellYn="YN") then
		        addSql = addSql & " and s.kaffaSellyn<>'X'"
		    else
		        addSql = addSql & " and s.kaffaSellyn='" & FRectExtSellYn & "'"
		    end if
		end if

        if (FRectExtDispYn<>"") then
            if (FRectExtDispYn="Y") then
                addSql = addSql & " and s.kaffaIsDisplay=1"
            else
                addSql = addSql & " and isNULL(s.kaffaIsDisplay,0)<>1"
            end if
        end if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(s.itemid) as cnt, CEILING(CAST(Count(s.itemid) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_Kaffa_reg_item s " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN [db_item].[dbo].tbl_item i on s.itemid = i.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_contents c on i.itemid = c.itemid" & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_Kaffa_category_mapping m " & VBCRLF
	    sqlStr = sqlStr & "     on i.cate_large=m.tencdl " & VBCRLF
	    sqlStr = sqlStr & "     and i.cate_mid=m.tencdm " & VBCRLF
	    sqlStr = sqlStr & " 	and i.cate_small=m.tencds " & VBCRLF
	    sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_multiLang_price P on s.itemid=p.itemid and P.sitename='CHNWEB' and currencyunit='WON'"
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
'rw sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Clng(FCurrPage) > Clng(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + Cstr(FPageSize * FCurrPage) & VBCRLF
        sqlStr = sqlStr & " i.* " & VBCRLF
        sqlStr = sqlStr & " ,(SELECT top 1 isNull(username,'') as username FROM [db_partner].[dbo].tbl_user_tenbyten WHERE userid = s.reguserid) as reguser, s.useyn, s.kaffaRegDateTime as kaffaregdate, s.lastupdate as kaffalastupdate " & VBCRLF
        sqlStr = sqlStr & " ,s.kaffaprice, s.kaffasellyn, s.kaffagoodno, s.regedOptCnt, s.rctSellCNT, s.accFailCnt, s.lastErrStr, isNULL(s.kaffaWeight,0) as kaffaWeight " & VBCRLF
        sqlStr = sqlStr & " ,isNULL(s.kaffaIsDisplay,0) as kaffaIsDisplay"
        sqlStr = sqlStr & " ,isNULL(s.KaffaDiscountPrice, s.kaffaprice) as KaffaDiscountPrice, convert(varchar(19),s.kaffaDiscount_Begin_dateTime,21) as kaffaDiscount_Begin_dateTime,convert(varchar(19),s.kaffaDiscount_End_dateTime,21) as kaffaDiscount_End_dateTime"
        sqlStr = sqlStr & " ,isNULL(P.orgprice,-999) as multiPrice, isNULL(P.maydiscountPrice,0) as maydiscountPrice"
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_Kaffa_reg_item s " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN [db_item].[dbo].tbl_item i on s.itemid = i.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_contents c on i.itemid = c.itemid"
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_Kaffa_category_mapping m " & VBCRLF
	    sqlStr = sqlStr & "     on i.cate_large=m.tencdl " & VBCRLF
	    sqlStr = sqlStr & "     and i.cate_mid=m.tencdm " & VBCRLF
	    sqlStr = sqlStr & " 	and i.cate_small=m.tencds " & VBCRLF
	    sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item_multiLang_price P on s.itemid=p.itemid and P.sitename='CHNWEB' and currencyunit='WON'"
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & addSql

		sqlStr = sqlStr & " ORDER BY "

		if FRectdiffMultiPrc<>"" then
		    sqlStr = sqlStr & " isNULL(P.maydiscountPrice,0) desc,isNULL(s.kaffaDiscount_Begin_DateTIME,'2999-12-31') asc,"
		end if

		IF (FRectOrdType = "B") Then
		    sqlStr = sqlStr & " i.itemscore DESC, i.itemid DESC"
		ElseIf (FRectOrdType = "BM") Then
		    sqlStr = sqlStr & " s.rctSellCNT DESC, i.itemscore DESC, s.itemid DESC"
		Else
		    sqlStr = sqlStr & " i.itemid DESC"
	    End If
'rw sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new cKaffaOneItem
					FItemList(i).Fitemid            = rsget("itemid")
					FItemList(i).Fmakerid           = rsget("makerid")
					FItemList(i).Fcate_large        = rsget("cate_large")
					FItemList(i).Fcate_mid          = rsget("cate_mid")
					FItemList(i).Fcate_small        = rsget("cate_small")
					FItemList(i).Fitemdiv           = rsget("itemdiv")
					FItemList(i).Fitemgubun         = rsget("itemgubun")
					FItemList(i).Fitemname          = db2html(rsget("itemname"))
					FItemList(i).Fsellcash          = rsget("sellcash")
					FItemList(i).Fbuycash           = rsget("buycash")
					FItemList(i).Forgprice          = rsget("orgprice")
					FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
					FItemList(i).Fsailprice         = rsget("sailprice")
					FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
					FItemList(i).Fmileage           = rsget("mileage")
					FItemList(i).Fregdate           = rsget("regdate")
					FItemList(i).Flastupdate        = rsget("lastupdate")
					FItemList(i).FsellEndDate       = rsget("sellEndDate")
					FItemList(i).Fsellyn            = rsget("sellyn")
					FItemList(i).Flimityn           = rsget("limityn")
					FItemList(i).Fdanjongyn         = rsget("danjongyn")
					FItemList(i).Fsailyn            = rsget("sailyn")
					FItemList(i).Fisusing           = rsget("isusing")
					FItemList(i).Fisextusing        = rsget("isextusing")
					FItemList(i).Fmwdiv             = rsget("mwdiv")
					FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
					FItemList(i).Fvatinclude        = rsget("vatinclude")
					FItemList(i).Fdeliverytype      = rsget("deliverytype")
					FItemList(i).Fdeliverarea       = rsget("deliverarea")
					FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
					FItemList(i).Fismobileitem      = rsget("ismobileitem")
					FItemList(i).Fpojangok          = rsget("pojangok")
					FItemList(i).Flimitno           = rsget("limitno")
					FItemList(i).Flimitsold         = rsget("limitsold")
					FItemList(i).Fevalcnt           = rsget("evalcnt")
					FItemList(i).Foptioncnt         = rsget("optioncnt")
					FItemList(i).Fitemrackcode      = rsget("itemrackcode")
					FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
					FItemList(i).Fbrandname         = db2html(rsget("brandname"))
					FItemList(i).FitemWeight		= rsget("itemWeight")
					FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
					FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
					FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
					FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
					FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
					FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
					FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
					FItemList(i).FdeliverOverseas	= rsget("deliverOverseas")
					FItemList(i).FRegUser			= rsget("reguser")
					FItemList(i).FKaffaUseYN		= rsget("useyn")
					FItemList(i).FKaffaregdate		= rsget("kaffaregdate")
					FItemList(i).FKaffalastupdate	= rsget("kaffalastupdate")
					FItemList(i).FKaffaprice		= rsget("kaffaprice")
					FItemList(i).FKaffasellyn		= rsget("kaffasellyn")
					FItemList(i).FKaffagoodno		= rsget("kaffagoodno")
					FItemList(i).FregedOptCnt		= rsget("regedOptCnt")
					FItemList(i).FrctSellCNT		= rsget("rctSellCNT")
					FItemList(i).FaccFailCnt		= rsget("accFailCnt")
					FItemList(i).FlastErrStr		= rsget("lastErrStr")
                    FItemList(i).FkaffaWeight       = rsget("kaffaWeight")
                    FItemList(i).FkaffaIsDisplay    = rsget("kaffaIsDisplay")

                    FItemList(i).FKaffaDiscountPrice = rsget("KaffaDiscountPrice")
                    FItemList(i).FkaffaDiscount_Begin_dateTime = rsget("kaffaDiscount_Begin_dateTime")
                    FItemList(i).FkaffaDiscount_End_dateTime   = rsget("kaffaDiscount_End_dateTime")
                    FItemList(i).FmultiPrice         = rsget("multiPrice")
                    FItemList(i).FmaydiscountPrice   = rsget("maydiscountPrice")


					'if (rsget("infoimageExists")>0) then
					'    FItemList(i).FinfoimageExists   = true
					'else
					'    FItemList(i).FinfoimageExists   = false
					'end if

					''//기본 배송비 정책 관련 추가
					'FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
					'FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")
					'FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function
End Class


Function CateLargeSelectBox(cdl)
	Dim sqlStr, vBody
	sqlStr = "select top 100 * from [db_item].[dbo].tbl_Cate_large"
	sqlStr = sqlStr + " where code_large<'999'"
	sqlStr = sqlStr + " and display_yn='Y'"
	sqlStr = sqlStr + " order by code_large"

	rsget.Open sqlStr, dbget, 1
	do until rsget.Eof
		vBody = vBody & "<option value=""" & rsget("code_large") & """ " & CHKIIF(CStr(cdl)=CStr(rsget("code_large")),"selected","") & ">" & db2html(rsget("code_nm")) & "</option>" & vbCrLf
		rsget.moveNext
	loop
	rsget.close

	CateLargeSelectBox = vBody
End Function


Function KaffaCate1SelectBox()
	Dim sqlStr, vBody
	sqlStr = "select cate1, cate_nm from [db_item].[dbo].[tbl_kaffa_category]"
	sqlStr = sqlStr + " where cate2 = 0 and cate3 = 0"
	sqlStr = sqlStr + " order by cate1 asc"

	rsget.Open sqlStr, dbget, 1
	do until rsget.Eof
		vBody = vBody & "<option value=""" & rsget("cate1") & """>" & rsget("cate_nm") & "</option>" & vbCrLf
		rsget.moveNext
	loop
	rsget.close

	KaffaCate1SelectBox = vBody
End Function


Function KaffaCate2SelectBox(arr, a, v)
	Dim i, vBody
	For i = 0 To UBound(arr,2)
		If CStr(arr(0,i)) = CStr(a) Then
			vBody = vBody & "<option value=""" & arr(1,i) & """ " & CHKIIF(CStr(arr(1,i))=CStr(v),"selected","") & ">" & arr(2,i) & "</option>" & vbCrLf
		End IF
	Next

	KaffaCate2SelectBox = vBody
End Function


Function KaffaCate3SelectBox(arr, a1, a2, v)
	Dim i, vBody
	For i = 0 To UBound(arr,2)
		If CStr(arr(0,i)) = CStr(a1) AND CStr(arr(1,i)) = CStr(a2) Then
			vBody = vBody & "<option value=""" & arr(2,i) & """ " & CHKIIF(CStr(arr(2,i))=CStr(v),"selected","") & ">" & arr(3,i) & "</option>" & vbCrLf
		End IF
	Next

	KaffaCate3SelectBox = vBody
End Function
%>