<%
'####################################################
' Description :  오프샵
' History : 2009.04.07 서동석 생성
'			2010.08.04 한용민 수정
'####################################################

'// 이전 달 마진 입력내역 있는지
function IsExistPrevMonthShopJungsanLog(shopid, makerid)
	dim sqlStr

	sqlStr = "select count(Logidx) as cnt from db_shop.dbo.tbl_shop_designer_maginLog"
	sqlStr = sqlStr + " where 1 = 1 "
	sqlStr = sqlStr + " and shopid='" & shopid & "'"
	sqlStr = sqlStr + " and makerid='" & makerid & "'"
	sqlStr = sqlStr + " and regdate < '" + Left(Now(), 7) + "-01' "

	rsget.Open sqlStr,dbget,1
		IsExistPrevMonthShopJungsanLog = (rsget("cnt") > 0)
	rsget.Close
end Function

class COffShopItem
	public floginsitename
	public fcompany_name
	public fcompany_no
	public fcountrylangcd
	public floginsite
	public Fuserid
	public Fuserpass
	public Fshopname
	public Fshopphone
	public Fshopzipcode
	public Fshopaddr1
	public Fshopaddr2
	public Fmanname
	public Fmanhp
	public Fmanphone
	public Fmanemail
	public Fisusing
	public FShopdiv
	public FShopdivName
	public Fstockbasedate
	public Fshopsocname
	public Fshopsocno
	public Fshopceoname
	public Fvieworder
	public Fgroupid
	public fcurrencyUnit
	public fcurrencyUnit_Pos
	public fexchangeRate
	public fbasedate
	public fcurrencyChar
	public fregdate
	public freguserid
	public flastuserid
	public fidx
	public fmultipleRate
    public fpyeong
	public FshopCountryCode
    public FcountryNamekr
	public fshopid
    public FdecimalPointLen
    public FdecimalPointCut
	public fothershopid
	public fsiteseq
	public flastupdate
	public flastadminuserid
    public Fismobileusing
    public Fmobileshopname
    public Fmobileshopimage
    public Fmobileworkhour
    public Fmobileclosedate
    public Fmobiletel
    public Fmobileaddr
    public Fmobilemapimage
    public Fmobilebysubway
    public Fmobilebybus
    public Fmobilelatitude
    public Fmobilelongitude
	public fadmindisplang
	public fsitename
	public Fctropen
	public FViewSort
	public FengName
	public FShopFax
	public FengAddress

	public function GetMobileShopImage()
		GetMobileShopImage = "http://webimage.10x10.co.kr/" + Fmobileshopimage
    end function

	public function GetMobileShopImage50X50()
		GetMobileShopImage50X50 = "http://webimage.10x10.co.kr/mobileshopimage/50X50/" + Fuserid + ".png"
    end function

	public function GetMobileMapImage()
		GetMobileMapImage = "http://webimage.10x10.co.kr/" + Fmobilemapimage
    end function

	'//사용안함
	'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 drawoffshop_commoncode 사용
	public function GetShopdivName()
	    '' 대표코드 2,4,6
		if FShopdiv="1" then
			GetShopdivName = "직영"
		elseif FShopdiv="2" then
			GetShopdivName = "직영[대표]"
		elseif FShopdiv="3" then
			GetShopdivName = "가맹점"
		elseif FShopdiv="4" then
			GetShopdivName = "가맹[대표]"
		elseif FShopdiv="5" then
			GetShopdivName = "도매"
		elseif FShopdiv="6" then
			GetShopdivName = "도매[대표]"
		elseif FShopdiv="7" then
			GetShopdivName = "해외"
		elseif FShopdiv="8" then
			GetShopdivName = "해외[대표]"
		elseif FShopdiv="9" then
			GetShopdivName = "ETC"
		elseif FShopdiv="11" then
			GetShopdivName = "ithinkso"
		elseif FShopdiv="12" then
			GetShopdivName = "ithinkso[대표]"
		elseif FShopdiv="13" then
			GetShopdivName = "대행"
		elseif FShopdiv="14" then
			GetShopdivName = "대행[대표]"
		else
			GetShopdivName = FShopdiv
		end if
	end function

    ''직영샵
	'//사용안함
	'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 getoffshop_commoncodegroup 사용
    function IsOwnShop()
        IsOwnShop = (FShopdiv="1") or (FShopdiv="2")
    end function

    ''가맹샵
	'//사용안함
	'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 getoffshop_commoncodegroup 사용
    function IsJoinShop()
        IsJoinShop = (FShopdiv="3") or (FShopdiv="4")
    end function

    ''도매(기존 해외였음(
	'//사용안함
	'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 getoffshop_commoncodegroup 사용
    function IsAbrodShop()
        IsAbrodShop = (FShopdiv="5") or (FShopdiv="6")
    end function

    ''해외
	'//사용안함
	'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 getoffshop_commoncodegroup 사용
    function IsForeignShop()
        IsForeignShop = (FShopdiv="7") or (FShopdiv="8")
    end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COffShopDesignerItem
	public FShopId
	public FShopName
	public FDesignerId
	public FDesignerName
	public FChargeDiv
	public FComm_cd
	public FDefaultMargin
	public FDefaultSuplyMargin
	public FSubtract
	public FRegdate
	public FAdminopen
	public FEtcjunsandetail
	public Fonlinedefaultmargine
	public FIsUsing
	public Fbrandisusing
	public FPartnerisusing
    public FShopisusing
	public FOnlineMWDiv
	public FAutoJungsan
	public FAutoJungsanDiv
    public Fjungsan_date_off
    public Fjungsan_date
    public FShopdiv
    public FdefaultCenterMwDiv
    public Fdefaultbeasongdiv		'배송구분(매장판매/업체배송)
    public Fjungsan_gubun
	public fitemregyn
	public fshopdivname
	public fisoffusing
	public fpurchaseType

    public function IsContractExists()
        IsContractExists = Not (IsNULL(FComm_cd) or (FComm_cd=""))
    end function

	'//사용안함
	'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 getoffshop_commoncodegroup 사용
    public function IsProtoTypeShop()
        IsProtoTypeShop = ((FShopdiv="2") or (FShopdiv="4") or (FShopdiv="6") or (FShopdiv="8") or (FShopdiv="12") or (FShopdiv="14"))
    end function

    ''직영샵
	'//사용안함
	'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 getoffshop_commoncodegroup 사용
    function IsOwnShop()
        IsOwnShop = (FShopdiv="1") or (FShopdiv="2")
    end function

    ''가맹샵
	'//사용안함
	'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 getoffshop_commoncodegroup 사용
    function IsJoinShop()
        IsJoinShop = (FShopdiv="3") or (FShopdiv="4")
    end function

    ''도매(기존 해외였음(
	'//사용안함
	'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 getoffshop_commoncodegroup 사용
    function IsAbrodShop()
        IsAbrodShop = (FShopdiv="5") or (FShopdiv="6")
    end function

    ''해외
	'//사용안함
	'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 getoffshop_commoncodegroup 사용
    function IsForeignShop()
        IsForeignShop = (FShopdiv="7") or (FShopdiv="8")
    end function

	'//사용안함
	'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 drawoffshop_commoncode 사용
    public function GetShopdivName()
	    '' 대표코드 2,4,6
		if FShopdiv="1" then
			GetShopdivName = "직영"
		elseif FShopdiv="2" then
			GetShopdivName = "직영[대표]"
		elseif FShopdiv="3" then
			GetShopdivName = "가맹점"
		elseif FShopdiv="4" then
			GetShopdivName = "가맹[대표]"
		elseif FShopdiv="5" then
			GetShopdivName = "도매"
		elseif FShopdiv="6" then
			GetShopdivName = "도매[대표]"
		elseif FShopdiv="9" then
			GetShopdivName = "ETC"
		elseif FShopdiv="11" then
			GetShopdivName = "ithinkso"
		elseif FShopdiv="12" then
			GetShopdivName = "ithinkso[대표]"
		else
			GetShopdivName = FShopdiv
		end if
	end function

	public function getMwName()
		if FOnlineMWDiv="M" then
			getMwName = "매입"
		elseif FOnlineMWDiv="W" then
			getMwName = "위탁"
		elseif FOnlineMWDiv="U" then
			getMwName = "업체"
		end if
	end function

	public function getChargeDivColor()
	    getChargeDivColor = getJungsanDivColor
'		if FChargeDiv="2" then
'			getChargeDivColor = "#000000"
'		elseif FChargeDiv="4" then
'			getChargeDivColor = "red"
'		elseif FChargeDiv="5" then
'			getChargeDivColor = "#FF44FF"
'		elseif FChargeDiv="6" then
'			getChargeDivColor = "blue"
'		elseif FChargeDiv="8" then
'			getChargeDivColor = "#FF44FF"
'		end if
	end function

	public function getChargeDivName()
	    getChargeDivName = getJungsanDivName
'		if FChargeDiv="2" then
'			getChargeDivName = "텐바이텐위탁"
'		elseif FChargeDiv="4" then
'			getChargeDivName = "텐바이텐매입"
'		elseif FChargeDiv="5" then
'			getChargeDivName = "매입출고정산"
'		elseif FChargeDiv="6" then
'			getChargeDivName = "업체위탁"
'		elseif FChargeDiv="8" then
'			getChargeDivName = "업체매입"
'		end if
	end function

    public function getJungsanDivColor()
		if FComm_cd="B011" then
			getJungsanDivColor = "#000000"
		elseif FComm_cd="B031" then
			getJungsanDivColor = "#FF44FF"
		elseif FComm_cd="B012" then
			getJungsanDivColor = "blue"
		elseif FComm_cd="B022" then
			getJungsanDivColor = "#FF44FF"
		elseif FComm_cd="B021" then
			getJungsanDivColor = "#FF44FF"
		elseif FComm_cd="B023" then
			getJungsanDivColor = "#55AA55"
		elseif FComm_cd="B013" then
			getJungsanDivColor = "#55AA55"
		end if
	end function

	public function getJungsanDivName()
		if FComm_cd="B011" then
			getJungsanDivName = "텐바이텐위탁"
		elseif FComm_cd="B031" then
			getJungsanDivName = "매입출고정산"
		elseif FComm_cd="B012" then
			getJungsanDivName = "업체위탁"
		elseif FComm_cd="B022" then
			getJungsanDivName = "매장매입"
		elseif FComm_cd="B021" then
			getJungsanDivName = "오프매입"
		elseif FComm_cd="B023" then
			getJungsanDivName = "가맹점매입"
		elseif FComm_cd="B013" then
			getJungsanDivName = "출고위탁"
		end if

    end function

	public function getDefaultBeasongDivName()
		if Fdefaultbeasongdiv="2" then
			getDefaultBeasongDivName = "업체배송"
		elseif Fdefaultbeasongdiv="0" then
			getDefaultBeasongDivName = "매장판매"
		else
			getDefaultBeasongDivName = Fdefaultbeasongdiv
		end if
    end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COffShopChargeUserItem
	public FChargeUser
	public FChargeName
	public FPartnerUserId
	public FregDate

	public function ChargeType2Name(byval ctype)
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COffshopMarginLogItem
    public Flogidx
    public Fshopid
    public Fmakerid
    public Fcomm_cd
    public Fdefaultmargin
    public Fdefaultsuplymargin
    public Fdefaultbeasongdiv
    public FactFlag
    public Fregdate
    public Freguserid
    public FdeleteYN
    public FDefaultCenterMwdiv



    public function getActFlagName()
        if (FactFlag="P") then
            getActFlagName = "-"
        elseif FactFlag="M" then
			getActFlagName = "수정"
		elseif FactFlag="I" then
			getActFlagName = "입력"
		elseif FactFlag="D" then
			getActFlagName = "삭제"
		else
		    getActFlagName = FactFlag
		end if
    end function

    public function getJungsanDivColor()
		if FComm_cd="B011" then
			getJungsanDivColor = "#000000"
		elseif FComm_cd="B031" then
			getJungsanDivColor = "#FF44FF"
		elseif FComm_cd="B012" then
			getJungsanDivColor = "blue"
		elseif FComm_cd="B022" then
			getJungsanDivColor = "#FF44FF"
		elseif FComm_cd="B021" then
			getJungsanDivColor = "#FF44FF"
		elseif FComm_cd="B023" then
			getJungsanDivColor = "#55AA55"
		elseif FComm_cd="B013" then
			getJungsanDivColor = "#55AA55"
		end if
	end function

	public function getJungsanDivName()
		if FComm_cd="B011" then
			getJungsanDivName = "텐바이텐위탁"
		elseif FComm_cd="B031" then
			getJungsanDivName = "매입출고정산"
		elseif FComm_cd="B012" then
			getJungsanDivName = "업체위탁"
		elseif FComm_cd="B022" then
			getJungsanDivName = "매장매입"
		elseif FComm_cd="B021" then
			getJungsanDivName = "오프매입"
		elseif FComm_cd="B023" then
			getJungsanDivName = "가맹점매입"
		elseif FComm_cd="B013" then
			getJungsanDivName = "출고위탁"
		end if

    end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COffShopChargeUser
    public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount
	public FRectShopID
	public FRectDesigner
	public FRectChargeType
    public FRectComm_cd
    public FRectShopusing
    public FRectPartnerusing
    public FRectBrandusing
    public FRectcurrencyUnit
    public FRectbasedate
    public frectidx
    public FRectOffUpBea
	public FRectHasContOnly
    public FRectShopDiv2
    public FRectIsUsing
    public frectshopname
    public FRectNotProtoTypeShop
    public frectcodegroup
    public frectshopdiv
	public frectsitename
	public FRectmaeipdiv
	public FRectBrandPurchaseType
	public FRectisoffusing
	public FRectadminopen
	public FRectloginsite
	public FRectcountrylangcd
	public FRectvieworder

    public Sub GetOffShopMarginLogList()
        dim sqlStr,i

        sqlStr = "select count(Logidx) as cnt from db_shop.dbo.tbl_shop_designer_maginLog"
        sqlStr = sqlStr + " where deleteYN='N'"
        if (FRectShopID<>"") then
            sqlStr = sqlStr + " and shopid='" & FRectShopID & "'"
        end if

        if (FRectDesigner<>"") then
            sqlStr = sqlStr + " and makerid='" & FRectDesigner & "'"
        end if

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " * "
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_designer_maginLog "
		sqlStr = sqlStr + " where deleteYN='N'"
        if (FRectShopID<>"") then
            sqlStr = sqlStr + " and shopid='" & FRectShopID & "'"
        end if

        if (FRectDesigner<>"") then
            sqlStr = sqlStr + " and makerid='" & FRectDesigner & "'"
        end if

		sqlStr = sqlStr + " order by Logidx desc"

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
				set FItemList(i) = new COffshopMarginLogItem

				FItemList(i).Flogidx               = rsget("logidx")
				FItemList(i).FShopId               = rsget("shopid")
                FItemList(i).Fmakerid              = rsget("makerid")
                FItemList(i).Fcomm_cd              = rsget("comm_cd")
                FItemList(i).Fdefaultmargin        = rsget("defaultmargin")
                FItemList(i).Fdefaultsuplymargin   = rsget("defaultsuplymargin")
                FItemList(i).Fdefaultbeasongdiv   = rsget("defaultbeasongdiv")
                FItemList(i).FactFlag              = rsget("actFlag")
                FItemList(i).Fregdate              = rsget("regdate")
                FItemList(i).Freguserid            = rsget("reguserid")
                FItemList(i).FdeleteYN             = rsget("deleteYN")
                FItemList(i).FDefaultCenterMwdiv             = rsget("DefaultCenterMwdiv")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

	public Sub GetErrContractList()
		dim sqlStr,i

		sqlStr = "select top 1000 d.* , c.socname_kor, c.maeipdiv, c.defaultmargine as onlinedefaultmargine, c.isusing "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d "
		sqlStr = sqlStr + " Left join [db_user].[dbo].tbl_user_c c "
		sqlStr = sqlStr + " on d.makerid=c.userid "
		sqlStr = sqlStr + " where ((d.shopid='streetshop000') or (d.shopid='streetshop800'))"
		sqlStr = sqlStr + " and "
		sqlStr = sqlStr + " ("
		sqlStr = sqlStr + "  (d.chargediv='4' and c.maeipdiv='W') or (d.chargediv='4' and c.maeipdiv='U')"
		sqlStr = sqlStr + " )"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDesignerItem

					FItemList(i).FShopId       = rsget("shopid")
					FItemList(i).FDesignerId       = rsget("makerid")
					FItemList(i).FDesignerName     = db2html(rsget("socname_kor"))
					FItemList(i).FChargeDiv        = rsget("chargediv")
					FItemList(i).FComm_cd          = rsget("comm_cd")
					FItemList(i).FDefaultMargin    = rsget("defaultmargin")
					FItemList(i).FDefaultSuplyMargin = rsget("defaultsuplymargin")
					FItemList(i).Fonlinedefaultmargine = rsget("onlinedefaultmargine")
					FItemList(i).FSubtract         = rsget("subtract")
					FItemList(i).FRegdate          = rsget("regdate")
					FItemList(i).FAdminopen          = rsget("adminopen")
					FItemList(i).FEtcjunsandetail  = db2html(rsget("etcjunsandetail"))
					FItemList(i).Fisusing = rsget("isusing")
					FItemList(i).FOnlineMWDiv = rsget("maeipdiv")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'//admin/offshop/exchangerate/exchangerate.asp
    public Sub fexchangerate_oneitem()
        dim SqlStr , sqlsearch

        if frectidx <> "" then
        	sqlsearch = sqlsearch & " and idx=" & frectidx & ""
        end if
        if frectcurrencyUnit <> "" then
        	sqlsearch = sqlsearch & " and currencyUnit='" & frectcurrencyUnit & "'"
        end if
        if frectsitename <> "" then
        	sqlsearch = sqlsearch & " and sitename='" & frectsitename & "'"
        end if

        SqlStr = "select"
		sqlStr = sqlStr & " idx, sitename, currencyUnit, exchangeRate, basedate, currencyChar, regdate, lastupdate"
		sqlStr = sqlStr & " , reguserid, lastuserid"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_exchangeRate"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr &"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount

        set FOneItem = new COffShopItem
        if Not rsget.Eof then

			FOneItem.fidx = rsget("idx")
			FOneItem.fsitename = rsget("sitename")
            FOneItem.fcurrencyUnit = rsget("currencyUnit")
            FOneItem.fexchangeRate = rsget("exchangeRate")
            FOneItem.fbasedate = rsget("basedate")
            FOneItem.fcurrencyChar = rsget("currencyChar")
            FOneItem.fregdate = rsget("regdate")
            FOneItem.flastupdate = rsget("lastupdate")
            FOneItem.freguserid = rsget("reguserid")
            FOneItem.flastuserid = rsget("lastuserid")

        end if
        rsget.close
    end Sub

	'//admin/offshop/exchangerate/exchangerate.asp
	public sub fexchangerate_list()
		dim sqlStr,i , sqlsearch

		'총 갯수 구하기
		sqlStr = "select"
		sqlStr = sqlStr & " count(*) as cnt"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_exchangeRate"

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx, sitename, currencyUnit, exchangeRate, basedate, currencyChar, regdate, lastupdate"
		sqlStr = sqlStr & " , reguserid, lastuserid"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_exchangeRate"
		sqlStr = sqlStr & " order by sitename asc, currencyUnit asc"

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
				set FItemList(i) = new COffShopItem

				FItemList(i).fidx = rsget("idx")
				FItemList(i).fsitename = rsget("sitename")
				FItemList(i).fcurrencyUnit = rsget("currencyUnit")
				FItemList(i).fexchangeRate = rsget("exchangeRate")
				FItemList(i).fcurrencyChar = rsget("currencyChar")
				FItemList(i).fbasedate = rsget("basedate")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).flastupdate = rsget("lastupdate")
				FItemList(i).freguserid = rsget("reguserid")
				FItemList(i).flastuserid = rsget("lastuserid")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

    '//admin/lib/popoffshopinfo.asp		'//common/offshop/pop_shopselect_pos.asp
	public Sub GetOffShopList()
		dim sqlStr , i , sqlsearch

		if FRectShopID<>"" then
			sqlsearch = sqlsearch + " and s.userid like '%" + FRectShopID + "%'"
		end if

		if (FRectIsUsing<>"") then
		    sqlsearch = sqlsearch + " and s.isusing ='"&FRectIsUsing&"'"
		end if

		if (FRectShopDiv2<>"") then
		    if (FRectShopDiv2="1") then
		        sqlsearch = sqlsearch + " and s.shopdiv in ('1','2')"
		    elseif (FRectShopDiv2="3") then
		        sqlsearch = sqlsearch + " and s.shopdiv in ('3','4')"
		    elseif (FRectShopDiv2="5") then
		        sqlsearch = sqlsearch + " and s.shopdiv in ('5','6')"
		    elseif (FRectShopDiv2="7") then
		        sqlsearch = sqlsearch + " and s.shopdiv in ('7','8')"
		    elseif (FRectShopDiv2="9") then
		        sqlsearch = sqlsearch + " and s.shopdiv in ('9')"
		    elseif (FRectShopDiv2="11") then
		        sqlsearch = sqlsearch + " and s.shopdiv in ('11','12')"
		    elseif (FRectShopDiv2="13") then
		        sqlsearch = sqlsearch + " and s.shopdiv in ('13','14')"
		    elseif (FRectShopDiv2="15") then
		        sqlsearch = sqlsearch + " and s.shopdiv in ('15','16')"
		    end if
		end if

        if (FRectNotProtoTypeShop<>"") then
            sqlsearch = sqlsearch + " and s.shopdiv not in ('2','4','6','8')"
        end if

		if frectshopname<>"" then
			sqlsearch = sqlsearch + " and s.shopname like '%" + frectshopname + "%'"
		end if
        if frectcurrencyUnit <> "" then
        	sqlsearch = sqlsearch & " and s.currencyUnit='" & frectcurrencyUnit & "'"
        end if
        if FRectloginsite <> "" then
        	sqlsearch = sqlsearch & " and s.loginsite='" & FRectloginsite & "'"
        end if
        if FRectcountrylangcd <> "" then
        	sqlsearch = sqlsearch & " and s.countrylangcd='" & FRectcountrylangcd & "'"
        end if
		if (FRectvieworder<>"") then
			if FRectvieworder="0" then
            	sqlsearch = sqlsearch + " and s.vieworder=0"
			else
				sqlsearch = sqlsearch + " and s.vieworder>=1"
			end if
        end if

		sqlStr = " select top 500 s.*, IsNull(s.ismobileusing, 'N') as ismobileusing, p.groupid, e.countryNamekr"
		sqlStr = sqlStr & " , c.codename as shopdivname"
		'sqlStr = sqlStr & " , oc.codename as loginsitename"
		sqlStr = sqlStr & " , g.company_no, g.company_name"
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_user s"
        sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr & " 	on s.userid=p.id"
        sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner_group g"
        sqlStr = sqlStr & " 	on p.groupid=g.groupid"
        sqlStr = sqlStr & " left join [db_shop].[dbo].[tbl_country_info] e"
        sqlStr = sqlStr & " 	on s.shopCountryCode=e.CountryCode"
        sqlStr = sqlStr & " left join db_shop.dbo.tbl_offshop_commoncode c "
        sqlStr = sqlStr & " 	on c.codekind = 'shopdiv' "
        sqlStr = sqlStr & " 	and c.codeid = s.shopdiv "
        sqlStr = sqlStr & " 	and c.useyn = 'Y' "
'        sqlStr = sqlStr & " left join db_shop.dbo.tbl_offshop_commoncode oc"
'        sqlStr = sqlStr & " 	on s.loginsite=oc.codeid"
'        sqlStr = sqlStr & " 	and oc.useyn='Y'"
'        sqlStr = sqlStr & " 	and oc.codekind='loginsite'"
'        sqlStr = sqlStr & " 	and oc.codegroup='MAIN'"
        sqlStr = sqlStr & " where 1=1 " + sqlsearch
        'sqlStr = sqlStr & " order by s.isusing desc , convert(int,s.shopdiv)+10 asc, s.userid asc"
        sqlStr = sqlStr & " order by convert(int,s.shopdiv)+10 asc, s.userid asc"

		'response.write sqlStr & "<br>"
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
				set FItemList(i) = new COffShopItem

					'FItemList(i).floginsitename         = rsget("loginsitename")
					FItemList(i).fcompany_name         = rsget("company_name")
					FItemList(i).fcompany_no         = rsget("company_no")
					FItemList(i).fcountrylangcd         = rsget("countrylangcd")
					FItemList(i).floginsite         = rsget("loginsite")
					FItemList(i).fadmindisplang         = rsget("admindisplang")
					FItemList(i).fpyeong         = rsget("pyeong")
					FItemList(i).Fuserid         = rsget("userid")
					FItemList(i).Fuserpass       = rsget("userpass")
					FItemList(i).Fshopname       = db2html(rsget("shopname"))
					FItemList(i).Fshopphone      = rsget("shopphone")
					FItemList(i).FshopCountryCode= rsget("shopCountryCode")
					FItemList(i).Fshopzipcode    = rsget("shopzipcode")
					FItemList(i).Fshopaddr1      = db2html(rsget("shopaddr1"))
					FItemList(i).Fshopaddr2      = db2html(rsget("shopaddr2"))
					FItemList(i).Fmanname        = db2html(rsget("manname"))
					FItemList(i).Fmanhp          = rsget("manhp")
					FItemList(i).Fmanphone       = rsget("manphone")
					FItemList(i).Fmanemail       = db2html(rsget("manemail"))
					FItemList(i).Fisusing   	 = rsget("isusing")
					FItemList(i).FShopdiv		 = rsget("shopdiv")
					FItemList(i).FShopdivName	 = rsget("shopdivname")
					FItemList(i).Fstockbasedate  = rsget("stockbasedate")
					FItemList(i).Fshopsocno		= rsget("shopsocno")
					FItemList(i).Fshopceoname  = db2html(rsget("shopceoname"))
					FItemList(i).Fvieworder		=rsget("vieworder")
                    FItemList(i).Fgroupid       =rsget("groupid")
					FItemList(i).fcurrencyUnit	=rsget("currencyUnit")
					FItemList(i).fcurrencyUnit_Pos =rsget("currencyUnit_Pos")
                    FItemList(i).fmultipleRate  =rsget("multipleRate")
                    FItemList(i).FexchangeRate  =rsget("exchangeRate")
                    FItemList(i).FcountryNamekr =db2html(rsget("countryNamekr"))
                    if IsNull(FItemList(i).FshopCountryCode) then
                    	FItemList(i).FshopCountryCode = "KR"
                    end if

                    FItemList(i).FdecimalPointLen = rsget("decimalPointLen")
                    FItemList(i).FdecimalPointCut = rsget("decimalPointCut")

                    FItemList(i).Fismobileusing 	= rsget("ismobileusing")
                    FItemList(i).Fmobileshopname 	= db2html(rsget("mobileshopname"))
                    FItemList(i).Fmobileshopimage 	= rsget("mobileshopimage")
                    FItemList(i).Fmobileworkhour 	= db2html(rsget("mobileworkhour"))
                    FItemList(i).Fmobileclosedate 	= db2html(rsget("mobileclosedate"))
                    FItemList(i).Fmobiletel 		= db2html(rsget("mobiletel"))
                    FItemList(i).Fmobileaddr 		= db2html(rsget("mobileaddr"))
                    FItemList(i).Fmobilemapimage 	= rsget("mobilemapimage")
                    FItemList(i).Fmobilebysubway 	= db2html(rsget("mobilebysubway"))
                    FItemList(i).Fmobilebybus 		= db2html(rsget("mobilebybus"))
                    FItemList(i).Fmobilelatitude 	= rsget("mobilelatitude")
                    FItemList(i).Fmobilelongitude 	= rsget("mobilelongitude")
   					  FItemList(i).Fctropen 	= rsget("ctropen")
   					  if isNull(FItemList(i).Fctropen) then   FItemList(i).Fctropen = 0
					FItemList(i).FViewSort 	= rsget("viewsort")
					FItemList(i).FengName 	= rsget("engName")
					FItemList(i).FShopFax 	= rsget("shopfax")
					FItemList(i).FengAddress 	= rsget("engAddress")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

    '//admin/offshop/othersite/Shoplinkothersite.asp
	public Sub getShoplinkothersitelist()
		dim sqlStr , i , sqlsearch

		if FRectShopID<>"" then
			sqlsearch = sqlsearch + " and u.userid='" + FRectShopID + "'"
		end if

		if (FRectIsUsing<>"") then
		    sqlsearch = sqlsearch + " and u.isusing ='"&FRectIsUsing&"'"
		end if

		if (FRectShopDiv2<>"") then
		    if (FRectShopDiv2="1") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('1')"
		    elseif (FRectShopDiv2="3") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('3')"
		    elseif (FRectShopDiv2="5") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('5')"
		    elseif (FRectShopDiv2="7") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('7')"
		    elseif (FRectShopDiv2="9") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('9')"
		    end if
		end if

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " u.userid as shopid , u.shopname ,u.shopdiv ,u.currencyUnit ,u.currencyUnit_Pos, u.multipleRate ,u.exchangeRate"
		sqlStr = sqlStr + " ,u.isusing ,u.vieworder"
		sqlStr = sqlStr + " ,e.countryNamekr"
		sqlStr = sqlStr + " ,l.siteseq ,l.othershopid ,l.lastupdate ,l.lastadminuserid"
		sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_user u"
        sqlStr = sqlStr + " left join [db_shop].[dbo].[tbl_country_info] e"
        sqlStr = sqlStr + " 	on u.shopCountryCode=e.CountryCode"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_othersitelink l"
		sqlStr = sqlStr + " 	on u.userid = l.shopid"
        sqlStr = sqlStr + " where 1=1 " + sqlsearch
        sqlStr = sqlStr + " order by u.shopdiv asc, u.shopid asc"

        'response.write sqlStr &"<Br>"
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
				set FItemList(i) = new COffShopItem

					FItemList(i).flastupdate  =rsget("lastupdate")
					FItemList(i).flastadminuserid  =rsget("lastadminuserid")
					FItemList(i).fsiteseq  =rsget("siteseq")
					FItemList(i).fothershopid  =rsget("othershopid")
					FItemList(i).fshopid         = rsget("shopid")
					FItemList(i).fshopname         = db2html(rsget("shopname"))
					FItemList(i).fshopdiv         = rsget("shopdiv")
					FItemList(i).fisusing         = rsget("isusing")
					FItemList(i).fvieworder         = rsget("vieworder")
					FItemList(i).fcurrencyUnit	=rsget("currencyUnit")
					FItemList(i).fcurrencyUnit_Pos =rsget("currencyUnit_Pos")
                    FItemList(i).fmultipleRate  =rsget("multipleRate")
                    FItemList(i).FexchangeRate  =rsget("exchangeRate")
		            FItemList(i).FcountryNamekr =db2html(rsget("countryNamekr"))

		            if IsNull(FItemList(i).FshopCountryCode) then
		            	FItemList(i).FshopCountryCode = "KR"
		            end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'//admin/lib/popshopupcheinfo.asp		'//admin/offshop/offupchelist.asp
	public function GetOffShopDesignerList1()
		dim sqlStr,i, sqlsearch

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and c.userid='" + FRectDesigner + "'"
		elseif FRectComm_cd<>"" then
			sqlsearch = sqlsearch + " and s.comm_cd='" + FRectComm_cd + "'"
		elseif FRectChargeType<>"" then
			sqlsearch = sqlsearch + " and s.chargediv='" + FRectChargeType + "'"
		end if
        if FRectShopusing<>"" then
		    sqlsearch = sqlsearch + " and u.isusing='" + FRectShopusing + "'"
		end if
        if (FRectDesigner="") then
    		if FRectPartnerusing<>"" then
    		    sqlsearch = sqlsearch + " and p.isusing='" + FRectPartnerusing + "'"
    		end if
    		if FRectBrandusing<>"" then
    		    sqlsearch = sqlsearch + " and c.isusing='" + FRectBrandusing + "'"
    		end if
		end if
        if (FRectOffUpBea<>"") then
            sqlsearch = sqlsearch + " and s.defaultbeasongdiv=2"
        end if
        if FRectHasContOnly="ON" then
        	sqlsearch = sqlsearch + " and IsNull(s.comm_cd,'') <> '' "
        elseif FRectHasContOnly="OFF" then
        	sqlsearch = sqlsearch + " and IsNull(s.comm_cd,'') = '' "
        elseif FRectHasContOnly<>"" then
            sqlsearch = sqlsearch + " and IsNull(s.comm_cd,'') <> '' "
        end if
		if FRectmaeipdiv <> "" then
		   sqlsearch = sqlsearch + " and c.maeipdiv = '" & FRectmaeipdiv & "' "
		end if
		if FRectBrandPurchaseType<>"" then
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
		if FRectisoffusing<>"" then
		    sqlsearch = sqlsearch + " and c.isoffusing='" + FRectisoffusing + "'"
		end if
		if FRectadminopen<>"" then
		    sqlsearch = sqlsearch + " and sr.adminopen='" + FRectadminopen + "'"
		end if

		sqlStr = " select count(c.userid) as cnt"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c with (nolock)"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user u with (nolock) on u.userid='" + FRectShopid + "'"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s with (nolock) on c.userid=s.makerid and s.shopid=u.userid "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock) on c.userid=p.id"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer sr with (nolock)"
		sqlStr = sqlStr + " 	on c.userid=sr.makerid"
		sqlStr = sqlStr + " 	and sr.shopid='streetshop000'"
		sqlStr = sqlStr + " where c.userdiv in ('02','20') " & sqlsearch

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount<1 then exit function

        sqlStr = " select top " + CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr + " c.userid, c.socname_kor, c.isoffusing"
		sqlStr = sqlStr + " ,u.userid as shopid, u.shopname, u.shopdiv, u.isusing as shopisusing"
		sqlStr = sqlStr + " ,s.chargediv, s.comm_cd, IsNull(s.defaultmargin,0) as defaultmargin "
		sqlStr = sqlStr + " ,IsNull(s.subtract,0) as subtract, s.regdate, s.etcjunsandetail "
		sqlStr = sqlStr + " ,IsNull(s.defaultCenterMwdiv,'') as defaultCenterMwdiv"
		sqlStr = sqlStr + " ,IsNull(c.defaultmargine,0) as onlinedefaultmargine, IsNULL(s.defaultsuplymargin,0) as defaultsuplymargin,  c.isusing as brandisUsing, c.maeipdiv"
		sqlStr = sqlStr + " ,p.isusing as partnerisusing, p.jungsan_date, p.jungsan_date_off, IsNull(s.defaultbeasongdiv,0) as defaultbeasongdiv"
		sqlStr = sqlStr + " ,p.jungsan_gubun, p.purchaseType"
		sqlStr = sqlStr + " ,IsNull(sr.adminopen,'N') as adminopen"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c with (nolock)"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user u with (nolock) on u.userid='" + FRectShopid + "'"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s with (nolock) on c.userid=s.makerid and s.shopid=u.userid "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock) on c.userid=p.id"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer sr with (nolock)"
		sqlStr = sqlStr + " 	on c.userid=sr.makerid"
		sqlStr = sqlStr + " 	and sr.shopid='streetshop000'"
		sqlStr = sqlStr + " where c.userdiv in ('02','20') " & sqlsearch
		sqlStr = sqlStr + " order by c.userid"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new COffShopDesignerItem

				FItemList(i).fpurchaseType           = rsget("purchaseType")
				FItemList(i).fisoffusing           = rsget("isoffusing")
				FItemList(i).FShopId           = rsget("shopid")
				FItemList(i).FShopName         = db2html(rsget("shopname"))
				FItemList(i).FDesignerId       = rsget("userid")
				FItemList(i).FDesignerName     = db2html(rsget("socname_kor"))
				FItemList(i).FChargeDiv        = rsget("chargediv")
				FItemList(i).FComm_cd          = rsget("comm_cd")
				FItemList(i).FDefaultMargin    = rsget("defaultmargin")
				FItemList(i).FDefaultSuplyMargin = rsget("defaultsuplymargin")
				FItemList(i).Fonlinedefaultmargine = rsget("onlinedefaultmargine")
				FItemList(i).FSubtract         		= rsget("subtract")
				FItemList(i).FRegdate          		= rsget("regdate")
				FItemList(i).FAdminopen          	= rsget("adminopen")
				FItemList(i).FEtcjunsandetail   	= db2html(rsget("etcjunsandetail"))
				FItemList(i).FOnlineMWDiv       	= rsget("maeipdiv")
				FItemList(i).Fbrandisusing      	= rsget("brandisUsing")
				FItemList(i).FPartnerisusing    	= rsget("partnerisUsing")
                FItemList(i).FShopisusing       	= rsget("shopisusing")
                FItemList(i).Fjungsan_date_off  	= rsget("jungsan_date_off")
                FItemList(i).Fjungsan_date  		= rsget("jungsan_date")
                FItemList(i).Fshopdiv       		= rsget("shopdiv")
                FItemList(i).FdefaultCenterMwdiv 	= rsget("defaultCenterMwdiv")

                FItemList(i).Fdefaultbeasongdiv 	= rsget("defaultbeasongdiv")
				FItemList(i).Fjungsan_gubun			= rsget("jungsan_gubun")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function

	'//admin/lib/popshopupcheinfo.asp
	public function GetOffShopDesignerList2()
		dim sqlStr,i , sqlsearch

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and c.userid='" + FRectDesigner + "'"
		end if

        if FRectShopusing<>"" then
		    sqlsearch = sqlsearch + " and u.isusing='" + FRectShopusing + "'"
		end if

        if (FRectDesigner="") then
    		if FRectPartnerusing<>"" then
    		    sqlsearch = sqlsearch + " and p.isusing='" + FRectPartnerusing + "'"
    		end if
    		if FRectBrandusing<>"" then
    		    sqlsearch = sqlsearch + " and c.isusing='" + FRectBrandusing + "'"
    		end if
		end if

        if (FRectOffUpBea<>"") then
            sqlsearch = sqlsearch + " and s.defaultbeasongdiv=2"
        end if

        if frectshopdiv <> "" then
        	sqlsearch = sqlsearch + " and co.maincodeid='"&frectshopdiv&"'"
        end if

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " c.userid, c.socname_kor,  c.isusing as brandisUsing, c.maeipdiv"
		sqlStr = sqlStr + " ,u.userid as shopid, u.shopname, u.shopdiv, u.isusing as shopisusing"
		sqlStr = sqlStr + " ,s.chargediv, s.comm_cd, IsNull(s.defaultmargin,0) as defaultmargin "
		sqlStr = sqlStr + " ,IsNull(s.subtract,0) as subtract, s.regdate, IsNull(s.adminopen,'N') as adminopen, s.etcjunsandetail "
		sqlStr = sqlStr + " ,IsNull(c.defaultmargine,0) as onlinedefaultmargine, IsNULL(s.defaultsuplymargin,0) as defaultsuplymargin"
		sqlStr = sqlStr + " ,IsNull(s.defaultCenterMwdiv,'') as defaultCenterMwdiv"
		sqlStr = sqlStr + " ,p.isusing as partnerisusing, p.jungsan_date, p.jungsan_date_off"
		sqlStr = sqlStr + " ,(CASE"
		sqlStr = sqlStr + " 	WHEN Left(u.userid,4)='cafe' then ('streetshop' +  u.userid )"
		sqlStr = sqlStr + " 	WHEN u.userid='streetshop091' THEN 'streetshop991'"
		sqlStr = sqlStr + "  	else u.userid"
		sqlStr = sqlStr + " end) as ordDummi"
		sqlStr = sqlStr + " , IsNull(s.defaultbeasongdiv,0) as defaultbeasongdiv , s.itemregyn, p.purchaseType"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr + " 	on u.userid<>''"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_offshop_commoncode co"
		sqlStr = sqlStr + " 	on codekind = 'shopdiv'"
		sqlStr = sqlStr + " 	and u.shopdiv = co.codeid"
		sqlStr = sqlStr + " 	and co.useyn='Y'"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s"
		sqlStr = sqlStr + " 	on c.userid=s.makerid and s.shopid=u.userid"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p"
		sqlStr = sqlStr + " 	on c.userid=p.id"
		sqlStr = sqlStr + " where c.userdiv in ('02','20') " & sqlsearch
        sqlStr = sqlStr + " order by convert(int,isnull(u.shopdiv,99)) desc,c.userid, ordDummi"

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
        FtotalCount  = FResultCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDesignerItem

				FItemList(i).fitemregyn           = rsget("itemregyn")
				FItemList(i).FShopId           = rsget("shopid")
				FItemList(i).FShopName         = db2html(rsget("shopname"))
				FItemList(i).FDesignerId       = rsget("userid")
				FItemList(i).FDesignerName     = db2html(rsget("socname_kor"))
				FItemList(i).FChargeDiv        = rsget("chargediv")
				FItemList(i).FComm_cd          = rsget("comm_cd")
				FItemList(i).FDefaultMargin    = rsget("defaultmargin")
				FItemList(i).FDefaultSuplyMargin = rsget("defaultsuplymargin")
				FItemList(i).Fonlinedefaultmargine = rsget("onlinedefaultmargine")
				FItemList(i).FSubtract         = rsget("subtract")
				FItemList(i).FRegdate          = rsget("regdate")
				FItemList(i).FAdminopen          = rsget("adminopen")
				FItemList(i).FEtcjunsandetail   = db2html(rsget("etcjunsandetail"))
				FItemList(i).FOnlineMWDiv       = rsget("maeipdiv")
				FItemList(i).Fbrandisusing      = rsget("brandisUsing")
				FItemList(i).FPartnerisusing    = rsget("partnerisUsing")
                FItemList(i).FShopisusing       = rsget("shopisusing")
                FItemList(i).Fjungsan_date_off  = rsget("jungsan_date_off")
                FItemList(i).Fjungsan_date  = rsget("jungsan_date")
                FItemList(i).Fshopdiv       = rsget("shopdiv")
                FItemList(i).FdefaultCenterMwdiv = rsget("defaultCenterMwdiv")
                FItemList(i).Fdefaultbeasongdiv 	= rsget("defaultbeasongdiv")
				FItemList(i).FpurchaseType 	= rsget("purchaseType")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	'//admin/offshop/offupchelist.asp
	public function GetOffShopbrandcontractlist()
		dim sqlStr,i , sqlsearch

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and c.userid='" + FRectDesigner + "'"
		end if
        if FRectShopusing<>"" then
		    sqlsearch = sqlsearch + " and u.isusing='" + FRectShopusing + "'"
		end if
        if (FRectDesigner="") then
    		if FRectPartnerusing<>"" then
    		    sqlsearch = sqlsearch + " and p.isusing='" + FRectPartnerusing + "'"
    		end if
    		if FRectBrandusing<>"" then
    		    sqlsearch = sqlsearch + " and c.isusing='" + FRectBrandusing + "'"
    		end if
		end if
        if (FRectOffUpBea<>"") then
            sqlsearch = sqlsearch + " and s.defaultbeasongdiv=2"
        end if
        if frectshopdiv <> "" then
        	sqlsearch = sqlsearch + " and co.maincodeid='"&frectshopdiv&"'"
        end if
        if FRectHasContOnly="ON" then
        	sqlsearch = sqlsearch + " and IsNull(s.comm_cd,'') <> '' "
        elseif FRectHasContOnly="OFF" then
        	sqlsearch = sqlsearch + " and IsNull(s.comm_cd,'') = '' "
        elseif FRectHasContOnly<>"" then
            sqlsearch = sqlsearch + " and IsNull(s.comm_cd,'') <> '' "
        end if
		if FRectmaeipdiv <> "" then
		   sqlsearch = sqlsearch + " and c.maeipdiv = '" & FRectmaeipdiv & "' "
		end if
		if FRectBrandPurchaseType<>"" then
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
		if FRectisoffusing<>"" then
		    sqlsearch = sqlsearch + " and c.isoffusing='" + FRectisoffusing + "'"
		end if
		if FRectadminopen<>"" then
		    sqlsearch = sqlsearch + " and sr.adminopen='" + FRectadminopen + "'"
		end if

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " c.userid, c.socname_kor,  c.isusing as brandisUsing, c.maeipdiv, c.isoffusing"
		sqlStr = sqlStr + " ,u.userid as shopid, u.shopname, u.shopdiv, u.isusing as shopisusing"
		sqlStr = sqlStr + " ,s.chargediv, s.comm_cd, IsNull(s.defaultmargin,0) as defaultmargin "
		sqlStr = sqlStr + " ,IsNull(s.subtract,0) as subtract, s.regdate, s.etcjunsandetail "
		sqlStr = sqlStr + " ,IsNull(c.defaultmargine,0) as onlinedefaultmargine, IsNULL(s.defaultsuplymargin,0) as defaultsuplymargin"
		sqlStr = sqlStr + " ,IsNull(s.defaultCenterMwdiv,'') as defaultCenterMwdiv"
		sqlStr = sqlStr + " ,p.isusing as partnerisusing, p.jungsan_date, p.jungsan_date_off, p.purchaseType"
		sqlStr = sqlStr + " , IsNull(s.defaultbeasongdiv,0) as defaultbeasongdiv , s.itemregyn"
		sqlStr = sqlStr + " ,(case"
		sqlStr = sqlStr + " 	when convert(int,isnull(u.shopdiv,99))%2=0 then convert(int,isnull(u.shopdiv,99))+10-2"
		sqlStr = sqlStr + " 	else convert(int,isnull(u.shopdiv,99))+10"
		sqlStr = sqlStr + " 	end) as ordertemp"		'/대표샵이 상위에 오기위한 정렬조건
		sqlStr = sqlStr + " ,IsNull(sr.adminopen,'N') as adminopen"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c with (nolock)"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user u with (nolock)"
		sqlStr = sqlStr + " 	on u.userid<>''"
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_offshop_commoncode co with (nolock)"
		sqlStr = sqlStr + " 	on codekind = 'shopdiv'"
		sqlStr = sqlStr + " 	and u.shopdiv = co.codeid"
		sqlStr = sqlStr + " 	and co.useyn='Y'"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s with (nolock)"
		sqlStr = sqlStr + " 	on c.userid=s.makerid and s.shopid=u.userid"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock)"
		sqlStr = sqlStr + " 	on c.userid=p.id"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer sr with (nolock)"
		sqlStr = sqlStr + " 	on c.userid=sr.makerid"
		sqlStr = sqlStr + " 	and sr.shopid='streetshop000'"
		sqlStr = sqlStr + " where c.userdiv in ('02','20') " & sqlsearch
		sqlStr = sqlStr + " order by ordertemp asc"

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
        FtotalCount  = FResultCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDesignerItem

				FItemList(i).fpurchaseType           = rsget("purchaseType")
				FItemList(i).fisoffusing           = rsget("isoffusing")
				FItemList(i).fitemregyn           = rsget("itemregyn")
				FItemList(i).FShopId           = rsget("shopid")
				FItemList(i).FShopName         = db2html(rsget("shopname"))
				FItemList(i).FDesignerId       = rsget("userid")
				FItemList(i).FDesignerName     = db2html(rsget("socname_kor"))
				FItemList(i).FChargeDiv        = rsget("chargediv")
				FItemList(i).FComm_cd          = rsget("comm_cd")
				FItemList(i).FDefaultMargin    = rsget("defaultmargin")
				FItemList(i).FDefaultSuplyMargin = rsget("defaultsuplymargin")
				FItemList(i).Fonlinedefaultmargine = rsget("onlinedefaultmargine")
				FItemList(i).FSubtract         = rsget("subtract")
				FItemList(i).FRegdate          = rsget("regdate")
				FItemList(i).FAdminopen          = rsget("adminopen")
				FItemList(i).FEtcjunsandetail   = db2html(rsget("etcjunsandetail"))
				FItemList(i).FOnlineMWDiv       = rsget("maeipdiv")
				FItemList(i).Fbrandisusing      = rsget("brandisUsing")
				FItemList(i).FPartnerisusing    = rsget("partnerisUsing")
                FItemList(i).FShopisusing       = rsget("shopisusing")
                FItemList(i).Fjungsan_date_off  = rsget("jungsan_date_off")
                FItemList(i).Fjungsan_date  = rsget("jungsan_date")
                FItemList(i).Fshopdiv       = rsget("shopdiv")
                FItemList(i).FdefaultCenterMwdiv = rsget("defaultCenterMwdiv")
                FItemList(i).Fdefaultbeasongdiv 	= rsget("defaultbeasongdiv")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	'//admin/offshop/offupchelist.asp
	public function GetOffShopbrandcontractdiff()
		dim sqlStr,i , sqlsearch

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and c.userid='" + FRectDesigner + "'"
		end if
        if FRectShopusing<>"" then
		    sqlsearch = sqlsearch + " and u.isusing='" + FRectShopusing + "'"
		end if
        if (FRectDesigner="") then
    		if FRectPartnerusing<>"" then
    		    sqlsearch = sqlsearch + " and p.isusing='" + FRectPartnerusing + "'"
    		end if
    		if FRectBrandusing<>"" then
    		    sqlsearch = sqlsearch + " and c.isusing='" + FRectBrandusing + "'"
    		end if
		end if
        if (FRectOffUpBea<>"") then
            sqlsearch = sqlsearch + " and s.defaultbeasongdiv=2"
        end if
        if frectshopdiv <> "" then
        	sqlsearch = sqlsearch + " and co.maincodeid='"&frectshopdiv&"'"
        end if
        if FRectHasContOnly="ON" then
        	sqlsearch = sqlsearch + " and IsNull(s.comm_cd,'') <> '' "
        elseif FRectHasContOnly="OFF" then
        	sqlsearch = sqlsearch + " and IsNull(s.comm_cd,'') = '' "
        elseif FRectHasContOnly<>"" then
            sqlsearch = sqlsearch + " and IsNull(s.comm_cd,'') <> '' "
        end if
		if FRectmaeipdiv <> "" then
		   sqlsearch = sqlsearch + " and c.maeipdiv = '" & FRectmaeipdiv & "' "
		end if
		if FRectBrandPurchaseType<>"" then
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
		if FRectisoffusing<>"" then
		    sqlsearch = sqlsearch + " and c.isoffusing='" + FRectisoffusing + "'"
		end if
		if FRectadminopen<>"" then
		    sqlsearch = sqlsearch + " and sr.adminopen='" + FRectadminopen + "'"
		end if
        if (FRectShopID<>"") then
            sqlsearch = sqlsearch + " and u.userid='"&FRectShopID&"'"
        end if

		sqlStr = " select count(c.userid) as cnt"
		sqlStr = sqlStr & " from [db_user].[dbo].tbl_user_c c with (nolock)" & vbcrlf
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shop_user u with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on u.userid<>''" & vbcrlf
		sqlStr = sqlStr & " 	and u.shopdiv in (3,4)" & vbcrlf	'-- 가맹점,가맹점대표
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_offshop_commoncode co with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on co.codekind = 'shopdiv'" & vbcrlf
		sqlStr = sqlStr & " 	and u.shopdiv = co.codeid" & vbcrlf
		sqlStr = sqlStr & " 	and co.useyn='Y'" & vbcrlf
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_designer sr with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on c.userid=sr.makerid" & vbcrlf
		sqlStr = sqlStr & " 	and sr.shopid='streetshop000'" & vbcrlf
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_designer s with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on c.userid=s.makerid" & vbcrlf
		sqlStr = sqlStr & " 	and s.shopid='streetshop800'" & vbcrlf
		sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner p with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on c.userid=p.id" & vbcrlf
		sqlStr = sqlStr & " where c.userdiv in ('02','20')" & vbcrlf
		sqlStr = sqlStr & " and sr.shopid is not null" & vbcrlf		'-- 직맹점대표계약 있음
		sqlStr = sqlStr & " and s.shopid is null" & sqlsearch		'-- 가맹점대표계약 없음

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount<1 then exit function

		sqlStr = " select top " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " c.userid, c.socname_kor,  c.isusing as brandisUsing, c.maeipdiv, c.isoffusing" & vbcrlf
		sqlStr = sqlStr & " ,u.userid as shopid, u.shopname, u.shopdiv, u.isusing as shopisusing" & vbcrlf
		sqlStr = sqlStr & " ,s.chargediv, s.comm_cd, IsNull(s.defaultmargin,0) as defaultmargin" & vbcrlf
		sqlStr = sqlStr & " ,IsNull(s.subtract,0) as subtract, s.regdate, s.etcjunsandetail" & vbcrlf
		sqlStr = sqlStr & " ,IsNull(c.defaultmargine,0) as onlinedefaultmargine, IsNULL(s.defaultsuplymargin,0) as defaultsuplymargin" & vbcrlf
		sqlStr = sqlStr & " ,IsNull(s.defaultCenterMwdiv,'') as defaultCenterMwdiv" & vbcrlf
		sqlStr = sqlStr & " ,p.isusing as partnerisusing, p.jungsan_date, p.jungsan_date_off, p.purchaseType" & vbcrlf
		sqlStr = sqlStr & " , IsNull(s.defaultbeasongdiv,0) as defaultbeasongdiv , s.itemregyn" & vbcrlf
		sqlStr = sqlStr & " ,IsNull(sr.adminopen,'N') as adminopen" & vbcrlf
		sqlStr = sqlStr & " from [db_user].[dbo].tbl_user_c c with (nolock)" & vbcrlf
		sqlStr = sqlStr & " join [db_shop].[dbo].tbl_shop_user u with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on u.userid<>''" & vbcrlf
		sqlStr = sqlStr & " 	and u.shopdiv in (3,4)" & vbcrlf	'-- 가맹점,가맹점대표
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_offshop_commoncode co with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on co.codekind = 'shopdiv'" & vbcrlf
		sqlStr = sqlStr & " 	and u.shopdiv = co.codeid" & vbcrlf
		sqlStr = sqlStr & " 	and co.useyn='Y'" & vbcrlf
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_designer sr with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on c.userid=sr.makerid" & vbcrlf
		sqlStr = sqlStr & " 	and sr.shopid='streetshop000'" & vbcrlf
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_designer s with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on c.userid=s.makerid" & vbcrlf
		sqlStr = sqlStr & " 	and s.shopid='streetshop800'" & vbcrlf
		sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner p with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on c.userid=p.id" & vbcrlf
		sqlStr = sqlStr & " where c.userdiv in ('02','20')" & vbcrlf
		sqlStr = sqlStr & " and sr.shopid is not null" & vbcrlf		'-- 직맹점대표계약 있음
		sqlStr = sqlStr & " and s.shopid is null" & sqlsearch		'-- 가맹점대표계약 없음
		sqlStr = sqlStr & " order by c.userid asc, u.userid asc" & vbcrlf

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new COffShopDesignerItem

				FItemList(i).fpurchaseType           = rsget("purchaseType")
				FItemList(i).fisoffusing           = rsget("isoffusing")
				FItemList(i).fitemregyn           = rsget("itemregyn")
				FItemList(i).FShopId           = rsget("shopid")
				FItemList(i).FShopName         = db2html(rsget("shopname"))
				FItemList(i).FDesignerId       = rsget("userid")
				FItemList(i).FDesignerName     = db2html(rsget("socname_kor"))
				FItemList(i).FChargeDiv        = rsget("chargediv")
				FItemList(i).FComm_cd          = rsget("comm_cd")
				FItemList(i).FDefaultMargin    = rsget("defaultmargin")
				FItemList(i).FDefaultSuplyMargin = rsget("defaultsuplymargin")
				FItemList(i).Fonlinedefaultmargine = rsget("onlinedefaultmargine")
				FItemList(i).FSubtract         = rsget("subtract")
				FItemList(i).FRegdate          = rsget("regdate")
				FItemList(i).FAdminopen          = rsget("adminopen")
				FItemList(i).FEtcjunsandetail   = db2html(rsget("etcjunsandetail"))
				FItemList(i).FOnlineMWDiv       = rsget("maeipdiv")
				FItemList(i).Fbrandisusing      = rsget("brandisUsing")
				FItemList(i).FPartnerisusing    = rsget("partnerisUsing")
                FItemList(i).FShopisusing       = rsget("shopisusing")
                FItemList(i).Fjungsan_date_off  = rsget("jungsan_date_off")
                FItemList(i).Fjungsan_date  = rsget("jungsan_date")
                FItemList(i).Fshopdiv       = rsget("shopdiv")
                FItemList(i).FdefaultCenterMwdiv = rsget("defaultCenterMwdiv")
                FItemList(i).Fdefaultbeasongdiv 	= rsget("defaultbeasongdiv")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	'//admin/lib/popshopupcheinfo.asp
	public function Getoffshopdivmainlist()
		dim sqlStr,i , sqlsearch

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and c.userid='" + FRectDesigner + "'"
		end if

        if FRectShopusing<>"" then
		    sqlsearch = sqlsearch + " and u.isusing='" + FRectShopusing + "'"
		end if

        if (FRectDesigner="") then
    		if FRectPartnerusing<>"" then
    		    sqlsearch = sqlsearch + " and p.isusing='" + FRectPartnerusing + "'"
    		end if
    		if FRectBrandusing<>"" then
    		    sqlsearch = sqlsearch + " and c.isusing='" + FRectBrandusing + "'"
    		end if
		end if

        if (FRectOffUpBea<>"") then
            sqlsearch = sqlsearch + " and s.defaultbeasongdiv=2"
        end if

		if frectcodegroup <> "" then
			sqlsearch = sqlsearch + " and codegroup = '"&frectcodegroup&"'"
		end if

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " c.userid, c.socname_kor,  c.isusing as brandisUsing, c.maeipdiv"
		sqlStr = sqlStr + " ,u.userid as shopid, u.shopname, u.shopdiv, u.isusing as shopisusing"
		sqlStr = sqlStr + " ,s.chargediv, s.comm_cd, IsNull(s.defaultmargin,0) as defaultmargin "
		sqlStr = sqlStr + " ,IsNull(s.subtract,0) as subtract, s.regdate, IsNull(s.adminopen,'N') as adminopen, s.etcjunsandetail "
		sqlStr = sqlStr + " ,IsNull(c.defaultmargine,0) as onlinedefaultmargine, IsNULL(s.defaultsuplymargin,0) as defaultsuplymargin"
		sqlStr = sqlStr + " ,IsNull(s.defaultCenterMwdiv,'') as defaultCenterMwdiv"
		sqlStr = sqlStr + " ,(CASE"
		sqlStr = sqlStr + " 	WHEN Left(u.userid,4)='cafe' then ('streetshop' +  u.userid )"
		sqlStr = sqlStr + " 	WHEN u.userid='streetshop091' THEN 'streetshop991'"
		sqlStr = sqlStr + "  	else u.userid"
		sqlStr = sqlStr + " end) as ordDummi"
		sqlStr = sqlStr + " , IsNull(s.defaultbeasongdiv,0) as defaultbeasongdiv , s.itemregyn"
		sqlStr = sqlStr + " ,co.codename as shopdivname"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr + " join db_shop.dbo.tbl_offshop_commoncode co"
		sqlStr = sqlStr + " 	on codekind = 'shopdiv'"
		sqlStr = sqlStr + " 	and u.shopdiv = co.codeid"
		sqlStr = sqlStr + " 	and co.useyn='Y'"
		sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " 	on c.userid<>''"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s"
		sqlStr = sqlStr + " 	on c.userid=s.makerid and u.userid = s.shopid"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by co.orderno asc"

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
        FtotalCount  = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDesignerItem

				FItemList(i).fshopdivname   = db2html(rsget("shopdivname"))
				FItemList(i).fitemregyn           = rsget("itemregyn")
				FItemList(i).FShopId           = rsget("shopid")
				FItemList(i).FShopName         = db2html(rsget("shopname"))
				FItemList(i).FDesignerId       = rsget("userid")
				FItemList(i).FDesignerName     = db2html(rsget("socname_kor"))
				FItemList(i).FChargeDiv        = rsget("chargediv")
				FItemList(i).FComm_cd          = rsget("comm_cd")
				FItemList(i).FDefaultMargin    = rsget("defaultmargin")
				FItemList(i).FDefaultSuplyMargin = rsget("defaultsuplymargin")
				FItemList(i).Fonlinedefaultmargine = rsget("onlinedefaultmargine")
				FItemList(i).FSubtract         = rsget("subtract")
				FItemList(i).FRegdate          = rsget("regdate")
				FItemList(i).FAdminopen          = rsget("adminopen")
				FItemList(i).FEtcjunsandetail   = db2html(rsget("etcjunsandetail"))
				FItemList(i).FOnlineMWDiv       = rsget("maeipdiv")
				FItemList(i).Fbrandisusing      = rsget("brandisUsing")
                FItemList(i).FShopisusing       = rsget("shopisusing")
                FItemList(i).Fshopdiv       = rsget("shopdiv")
                FItemList(i).FdefaultCenterMwdiv = rsget("defaultCenterMwdiv")
                FItemList(i).Fdefaultbeasongdiv 	= rsget("defaultbeasongdiv")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

    ''대표마진 불일치
    '///admin/offshop/offupchelist.asp
    public function GetOffShopbrandcontractlisterror()
		dim sqlStr, i, sqlsearch

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and c.userid='" + FRectDesigner + "'"
		elseif FRectComm_cd<>"" then
			sqlsearch = sqlsearch + " and s2.comm_cd='" + FRectComm_cd + "'"
		elseif FRectChargeType<>"" then
			sqlsearch = sqlsearch + " and s2.chargediv='" + FRectChargeType + "'"
		end if
        if FRectShopusing<>"" then
		    sqlsearch = sqlsearch + " and u.isusing='" + FRectShopusing + "'"
		end if
        if (FRectDesigner="") then
    		if FRectPartnerusing<>"" then
    		    sqlsearch = sqlsearch + " and p.isusing='" + FRectPartnerusing + "'"
    		end if
    		if FRectBrandusing<>"" then
    		    sqlsearch = sqlsearch + " and c.isusing='" + FRectBrandusing + "'"
    		end if
		end if
        if (FRectOffUpBea<>"") then
            sqlsearch = sqlsearch + " and s1.defaultbeasongdiv=2"
        end if
        if (FRectShopID<>"") then
            sqlsearch = sqlsearch + " and s1.shopid='"&FRectShopID&"'"
        end if
		if FRectmaeipdiv <> "" then
			sqlsearch = sqlsearch + " and c.maeipdiv = '" & FRectmaeipdiv & "' "
		end if
		if FRectBrandPurchaseType<>"" then
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
		if FRectisoffusing<>"" then
		    sqlsearch = sqlsearch + " and c.isoffusing='" + FRectisoffusing + "'"
		end if
		if FRectadminopen<>"" then
		    sqlsearch = sqlsearch + " and sr.adminopen='" + FRectadminopen + "'"
		end if

        sqlStr = " select count(c.userid) as cnt "
        sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c with (nolock)"
        sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s1 with (nolock) on c.userid=s1.makerid "
        sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s2 with (nolock)"
        sqlStr = sqlStr + " 	on s1.makerid=s2.makerid"
        sqlStr = sqlStr + "  	and ("
        sqlStr = sqlStr + " 		(s1.shopid='streetshop000' and Left(s2.shopid,11)='streetshop0' and s1.shopid<>s2.shopid)"
        sqlStr = sqlStr + " 		or (s1.shopid='streetshop800' and Left(s2.shopid,12)<>'streetshop87' and Left(s2.shopid,11)='streetshop8' and s1.shopid<>s2.shopid)"
        sqlStr = sqlStr + " 		or (s1.shopid='streetshop870' and Left(s2.shopid,12)='streetshop87' and s1.shopid<>s2.shopid)"
        sqlStr = sqlStr + " 	)"
        sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user u with (nolock)"
        sqlStr = sqlStr + " 	on s2.shopid=u.userid"
        sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock)"
        sqlStr = sqlStr + " 	on c.userid=p.id"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer sr with (nolock)"
		sqlStr = sqlStr + " 	on c.userid=sr.makerid"
		sqlStr = sqlStr + " 	and sr.shopid='streetshop000'"
        sqlStr = sqlStr + " where ((s1.comm_cd<>s2.comm_cd) or (s1.defaultmargin<>s2.defaultmargin) or (s1.defaultsuplymargin<>s2.defaultsuplymargin))"
        ''업체위탁,매장매입은 재낌.
        sqlStr = sqlStr + " and s2.comm_cd<>'B012'"
        sqlStr = sqlStr + " and s2.comm_cd<>'B022' " & sqlsearch

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount<1 then exit function

        sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " c.userid, c.socname_kor, c.isoffusing"
		sqlStr = sqlStr + " ,u.userid as shopid, u.shopname, u.shopdiv, u.isusing as shopisusing"
		sqlStr = sqlStr + " ,s2.chargediv, s2.comm_cd, IsNull(s2.defaultmargin,0) as defaultmargin "
		sqlStr = sqlStr + " ,IsNull(s2.subtract,0) as subtract, s2.regdate, s2.etcjunsandetail "
		sqlStr = sqlStr + " ,IsNull(c.defaultmargine,0) as onlinedefaultmargine, IsNULL(s2.defaultsuplymargin,0) as defaultsuplymargin,  c.isusing as brandisUsing, c.maeipdiv"
		sqlStr = sqlStr + " ,p.isusing as partnerisusing, p.jungsan_date, p.jungsan_date_off, p.purchaseType"
		sqlStr = sqlStr + " ,IsNull(sr.adminopen,'N') as adminopen"
        sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c with (nolock)"
        sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s1 with (nolock) on c.userid=s1.makerid "
        sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s2 with (nolock)"
        sqlStr = sqlStr + " 	on s1.makerid=s2.makerid"
        sqlStr = sqlStr + "  	and ("
        sqlStr = sqlStr + " 		(s1.shopid='streetshop000' and Left(s2.shopid,11)='streetshop0' and s1.shopid<>s2.shopid)"
        sqlStr = sqlStr + " 		or (s1.shopid='streetshop800' and Left(s2.shopid,12)<>'streetshop87' and Left(s2.shopid,11)='streetshop8' and s1.shopid<>s2.shopid)"
        sqlStr = sqlStr + " 		or (s1.shopid='streetshop870' and Left(s2.shopid,12)='streetshop87' and s1.shopid<>s2.shopid)"
        sqlStr = sqlStr + " 	)"
        sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user u with (nolock)"
        sqlStr = sqlStr + " 	on s2.shopid=u.userid"
        sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock)"
        sqlStr = sqlStr + " 	on c.userid=p.id"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer sr with (nolock)"
		sqlStr = sqlStr + " 	on c.userid=sr.makerid"
		sqlStr = sqlStr + " 	and sr.shopid='streetshop000'"
        sqlStr = sqlStr + " where ((s1.comm_cd<>s2.comm_cd) or (s1.defaultmargin<>s2.defaultmargin) or (s1.defaultsuplymargin<>s2.defaultsuplymargin))"
        sqlStr = sqlStr + " and s2.comm_cd<>'B012'"
        sqlStr = sqlStr + " and s2.comm_cd<>'B022' " & sqlsearch
        sqlStr = sqlStr + " order by s1.makerid, s1.shopid"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new COffShopDesignerItem

				FItemList(i).fpurchaseType = rsget("purchaseType")
				FItemList(i).fisoffusing           = rsget("isoffusing")
				FItemList(i).FShopId           = rsget("shopid")
				FItemList(i).FShopName         = db2html(rsget("shopname"))
				FItemList(i).FDesignerId       = rsget("userid")
				FItemList(i).FDesignerName     = db2html(rsget("socname_kor"))
				FItemList(i).FChargeDiv        = rsget("chargediv")
				FItemList(i).FComm_cd          = rsget("comm_cd")
				FItemList(i).FDefaultMargin    = rsget("defaultmargin")
				FItemList(i).FDefaultSuplyMargin = rsget("defaultsuplymargin")
				FItemList(i).Fonlinedefaultmargine = rsget("onlinedefaultmargine")
				FItemList(i).FSubtract         = rsget("subtract")
				FItemList(i).FRegdate          = rsget("regdate")
				FItemList(i).FAdminopen          = rsget("adminopen")
				FItemList(i).FEtcjunsandetail   = db2html(rsget("etcjunsandetail"))
				FItemList(i).FOnlineMWDiv       = rsget("maeipdiv")
				FItemList(i).Fbrandisusing      = rsget("brandisUsing")
				FItemList(i).FPartnerisusing    = rsget("partnerisUsing")
                FItemList(i).FShopisusing       = rsget("shopisusing")
                FItemList(i).Fjungsan_date_off  = rsget("jungsan_date_off")
                FItemList(i).Fjungsan_date  = rsget("jungsan_date")
                FItemList(i).Fshopdiv       = rsget("shopdiv")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	public function GetOffShopDesignerList()
		dim sqlStr,i

		sqlStr = " select A.id, A.company_name, A.userid as shopid,"
		sqlStr = sqlStr + " s.chargediv, s.comm_cd, IsNull(s.defaultmargin,0) as defaultmargin, "
		sqlStr = sqlStr + " IsNull(s.subtract,0) as subtract, s.regdate, IsNull(s.adminopen,'N') as adminopen, s.etcjunsandetail "
		sqlStr = sqlStr + " ,IsNull(s.defaultCenterMwdiv,'') as defaultCenterMwdiv"
		sqlStr = sqlStr + " ,IsNull(A.defaultmargine,0) as onlinedefaultmargine, IsNULL(s.defaultsuplymargin,0) as defaultsuplymargin, A.shopdiv"
		sqlStr = sqlStr + " from (select p.id, p.company_name, u.userid, u.shopdiv, p.userdiv, c.defaultmargine from [db_partner].[dbo].tbl_partner p"
		sqlStr = sqlStr + " , [db_user].[dbo].tbl_user_c c "
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user u on u.userid<>''"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and u.userid='" + FRectShopID + "'"
		end if

		sqlStr = sqlStr + " where p.userdiv='9999'"
		sqlStr = sqlStr + " and p.id=c.userid"
		sqlStr = sqlStr + " and c.userdiv in ('02','20')"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and p.id='" + FRectDesigner + "'"
		end if

		sqlStr = sqlStr + " ) as A"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s on A.userid=s.shopid and A.id=s.makerid"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and s.shopid='" + FRectShopID + "'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and s.makerid='" + FRectDesigner + "'"
		end if

        if FRectComm_cd<>"" then
		    sqlStr = sqlStr + " and s.comm_cd='" + FRectComm_cd + "'"
		end if

		if FRectChargeType<>"" then
			sqlStr = sqlStr + " and s.chargediv='" + FRectChargeType + "'"
		end if

		sqlStr = sqlStr + " where A.userdiv='9999'"

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and A.id='" + FRectDesigner + "'"
		end if

		sqlStr = sqlStr + " order by chargediv desc"

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopDesignerItem

				FItemList(i).FShopId       = rsget("shopid")
				FItemList(i).FDesignerId       = rsget("id")
				FItemList(i).FDesignerName     = db2html(rsget("company_name"))
				FItemList(i).FChargeDiv        = rsget("chargediv")
				FItemList(i).FComm_cd          = rsget("comm_cd")
				FItemList(i).FDefaultMargin    = rsget("defaultmargin")
				FItemList(i).FDefaultSuplyMargin = rsget("defaultsuplymargin")
				FItemList(i).Fonlinedefaultmargine = rsget("onlinedefaultmargine")
				FItemList(i).FSubtract         = rsget("subtract")
				FItemList(i).FRegdate          = rsget("regdate")
				FItemList(i).FAdminopen          = rsget("adminopen")
				FItemList(i).FEtcjunsandetail  = db2html(rsget("etcjunsandetail"))
                FItemList(i).Fshopdiv       = rsget("shopdiv")
                FItemList(i).FdefaultCenterMwdiv = rsget("defaultCenterMwdiv")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function

	public function GetChargeUserList()
		dim sqlStr,i

		sqlStr = " select count(chargeuser) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_chargeuser"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " chargeuser, chargename, partneruserid, regdate"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_chargeuser"

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
				set FItemList(i) = new COffShopChargeUserItem

				FItemList(i).FChargeUser        = rsget("chargeuser")
				FItemList(i).FChargeName     	= html2db(rsget("chargename"))
				FItemList(i).FPartnerUserId     = rsget("partneruserid")
				FItemList(i).FregDate     		= rsget("regdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

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
