<%
'####################################################
' Description :  해외 상품 클래스
' History : 2013.05.02 한용민 생성
'####################################################

function drawSelectBoxsitecurrencyunit(sitename, selectBoxName,selectedId,chplg)
	dim tmp_str,query1

	if sitename = "" then exit function

	query1 = " select sitename,currencyunit, currencychar"
	query1 = query1 & " from db_item.dbo.tbl_exchangeRate with (nolock)"
	query1 = query1 & " where sitename='"& sitename &"'"
	query1 = query1 & " order by currencyunit asc"

	'response.write query1 & "<br>"
	rsget.Open query1,dbget,1

	%>
	<select class="select" name="<%=selectBoxName%>" <%=chplg%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>SELECT</option>
	<%
		if  not rsget.EOF  then
		rsget.Movefirst

		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("currencyunit")) then
		tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("currencyunit")&"' "&tmp_str&">"&rsget("currencyunit")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
		end if
	rsget.close
	response.write("</select>")
	'response.write query1 &"<Br>"
end function

function getcountrylang(sitename)
	dim tmp_str,query1

	if sitename = "" then exit function

	query1 = " select sitename, countrylangcd"
	query1 = query1 & " from db_item.dbo.tbl_exchangeRate with (nolock)"
	query1 = query1 & " where sitename='"& sitename &"'"
	query1 = query1 & " group by sitename, countrylangcd"

	'response.write query1 & "<br>"
	rsget.Open query1,dbget,1
	if  not rsget.EOF  then	
		getcountrylang = rsget.getrows()
	end if
	rsget.close
end function

function getMultiSiteSitenameByCode(iVal)
    if (iVal="WSLWEB") then
        getMultiSiteSitenameByCode = "wholesale.10x10.co.kr"
    elseif (iVal="TPL") then
        getMultiSiteSitenameByCode = "3pl.10x10shop.com"
    elseif (iVal="CHNWEB") then
        getMultiSiteSitenameByCode = "www.10x10shop.com"
    elseif (iVal="ITSWEB") then
        getMultiSiteSitenameByCode = "www.ithinksoweb.com"
    elseif (iVal="11STMY") then
        getMultiSiteSitenameByCode = "www.11st.my"
    elseif (iVal="ZILINGO") then
        getMultiSiteSitenameByCode = "zilingo.com"
    elseif (iVal="ETSY") then
        getMultiSiteSitenameByCode = "etsy.com"
    elseif (iVal="SHOPIFY") then
        getMultiSiteSitenameByCode = "www.shopify.com"
    end if
end function

'//해외언어		'/2013.05.02 한용민 생성
function drawSelectboxMultiLangCountrycd(selBoxName, selVal, chplg)
%>
	<select name="<%= selBoxName %>" <%= chplg %> class="select">
		<option value="" <% if selVal = "" then response.write " selected" %> >SELECT</option>
		<option value="KR" <% if selVal = "KR" then response.write " selected" %> >KR</option>
		<option value="EN" <% if selVal = "EN" then response.write " selected" %> >EN</option>
		<option value="CN" <% if selVal = "CN" then response.write " selected" %> >CN</option>
		<option value="JP" <% if selVal = "JP" then response.write " selected" %> >JP</option>
		<option value="ITSWEB" <% if selVal = "ITSWEB" then response.write " selected" %> >ITSWEB Only</option>
	</select>
<%
end function

'//해외사이트구분		'/2013.05.02 한용민 생성
function drawSelectboxMultiSiteSitename(selBoxName, selVal, chplg)
%>
	<select name="<%= selBoxName %>" <%= chplg %> class="select">
		<option value="" <% if selVal = "" then response.write " selected" %> >SELECT</option>
		<option value="WSLWEB" <% if selVal = "WSLWEB" then response.write " selected" %> >wholesale.10x10.co.kr</option>
		<option value="TPL" <% if selVal = "TPL" then response.write " selected" %> >3pl.10x10.co.kr</option>
		<option value="CHNWEB" <% if selVal = "CHNWEB" then response.write " selected" %> >www.10x10shop.com</option>
		<option value="ITSWEB" <% if selVal = "ITSWEB" then response.write " selected" %> >www.ithinksoweb.com</option>
		<option value="11STMY" <% if selVal = "11STMY" then response.write " selected" %> >www.11st.my</option>
		<option value="ZILINGO" <% if selVal = "ZILINGO" then response.write " selected" %> >zilingo.com</option>
		<option value="ETSY" <% if selVal = "ETSY" then response.write " selected" %> >etsy.com</option>
		<option value="SHOPIFY" <% if selVal = "SHOPIFY" then response.write " selected" %> >www.shopify.com</option>
	</select>
<%
end function

Class COverseasItemDetail
	public fitemoption_wholesale
	public fsitename
	public fforeignorgprice
	public foff_isusing
	public Foptisusing
	public Foptsellyn
	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold
	public Foptionname
	public Foptiontypename
	public Frealstock
	public Fipkumdiv2
	public Fipkumdiv4
	public Fipkumdiv5
	public Foffconfirmno
	public FLastUpdate
	public Frestockdate
	public Foptaddprice
	public Fbarcode
	public Fupchemanagecode
	public fcurrencyUnit
	public fcurrencyChar
	public fexchangeRate
    public Fitemid
    public Fmakerid
    public FCate_large
    public FCate_mid
    public FCate_small
    public Fitemdiv
    public Fitemgubun
    public Fitemname
    public Fitemname10x10
    public Fsellcash
    public Fbuycash
    public Forgprice
    public Forgsuplycash
    public Fsailprice
    public Fsailsuplycash
    public Fmileage
    public Fregdate
    public FsellEndDate
    public Fsellyn
    public Flimityn
    public Fdanjongyn
    public Fsailyn
    public Fisusing
    public Fisextusing
    public Fmwdiv
    public Fspecialuseritem
    public Fvatinclude
    public Fdeliverytype
    public Fdeliverarea
    public Fdeliverfixday
    public Fismobileitem
    public Fpojangok
    public Flimitno
    public Flimitsold
    public Fevalcnt
    public Foptioncnt
    public Fitemrackcode
    public FReIpgodate
    public Fbrandname
    public Ftitleimage
    public Fmainimage
    public Fmainimage2
    public Fsmallimage
    public Flistimage
    public Flistimage120
    public Fbasicimage
    public Fbasicimage600
    public Ficon1image
    public Ficon2image
    public Fitemcouponyn
    public Fcurritemcouponidx
    public Fitemcoupontype
    public Fitemcouponvalue
    public Fitemcopy
    public FExistMultiLang
    public FRegUserID
    public FavailPayType
    public FtenOnlyYn
    public FfrontMakerid		'프론트표시용 그룹 브랜드ID
	public fuseyn
	public fitemname_en
	public foptiontypename_en
	public foptionname_en
	public fitemsource_en
	public fsourcearea_en
	public fitemcopy_en
	public fitemsize_en
	public fmakername_en
	public fkeywords_en
	public fitemname_10x10
	public foptiontypename_10x10
	public foptionname_10x10
	public fitemsource_10x10
	public fsourcearea_10x10
	public fitemcopy_10x10
	public fitemsize_10x10
	public fmakername_10x10
	public fkeywords_10x10
	public F10x10itemoption
	public F10x10optionname
	public F10x10optisusing
	public F10x10optiontypename
	public FNotReg
	public fcountrycd
	public fisusing_off

    ''tbl_item_Contents
    public Fkeywords
    public Fsourcearea
    public Fmakername
    public Fitemsource
    public Fitemsize
	public FitemWeight
    public Fusinghtml
    public Fitemcontent
    public Fordercomment
    public Fdesignercomment
    public Fsellcount
    public Ffavcount
    public Frecentsellcount
    public Frecentfavcount
    public Frecentpoints
    public Frecentpcount
	public FrequireMakeDay
    public FreserveItemTp       ''단독(예약)구매상품
    public Flinkurl

    ''Etc
    public Fcouponbuyprice
    public FCate_large_Name
    public FCate_Mid_Name
    public FCate_Small_Name
    public FinfoDiv			'품번구분번호
    public FsafetyYn		'안정인증대상 여부
    public FsafetyDiv
    public FsafetyNum
    public FinfoimageExists

    '' 기본 배송비 정책 관련 tbl_user_c
    public FdefaultFreeBeasongLimit
    public FdefaultDeliverPay
    public FdefaultDeliveryType
	public FdeliverOverseas
    public FavgDLvDate
    public Fitemoption

    ''tbl_item_multiSite_regItem
    public fSiteisusing
    public FareaCode11st
    public fmultiplerate
    public flinkPriceType
	public fwonprice
	public fpricearr
	public fcountryLangCDarr
	public fmultilangcnt

    public function getlinkPriceTypeName()
        if isNULL(flinkPriceType) then Exit function

        if (flinkPriceType=1) then
            getlinkPriceTypeName = "SellPrice"
        elseif (flinkPriceType=2) then
            getlinkPriceTypeName = "OrgPrice"
        end if
    end function

    public function IsReserveOnlyItem ''단독(예약)구매상품
        IsReserveOnlyItem = false
        if IsNULL(FreserveItemTp) THEN Exit function
        IsReserveOnlyItem = (FreserveItemTp=1)
    end function

    public Function IsSoldOut()
		IsSoldOut = (FSellYn<>"Y") or ((FLimitYn="Y") and (GetLimitEa()<1))
	end function

    public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	public function GetLimitStockNo()
		GetLimitStockNo = GetCheckStockNo + Fipkumdiv4 + Fipkumdiv2
	end function

	public function GetCheckStockNo()
		GetCheckStockNo = Frealstock + GetTodayBaljuNo
	end function

	public function GetTodayBaljuNo()
		GetTodayBaljuNo = Fipkumdiv5 + Foffconfirmno
	end function

    public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" or Fdeliverytype="9" or Fdeliverytype="7" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
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

	'// 상품 쿠폰 내용
	public function GetCouponDiscountStr() '!
		Select Case Fitemcoupontype
			Case "1"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "%"
			Case "2"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "원"
			Case "3"
				GetCouponDiscountStr ="무료배송"
			Case Else
				GetCouponDiscountStr = Fitemcoupontype
		End Select
	end function

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

	'// 우수회원샵 상품 여부
	public Function IsSpecialUserItem() '!
	    dim uLevel
	    ''uLevel = GetLoginUserLevel()
		IsSpecialUserItem = (FSpecialUserItem>0) ''and (uLevel>0 and uLevel<>5)
	end Function

	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "특정"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function

	''재입고 상품 여부 (7일)
	public function IsReIpgoItem()
	    IsReIpgoItem = False
	    if IsNULL(FReIpgodate) then Exit Function

	    IsReIpgoItem = DateDiff("d",FReIpgodate,now())<8

    end function

	'# 상품의 진행중인 할인코드 접수
	public Sub getSeleCode(byREF saleCode, byREF saleName)
		Dim strSql
		strSql = "select sm.sale_code, sm.sale_name " &_
				" from db_event.dbo.tbl_sale as sm with (nolock) " &_
				" 	join db_event.dbo.tbl_saleItem as si with (nolock) " &_
				" 		on sm.sale_code=si.sale_code " &_
				" where si.itemid=" & Fitemid &_
				" 	and getdate() between sm.sale_startdate and dateadd(d,1,sm.sale_enddate) "
		rsget.Open strSql,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			saleCode = rsget("sale_code")
			saleName = rsget("sale_name")
		end if
		rsget.Close
	end Sub

    Private Sub Class_Initialize()
        Foptioncnt = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COverSeasItem
    public FOneItem
	public FItemList()
	public farrlist
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public frectcolorcode
	public FRectMakerid
    public FRectItemID
    public FRectItemName
	Public FRectMultiLanguage
    public FRectSellYN
    public FRectIsUsing
    public FRectDanjongYN
    public FRectMWDiv
    public FRectLimitYN
	public FRectVatYN
	public FRectSailYN
	public FRectCouponYN
	public FRectDeliveryType
	public FRectIsOversea
	public FRectIsWeight
	public FRectRackcode
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public FRectSortDiv
	public FRectSortDiv2
	public FRectItemDiv
	public FRectSellcash1
	public FRectSellcash2
	public FRectMinusMigin
	public FRectMarginUP
	public FRectMarginDown
	public FRectItemGubun
	public FRectNoBarcode
	public FRectNoUpcheBarcode
	public FRectCountryCd
	public FRectLangGubun
	public FRectRegUserID
	public FRectRegDate1
	public FRectRegDate2
	public FRectIsReg
	public FRectuseyn
    public FRectSitename
	Public FRectShopid
	public frectcurrencyunit
	public FRectsiteisusing

    public function getSiteCountryLangCD(isitename)
        dim sqlStr
        sqlStr = "select top 1 countryLangCd from db_item.dbo.tbl_exchangeRate with (nolock)"
        sqlStr = sqlStr & " where sitename='"&isitename&"'"

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1

        if not rsget.Eof then
            getSiteCountryLangCD = rsget("countryLangCd")
        end if

        rsget.Close
    end function

	'/admin/itemmaster/overseas/itemlist.asp/
	'/이거 수정시 GetOverSeasItemList_excel 이펑션도 고쳐야함
	public function GetForeignItemList()
        dim sqlStr, addSql, i
        
        if FRectSitename="" or isnull(FRectSitename) then exit function

		if FRectCountryCd<>"" then
			if FRectCountryCd="o" then
		        addSql = addSql & " and uu.sitename is not null"
			elseif FRectCountryCd="x" then
				addSql = addSql & " and uu.sitename is null"
			end if
		end if
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if
        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if
        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if
        if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if
        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if
        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if
        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if
        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if
        if (FRectLimitYN <> "") then
            addSql = addSql & " and i.limityn='" + FRectLimitYN + "'"
        end if
        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
        end if
        if FRectIsWeight<>"" then
			if FRectIsWeight="Y" then
				addSql = addSql + " and isnull(i.itemWeight,0)>0 "
			else
				addSql = addSql + " and isnull(i.itemWeight,0)<=0 "
			end if
		end if
		IF FRectSellcash1 <> "" Then
			addSql = addSql + " and i.sellcash >= '" & FRectSellcash1 & "' "
		End IF
		IF FRectSellcash2 <> "" Then
			addSql = addSql + " and i.sellcash <= '" & FRectSellcash2 & "' "
		End IF
		If FRectRegDate1 <> "" Then
			addSql = addSql + " and uu.lastupdate >= '" & FRectRegDate1 & "' "
		End If
        if (FRectsiteisusing <> "") then
            addSql = addSql & " and uu.isusing='" & FRectsiteisusing & "'"
        end if

		sqlStr = "create table #tmp_exchangeRatecurrencyunitgroup ("
        sqlStr = sqlStr & " 	sitename nvarchar(32)"
        sqlStr = sqlStr & " 	,currencyunit nvarchar(16)"
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " insert into #tmp_exchangeRatecurrencyunitgroup"
        sqlStr = sqlStr & " 	select sitename, currencyunit"
        sqlStr = sqlStr & " 	from db_item.dbo.tbl_exchangeRate with (nolock)"
        sqlStr = sqlStr & " 	where sitename='"& FRectSitename &"'"
        sqlStr = sqlStr & " 	group by sitename, currencyunit"

		'response.write sqlStr & "<br>"
		dbget.execute sqlStr

		sqlStr = "create table #tmp_exchangeRatecountryLangCDgroup ("
        sqlStr = sqlStr & " 	sitename nvarchar(32)"
        sqlStr = sqlStr & " 	,countryLangCD varchar(32)"
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " insert into #tmp_exchangeRatecountryLangCDgroup"
        sqlStr = sqlStr & " 	select sitename, countryLangCD"
        sqlStr = sqlStr & " 	from db_item.dbo.tbl_exchangeRate with (nolock)"
        sqlStr = sqlStr & " 	where sitename='"& FRectSitename &"'"
        sqlStr = sqlStr & " 	group by sitename, countryLangCD"

		'response.write sqlStr & "<br>"
		dbget.execute sqlStr

		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i with (nolock)"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_multiSite_regItem] uu with (nolock)"
        sqlStr = sqlStr & " 	on i.itemid=uu.itemid AND uu.sitename = '"&FRectSitename&"'"

		if FRectCountryCd <> "" then
			if FRectLangGubun <> "" then
				sqlStr = sqlStr & " join db_item.dbo.tbl_item_multiLang ml with (nolock)"
				sqlStr = sqlStr & " 	on i.itemid = ml.itemid"
				sqlStr = sqlStr & " 	and ml.countryCd='"& FRectLangGubun &"'"
			end if
			if FRectCountryCd="x" then
				if FRectSitename="WSLWEB" then
					sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_designer sd with (nolock)"
					sqlStr = sqlStr & " 	on i.makerid = sd.makerid"
					'sqlStr = sqlStr & " 	and sd.shopid='streetshop700'"		'/홀쎄일은 해외대표계약있는것만
					sqlStr = sqlStr & " 	and sd.shopid='streetshop000'"		'/홀쎄일은 직영대표계약있는것만
				end if
			end if
		end if

        sqlStr = sqlStr & " where i.itemid<>0" & addSql		'/and i.deliverytype in(1,4) and (i.mwdiv='M' or i.mwdiv='W')

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		dbget.CommandTimeout = 60*1   ' 1분
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
        rsget.Close

        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " uu.multilangcnt, i.itemname as itemname10x10, i.*"
        sqlStr = sqlStr & " ,isnull(uu.isusing,'') as siteisusing"		', ml.useyn
        sqlStr = sqlStr & " ,substring(STUFF(("
        sqlStr = sqlStr & " 	SELECT Top 100 '|^|' + ee.currencyUnit + '|*|' + cast(mpp.orgprice as varchar(10)) + '|*|'"
        sqlStr = sqlStr & " 	+ cast(mpp.wonprice as varchar(10)) + '|*|' + cast(mpp.lastexchangeRate as varchar(10))"
        sqlStr = sqlStr & " 	FROM #tmp_exchangeRatecurrencyunitgroup ee"
        sqlStr = sqlStr & " 	join db_item.dbo.tbl_item_multiLang_price mpp with (nolock)"
        sqlStr = sqlStr & " 		on ee.sitename=mpp.sitename"
        sqlStr = sqlStr & " 		and ee.currencyunit = mpp.currencyUnit"
        sqlStr = sqlStr & " 	WHERE ee.sitename = uu.sitename and mpp.itemid = uu.itemid"
        sqlStr = sqlStr & " 	ORDER BY ee.currencyUnit asc"
        sqlStr = sqlStr & " FOR XML PATH('')), 1, 1, ''),3,4000) as pricearr"
        sqlStr = sqlStr & " ,substring(STUFF(("
        sqlStr = sqlStr & " 	SELECT Top 100 '|^|' + ee.countryLangCD + '|*|' + mll.itemname"
        sqlStr = sqlStr & " 	FROM #tmp_exchangeRatecountryLangCDgroup ee"
        sqlStr = sqlStr & " 	join [db_item].[dbo].[tbl_item_multiLang] mll with (nolock)"
        sqlStr = sqlStr & " 		on ee.countryLangCD = mll.countryCd"
        sqlStr = sqlStr & " 		and mll.useyn='Y'"
        sqlStr = sqlStr & " 	WHERE ee.sitename = uu.sitename and uu.itemid=mll.itemid"
        sqlStr = sqlStr & " 	ORDER BY ee.countryLangCD asc"
        sqlStr = sqlStr & " FOR XML PATH('')), 1, 1, ''),3,4000) as countryLangCDarr"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i with (nolock)"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_multiSite_regItem] uu with (nolock)"
        sqlStr = sqlStr & " 	on i.itemid=uu.itemid AND uu.sitename = '"&FRectSitename&"'"

		if FRectCountryCd <> "" then
			if FRectLangGubun <> "" then
				sqlStr = sqlStr & " join db_item.dbo.tbl_item_multiLang ml with (nolock)"
				sqlStr = sqlStr & " 	on i.itemid = ml.itemid"
				sqlStr = sqlStr & " 	and ml.countryCd='"& FRectLangGubun &"'"
			end if
			if FRectCountryCd="x" then
				if FRectSitename="WSLWEB" then
					sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_designer sd with (nolock)"
					sqlStr = sqlStr & " 	on i.makerid = sd.makerid"
					'sqlStr = sqlStr & " 	and sd.shopid='streetshop700'"		'/홀쎄일은 해외대표계약있는것만
					sqlStr = sqlStr & " 	and sd.shopid='streetshop000'"		'/홀쎄일은 직영대표계약있는것만
				end if
			end if
		end if

        sqlStr = sqlStr & " where i.itemid<>0" & addSql		'/and i.deliverytype in(1,4) and (i.mwdiv='M' or i.mwdiv='W')

		IF FRectSortDiv="new" Then
			sqlStr = sqlStr & " Order by i.itemid desc"
		ElseIf FRectSortDiv = "best" Then
			sqlStr = sqlStr & " ORDER BY i.itemScore DESC, i.itemid desc"
		ElseIf FRectSortDiv = "min" Then
			sqlStr = sqlStr & " ORDER BY i.sellcash ASC, i.itemid desc"
		ElseIf FRectSortDiv = "hi" Then
			sqlStr = sqlStr & " ORDER BY i.sellcash DESC, i.itemid desc"
		ElseIf FRectSortDiv = "hs" Then
			sqlStr = sqlStr & " ORDER BY i.orgprice-i.sellcash DESC, i.itemid desc"
		ELSEIF FRectSortDiv="cashH" Then
			sqlStr = sqlStr & " Order by i.SellCash desc, i.itemid desc"
		ELSEIF FRectSortDiv="cashL" Then
			sqlStr = sqlStr & " Order by i.SellCash asc, i.itemid desc"
		ELSEIF FRectSortDiv="best" Then
			sqlStr = sqlStr & " Order by i.ItemScore desc, i.itemid desc"
		ELSEIF FRectSortDiv = "weightup" Then
			sqlStr = sqlStr & " Order by i.itemWeight desc, i.itemid desc"
		ElseIf FRectSortDiv = "weightdown" Then
			sqlStr = sqlStr & " Order by i.itemWeight asc, i.itemid desc"
		ELSE
			sqlStr = sqlStr & " Order by i.itemid desc"
		End IF

		'response.write sqlStr & "<br>"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		dbget.CommandTimeout = 60*1   ' 1분
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new COverseasItemDetail

				FItemList(i).fmultilangcnt	= rsget("multilangcnt")
				FItemList(i).fcountryLangCDarr	= rsget("countryLangCDarr")
				FItemList(i).fpricearr			= rsget("pricearr")
				FItemList(i).fsiteisusing            = rsget("siteisusing")
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname10x10		= db2html(rsget("itemname10x10"))
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
                'FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	'/이거 수정시 GetForeignItemList 이펑션도 고쳐야함
	public function GetOverSeasItemList_excel()
        dim sqlStr, addSql, i

        if FRectSitename="" or isnull(FRectSitename) then exit function

		if FRectCountryCd<>"" then
			if FRectCountryCd="o" then
		        addSql = addSql & " and uu.sitename is not null"
			elseif FRectCountryCd="x" then
				addSql = addSql & " and uu.sitename is null"
			end if
		end if
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if
        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if
        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if
        if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if
        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if
        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if
        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if
        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if
        if (FRectLimitYN <> "") then
            addSql = addSql & " and i.limityn='" + FRectLimitYN + "'"
        end if
        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
        end if
        if FRectIsWeight<>"" then
			if FRectIsWeight="Y" then
				addSql = addSql + " and isnull(i.itemWeight,0)>0 "
			else
				addSql = addSql + " and isnull(i.itemWeight,0)<=0 "
			end if
		end if
		IF FRectSellcash1 <> "" Then
			addSql = addSql + " and i.sellcash >= '" & FRectSellcash1 & "' "
		End IF
		IF FRectSellcash2 <> "" Then
			addSql = addSql + " and i.sellcash <= '" & FRectSellcash2 & "' "
		End IF
		If FRectRegDate1 <> "" Then
			addSql = addSql + " and uu.lastupdate >= '" & FRectRegDate1 & "' "
		End If
        if (FRectsiteisusing <> "") then
            addSql = addSql & " and uu.isusing='" & FRectsiteisusing & "'"
        end if

        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.itemid, isnull(o.itemoption,'0000') as itemoption"
        sqlStr = sqlStr & " , replace(replace(replace(replace(i.itemname,char(9),''),char(10),''),char(13),''),'""','''') as itemname_10x10"
		sqlStr = sqlStr & " , i.makerid, replace(replace(replace(replace(o.optiontypename,char(9),''),char(10),''),char(13),''),'""','''') as optiontypename_10x10"
		sqlStr = sqlStr & " , replace(replace(replace(replace(o.optionname,char(9),''),char(10),''),char(13),''),'""','''') as optionname_10x10"
        sqlStr = sqlStr & " , ic.itemsource as itemsource_10x10, ic.sourcearea as sourcearea_10x10" & vbcrlf
        sqlStr = sqlStr & " , ic.designercomment as itemcopy_10x10, ic.itemsize as itemsize_10x10, ic.makername as makername_10x10, ic.keywords as keywords_10x10" & vbcrlf
        sqlStr = sqlStr & " , replace(replace(replace(replace(mle.itemname,char(9),''),char(10),''),char(13),''),'""','''') as itemname_en"
		sqlStr = sqlStr & " , replace(replace(replace(replace(mo.optiontypename,char(9),''),char(10),''),char(13),''),'""','''') as optiontypename_en"
		sqlStr = sqlStr & " , replace(replace(replace(replace(mo.optionname,char(9),''),char(10),''),char(13),''),'""','''') as optionname_en"
		sqlStr = sqlStr & " , mle.itemsource as itemsource_en, mle.sourcearea as sourcearea_en" & vbcrlf
        sqlStr = sqlStr & " , mle.itemcopy as itemcopy_en, mle.itemsize as itemsize_en, mle.makername as makername_en, mle.keywords as keywords_en" & vbcrlf
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i with (nolock)"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_multiSite_regItem] uu with (nolock)"
        sqlStr = sqlStr & " 	on i.itemid=uu.itemid AND uu.sitename = '"&FRectSitename&"'"

		if FRectCountryCd<>"" then
			if not(FRectCountryCd="o" or FRectCountryCd="x") then
				sqlStr = sqlStr & " join db_item.dbo.tbl_item_multiLang ml with (nolock)"
				sqlStr = sqlStr & " 	on i.itemid = ml.itemid"
				sqlStr = sqlStr & " 	and ml.countryCd='"& FRectCountryCd &"'"
			end if
			if FRectCountryCd="x" then
				if FRectSitename="WSLWEB" then
					sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_designer sd with (nolock)"
					sqlStr = sqlStr & " 	on i.makerid = sd.makerid"
					'sqlStr = sqlStr & " 	and sd.shopid='streetshop700'"		'/홀쎄일은 해외대표계약있는것만
					sqlStr = sqlStr & " 	and sd.shopid='streetshop000'"		'/홀쎄일은 직영대표계약있는것만
				end if
			end if
		end if

        sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c with (nolock) "
        sqlStr = sqlStr & " 	on i.makerid=c.userid"
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_option o with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on i.itemid = o.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and o.isusing='Y'" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_Contents ic with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on i.itemid = ic.itemid" & vbcrlf
		sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang] mle with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on i.itemid=mle.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and mle.countryCd='EN'" & vbcrlf
		sqlStr = sqlStr & " 	and mle.useyn='Y'" & vbcrlf
		sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang_option] mo with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on i.itemid=mo.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and o.itemoption = mo.itemoption" & vbcrlf
		sqlStr = sqlStr & " 	and mo.countryCd='EN'" & vbcrlf
        sqlStr = sqlStr & " where i.itemid<>0 " & addSql	'/and i.deliverytype in(1,4) and (i.mwdiv='M' or i.mwdiv='W')
       	sqlStr = sqlStr & " order by i.itemid desc"

		'response.write sqlStr & "<br>"
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new COverseasItemDetail
				FItemList(i).Fitemgubun = "10"
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fitemoption            = rsget("itemoption")
                FItemList(i).fmakerid            = rsget("makerid")
				FItemList(i).fitemname_10x10    		= db2html(rsget("itemname_10x10"))
				FItemList(i).foptiontypename_10x10    		= db2html(rsget("optiontypename_10x10"))
				FItemList(i).foptionname_10x10    		= db2html(rsget("optionname_10x10"))
				FItemList(i).fitemsource_10x10    		= db2html(rsget("itemsource_10x10"))
				FItemList(i).fsourcearea_10x10    		= db2html(rsget("sourcearea_10x10"))
				FItemList(i).fitemcopy_10x10    		= db2html(rsget("itemcopy_10x10"))
				FItemList(i).fitemsize_10x10    		= db2html(rsget("itemsize_10x10"))
				FItemList(i).fmakername_10x10    		= db2html(rsget("makername_10x10"))
				FItemList(i).fkeywords_10x10    		= db2html(rsget("keywords_10x10"))
				FItemList(i).fitemname_en    		= db2html(rsget("itemname_en"))
				FItemList(i).foptiontypename_en    		= db2html(rsget("optiontypename_en"))
				FItemList(i).foptionname_en    		= db2html(rsget("optionname_en"))
				FItemList(i).fitemsource_en    		= db2html(rsget("itemsource_en"))
				FItemList(i).fsourcearea_en    		= db2html(rsget("sourcearea_en"))
				FItemList(i).fitemcopy_en    		= db2html(rsget("itemcopy_en"))
				FItemList(i).fitemsize_en    		= db2html(rsget("itemsize_en"))
				FItemList(i).fmakername_en    		= db2html(rsget("makername_en"))
				FItemList(i).fkeywords_en    		= db2html(rsget("keywords_en"))
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	' 사용안함
	public function GetOverSeasItemList()
        dim sqlStr, addSql, i

        if FRectSitename="" then exit function

        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if

        if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if (FRectLimitYN <> "") then
            addSql = addSql & " and i.limityn='" + FRectLimitYN + "'"
        end if

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

		sqlStr = "create table #tmp_exchangeRatecurrencyunitgroup ("
        sqlStr = sqlStr & " 	sitename nvarchar(32)"
        sqlStr = sqlStr & " 	,currencyunit nvarchar(16)"
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " insert into #tmp_exchangeRatecurrencyunitgroup"
        sqlStr = sqlStr & " 	select sitename, currencyunit"
        sqlStr = sqlStr & " 	from db_item.dbo.tbl_exchangeRate with (nolock)"
        sqlStr = sqlStr & " 	where sitename='"& FRectSitename &"'"
        sqlStr = sqlStr & " 	group by sitename, currencyunit"

		'response.write sqlStr & "<br>"
		dbget.execute sqlStr

		sqlStr = "create table #tmp_exchangeRatecountryLangCDgroup ("
        sqlStr = sqlStr & " 	sitename nvarchar(32)"
        sqlStr = sqlStr & " 	,countryLangCD varchar(32)"
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " insert into #tmp_exchangeRatecountryLangCDgroup"
        sqlStr = sqlStr & " 	select sitename, countryLangCD"
        sqlStr = sqlStr & " 	from db_item.dbo.tbl_exchangeRate with (nolock)"
        sqlStr = sqlStr & " 	where sitename='"& FRectSitename &"'"
        sqlStr = sqlStr & " 	group by sitename, countryLangCD"

		'response.write sqlStr & "<br>"
		dbget.execute sqlStr

		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].[tbl_item_multiSite_regItem] uu with (nolock)"
        sqlStr = sqlStr & " join [db_item].[dbo].tbl_item i with (nolock) "
        sqlStr = sqlStr & " 	on uu.itemid = i.itemid and uu.sitename = '"&FRectSitename&"'"

		if FRectCountryCd<>"" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_multiLang ml with (nolock)"
			sqlStr = sqlStr & " 	on uu.itemid = ml.itemid"
			sqlStr = sqlStr & " 	and ml.countryCd='"& FRectCountryCd &"'"
		end if

        sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c with (nolock) "
        sqlStr = sqlStr & " 	on i.makerid=c.userid"
        sqlStr = sqlStr & " where i.itemid<>0 " & addSql	'/and i.deliverytype in(1,4) and (i.mwdiv='M' or i.mwdiv='W')

		'response.write sqlStr & "<br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

		if FTotalCount < 1 then exit function

        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " uu.multilangcnt, i.itemname as itemname10x10, i.*"
        sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliverPay,0) as defaultDeliverPay"
        sqlStr = sqlStr & " , IsNULL(defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " ,isnull(uu.isusing,'') as siteisusing"		', ml.useyn
        sqlStr = sqlStr & " ,substring(STUFF(("
        sqlStr = sqlStr & " 	SELECT Top 100 '|^|' + ee.currencyUnit + '|*|' + cast(mpp.orgprice as varchar(10)) + '|*|'"
        sqlStr = sqlStr & " 	+ cast(mpp.wonprice as varchar(10)) + '|*|' + cast(mpp.lastexchangeRate as varchar(10))"
        sqlStr = sqlStr & " 	FROM #tmp_exchangeRatecurrencyunitgroup ee"
        sqlStr = sqlStr & " 	join db_item.dbo.tbl_item_multiLang_price mpp with (nolock)"
        sqlStr = sqlStr & " 		on ee.sitename=mpp.sitename"
        sqlStr = sqlStr & " 		and ee.currencyunit = mpp.currencyUnit"
        sqlStr = sqlStr & " 	WHERE ee.sitename = uu.sitename and mpp.itemid = uu.itemid"
        sqlStr = sqlStr & " 	ORDER BY ee.currencyUnit asc"
        sqlStr = sqlStr & " FOR XML PATH('')), 1, 1, ''),3,4000) as pricearr"
        sqlStr = sqlStr & " ,substring(STUFF(("
        sqlStr = sqlStr & " 	SELECT Top 100 '|^|' + ee.countryLangCD + '|*|' + mll.itemname"
        sqlStr = sqlStr & " 	FROM #tmp_exchangeRatecountryLangCDgroup ee"
        sqlStr = sqlStr & " 	join [db_item].[dbo].[tbl_item_multiLang] mll with (nolock)"
        sqlStr = sqlStr & " 		on ee.countryLangCD = mll.countryCd"
        sqlStr = sqlStr & " 		and mll.useyn='Y'"
        sqlStr = sqlStr & " 	WHERE ee.sitename = uu.sitename and uu.itemid=mll.itemid"
        sqlStr = sqlStr & " 	ORDER BY ee.countryLangCD asc"
        sqlStr = sqlStr & " FOR XML PATH('')), 1, 1, ''),3,4000) as countryLangCDarr"
        sqlStr = sqlStr & " from [db_item].[dbo].[tbl_item_multiSite_regItem] uu with (nolock)"
        sqlStr = sqlStr & " join [db_item].[dbo].tbl_item i with (nolock) "
        sqlStr = sqlStr & " 	on uu.itemid = i.itemid and uu.sitename = '"&FRectSitename&"'"

		if FRectCountryCd<>"" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_multiLang ml with (nolock)"
			sqlStr = sqlStr & " 	on uu.itemid = ml.itemid"
			sqlStr = sqlStr & " 	and ml.countryCd='"& FRectCountryCd &"'"
		end if

        sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c with (nolock) "
        sqlStr = sqlStr & " 	on i.makerid=c.userid"
        sqlStr = sqlStr & " where i.itemid<>0 " & addSql	'/and i.deliverytype in(1,4) and (i.mwdiv='M' or i.mwdiv='W')
       	sqlStr = sqlStr & " order by i.itemid desc"

		'response.write sqlStr & "<br>"
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new COverseasItemDetail

				FItemList(i).fmultilangcnt	= rsget("multilangcnt")
				FItemList(i).fcountryLangCDarr	= rsget("countryLangCDarr")
				FItemList(i).fpricearr			= rsget("pricearr")
				FItemList(i).fsiteisusing            = rsget("siteisusing")
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname10x10		= db2html(rsget("itemname10x10"))
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
                'FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	'/company/ch/target_itemlist.asp
	public function GetOverSeasTargetItemList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemDiv<> "") then
            addSql = addSql & " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

		if FRectLimityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            addSql = addSql + " and i.limityn='" + FRectLimityn + "'"
        end if

        if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if

        if FRectSailYn<>"" then
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatinclude='" + FRectVatYn + "'"
        end if

        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
        end if
        if FRectIsWeight<>"" then
			if FRectIsWeight="Y" then
				addSql = addSql + " and isnull(i.itemWeight,0)>0 "
			else
				addSql = addSql + " and isnull(i.itemWeight,0)<=0 "
			end if
		end if

        If FRectMinusMigin <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' "
        	addSql = addSql + " and ("
        	addSql = addSql + " 		(i.sellcash <= i.buycash) or "
        	addSql = addSql + " 		(i.itemcouponyn = 'Y' and i.curritemcouponidx is Not NULL and "
        	addSql = addSql + " 			(select "
        	addSql = addSql + " 				case itemcoupontype "
        	addSql = addSql + " 					when 1 then i.sellcash-i.sellcash*(itemcouponvalue/100) "
        	addSql = addSql + " 					else i.sellcash-itemcouponvalue "
        	addSql = addSql + " 				end "
        	addSql = addSql + " 			from db_item.dbo.tbl_item_coupon_master with (nolock) where itemcouponidx = i.curritemcouponidx"
        	addSql = addSql + " 			) < (Select top 1 D.couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail D with (nolock) Where D.itemcouponidx=i.curritemcouponidx and D.itemid=i.itemid) "
        	addSql = addSql + " 		)"
        	addSql = addSql + " 	)"
        End If

        If FRectMarginUP <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) >= " & FRectMarginUP & " "
        End If

        If FRectMarginDown <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) <= " & FRectMarginDown & " "
        End If

		if frectcolorcode <> "" then
			addSql = addSql + " and co.colorcode = "&frectcolorcode&""
		end if

		If FRectRegUserID <> "" Then
			addSql = addSql + " and uu.reguserid = '"&FRectRegUserID&"' and reguserid is not null "
		End IF

		If FRectIsReg = "o" Then
			addSql = addSql + " and uu.itemid is not null "
		ElseIf FRectIsReg = "x" Then
			addSql = addSql + " and uu.itemid is null "
		End IF

		If FRectRegDate1 <> "" Then
			addSql = addSql + " and uu.lastupdate >= '" & FRectRegDate1 & "' "
		End If

		If FRectRegDate2 <> "" Then
			addSql = addSql + " and uu.lastupdate <= '" & FRectRegDate2 & " 23:59:59' "
		End If

		IF FRectSellcash1 <> "" Then
			addSql = addSql + " and i.sellcash >= '" & FRectSellcash1 & "' "
		End IF

		IF FRectSellcash2 <> "" Then
			addSql = addSql + " and i.sellcash <= '" & FRectSellcash2 & "' "
		End IF

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i with (nolock)"
        sqlStr = sqlStr & "     left join [db_temp].[dbo].[tbl_해외판매상품알바관리로그] uu with (nolock) on i.itemid=uu.itemid AND uu.countryCd = 'kr'"

		if frectcolorcode <> "" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_colorOption co"
			sqlStr = sqlStr & " 	on i.itemid = co.itemid"
    	end if

        sqlStr = sqlStr & " where i.itemid<>0 " & addSql	'/and i.deliverytype in(1,4) and (i.mwdiv='M' or i.mwdiv='W')

		'rw sqlStr
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.*"
        sqlStr = sqlStr & " , IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(defaultDeliverPay,0) as defaultDeliverPay, IsNULL(defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " , IsNULL(A.itemid,0) as infoimageExists"
        sqlStr = sqlStr & " , (select count(itemid) from [db_item].[dbo].[tbl_item_multiLang] with (nolock) where itemid = i.itemid AND countryCd = 'kr') as existmultilang"
        sqlStr = sqlStr & " , IsNull(uu.reguserid,'') as reguserid"
        'sqlStr = sqlStr & " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i with (nolock) "
        sqlStr = sqlStr & "     left join [db_item].[dbo].tbl_item_addimage A with (nolock) on i.itemid=A.itemid and A.ImgType=1 and A.Gubun=1"
        sqlStr = sqlStr & "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
        sqlStr = sqlStr & "     left join [db_temp].[dbo].[tbl_해외판매상품알바관리로그] uu with (nolock) on i.itemid=uu.itemid AND uu.countryCd = 'kr'"
        'sqlStr = sqlStr & "    left join [db_item].[dbo].tbl_item_Contents s with (nolock)"
        'sqlStr = sqlStr & "    on i.itemid=s.itemid"

		if frectcolorcode <> "" then
			sqlStr = sqlStr & " join db_item.dbo.tbl_item_colorOption co with (nolock)"
			sqlStr = sqlStr & " 	on i.itemid = co.itemid"
    	end if

        sqlStr = sqlStr & " where i.itemid<>0 " & addSql	'/and i.deliverytype in(1,4) and (i.mwdiv='M' or i.mwdiv='W')

		IF FRectSortDiv="new" Then
			sqlStr = sqlStr & " Order by i.itemid desc, "
		ElseIf FRectSortDiv = "best" Then
			sqlStr = sqlStr & " ORDER BY i.itemScore DESC, "
		ElseIf FRectSortDiv = "min" Then
			sqlStr = sqlStr & " ORDER BY i.sellcash ASC, "
		ElseIf FRectSortDiv = "hi" Then
			sqlStr = sqlStr & " ORDER BY i.sellcash DESC, "
		ElseIf FRectSortDiv = "hs" Then
			sqlStr = sqlStr & " ORDER BY i.orgprice-i.sellcash DESC, "
		ELSEIF FRectSortDiv="cashH" Then
			sqlStr = sqlStr & " Order by i.SellCash desc, "
		ELSEIF FRectSortDiv="cashL" Then
			sqlStr = sqlStr & " Order by i.SellCash asc, "
		ELSEIF FRectSortDiv="best" Then
			sqlStr = sqlStr & " Order by i.ItemScore desc, "
		ELSE
			sqlStr = sqlStr & " Order by "
		End IF

		If FRectSortDiv2 = "weightup" Then
			sqlStr = sqlStr & " i.itemWeight desc "
		ElseIf FRectSortDiv2 = "weightdown" Then
			sqlStr = sqlStr & " i.itemWeight asc "
		End If

		'response.write  sqlStr
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new COverseasItemDetail

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
                'FItemList(i).Fcouponbuyprice    = rsget("couponbuyprice")	'쿠폰적용 매입가

                if (rsget("infoimageExists")>0) then
                    FItemList(i).FinfoimageExists   = true
                else
                    FItemList(i).FinfoimageExists   = false
                end if

                ''//기본 배송비 정책 관련 추가
                FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")
                If rsget("existmultilang") > 0 Then
                	FItemList(i).FExistMultiLang = "Y"
                Else
            		FItemList(i).FExistMultiLang = "N"
            	End IF
            	FItemList(i).FRegUserID = rsget("reguserid")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

    '//admin/itemmaster/overseas/popItemContent.asp
	public Sub GetOverSeasTargetItem()
		dim sqlstr,i
        dim iCountryLangCD : iCountryLangCD = getSiteCountryLangCD(FRectSitename)

		sqlstr = "SELECT mi.*, R.isusing as Siteisusing"
		sqlstr = sqlstr & " FROM [db_item].[dbo].[tbl_item_multiLang] AS mi with (nolock)"
		sqlstr = sqlstr & " left Join db_item.dbo.tbl_item_multiSite_regItem R with (nolock)"
		sqlstr = sqlstr & "     on mi.itemid=R.itemid"
		sqlstr = sqlstr & "     and R.sitename='"&FRectSitename&"'"
		sqlstr = sqlstr & " WHERE mi.itemid = '" & FRectItemID & "' AND mi.countryCd = '" & FRectMultiLanguage & "'"

		'response.write sqlstr & "<br>"
		rsget.Open sqlstr,dbget,1

		ftotalcount = rsget.recordcount

		set FOneItem = new COverseasItemDetail
		if Not rsget.Eof then
			FOneItem.Fitemname 		= db2html(rsget("itemname"))
			FOneItem.Fitemcopy 		= db2html(rsget("itemcopy"))
			FOneItem.Fitemcontent 	= db2html(rsget("itemcontent"))
			FOneItem.Fitemsource	= db2html(rsget("itemsource"))
			FOneItem.Fitemsize		= db2html(rsget("itemsize"))
			FOneItem.Fsourcearea	= db2html(rsget("sourcearea"))
			FOneItem.Fmakername		= db2html(rsget("makername"))
			FOneItem.fuseyn		    = rsget("useyn")
			FOneItem.fcountrycd		= rsget("countrycd")
			FOneItem.fkeywords 		= db2html(rsget("keywords"))
			FOneItem.fSiteisusing   = rsget("Siteisusing")
			FOneItem.FareaCode11st   = rsget("areaCode11st")
		else
			FOneItem.Fitemname 		= ""
			FOneItem.Fitemcopy 		= ""
			FOneItem.Fitemcontent 	= ""
			FOneItem.Fitemsource	= ""
			FOneItem.Fitemsize		= ""
			FOneItem.Fsourcearea	= ""
			FOneItem.Fmakername		= ""
			FOneItem.fuseyn	= ""
			'FOneItem.fcountrycd	= iCountryLangCD
			FOneItem.fcountrycd	= ""
			FOneItem.fkeywords = ""
			FOneItem.fSiteisusing   = ""
			FOneItem.FareaCode11st   = ""
		end if
		rsget.Close

	end Sub

	public function GetOverSeasTargetItemListXLS()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemDiv<> "") then
            addSql = addSql & " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

		if FRectLimityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            addSql = addSql + " and i.limityn='" + FRectLimityn + "'"
        end if

        if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if

        if FRectSailYn<>"" then
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatinclude='" + FRectVatYn + "'"
        end if

        if FRectDeliveryType<>"" then
        	  addSql = addSql + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
        end if
        if FRectIsWeight<>"" then
			if FRectIsWeight="Y" then
				addSql = addSql + " and isnull(i.itemWeight,0)>0 "
			else
				addSql = addSql + " and isnull(i.itemWeight,0)<=0 "
			end if
		end if

        If FRectMarginUP <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) >= " & FRectMarginUP & " "
        End If

        If FRectMarginDown <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.orgprice <> 0 and ((1-(i.orgsuplycash/i.orgprice))*100) <= " & FRectMarginDown & " "
        End If

		if frectcolorcode <> "" then
			addSql = addSql + " and co.colorcode = "&frectcolorcode&""
		end if

        '// 상품리스트_옵션_무.xlsx
'        sqlStr = "select "
'        sqlStr = sqlStr & " m.itemid, i.sellyn, m.itemname, i.brandname, m.makername, m.sourcearea, c.keywords, "
'        sqlStr = sqlStr & " i.sellcash, i.buycash, i.sellcash, 'http://www.10x10.co.kr/shopping/category_prd.asp?itemid=' + convert(varchar,m.itemid) as linkurl, "
'        sqlStr = sqlStr & " i.itemWeight, m.itemContent, m.itemsource, m.itemsize, '', '', '', '0' "
'        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i with (nolock) "
'        sqlStr = sqlStr & "    inner join db_item.dbo.tbl_item_multiLang as m with (nolock) on i.itemid = m.itemid "
'        sqlStr = sqlStr & "    left join db_item.dbo.tbl_item_Contents as c with (nolock) on i.itemid = c.itemid "
'        sqlStr = sqlStr & " where m.countryCd = 'kr' and i.deliverytype in(1,4) and i.sellyn = 'Y' and i.itemid not in(179229, 179233) and i.makerid <> 'urbanshop' "
'        sqlStr = sqlStr & " and i.itemid not in(select itemid from db_item.dbo.tbl_item_option_Multiple) "
'        sqlStr = sqlStr & " and i.itemid not in(select itemid from db_item.dbo.tbl_item_option) "
'        sqlStr = sqlStr & " and i.itemid<>0" & addSql

        '// 상품리스트_옵션_유.xlsx
        sqlStr = "select "
        sqlStr = sqlStr & " m.itemid, i.sellyn, m.itemname, i.brandname, m.makername, m.sourcearea, c.keywords, "
        sqlStr = sqlStr & " i.sellcash, i.buycash, i.sellcash, 'http://www.10x10.co.kr/shopping/category_prd.asp?itemid=' + convert(varchar,m.itemid) as linkurl, "
        sqlStr = sqlStr & " i.itemWeight, m.itemContent, m.itemsource, m.itemsize, o.itemoption, case when o.optionTypeName = '' then '선택' else o.optionTypeName + ' 선택' end, o.optionname, '1' "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i with (nolock) "
        sqlStr = sqlStr & "    inner join db_item.dbo.tbl_item_multiLang as m with (nolock) on i.itemid = m.itemid "
        sqlStr = sqlStr & "    inner join db_item.dbo.tbl_item_option as o with (nolock) on i.itemid = o.itemid "
        sqlStr = sqlStr & "    left join db_item.dbo.tbl_item_Contents as c with (nolock) on i.itemid = c.itemid "
        sqlStr = sqlStr & " where m.countryCd = 'kr' and i.sellyn = 'Y' and i.itemid not in(179229, 179233) and i.makerid <> 'urbanshop' "
        sqlStr = sqlStr & " and i.itemid not in (select itemid from db_item.dbo.tbl_item_option_Multiple with (nolock)) "
        sqlStr = sqlStr & " and i.itemid<>0" & addSql	'/and i.deliverytype in(1,4) and (i.mwdiv='M' or i.mwdiv='W')

		IF FRectSortDiv="new" Then
			sqlStr = sqlStr & " Order by i.itemid desc "
		ElseIf FRectSortDiv = "best" Then
			sqlStr = sqlStr & " ORDER BY i.itemScore DESC"
		ElseIf FRectSortDiv = "min" Then
			sqlStr = sqlStr & " ORDER BY i.sellcash ASC"
		ElseIf FRectSortDiv = "hi" Then
			sqlStr = sqlStr & " ORDER BY i.sellcash DESC"
		ElseIf FRectSortDiv = "hs" Then
			sqlStr = sqlStr & " ORDER BY i.orgprice-i.sellcash DESC"
		ELSEIF FRectSortDiv="cashH" Then
			sqlStr = sqlStr & " Order by i.SellCash desc "
		ELSEIF FRectSortDiv="cashL" Then
			sqlStr = sqlStr & " Order by i.SellCash"
		ELSEIF FRectSortDiv="best" Then
			sqlStr = sqlStr & " Order by i.ItemScore desc "
		ELSE
			sqlStr = sqlStr & " Order by i.itemid desc "
		End IF
       ' sqlStr = sqlStr & " order by i.itemid desc"

		'response.write  sqlStr
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
        FtotalPage =  Clng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = FTotalCount

        if (FResultCount<1) then FResultCount=0

        i=0
        if  not rsget.EOF  then
            farrlist = rsget.getrows()
        end if
        rsget.Close
    end function

	Public Sub getItem11STMYOptionInfo
		Dim sqlStr, addSql, i
		sqlstr = ""
		sqlstr = sqlstr & " SELECT "
		sqlstr = sqlstr & " o.itemid, o.itemoption, mo.optiontypename "
		sqlstr = sqlstr & " , mo.optionname, mo.isusing ,o.itemoption as itemoption10x10, o.optiontypename as optiontypename10x10 "
		sqlstr = sqlstr & " , o.optionname as optionname10x10, o.isusing as isusing10x10, mo.itemoption as regedoption "
		sqlstr = sqlstr & " FROM [db_item].[dbo].tbl_item_option as o with (nolock) "
		sqlstr = sqlstr & " LEFT JOIN [db_item].[dbo].[tbl_item_multiLang_option] as mo with (nolock) on o.itemid = mo.itemid and o.itemoption = mo.itemoption and mo.countryCd = 'EN' "
		sqlstr = sqlstr & " WHERE o.itemid='" & CStr(FRectItemID) & "'"
		sqlstr = sqlstr & " ORDER BY o.itemoption ASC"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				SET FItemList(i) = new COverseasItemDetail
					FItemList(i).FItemid				= rsget("itemid")
					FItemList(i).FItemoption			= rsget("itemoption")
					FItemList(i).F10x10itemoption		= rsget("itemoption10x10")
	'				FItemList(i).Fregedoption			= rsget("regedoption")
					If isNull(rsget("regedoption")) Then
						FItemList(i).FNotReg = "o"
						FItemList(i).FItemoption 		= rsget("itemoption10x10")
					End If
					FItemList(i).FOptisusing			= rsget("isusing")
					FItemList(i).F10x10optisusing		= rsget("isusing10x10")
					If FItemList(i).FNotReg = "o" Then
						FItemList(i).FOptisusing 		= rsget("isusing10x10")
					End If
					FItemList(i).FOptionname			= db2html(rsget("optionname"))
					FItemList(i).F10x10optionname		= db2html(rsget("optionname10x10"))
					If FItemList(i).FNotReg = "o" Then
						FItemList(i).FOptionname 		= db2html(rsget("optionname10x10"))
					End If
					FItemList(i).FOptiontypename 		= db2html(rsget("optiontypename"))
					FItemList(i).F10x10optiontypename	= db2html(rsget("optiontypename10x10"))
					If FItemList(i).FNotReg = "o" Then
						FItemList(i).FOptiontypename 	= db2html(rsget("optiontypename10x10"))
					End If
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

    '/admin/itemmaster/overseas/popItemContent.asp
	public Sub GetItemOptionInfo
		dim sqlstr,i

		sqlstr = " select top 1000"
		sqlstr = sqlstr + " o.optiontypename, o.optionname, o.isusing, o.optsellyn, o.optlimityn, o.optlimitno, o.optlimitsold"
		sqlstr = sqlstr + " ,IsNull(sm.realstock,0) as realstock, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv2,0) as ipkumdiv2, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv4,0) as ipkumdiv4, "
		sqlstr = sqlstr + " IsNull(sm.ipkumdiv5,0) as ipkumdiv5, "
		sqlstr = sqlstr + " IsNull(sm.offconfirmno,0) as offconfirmno, "
		sqlstr = sqlstr + " sm.lastupdate, cur.stockreipgodate, IsNull(o.optaddprice,0) as optaddprice"
		sqlstr = sqlstr + " , IsNull(cur.barcode,'') AS barcode, IsNull(cur.upchemanagecode,'') AS upchemanagecode "
		sqlstr = sqlstr + " ,os.upchemanagecode, os.itemid, os.itemgubun, os.itemoption"
		sqlstr = sqlstr + " , isnull(ii.isusing,'N') as off_isusing"
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item_option_stock as os with (nolock)"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option as o with (nolock)"
		sqlstr = sqlstr + " 	on os.itemid=o.itemid"
		sqlstr = sqlstr + " 	and os.itemoption = o.itemoption"
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm with (nolock)"
		sqlstr = sqlstr + " 	on sm.itemgubun='10' and os.itemid=sm.itemid and os.itemoption=sm.itemoption"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option_Stock cur with (nolock) "
		sqlstr = sqlstr + " 	on cur.itemgubun='10' and os.itemid=cur.itemid and os.itemoption=cur.itemoption"
		sqlstr = sqlstr + " left join db_shop.dbo.tbl_shop_item ii with (nolock)"
		sqlstr = sqlstr + " 	on os.itemid=ii.shopitemid"
		sqlstr = sqlstr + " 	and os.itemoption=ii.itemoption"
		sqlstr = sqlstr + " 	and ii.itemgubun='10'"
		sqlstr = sqlstr + " where os.itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " order by os.itemoption asc"

		'response.write sqlstr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COverseasItemDetail

				FItemList(i).foff_isusing		= rsget("off_isusing")
				FItemList(i).fitemgubun		= rsget("itemgubun")
				FItemList(i).fupchemanagecode		= rsget("upchemanagecode")
				FItemList(i).Fitemid		= rsget("itemid")
				FItemList(i).Fitemoption	= rsget("itemoption")
				FItemList(i).Foptisusing	= rsget("isusing")
				FItemList(i).Foptsellyn		= rsget("optsellyn")
				FItemList(i).Foptlimityn	= rsget("optlimityn")
				FItemList(i).Foptlimitno	= rsget("optlimitno")
				FItemList(i).Foptlimitsold	= rsget("optlimitsold")
				FItemList(i).Foptionname	= db2html(rsget("optionname"))
				FItemList(i).Foptiontypename = db2html(rsget("optiontypename"))
				FItemList(i).Frealstock		 = rsget("realstock")
				FItemList(i).Fipkumdiv2		 = rsget("ipkumdiv2")
				FItemList(i).Fipkumdiv4		 = rsget("ipkumdiv4")
				FItemList(i).Fipkumdiv5		 = rsget("ipkumdiv5")
				FItemList(i).Foffconfirmno	 = rsget("offconfirmno")
				FItemList(i).FLastUpdate	 = rsget("lastupdate")
				FItemList(i).Frestockdate	 = rsget("stockreipgodate")

				if (Len(FItemList(i).Frestockdate) = 10) then
					FItemList(i).Frestockdate = "<b>" & FItemList(i).Frestockdate & "</b>"
				else
					FItemList(i).Frestockdate = "0000-00-00"
				end if

				FItemList(i).Foptaddprice	= rsget("optaddprice")
				FItemList(i).Fbarcode		= rsget("barcode")
				FItemList(i).Fupchemanagecode = rsget("upchemanagecode")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

    '/admin/itemmaster/overseas/popItemContent.asp
	public Sub GetOverSeasItemOptionList
		dim sqlstr,i

		if frectCountryCd="" or FRectItemID="" then exit Sub

		sqlstr = "select top 1000"
		sqlstr = sqlstr + " isnull(os.itemgubun,'10') as itemgubun, o.itemid, o.itemoption"
		sqlstr = sqlstr + " , mo.itemoption as itemoption_wholesale, isnull(mo.optiontypename,o.optiontypename) as optiontypename"
		sqlstr = sqlstr + " , isnull(mo.optionname,o.optionname) as optionname, isnull(mo.isusing,o.isusing) as isusing"
		sqlstr = sqlstr + " , o.itemoption as itemoption10x10, o.optiontypename as optiontypename10x10, o.optionname as optionname10x10"
		sqlstr = sqlstr + " , o.isusing as isusing10x10"
		sqlstr = sqlstr + " , os.upchemanagecode, isnull(ii.isusing,'N') as off_isusing"
		sqlstr = sqlstr + " from [db_item].[dbo].tbl_item_option as o with (nolock)"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_multiLang_option as mo with (nolock)"
		sqlstr = sqlstr + " 	on o.itemid = mo.itemid"
		sqlstr = sqlstr + " 	and o.itemoption = mo.itemoption"
		sqlstr = sqlstr + " 	and mo.CountryCd='"&frectCountryCd&"'"
		sqlstr = sqlstr + " left join [db_item].[dbo].tbl_item_option_stock as os with (nolock)"
		sqlstr = sqlstr + " 	on os.itemgubun=10"
		sqlstr = sqlstr + " 	and os.itemid=o.itemid"
		sqlstr = sqlstr + " 	and os.itemoption=o.itemoption"
		sqlstr = sqlstr + " left join db_shop.dbo.tbl_shop_item ii with (nolock)"
		sqlstr = sqlstr + " 	on o.itemid=ii.shopitemid"
		sqlstr = sqlstr + " 	and o.itemoption=ii.itemoption"
		sqlstr = sqlstr + " 	and ii.itemgubun='10'"
		sqlstr = sqlstr + " where o.itemid=" + CStr(FRectItemID)
		sqlstr = sqlstr + " order by o.itemoption asc"

		'response.write sqlstr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COverseasItemDetail

				FItemList(i).fitemgubun				= rsget("itemgubun")
				FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fitemoption			= rsget("itemoption")
				FItemList(i).fitemoption_wholesale			= rsget("itemoption_wholesale")
				FItemList(i).Foptionname			= db2html(rsget("optionname"))
				FItemList(i).Foptiontypename 		= db2html(rsget("optiontypename"))
				FItemList(i).Foptisusing			= rsget("isusing")

				If isNull(rsget("itemoption_wholesale")) Then
					FItemList(i).FNotReg = "o"
				End If

				FItemList(i).F10x10itemoption		= rsget("itemoption10x10")
				FItemList(i).F10x10optiontypename	= db2html(rsget("optiontypename10x10"))
				FItemList(i).F10x10optionname		= db2html(rsget("optionname10x10"))
				FItemList(i).F10x10optisusing		= rsget("isusing10x10")
				FItemList(i).foff_isusing		= rsget("off_isusing")
				FItemList(i).fupchemanagecode		= rsget("upchemanagecode")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.close
	end Sub

    '/admin/itemmaster/overseas/popItemContent.asp
	public Sub GetOverSeasItemprice
		dim sqlstr,i

		sqlStr = "create table #tmp_exchangeRategroup ("
        sqlStr = sqlStr & " 	sitename nvarchar(32)"
        sqlStr = sqlStr & " 	,currencyunit nvarchar(16)"
        sqlStr = sqlStr & " 	,currencyChar nvarchar(50)"
        sqlStr = sqlStr & " 	,exchangeRate money"
        sqlStr = sqlStr & " 	,multiplerate float"
        sqlStr = sqlStr & " 	,linkPriceType int"
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " insert into #tmp_exchangeRategroup"
        sqlStr = sqlStr & " 	select sitename, currencyunit, currencyChar, exchangeRate, multiplerate, linkPriceType"
        sqlStr = sqlStr & " 	from db_item.dbo.tbl_exchangeRate with (nolock)"
        sqlStr = sqlStr & " 	where sitename='"& FRectSitename &"'"
        sqlStr = sqlStr & " 	group by sitename, currencyunit, currencyChar, exchangeRate, multiplerate, linkPriceType"

		'response.write sqlStr & "<br>"
		dbget.execute sqlStr

		sqlstr = " select"
		sqlstr = sqlstr & " r.sitename, r.currencyUnit, r.currencyChar, isnull(r.exchangeRate,1) as exchangeRate"
		sqlstr = sqlstr & " , mp.itemid, isnull(mp.orgprice,0) as orgprice, isnull(r.multiplerate,1) as multiplerate, isnull(r.linkPriceType,1) as linkPriceType"
		sqlstr = sqlstr & " , isnull(mp.wonprice,0) as wonprice"
		sqlstr = sqlstr & " ,(case when isnull(mp.itemid,'') ='' then 'o' else '' end) as NotReg"
		sqlstr = sqlstr & " from #tmp_exchangeRategroup r"
		sqlstr = sqlstr & " left join db_item.dbo.tbl_item_multiLang_price mp with (nolock)"
		sqlstr = sqlstr & " 	on r.sitename=mp.sitename"
		sqlstr = sqlstr & " 	and r.currencyUnit=mp.currencyUnit"
		sqlstr = sqlstr & " 	and mp.itemid="& frectitemid &""
		sqlstr = sqlstr & " where r.sitename= '"&frectSiteName&"'"
		sqlstr = sqlstr & " order by r.currencyunit asc"

		'response.write sqlstr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COverseasItemDetail

				FItemList(i).fCountryCd				= rsget("sitename")
				FItemList(i).fcurrencyUnit				= rsget("currencyUnit")
				FItemList(i).fcurrencyChar				= rsget("currencyChar")
				FItemList(i).fexchangeRate				= rsget("exchangeRate")
				FItemList(i).fitemid				= rsget("itemid")
				FItemList(i).forgprice				= rsget("orgprice")
				FItemList(i).fmultiplerate				= rsget("multiplerate")
				FItemList(i).flinkPriceType         = rsget("linkPriceType")
				FItemList(i).fwonprice				= rsget("wonprice")
				FItemList(i).fNotReg				= rsget("NotReg")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

	public Sub GetOverSeasDefaultPriceInfo
		dim sqlstr,i

		sqlStr = "select top 1 e.exchangeRate, e.multipleRate, linkPriceType "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_shop].[dbo].tbl_shop_user s "
        sqlStr = sqlStr & " 	join db_item.dbo.tbl_exchangeRate e "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and s.currencyUnit = e.currencyUnit "
        sqlStr = sqlStr & " 		and s.loginsite = e.sitename "
        sqlStr = sqlStr & " 		and s.countrylangcd = e.countrylangcd "
        sqlStr = sqlStr & " where s.userid = '"& FRectShopid &"' "
		''response.write sqlstr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COverseasItemDetail

				FItemList(i).fexchangeRate			= rsget("exchangeRate")
				FItemList(i).fmultiplerate			= rsget("multiplerate")
				FItemList(i).flinkPriceType         = rsget("linkPriceType")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
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

class cexchangerate_item
	public fidx
	public fsitename
	public fcurrencyUnit
	public fexchangeRate
	public fcurrencyChar
	public fbasedate
	public fregdate
	public flastupdate
	public freguserid
	public flastuserid
	public fcountryLangCD
	public fmultiplerate
	public flinkPriceType
	public FMakerid

    public function getlinkPriceTypeName()
        if isNULL(flinkPriceType) then Exit function

        if (flinkPriceType=1) then
            getlinkPriceTypeName = "SellPrice"
        elseif (flinkPriceType=2) then
            getlinkPriceTypeName = "OrgPrice"
        end if
    end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class cexchangerate
    public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount

	public frectidx
	public frectcurrencyUnit
	public frectsitename

	'//common/overseas/exchangerate/exchangerate.asp
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
		sqlStr = sqlStr & " , reguserid, lastuserid, countryLangCD, multiplerate, linkPriceType, makerid"
		sqlStr = sqlStr & " from db_item.dbo.tbl_exchangeRate with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr &"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount

        set FOneItem = new cexchangerate_item
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
            FOneItem.fcountryLangCD = rsget("countryLangCD")
			FOneItem.fmultiplerate = rsget("multiplerate")
			FOneItem.flinkPriceType = rsget("linkPriceType")
			FOneItem.FMakerid = rsget("makerid")
        end if
        rsget.close
    end Sub

	'//common/overseas/exchangerate/exchangerate.asp
	public sub fexchangerate_list()
		dim sqlStr,i , sqlsearch

		'총 갯수 구하기
		sqlStr = "select"
		sqlStr = sqlStr & " count(*) as cnt"
		sqlStr = sqlStr & " from db_item.dbo.tbl_exchangeRate"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " idx, sitename, currencyUnit, exchangeRate, basedate, currencyChar, regdate, lastupdate"
		sqlStr = sqlStr & " , reguserid, lastuserid, countryLangCD, multiplerate, linkPriceType, makerid"
		sqlStr = sqlStr & " from db_item.dbo.tbl_exchangeRate with (nolock)"
		sqlStr = sqlStr & " order by sitename asc, idx desc"

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
				set FItemList(i) = new cexchangerate_item

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
				FItemList(i).fcountryLangCD = rsget("countryLangCD")
				FItemList(i).fmultiplerate = rsget("multiplerate")
				FItemList(i).flinkPriceType = rsget("linkPriceType")
				FItemList(i).FMakerid = rsget("makerid")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

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