<%

class CLocationItem
	public Fcompanyid
	public Flocationid
	public Flocation_name
	public Fdivcd

	public Fuseforipgo
	public Fuseforchulgo
	public Fuseformaeip
	public Fuseformove

	public Fcompany_tel
	public Fcompany_fax
	public ftel
	public Freturn_zipcode				'업체 사무실주소로 전용
	public Freturn_address				'업체 사무실주소로 전용
	public Freturn_address2				'업체 사무실주소로 전용

	public Fmanager_name
	public Fmanager_phone
	public Fmanager_hp
	public Fmanager_email

	public Fdeliver_name
	public Fdeliver_phone
	public Fdeliver_hp
	public Fdeliver_email

	public Fsocno
	public Fsocname
	public Fceoname
	public Faddress
	public fmanager_address
	public Fbisstatus					'업태
	public Fbistype						'업종

	public Fdefaultsellcountday
	public Fdefaultstockneedday

	public Fdefaultdeliverytype			'배송유형
	public Fdefaultpurchasetype			'매입구분
	public Fdefaultpurchasemargin		'매입마진
	public Fdefaultsupplymargin			'공급마진

	public Fuseforeigndata
	public Fcurrencyunit
	public fcurrencyunit_pos
    public FcurrencyChar

	public Fuseyn
	public Fregdate
	public Flastupdate

	public function GetDivCDString()
		if Fdivcd="M" then
			GetDivCDString = "매입처"
		elseif Fdivcd="C" then
			GetDivCDString = "출고처"
		elseif Fdivcd="E" then
			GetDivCDString = "이동처"
		else
			GetDivCDString = Fdivcd
		end if
	end function

	public function GetDefaultDeliveryTypeString()
		if Fdefaultdeliverytype="SN" then
			GetDefaultDeliveryTypeString = "자체일반배송"
		elseif Fdefaultdeliverytype="SF" then
			GetDefaultDeliveryTypeString = "자체무료배송"
		elseif Fdefaultdeliverytype="UF" then
			GetDefaultDeliveryTypeString = "업체무료배송"
		elseif Fdefaultdeliverytype="UC" then
			GetDefaultDeliveryTypeString = "업체조건배송"
		else
			GetDefaultDeliveryTypeString = Fdefaultdeliverytype
		end if
	end function

	public function GetDefaultPurchaseTypeString()
		if Fdefaultpurchasetype="M" then
			GetDefaultPurchaseTypeString = "매입"
		elseif Fdefaultpurchasetype="W" then
			GetDefaultPurchaseTypeString = "위탁"
		elseif Fdefaultpurchasetype="U" then
			GetDefaultPurchaseTypeString = "업배"
		else
			GetDefaultPurchaseTypeString = Fdefaultpurchasetype
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CLocation
	public FItemList()
	public FOneItem

	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage

	public FRectCompanyId
	public FRectLocationId
	public FRectLocationName
	public FRectDivCD
    public FRectDeliverName
    public FRectUseYN

    public FRectSearchFrom
    public FRectSearchTo

	public Sub GetLocationList
		dim sqlStr,i
		sqlStr = "select count(locationid) as cnt "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner "
		sqlStr = sqlStr + " where 1=1"

		if FRectCompanyId<>"" then
			sqlStr = sqlStr + " and companyid = '" + FRectCompanyId + "'"
		end if

		if FRectLocationId<>"" then
			sqlStr = sqlStr + " and locationid = '" + FRectLocationId + "'"
		end if

		if FRectLocationName<>"" then
			sqlStr = sqlStr + " and location_name like '%" + FRectLocationName + "%'"
		end if

		if FRectDivCD<>"" then
			sqlStr = sqlStr + " and divcd = '" + FRectDivCD + "'"
		end if

		if FRectDeliverName<>"" then
			sqlStr = sqlStr + " and deliver_name like '%" + FRectDeliverName + "%'"
		end if

		if FRectUseYN<>"" then
			sqlStr = sqlStr + " and useyn = '" + FRectUseYN + "'"
		end if

		if FRectSearchFrom<>"" then
			if (FRectSearchFrom = FRectSearchTo) then
				sqlStr = sqlStr + " and ((location_name like '" + FRectSearchFrom + "%') or (locationid like '" + FRectSearchFrom + "%'))"
			else
				sqlStr = sqlStr + " and ((location_name >= '" + FRectSearchFrom + "' and location_name < '" + FRectSearchTo + "') or (locationid >= '" + FRectSearchFrom + "' and locationid< '" + FRectSearchTo + "')) "
				sqlStr = sqlStr + " "
			end if
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		''#################################################
		''현재 페이지 리스트.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " l.* "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner l "
		sqlStr = sqlStr + " where 1=1"

		if FRectCompanyId<>"" then
			sqlStr = sqlStr + " and companyid = '" + FRectCompanyId + "'"
		end if

		if FRectLocationId<>"" then
			sqlStr = sqlStr + " and locationid = '" + FRectLocationId + "'"
		end if

		if FRectLocationName<>"" then
			sqlStr = sqlStr + " and location_name like '%" + FRectLocationName + "%'"
		end if

		if FRectDivCD<>"" then
			sqlStr = sqlStr + " and divcd = '" + FRectDivCD + "'"
		end if

		if FRectDeliverName<>"" then
			sqlStr = sqlStr + " and deliver_name like '%" + FRectDeliverName + "%'"
		end if

		if FRectUseYN<>"" then
			sqlStr = sqlStr + " and useyn = '" + FRectUseYN + "'"
		end if

		if FRectSearchFrom<>"" then
			if (FRectSearchFrom = FRectSearchTo) then
				sqlStr = sqlStr + " and ((location_name like '" + FRectSearchFrom + "%') or (locationid like '" + FRectSearchFrom + "%'))"
			else
				sqlStr = sqlStr + " and ((location_name >= '" + FRectSearchFrom + "' and location_name < '" + FRectSearchTo + "') or (locationid >= '" + FRectSearchFrom + "' and locationid< '" + FRectSearchTo + "')) "
				sqlStr = sqlStr + " "
			end if
		end if

		sqlStr = sqlStr + " order by l.locationid "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr

		''올림.
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CLocationItem

				FItemList(i).Fcompanyid        = rsget("companyid")
				FItemList(i).Flocationid       = rsget("locationid")
				FItemList(i).Flocation_name    = db2html(rsget("location_name"))
				FItemList(i).Fdivcd            = rsget("divcd")

				FItemList(i).Fuseforipgo       = rsget("useforipgo")
				FItemList(i).Fuseforchulgo     = rsget("useforchulgo")
				FItemList(i).Fuseformaeip      = rsget("useformaeip")
				FItemList(i).Fuseformove       = rsget("useformove")

				FItemList(i).Fcompany_tel     = rsget("company_tel")
				FItemList(i).Fcompany_fax     = rsget("company_fax")

				FItemList(i).Freturn_zipcode  = rsget("return_zipcode")
				FItemList(i).Freturn_address  = db2html(rsget("return_address"))
				FItemList(i).Freturn_address2 = db2html(rsget("return_address2"))

				FItemList(i).Fmanager_name    = db2html(rsget("manager_name"))
				FItemList(i).Fmanager_phone   = rsget("manager_phone")
				FItemList(i).Fmanager_hp      = rsget("manager_hp")
				FItemList(i).Fmanager_email   = db2html(rsget("manager_email"))

				FItemList(i).Fdeliver_name    = db2html(rsget("deliver_name"))
				FItemList(i).Fdeliver_phone   = rsget("deliver_phone")
				FItemList(i).Fdeliver_hp      = rsget("deliver_hp")
				FItemList(i).Fdeliver_email   = db2html(rsget("deliver_email"))

				FItemList(i).Fsocno   			= db2html(rsget("socno"))
				FItemList(i).Fsocname   		= db2html(rsget("socname"))
				FItemList(i).Fceoname   		= db2html(rsget("ceoname"))
				FItemList(i).Faddress   		= db2html(rsget("address"))
				FItemList(i).Fbisstatus   		= db2html(rsget("bisstatus"))				'업태
				FItemList(i).Fbistype   		= db2html(rsget("bistype"))					'업종

				FItemList(i).Fdefaultsellcountday      = rsget("defaultsellcountday")
				FItemList(i).Fdefaultstockneedday      = rsget("defaultstockneedday")

				FItemList(i).Fdefaultdeliverytype      = rsget("defaultdeliverytype")
				FItemList(i).Fdefaultpurchasetype      = rsget("defaultpurchasetype")
				FItemList(i).Fdefaultpurchasemargin    = rsget("defaultpurchasemargin")
				FItemList(i).Fdefaultsupplymargin      = rsget("defaultsupplymargin")

				FItemList(i).Fuseyn           = rsget("useyn")
				FItemList(i).Fregdate         = rsget("indt")
				FItemList(i).Flastupdate      = rsget("updt")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub

	public Sub GetOneLocation
		dim sqlStr
		sqlStr = "select top 1 "
		sqlStr = sqlStr + " 	p.* " + vbcrlf
		sqlStr = sqlStr + " 	, (CASE " + vbcrlf
		sqlStr = sqlStr + " 			WHEN (u.shopdiv in ('7', '8')) THEN 'Y' " + vbcrlf
		sqlStr = sqlStr + " 			ELSE 'N' " + vbcrlf
		sqlStr = sqlStr + " 		END " + vbcrlf
		sqlStr = sqlStr + " 	) as useforeigndata " + vbcrlf
		sqlStr = sqlStr + " 	, u.currencyunit" + vbcrlf
		sqlStr = sqlStr + " 	, u.currencyunit_pos" + vbcrlf
		sqlStr = sqlStr + "     , (CASE " + vbcrlf
		sqlStr = sqlStr + " 			WHEN (ISNULL(R.currencyChar,'')= '') THEN '￦' " + vbcrlf
		sqlStr = sqlStr + " 			ELSE R.currencyChar " + vbcrlf
		sqlStr = sqlStr + " 		END " + vbcrlf
		sqlStr = sqlStr + " 	) as currencyChar " + vbcrlf
		sqlStr = sqlStr + " from " + vbcrlf
		sqlStr = sqlStr + " 	[db_partner].[dbo].tbl_partner p with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_user u with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and p.id = u.userid " + vbcrlf
		sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_exchangeRate R with (nolock)"
		''sqlStr = sqlStr + " 	on U.currencyunit=R.currencyunit"               ''2016/09/06 아래로 변경
		sqlStr = sqlStr + " 	on U.currencyunit_POS=R.currencyunit"
		sqlStr = sqlStr + " where id='" + html2db(FRectLocationId) + "'"

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		set FOneItem = new CLocationItem
		if Not rsget.Eof then

			FOneItem.Fcompanyid       	= "10x10"
			FOneItem.Flocationid       = rsget("id")
			FOneItem.Flocation_name    = db2html(rsget("company_name"))
			FOneItem.ftel      = rsget("tel")
			FOneItem.Fmanager_name    = db2html(rsget("manager_name"))
			FOneItem.Fmanager_phone   = rsget("manager_phone")
			FOneItem.Fmanager_hp      = rsget("manager_hp")

			FOneItem.Fdeliver_name    = rsget("deliver_name")
			FOneItem.Fdeliver_phone   = rsget("deliver_phone")

			FOneItem.Fsocno   			= db2html(rsget("company_no"))
			FOneItem.Fsocname   		= db2html(rsget("company_name"))
			FOneItem.Fceoname   		= db2html(rsget("ceoname"))
			FOneItem.Faddress   		= db2html(rsget("address"))
			FOneItem.fmanager_address   		= db2html(rsget("manager_address"))
			FOneItem.Fbisstatus   		= db2html(rsget("company_uptae"))					'업태
			FOneItem.Fbistype   		= db2html(rsget("company_upjong"))					'업종

			FOneItem.Fuseforeigndata	= db2html(rsget("useforeigndata"))
			FOneItem.Fcurrencyunit		= db2html(rsget("currencyunit"))
			FOneItem.fcurrencyunit_pos		= db2html(rsget("currencyunit_pos"))
            FOneItem.FcurrencyChar      = db2html(rsget("currencyChar"))


		end if
		rsget.close

	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
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
end class


'거래처/이동처 구분
Sub DrawLocationDivcdBox(byval divcdname, byval divcdval)
	dim buf,i

	buf = "<select class='select' name='" & divcdname & "'>"

	if (""=CStr(divcdval)) then
		buf = buf + "<option value='' selected>전체</option>"
	else
		buf = buf + "<option value='' >전체</option>"
    end if

	if ("M"=CStr(divcdval)) then
		buf = buf + "<option value='M' selected>매입처</option>"
	else
		buf = buf + "<option value='M' >매입처</option>"
    end if

	if ("C"=CStr(divcdval)) then
		buf = buf + "<option value='C' selected>출고처</option>"
	else
		buf = buf + "<option value='C' >출고처</option>"
    end if

	if ("E"=CStr(divcdval)) then
		buf = buf + "<option value='E' selected>이동처</option>"
	else
		buf = buf + "<option value='E' >이동처</option>"
    end if

    buf = buf + "</select>"

    response.write buf
end Sub

'거래처/이동처 사용구분
Sub DrawLocationUseYNBox(byval useynname, byval useynval)
	dim buf,i

	buf = "<select class='select' name='" & useynname & "'>"

	if ("Y"=CStr(useynval)) then
		buf = buf + "<option value='Y' selected>사용함</option>"
	else
		buf = buf + "<option value='Y' >사용함</option>"
    end if

	if ("N"=CStr(useynval)) then
		buf = buf + "<option value='N' selected>사용안함</option>"
	else
		buf = buf + "<option value='N' >사용안함</option>"
    end if

    buf = buf + "</select>"

    response.write buf
end Sub

%>
