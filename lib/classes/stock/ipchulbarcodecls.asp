<%
'###########################################################
' Description : 주문
' Hieditor : 2015.05.27 이상구 생성
'			 2016.09.08 한용민 수정
'###########################################################

public function GetShopCountrylangcd(shopid)
    dim sqlStr
    sqlStr = "select top 1 countrylangcd from  db_shop.dbo.tbl_shop_user u"
    sqlStr = sqlStr & " where userid='"&shopid&"'"

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
		GetShopCountrylangcd = rsget("countrylangcd")
    end if
	rsget.Close
end function

public function GetDivCDString(divcd)
	if divcd="IP" then
		GetDivCDString = "입고"
	elseif divcd="CH" then
		GetDivCDString = "출고"
	elseif divcd="MV" then
		GetDivCDString = "이동"
	else
		GetDivCDString = divcd
	end if
end function

public function GetDivCDColor(divcd)
	if divcd="IP" then
		GetDivCDColor = "blue"
	elseif divcd="CH" then
		GetDivCDColor = "red"
	elseif divcd="MV" then
		GetDivCDColor = "black"
	else
		GetDivCDColor = "black"
	end if
end function

class CStorageMasterItem
	public FStatecd
	public fforeign_statecd
	public Fidx
	public Fcompanyid
	public Fordercode
	public Fdivcd
	public Fstatusdivcd					'상태
	public Ftypecd						'매입구분
	public Flocationidfrom
	public Flocationidto
	public Fregisterid
	public Ffinisherid
	public Flocationnamefrom
	public Flocationnameto
	public Fregistername
	public Ffinishername
	public Finvoicetype					'송장
	public Finvoiceno
	public Fsumcustomerprice			'소비자가
	public Fsumsupplyprice				'공급가
	public Fsumpurchaseprice			'매입가
	public Fjungsaninsertyn				'정산입력 여부
	public Fregistermemo
	public Ffinishermemo
	public Fsmssenddt
	public Femailsenddt
	public Ffaxsendidx
	public Frequestdt
	public Ffinishdt					'입고일/출고일/이동입고일
	public Fisforeignorder
	public Fcurrencyunit
	public Fforeignordershopid
	public Fuseyn
	public Fregdate
	public Flastupdate
	public fbaljuid
	public ftotalsellcash
	public ftotalsuplycash
	public fjumunsellcash
	public fjumunsuplycash
	public fscheduledate
	public fbeasongdate
	public fipgodate
	public fbaljucode
	public fjumunforeign_sellcash
	public fjumunforeign_suplycash
	public ftotalforeign_sellcash
	public ftotalforeign_suplycash
	public fcurrencychar
	public ftotaldeliverpriceforeign
	public ffreightterm
	public fopenstate
	public fshippingaddress
	public finvoiceaddress
	public finvoiceidx

	public function IsFixed()
		if Fstatecd=" " then
			IsFixed = false
			exit function
		end if

		if (Fstatecd>="2") then
			IsFixed = true
		else
			IsFixed = false
		end if
	end function

    public function IsShippingPriceExistsState()
		'// 출고작업 이후에 배송비가 입력된다.
        IsShippingPriceExistsState = (Fstatecd >= "7")
    end function

    public function IsInvoiceExistsState()
        IsInvoiceExistsState = ((FshippingAddress <> "") or (FinvoiceAddress <> ""))
    end function

    public function getShippingPrice()
        getShippingPrice = 0
        IF Not IsShippingPriceExistsState then Exit function

		IF IsNull(FtotalDeliverPriceForeign) then Exit function

		getShippingPrice = FtotalDeliverPriceForeign
    end function

    public function getOrderNumFormat()
        getOrderNumFormat = Fbaljucode
    end function

    public function getItemArrNameFormat()
        getItemArrNameFormat = FfirstItemName&"... etc"
    end function

    public function getOrderDateFormat()
        getOrderDateFormat = Replace(Left(Cstr(Fregdate),10),"-",".")
    end function

    public function getOrderStateFormat()
        dim buf : buf= "order"
        'SHIPPING

        if (Fstatecd=" ") then
            if (Fforeign_statecd=7) then
                buf = "confirmed"
            elseif (Fforeign_statecd=3) then
                buf = "confirmed"
            elseif (Fforeign_statecd=0) then
                buf = "order"
            else
                buf = "order"
            end if
        elseif (Fstatecd="0") then
            buf = "packing" ''주문접수 = 패킹지시
        elseif (Fstatecd="1") then
            buf = "packing" ''주문확인 = 패킹중
        elseif (Fstatecd="5") then
            buf = "packing" ''배송준비 = 패킹중
        elseif (Fstatecd="6") then
            buf = "waiting" ''출고대기 = 패킹중
        elseif (Fstatecd="7") then
            buf = "shipping"  ''출고완료
        elseif (Fstatecd="8") then
            buf = "shipping"  ''검품완료
        elseif (Fstatecd="9") then
            buf = "delivered"  ''입고완료
        end if

        if (buf="order") then
            getOrderStateFormat = buf
        elseif (buf="confirmed") then
            getOrderStateFormat = "<span class='cr000'>"&buf&"</span>"
        elseif (buf="packing") then
            getOrderStateFormat = "<span class='crABl'>"&buf&"</span>"
        elseif (buf="waiting") then
            getOrderStateFormat = "<span class='crABl'>"&buf&"</span>"
        elseif (buf="shipping") then
            getOrderStateFormat = "<span class='crRed'>"&buf&"</span>"
        else
            getOrderStateFormat = "<span class='cr000'>"&buf&"</span>"
        end if

    end function

	public function GetDivCDString()
		if Fdivcd="IP" then
			GetDivCDString = "입고"
		elseif Fdivcd="CH" then
			GetDivCDString = "출고"
		elseif Fdivcd="MV" then
			GetDivCDString = "이동"
		else
			GetDivCDString = Fdivcd
		end if
	end function

	public function GetMWDivString()
		if Ftypecd="M" then
			GetMWDivString = "매입"
		elseif Ftypecd="CH" then
			GetMWDivString = "위탁"
		else
			GetMWDivString = Ftypecd
		end if
	end function

	public function GetStateCDString()
		if Fstatusdivcd="0" then
			GetStateCDString = "작성중"
		elseif Fstatusdivcd="2" then
			GetStateCDString = "주문접수"
		elseif Fstatusdivcd="5" then
			GetStateCDString = "업체확인"
		elseif Fstatusdivcd="7" then
			GetStateCDString = "출고완료"
		elseif Fstatusdivcd="9" then
			GetStateCDString = "도착확인"
		else
			GetStateCDString = Fstatusdivcd
		end if
	end function

	function GetStateCDColor()
		if Fstatusdivcd="0" then
			GetStateCDColor = "#AAAAAA"
		elseif Fstatusdivcd="2" then
			GetStateCDColor = "#3333CC"
		elseif Fstatusdivcd="5" then
			GetStateCDColor = "#33CC33"
		elseif Fstatusdivcd="7" then
			GetStateCDColor = "#CC3333"
		elseif Fstatusdivcd="9" then
			GetStateCDColor = "#000000"
		else
			GetStateCDColor = Fstatusdivcd
		end if
	end function

	public function GetCodeString()
		dim tmp

		if ((IsNull(Fdivcd)) or (Fdivcd = "")) then
			GetCodeString = ""
		else
			tmp = 1000000 + Fidx
			tmp = Right(CStr(tmp), 6)

			GetCodeString = Fdivcd & tmp
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class CStorageMaster
	public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage

	public FRectCompanyId
	public FRectMasterIdx
	public FRectDivCD						'입고/이동/출고
    public FRectUseYN
    public FRectLocationIdFrom
    public FRectLocationIdTo
    public FRectArrMasterIDX
    public FRectFaxIDX
    public FRectOrderCode					'주문코드
    public FRectTypeCD						'매입/위탁
    public FRectSearchDateType				'조건없음/작성일/입고요청일/입고일
    public FRectYyyymmddfrom
    public FRectYyyymmddto
    public FRectStatusDivCD					'상태
    public FRectRegisterMemo
    public FRectShopId
    public frectbaljucode
    public frectsitename
    public FRectOrderIDX
    public FRectAuthMode

	'/admin/fran/viewordersheet.asp
    public function getShopOneOrderMaster()
        dim sqlStr, i

		if (FRectAuthMode = "none") then
        	sqlStr = "exec db_shop.dbo.[sp_Ten_Shop_Order_Front_GetOneOrder_InSafe] '"&frectsitename&"','"&FRectOrderIDX&"'"
		else
			sqlStr = "exec db_shop.dbo.[sp_Ten_Shop_Order_Front_GetOneOrder] '"&FRectForeignOrderShopid&"','"&frectsitename&"','"&FRectOrderIDX&"'"
		end if
		
		'response.write sqlStr &"<Br>"
		'response.end
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FResultCount = rsget.RecordCount

        if  not rsget.EOF  then
            set FOneItem = new CStorageMasterItem

            FOneItem.Fidx                    = rsget("idx")
            FOneItem.Fbaljuid                = rsget("baljuid")
            FOneItem.Ftotalsellcash          = rsget("totalsellcash")
            FOneItem.Ftotalsuplycash         = rsget("totalsuplycash")
            FOneItem.Fjumunsellcash          = rsget("jumunsellcash")
            FOneItem.Fjumunsuplycash         = rsget("jumunsuplycash")
            FOneItem.Fregdate                = rsget("regdate")
            FOneItem.Fscheduledate           = rsget("scheduledate")
            FOneItem.Fbeasongdate            = rsget("beasongdate")
            FOneItem.Fipgodate               = rsget("ipgodate")
            FOneItem.Fbaljucode              = rsget("baljucode")
            FOneItem.Fstatecd                = rsget("statecd")
            FOneItem.FcurrencyUnit           = rsget("currencyUnit")
            FOneItem.Fforeign_statecd        = rsget("foreign_statecd")
            FOneItem.Fjumunforeign_sellcash  = rsget("jumunforeign_sellcash")
            FOneItem.Fjumunforeign_suplycash = rsget("jumunforeign_suplycash")
            FOneItem.Ftotalforeign_sellcash  = rsget("totalforeign_sellcash")
            FOneItem.Ftotalforeign_suplycash = rsget("totalforeign_suplycash")
            ''FOneItem.FfirstItemName          = rsget("firstItemName")
            FOneItem.FcurrencyChar           = rsget("currencyChar")
			FOneItem.FtotalDeliverPriceForeign  = rsget("totalDeliverPriceForeign")
			FOneItem.FfreightTerm           	= rsget("freightTerm")
			FOneItem.FopenState           		= rsget("openState")
			FOneItem.FshippingAddress       	= rsget("shippingAddress")
			FOneItem.FinvoiceAddress        	= rsget("invoiceAddress")
			FOneItem.finvoiceidx        	= rsget("invoiceidx")

        end if
        rsget.Close

    end function

	'해당주문과 관련된 패킹일(발주일) 목록
	public function GetPackingDayList()
		dim query1, sqlStr,i, tmpstr

		if FRectMasterIdx="" then exit function

		query1 = " select distinct " + vbcrlf
		query1 = query1 + " 	CONVERT(VARCHAR(10),b.baljudate,21) as yyyymmdd " + vbcrlf
		query1 = query1 + " from " + vbcrlf
		query1 = query1 + " 	[db_storage].[dbo].tbl_shopbalju b " + vbcrlf
		query1 = query1 + " 	JOIN [db_storage].[dbo].tbl_ordersheet_master m " + vbcrlf
		query1 = query1 + " 	ON " + vbcrlf
		query1 = query1 + " 		b.baljucode = m.baljucode " + vbcrlf
		query1 = query1 + " where " + vbcrlf
		query1 = query1 + " 	CONVERT(VARCHAR(10),b.baljudate,21) = ( " + vbcrlf
		query1 = query1 + " 		select " + vbcrlf
		query1 = query1 + " 			CONVERT(VARCHAR(10),b.baljudate,21) as yyyymmdd " + vbcrlf
		query1 = query1 + " 		from " + vbcrlf
		query1 = query1 + " 			[db_storage].[dbo].tbl_shopbalju b " + vbcrlf
		query1 = query1 + " 			JOIN [db_storage].[dbo].tbl_ordersheet_master m " + vbcrlf
		query1 = query1 + " 			ON " + vbcrlf
		query1 = query1 + " 				b.baljucode = m.baljucode " + vbcrlf
		query1 = query1 + " 		where " + vbcrlf
		query1 = query1 + " 			m.idx = " & CStr(FRectMasterIdx) & " " + vbcrlf
		query1 = query1 + " 	) " + vbcrlf
		query1 = query1 + " 	and m.baljuid = '" & CStr(FRectShopId) & "' " + vbcrlf
		query1 = query1 + " order by " + vbcrlf
		query1 = query1 + " 	CONVERT(VARCHAR(10),b.baljudate,21) desc " + vbcrlf

		'response.write query1 & "<br>"
		rsget.Open query1,dbget,1
		tmpstr = ""
		if Not rsget.Eof then
			do until rsget.eof
				tmpstr = tmpstr + ", " + rsget("yyyymmdd")
				rsget.movenext
			loop
		end if
		rsget.close

		if (LEN(tmpstr) > 2) then
			tmpstr = RIGHT(tmpstr, (LEN(tmpstr) - 2))
		end if
		GetPackingDayList = tmpstr
	end function

	'해당주문과 관련된 패킹일(발주일) 관련 주문목록
	public function GetOrderCodeList()
		dim query1, sqlStr,i, tmpstr

		if FRectMasterIdx="" then exit function

		query1 = " select distinct " + vbcrlf
		query1 = query1 + " 	m.baljucode " + vbcrlf
		query1 = query1 + " from " + vbcrlf
		query1 = query1 + " 	[db_storage].[dbo].tbl_shopbalju b " + vbcrlf
		query1 = query1 + " 	JOIN [db_storage].[dbo].tbl_ordersheet_master m " + vbcrlf
		query1 = query1 + " 	ON " + vbcrlf
		query1 = query1 + " 		b.baljucode = m.baljucode " + vbcrlf
		query1 = query1 + " where " + vbcrlf
		query1 = query1 + " 	CONVERT(VARCHAR(10),b.baljudate,21) = ( " + vbcrlf
		query1 = query1 + " 		select " + vbcrlf
		query1 = query1 + " 			CONVERT(VARCHAR(10),b.baljudate,21) as yyyymmdd " + vbcrlf
		query1 = query1 + " 		from " + vbcrlf
		query1 = query1 + " 			[db_storage].[dbo].tbl_shopbalju b " + vbcrlf
		query1 = query1 + " 			JOIN [db_storage].[dbo].tbl_ordersheet_master m " + vbcrlf
		query1 = query1 + " 			ON " + vbcrlf
		query1 = query1 + " 				b.baljucode = m.baljucode " + vbcrlf
		query1 = query1 + " 		where " + vbcrlf
		query1 = query1 + " 			m.idx = " & CStr(FRectMasterIdx) & " " + vbcrlf
		query1 = query1 + " 	) " + vbcrlf
		query1 = query1 + " 	and m.baljuid = '" & CStr(FRectShopId) & "' " + vbcrlf
		query1 = query1 + " order by " + vbcrlf
		query1 = query1 + " 	m.baljucode desc " + vbcrlf

		'response.write query1 & "<br>"
		rsget.Open query1,dbget,1
		tmpstr = ""
		if Not rsget.Eof then
			do until rsget.eof

				tmpstr = tmpstr + ", " + rsget("baljucode")
				rsget.movenext
			loop
		end if
		rsget.close

		if (LEN(tmpstr) > 2) then
			tmpstr = RIGHT(tmpstr, (LEN(tmpstr) - 2))
		end if
		GetOrderCodeList = tmpstr
	end function

	public Sub GetStorageMasterList
		dim sqlStr,i

		sqlStr = "select count(idx) as cnt "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master "
		sqlStr = sqlStr + " where 1=1"

		if FRectCompanyId<>"" then
			sqlStr = sqlStr + " and companyid = '" + FRectCompanyId + "'"
		end if

		if FRectDivCD<>"" then
			sqlStr = sqlStr + " and divcd = '" + FRectDivCD + "'"
		end if

		if FRectStatusDivCD<>"" then
			sqlStr = sqlStr + " and statusdivcd = '" + FRectStatusDivCD + "'"
		end if

		if ((FRectUseYN<>"") and (FRectUseYN<>"all")) then
			if (FRectUseYN = "Y") then
				sqlStr = sqlStr + " and useyn  = 'Y' "
			else
				sqlStr = sqlStr + " and useyn  <> 'Y' "
			end if
		end if

		if FRectLocationIdFrom<>"" then
			sqlStr = sqlStr + " and locationidfrom = '" + FRectLocationIdFrom + "'"
		end if

		if FRectLocationIdTo<>"" then
			sqlStr = sqlStr + " and locationidto = '" + FRectLocationIdTo + "'"
		end if

		if FRectOrderCode<>"" then
			sqlStr = sqlStr + " and ordercode = '" + FRectOrderCode + "'"
		end if

		if FRectTypeCD<>"" then
			sqlStr = sqlStr + " and typecd = '" + FRectTypeCD + "'"
		end if

		if FRectSearchDateType<>"" then
			if FRectSearchDateType = "W" then
				sqlStr = sqlStr + " and regdate >= '" + FRectYyyymmddfrom + " 00:00:00'"
				sqlStr = sqlStr + " and regdate <= '" + FRectYyyymmddto + " 23:59:59'"
			end if
			if FRectSearchDateType = "R" then
				sqlStr = sqlStr + " and requestdt >= '" + FRectYyyymmddfrom + " 00:00:00'"
				sqlStr = sqlStr + " and requestdt <= '" + FRectYyyymmddto + " 23:59:59'"
			end if
			if FRectSearchDateType = "F" then
				sqlStr = sqlStr + " and finishdt >= '" + FRectYyyymmddfrom + " 00:00:00'"
				sqlStr = sqlStr + " and finishdt <= '" + FRectYyyymmddto + " 23:59:59'"
			end if
		end if

		if FRectRegisterMemo<>"" then
			sqlStr = sqlStr + " and registermemo like '%" + FRectRegisterMemo + "%'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.* "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m "
		sqlStr = sqlStr + " where 1=1"

		if FRectCompanyId<>"" then
			sqlStr = sqlStr + " and companyid = '" + FRectCompanyId + "'"
		end if

		if FRectDivCD<>"" then
			sqlStr = sqlStr + " and divcd = '" + FRectDivCD + "'"
		end if

		if FRectStatusDivCD<>"" then
			sqlStr = sqlStr + " and statusdivcd = '" + FRectStatusDivCD + "'"
		end if

		if ((FRectUseYN<>"") and (FRectUseYN<>"all")) then
			if (FRectUseYN = "Y") then
				sqlStr = sqlStr + " and useyn  = 'Y' "
			else
				sqlStr = sqlStr + " and useyn  <> 'Y' "
			end if
		end if

		if FRectLocationIdFrom<>"" then
			sqlStr = sqlStr + " and locationidfrom = '" + FRectLocationIdFrom + "'"
		end if

		if FRectLocationIdTo<>"" then
			sqlStr = sqlStr + " and locationidto = '" + FRectLocationIdTo + "'"
		end if

		if FRectOrderCode<>"" then
			sqlStr = sqlStr + " and ordercode = '" + FRectOrderCode + "'"
		end if

		if FRectTypeCD<>"" then
			sqlStr = sqlStr + " and typecd = '" + FRectTypeCD + "'"
		end if

		if FRectSearchDateType<>"" then
			if FRectSearchDateType = "W" then
				sqlStr = sqlStr + " and regdate >= '" + FRectYyyymmddfrom + " 00:00:00'"
				sqlStr = sqlStr + " and regdate <= '" + FRectYyyymmddto + " 23:59:59'"
			end if
			if FRectSearchDateType = "R" then
				sqlStr = sqlStr + " and requestdt >= '" + FRectYyyymmddfrom + " 00:00:00'"
				sqlStr = sqlStr + " and requestdt <= '" + FRectYyyymmddto + " 23:59:59'"
			end if
			if FRectSearchDateType = "F" then
				sqlStr = sqlStr + " and finishdt >= '" + FRectYyyymmddfrom + " 00:00:00'"
				sqlStr = sqlStr + " and finishdt <= '" + FRectYyyymmddto + " 23:59:59'"
			end if
		end if

		if FRectRegisterMemo<>"" then
			sqlStr = sqlStr + " and registermemo like '%" + FRectRegisterMemo + "%'"
		end if

		sqlStr = sqlStr + " order by m.idx desc "
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
				set FItemList(i) = new CStorageMasterItem

				FItemList(i).Fidx        		= rsget("idx")
				FItemList(i).Fcompanyid       	= rsget("companyid")
				FItemList(i).Fordercode       	= rsget("ordercode")
				FItemList(i).Fdivcd    			= rsget("divcd")
				FItemList(i).Fstatusdivcd    	= rsget("statusdivcd")
				FItemList(i).Ftypecd    		= rsget("typecd")
				FItemList(i).Flocationidfrom    = rsget("locationidfrom")
				FItemList(i).Flocationidto     	= rsget("locationidto")
				FItemList(i).Fregisterid      	= rsget("registerid")
				FItemList(i).Ffinisherid       	= rsget("finisherid")
				FItemList(i).Flocationnamefrom  = db2html(rsget("locationnamefrom"))
				FItemList(i).Flocationnameto    = db2html(rsget("locationnameto"))
				FItemList(i).Fregistername     	= db2html(rsget("registername"))
				FItemList(i).Ffinishername     	= db2html(rsget("finishername"))
				FItemList(i).Finvoicetype		= rsget("invoicetype")
				FItemList(i).Finvoiceno			= rsget("invoiceno")
				FItemList(i).Fsumcustomerprice  = rsget("sumcustomerprice")
				FItemList(i).Fsumsupplyprice  	= rsget("sumsupplyprice")
				FItemList(i).Fsumpurchaseprice 	= rsget("sumpurchaseprice")
				FItemList(i).Fregistermemo    	= db2html(rsget("registermemo"))
				FItemList(i).Ffinishermemo    	= db2html(rsget("finishermemo"))
				FItemList(i).Fsmssenddt 		= rsget("smssenddt")
				FItemList(i).Femailsenddt 		= rsget("emailsenddt")
				FItemList(i).Ffaxsendidx		= rsget("faxsendidx")
				FItemList(i).Frequestdt 		= rsget("requestdt")
				FItemList(i).Ffinishdt 			= rsget("finishdt")
				FItemList(i).Fuseyn 			= rsget("useyn")
				FItemList(i).Fregdate         	= rsget("regdate")
				FItemList(i).Flastupdate      	= rsget("lastupdt")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub

	public Sub GetOneStorageMaster
		dim sqlStr, sqlsearch

		if frectbaljucode="" and FRectMasterIdx="" then exit Sub

		if frectbaljucode<>"" then
			sqlsearch = sqlsearch & " and m.baljucode='"& frectbaljucode &"'"
		end if
		if FRectMasterIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx="& FRectMasterIdx &""
		end if

		sqlStr = " select top 1 " + vbcrlf
		sqlStr = sqlStr + " 	m.* " + vbcrlf
		sqlStr = sqlStr + " 	, (CASE " + vbcrlf
		sqlStr = sqlStr + " 			WHEN (ubalju.shopdiv in ('7', '8')) THEN 'Y' " + vbcrlf
		sqlStr = sqlStr + " 			ELSE 'N' " + vbcrlf
		sqlStr = sqlStr + " 		END " + vbcrlf
		sqlStr = sqlStr + " 	) as isforeignorder " + vbcrlf
		sqlStr = sqlStr + " 	, ubalju.currencyunit" + vbcrlf
		sqlStr = sqlStr + " 	, IsNull(m.scheduledate, m.regdate) as requestdt " + vbcrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m with (nolock)" + vbcrlf
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user ubalju with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on m.baljuid = ubalju.userid " + vbcrlf
		sqlStr = sqlStr + " where 1=1 " & sqlsearch

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		set FOneItem = new CStorageMasterItem
		if Not rsget.Eof then
			FOneItem.Fidx        		= rsget("idx")
			FOneItem.Fcompanyid       	= "10x10"
			FOneItem.Fordercode       	= rsget("baljucode")
			FOneItem.Fsumcustomerprice  = rsget("jumunsellcash")
			FOneItem.Fsumsupplyprice  	= rsget("totalsuplycash")
			FOneItem.Fisforeignorder  	= rsget("isforeignorder")
			FOneItem.Fcurrencyunit  	= rsget("currencyunit")
			FOneItem.Fforeignordershopid  	= rsget("baljuid")
			FOneItem.Flocationidfrom  	= rsget("targetid")
			FOneItem.Flocationidto  	= rsget("baljuid")
			FOneItem.Flocationnamefrom  = rsget("targetname")
			FOneItem.Flocationnameto  	= rsget("baljuname")
			FOneItem.Ffinishdt  		= rsget("ipgodate")
			FOneItem.Frequestdt  		= rsget("requestdt")
			FOneItem.Fregistermemo  	= rsget("comment")
			FOneItem.FStatecd  	= rsget("Statecd")
			FOneItem.fforeign_statecd  	= rsget("foreign_statecd")
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

class CStorageDetailItem
	public fcurrencyunit
	public fitemweight
	public fdeliverOverseas
	public fsocname
	public fsocname_kor
	public fexchangeRate
	public fmultipleRate
	public fbaljuid
	public fdetailidx
	public falinkcode
	public fbaljudate
	public finnerboxno
	public fcartoonboxno
	public fmakerid
	public fitemname
	public Fbaljuitemno
	public frealitemno
	public fsellcash
	public fsuplycash
	public fbuycash
	public fdefaultsuplymargin
	public ftotalsuplycash
	public fitemname_10x10
	public foptionname_10x10
	public fitemsource_10x10
	public fsourcearea_10x10
	public fitemname_en
	public foptionname_en
	public fitemsource_en
	public fsourcearea_en
	public fpublicbarcode
	public Fbaljucode
	public Fidx
	public Fmasteridx
	public Fprdcode
	public Ftypecd				'매입구분
	public Fbarcode				'사용안함
	public Fgroupid				'그룹(미사용)
	public Flocationid			'매입처
	public Fbrandid				'미사용
	public Fgeneralbarcode
	public Fprdbarcode
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemoptionname
	public Fmainimageurl
	public Fprdname
	public Fgroupname
	public Flocationname
	public Fcustomerprice
	public Fsellprice
	public Fsupplyprice
	public Fpurchaseprice
	public Frequestedno
	public Ffixedno				'주문확정
	public Ferrorcd				'주문수량 오차 사유
	public Ferrorcomment		'주문수량 오차 사유 직접입력(오차가 없어도 입력가능)
	public Flcitemname
	public Flcitemoptionname
	public Flcprice
	public Fuseyn
	public Flastupdt
	public Fregdate
	public Fboxno
	public FItemCouponYN
	public Fitemcoupontype
	public Fitemcouponvalue
	public Fsaleyn
	public Fsaleprice
	public fitemcopy
	public Frealstockno
	public Flocation_name
	public fcatename1
	public fcatename2
	public fcatename_cn_gan2
	public fcatename_cn_bun2
	public fcatename3
	public fcatename_cn_gan3
	public fcatename_cn_bun3
	public fcatename_eng2
	public fcatename_eng3
	public fcatecdl
	public fcatecdm
	public fcatecdn
	public fitemsize_10x10
	public fitemsize_en
	public fitemrackcode
	public fprtidx
	public fsubitemrackcode
	public Fforeign_sellcash
	public Fforeign_suplycash
	public FmLitemname
	public FmLitemOptionname
	public FupchemanageCode
	public FImageList
	public FImageSmall
	public fextbarcode
	public Fstatecd
	public Fforeign_statecd
	public Fshopitemid
	public Fshopitemname
	public Fshopitemoptionname
	public Fshopitemprice

    public function getOptionDpFormat()
    	dim tmpOptionname
   
		if Fitemoption<>"" and Fitemoption<>"0000" then
			tmpOptionname = "Option : "&FmLitemOptionname
		end if

        getOptionDpFormat = tmpOptionname
    end function

	'// 세일포함 실제가격
	public Function getRealPrice() '!
		getRealPrice = Fcustomerprice

		'if (IsSpecialUserItem()) then
		'	getRealPrice = getSpecialShopItemPrice(FSellCash)
		'end if
	end Function

	'// 상품 쿠폰 여부
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function

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

	public function GetTypeCDString()
		dim tmp

		if (Ftypecd = "M") then
			GetTypeCDString = "매입"
		elseif (Ftypecd = "W") then
			GetTypeCDString = "위탁"
		else
			GetTypeCDString = Ftypecd
		end if
	end function

	public function GetErrorCDString()
		dim tmp

		if (Ferrorcd = "T") then
			GetErrorCDString = "일시품절"
		elseif (Ferrorcd = "E") then
			GetErrorCDString = "단종"
		elseif (Ferrorcd = "C") then
			GetErrorCDString = "직접입력"
		else
			GetErrorCDString = Ferrorcd
		end if
	end function

	public function GetBarCode()
		GetBarCode = Fitemgubun & Format00(6,Fshopitemid) & Fitemoption
		if (Fshopitemid >= 1000000) then
			getBarCode = CStr(Fitemgubun) + Format00(8,Fshopitemid) + Fitemoption
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class CStorageDetail
	public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage
	public FPageCount

	public frectbaljucode
	public FRectCompanyId
	public FRectMasterIdx
	public FRectIsForeignOrder
	public FRectForeignOrderShopid
	public FRectShopId
	public FRectBoxNo
	public FRectmakerid
	public FRectstartdate
	public FRectenddate
	public FRectChulgoYN
	public FRectStatecd
	public FRectItemid
	public FRectinnerboxno
	public FRectShopDiv
	public FRectShowDeleted
	public FRectMichulgoReason
	public FRectDateType
	public FRectcartoonboxmasteridx
	public FRectPrdCode
	public FRectGeneralBarcode
	public FRectItemName
	public FRectisforeignprint
	public FRectSellYN
	public FRectisUsing
	public frectitembarcodearr
	Public FRectShowSupplyCash
	public frectsitename
	public FRectOrderIDX
	public FRectAuthMode
	public FRectCDL
	public FRectCDM
	public FRectCDS
	public FRectShopItemName
	public FRectCurrentStockExist
	public FRectRealStockOneMore
	public FRectShopItemNameInserted
	public FRectLocationId
	public FRectitemgubun

	public function GetMaxBoxNo()
		dim sqlStr,i, sqlsearch

		if FRectMasterIdx="" then exit function

		if FRectMasterIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx="& FRectMasterIdx &""
		end if

		sqlStr = " select max(isnull(d.packingstate,0)) as no "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m "
		sqlStr = sqlStr + " 	JOIN [db_storage].[dbo].tbl_ordersheet_detail d "
		sqlStr = sqlStr + " 	ON "
		sqlStr = sqlStr + " 		d.masteridx = m.idx "
		sqlStr = sqlStr + " where d.deldt is Null " & sqlsearch

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
			GetMaxBoxNo = rsget("no")
		rsget.Close
	end function

	public function GetMaxBoxNoByBox()
		dim sqlStr,i

		sqlStr = " select max(isnull(d.packingstate,0)) as no "
		sqlStr = sqlStr + " from " + vbcrlf
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m " + vbcrlf
		sqlStr = sqlStr + " 	JOIN [db_storage].[dbo].tbl_ordersheet_detail d " + vbcrlf
		sqlStr = sqlStr + " 	ON " + vbcrlf
		sqlStr = sqlStr + " 		d.masteridx = m.idx " + vbcrlf
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.idx in ( "
		sqlStr = sqlStr + " 		select  "
		sqlStr = sqlStr + " 		 	m.idx "
		sqlStr = sqlStr + " 		 from  "
		sqlStr = sqlStr + " 		 	[db_storage].[dbo].tbl_shopbalju b  "
		sqlStr = sqlStr + " 		 	JOIN [db_storage].[dbo].tbl_ordersheet_master m  "
		sqlStr = sqlStr + " 		 	ON  "
		sqlStr = sqlStr + " 		 		b.baljucode = m.baljucode  "
		sqlStr = sqlStr + " 		 where  "
		sqlStr = sqlStr + " 			CONVERT(VARCHAR(10),b.baljudate,21) = ( " + vbcrlf
		sqlStr = sqlStr + " 				select " + vbcrlf
		sqlStr = sqlStr + " 					CONVERT(VARCHAR(10),b.baljudate,21) as yyyymmdd " + vbcrlf
		sqlStr = sqlStr + " 		 		from  "
		sqlStr = sqlStr + " 		 			[db_storage].[dbo].tbl_shopbalju b  "
		sqlStr = sqlStr + " 		 			JOIN [db_storage].[dbo].tbl_ordersheet_master m  "
		sqlStr = sqlStr + " 		 			ON  "
		sqlStr = sqlStr + " 		 				b.baljucode = m.baljucode  "
		sqlStr = sqlStr + " 				where " + vbcrlf
		sqlStr = sqlStr + " 					m.idx = " & CStr(FRectMasterIdx) & " " + vbcrlf
		sqlStr = sqlStr + " 			) " + vbcrlf
		sqlStr = sqlStr + " 			and m.baljuid = '" & CStr(FRectShopId) & "' " + vbcrlf
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and m.deldt is Null " + vbcrlf
		sqlStr = sqlStr + " 	and d.deldt is Null " + vbcrlf
		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
			GetMaxBoxNoByBox = rsget("no")
		rsget.Close
	end function

	public function GetCartonBoxNo(shopid, baljudate, innerboxno)
		dim sqlStr,i

		GetCartonBoxNo = ""

		sqlStr = " select top 1 d.cartoonboxno " + vbcrlf
		sqlStr = sqlStr + " from " + vbcrlf
		sqlStr = sqlStr + " 	db_storage.dbo.tbl_cartoonbox_detail d " + vbcrlf
		sqlStr = sqlStr + " 	left join db_storage.dbo.tbl_cartoonbox_master m " + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		m.idx = d.masteridx " + vbcrlf
		sqlStr = sqlStr + " where " + vbcrlf
		sqlStr = sqlStr + " 	1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 	and d.shopid = '" + CStr(shopid) + "' " + vbcrlf
		sqlStr = sqlStr + " 	and d.baljudate = '" + CStr(baljudate) + "' " + vbcrlf
		sqlStr = sqlStr + " 	and d.innerboxno = " + CStr(innerboxno) + " " + vbcrlf
		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			GetCartonBoxNo = rsget("cartoonboxno")
		end if
		rsget.Close
	end function

	public function GetMaxCartonBoxNo(shopid, baljudate)
		dim sqlStr,i

		GetMaxCartonBoxNo = ""

		sqlStr = " select max(d.cartoonboxno) as cartoonboxno " + vbcrlf
		sqlStr = sqlStr + " from " + vbcrlf
		sqlStr = sqlStr + " 	db_storage.dbo.tbl_cartoonbox_detail d " + vbcrlf
		sqlStr = sqlStr + " 	left join db_storage.dbo.tbl_cartoonbox_master m " + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		m.idx = d.masteridx " + vbcrlf
		sqlStr = sqlStr + " where " + vbcrlf
		sqlStr = sqlStr + " 	1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 	and d.shopid = '" + CStr(shopid) + "' " + vbcrlf
		sqlStr = sqlStr + " 	and d.baljudate = '" + CStr(baljudate) + "' " + vbcrlf
		'response.write sqlStr

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			GetMaxCartonBoxNo = rsget("cartoonboxno")
		end if
		rsget.Close
	end function

	'/2016.09.06 한용민 생성
	'//common/offshop/order/orderitem_info_foreign.asp		'//common/popOrderSheet_foreign_excel.asp
	public Sub Getordersheet_foreign_detail
		dim sqlStr, i, sqlsearch, tmpstr

		if FRectcartoonboxmasteridx<>"" then
			sqlsearch = sqlsearch & " and cd.masteridx = "& FRectcartoonboxmasteridx &"" & vbcrlf
		end if
		if FRectMasterIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx = "& FRectMasterIdx &"" & vbcrlf
		end if
		if FRectshopid<>"" then
			sqlsearch = sqlsearch & " and m.baljuid = '"& FRectshopid &"'" & vbcrlf
		end if
		if frectbaljucode<>"" then
			sqlsearch = sqlsearch & " and m.baljucode = '"& frectbaljucode &"'" & vbcrlf
		end if
		if FRectmakerid<>"" then
			sqlsearch = sqlsearch & " and d.makerid = '" & FRectmakerid & "'" & vbcrlf
		end if
		if FRectstartdate<>"" and FRectenddate<>"" then
			if (FRectDateType = "C") then
				'sqlsearch = sqlsearch & " and m.statecd >= '7'" & vbcrlf
				sqlsearch = sqlsearch & " and IsNull(m.beasongdate, m.ipgodate) >= '" & FRectstartdate & "'" & vbcrlf
				sqlsearch = sqlsearch & " and IsNull(m.beasongdate, m.ipgodate) < '" & FRectenddate & "'" & vbcrlf
			elseif (FRectDateType = "J") then
				'sqlsearch = sqlsearch & " and m.statecd >= '0' "
				sqlsearch = sqlsearch & " and m.regdate >= '" & FRectstartdate & "'" & vbcrlf
				sqlsearch = sqlsearch & " and m.regdate < '" & FRectenddate & "'" & vbcrlf
			else
				sqlsearch = sqlsearch & " and b.baljudate>='" & FRectstartdate & "'" & vbcrlf
				sqlsearch = sqlsearch & " and b.baljudate<'" & FRectenddate & "'" & vbcrlf
			end if
		end if
		if FRectChulgoYN<>"" then
			if FRectChulgoYN = "N" then
				sqlsearch = sqlsearch & " and m.statecd < 7 " & vbcrlf
			else
				sqlsearch = sqlsearch & " and m.statecd >= 7" & vbcrlf
			end if
		end if
		if FRectStatecd<>"" then
			sqlsearch = sqlsearch & " and m.statecd = '" & FRectStatecd & "'" & vbcrlf
		end if
		if FRectItemid<>"" then
			sqlsearch = sqlsearch & " and d.itemid = " & FRectItemid & "" & vbcrlf
		end if
		if FRectinnerboxno<>"" then
			sqlsearch = sqlsearch & " and IsNull(d.packingstate, '0') = '" & FRectinnerboxno & "'" & vbcrlf
		end if
		if FRectShopDiv<>"" then
			'참고 : /lib/classes/offshop/offshopchargecls.asp
			if FRectShopDiv="franchisee" then
				'가맹점
				sqlsearch = sqlsearch & " and u.shopdiv in ('3', '4')" & vbcrlf
			elseif FRectShopDiv="direct" then
				'직영점
				sqlsearch = sqlsearch & " and u.shopdiv in ('1', '2')" & vbcrlf
			elseif FRectShopDiv="foreign" then
				'해외
				sqlsearch = sqlsearch & " and u.shopdiv in ('7', '8')" & vbcrlf
			elseif FRectShopDiv="buy" then
				'도매
				sqlsearch = sqlsearch & " and u.shopdiv in ('5', '6')" & vbcrlf
			else
				'기타
				sqlsearch = sqlsearch & " and u.shopdiv = '9'" & vbcrlf
			end if
		end if
		if FRectShowDeleted<>"" then
			if FRectShowDeleted = "N" then
				sqlsearch = sqlsearch & " and m.deldt is null "
				sqlsearch = sqlsearch & " and d.deldt is null "
			end if
		end if
		if FRectMichulgoReason <> "" then
			tmpstr = ""

			if (InStr(1, FRectMichulgoReason, "5", 1) > 0) then
				if (tmpstr <> "") then
					tmpstr = tmpstr & " or "
				end if
				tmpstr = tmpstr & " d.comment = '5일내출고' "
			end if

			if (InStr(1, FRectMichulgoReason, "S", 1) > 0) then
				if (tmpstr <> "") then
					tmpstr = tmpstr & " or "
				end if
				tmpstr = tmpstr & " d.comment = '재고부족' "
			end if

			if (InStr(1, FRectMichulgoReason, "T", 1) > 0) then
				if (tmpstr <> "") then
					tmpstr = tmpstr & " or "
				end if
				tmpstr = tmpstr & " d.comment = '일시품절' "
			end if

			if (InStr(1, FRectMichulgoReason, "D", 1) > 0) then
				if (tmpstr <> "") then
					tmpstr = tmpstr & " or "
				end if
				tmpstr = tmpstr & " d.comment = '단종' "
			end if

			'기타
			if (InStr(1, FRectMichulgoReason, "E", 1) > 0) then
				if (tmpstr <> "") then
					tmpstr = tmpstr & " or "
				end if
				tmpstr = tmpstr & " IsNull(d.comment, '') not in ('5일내출고', '재고부족', '일시품절', '단종') "
			end if

			if (tmpstr <> "") then
				sqlsearch = sqlsearch & " and d.baljuitemno > d.realitemno "
				sqlsearch = sqlsearch & " and (" + CStr(tmpstr) + ") "
			end if
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.baljucode, m.alinkcode, m.baljuid, convert(varchar(10),m.regdate,121) as regdate, convert(varchar(10), b.baljudate, 21) as baljudate" & vbcrlf
		sqlStr = sqlStr & " , m.statecd, m.foreign_statecd" & vbcrlf
		sqlStr = sqlStr & " , IsNull(d.packingstate, '0') as innerboxno, cd.cartoonboxno, d.makerid, d.itemgubun, d.itemid, d.itemoption, d.itemname" & vbcrlf
		sqlStr = sqlStr & " , d.itemoptionname, d.baljuitemno, d.realitemno" & vbcrlf
		sqlStr = sqlStr & " , (case when u.shopdiv='7' then d.foreign_sellcash else d.sellcash end) as sellcash" & vbcrlf
		sqlStr = sqlStr & " , (case when u.shopdiv='7' then d.foreign_suplycash else d.suplycash end) as suplycash" & vbcrlf
		sqlStr = sqlStr & " , (case" & vbcrlf
		sqlStr = sqlStr & " 	when u.shopdiv='7'" & vbcrlf
		sqlStr = sqlStr & " 		then (case when d.foreign_suplycash<>0 and d.foreign_sellcash<>0 then round( (100-((d.foreign_suplycash/d.foreign_sellcash)*100)) ,2) else 0 end)" & vbcrlf
		sqlStr = sqlStr & " 	else" & vbcrlf
		sqlStr = sqlStr & " 		(case when d.suplycash<>0 and d.sellcash<>0 then round( (100-((d.suplycash/d.sellcash)*100)) ,2) else 0 end)" & vbcrlf
		sqlStr = sqlStr & " 	end" & vbcrlf
		sqlStr = sqlStr & " ) as defaultsuplymargin" & vbcrlf
		sqlStr = sqlStr & " , (case when u.shopdiv='7' then d.foreign_suplycash*d.realitemno else d.suplycash*d.realitemno end) as totalsuplycash" & vbcrlf
		sqlStr = sqlStr & " , m.currencyunit" & vbcrlf
		sqlStr = sqlStr & " , (case when i.itemid is not null then i.itemname else ii.shopitemname end) as itemname_10x10" & vbcrlf
		sqlStr = sqlStr & " , (case when i.itemid is not null then o.optionname else ii.shopitemoptionname end) as optionname_10x10" & vbcrlf
		sqlStr = sqlStr & " , i.itemweight, i.deliverOverseas, ic.itemsource as itemsource_10x10, ic.sourcearea as sourcearea_10x10" & vbcrlf
		sqlStr = sqlStr & " , ml.itemname as itemname_en, mo.optionname as optionname_en, ml.itemsource as itemsource_en, ml.sourcearea as sourcearea_en" & vbcrlf
		sqlStr = sqlStr & " , d.idx as detailidx, isnull(l.lcprice,0) as lcprice, l.exchangeRate, l.multipleRate" & vbcrlf
		sqlStr = sqlStr & " , i.listimage,  ii.offimglist" & vbcrlf ''2017/06/26추가
'		sqlStr = sqlStr & " , c1.code_nm as catename1" & vbcrlf
'		sqlStr = sqlStr & " , c2.code_nm as catename2, c2.code_nm_eng as catename_eng2, c2.code_nm_cn_gan as catename_cn_gan2, c2.code_nm_cn_bun as catename_cn_bun2" & vbcrlf
'		sqlStr = sqlStr & " , c3.code_nm as catename3, c3.code_nm_eng as catename_eng3, c3.code_nm_cn_gan as catename_cn_gan3, c3.code_nm_cn_bun as catename_cn_bun3" & vbcrlf
		sqlStr = sqlStr & " , ii.extbarcode , ii.catecdl, ii.catecdm, ii.catecdn" & vbcrlf
		sqlStr = sqlStr & " from [db_storage].[dbo].tbl_ordersheet_master m" & vbcrlf
		sqlStr = sqlStr & " join [db_storage].[dbo].tbl_ordersheet_detail d" & vbcrlf
		sqlStr = sqlStr & " 	on m.idx = d.masteridx" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item i" & vbcrlf
		sqlStr = sqlStr & " 	on d.itemgubun = '10' and d.itemid = i.itemid" & vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_item ii" & vbcrlf
		sqlStr = sqlStr & " 	on d.itemgubun=ii.itemgubun	and d.itemid=ii.shopitemid and d.itemoption=ii.itemoption" & vbcrlf
		sqlStr = sqlStr & " left join [db_storage].[dbo].tbl_shopbalju b" & vbcrlf
		sqlStr = sqlStr & " 	on m.baljucode = b.baljucode and m.baljuid = b.baljuid" & vbcrlf
		sqlStr = sqlStr & " left join [db_storage].[dbo].tbl_cartoonbox_detail cd" & vbcrlf
		sqlStr = sqlStr & " 	on convert(varchar(10), b.baljudate, 21) = convert(varchar(10), cd.baljudate, 21)" & vbcrlf
		sqlStr = sqlStr & " 	and b.baljuid = cd.shopid" & vbcrlf
		sqlStr = sqlStr & " 	and IsNull(d.packingstate, 0) = cd.innerboxno" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_option o" & vbcrlf
		sqlStr = sqlStr & " 	on d.itemgubun = '10'" & vbcrlf
		sqlStr = sqlStr & " 	and d.itemid = o.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and d.itemoption = o.itemoption" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_Contents ic" & vbcrlf
		sqlStr = sqlStr & " 	on i.itemid = ic.itemid" & vbcrlf
		sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang] ml" & vbcrlf
		sqlStr = sqlStr & " 	on d.itemid=ml.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and ml.countryCd='EN'" & vbcrlf
		sqlStr = sqlStr & " 	and ml.useyn='Y'" & vbcrlf
		sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang_option] mo" & vbcrlf
		sqlStr = sqlStr & " 	on d.itemid=mo.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and d.itemoption = mo.itemoption" & vbcrlf
		sqlStr = sqlStr & " 	and mo.countryCd='EN'" & vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user u" & vbcrlf
		sqlStr = sqlStr & " 	on m.baljuid=u.userid" & vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_locale_item l " & VbCrLf
		sqlStr = sqlStr & " 	on m.baljuid = l.shopid " & VbCrLf
		sqlStr = sqlStr & " 	and d.itemgubun = l.itemgubun " & VbCrLf
		sqlStr = sqlStr & " 	and d.itemid = l.shopitemid " & VbCrLf
		sqlStr = sqlStr & " 	and d.itemoption = l.itemoption " & VbCrLf

'		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_large] as c1"
'		sqlStr = sqlStr & " 	on ii.catecdl=c1.code_large"
'		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_mid] as c2"
'		sqlStr = sqlStr & " 	on ii.catecdl=c2.code_large"
'		sqlStr = sqlStr & " 	and ii.catecdm=c2.code_mid"
'		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_small] as c3"
'		sqlStr = sqlStr & " 	on ii.catecdl=c3.code_large"
'		sqlStr = sqlStr & " 	and ii.catecdm=c3.code_mid"
'		sqlStr = sqlStr & " 	and ii.catecdn=c3.code_small"

		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by cartoonboxno asc, innerboxno asc, d.makerid asc, d.itemgubun asc, d.itemid asc, d.itemoption asc" & vbcrlf		'm.baljucode desc

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CStorageDetailItem
				FItemList(i).Fstatecd    		= rsget("statecd")
				FItemList(i).Fforeign_statecd    		= rsget("foreign_statecd")
				FItemList(i).fitemweight    		= rsget("itemweight")
				FItemList(i).fdeliverOverseas    		= rsget("deliverOverseas")
				FItemList(i).fbaljucode    		= rsget("baljucode")
				FItemList(i).falinkcode    		= rsget("alinkcode")
				FItemList(i).fbaljuid    		= rsget("baljuid")
				FItemList(i).fregdate    		= rsget("regdate")
				FItemList(i).fbaljudate    		= rsget("baljudate")
				FItemList(i).finnerboxno    		= rsget("innerboxno")
				FItemList(i).fcartoonboxno    		= rsget("cartoonboxno")
				FItemList(i).fmakerid    		= rsget("makerid")
				FItemList(i).fitemgubun    		= rsget("itemgubun")
				FItemList(i).fitemid    		= rsget("itemid")
				FItemList(i).fitemoption    		= rsget("itemoption")
				FItemList(i).fitemname    		= db2html(rsget("itemname"))
				FItemList(i).fitemoptionname    		= db2html(rsget("itemoptionname"))
				FItemList(i).Fbaljuitemno    		= rsget("baljuitemno")
				FItemList(i).frealitemno    		= rsget("realitemno")
				FItemList(i).fsellcash    		= rsget("sellcash")
				FItemList(i).fsuplycash    		= rsget("suplycash")
				FItemList(i).fdefaultsuplymargin    		= rsget("defaultsuplymargin")
				FItemList(i).ftotalsuplycash    		= rsget("totalsuplycash")
				FItemList(i).fcurrencyunit    		= rsget("currencyunit")
				FItemList(i).fitemname_10x10    		= db2html(rsget("itemname_10x10"))
				FItemList(i).foptionname_10x10    		= db2html(rsget("optionname_10x10"))
				FItemList(i).fitemsource_10x10    		= db2html(rsget("itemsource_10x10"))
				FItemList(i).fsourcearea_10x10    		= db2html(rsget("sourcearea_10x10"))
				FItemList(i).fitemname_en    		= db2html(rsget("itemname_en"))
				FItemList(i).foptionname_en    		= db2html(rsget("optionname_en"))
				FItemList(i).fitemsource_en    		= db2html(rsget("itemsource_en"))
				FItemList(i).fsourcearea_en    		= db2html(rsget("sourcearea_en"))
				FItemList(i).fdetailidx    		= rsget("detailidx")
				FItemList(i).flcprice         	= rsget("lcprice")
				FItemList(i).fexchangeRate     	= rsget("exchangeRate")
				FItemList(i).fmultipleRate     	= rsget("multipleRate")
'				FItemList(i).fcatename1    = db2html(rsget("catename1"))
'				FItemList(i).fcatename2    = db2html(rsget("catename2"))
'				FItemList(i).fcatename_cn_gan2    = db2html(rsget("catename_cn_gan2"))
'				FItemList(i).fcatename_cn_bun2    = db2html(rsget("catename_cn_bun2"))
'				FItemList(i).fcatename3    = db2html(rsget("catename3"))
'				FItemList(i).fcatename_cn_gan3    = db2html(rsget("catename_cn_gan3"))
'				FItemList(i).fcatename_cn_bun3    = db2html(rsget("catename_cn_bun3"))
'				FItemList(i).fcatename_eng2    = db2html(rsget("catename_eng2"))
'				FItemList(i).fcatename_eng3    = db2html(rsget("catename_eng3"))
				FItemList(i).fcatecdl     	= rsget("catecdl")
				FItemList(i).fcatecdm     	= rsget("catecdm")
				FItemList(i).fcatecdn     	= rsget("catecdn")
				FItemList(i).fextbarcode     	= rsget("extbarcode")

				if (IsNull(rsget("listimage"))) then
					FItemList(i).Fmainimageurl  = webImgSSLUrl & "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("offimglist")
				else
					FItemList(i).Fmainimageurl  = webImgSSLUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
				end if
				
				i=i+1
				rsget.movenext
			loop
		end if
	rsget.close
	end sub

	'/common/barcode/inc_barcodeprint_off.asp	'/common/barcode/inc_paperbarcodeprint_off.asp
	public Sub Getjumundetaillist
		dim sqlStr, i, iCountrylangCd, sqlsearch

        if (FRectForeignOrderShopid<>"") then
            iCountrylangCd= GetShopCountrylangcd(FRectForeignOrderShopid)
        end if

		if frectitembarcodearr<>"" then
			frectitembarcodearr = replace(frectitembarcodearr, ",", "','")
			frectitembarcodearr = "'" & frectitembarcodearr & "'"
			sqlsearch = sqlsearch & " and [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun, d.itemid, d.itemoption) in ("& frectitembarcodearr &")"
		end if
		if frectbaljucode<>"" then
			sqlsearch = sqlsearch & " and m.baljucode='"& frectbaljucode &"'"
		end if
		if FRectMasterIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx="& FRectMasterIdx &""
		end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and d.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and d.itemid in (" + FRectItemid + ")"
            end if
        end if
        if (FRectMakerid <> "") then
            sqlsearch = sqlsearch & " and d.makerid='" + FRectMakerid + "'"
        end if
        ''//상품명 검색 수정  2016/04/04 최대 N건 가능 by eastone
        if (FRectItemName <> "") then
			''addSql = addSql & " and s.shopitemname like '%" + html2db(FRectItemName) + "%'"

			sqlStr = " select top 1000 B.itemid into #TMPSearchItem"
			sqlStr = sqlStr + " from [DBAPPWISH].db_AppWish.dbo.tbl_item_SearchBase B with (nolock)"
			if (FRectMakerid <> "") then
    			sqlStr = sqlStr + " Join [DBAPPWISH].[db_AppWish].dbo.tbl_item ai with (nolock)"
            	sqlStr = sqlStr + " on B.itemid=ai.itemid"
            	sqlStr = sqlStr + " and ai.makerid='"&FRectMakerid&"'"
	        end if
	        sqlStr = sqlStr + " where contains(B.searchKey,'""" + CStr(FRectItemName) + """') "
            sqlStr = sqlStr + " order by B.itemid desc "
            dbget.Execute sqlStr

            sqlsearch = sqlsearch & " and d.itemid in (select itemid from #TMPSearchItem )"  ''2016/04/04
		end if
		if FRectPrdCode<>"" then
			if (Len(FRectPrdCode) = 12) then
				'sqlsearch = sqlsearch + " 	and d.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemid = " + RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			else
				'sqlsearch = sqlsearch + " 	and d.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemid = " + RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			end if
		end if
		if FRectGeneralBarcode<>"" then
			sqlsearch = sqlsearch + " 	and s.extbarcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if
        if (FRectSellYN="YS") then
            sqlsearch = sqlsearch & " and i.sellyn<>'N'"
        elseif( FRectSellYN="SR") then
        	  sqlsearch = sqlsearch & " and i.sellyn='N' and r.itemid is not null "
        elseif (FRectSellYN <> "") then
            sqlsearch = sqlsearch & " and i.sellyn='" + FRectSellYN + "'"
        end if
        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and s.isusing='" + FRectIsUsing + "'"
        end if
        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " 	and s.catecdl='" + FRectCDL + "'" & VbCrLf
        end if
        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " 	and s.catecdm='" + FRectCDM + "'" & VbCrLf
        end if
        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " 	and s.catecdn='" + FRectCDS + "'" & VbCrLf
        end if
		if FRectShopItemName<>"" then
			sqlsearch = sqlsearch + " 	and f.lcitemname like '%" + FRectShopItemName + "%'" & VbCrLf
		end if
		if FRectCurrentStockExist = "Y" then
			sqlsearch = sqlsearch + " 	and sc.shopid is not null " & VbCrLf
		end if
		if FRectRealStockOneMore = "Y" then
			sqlsearch = sqlsearch + " 	and IsNull(sc.realstockno, 0) > 0 " & VbCrLf
		end if
		if FRectShopItemNameInserted = "Y" then
			sqlsearch = sqlsearch + " 	and f.shopid is not null " & VbCrLf
		end if
		if FRectshopid <> "" then
			sqlsearch = sqlsearch + " 	and m.baljuid = '"& FRectshopid &"'" & VbCrLf
		end if
		if FRectitemgubun <> "" then
			sqlsearch = sqlsearch + " 	and d.itemgubun='"& FRectitemgubun &"'" & VbCrLf
		end if

		sqlStr = " select count(d.idx) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m with (nolock)"
		sqlStr = sqlStr + " 	JOIN [db_storage].[dbo].tbl_ordersheet_detail d with (nolock)"
		sqlStr = sqlStr + " 	ON "
		sqlStr = sqlStr + " 		d.masteridx = m.idx "
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun = '10' "
		sqlStr = sqlStr + " 		and d.itemid = i.itemid "
		sqlStr = sqlStr & " 	left join db_item.dbo.tbl_item_Contents ic with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 		on i.itemid = ic.itemid" & vbcrlf
		sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun=s.itemgubun "
		sqlStr = sqlStr + " 		and d.itemid=s.shopitemid "
		sqlStr = sqlStr + " 		and d.itemoption=s.itemoption "
		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_stock o with (nolock)"
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun=o.itemgubun "
		sqlStr = sqlStr + " 		and d.itemid=o.itemid "
		sqlStr = sqlStr + " 		and d.itemoption=o.itemoption "
		sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_locale_item f with (nolock)"
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and f.shopid = '" & CStr(FRectForeignOrderShopid) & "' "
		sqlStr = sqlStr + " 		and d.itemgubun=f.itemgubun "
		sqlStr = sqlStr + " 		and d.itemid=f.shopitemid "
		sqlStr = sqlStr + " 		and d.itemoption=f.itemoption "
		sqlStr = sqlStr + " 	left join [db_partner].[dbo].tbl_partner p with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and d.makerid=p.id " + vbcrlf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c c with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on d.makerid=c.userid " & vbcrlf
        sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang_price] r with (nolock)"
        sqlStr = sqlStr & "  	on r.sitename = 'WSLWEB'"
        sqlStr = sqlStr & "  	and r.currencyUnit = 'USD'"
        sqlStr = sqlStr & "  	and d.itemid=r.itemid"
		sqlStr = sqlStr & " left join [db_summary].[dbo].tbl_current_shopstock_summary sc with (nolock)" & VbCrLf
		sqlStr = sqlStr & " 	on sc.shopid = '" & CStr(FRectForeignOrderShopid) & "' " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemgubun = d.itemgubun " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemid = d.itemid " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemoption = d.itemoption " & VbCrLf
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on d.itemid = a.itemid"	& vbcrlf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang] Lni with (nolock)"
            sqlStr = sqlStr + "  on Lni.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and s.itemgubun='10'"
            sqlStr = sqlStr + "  and s.shopitemid=Lni.itemid"

            sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang_option] Lno with (nolock)"
            sqlStr = sqlStr + "  on Lno.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and s.itemgubun='10'"
            sqlStr = sqlStr + "  and s.shopitemid=Lno.itemid"
            sqlStr = sqlStr + "  and s.itemoption=Lno.itemoption"
		end if

		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 " & sqlsearch
		sqlStr = sqlStr + " 	and d.deldt is Null " + vbcrlf

		if (FRectBoxNo <> "") then
			sqlStr = sqlStr + " 	and isnull(d.packingstate,0) = " & CStr(FRectBoxNo) & " " + vbcrlf
		end if

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " " + vbcrlf
		sqlStr = sqlStr + " 	m.baljucode " + vbcrlf
		sqlStr = sqlStr + " 	, d.idx " + vbcrlf
		sqlStr = sqlStr + " 	, d.masteridx " + vbcrlf
		sqlStr = sqlStr + " 	, [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun,d.itemid,d.itemoption)  as prdcode " + vbcrlf
		sqlStr = sqlStr + " 	, [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun,d.itemid,d.itemoption)  as prdbarcode " + vbcrlf
		sqlStr = sqlStr + " 	, d.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	, d.itemid " + vbcrlf
		sqlStr = sqlStr + " 	, d.itemoption " + vbcrlf
		sqlStr = sqlStr + " 	, d.itemoptionname " + vbcrlf
		sqlStr = sqlStr + " 	, d.itemname as prdname " + vbcrlf
		sqlStr = sqlStr + " 	, d.makerid as locationid " + vbcrlf
		sqlStr = sqlStr + " 	, p.company_name as locationname, c.socname, c.socname_kor" + vbcrlf
		sqlStr = sqlStr & " 	, d.sellcash as customerprice, d.sellcash as sellprice" + vbcrlf

		If (FRectShowSupplyCash = "Y") Then
			sqlStr = sqlStr & " 	, d.suplycash as supplyprice" + vbcrlf			'// buycash 는 매입가, suplycash 는 공급가
		Else
			sqlStr = sqlStr & " 	, d.buycash as supplyprice" + vbcrlf			'// buycash 는 매입가, suplycash 는 공급가
		End If

		sqlStr = sqlStr + " 	, d.baljuitemno as requestedno " + vbcrlf
		sqlStr = sqlStr + " 	, d.realitemno as fixedno " + vbcrlf
		sqlStr = sqlStr + " 	, (CASE " + vbcrlf
		sqlStr = sqlStr + " 			WHEN m.deldt is not null or d.deldt is not null THEN 'N' " + vbcrlf
		sqlStr = sqlStr + " 			ELSE 'Y' " + vbcrlf
		sqlStr = sqlStr + " 		END " + vbcrlf
		sqlStr = sqlStr + " 	) as useyn " + vbcrlf
		sqlStr = sqlStr + " 	, o.barcode as generalbarcode " + vbcrlf
		sqlStr = sqlStr + " 	, i.smallimage, i.listimage" + vbcrlf
		sqlStr = sqlStr + " 	, s.offimgsmall, s.offimglist, s.extbarcode, s.itemcopy" + vbcrlf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + " 	, isnull(isNULL(Lni.itemname,f.lcitemname),d.itemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " 	, isnull(isNULL(Lno.optionname,f.lcitemoptionname),d.itemoptionname) as lcitemoptionname " + vbcrlf
			sqlStr = sqlStr & " 	, Lni.sourcearea as sourcearea_en, Lni.itemsource as itemsource_en, Lni.itemsize as itemsize_en" + vbcrlf
		else
    		sqlStr = sqlStr + " 	, isnull(f.lcitemname,d.itemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " 	, isnull(f.lcitemoptionname,d.itemoptionname) as lcitemoptionname" + vbcrlf
    		sqlStr = sqlStr & " 	, '' as sourcearea_en, '' as itemsource_en, '' as itemsize_en" + vbcrlf
	    end if

		sqlStr = sqlStr + " 	, isnull(f.lcprice,0) as lcprice" + vbcrlf
		sqlStr = sqlStr + " 	, isnull(d.packingstate,0) as boxno " + vbcrlf
		sqlStr = sqlStr & " , (CASE " & VbCrLf
		sqlStr = sqlStr & " 		WHEN IsNull(sc.realstockno, 0) <= 0 THEN 0 " & VbCrLf
		sqlStr = sqlStr & " 		ELSE sc.realstockno " & VbCrLf
		sqlStr = sqlStr & " 	END " & VbCrLf
		sqlStr = sqlStr & " ) as realstockno " & VbCrLf
		sqlStr = sqlStr & " , (CASE" & VbCrLf
		sqlStr = sqlStr & " 		WHEN IsNull(f.lcprice, 0) > 0 THEN 'Y'" & VbCrLf
		sqlStr = sqlStr & " 		ELSE 'N'" & VbCrLf
		sqlStr = sqlStr & " 	END" & VbCrLf
		sqlStr = sqlStr & " ) as saleyn" & VbCrLf
		sqlStr = sqlStr & " , c1.code_nm as catename1" & vbcrlf
		sqlStr = sqlStr & " , c2.code_nm as catename2, c2.code_nm_eng as catename_eng2, c2.code_nm_cn_gan as catename_cn_gan2, c2.code_nm_cn_bun as catename_cn_bun2" & vbcrlf
		sqlStr = sqlStr & " , c3.code_nm as catename3, c3.code_nm_eng as catename_eng3, c3.code_nm_cn_gan as catename_cn_gan3, c3.code_nm_cn_bun as catename_cn_bun3" & vbcrlf
		sqlStr = sqlStr & " , ic.sourcearea as sourcearea_10x10, ic.itemsource as itemsource_10x10, ic.itemsize as itemsize_10x10"
        sqlstr = sqlstr & " , isnull((case when d.itemgubun='10' then i.itemrackcode else s.offitemrackcode end),'') as itemrackcode"
        sqlstr = sqlstr & " , isnull(c.prtidx,'') as prtidx, isnull(a.subitemrackcode,'') as subitemrackcode" & vbcrlf
		sqlStr = sqlStr + " from " + vbcrlf
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	JOIN [db_storage].[dbo].tbl_ordersheet_detail d with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	ON " + vbcrlf
		sqlStr = sqlStr + " 		d.masteridx = m.idx " + vbcrlf
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item i with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemgubun = '10' " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemid = i.itemid " + vbcrlf
		sqlStr = sqlStr & " 	left join db_item.dbo.tbl_item_Contents ic with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 		on i.itemid = ic.itemid" & vbcrlf
		sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_item s with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemgubun=s.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemid=s.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemoption=s.itemoption " + vbcrlf
		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_stock o with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemgubun=o.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemid=o.itemid " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemoption=o.itemoption " + vbcrlf
		sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_locale_item f with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and f.shopid = '" & CStr(FRectForeignOrderShopid) & "' "
		sqlStr = sqlStr + " 		and d.itemgubun=f.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemid=f.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemoption=f.itemoption " + vbcrlf
		sqlStr = sqlStr + " 	left join [db_partner].[dbo].tbl_partner p with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and d.makerid=p.id " + vbcrlf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c c with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on d.makerid=c.userid " & vbcrlf
        sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang_price] r with (nolock)"
        sqlStr = sqlStr & "  	on r.sitename = 'WSLWEB'"
        sqlStr = sqlStr & "  	and r.currencyUnit = 'USD'"
        sqlStr = sqlStr & "  	and d.itemid=r.itemid"
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_large] as c1 with (nolock)"
		sqlStr = sqlStr & " 	on s.catecdl=c1.code_large"
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_mid] as c2 with (nolock)"
		sqlStr = sqlStr & " 	on s.catecdl=c2.code_large"
		sqlStr = sqlStr & " 	and s.catecdm=c2.code_mid"
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_small] as c3 with (nolock)"
		sqlStr = sqlStr & " 	on s.catecdl=c3.code_large"
		sqlStr = sqlStr & " 	and s.catecdm=c3.code_mid"
		sqlStr = sqlStr & " 	and s.catecdn=c3.code_small"
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on d.itemid = a.itemid"	& vbcrlf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang] Lni with (nolock)"
            sqlStr = sqlStr + "  on Lni.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and d.itemgubun='10'"
            sqlStr = sqlStr + "  and d.itemid=Lni.itemid"

            sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang_option] Lno with (nolock)"
            sqlStr = sqlStr + "  on Lno.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and d.itemgubun='10'"
            sqlStr = sqlStr + "  and d.itemid=Lno.itemid"
            sqlStr = sqlStr + "  and d.itemoption=Lno.itemoption"
		end if

		sqlStr = sqlStr & " left join [db_summary].[dbo].tbl_current_shopstock_summary sc with (nolock)" & VbCrLf
		sqlStr = sqlStr & " 	on sc.shopid = '" & CStr(FRectForeignOrderShopid) & "' " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemgubun = d.itemgubun " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemid = d.itemid " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemoption = d.itemoption " & VbCrLf
		sqlStr = sqlStr + " where " + vbcrlf
		sqlStr = sqlStr + " 	1 = 1 " & sqlsearch
		sqlStr = sqlStr + " 	and d.deldt is Null " + vbcrlf

		if (FRectBoxNo <> "") then
			sqlStr = sqlStr + " 	and isnull(d.packingstate,0) = " & CStr(FRectBoxNo) & " " + vbcrlf
		end if

		sqlStr = sqlStr + " order by isnull(d.packingstate,0), d.itemgubun, d.itemid, d.itemoption " + vbcrlf
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient

		'response.write sqlStr
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new CStorageDetailItem

				FItemList(i).Fprdcode    		= db2html(rsget("prdcode"))
				FItemList(i).Fprdname    		= db2html(rsget("prdname"))
				FItemList(i).Flocationid    	= db2html(rsget("locationid"))
				FItemList(i).Flocationname    	= db2html(rsget("locationname"))
				FItemList(i).fsocname    = db2html(rsget("socname"))
				FItemList(i).fsocname_kor    = db2html(rsget("socname_kor"))
				FItemList(i).Fitemgubun    		= db2html(rsget("itemgubun"))
				FItemList(i).Fitemid    		= db2html(rsget("itemid"))
				FItemList(i).Fitemoption    	= db2html(rsget("itemoption"))
				FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
				FItemList(i).fitemcopy    = db2html(rsget("itemcopy"))
				FItemList(i).Fprdbarcode    	= db2html(rsget("prdbarcode"))
				FItemList(i).Fgeneralbarcode    = db2html(rsget("generalbarcode"))
				FItemList(i).Fcustomerprice    	= rsget("customerprice")
				FItemList(i).Fsupplyprice    	= rsget("supplyprice")
				FItemList(i).Fsellprice       	= rsget("sellprice")

				if (IsNull(rsget("listimage")) = True) then
					FItemList(i).Fmainimageurl  = webImgSSLUrl & "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("offimglist")
				else
					FItemList(i).Fmainimageurl  = webImgSSLUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
				end if

				FItemList(i).Fuseyn    			= db2html(rsget("useyn"))
				FItemList(i).Flcitemname    	= db2html(rsget("lcitemname"))
				FItemList(i).Flcitemoptionname  = db2html(rsget("lcitemoptionname"))
				FItemList(i).Flcprice    		= db2html(rsget("lcprice"))
				FItemList(i).Frealstockno       = rsget("realstockno")
				FItemList(i).fsaleyn    		= rsget("saleyn")
				FItemList(i).Fbaljucode    		= rsget("baljucode")
				FItemList(i).fpublicbarcode    		= rsget("extbarcode")
				FItemList(i).Fidx        		= rsget("idx")
				FItemList(i).Fmasteridx       	= rsget("masteridx")
				FItemList(i).Frequestedno    	= db2html(rsget("requestedno"))
				FItemList(i).Ffixedno    		= db2html(rsget("fixedno"))
				FItemList(i).Fboxno		       	= rsget("boxno")
				FItemList(i).fcatename1    = db2html(rsget("catename1"))
				FItemList(i).fcatename2    = db2html(rsget("catename2"))
				FItemList(i).fcatename_cn_gan2    = db2html(rsget("catename_cn_gan2"))
				FItemList(i).fcatename_cn_bun2    = db2html(rsget("catename_cn_bun2"))
				FItemList(i).fcatename3    = db2html(rsget("catename3"))
				FItemList(i).fcatename_cn_gan3    = db2html(rsget("catename_cn_gan3"))
				FItemList(i).fcatename_cn_bun3    = db2html(rsget("catename_cn_bun3"))
				FItemList(i).fcatename_eng2    = db2html(rsget("catename_eng2"))
				FItemList(i).fcatename_eng3    = db2html(rsget("catename_eng3"))
				FItemList(i).fsourcearea_10x10    		= db2html(rsget("sourcearea_10x10"))
				FItemList(i).fitemsource_10x10    		= db2html(rsget("itemsource_10x10"))
				FItemList(i).fitemsize_10x10    		= db2html(rsget("itemsize_10x10"))
				FItemList(i).fsourcearea_en    		= db2html(rsget("sourcearea_en"))
				FItemList(i).fitemsource_en    		= db2html(rsget("itemsource_en"))
				FItemList(i).fitemsize_en    		= db2html(rsget("itemsize_en"))
				FItemList(i).Fitemrackcode = rsget("itemrackcode")
				FItemList(i).fprtidx = rsget("prtidx")
				FItemList(i).fsubitemrackcode = rsget("subitemrackcode")

				i=i+1
				rsget.movenext
			loop
		end if
	rsget.close
	end sub

	public Sub GetStorageDetailList
		dim sqlStr, i, iCountrylangCd, sqlsearch

        if (FRectForeignOrderShopid<>"") then
            iCountrylangCd= GetShopCountrylangcd(FRectForeignOrderShopid)
        end if

		if frectitembarcodearr<>"" then
			frectitembarcodearr = replace(frectitembarcodearr, ",", "','")
			frectitembarcodearr = "'" & frectitembarcodearr & "'"
			sqlsearch = sqlsearch & " and [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun, d.itemid, d.itemoption) in ("& frectitembarcodearr &")"
		end if
		if frectbaljucode<>"" then
			sqlsearch = sqlsearch & " and m.baljucode='"& frectbaljucode &"'"
		end if
		if FRectMasterIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx="& FRectMasterIdx &""
		end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and d.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and d.itemid in (" + FRectItemid + ")"
            end if
        end if
        if (FRectMakerid <> "") then
            sqlsearch = sqlsearch & " and d.makerid='" + FRectMakerid + "'"
        end if
        ''//상품명 검색 수정  2016/04/04 최대 N건 가능 by eastone
        if (FRectItemName <> "") then
			''addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"

			sqlStr = " select top 1000 B.itemid into #TMPSearchItem"
			sqlStr = sqlStr + " from [DBAPPWISH].db_AppWish.dbo.tbl_item_SearchBase B"
			if (FRectMakerid <> "") then
    			sqlStr = sqlStr + " Join [DBAPPWISH].[db_AppWish].dbo.tbl_item ai"
            	sqlStr = sqlStr + " on B.itemid=ai.itemid"
            	sqlStr = sqlStr + " and ai.makerid='"&FRectMakerid&"'"
	        end if
	        sqlStr = sqlStr + " where contains(B.searchKey,'""" + CStr(FRectItemName) + """') "
            sqlStr = sqlStr + " order by B.itemid desc "
            dbget.Execute sqlStr

            sqlsearch = sqlsearch & " and d.itemid in (select itemid from #TMPSearchItem )"  ''2016/04/04
		end if
		if FRectPrdCode<>"" then
			if (Len(FRectPrdCode) = 12) then
				'sqlsearch = sqlsearch + " 	and d.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemid = " + RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			else
				'sqlsearch = sqlsearch + " 	and d.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemid = " + RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			end if
		end if
		if FRectGeneralBarcode<>"" then
			sqlsearch = sqlsearch + " 	and o.barcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if
        if (FRectSellYN="YS") then
            sqlsearch = sqlsearch & " and i.sellyn<>'N'"
        elseif( FRectSellYN="SR") then
        	  sqlsearch = sqlsearch & " and i.sellyn='N' and r.itemid is not null "
        elseif (FRectSellYN <> "") then
            sqlsearch = sqlsearch & " and i.sellyn='" + FRectSellYN + "'"
        end if
        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and i.isusing='" + FRectIsUsing + "'"
        end if

		sqlStr = " select count(d.idx) as cnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m "
		sqlStr = sqlStr + " 	JOIN [db_storage].[dbo].tbl_ordersheet_detail d "
		sqlStr = sqlStr + " 	ON "
		sqlStr = sqlStr + " 		d.masteridx = m.idx "
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item i "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun = '10' "
		sqlStr = sqlStr + " 		and d.itemid = i.itemid "
		sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun=s.itemgubun "
		sqlStr = sqlStr + " 		and d.itemid=s.shopitemid "
		sqlStr = sqlStr + " 		and d.itemoption=s.itemoption "
		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_stock o "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and d.itemgubun=o.itemgubun "
		sqlStr = sqlStr + " 		and d.itemid=o.itemid "
		sqlStr = sqlStr + " 		and d.itemoption=o.itemoption "
		sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_locale_item f "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and f.shopid = '" & CStr(FRectForeignOrderShopid) & "' "
		sqlStr = sqlStr + " 		and d.itemgubun=f.itemgubun "
		sqlStr = sqlStr + " 		and d.itemid=f.shopitemid "
		sqlStr = sqlStr + " 		and d.itemoption=f.itemoption "
		sqlStr = sqlStr + " 	left join [db_partner].[dbo].tbl_partner p " + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and d.makerid=p.id " + vbcrlf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c c" & vbcrlf
		sqlStr = sqlStr & " 	on d.makerid=c.userid " & vbcrlf
        sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang_price] r"
        sqlStr = sqlStr & "  	on r.sitename = 'WSLWEB'"
        sqlStr = sqlStr & "  	and r.currencyUnit = 'USD'"
        sqlStr = sqlStr & "  	and d.itemid=r.itemid"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 " & sqlsearch
		sqlStr = sqlStr + " 	and d.deldt is Null " + vbcrlf

		if (FRectBoxNo <> "") then
			sqlStr = sqlStr + " 	and isnull(d.packingstate,0) = " & CStr(FRectBoxNo) & " " + vbcrlf
		end if

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " " + vbcrlf
		sqlStr = sqlStr + " 	m.baljucode " + vbcrlf
		sqlStr = sqlStr + " 	, d.idx " + vbcrlf
		sqlStr = sqlStr + " 	, d.masteridx " + vbcrlf
		sqlStr = sqlStr + " 	, [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun,d.itemid,d.itemoption)  as prdcode " + vbcrlf
		sqlStr = sqlStr + " 	, [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun,d.itemid,d.itemoption)  as prdbarcode " + vbcrlf
		sqlStr = sqlStr + " 	, d.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	, d.itemid " + vbcrlf
		sqlStr = sqlStr + " 	, d.itemoption " + vbcrlf
		sqlStr = sqlStr + " 	, d.itemoptionname " + vbcrlf
		sqlStr = sqlStr + " 	, d.itemname as prdname " + vbcrlf
		sqlStr = sqlStr + " 	, d.makerid as locationid " + vbcrlf
		sqlStr = sqlStr + " 	, p.company_name as locationname, c.socname, c.socname_kor" + vbcrlf
		sqlStr = sqlStr & " 	, d.sellcash as customerprice, d.sellcash as sellprice" + vbcrlf
		sqlStr = sqlStr & " 	, d.buycash" + vbcrlf	'// buycash 매입가

		If (FRectShowSupplyCash = "Y") Then
			sqlStr = sqlStr & " 	, d.suplycash as supplyprice" + vbcrlf			'// buycash 는 매입가, suplycash 는 공급가
		Else
			sqlStr = sqlStr & " 	, d.buycash as supplyprice" + vbcrlf			'// buycash 는 매입가, suplycash 는 공급가
		End If

		sqlStr = sqlStr + " 	, d.baljuitemno as requestedno " + vbcrlf
		sqlStr = sqlStr + " 	, d.realitemno as fixedno " + vbcrlf
		sqlStr = sqlStr + " 	, (CASE " + vbcrlf
		sqlStr = sqlStr + " 			WHEN m.deldt is not null or d.deldt is not null THEN 'N' " + vbcrlf
		sqlStr = sqlStr + " 			ELSE 'Y' " + vbcrlf
		sqlStr = sqlStr + " 		END " + vbcrlf
		sqlStr = sqlStr + " 	) as useyn " + vbcrlf
		sqlStr = sqlStr + " 	, o.barcode as generalbarcode " + vbcrlf
		sqlStr = sqlStr + " 	, i.smallimage, i.listimage" + vbcrlf
		sqlStr = sqlStr + " 	, s.offimgsmall, s.offimglist, s.extbarcode" + vbcrlf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + " 	, isnull(isNULL(Lni.itemname,f.lcitemname),d.itemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " 	, isnull(isNULL(Lno.optionname,f.lcitemoptionname),d.itemoptionname) as lcitemoptionname " + vbcrlf
		else
    		sqlStr = sqlStr + " 	, isnull(f.lcitemname,d.itemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " 	, isnull(f.lcitemoptionname,d.itemoptionname) as lcitemoptionname" + vbcrlf
	    end if

		sqlStr = sqlStr + " 	, isnull(f.lcprice,0) as lcprice" + vbcrlf
		sqlStr = sqlStr + " 	, isnull(d.packingstate,0) as boxno " + vbcrlf
		sqlStr = sqlStr + " from " + vbcrlf
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m " + vbcrlf
		sqlStr = sqlStr + " 	JOIN [db_storage].[dbo].tbl_ordersheet_detail d " + vbcrlf
		sqlStr = sqlStr + " 	ON " + vbcrlf
		sqlStr = sqlStr + " 		d.masteridx = m.idx " + vbcrlf
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item i " + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemgubun = '10' " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemid = i.itemid " + vbcrlf
		sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_item s " + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemgubun=s.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemid=s.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemoption=s.itemoption " + vbcrlf
		sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_option_stock o " + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemgubun=o.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemid=o.itemid " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemoption=o.itemoption " + vbcrlf
		sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_locale_item f " + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and f.shopid = '" & CStr(FRectForeignOrderShopid) & "' "
		sqlStr = sqlStr + " 		and d.itemgubun=f.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemid=f.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 		and d.itemoption=f.itemoption " + vbcrlf
		sqlStr = sqlStr + " 	left join [db_partner].[dbo].tbl_partner p " + vbcrlf
		sqlStr = sqlStr + " 	on " + vbcrlf
		sqlStr = sqlStr + " 		1 = 1 " + vbcrlf
		sqlStr = sqlStr + " 		and d.makerid=p.id " + vbcrlf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c c" & vbcrlf
		sqlStr = sqlStr & " 	on d.makerid=c.userid " & vbcrlf
        sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang_price] r"
        sqlStr = sqlStr & "  	on r.sitename = 'WSLWEB'"
        sqlStr = sqlStr & "  	and r.currencyUnit = 'USD'"
        sqlStr = sqlStr & "  	and d.itemid=r.itemid"

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang] Lni"
            sqlStr = sqlStr + "  on Lni.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and d.itemgubun='10'"
            sqlStr = sqlStr + "  and d.itemid=Lni.itemid"

            sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang_option] Lno"
            sqlStr = sqlStr + "  on Lno.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  and d.itemgubun='10'"
            sqlStr = sqlStr + "  and d.itemid=Lno.itemid"
            sqlStr = sqlStr + "  and d.itemoption=Lno.itemoption"
		end if

		sqlStr = sqlStr + " where " + vbcrlf
		sqlStr = sqlStr + " 	1 = 1 " & sqlsearch
		sqlStr = sqlStr + " 	and d.deldt is Null " + vbcrlf

		if (FRectBoxNo <> "") then
			sqlStr = sqlStr + " 	and isnull(d.packingstate,0) = " & CStr(FRectBoxNo) & " " + vbcrlf
		end if

		sqlStr = sqlStr + " order by isnull(d.packingstate,0), d.itemgubun, d.itemid, d.itemoption " + vbcrlf
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient

		'response.write sqlStr
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new CStorageDetailItem
				FItemList(i).Fbaljucode    		= rsget("baljucode")
				FItemList(i).fpublicbarcode    		= rsget("extbarcode")
				FItemList(i).Fidx        		= rsget("idx")
				FItemList(i).Fmasteridx       	= rsget("masteridx")
				FItemList(i).Fprdcode    		= db2html(rsget("prdcode"))
				FItemList(i).Fprdbarcode    	= db2html(rsget("prdbarcode"))
				FItemList(i).Fitemgubun    		= db2html(rsget("itemgubun"))
				FItemList(i).Fitemid    		= db2html(rsget("itemid"))
				FItemList(i).Fitemoption    	= db2html(rsget("itemoption"))
				FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
				FItemList(i).Fprdname    		= db2html(rsget("prdname"))
				FItemList(i).Flocationid    	= db2html(rsget("locationid"))
				FItemList(i).Flocationname    	= db2html(rsget("locationname"))
				FItemList(i).fsocname    = db2html(rsget("socname"))
				FItemList(i).fsocname_kor    = db2html(rsget("socname_kor"))
				FItemList(i).Fcustomerprice    	= rsget("customerprice")
				FItemList(i).Fsellprice       	= rsget("sellprice")
				FItemList(i).Fsupplyprice    	= rsget("supplyprice")
				FItemList(i).fbuycash    	= rsget("buycash")
				FItemList(i).Frequestedno    	= db2html(rsget("requestedno"))
				FItemList(i).Ffixedno    		= db2html(rsget("fixedno"))
				FItemList(i).Fuseyn    			= db2html(rsget("useyn"))

				if (IsNull(rsget("listimage")) = True) then
					FItemList(i).Fmainimageurl  = webImgSSLUrl & "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("offimglist")
				else
					FItemList(i).Fmainimageurl  = webImgSSLUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
				end if

				FItemList(i).Fgeneralbarcode    = db2html(rsget("generalbarcode"))
				FItemList(i).Flcitemname    	= db2html(rsget("lcitemname"))
				FItemList(i).Flcitemoptionname  = db2html(rsget("lcitemoptionname"))
				FItemList(i).Flcprice    		= db2html(rsget("lcprice"))
				FItemList(i).Fboxno		       	= rsget("boxno")

				i=i+1
				rsget.movenext
			loop
		end if
	rsget.close
	end sub

	'/admin/fran/viewordersheet.asp
    public function getShopOneOrderDetailList()
        dim sqlStr, i

		if (FRectAuthMode = "none") then
        	sqlStr = "exec db_shop.dbo.[sp_Ten_Shop_Order_Front_GetOneOrder_Detail_InSafe] '"&frectsitename&"','"&FRectOrderIDX&"'"
		else
			sqlStr = "exec db_shop.dbo.[sp_Ten_Shop_Order_Front_GetOneOrder_Detail] '"&FRectForeignOrderShopid&"','"&frectsitename&"','"&FRectOrderIDX&"'"
		end if

		'response.write sqlStr & "<br>"
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FResultCount = rsget.RecordCount
        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            do until rsget.EOF
                set FItemList(i) = new CStorageDetailItem

                FItemList(i).Fdetailidx         = rsget("idx")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fitemoption        = rsget("itemoption")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).fsocname         = rsget("socname")
                FItemList(i).fsocname_kor         = rsget("socname_kor")
                FItemList(i).Fitemname          = rsget("itemname")
                FItemList(i).Fitemoptionname    = rsget("itemoptionname")
                FItemList(i).Fbaljuitemno       = rsget("baljuitemno")
                FItemList(i).Frealitemno        = rsget("realitemno")
                FItemList(i).Fforeign_sellcash  = rsget("foreign_sellcash")
                FItemList(i).Fforeign_suplycash = rsget("foreign_suplycash")
                FItemList(i).FmLitemname        = rsget("mLitemname")
                FItemList(i).FmLitemOptionname  = rsget("mLitemOptionname")
                FItemList(i).FupchemanageCode   = rsget("upchemanageCode")
	           	FItemList(i).FImageList 		= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
	           	FItemList(i).FImageSmall 		= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	'/common/barcode/inc_barcodeprint_on.asp	'/common/barcode/inc_paperbarcodeprint_on.asp
	public Sub GetjumundetaillistByBox
		dim sqlStr,i, iCountrylangCd, sqlsearch

        if (FRectForeignOrderShopid<>"") then
            iCountrylangCd= GetShopCountrylangcd(FRectForeignOrderShopid)
        end if

		if frectitembarcodearr<>"" then
			frectitembarcodearr = replace(frectitembarcodearr, ",", "','")
			frectitembarcodearr = "'" & frectitembarcodearr & "'"
			sqlsearch = sqlsearch & " and [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun, d.itemid, d.itemoption) in ("& frectitembarcodearr &")"
		end if
		if frectbaljucode<>"" then
			sqlsearch = sqlsearch & " and m.baljucode='"& frectbaljucode &"'"
		end if
		if FRectMasterIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx="& FRectMasterIdx &""
		end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and d.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and d.itemid in (" + FRectItemid + ")"
            end if
        end if
        if (FRectMakerid <> "") then
            sqlsearch = sqlsearch & " and d.makerid='" + FRectMakerid + "'"
        end if
        ''//상품명 검색 수정  2016/04/04 최대 N건 가능 by eastone
        if (FRectItemName <> "") then
			''addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"

			sqlStr = " select top 1000 B.itemid into #TMPSearchItem"
			sqlStr = sqlStr + " from [DBAPPWISH].db_AppWish.dbo.tbl_item_SearchBase B with (nolock)"
			if (FRectMakerid <> "") then
    			sqlStr = sqlStr + " Join [DBAPPWISH].[db_AppWish].dbo.tbl_item ai with (nolock)"
            	sqlStr = sqlStr + " on B.itemid=ai.itemid"
            	sqlStr = sqlStr + " and ai.makerid='"&FRectMakerid&"'"
	        end if
	        sqlStr = sqlStr + " where contains(B.searchKey,'""" + CStr(FRectItemName) + """') "
            sqlStr = sqlStr + " order by B.itemid desc "
            dbget.Execute sqlStr

            sqlsearch = sqlsearch & " and d.itemid in (select itemid from #TMPSearchItem )"  ''2016/04/04
		end if
		if FRectPrdCode<>"" then
			if (Len(FRectPrdCode) = 12) then
				'sqlsearch = sqlsearch + " 	and d.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemid = " + RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			else
				'sqlsearch = sqlsearch + " 	and d.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemid = " + RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			end if
		end if
		if FRectGeneralBarcode<>"" then
			sqlsearch = sqlsearch + " 	and s.extbarcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if
        if (FRectSellYN="YS") then
            sqlsearch = sqlsearch & " and i.sellyn<>'N'"
        elseif( FRectSellYN="SR") then
        	  sqlsearch = sqlsearch & " and i.sellyn='N' and r.itemid is not null "
        elseif (FRectSellYN <> "") then
            sqlsearch = sqlsearch & " and i.sellyn='" + FRectSellYN + "'"
        end if
        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and i.isusing='" + FRectIsUsing + "'"
        end if
        if (FRectCDL<>"") then
            sqlsearch = sqlsearch + " 	and s.catecdl='" + FRectCDL + "'" & VbCrLf
        end if
        if (FRectCDM<>"") then
            sqlsearch = sqlsearch + " 	and s.catecdm='" + FRectCDM + "'" & VbCrLf
        end if
        if (FRectCDS<>"") then
            sqlsearch = sqlsearch + " 	and s.catecdn='" + FRectCDS + "'" & VbCrLf
        end if
		if FRectShopItemName<>"" then
			sqlsearch = sqlsearch + " 	and f.lcitemname like '%" + FRectShopItemName + "%'" & VbCrLf
		end if
		if FRectCurrentStockExist = "Y" then
			sqlsearch = sqlsearch + " 	and sc.shopid is not null " & VbCrLf
		end if
		if FRectRealStockOneMore = "Y" then
			sqlsearch = sqlsearch + " 	and IsNull(sc.realstockno, 0) > 0 " & VbCrLf
		end if
		if FRectShopItemNameInserted = "Y" then
			sqlsearch = sqlsearch + " 	and f.shopid is not null " & VbCrLf
		end if
		if FRectshopid <> "" then
			sqlsearch = sqlsearch + " 	and m.baljuid = '"& FRectshopid &"'" & VbCrLf
		end if
		if FRectitemgubun <> "" then
			sqlsearch = sqlsearch + " 	and d.itemgubun='"& FRectitemgubun &"'" & VbCrLf
		end if

		sqlStr = " select count(d.idx) as cnt "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m with (nolock)"
		sqlStr = sqlStr + " JOIN [db_storage].[dbo].tbl_ordersheet_detail d with (nolock)"
		sqlStr = sqlStr + " 	ON d.masteridx = m.idx "
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + " 	on d.itemgubun = '10' "
		sqlStr = sqlStr + " 	and d.itemid = i.itemid "
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s with (nolock)"
		sqlStr = sqlStr + " 	on d.itemgubun=s.itemgubun "
		sqlStr = sqlStr + " 	and d.itemid=s.shopitemid "
		sqlStr = sqlStr + " 	and d.itemoption=s.itemoption "
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock o with (nolock)"
		sqlStr = sqlStr + " 	on d.itemgubun=o.itemgubun "
		sqlStr = sqlStr + " 	and d.itemid=o.itemid "
		sqlStr = sqlStr + " 	and d.itemoption=o.itemoption "
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_locale_item f with (nolock)"
		sqlStr = sqlStr + " 	on f.shopid = '" & CStr(FRectForeignOrderShopid) & "' "
		sqlStr = sqlStr + " 	and d.itemgubun=f.itemgubun "
		sqlStr = sqlStr + " 	and d.itemid=f.shopitemid "
		sqlStr = sqlStr + " 	and d.itemoption=f.itemoption "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on d.makerid=p.id " + vbcrlf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c c with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on d.makerid=c.userid " & vbcrlf
        sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang_price] r with (nolock)"
        sqlStr = sqlStr & "  	on r.sitename = 'WSLWEB'"
        sqlStr = sqlStr & "  	and r.currencyUnit = 'USD'"
        sqlStr = sqlStr & "  	and d.itemid=r.itemid"
		sqlStr = sqlStr & " left join [db_summary].[dbo].tbl_current_shopstock_summary sc with (nolock)" & VbCrLf
		sqlStr = sqlStr & " 	on sc.shopid = '" & CStr(FRectForeignOrderShopid) & "' " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemgubun = d.itemgubun " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemid = d.itemid " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemoption = d.itemoption " & VbCrLf
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on d.itemid = a.itemid"	& vbcrlf
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 " & sqlsearch
		sqlStr = sqlStr + " 	and m.idx in ( "
		sqlStr = sqlStr + " 		select  "
		sqlStr = sqlStr + " 		 	m.idx "
		sqlStr = sqlStr + " 		 from  "
		sqlStr = sqlStr + " 		 	[db_storage].[dbo].tbl_shopbalju b with (nolock)"
		sqlStr = sqlStr + " 		 	JOIN [db_storage].[dbo].tbl_ordersheet_master m with (nolock)"
		sqlStr = sqlStr + " 		 	ON  "
		sqlStr = sqlStr + " 		 		b.baljucode = m.baljucode  "
		sqlStr = sqlStr + " 		 where  "
		sqlStr = sqlStr + " 			CONVERT(VARCHAR(10),b.baljudate,21) = ( " + vbcrlf
		sqlStr = sqlStr + " 				select " + vbcrlf
		sqlStr = sqlStr + " 					CONVERT(VARCHAR(10),b.baljudate,21) as yyyymmdd " + vbcrlf
		sqlStr = sqlStr + " 		 		from  "
		sqlStr = sqlStr + " 		 			[db_storage].[dbo].tbl_shopbalju b with (nolock)"
		sqlStr = sqlStr + " 		 			JOIN [db_storage].[dbo].tbl_ordersheet_master m with (nolock)"
		sqlStr = sqlStr + " 		 			ON  "
		sqlStr = sqlStr + " 		 				b.baljucode = m.baljucode  "
		sqlStr = sqlStr + " 				where " + vbcrlf
		sqlStr = sqlStr + " 					m.idx = " & CStr(FRectMasterIdx) & " " + vbcrlf
		sqlStr = sqlStr + " 			) " + vbcrlf
		sqlStr = sqlStr + " 			and m.baljuid = '" & CStr(FRectShopId) & "' " + vbcrlf
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and m.deldt is Null " + vbcrlf
		sqlStr = sqlStr + " 	and d.deldt is Null " + vbcrlf

		if (FRectBoxNo <> "") then
			sqlStr = sqlStr + " 	and isnull(d.packingstate,0) = " & CStr(FRectBoxNo) & " " + vbcrlf
		end if

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " " + vbcrlf
		sqlStr = sqlStr + " m.baljucode " + vbcrlf
		sqlStr = sqlStr + " , d.idx " + vbcrlf
		sqlStr = sqlStr + " , d.masteridx " + vbcrlf
		sqlStr = sqlStr + " , [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun,d.itemid,d.itemoption)  as prdcode " + vbcrlf
		sqlStr = sqlStr + " , [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun,d.itemid,d.itemoption)  as prdbarcode " + vbcrlf
		sqlStr = sqlStr + " , d.itemgubun " + vbcrlf
		sqlStr = sqlStr + " , d.itemid " + vbcrlf
		sqlStr = sqlStr + " , d.itemoption " + vbcrlf
		sqlStr = sqlStr + " , d.itemoptionname " + vbcrlf
		sqlStr = sqlStr + " , d.itemname as prdname " + vbcrlf
		sqlStr = sqlStr + " , d.makerid as locationid " + vbcrlf
		sqlStr = sqlStr + " , p.company_name as locationname, c.socname, c.socname_kor" + vbcrlf
		sqlStr = sqlStr + " , d.sellcash as customerprice, d.sellcash as sellprice, d.buycash as supplyprice" + vbcrlf
		sqlStr = sqlStr + " 	, d.baljuitemno as requestedno " + vbcrlf
		sqlStr = sqlStr + " , d.realitemno as fixedno " + vbcrlf
		sqlStr = sqlStr + " , (CASE " + vbcrlf
		sqlStr = sqlStr + " 		WHEN m.deldt is not null or d.deldt is not null THEN 'N' " + vbcrlf
		sqlStr = sqlStr + " 		ELSE 'Y' " + vbcrlf
		sqlStr = sqlStr + " 	END " + vbcrlf
		sqlStr = sqlStr + " ) as useyn " + vbcrlf
		sqlStr = sqlStr + " , o.barcode as generalbarcode " + vbcrlf
		sqlStr = sqlStr + " , i.smallimage, i.listimage" + vbcrlf
		sqlStr = sqlStr + " , s.offimgsmall, s.offimglist, s.extbarcode, s.itemcopy" + vbcrlf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + " , isnull(isNULL(Lni.itemname,f.lcitemname),d.itemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " , isnull(isNULL(Lno.optionname,f.lcitemoptionname),d.itemoptionname) as lcitemoptionname " + vbcrlf
			sqlStr = sqlStr & " , Lni.sourcearea as sourcearea_en, Lni.itemsource as itemsource_en, Lni.itemsize as itemsize_en" + vbcrlf
		else
    		sqlStr = sqlStr + " , isnull(f.lcitemname,d.itemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " , isnull(f.lcitemoptionname,d.itemoptionname) as lcitemoptionname" + vbcrlf
    		sqlStr = sqlStr & " , '' as sourcearea_en, '' as itemsource_en, '' as itemsize_en" + vbcrlf
	    end if

		sqlStr = sqlStr + " , isnull(f.lcprice,0) as lcprice" + vbcrlf
		sqlStr = sqlStr + " , isnull(d.packingstate,0) as boxno " + vbcrlf
		sqlStr = sqlStr & " , (CASE " & VbCrLf
		sqlStr = sqlStr & " 		WHEN IsNull(sc.realstockno, 0) <= 0 THEN 0 " & VbCrLf
		sqlStr = sqlStr & " 		ELSE sc.realstockno " & VbCrLf
		sqlStr = sqlStr & " 	END " & VbCrLf
		sqlStr = sqlStr & " ) as realstockno " & VbCrLf
		sqlStr = sqlStr & " , (CASE" & VbCrLf
		sqlStr = sqlStr & " 		WHEN IsNull(f.lcprice, 0) > 0 THEN 'Y'" & VbCrLf
		sqlStr = sqlStr & " 		ELSE 'N'" & VbCrLf
		sqlStr = sqlStr & " 	END" & VbCrLf
		sqlStr = sqlStr & " ) as saleyn" & VbCrLf
		sqlStr = sqlStr & " , c1.code_nm as catename1" & vbcrlf
		sqlStr = sqlStr & " , c2.code_nm as catename2, c2.code_nm_eng as catename_eng2, c2.code_nm_cn_gan as catename_cn_gan2, c2.code_nm_cn_bun as catename_cn_bun2" & vbcrlf
		sqlStr = sqlStr & " , c3.code_nm as catename3, c3.code_nm_eng as catename_eng3, c3.code_nm_cn_gan as catename_cn_gan3, c3.code_nm_cn_bun as catename_cn_bun3" & vbcrlf
		sqlStr = sqlStr & " , ic.sourcearea as sourcearea_10x10, ic.itemsource as itemsource_10x10, ic.itemsize as itemsize_10x10"
		sqlstr = sqlstr & " , isnull((case when d.itemgubun='10' then i.itemrackcode else s.offitemrackcode end),'') as itemrackcode"
		sqlstr = sqlstr & " , isnull(c.prtidx,'') as prtidx, isnull(a.subitemrackcode,'') as subitemrackcode" & vbcrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m with (nolock)" + vbcrlf
		sqlStr = sqlStr + " JOIN [db_storage].[dbo].tbl_ordersheet_detail d with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	ON d.masteridx = m.idx " + vbcrlf
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on d.itemgubun = '10' " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemid = i.itemid " + vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_Contents ic with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on i.itemid = ic.itemid" & vbcrlf
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on d.itemgubun=s.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemid=s.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemoption=s.itemoption " + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock o with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on d.itemgubun=o.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemid=o.itemid " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemoption=o.itemoption " + vbcrlf
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_locale_item f with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on f.shopid = '" & CStr(FRectForeignOrderShopid) & "' "
		sqlStr = sqlStr + " 	and d.itemgubun=f.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemid=f.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemoption=f.itemoption " + vbcrlf
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on d.makerid=p.id " + vbcrlf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c c with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on d.makerid=c.userid " & vbcrlf
        sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang_price] r with (nolock)"
        sqlStr = sqlStr & "  	on r.sitename = 'WSLWEB'"
        sqlStr = sqlStr & "  	and r.currencyUnit = 'USD'"
        sqlStr = sqlStr & "  	and d.itemid=r.itemid"
		sqlStr = sqlStr & " left join [db_summary].[dbo].tbl_current_shopstock_summary sc with (nolock)" & VbCrLf
		sqlStr = sqlStr & " 	on sc.shopid = '" & CStr(FRectForeignOrderShopid) & "' " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemgubun = d.itemgubun " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemid = d.itemid " & VbCrLf
		sqlStr = sqlStr & " 	and sc.itemoption = d.itemoption " & VbCrLf
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_large] as c1 with (nolock)"
		sqlStr = sqlStr & " 	on s.catecdl=c1.code_large"
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_mid] as c2 with (nolock)"
		sqlStr = sqlStr & " 	on s.catecdl=c2.code_large"
		sqlStr = sqlStr & " 	and s.catecdm=c2.code_mid"
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_small] as c3 with (nolock)"
		sqlStr = sqlStr & " 	on s.catecdl=c3.code_large"
		sqlStr = sqlStr & " 	and s.catecdm=c3.code_mid"
		sqlStr = sqlStr & " 	and s.catecdn=c3.code_small"
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on d.itemid = a.itemid"	& vbcrlf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang] Lni with (nolock)"
            sqlStr = sqlStr + "  	on Lni.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  	and d.itemgubun='10'"
            sqlStr = sqlStr + "  	and d.itemid=Lni.itemid"

            sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang_option] Lno with (nolock)"
            sqlStr = sqlStr + "  	on Lno.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  	and d.itemgubun='10'"
            sqlStr = sqlStr + "  	and d.itemid=Lno.itemid"
            sqlStr = sqlStr + "  	and d.itemoption=Lno.itemoption"
		end if
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 " & sqlsearch
		sqlStr = sqlStr + " 	and m.idx in ( "
		sqlStr = sqlStr + " 		select  "
		sqlStr = sqlStr + " 		 	m.idx "
		sqlStr = sqlStr + " 		 from  "
		sqlStr = sqlStr + " 		 	[db_storage].[dbo].tbl_shopbalju b with (nolock)"
		sqlStr = sqlStr + " 		 	JOIN [db_storage].[dbo].tbl_ordersheet_master m with (nolock)"
		sqlStr = sqlStr + " 		 	ON  "
		sqlStr = sqlStr + " 		 		b.baljucode = m.baljucode  "
		sqlStr = sqlStr + " 		 where  "
		sqlStr = sqlStr + " 			CONVERT(VARCHAR(10),b.baljudate,21) = ( " + vbcrlf
		sqlStr = sqlStr + " 				select " + vbcrlf
		sqlStr = sqlStr + " 					CONVERT(VARCHAR(10),b.baljudate,21) as yyyymmdd " + vbcrlf
		sqlStr = sqlStr + " 		 		from  "
		sqlStr = sqlStr + " 		 			[db_storage].[dbo].tbl_shopbalju b with (nolock)"
		sqlStr = sqlStr + " 		 			JOIN [db_storage].[dbo].tbl_ordersheet_master m with (nolock)"
		sqlStr = sqlStr + " 		 			ON  "
		sqlStr = sqlStr + " 		 				b.baljucode = m.baljucode  "
		sqlStr = sqlStr + " 				where " + vbcrlf
		sqlStr = sqlStr + " 					m.idx = " & CStr(FRectMasterIdx) & " " + vbcrlf
		sqlStr = sqlStr + " 			) " + vbcrlf
		sqlStr = sqlStr + " 			and m.baljuid = '" & CStr(FRectShopId) & "' " + vbcrlf
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and m.deldt is Null " + vbcrlf
		sqlStr = sqlStr + " 	and d.deldt is Null " + vbcrlf

		if (FRectBoxNo <> "") then
			sqlStr = sqlStr + " 	and isnull(d.packingstate,0) = " & CStr(FRectBoxNo) & " " + vbcrlf
		end if

		sqlStr = sqlStr + " order by m.baljucode, isnull(d.packingstate,0), d.itemgubun, d.itemid, d.itemoption " + vbcrlf
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
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
				set FItemList(i) = new CStorageDetailItem

				FItemList(i).Fprdcode    		= db2html(rsget("prdcode"))
				FItemList(i).Fprdname    		= db2html(rsget("prdname"))
				FItemList(i).Flocationid    	= db2html(rsget("locationid"))
				FItemList(i).Flocationname    	= db2html(rsget("locationname"))
				FItemList(i).fsocname    = db2html(rsget("socname"))
				FItemList(i).fsocname_kor    = db2html(rsget("socname_kor"))
				FItemList(i).Fitemgubun    		= db2html(rsget("itemgubun"))
				FItemList(i).Fitemid    		= db2html(rsget("itemid"))
				FItemList(i).Fitemoption    	= db2html(rsget("itemoption"))
				FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
				FItemList(i).fitemcopy    = db2html(rsget("itemcopy"))
				FItemList(i).Fprdbarcode    	= db2html(rsget("prdbarcode"))
				FItemList(i).Fgeneralbarcode    = db2html(rsget("generalbarcode"))
				FItemList(i).Fcustomerprice    	= rsget("customerprice")
				FItemList(i).Fsupplyprice    	= rsget("supplyprice")
				FItemList(i).Fsellprice       	= rsget("sellprice")

				if (IsNull(rsget("listimage")) = True) then
					FItemList(i).Fmainimageurl  = webImgSSLUrl & "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("offimglist")
				else
					FItemList(i).Fmainimageurl  = webImgSSLUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
				end if

				FItemList(i).Fuseyn    			= db2html(rsget("useyn"))
				FItemList(i).Flcitemname    	= db2html(rsget("lcitemname"))
				FItemList(i).Flcitemoptionname  = db2html(rsget("lcitemoptionname"))
				FItemList(i).Flcprice    		= db2html(rsget("lcprice"))
				FItemList(i).Frealstockno       = rsget("realstockno")
				FItemList(i).fsaleyn    		= rsget("saleyn")
				FItemList(i).Fbaljucode    		= rsget("baljucode")
				FItemList(i).fpublicbarcode    		= rsget("extbarcode")
				FItemList(i).Fidx        		= rsget("idx")
				FItemList(i).Fmasteridx       	= rsget("masteridx")
				FItemList(i).Frequestedno    	= db2html(rsget("requestedno"))
				FItemList(i).Ffixedno    		= db2html(rsget("fixedno"))
				FItemList(i).Fboxno		       	= rsget("boxno")
				FItemList(i).fcatename1    = db2html(rsget("catename1"))
				FItemList(i).fcatename2    = db2html(rsget("catename2"))
				FItemList(i).fcatename_cn_gan2    = db2html(rsget("catename_cn_gan2"))
				FItemList(i).fcatename_cn_bun2    = db2html(rsget("catename_cn_bun2"))
				FItemList(i).fcatename3    = db2html(rsget("catename3"))
				FItemList(i).fcatename_cn_gan3    = db2html(rsget("catename_cn_gan3"))
				FItemList(i).fcatename_cn_bun3    = db2html(rsget("catename_cn_bun3"))
				FItemList(i).fcatename_eng2    = db2html(rsget("catename_eng2"))
				FItemList(i).fcatename_eng3    = db2html(rsget("catename_eng3"))
				FItemList(i).fsourcearea_10x10    		= db2html(rsget("sourcearea_10x10"))
				FItemList(i).fitemsource_10x10    		= db2html(rsget("itemsource_10x10"))
				FItemList(i).fitemsize_10x10    		= db2html(rsget("itemsize_10x10"))
				FItemList(i).fsourcearea_en    		= db2html(rsget("sourcearea_en"))
				FItemList(i).fitemsource_en    		= db2html(rsget("itemsource_en"))
				FItemList(i).fitemsize_en    		= db2html(rsget("itemsize_en"))
				FItemList(i).Fitemrackcode = rsget("itemrackcode")
				FItemList(i).fprtidx = rsget("prtidx")
				FItemList(i).fsubitemrackcode = rsget("subitemrackcode")

				i=i+1
				rsget.movenext
			loop
		end if
	rsget.close
	end sub

	public Sub GetStorageDetailListByBox
		dim sqlStr,i, iCountrylangCd, sqlsearch

        if (FRectForeignOrderShopid<>"") then
            iCountrylangCd= GetShopCountrylangcd(FRectForeignOrderShopid)
        end if

		if frectitembarcodearr<>"" then
			frectitembarcodearr = replace(frectitembarcodearr, ",", "','")
			frectitembarcodearr = "'" & frectitembarcodearr & "'"
			sqlsearch = sqlsearch & " and [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun, d.itemid, d.itemoption) in ("& frectitembarcodearr &")"
		end if
		if frectbaljucode<>"" then
			sqlsearch = sqlsearch & " and m.baljucode='"& frectbaljucode &"'"
		end if
		if FRectMasterIdx<>"" then
			sqlsearch = sqlsearch & " and m.idx="& FRectMasterIdx &""
		end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and d.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	sqlsearch = sqlsearch & " and d.itemid in (" + FRectItemid + ")"
            end if
        end if
        if (FRectMakerid <> "") then
            sqlsearch = sqlsearch & " and d.makerid='" + FRectMakerid + "'"
        end if
        ''//상품명 검색 수정  2016/04/04 최대 N건 가능 by eastone
        if (FRectItemName <> "") then
			''addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"

			sqlStr = " select top 1000 B.itemid into #TMPSearchItem"
			sqlStr = sqlStr + " from [DBAPPWISH].db_AppWish.dbo.tbl_item_SearchBase B"
			if (FRectMakerid <> "") then
    			sqlStr = sqlStr + " Join [DBAPPWISH].[db_AppWish].dbo.tbl_item ai"
            	sqlStr = sqlStr + " on B.itemid=ai.itemid"
            	sqlStr = sqlStr + " and ai.makerid='"&FRectMakerid&"'"
	        end if
	        sqlStr = sqlStr + " where contains(B.searchKey,'""" + CStr(FRectItemName) + """') "
            sqlStr = sqlStr + " order by B.itemid desc "
            dbget.Execute sqlStr

            sqlsearch = sqlsearch & " and d.itemid in (select itemid from #TMPSearchItem )"  ''2016/04/04
		end if
		if FRectPrdCode<>"" then
			if (Len(FRectPrdCode) = 12) then
				'sqlsearch = sqlsearch + " 	and d.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemid = " + RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			else
				'sqlsearch = sqlsearch + " 	and d.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemid = " + RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and d.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			end if
		end if
		if FRectGeneralBarcode<>"" then
			sqlsearch = sqlsearch + " 	and o.barcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if
        if (FRectSellYN="YS") then
            sqlsearch = sqlsearch & " and i.sellyn<>'N'"
        elseif( FRectSellYN="SR") then
        	  sqlsearch = sqlsearch & " and i.sellyn='N' and r.itemid is not null "
        elseif (FRectSellYN <> "") then
            sqlsearch = sqlsearch & " and i.sellyn='" + FRectSellYN + "'"
        end if
        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and i.isusing='" + FRectIsUsing + "'"
        end if

		sqlStr = " select count(d.idx) as cnt "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m "
		sqlStr = sqlStr + " JOIN [db_storage].[dbo].tbl_ordersheet_detail d "
		sqlStr = sqlStr + " 	ON d.masteridx = m.idx "
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i "
		sqlStr = sqlStr + " 	on d.itemgubun = '10' "
		sqlStr = sqlStr + " 	and d.itemid = i.itemid "
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " 	on d.itemgubun=s.itemgubun "
		sqlStr = sqlStr + " 	and d.itemid=s.shopitemid "
		sqlStr = sqlStr + " 	and d.itemoption=s.itemoption "
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock o "
		sqlStr = sqlStr + " 	on d.itemgubun=o.itemgubun "
		sqlStr = sqlStr + " 	and d.itemid=o.itemid "
		sqlStr = sqlStr + " 	and d.itemoption=o.itemoption "
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_locale_item f "
		sqlStr = sqlStr + " 	on f.shopid = '" & CStr(FRectForeignOrderShopid) & "' "
		sqlStr = sqlStr + " 	and d.itemgubun=f.itemgubun "
		sqlStr = sqlStr + " 	and d.itemid=f.shopitemid "
		sqlStr = sqlStr + " 	and d.itemoption=f.itemoption "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p " + vbcrlf
		sqlStr = sqlStr + " 	on d.makerid=p.id " + vbcrlf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c c" & vbcrlf
		sqlStr = sqlStr & " 	on d.makerid=c.userid " & vbcrlf
        sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang_price] r"
        sqlStr = sqlStr & "  	on r.sitename = 'WSLWEB'"
        sqlStr = sqlStr & "  	and r.currencyUnit = 'USD'"
        sqlStr = sqlStr & "  	and d.itemid=r.itemid"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 " & sqlsearch
		sqlStr = sqlStr + " 	and m.idx in ( "
		sqlStr = sqlStr + " 		select  "
		sqlStr = sqlStr + " 		 	m.idx "
		sqlStr = sqlStr + " 		 from  "
		sqlStr = sqlStr + " 		 	[db_storage].[dbo].tbl_shopbalju b  "
		sqlStr = sqlStr + " 		 	JOIN [db_storage].[dbo].tbl_ordersheet_master m  "
		sqlStr = sqlStr + " 		 	ON  "
		sqlStr = sqlStr + " 		 		b.baljucode = m.baljucode  "
		sqlStr = sqlStr + " 		 where  "
		sqlStr = sqlStr + " 			CONVERT(VARCHAR(10),b.baljudate,21) = ( " + vbcrlf
		sqlStr = sqlStr + " 				select " + vbcrlf
		sqlStr = sqlStr + " 					CONVERT(VARCHAR(10),b.baljudate,21) as yyyymmdd " + vbcrlf
		sqlStr = sqlStr + " 		 		from  "
		sqlStr = sqlStr + " 		 			[db_storage].[dbo].tbl_shopbalju b  "
		sqlStr = sqlStr + " 		 			JOIN [db_storage].[dbo].tbl_ordersheet_master m  "
		sqlStr = sqlStr + " 		 			ON  "
		sqlStr = sqlStr + " 		 				b.baljucode = m.baljucode  "
		sqlStr = sqlStr + " 				where " + vbcrlf
		sqlStr = sqlStr + " 					m.idx = " & CStr(FRectMasterIdx) & " " + vbcrlf
		sqlStr = sqlStr + " 			) " + vbcrlf
		sqlStr = sqlStr + " 			and m.baljuid = '" & CStr(FRectShopId) & "' " + vbcrlf
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and m.deldt is Null " + vbcrlf
		sqlStr = sqlStr + " 	and d.deldt is Null " + vbcrlf

		if (FRectBoxNo <> "") then
			sqlStr = sqlStr + " 	and isnull(d.packingstate,0) = " & CStr(FRectBoxNo) & " " + vbcrlf
		end if

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " " + vbcrlf
		sqlStr = sqlStr + " m.baljucode " + vbcrlf
		sqlStr = sqlStr + " , d.idx " + vbcrlf
		sqlStr = sqlStr + " , d.masteridx " + vbcrlf
		sqlStr = sqlStr + " , [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun,d.itemid,d.itemoption)  as prdcode " + vbcrlf
		sqlStr = sqlStr + " , [db_storage].[dbo].[uf_getTenBarCodeType](d.itemgubun,d.itemid,d.itemoption)  as prdbarcode " + vbcrlf
		sqlStr = sqlStr + " , d.itemgubun " + vbcrlf
		sqlStr = sqlStr + " , d.itemid " + vbcrlf
		sqlStr = sqlStr + " , d.itemoption " + vbcrlf
		sqlStr = sqlStr + " , d.itemoptionname " + vbcrlf
		sqlStr = sqlStr + " , d.itemname as prdname " + vbcrlf
		sqlStr = sqlStr + " , d.makerid as locationid " + vbcrlf
		sqlStr = sqlStr + " , p.company_name as locationname, c.socname, c.socname_kor" + vbcrlf
		sqlStr = sqlStr + " , d.sellcash as customerprice, d.sellcash as sellprice, d.buycash as supplyprice" + vbcrlf
		sqlStr = sqlStr + " , d.realitemno as fixedno " + vbcrlf
		sqlStr = sqlStr + " , (CASE " + vbcrlf
		sqlStr = sqlStr + " 		WHEN m.deldt is not null or d.deldt is not null THEN 'N' " + vbcrlf
		sqlStr = sqlStr + " 		ELSE 'Y' " + vbcrlf
		sqlStr = sqlStr + " 	END " + vbcrlf
		sqlStr = sqlStr + " ) as useyn " + vbcrlf
		sqlStr = sqlStr + " , o.barcode as generalbarcode " + vbcrlf
		sqlStr = sqlStr + " , i.smallimage, i.listimage" + vbcrlf
		sqlStr = sqlStr + " , s.offimgsmall, s.offimglist" + vbcrlf

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + " 	, isnull(isNULL(Lni.itemname,f.lcitemname),d.itemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " 	, isnull(isNULL(Lno.optionname,f.lcitemoptionname),d.itemoptionname) as lcitemoptionname " + vbcrlf
		else
    		sqlStr = sqlStr + " 	, isnull(f.lcitemname,d.itemname) as lcitemname" + vbcrlf
    		sqlStr = sqlStr + " 	, isnull(f.lcitemoptionname,d.itemoptionname) as lcitemoptionname" + vbcrlf
	    end if

		sqlStr = sqlStr + " , isnull(f.lcprice,0) as lcprice" + vbcrlf
		sqlStr = sqlStr + " , isnull(d.packingstate,0) as boxno " + vbcrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m " + vbcrlf
		sqlStr = sqlStr + " JOIN [db_storage].[dbo].tbl_ordersheet_detail d " + vbcrlf
		sqlStr = sqlStr + " 	ON d.masteridx = m.idx " + vbcrlf
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item i " + vbcrlf
		sqlStr = sqlStr + " 	on d.itemgubun = '10' " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemid = i.itemid " + vbcrlf
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s " + vbcrlf
		sqlStr = sqlStr + " 	on d.itemgubun=s.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemid=s.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemoption=s.itemoption " + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock o " + vbcrlf
		sqlStr = sqlStr + " 	on d.itemgubun=o.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemid=o.itemid " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemoption=o.itemoption " + vbcrlf
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_locale_item f " + vbcrlf
		sqlStr = sqlStr + " 	on f.shopid = '" & CStr(FRectForeignOrderShopid) & "' "
		sqlStr = sqlStr + " 	and d.itemgubun=f.itemgubun " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemid=f.shopitemid " + vbcrlf
		sqlStr = sqlStr + " 	and d.itemoption=f.itemoption " + vbcrlf
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p " + vbcrlf
		sqlStr = sqlStr + " 	on d.makerid=p.id " + vbcrlf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c c" & vbcrlf
		sqlStr = sqlStr & " 	on d.makerid=c.userid " & vbcrlf
        sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_multiLang_price] r"
        sqlStr = sqlStr & "  	on r.sitename = 'WSLWEB'"
        sqlStr = sqlStr & "  	and r.currencyUnit = 'USD'"
        sqlStr = sqlStr & "  	and d.itemid=r.itemid"

		if (iCountrylangCd<>"") and (iCountrylangCd<>"KR") then
		    sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang] Lni"
            sqlStr = sqlStr + "  	on Lni.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  	and d.itemgubun='10'"
            sqlStr = sqlStr + "  	and d.itemid=Lni.itemid"

            sqlStr = sqlStr + "  left join db_item.[dbo].[tbl_item_multiLang_option] Lno"
            sqlStr = sqlStr + "  	on Lno.countryCd='"&iCountrylangCd&"'"
            sqlStr = sqlStr + "  	and d.itemgubun='10'"
            sqlStr = sqlStr + "  	and d.itemid=Lno.itemid"
            sqlStr = sqlStr + "  	and d.itemoption=Lno.itemoption"
		end if
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 " & sqlsearch
		sqlStr = sqlStr + " 	and m.idx in ( "
		sqlStr = sqlStr + " 		select  "
		sqlStr = sqlStr + " 		 	m.idx "
		sqlStr = sqlStr + " 		 from  "
		sqlStr = sqlStr + " 		 	[db_storage].[dbo].tbl_shopbalju b  "
		sqlStr = sqlStr + " 		 	JOIN [db_storage].[dbo].tbl_ordersheet_master m  "
		sqlStr = sqlStr + " 		 	ON  "
		sqlStr = sqlStr + " 		 		b.baljucode = m.baljucode  "
		sqlStr = sqlStr + " 		 where  "
		sqlStr = sqlStr + " 			CONVERT(VARCHAR(10),b.baljudate,21) = ( " + vbcrlf
		sqlStr = sqlStr + " 				select " + vbcrlf
		sqlStr = sqlStr + " 					CONVERT(VARCHAR(10),b.baljudate,21) as yyyymmdd " + vbcrlf
		sqlStr = sqlStr + " 		 		from  "
		sqlStr = sqlStr + " 		 			[db_storage].[dbo].tbl_shopbalju b  "
		sqlStr = sqlStr + " 		 			JOIN [db_storage].[dbo].tbl_ordersheet_master m  "
		sqlStr = sqlStr + " 		 			ON  "
		sqlStr = sqlStr + " 		 				b.baljucode = m.baljucode  "
		sqlStr = sqlStr + " 				where " + vbcrlf
		sqlStr = sqlStr + " 					m.idx = " & CStr(FRectMasterIdx) & " " + vbcrlf
		sqlStr = sqlStr + " 			) " + vbcrlf
		sqlStr = sqlStr + " 			and m.baljuid = '" & CStr(FRectShopId) & "' " + vbcrlf
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and m.deldt is Null " + vbcrlf
		sqlStr = sqlStr + " 	and d.deldt is Null " + vbcrlf

		if (FRectBoxNo <> "") then
			sqlStr = sqlStr + " 	and isnull(d.packingstate,0) = " & CStr(FRectBoxNo) & " " + vbcrlf
		end if

		sqlStr = sqlStr + " order by m.baljucode, isnull(d.packingstate,0), d.itemgubun, d.itemid, d.itemoption " + vbcrlf
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
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
				set FItemList(i) = new CStorageDetailItem
				FItemList(i).Fbaljucode    		= rsget("baljucode")
				FItemList(i).Fidx        		= rsget("idx")
				FItemList(i).Fmasteridx       	= rsget("masteridx")
				FItemList(i).Fprdcode    		= db2html(rsget("prdcode"))
				FItemList(i).Fprdbarcode    	= db2html(rsget("prdbarcode"))
				FItemList(i).Fitemgubun    		= db2html(rsget("itemgubun"))
				FItemList(i).Fitemid    		= db2html(rsget("itemid"))
				FItemList(i).Fitemoption    	= db2html(rsget("itemoption"))
				FItemList(i).Fitemoptionname    = db2html(rsget("itemoptionname"))
				FItemList(i).Fprdname    		= db2html(rsget("prdname"))
				FItemList(i).Flocationid    	= db2html(rsget("locationid"))
				FItemList(i).Flocationname    	= db2html(rsget("locationname"))
				FItemList(i).fsocname    = db2html(rsget("socname"))
				FItemList(i).fsocname_kor    = db2html(rsget("socname_kor"))
				FItemList(i).Fcustomerprice    	= rsget("customerprice")
				FItemList(i).Fsellprice       	= rsget("sellprice")
				FItemList(i).Fsupplyprice    	= rsget("supplyprice")
				FItemList(i).Ffixedno    		= db2html(rsget("fixedno"))
				FItemList(i).Fuseyn    			= db2html(rsget("useyn"))

				if (IsNull(rsget("listimage")) = True) then
					FItemList(i).Fmainimageurl  = webImgSSLUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("offimglist")
				else
					FItemList(i).Fmainimageurl  = webImgSSLUrl + "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
				end if

				FItemList(i).Fgeneralbarcode    = db2html(rsget("generalbarcode"))
				FItemList(i).Flcitemname    	= db2html(rsget("lcitemname"))
				FItemList(i).Flcitemoptionname  = db2html(rsget("lcitemoptionname"))
				FItemList(i).Flcprice    		= db2html(rsget("lcprice"))
				FItemList(i).Fboxno		       	= rsget("boxno")
				i=i+1
				rsget.movenext
			loop
		end if
	rsget.close
	end sub

	'/온라인 상품 리스트	'/2016.12.16 한용민 생성
	'/common/barcode/inc_barcodeprint_on.asp	'/common/barcode/inc_paperbarcodeprint_on.asp
	public Sub GetProductListOnline
		dim sqlStr, i, sqlsearch

		if frectitembarcodearr<>"" then
			frectitembarcodearr = replace(frectitembarcodearr, ",", "','")
			frectitembarcodearr = "'" & frectitembarcodearr & "'"
			sqlsearch = sqlsearch & " and [db_storage].[dbo].[uf_getTenBarCodeType]('10', i.itemid, isnull(o.itemoption,'0000')) in ("& frectitembarcodearr &")"
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
        if (FRectMakerid <> "") then
            sqlsearch = sqlsearch & " and i.makerid='" + FRectMakerid + "'"
        end if
        ''//상품명 검색 수정  2016/04/04 최대 N건 가능 by eastone
        if (FRectItemName <> "") then
			''addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"

			sqlStr = " select top 1000 B.itemid into #TMPSearchItem"
			sqlStr = sqlStr + " from [DBAPPWISH].db_AppWish.dbo.tbl_item_SearchBase B with (nolock)"
			if (FRectMakerid <> "") then
    			sqlStr = sqlStr + " Join [DBAPPWISH].[db_AppWish].dbo.tbl_item ai with (nolock)"
            	sqlStr = sqlStr + " 	on B.itemid=ai.itemid"
            	sqlStr = sqlStr + " 	and ai.makerid='"&FRectMakerid&"'"
	        end if
	        sqlStr = sqlStr + " where contains(B.searchKey,'""" + CStr(FRectItemName) + """') "
            sqlStr = sqlStr + " order by B.itemid desc "
            dbget.Execute sqlStr

            sqlsearch = sqlsearch & " and i.itemid in (select itemid from #TMPSearchItem )"  ''2016/04/04
		end if
		if FRectPrdCode<>"" then
			if (Len(FRectPrdCode) = 12) then
				'sqlsearch = sqlsearch + " 	and i.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and i.itemid = " + RIGHT(LEFT(CStr(FRectPrdCode), 8), 6) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and o.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			else
				'sqlsearch = sqlsearch + " 	and i.itemgubun = '" + LEFT(CStr(FRectPrdCode), 2) + "' " & VbCrLf
				sqlsearch = sqlsearch + " 	and i.itemid = " + RIGHT(LEFT(CStr(FRectPrdCode), (Len(FRectPrdCode) - 4)), (Len(FRectPrdCode) - 6)) + " " & VbCrLf
				sqlsearch = sqlsearch + " 	and o.itemoption = '" + RIGHT(CStr(FRectPrdCode), 4) + "' " & VbCrLf
			end if
		end if
		if FRectGeneralBarcode<>"" then
			sqlsearch = sqlsearch + " 	and s.barcode = '" + CStr(FRectGeneralBarcode) + "'" + VbCrlf
		end if
        if (FRectSellYN="YS") then
            sqlsearch = sqlsearch & " and i.sellyn<>'N'"
        elseif( FRectSellYN="SR") then
        	  sqlsearch = sqlsearch & " and i.sellyn='N' and r.itemid is not null "
        elseif (FRectSellYN <> "") then
            sqlsearch = sqlsearch & " and i.sellyn='" + FRectSellYN + "'"
        end if
        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and i.isusing='" + FRectIsUsing + "'"
        end if

		sqlStr = "select count(i.itemid) as cnt" + vbcrlf
		sqlStr = sqlStr + " from db_item.dbo.tbl_item i with (nolock)" + vbcrlf
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item_option o with (nolock)" & vbcrlf
		sqlStr = sqlStr + " 	on i.itemid = o.itemid" & vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on s.itemgubun='10'" + vbcrlf
		sqlStr = sqlStr + " 	and i.itemid=s.itemid " + vbcrlf
		sqlStr = sqlStr + " 	and isnull(o.itemoption,'0000')=s.itemoption " + vbcrlf
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on i.makerid=p.id " + vbcrlf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c c with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on i.makerid=c.userid " & vbcrlf
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on i.itemid = a.itemid"	& vbcrlf

		if FRectisforeignprint="Y" then
		    sqlStr = sqlStr + " left join db_item.[dbo].[tbl_item_multiLang] Lni with (nolock)"
	        sqlStr = sqlStr + "  	on Lni.countryCd='EN'"
	        sqlStr = sqlStr + "  	and i.itemid=Lni.itemid"
	        sqlStr = sqlStr + " left join db_item.[dbo].[tbl_item_multiLang_option] Lno with (nolock)"
	        sqlStr = sqlStr + "  	on Lno.countryCd='EN'"
	        sqlStr = sqlStr + "  	and i.itemid=Lno.itemid"
	        sqlStr = sqlStr + "  	and o.itemoption=Lno.itemoption"
	        sqlStr = sqlStr + " left join db_item.[dbo].[tbl_item_multiLang_price] r with (nolock)"
	        sqlStr = sqlStr + "  	on r.sitename = 'WSLWEB'"
	        sqlStr = sqlStr + "  	and r.currencyUnit = 'USD'"
	        sqlStr = sqlStr + "  	and i.itemid=r.itemid"
	    end if

		sqlStr = sqlStr + " where i.itemid<>0 " & sqlsearch

		'response.write sqlStr & "<Br>"
		'response.end
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " '10' as itemgubun " + vbcrlf
		sqlStr = sqlStr + " , i.itemid " + vbcrlf
		sqlStr = sqlStr + " , isnull(o.itemoption,'0000') as itemoption " + vbcrlf
		sqlStr = sqlStr + " , o.optionname " + vbcrlf
		sqlStr = sqlStr + " , i.itemname as prdname " + vbcrlf
		sqlStr = sqlStr + " , i.makerid as locationid " + vbcrlf
		sqlStr = sqlStr + " , p.company_name as locationname, c.socname, c.socname_kor" + vbcrlf
		sqlStr = sqlStr + " , i.orgprice as customerprice, i.sellcash as sellprice, i.orgsuplycash as supplyprice, i.sailprice" + vbcrlf
		sqlStr = sqlStr + " , i.itemcouponvalue, i.itemcoupontype, i.ItemCouponYN, i.sailyn, '1' as fixedno " + vbcrlf
		sqlStr = sqlStr + " , i.isusing as useyn " + vbcrlf
		sqlStr = sqlStr + " , s.barcode as generalbarcode " + vbcrlf
		sqlStr = sqlStr + " , i.smallimage, i.listimage" + vbcrlf
		sqlStr = sqlStr & " , c1.code_nm as catename1" & vbcrlf
		sqlStr = sqlStr & " , c2.code_nm as catename2, c2.code_nm_eng as catename_eng2, c2.code_nm_cn_gan as catename_cn_gan2, c2.code_nm_cn_bun as catename_cn_bun2" & vbcrlf
		sqlStr = sqlStr & " , c3.code_nm as catename3, c3.code_nm_eng as catename_eng3, c3.code_nm_cn_gan as catename_cn_gan3, c3.code_nm_cn_bun as catename_cn_bun3" & vbcrlf

		if FRectisforeignprint="Y" then
			sqlStr = sqlStr + " , Lni.itemname as lcitemname, Lno.optionname as lcitemoptionname, isnull(r.orgprice,0) as orgprice" + vbcrlf
	    end if

        sqlstr = sqlstr & " , isnull(i.itemrackcode,'') as itemrackcode, c.prtidx, isnull(a.subitemrackcode,'') as subitemrackcode" & vbcrlf
		sqlStr = sqlStr + " from db_item.dbo.tbl_item i with (nolock)" + vbcrlf
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item_option o with (nolock)" & vbcrlf
		sqlStr = sqlStr + " 	on i.itemid = o.itemid" & vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on s.itemgubun='10'" + vbcrlf
		sqlStr = sqlStr + " 	and i.itemid=s.itemid " + vbcrlf
		sqlStr = sqlStr + " 	and isnull(o.itemoption,'0000')=s.itemoption " + vbcrlf
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p with (nolock)" + vbcrlf
		sqlStr = sqlStr + " 	on i.makerid=p.id " + vbcrlf
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c c with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on i.makerid=c.userid " & vbcrlf
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_large] as c1 with (nolock)"
		sqlStr = sqlStr & " 	on i.cate_large=c1.code_large"
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_mid] as c2 with (nolock)"
		sqlStr = sqlStr & " 	on i.cate_large=c2.code_large"
		sqlStr = sqlStr & " 	and i.cate_mid=c2.code_mid"
		sqlStr = sqlStr & " left JOIN [db_item].[dbo].[tbl_Cate_small] as c3 with (nolock)"
		sqlStr = sqlStr & " 	on i.cate_large=c3.code_large"
		sqlStr = sqlStr & " 	and i.cate_mid=c3.code_mid"
		sqlStr = sqlStr & " 	and i.cate_small=c3.code_small"
        sqlstr = sqlstr & " left join [db_item].[dbo].[tbl_item_logics_addinfo] a with (nolock)"	& vbcrlf
        sqlstr = sqlstr & " 	on i.itemid = a.itemid"	& vbcrlf

		if FRectisforeignprint="Y" then
		    sqlStr = sqlStr + " left join db_item.[dbo].[tbl_item_multiLang] Lni with (nolock)"
	        sqlStr = sqlStr + "  	on Lni.countryCd='EN'"
	        sqlStr = sqlStr + "  	and i.itemid=Lni.itemid"
	        sqlStr = sqlStr + " left join db_item.[dbo].[tbl_item_multiLang_option] Lno with (nolock)"
	        sqlStr = sqlStr + "  	on Lno.countryCd='EN'"
	        sqlStr = sqlStr + "  	and i.itemid=Lno.itemid"
	        sqlStr = sqlStr + "  	and o.itemoption=Lno.itemoption"
	        sqlStr = sqlStr + " left join db_item.[dbo].[tbl_item_multiLang_price] r with (nolock)"
	        sqlStr = sqlStr + "  	on r.sitename = 'WSLWEB'"
	        sqlStr = sqlStr + "  	and r.currencyUnit = 'USD'"
	        sqlStr = sqlStr + "  	and i.itemid=r.itemid"
	    end if

		sqlStr = sqlStr + " where i.itemid<>0 " & sqlsearch
		sqlStr = sqlStr + " order by i.itemid desc, isnull(o.itemoption,'0000') asc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new CStorageDetailItem
				FItemList(i).Fitemgubun    		= rsget("itemgubun")
				FItemList(i).Fitemid    		= rsget("itemid")
				FItemList(i).Fshopitemid		= rsget("itemid")
				FItemList(i).Fitemoption    	= rsget("itemoption")
				FItemList(i).Fitemoptionname    = db2html(rsget("optionname"))
				FItemList(i).Fshopitemoptionname = db2html(rsget("optionname"))
				FItemList(i).Fprdname    		= db2html(rsget("prdname"))
				FItemList(i).Fshopitemname		= db2html(rsget("prdname"))
				FItemList(i).Flocationid    	= db2html(rsget("locationid"))
				FItemList(i).fmakerid    	= db2html(rsget("locationid"))
				FItemList(i).Flocationname    	= db2html(rsget("locationname"))
				FItemList(i).fsocname    = db2html(rsget("socname"))
				FItemList(i).fsocname_kor    = db2html(rsget("socname_kor"))
				FItemList(i).Fcustomerprice    	= rsget("customerprice")
				FItemList(i).Fsellprice       	= rsget("sellprice")
				FItemList(i).Fshopitemprice       	= rsget("sellprice")
				FItemList(i).Fsupplyprice    	= rsget("supplyprice")
				FItemList(i).FItemCouponYN    	= rsget("ItemCouponYN")
				FItemList(i).Fitemcoupontype    	= rsget("itemcoupontype")
				FItemList(i).Fitemcouponvalue    	= rsget("itemcouponvalue")
				FItemList(i).Fsaleyn    	= rsget("sailyn")
				FItemList(i).Fsaleprice    	= rsget("sailprice")
				FItemList(i).Ffixedno    		= rsget("fixedno")
				FItemList(i).Fuseyn    			= rsget("useyn")
				FItemList(i).Fmainimageurl  = webImgSSLUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
				FItemList(i).FImageSmall  = webImgSSLUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).Fgeneralbarcode    = db2html(rsget("generalbarcode"))
				FItemList(i).fcatename1    = db2html(rsget("catename1"))
				FItemList(i).fcatename2    = db2html(rsget("catename2"))
				FItemList(i).fcatename_cn_gan2    = db2html(rsget("catename_cn_gan2"))
				FItemList(i).fcatename_cn_bun2    = db2html(rsget("catename_cn_bun2"))
				FItemList(i).fcatename3    = db2html(rsget("catename3"))
				FItemList(i).fcatename_cn_gan3    = db2html(rsget("catename_cn_gan3"))
				FItemList(i).fcatename_cn_bun3    = db2html(rsget("catename_cn_bun3"))
				FItemList(i).fcatename_eng2    = db2html(rsget("catename_eng2"))
				FItemList(i).fcatename_eng3    = db2html(rsget("catename_eng3"))

				if FRectisforeignprint="Y" then
					FItemList(i).Flcitemname    	= db2html(rsget("lcitemname"))
					FItemList(i).Flcitemoptionname  = db2html(rsget("lcitemoptionname"))
					FItemList(i).Flcprice    		= rsget("orgprice")
				end if

				FItemList(i).Fitemrackcode = rsget("itemrackcode")
				FItemList(i).fprtidx = rsget("prtidx")
				FItemList(i).fsubitemrackcode = rsget("subitemrackcode")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 50
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
%>
