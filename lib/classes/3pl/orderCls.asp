<%
'###########################################################
' Description : 주문 클래스
' Hieditor : 2009.04.17 이상구 생성
'			 2010.01.03 한용민 수정
'###########################################################

function TicketOrderCheck(iorderserial,byRef mayTicketCancelChargePro,byRef ticketCancelDisabled,byRef ticketCancelStr)
    Dim sqlStr, D9Day, D6Day, D2Day, DDay, returnExpiredate
    Dim nowDate, R8Day

    mayTicketCancelChargePro = 0
    ticketCancelDisabled     = false

    sqlStr = " select top 1 "
    sqlStr = sqlStr & "  dateadd(d,-9,tk_StSchedule) as D9"
    sqlStr = sqlStr & " ,dateadd(d,-6,tk_StSchedule) as D6"
    sqlStr = sqlStr & " ,dateadd(d,-2,tk_StSchedule) as D2"
    sqlStr = sqlStr & " ,tk_StSchedule as Dday"
    sqlStr = sqlStr & " ,tk_EdSchedule"
    sqlStr = sqlStr & " ,returnExpiredate"
    sqlStr = sqlStr & " ,getdate() as nowDate"
	sqlStr = sqlStr & " ,dateadd(d,8,m.regDate) as R8"
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m "
	sqlStr = sqlStr & " 	join db_order.dbo.tbl_order_detail d "
	sqlStr = sqlStr & " 	on m.orderserial = d.orderserial "
    sqlStr = sqlStr & "	    Join db_item.dbo.tbl_ticket_Schedule s"
    sqlStr = sqlStr & "	    on d.itemid=s.tk_itemid"
    sqlStr = sqlStr & "	    and d.itemoption=s.tk_itemoption"
    sqlStr = sqlStr & " where d.orderserial='"&iorderserial&"'"
    sqlStr = sqlStr & " and d.itemid<>0"
    sqlStr = sqlStr & " and d.cancelyn<>'Y'"
	''rw sqlStr

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
		D9Day               = rsget("D9")
		D6Day               = rsget("D6")
		D2Day               = rsget("D2")
		DDay                = rsget("Dday")
		returnExpiredate    = rsget("returnExpiredate")
		nowDate             = rsget("nowDate")
		R8Day               = rsget("R8")			'// 예매일+8일
    end if
	rsget.close

    if (returnExpiredate="") then Exit function

    ' if (nowDate<D10Day) then
    '     exit function
    ' end If

    if (nowDate>returnExpiredate) then
        ticketCancelDisabled = true
        ticketCancelStr      = "취소 마감기간은 "&CStr(returnExpiredate)&"입니다."
        Exit function
    end If

    if (nowDate<D9Day) and (nowDate=>R8Day) Then
		'//예매 후 8일~관람일 10일전까지, 장당 2,000원(티켓금액의 10%한도)
        mayTicketCancelChargePro = 2000
        ticketCancelStr = "예매 후 8일~관람일 10일전 취소시 (관람일 : "&CStr(Dday)&") "
        Exit function
    end if

    if (nowDate>D9Day) and (nowDate=<D6Day) then
        mayTicketCancelChargePro = 10
        ticketCancelStr = "관람일 9일~7일전 취소시 (관람일 : "&CStr(Dday)&") "
        Exit function
    end if

    if (nowDate>D6Day) and (nowDate=<D2Day) then
        mayTicketCancelChargePro = 20
        ticketCancelStr = "관람일 6일~3일전 취소시 (관람일 : "&CStr(Dday)&") "
        Exit function
    end if

    if (nowDate>D2Day) and (nowDate=<DDay) then
        mayTicketCancelChargePro = 30
        ticketCancelStr = "관람일 2일~1일전 취소시 (관람일 : "&CStr(Dday)&") "
        Exit function
    end if

end Function

'// 여행상품
function TravelOrderCheck(iorderserial,byRef mayTravelCancelChargePrice,byRef travelCancelDisabled,byRef travelCancelStr)
    Dim sqlStr

	'// 발권일 다음날부터 취소수수료 발생
	'// 출발 6일전부터는 취소불가

    mayTravelCancelChargePrice = 0
    travelCancelDisabled     = False

    sqlStr = " select top 1 "
    sqlStr = sqlStr & "  	(case when DateDiff(d,s.returnExpireDate, getdate()) > 0 then 'N' else 'Y' end) as cancelOK "
    sqlStr = sqlStr & " 	,(case when DateDiff(d,d.beasongdate, getdate()) <= 0 then 0 else ti.bookingCharge end) as cancelCharge "
    sqlStr = sqlStr & " 	,(case "
    sqlStr = sqlStr & " 			when DateDiff(d,s.returnExpireDate, getdate()) > 0 then '출발 6일전 취소환불불가' "
    sqlStr = sqlStr & " 			when DateDiff(d,d.beasongdate, getdate()) > 0 then '취소 수수료 차감' "
    sqlStr = sqlStr & " 			else ''  "
    sqlStr = sqlStr & " 	end) as cancelSTR "
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m "
    sqlStr = sqlStr & " 	join db_order.dbo.tbl_order_detail d "
    sqlStr = sqlStr & " 	on m.orderserial = d.orderserial "
    sqlStr = sqlStr & " 	Join db_item.dbo.tbl_ticket_Schedule s "
    sqlStr = sqlStr & " 	on d.itemid=s.tk_itemid "
    sqlStr = sqlStr & " 	and d.itemoption=s.tk_itemoption "
    sqlStr = sqlStr & " 	join db_item.[dbo].[tbl_ticket_itemInfo] ti "
    sqlStr = sqlStr & " 	on ti.itemid = d.itemid "
    sqlStr = sqlStr & " where d.orderserial='"&iorderserial&"' "
    sqlStr = sqlStr & " and d.itemid<>0 "
    sqlStr = sqlStr & " and d.cancelyn<>'Y' "
    sqlStr = sqlStr & " order by d.beasongdate "

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
		travelCancelDisabled		= (rsget("cancelOK") = "N")
		mayTravelCancelChargePrice	= rsget("cancelCharge")
		travelCancelStr				= rsget("cancelSTR")
    end if
	rsget.close

end function

function TravelOrderCheckArr(iorderserial)
    Dim sqlStr

	'// 발권일 다음날부터 취소수수료 발생
	'// 출발 6일전부터는 취소불가

    TravelOrderCheckArr = ""

    sqlStr = " select d.idx as orderdetailidx "
	sqlStr = sqlStr & "  	,(case when DateDiff(d,s.returnExpireDate, getdate()) > 0 then 'N' else 'Y' end) as cancelOK "
	sqlStr = sqlStr & " 	,(case when DateDiff(d,d.beasongdate, getdate()) <= 0 then 0 else ti.bookingCharge end) as cancelCharge "
	sqlStr = sqlStr & " 	,(case "
    sqlStr = sqlStr & " 			when DateDiff(d,s.returnExpireDate, getdate()) > 0 then '출발 6일전 취소환불불가' "
    sqlStr = sqlStr & " 			when DateDiff(d,d.beasongdate, getdate()) > 0 then '취소 수수료 차감' "
	sqlStr = sqlStr & " 			else ''  "
    sqlStr = sqlStr & " 	end) as cancelSTR "
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m "
    sqlStr = sqlStr & " 	join db_order.dbo.tbl_order_detail d "
    sqlStr = sqlStr & " 	on m.orderserial = d.orderserial "
    sqlStr = sqlStr & " 	Join db_item.dbo.tbl_ticket_Schedule s "
    sqlStr = sqlStr & " 	on d.itemid=s.tk_itemid "
    sqlStr = sqlStr & " 	and d.itemoption=s.tk_itemoption "
    sqlStr = sqlStr & " 	join db_item.[dbo].[tbl_ticket_itemInfo] ti "
    sqlStr = sqlStr & " 	on ti.itemid = d.itemid "
    sqlStr = sqlStr & " where d.orderserial='"&iorderserial&"' "
    sqlStr = sqlStr & " and d.itemid<>0 "
    sqlStr = sqlStr & " and d.cancelyn<>'Y' "
    sqlStr = sqlStr & " order by d.beasongdate "

	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdText
    if Not rsget.Eof then
		TravelOrderCheckArr = rsget.getRows()
    end if
	rsget.close

end function

function GetOrderserialWithOutmallOrderserial(ioutmallorderserial, byRef iorderserial)
	dim sqlStr

	iorderserial = ""

    sqlStr = " select top 1 orderserial, sellsite "
    sqlStr = sqlStr & "  from "
    sqlStr = sqlStr & "  [db_threepl].[dbo].[tbl_xSite_TMPOrder] "
    sqlStr = sqlStr & "  where outmallorderserial = '" + CStr(ioutmallorderserial) + "' "
    sqlStr = sqlStr & "  and orderserial is not null"

    rsget_TPL.CursorLocation = adUseClient
    rsget_TPL.Open sqlStr, dbget_TPL, adOpenForwardOnly, adLockReadOnly
    if Not rsget_TPL.Eof then
		iorderserial	= rsget_TPL("orderserial")
    end if
	rsget_TPL.close

end Function

function ereg(strOriginalString, strPattern, varIgnoreCase)
    ' Function matches pattern, returns true or false
    ' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)
    dim objRegExp : set objRegExp = new RegExp
    with objRegExp
        .Pattern = strPattern
        .IgnoreCase = varIgnoreCase
        .Global = True
    end with
    ereg = objRegExp.test(strOriginalString)
    set objRegExp = nothing
end Function

function ereg_replace(strOriginalString, strPattern, strReplacement, varIgnoreCase)
    ' Function replaces pattern with replacement
    ' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)
    dim objRegExp : set objRegExp = new RegExp
    with objRegExp
        .Pattern = strPattern
        .IgnoreCase = varIgnoreCase
        .Global = True
    end with
    ereg_replace = objRegExp.replace(strOriginalString, strReplacement)
    set objRegExp = nothing
end function

function GetUseridWithAsterisk(curUserid, useAsterisk)
	dim resultStr, leftLen, rightLen

	If IsNull(useAsterisk) Then
		useAsterisk = True
	End If

	If useAsterisk = False Then
		GetUseridWithAsterisk = curUserid
		Exit Function
	End If

	resultStr = "ERR"
	If IsNull(curUserid) Then
		GetUseridWithAsterisk = resultStr
		Exit Function
	End If

	'// 가운데 3글자
	If Len(curUserid) <= 3 Then
		resultStr = ereg_replace(curUserid, ".", "*", True)
		GetUseridWithAsterisk = resultStr
		Exit Function
	End If

	If (Len(curUserid) - 3) Mod 2 = 0 Then
		leftLen = (Len(curUserid) - 3) / 2
		rightLen = Len(curUserid) - 3 - leftLen
	Else
		leftLen = Int((Len(curUserid) - 3) / 2) + 1
		rightLen = Len(curUserid) - 3 - leftLen
	End If

	resultStr = Left(curUserid, leftLen) & ereg_replace(Mid(curUserid, 3, 3), ".", "*", True) & Right(curUserid, rightLen)
	GetUseridWithAsterisk = resultStr
end Function

function GetUsernameWithAsterisk(curUsername, useAsterisk)
	dim resultStr, leftLen, rightLen

	If IsNull(useAsterisk) Then
		useAsterisk = True
	End If

	If useAsterisk = False Then
		GetUsernameWithAsterisk = curUsername
		Exit Function
	End If

	resultStr = "ERR"
	If IsNull(curUsername) Then
		GetUsernameWithAsterisk = resultStr
		Exit Function
	End If

	'// 가운데 1글자
	If Len(curUsername) <= 1 Then
		resultStr = ereg_replace(curUsername, ".", "*", True)
		GetUsernameWithAsterisk = resultStr
		Exit Function
	End If

	If (Len(curUsername) - 1) Mod 2 = 0 Then
		leftLen = (Len(curUsername) - 1) / 2
		rightLen = Len(curUsername) - 1 - leftLen
	Else
		leftLen = Int((Len(curUsername) - 1) / 2) + 1
		rightLen = Len(curUsername) - 1 - leftLen
	End If

	resultStr = Left(curUsername, leftLen) & ereg_replace(Mid(curUsername, 1, 1), ".", "*", True) & Right(curUsername, rightLen)
	GetUsernameWithAsterisk = resultStr
end Function

Class COrderDetailItemMakerGroupInfoItem
	public Fgroupid
	public Fmakerid

	public Fcompany_name
	public Fcompany_no
	public Fceoname
	public Fcompany_uptae
	public Fcompany_upjong
	public Fcompany_zipcode
	public Fcompany_address
	public Fcompany_address2
	public Fcompany_tel
	public Fcompany_fax
	public Freturn_zipcode
	public Freturn_address
	public Freturn_address2
	public Fmanager_name
	public Fmanager_phone
	public Fmanager_hp
	public Fmanager_email
	public Fdeliver_name
	public Fdeliver_phone
	public Fdeliver_hp
	public Fdeliver_email
	public Fregdate
	public Flastupdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CEtcPaymentItem

	public Facctdiv
	public FacctdivName
	public Facctamount
	public FrealPayedsum
	public FacctAuthCode
	public FacctAuthDate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CUpcheBeasongPayItem

	public Fmakerid
	public Fdefaultfreebeasonglimit
	public Fdefaultdeliverpay

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COrderItemSummaryItem

	public Ftenbeacnt
	public Fupbeacnt
	public Fbrandcnt

	Private Sub Class_Initialize()
		Ftenbeacnt = 0
		Fupbeacnt = 0
		Fbrandcnt = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class




Class COrderDetailItem
    public Fidx
	public Forderserial
	public Fprdcode
	public Fitemid
	public Fitemoption
	public Fmasteridx
	public Fmakerid
	public Fitemno
	public Fitemcost
	public Fmileage
	public Fcancelyn
	public Fcurrstate
	public Fsongjangno
	public Fsongjangdiv
	public Fitemname
	public Fitemoptionname

	public Forgsuplycash
	public FbuycashCouponNotApplied
	public Fbuycash

	public Fvatinclude
	public Fbeasongdate
	public Fisupchebeasong
	public Fissailitem
	public Fupcheconfirmdate
	public Foitemdiv
    public FListImage
    public FSmallImage
    public Frequiredetail

    public Fsongjangdivname
    public Ffindurl

    public Forgitemcost					'소비자가
    public FitemcostCouponNotApplied	'판매가(할인가)
    public FplusSaleDiscount			'플러스세일할인액
    public FspecialshopDiscount			'우수고객할인액
	public FetcDiscount					'기타할인액

	Public FodlvType
	public fodlvfixday

    '''기존 버전 고려
    public function getItemcostCouponNotApplied
        if (FitemcostCouponNotApplied<>0) then
            getItemcostCouponNotApplied = FitemcostCouponNotApplied
        else
            getItemcostCouponNotApplied = FItemCost
        end if
    end function

    ''주문제작 상품
    public function IsRequireDetailExistsItem()
        IsRequireDetailExistsItem = (Foitemdiv="06") or (Frequiredetail<>"")
    end function

    public function getRequireDetailHtml()
		getRequireDetailHtml = nl2br(Frequiredetail)

		getRequireDetailHtml = replace(getRequireDetailHtml,CAddDetailSpliter,"<br><br>")
	end function

    ''소비자가
    public Forgprice
    public Fbonuscouponidx
    public Fitemcouponidx
    public FreducedPrice

	'상품할인 적용 주문인지 체크
    public function IsSaleDiscountItem()
        IsSaleDiscountItem = (GetSaleDiscountPrice() > 0)
    end function

	'상품쿠폰 적용 주문인지 체크
    public function IsItemCouponDiscountItem()
        IsItemCouponDiscountItem = false
        if (Not IsNull(Fitemcouponidx) and (Fitemcouponidx<>0)) then
            IsItemCouponDiscountItem = true
        end if
    end function

    '보너스쿠폰 적용 주문인지 체크
    public function IsBonusCouponDiscountItem()
        IsBonusCouponDiscountItem = false
        if (Not IsNull(Fbonuscouponidx) and (Fbonuscouponidx<>0) and (GetItemCouponPrice > GetBonusCouponPrice))  then
            IsBonusCouponDiscountItem = true
        end if
    end function

	'기타할인 적용 주문인지 체크
    public function IsEtcDiscountItem()
        IsEtcDiscountItem = false
        if (Not IsNull(Fbonuscouponidx) and (Fbonuscouponidx<>0) and (GetBonusCouponPrice > GetEtcDiscountPrice))  then
            IsEtcDiscountItem = true
        end if
    end function

	'// 매입가 할인적용되었는지
    public function IsBuyCashSaleApplied()
		IsBuyCashSaleApplied = (Forgsuplycash > FbuycashCouponNotApplied) and (FbuycashCouponNotApplied <> 0)
    end function

	'// 매입가 상품쿠폰적용되었는지
    public function IsBuyCashItemCouponApplied()
		IsBuyCashItemCouponApplied = (FbuycashCouponNotApplied > Fbuycash)
    end function

	'// 플러스 세일상품
    public function IsPlusSaleItem()
		IsPlusSaleItem = (FplusSaleDiscount <> 0)
    end function

	'// 마일리지 삽 상품
    public function IsMileageShopItem()
		IsMileageShopItem = (Foitemdiv = 82)
    end function

    '우수고객할인 적용 주문인지 체크
    public function IsSpecialShopDiscountItem()
        if (FitemcostCouponNotApplied = 0) then
        	'과거데이타
        	if (Not IsItemCouponDiscountItem) and (Not IsBonusCouponDiscountItem) and (Fissailitem = "N") then
        		'TODO : 소비자가변경, 옵션가변경이 있는경우 부정확한 값이 된다.
        		GetItemCouponDiscountPrice = (Forgprice - Fitemcost) = 0
        		exit function
        	end if

        	GetItemCouponDiscountPrice = false
        	exit function
        end if

		if (FspecialshopDiscount > 0) then
			IsSpecialShopDiscountItem = true
		else
			IsSpecialShopDiscountItem = false
		end if
    end function

	'상품쿠폰할인액
    public function GetItemCouponDiscountPrice()
        if (FitemcostCouponNotApplied = 0) then
        	'과거데이타
        	if (IsItemCouponDiscountItem = true) and (Not IsBonusCouponDiscountItem) and (Fissailitem = "N") then
        		'TODO : 소비자가변경, 옵션가변경, 우수고객할인이 있는경우 부정확한 값이 된다.
        		GetItemCouponDiscountPrice = Forgprice - Fitemcost
        		exit function
        	end if

        	GetItemCouponDiscountPrice = 0
        	exit function
        end if

        GetItemCouponDiscountPrice = FitemcostCouponNotApplied - Fitemcost
    end function

	'보너스쿠폰할인액
    public function GetBonusCouponDiscountPrice()
        GetBonusCouponDiscountPrice = GetItemCouponPrice - GetBonusCouponPrice
    end function

	'기타할인할인액
	public function GetEtcDiscountDiscountPrice()
        GetEtcDiscountDiscountPrice = GetBonusCouponPrice - GetEtcDiscountPrice
    end function

	'상품할인액
    public function GetSaleDiscountPrice()
        if (FitemcostCouponNotApplied = 0) then
        	'과거데이타
        	if (Not IsBonusCouponDiscountItem) and (Not IsItemCouponDiscountItem) and (Fissailitem = "Y") then
        		'TODO : 소비자가변경, 옵션가변경, 우수고객할인이 있는경우 부정확한 값이 된다.
        		GetSaleDiscountPrice = Forgprice - Fitemcost
        		exit function
        	end if

        	GetSaleDiscountPrice = 0
        	exit function
        end if

        GetSaleDiscountPrice = (Forgitemcost - (FitemcostCouponNotApplied + FplusSaleDiscount + FspecialshopDiscount))
    end function

    public function IsOldJumun()
    	'2011년 4월 1일 이전 주문 또는 그 주문에 대한 마이너스주문
    	IsOldJumun = (Forgitemcost = 0)
    end function

	public function GetOrgItemCostColor()
		if IsOldJumun then
			GetOrgItemCostColor = "gray"
		else
			GetOrgItemCostColor = "black"
		end if
	end function

	public function GetOrgItemCostPrice()
		if IsOldJumun then
			GetOrgItemCostPrice = Forgprice
		else
			GetOrgItemCostPrice = Forgitemcost
		end if
	end function

	public function GetSaleColor()
		if IsOldJumun then
			if (Fissailitem = "Y") or (Fissailitem = "P") or ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
				GetSaleColor = "red"
			else
				GetSaleColor = "black"
			end if
		else
			if (Forgitemcost <> FitemcostCouponNotApplied) then
				GetSaleColor = "red"
			else
				GetSaleColor = "black"
			end if
		end if
	end function

	public function GetSalePrice()
		if IsOldJumun then
			if (Fissailitem = "Y") or (Fissailitem = "P") or ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
				GetSalePrice = Fitemcost
			else
				GetSalePrice = Forgprice
			end if
		else
			GetSalePrice = FitemcostCouponNotApplied
		end if
	end function

	public function GetSaleText()
		dim result

		result = ""
		if IsOldJumun then
			if (Fissailitem = "Y") or (Fissailitem = "P") or ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
				if (Fissailitem = "Y") then
					if (Forgprice <= Fitemcost) then
						result = result + "할인상품 + 소비자가 인하" + vbCrLf
					else
						result = result + "할인상품" + vbCrLf
					end if
				end if
				if (Fissailitem = "P") then
					result = result + "플러스할인" + vbCrLf
				end if
				if ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
					result = result + "우수고객할인 또는 소비자가/옵션가 변동" + vbCrLf
				end if
			else
				result = "정상가격"
			end if
		else
			if (Forgitemcost <> FitemcostCouponNotApplied) then
				if (Fissailitem = "Y") then
					result = result + "할인상품 : " + CStr(GetSaleDiscountPrice) + "원" + vbCrLf
				end if
				if (FplusSaleDiscount > 0) then
					result = result + "플러스할인 : " + CStr(FplusSaleDiscount) + "원" + vbCrLf
				end if
				if (FspecialshopDiscount > 0) then
					result = result + "우수회원할인 : " + CStr(FspecialshopDiscount) + "원" + vbCrLf
				end if
			else
				result = "정상가격"
			end if
		end if

		GetSaleText = result
	end function

	public function GetItemCouponColor()
		if (IsItemCouponDiscountItem = true) then
			GetItemCouponColor = "green"
		else
			GetItemCouponColor = "black"
		end if
	end function

	public function GetItemCouponPrice()
		GetItemCouponPrice = Fitemcost
	end function

	public function GetItemCouponText()
		dim result

		result = ""
		if IsOldJumun then
			if (IsItemCouponDiscountItem = true) then
				if (GetSalePrice <> GetItemCouponPrice) then
					result = result + "상품쿠폰적용상품" + vbCrLf
				else
					result = result + "배송비쿠폰적용상품" + vbCrLf
				end if
			else
				result = "정상가격"
			end if
		else
			if (IsItemCouponDiscountItem = true) then
				if (GetItemCouponDiscountPrice = 0) then
					result = result + "배송비쿠폰적용상품" + vbCrLf
				else
					result = result + "상품쿠폰 : " + CStr(GetItemCouponDiscountPrice) + "원" + vbCrLf
				end if
			else
				result = "정상가격"
			end if
		end if

		GetItemCouponText = result
	end function

	public function GetBonusCouponColor()
		if (IsBonusCouponDiscountItem = true) then
			GetBonusCouponColor = "purple"
		else
			GetBonusCouponColor = "black"
		end if
	end function

	public function GetBonusCouponPrice()
		GetBonusCouponPrice = (FreducedPrice + FetcDiscount)
	end function

	public function GetBonusCouponText()
		dim result

		result = ""
		if IsOldJumun then
			if (IsBonusCouponDiscountItem = true) then
				result = result + "보너스쿠폰" + vbCrLf
			else
				result = "정상가격"
			end if
		else
			if (IsBonusCouponDiscountItem = true) then
				result = result + "보너스쿠폰 : " + CStr(GetBonusCouponDiscountPrice) + "원" + vbCrLf
			else
				result = "정상가격"
			end if
		end if

		GetBonusCouponText = result
	end function

	public function GetEtcDiscountColor()
		if (IsEtcDiscountItem = true) then
			GetEtcDiscountColor = "red"
		else
			GetEtcDiscountColor = "black"
		end if
	end function

	public function GetEtcDiscountPrice()
		GetEtcDiscountPrice = FreducedPrice
	end function

	public function GetEtcDiscountText()
		dim result

		result = ""
		if IsOldJumun then
			if (IsEtcDiscountItem = true) then
				result = result + "기타할인" + vbCrLf
			else
				result = "정상가격"
			end if
		else
			if (IsEtcDiscountItem = true) then
				result = result + "기타할인 : " + CStr(GetEtcDiscountDiscountPrice) + "원" + vbCrLf
			else
				result = "정상가격"
			end if
		end if

		GetEtcDiscountText = result
	end function

	public function GetSaleBuycashColor()
		if (IsBuyCashSaleApplied = true) then
			GetSaleBuycashColor = "red"
		else
			GetSaleBuycashColor = "black"
		end if
	end function

	public function GetSaleBuycashText()
		dim result

		result = ""

		if (IsBuyCashSaleApplied = true) then
			result = result + "매입가세일적용" + vbCrLf
		else
			result = "정상가격"
		end if

		GetSaleBuycashText = result
	end function

	public function GetItemCouponBuycashColor()
		if (IsBuyCashItemCouponApplied = true) then
			GetItemCouponBuycashColor = "green"
		else
			GetItemCouponBuycashColor = "black"
		end if
	end function

	public function GetItemCouponBuycashText()
		dim result

		result = ""

		if (IsBuyCashItemCouponApplied = true) then
			result = result + "매입가상품쿠폰적용" + vbCrLf
		else
			result = "정상가격"
		end if

		GetItemCouponBuycashText = result
	end function

    ''All@ 할인된가격
    public function getAllAtDiscountedPrice()
        getAllAtDiscountedPrice = 0

	    if (IsNULL(Fbonuscouponidx) or (Fbonuscouponidx = 0)) and (Fitemcost > Freducedprice) then
	            getAllAtDiscountedPrice = Fitemcost - Freducedprice
	    else
	        getAllAtDiscountedPrice = 0
	    end if
    end function

    '' %할인권 할인금액 or 카드 할인금액
    public function getPercentBonusCouponDiscountedPrice()
        getPercentBonusCouponDiscountedPrice = 0

        if (Freducedprice <> 0) then
            if (Fbonuscouponidx <> 0)  and (Fitemcost > Freducedprice) then
                getPercentBonusCouponDiscountedPrice = Fitemcost - Freducedprice
            end if
        end if
    end function

	public function CancelStateStr()
		CancelStateStr = "정상"

		if Fcancelyn="Y" then
			CancelStateStr ="취소"
		elseif Fcancelyn="D" then
			CancelStateStr ="삭제"
		elseif Fcancelyn="A" then
			CancelStateStr ="추가"
		end if
	end function

	public function CancelStateColor()
		if FCancelYn="D" then
			CancelStateColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelStateColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelStateColor = "#000000"
		elseif UCase(FCancelYn)="A" then
			CancelStateColor = "#0000FF"
		end if
	end function

	Public function GetStateName()
        if FCurrState="2" then
            if FIsUpchebeasong="Y" then
		        GetStateName = "업체통보"
		    else
		        GetStateName = "물류통보"
		    end if
	    elseif FCurrState="3" then
		    GetStateName = "상품준비"
	    elseif FCurrState="7" then
		    GetStateName = "출고완료"
		elseif FCurrState="0" then
		    GetStateName = ""
	    else
		    GetStateName = FCurrState
	    end if
	 end Function

	public function GetStateColor()
	    if FCurrState="2" then
			GetStateColor="#000000"
		elseif FCurrState="3" then
			GetStateColor="#CC9933"
		elseif FCurrState="7" then
			GetStateColor="#FF0000"
		else
			GetStateColor="#000000"
		end if
	end function

	'세일상품
	public function IsSaleItem()
        IsSaleItem = (FIsSailItem="Y") or (FplussaleDiscount>0) or (FspecialShopDiscount>0)  '''or (FIsSailItem="P")  플러스세일인 플러스 세일금액이 있으면. 으로 바뀜. 20110401 부터
        IsSaleItem = IsSaleItem and (Forgitemcost>FitemcostCouponNotApplied)
    end function

	'상품쿠폰
    public function IsItemCouponAssignedItem()
        IsItemCouponAssignedItem = (Fitemcouponidx>0) and (FitemcostCouponNotApplied>FItemCost)
    end function
	'보너스쿠폰
    public function IsSaleBonusCouponAssignedItem()
        IsSaleBonusCouponAssignedItem = (Fbonuscouponidx>0)
    end function

     ''마일리지샵 상품
    public function IsMileShopSangpum()
		IsMileShopSangpum = false

		if Foitemdiv="82" then
			IsMileShopSangpum = true
		end if
	end function

	'' 마스터 현재상태를 같이 넘겨야함.
    public function GetItemDeliverStateName(CurrMasterIpkumDiv, CurrMasterCancelyn)
        if ((CurrMasterCancelyn="Y") or (CurrMasterCancelyn="D") or (Fcancelyn="Y")) then
            GetItemDeliverStateName = "취소"
        else
            if (CurrMasterIpkumDiv="0") then
                GetItemDeliverStateName = "결제오류"
            elseif (CurrMasterIpkumDiv="1") then
                GetItemDeliverStateName = "주문실패"
            elseif (CurrMasterIpkumDiv="2") or (CurrMasterIpkumDiv="3") then
                GetItemDeliverStateName = "주문접수"
            elseif (CurrMasterIpkumDiv="9") then
                GetItemDeliverStateName = "반품"
            else
                if (IsNull(Fcurrstate) or (Fcurrstate=0)) then
            		GetItemDeliverStateName = "결제완료"
                elseif Fcurrstate="2" then
                    GetItemDeliverStateName = "주문통보"
            	elseif Fcurrstate="3" then
            		GetItemDeliverStateName = "상품준비중"
            	elseif Fcurrstate="7" then
            		GetItemDeliverStateName = "출고완료"
            	else
            		GetItemDeliverStateName = ""
            	end if
            end if
        end if
    end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COrderMasterItem
	public Forderserial
	public Fidx
	public Fjumundiv
	public Fuserid
	public Faccountname
	public Faccountdiv
	public Faccountno
	public Ftotalvat
	public Ftotalcost
	public Ftotalmileage
	public Ftotalsum
	public Fipkumdiv
	public Fipkumdate
	public Fregdate
	public Fbeadaldiv
	public Fbeadaldate
	public Fcancelyn
	public Fbuyname
	public Fbuyphone
	public Fbuyhp
	public Fbuyemail
	public Freqname
	public Freqzipcode
	public Freqaddress
	public Freqphone
	public Freqhp
	public Freqemail
	public Fcomment
	public Fdeliverno
	public Fsitename
	public Fpartnercompanyname
	public Fpaygatetid
	public Fdiscountrate
	public Fsubtotalprice
	public Fresultmsg
	public Frduserid
	public Fmiletotalprice
	public Fjungsanflag
	public Freqzipaddr
	public Fauthcode
	public Fsongjangdiv
	public Frdsite

	public Ftencardspend
	public FbCpnIdx

	public Fbeasongmemo

	public FInsureCd
	public Fcashreceiptreq
	public FcashreceiptTid
	public FcashreceiptIdx
	public Finireceipttid
	public Freferip
	public Fuserlevel
	public Flinkorderserial
	public Fspendmembership
	public Fsentenceidx
	public Fbaljudate
	public FuserDisplayYn

	public Fpggubun
	Public Fordersheetyn

	public Fallatdiscountprice

	'보조결제
	public FsumPaymentEtc

	'배송비 쿠폰 사용금액
	Public FDeliverpriceCouponNotApplied
	Public FDeliverprice

	'상품쿠폰적용안한 판매가(할인가 : 우수회원,플러스세일은 적용)
	public FsubtotalpriceCouponNotApplied

	public Fcash_receipt_tid

    ''플라워주문 관련
    public Freqdate
	public Freqtime
	public Fcardribbon
	public Fmessage
	public Ffromname

	''해외배송관련
	public FDlvcountryCode

	public FcountryNameKr
	public FcountryNameEn
	public FemsAreaCode
    public FemsZipCode
    public FitemGubunName
    public FgoodNames
    public FitemWeigth
    public FitemUsDollar
    public FemsInsureYn
    public FemsInsurePrice
    public FemsDlvCost

    ''OkCashbag 추가
    public FokcashbagSpend

	Public FspendTenCash
	Public Fspendgiftmoney
	public Forgorderserial

    '''주결제수단 금액 = subtotalPrice-FsumPaymentEtc
    public function TotalMajorPaymentPrice()
        TotalMajorPaymentPrice = FsubtotalPrice-FsumPaymentEtc
    end function

	'증빙서류 발급가능한지
    public function GetPaperAvailableString()
        GetPaperAvailableString = ""

        if (Fcancelyn = "Y") then
        	GetPaperAvailableString = "취소된 주문입니다."
        	exit function
        end if

        if (FIpkumDiv < 4) then
        	GetPaperAvailableString = "결제이전 주문입니다."
        	exit function
        end if

        if (Faccountdiv <> "7") and (Faccountdiv <> "20") and (sumPaymentEtc < 1) then
        	GetPaperAvailableString = "발행대상 금액이 없습니다."
        	exit function
        end if
    end function

	'증빙서류신청이 있었는지
    public function IsPaperRequestExist()
        IsPaperRequestExist = false

        if (IsPaperRequested or IsPaperFinished) then
        	IsPaperRequestExist = true
        end if
    end function

	'증빙서류 종류
    public function GetPaperType()
        GetPaperType = ""

        if (FcashreceiptReq = "R") or (FcashreceiptReq = "S") then
        	GetPaperType = "R"
        	Exit function
        end if

        if (FcashreceiptReq = "T") or (FcashreceiptReq = "U") then
        	GetPaperType = "T"
        	exit function
        end if

        if (Faccountdiv = "7") or (Faccountdiv = "20") and (FAuthCode <> "") then
        	GetPaperType = "R"
        end if
    end function

	'증빙서류 TID (세금계산서는 주문번호로 별도 검색)
    public function GetPaperTID()
        GetPaperTID = ""

        if Not IsPaperRequestExist then
        	exit function
        end if

        if Not IsPaperFinished then
        	exit function
        end if

        if GetPaperType <> "R" then
        	exit function
        end if

        if (Faccountdiv = "20") then
        	if IsNull(Fcash_receipt_tid) or (Fcash_receipt_tid = "") then
        		GetPaperTID = Fpaygatetid
        	else
        		GetPaperTID = Fcash_receipt_tid
        	end if
        else
        	GetPaperTID = Fcash_receipt_tid
        end if
    end function

	'증빙서류 발급신청상태인지
    public function IsPaperRequested()
        IsPaperRequested = false

        if (Faccountdiv = "7") or (Faccountdiv = "20") then
        	if ((FcashreceiptReq = "R") or (FcashreceiptReq = "T")) and (IsNull(FAuthCode) or FAuthCode = "") then
        		IsPaperRequested = true
        	end if
		else
			if (FcashreceiptReq = "R") or (FcashreceiptReq = "T") then
				IsPaperRequested = true
			end if
        end if
    end function

	'증빙서류 발급완료상태인지
    public function IsPaperFinished()
        IsPaperFinished = false

        if (Faccountdiv = "7") or (Faccountdiv = "20") then
        	if ((FcashreceiptReq = "R") or (FcashreceiptReq = "T")) and (FAuthCode <> "") then
        		IsPaperFinished = true
        	elseif (FAuthCode <> "") then
        		IsPaperFinished = true
        	end if
		else
			if (FcashreceiptReq = "S") or (FcashreceiptReq = "U") then
				IsPaperFinished = true
			end if
        end if
    end function

    ''데이콤 가상계좌 결제인지
    public function IsDacomCyberAccountPay()
        IsDacomCyberAccountPay = false
        if (FAccountdiv<>"7") then Exit function

        if (FAccountNo="국민 470301-01-014754") _
            or (FAccountNo="신한 100-016-523130") _
            or (FAccountNo="우리 092-275495-13-001") _
            or (FAccountNo="하나 146-910009-28804") _
            or (FAccountNo="기업 277-028182-01-046") _
            or (FAccountNo="농협 029-01-246118") then
                IsDacomCyberAccountPay = false
        else
            IsDacomCyberAccountPay = true
        end if
    end function

	''해외배송인지여부
	public function IsForeignDeliver()
        IsForeignDeliver = (Not IsNULL(FDlvcountryCode)) and (FDlvcountryCode<>"") and (FDlvcountryCode<>"KR") and (FDlvcountryCode<>"ZZ") and (FDlvcountryCode<>"Z4") and (FDlvcountryCode<>"QQ")
    end function

    ''군부대배송
    public function IsArmiDeliver()
        IsArmiDeliver = (Not IsNULL(FDlvcountryCode)) and (FDlvcountryCode="ZZ")
    end function

    ''퀵배송
    public function IsQuickDeliver()
        IsQuickDeliver = (Not IsNULL(FDlvcountryCode)) and (FDlvcountryCode="QQ")
    end function

    public function IsOldJumun()
    	'2011년 4월 1일 이전 주문 또는 그 주문에 대한 마이너스주문
    	IsOldJumun = (FsubtotalpriceCouponNotApplied = 0)
    end function

    public function IsErrSubtotalPrice()
        IsErrSubtotalPrice = (Fsubtotalprice <> (Ftotalsum - (Ftencardspend + Fmiletotalprice + Fspendmembership + Fallatdiscountprice)))
    end function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(Fipkumdiv)="0") or (CStr(Fipkumdiv)="1") or (CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

    ''결제했는지 여부
    public function IsPayedOrder()
        IsPayedOrder = (FIpkumdiv>3) and (FIpkumdiv<9)
    end function

	'직접수령여부
    public function IsReceiveSiteOrder
        IsReceiveSiteOrder = (Fjumundiv="7")
    end Function

    public function GetMasterDeliveryName()
        GetMasterDeliveryName = ""
        if IsNULL(Fsongjangdiv) then Exit function

        if Fsongjangdiv="24" then
            GetMasterDeliveryName = "사가와"
        elseif Fsongjangdiv="2" then
            GetMasterDeliveryName = "현대"
        else
            GetMasterDeliveryName = Fsongjangdiv
        end if
    end function

	'/사용중지 공용펑션에 공통함수 같이 쓸것 2016.06.30 한용민
	public function GetUserLevelColor()
		if Fuserlevel="1" then
			GetUserLevelColor = "#44DD44"   ''Green
		elseif Fuserlevel="2" then
			GetUserLevelColor = "#4444FF"   ''BLUE
		elseif Fuserlevel="3" then
			GetUserLevelColor = "#FF1111"   ''VIP SILVER
		elseif Fuserlevel="4" then
			GetUserLevelColor = "#7D2448"   ''VIP GOLD
		elseif Fuserlevel="6" then
			GetUserLevelColor = "red"  ''VVIP
		elseif Fuserlevel="9" then
			GetUserLevelColor = "#FF11FF"  '' BIZ
		elseif Fuserlevel="7" then
			GetUserLevelColor = "black"  '' staff
		elseif Fuserlevel="7" then
			GetUserLevelColor = "#FF11FF"  '' famliy
		elseif Fuserlevel="5" then
			GetUserLevelColor = "#FF6611"  ''orange
		elseif Fuserlevel="0" then
			GetUserLevelColor = "#DDDD22"  ''yellow
		else
			GetUserLevelColor = "#000000"
		end if
	end function

	'/사용중지 공용펑션에 공통함수 같이 쓸것 2016.06.30 한용민
	public function GetUserLevelName()

		if Fuserlevel="1" then
			GetUserLevelName = "Green"   		''Green
		elseif Fuserlevel="2" then
			GetUserLevelName = "Blue"   		''BLUE
		elseif Fuserlevel="3" then
			GetUserLevelName = "VIP Silver"   	''VIP SILVER
		elseif Fuserlevel="4" then
			GetUserLevelName = "VIP Gold"   	''VIP GOLD
		elseif Fuserlevel="6" then
			GetUserLevelName = "VVIP"   		''VVIP
		elseif Fuserlevel="9" then
			GetUserLevelName = "BIZ"  		'' BIZ
		elseif Fuserlevel="7" then
			GetUserLevelName = "Staff"  		'' staff
		elseif Fuserlevel="5" then
			GetUserLevelName = "Orange"  		''orange
		elseif Fuserlevel="0" then
			GetUserLevelName = "Yellow"  		''yellow
		else
			GetUserLevelName = "Yellow"			''??
		end if
	end function

	public function GetJumunDivName()
		if Fjumundiv="1" then
			GetJumunDivName = "웹주문"
		elseif Fjumundiv="3" then
			GetJumunDivName = "예약주문"
		elseif Fjumundiv="4" then
			GetJumunDivName = "티켓"
		elseif Fjumundiv="5" then
			GetJumunDivName = "외부몰"
		elseif Fjumundiv="6" then
			'// 아카데미DIY상품 -> 맞교환
			GetJumunDivName = "맞교환"
		elseif Fjumundiv="7" then
			GetJumunDivName = "현장수령"
		elseif Fjumundiv="8" then
			GetJumunDivName = "강좌주문"
		elseif Fjumundiv="9" then
			GetJumunDivName = "마이너스"
		else
			GetJumunDivName = Fjumundiv
		end if
	end function


	public function CancelYnName()
		CancelYnName = "정상"

		if Fcancelyn="Y" then
			CancelYnName ="취소"
		elseif Fcancelyn="D" then
			CancelYnName ="삭제"
		elseif Fcancelyn="A" then
			CancelYnName ="추가"
		end if
	end function

	public function CancelYnColor()
		CancelYnColor = "#000000"

		if FCancelYn="D" then
			CancelYnColor = "#FF0000"
		elseif UCase(FCancelYn)="Y" then
			CancelYnColor = "#FF0000"
		elseif FCancelYn="N" then
			CancelYnColor = "#000000"
		end if
	end function


	public function IpkumDivColor()
		if Fipkumdiv="0" then
			IpkumDivColor="#FF0000"
		elseif Fipkumdiv="1" then
			IpkumDivColor="#44BBBB"
		elseif Fipkumdiv="2" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="3" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="4" then
			IpkumDivColor="#0000FF"
		elseif Fipkumdiv="5" then
			IpkumDivColor="#CC9933"
		elseif Fipkumdiv="6" then
			IpkumDivColor="#FF00FF"
		elseif Fipkumdiv="7" then
			IpkumDivColor="#EE2222"
		elseif Fipkumdiv="8" then
			IpkumDivColor="#EE2222"
		elseif Fipkumdiv="9" then
			IpkumDivColor="#FF0000"
		end if
	end function

	Public function JumunMethodName()
		if Faccountdiv="7" then
			JumunMethodName="무통장"
		elseif Faccountdiv="14" then
			JumunMethodName="편의점결제"
		elseif Faccountdiv="100" then
			JumunMethodName="신용카드"
		elseif Faccountdiv="20" then
			JumunMethodName="실시간이체"
		elseif Faccountdiv="30" then
			JumunMethodName="포인트"
		elseif Faccountdiv="50" then
			JumunMethodName="입점몰결제"
		elseif Faccountdiv="80" then
			JumunMethodName="All@카드"
		elseif Faccountdiv="90" then
			JumunMethodName="상품권결제"
		elseif Faccountdiv="110" then
			JumunMethodName="OK+신용"
		elseif Faccountdiv="400" then
			JumunMethodName="핸드폰결제"
		elseif Faccountdiv="550" then
			JumunMethodName="기프팅"
		elseif Faccountdiv="560" then
			JumunMethodName="기프티콘"
		end if
	end function

	Public function IpkumDivName()
		if Fipkumdiv="0" then
			IpkumDivName="주문대기"
		elseif Fipkumdiv="1" then
			IpkumDivName="주문실패"
		elseif Fipkumdiv="2" then
			IpkumDivName="주문접수"
		elseif Fipkumdiv="3" then
			IpkumDivName="주문접수(3)"
		elseif Fipkumdiv="4" then
			IpkumDivName="결제완료"
		elseif Fipkumdiv="5" then
			IpkumDivName="주문통보"
		elseif Fipkumdiv="6" then
			IpkumDivName="상품준비"
		elseif Fipkumdiv="7" then
			IpkumDivName="일부출고"
	    elseif Fipkumdiv="8" then
			IpkumDivName="상품출고"
		else
			IpkumDivName=Fipkumdiv
		end if
	end Function

	Public function NormalUpcheDeliverState()
		 if IsNull(FCurrState) then
			 NormalUpcheDeliverState = "결제완료"
		 elseif FCurrState="3" then
			 NormalUpcheDeliverState = "상품준비"
		 elseif FCurrState="7" then
			 NormalUpcheDeliverState = "상품출고"
		 else
			 NormalUpcheDeliverState = ""
		 end if
	 end Function

	public function UpCheDeliverStateColor()
		if IsNull(FCurrState) then
			UpCheDeliverStateColor="#3300CC"
		elseif FCurrState="3" then
			UpCheDeliverStateColor="#0000FF"
		elseif FCurrState="7" then
			UpCheDeliverStateColor="#FF0000"
		else
			UpCheDeliverStateColor="#000000"
		end if
	end function


	public function SiteNameColor()
		if Fsitename<>"10x10" then
			SiteNameColor = "#55AA22"
		else
			SiteNameColor = "#000000"
		end if
	end function


	public function SubTotalColor()
		if FSubtotalPrice<0 then
			SubTotalColor = "#DD3333"
		else
			SubTotalColor = "#000000"
		end if
	end function

    ''플라워 지정일 배송 주문 존재여부
    public function IsFixDeliverItemExists()
        IsFixDeliverItemExists = Not IsNULL(Freqdate)
    end function

    '' 플라워 지정일 시각
    public function GetReqTimeText()
        if IsNULL(Freqtime) then Exit function
        GetReqTimeText = Freqtime & "~" & (Freqtime+2) & "시 경"
    end Function

    public function GetPggubunName()
		Select Case Fpggubun
			Case "KA"
				GetPggubunName = "카카오페이"
			Case "IN"
				GetPggubunName = "이니시스"
			Case "DA"
				GetPggubunName = "엘지데이콤"
			Case "NP"
				GetPggubunName = "네이버페이"
			Case "PY"
				GetPggubunName = "페이코"
			Case Else
				GetPggubunName = Fpggubun
		End Select
    end function

	Private Sub Class_Initialize()
        FokcashbagSpend = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COrderMaster
	public FOneItem
	public FItemList()

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectOrderSerial
	public FRectUserID
	public FRectBuyname
	public FRectReqName
	public FRectIpkumName
	public FRectSubTotalPrice

	public FRectBuyHp
	public FRectReqHp
	public FRectBuyPhone
	public FRectReqPhone
	public FRectReqSongjangNo

	public FRectRegStart
	public FRectRegEnd

	public FRectExtSiteName
	public FRectIsMinus
	public FRectIsLecture
	public FRectIsFlower

    public FRectOldOrder
    public FRectDetailIdx
    public FRectIsForeign
	public FRectIsForeignDirect
	public FRectIsQuick
	public FRectJumunItem
	public FRectSongjangno

	Public FTotItemNo
	public FTotItemKind

	public FRectForMail
	public FRectIncMainPayment

    ''detail query 후
    public function GetItemCostSum()

    end function

    public function GetImageFolderName(byval itemid)
		GetImageFolderName = "0" + CStr(Clng(itemid\10000))
	end function

	public function BeasongCD2Name(byval v)
		if v="0101" then
			BeasongCD2Name = "일반택배"
		elseif v="0201" then
			BeasongCD2Name = "포장배송A"
		elseif v="0202" then
			BeasongCD2Name = "포장배송B"
		elseif v="0203" then
			BeasongCD2Name = "포장배송C"
		elseif v="0301" then
			BeasongCD2Name = "직접수령"
		elseif v="0501" then
			BeasongCD2Name = "무료배송"
		end if

		''2011-04
		if v="1000" then
		    BeasongCD2Name = "텐바이텐"
		elseif v="2000" then
			BeasongCD2Name = "업체"
		elseif v="0999" then
			BeasongCD2Name = "해외"
		elseif v="0901" then
			BeasongCD2Name = "착불"
		elseif Left(v,2)="90" then
		    BeasongCD2Name = "업체조건"
		end if
	end function

	public function BeasongOptionString(byval beasongoptionname)
		dim result

		result = ""
		if (Not IsNull(beasongoptionname)) and (beasongoptionname <> "") and (beasongoptionname <> "-") then
			result = beasongoptionname
		end if

		if (result <> "") then
			result = " - " + result
		end if

		BeasongOptionString = result
	end function

	public function BeasongPay()
		dim i
		for i=0 to FResultCount-1
			if FItemList(i).FItemID=0 then
				BeasongPay = FItemList(i).Fitemcost
				Exit For
			end if
		next
	end Function

	public function BeasongOptionStr()
		dim i
		for i=0 to FResultCount-1
			if FItemList(i).FItemID=0 then
				BeasongOptionStr = BeasongCD2Name(FItemList(i).Fitemoption)
				Exit For
			end if
		next
	end function

	public Sub QuickSearchOrderList()
		dim sqlStr, i
		dim addSql, tmporderserial

		addSql = ""

		if (FRectOrderSerial<>"") then
			addSql = addSql + " and m.orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectRegStart<>"") then
			tmporderserial = Mid(Replace(FRectRegStart, "-", ""), 3, 100) & "00000"
			addSql = addSql + " and m.orderserial >='" + CStr(tmporderserial) + "'"
			addSql = addSql + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			addSql = addSql + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectUserID<>"") then
			addSql = addSql + " and m.userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			addSql = addSql + " and m.buyname = '" + FRectBuyname + "'"  ''like
		end if

		if (FRectReqName<>"") then
			addSql = addSql + " and m.reqname = '" + FRectReqName + "'"  ''like
		end if

		if (FRectIpkumName<>"") then
			addSql = addSql + " and m.accountname = '" + FRectIpkumName + "'" ''like
		end if

		if (FRectSubTotalPrice<>"") then
			addSql = addSql + " and m.subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectBuyHp<>"") then
			addSql = addSql + " and m.buyhp='" + FRectBuyHp + "'"
		end if

		if (FRectReqHp<>"") then
			addSql = addSql + " and m.reqhp='" + FRectReqHp + "'"
		end if

		if (FRectBuyPhone<>"") then
			addSql = addSql + " and m.buyphone='" + FRectBuyPhone + "'"
		end if

		if (FRectReqPhone<>"") then
			addSql = addSql + " and m.reqphone='" + FRectReqPhone + "'"
		end if

		if (FRectReqSongjangNo<>"") then
			addSql = addSql + " and m.deliverno='" + FRectReqSongjangNo + "'"
		end if

		if (FRectIsFlower="Y") then
			addSql = addSql + " and m.cardribbon is Not NULL "
		end if

		if (FRectIsLecture="Y") then
			addSql = addSql + " and ((m.reqzipaddr='') or (m.reqzipaddr is NULL)) "
		end if

		if (FRectIsMinus="Y") then
			addSql = addSql + " and m.jumundiv='9' "
		end if

        if (FRectIsForeign<>"") then
            addSql = addSql + " and IsNULL(m.dlvcountryCode,'KR') not in ('KR', 'ZZ', 'QQ')"
        end if

        if (FRectIsForeignDirect<>"") then
            addSql = addSql + " and n.orderserial is not NULL "
        end if

        if (FRectIsQuick<>"") then
            addSql = addSql + " and IsNULL(m.dlvcountryCode,'KR') = 'QQ'"
        end if

		if (FRectExtSiteName<>"") then
			addSql = addSql + " and ((m.sitename='" + FRectExtSiteName + "') or (m.rdsite='" + FRectExtSiteName + "')) "
		end if

		if (FRectJumunItem <> "") and (FRectUserID <> "") then
			if IsNumeric(FRectJumunItem) then
				'// 상품코드
				addSql = addSql + " and ( "
				addSql = addSql + " 	select count(*) as cnt "
				addSql = addSql + " 	from "
				addSql = addSql + " 	[db_threepl].[dbo].tbl_order_detail d "
				addSql = addSql + " 	where m.orderserial = d.orderserial and d.itemid = " + CStr(FRectJumunItem) + " "
				addSql = addSql + " ) > 0 "
			else
				'// 상품명
				addSql = addSql + " and ( "
				addSql = addSql + " 	select count(*) as cnt "
				addSql = addSql + " 	from "
				addSql = addSql + " 	[db_threepl].[dbo].tbl_order_detail d "
				addSql = addSql + " 	where m.orderserial = d.orderserial and d.itemname like '%" + CStr(FRectJumunItem) + "%' "
				addSql = addSql + " ) > 0 "
			end if
		end if

		if (FRectSongjangno <> "") then
			addSql = addSql + " 	and ( "
			addSql = addSql + " 		select count(*) as cnt "
			addSql = addSql + " 		from "
			addSql = addSql + " 		[db_threepl].[dbo].tbl_order_detail d "
			addSql = addSql + " 		where m.orderserial = d.orderserial and replace(d.songjangno, '-', '') = '" & FRectSongjangno & "' "
			addSql = addSql + " 	) > 0 "
		end if


		'// ===================================================================
		''갯수
		sqlStr = "select count(*) as cnt "
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
		else
    		sqlStr = sqlStr + " from [db_threepl].[dbo].tbl_order_master m"
    	end if

        ''sqlStr = sqlStr + " left join db_order.dbo.tbl_change_order c "
        ''sqlStr = sqlStr + " on "
        ''sqlStr = sqlStr + " 	1 = 1 "
        ''sqlStr = sqlStr + " 	and m.orderserial = c.chgorderserial "
        ''sqlStr = sqlStr + " 	and c.deldate is null "

		if (FRectIsForeignDirect<>"") then
			''sqlStr = sqlStr + " left join [db_order].[dbo].[tbl_order_custom_number] n "
			''sqlStr = sqlStr + " on "
			''sqlStr = sqlStr + " 	m.orderserial = n.orderserial "
		end if

		sqlStr = sqlStr + " where 1=1"

		sqlStr = sqlStr + addSql


		rsget_TPL.Open sqlStr,dbget_TPL,1
			FTotalCount = rsget_TPL("cnt")
		rsget_TPL.close


		'// ===================================================================
		''데이타.
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.*, IsNull(m.sumPaymentEtc, 0) as sumPaymentEtc, IsNull(m.subtotalpriceCouponNotApplied, 0) as subtotalpriceCouponNotApplied  "
		sqlStr = sqlStr + " , p.partnercompanyname "
		sqlStr = sqlStr + " from [db_threepl].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " left join [db_threepl].[dbo].[tbl_partnerinfo] p on m.sitename = p.partnercompanyid "

		sqlStr = sqlStr + " where 1=1"

		sqlStr = sqlStr + addSql

		sqlStr = sqlStr + " order by m.idx desc"

		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr,dbget_TPL,1
		''rw sqlStr


		FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if not rsget_TPL.Eof then
			rsget_TPL.absolutepage = FCurrPage
			do until rsget_TPL.eof
				set FItemList(i) = new COrderMasterItem
				FItemList(i).Forderserial       = rsget_TPL("orderserial")
				FItemList(i).Fjumundiv	        = rsget_TPL("jumundiv")
				FItemList(i).Fuserid			= rsget_TPL("userid")
				FItemList(i).Faccountname		= db2Html(rsget_TPL("accountname"))
				FItemList(i).Faccountdiv		= trim(rsget_TPL("accountdiv"))
				FItemList(i).Faccountno	        = rsget_TPL("accountno")

				FItemList(i).Ftotalmileage      = rsget_TPL("totalmileage")
				FItemList(i).Ftotalsum	        = rsget_TPL("totalsum")
				FItemList(i).Fipkumdiv	        = rsget_TPL("ipkumdiv")
				FItemList(i).Fipkumdate	        = rsget_TPL("ipkumdate")
				FItemList(i).Fregdate			= rsget_TPL("regdate")
				FItemList(i).Fbaljudate			= rsget_TPL("baljudate")
				FItemList(i).Fbeadaldate		= rsget_TPL("beadaldate")
				FItemList(i).Fcancelyn	        = rsget_TPL("cancelyn")

				FItemList(i).Fbuyname			= db2Html(rsget_TPL("buyname"))
				FItemList(i).Fbuyphone	        = rsget_TPL("buyphone")
				FItemList(i).Fbuyhp				= rsget_TPL("buyhp")
				FItemList(i).Fbuyemail	        = rsget_TPL("buyemail")
				FItemList(i).Freqname			= db2Html(rsget_TPL("reqname"))

				FItemList(i).Freqzipcode		= rsget_TPL("reqzipcode")
				FItemList(i).Freqzipaddr		= db2Html(rsget_TPL("reqzipaddr"))
				FItemList(i).Freqaddress		= db2Html(rsget_TPL("reqaddress"))
				FItemList(i).Freqphone	        = rsget_TPL("reqphone")
				FItemList(i).Freqhp				= rsget_TPL("reqhp")
				FItemList(i).Freqemail	        = rsget_TPL("reqemail")
				FItemList(i).Fcomment			= db2Html(rsget_TPL("comment"))

				FItemList(i).Fdeliverno	        = rsget_TPL("deliverno")

				FItemList(i).Fsitename	        = rsget_TPL("sitename")
				FItemList(i).Fpaygatetid		= rsget_TPL("paygatetid")
				FItemList(i).Fdiscountrate		= rsget_TPL("discountrate")
				FItemList(i).Fsubtotalprice		= rsget_TPL("subtotalprice")
				FItemList(i).Fresultmsg			= rsget_TPL("resultmsg")
				FItemList(i).Frduserid			= rsget_TPL("rduserid")
				FItemList(i).Fmiletotalprice	= rsget_TPL("miletotalprice")
				if IsNULL(FItemList(i).Fmiletotalprice) then FItemList(i).Fmiletotalprice=0

				FItemList(i).Fauthcode		    = rsget_TPL("authcode")
				FItemList(i).Ftencardspend		= rsget_TPL("tencardspend")
				FItemList(i).Fuserlevel		    = rsget_TPL("userlevel")
				FItemList(i).Fspendmembership	= rsget_TPL("spendmembership")

                FItemList(i).Fallatdiscountprice = rsget_TPL("allatdiscountprice")

                FItemList(i).Freqdate    		= rsget_TPL("reqdate")
                FItemList(i).Freqtime    		= rsget_TPL("reqtime")
                FItemList(i).Fcardribbon 		= rsget_TPL("cardribbon")
                FItemList(i).Fmessage    		= rsget_TPL("message")
                FItemList(i).Ffromname   		= rsget_TPL("fromname")

                FItemList(i).FDlvcountryCode 	= rsget_TPL("DlvcountryCode")

                FItemList(i).FsumPaymentEtc 					= rsget_TPL("sumPaymentEtc")
                FItemList(i).FsubtotalpriceCouponNotApplied 	= rsget_TPL("subtotalpriceCouponNotApplied")

				FItemList(i).Frdsite			= rsget_TPL("rdsite")

                If isNull(rsget_TPL("userDisplayYn")) Then
                	FItemList(i).FuserDisplayYn	= "Y"
                Else
                	FItemList(i).FuserDisplayYn	= rsget_TPL("userDisplayYn")
                End If

                FItemList(i).Fpartnercompanyname		= db2Html(rsget_TPL("partnercompanyname"))

				rsget_TPL.movenext
				i=i+1
			loop
		end if
		rsget_TPL.Close
	end sub



	public Sub QuickSearchOrderMaster()
		dim sqlStr, i

		sqlStr = "select top 1 m.*, IsNull(m.sumPaymentEtc, 0) as sumPaymentEtc, IsNull(m.subtotalpriceCouponNotApplied, 0) as subtotalpriceCouponNotApplied "
		sqlStr = sqlStr + ", ( select sum(IsNULL(itemCostCouponNotApplied,0))  "
		sqlStr = sqlStr + "	    from [db_threepl].[dbo].tbl_order_detail  "
		sqlStr = sqlStr + "	    where orderserial=m.Orderserial "
		sqlStr = sqlStr + "	    and itemid=0  "
		sqlStr = sqlStr + "	    and cancelyn<>'Y' "
		sqlStr = sqlStr + "	) as deliverpriceCouponNotApplied "
		sqlStr = sqlStr + "	,(  select sum(itemcost)  "
		sqlStr = sqlStr + "	    from [db_threepl].[dbo].tbl_order_detail  "
		sqlStr = sqlStr + "	    where orderserial=m.Orderserial "
		sqlStr = sqlStr + "	    and itemid=0  "
		sqlStr = sqlStr + "	    and cancelyn<>'Y' "
		sqlStr = sqlStr + "	) as deliverprice"
		sqlStr = sqlStr + "	, IsNull(m.pggubun,'') as pggubun "
	    sqlStr = sqlStr + " from [db_threepl].[dbo].tbl_order_master m"
		sqlStr = sqlStr + " where m.idx<>0"

		if (FRectOrderSerial<>"") then
			sqlStr = sqlStr + " and m.orderserial='" + FRectOrderSerial + "'"
		end if

		if (FRectRegStart<>"") then
			sqlStr = sqlStr + " and m.regdate >='" + CStr(FRectRegStart) + "'"
		end if

		if (FRectRegEnd<>"") then
			sqlStr = sqlStr + " and m.regdate <'" + CStr(FRectRegEnd) + "'"
		end if

		if (FRectUserID<>"") then
			sqlStr = sqlStr + " and m.userid='" + FRectUserID + "'"
		end if

		if (FRectBuyname<>"") then
			sqlStr = sqlStr + " and m.buyname = '" + FRectBuyname + "'"  ''like
		end if

		if (FRectReqName<>"") then
			sqlStr = sqlStr + " and m.reqname = '" + FRectReqName + "'" ''like
		end if

		if (FRectIpkumName<>"") then
			sqlStr = sqlStr + " and m.accountname ='" + FRectIpkumName + "'" ''like
		end if

		if (FRectSubTotalPrice<>"") then
			sqlStr = sqlStr + " and m.subtotalprice =" + CStr(FRectSubTotalPrice) + ""
		end if

		if (FRectBuyHp<>"") then
			sqlStr = sqlStr + " and m.buyhp='" + FRectBuyHp + "'"
		end if

		if (FRectReqHp<>"") then
			sqlStr = sqlStr + " and m.reqhp='" + FRectReqHp + "'"
		end if

		if (FRectBuyPhone<>"") then
			sqlStr = sqlStr + " and m.buyphone='" + FRectBuyPhone + "'"
		end if

		if (FRectReqPhone<>"") then
			sqlStr = sqlStr + " and m.reqphone='" + FRectReqPhone + "'"
		end if

		if (FRectReqSongjangNo<>"") then
			sqlStr = sqlStr + " and m.deliverno='" + FRectReqSongjangNo + "'"
		end if

		sqlStr = sqlStr + " order by m.orderserial desc"
        ''sqlStr = sqlStr + " order by idx desc"

		''response.write sqlStr

		rsget_TPL.pagesize = FPageSize
		rsget_TPL.Open sqlStr,dbget_TPL,1

		if not rsget_TPL.Eof then
		        FTotalCount = 1
		end if

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget_TPL.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

        if not rsget_TPL.Eof then
	        set FOneItem = new COrderMasterItem

			FOneItem.Fidx		           	= rsget_TPL("idx")
			FOneItem.Forderserial           = rsget_TPL("orderserial")
			FOneItem.Fjumundiv	            = rsget_TPL("jumundiv")
			FOneItem.Fuserid		        = rsget_TPL("userid")
			FOneItem.Faccountname	        = db2Html(rsget_TPL("accountname"))
			FOneItem.Faccountdiv	        = trim(rsget_TPL("accountdiv"))
			FOneItem.Faccountno	            = rsget_TPL("accountno")

			FOneItem.Ftotalmileage          = rsget_TPL("totalmileage")
			FOneItem.Ftotalsum	            = rsget_TPL("totalsum")
			FOneItem.Fipkumdiv	            = rsget_TPL("ipkumdiv")
			FOneItem.Fipkumdate	            = rsget_TPL("ipkumdate")
			FOneItem.Fregdate		        = rsget_TPL("regdate")
			FOneItem.Fbaljudate		        = rsget_TPL("baljudate")
			FOneItem.Fbeadaldate	        = rsget_TPL("beadaldate")
			FOneItem.Fcancelyn	            = rsget_TPL("cancelyn")
			FOneItem.Fbuyname		        = db2Html(rsget_TPL("buyname"))
			FOneItem.Fbuyphone	            = rsget_TPL("buyphone")
			FOneItem.Fbuyhp		            = rsget_TPL("buyhp")
			FOneItem.Fbuyemail	            = rsget_TPL("buyemail")
			FOneItem.Freqname		        = db2Html(rsget_TPL("reqname"))
			FOneItem.Freqzipcode	        = rsget_TPL("reqzipcode")
			FOneItem.Freqaddress	        = db2Html(rsget_TPL("reqaddress"))
			FOneItem.Freqphone	            = rsget_TPL("reqphone")
			FOneItem.Freqhp		            = rsget_TPL("reqhp")
			FOneItem.Freqemail	            = rsget_TPL("reqemail")
			FOneItem.Fcomment		        = db2Html(rsget_TPL("comment"))
			FOneItem.Fdeliverno	            = rsget_TPL("deliverno")
			FOneItem.Fsitename	            = rsget_TPL("sitename")
			FOneItem.Fpaygatetid	        = rsget_TPL("paygatetid")
			FOneItem.Fdiscountrate	        = rsget_TPL("discountrate")
			FOneItem.Fsubtotalprice	        = rsget_TPL("subtotalprice")
			FOneItem.Fresultmsg		        = rsget_TPL("resultmsg")
			FOneItem.Frduserid		        = rsget_TPL("rduserid")
			FOneItem.Fmiletotalprice	    = rsget_TPL("miletotalprice")

			FOneItem.FInsureCd           	= rsget_TPL("InsureCd")

			if IsNULL(FOneItem.Fmiletotalprice) then FOneItem.Fmiletotalprice=0

			FOneItem.Fjungsanflag		    = rsget_TPL("jungsanflag")
			FOneItem.Freqzipaddr		    = db2Html(rsget_TPL("reqzipaddr"))
			FOneItem.Fauthcode		        = rsget_TPL("authcode")
			FOneItem.Fcashreceiptreq		= rsget_TPL("cashreceiptreq")

			FOneItem.Ftencardspend		    = rsget_TPL("tencardspend")
			FOneItem.FbCpnIdx		    	= rsget_TPL("bCpnIdx")

			FOneItem.Fuserlevel		        = rsget_TPL("userlevel")
			FOneItem.Fspendmembership	    = rsget_TPL("spendmembership")
			FOneItem.Fallatdiscountprice    = rsget_TPL("allatdiscountprice")

			FOneItem.Freqdate    = rsget_TPL("reqdate")
            FOneItem.Freqtime    = rsget_TPL("reqtime")
            FOneItem.Fcardribbon = rsget_TPL("cardribbon")
            FOneItem.Fmessage    = rsget_TPL("message")
            FOneItem.Ffromname   = rsget_TPL("fromname")

            FOneItem.FDlvcountryCode = rsget_TPL("DlvcountryCode")
            FOneItem.Frdsite	= rsget_TPL("rdsite")

            FOneItem.FsumPaymentEtc 					= rsget_TPL("sumPaymentEtc")
            FOneItem.FsubtotalpriceCouponNotApplied 	= rsget_TPL("subtotalpriceCouponNotApplied")

			FOneItem.FDeliverpriceCouponNotApplied = rsget_TPL("deliverpriceCouponNotApplied")
			FOneItem.Fdeliverprice = rsget_TPL("deliverprice")

			FOneItem.Fpggubun 			= rsget_TPL("pggubun")
    		FOneItem.Fordersheetyn 		= rsget_TPL("ordersheetyn")
	    end if
		rsget_TPL.Close
	end sub

	public Sub QuickSearchOrderDetail()
		dim sqlStr
		dim i

		'orgitemcost 				: 소비자가
		'itemcostCouponNotApplied 	: 판매가(할인가)
		'itemcost 					: 상품쿠폰/플러스세일할인/우수고객할인 적용된 금액
		'reducedPrice 				: 보너스쿠폰적용가+기타할인적용가
		'plusSaleDiscount 			: 플러스세일할인액
		'specialshopDiscount 		: 우수고객할인액
		'etcDiscount				: 기타할인(하나카드 할인 등)

		'orgsuplycash				: 원매입가
		'buycashCouponNotApplied	: 할인매입가
		'buycash					: 쿠폰적용매입가

		sqlStr = "select d.idx, d.orderserial,d.prdcode, d.itemid,d.itemoption,d.itemno,d.itemcost,d.reducedPrice,d.buycash, d.oitemdiv "
		sqlStr = sqlStr + " ,d.mileage,d.cancelyn "
		sqlStr = sqlStr + " ,d.itemname, d.makerid, i.mainimageurl "
		sqlStr = sqlStr + " , 0 as orgprice, 0 as orgsuplycash, 0 as buycashCouponNotApplied, 0 as optionaddprice, d.itemoptionname "
		sqlStr = sqlStr + " ,d.currstate, d.upcheconfirmdate, d.songjangdiv, d.songjangno"
		sqlStr = sqlStr + " ,d.beasongdate, d.isupchebeasong, d.issailitem, d.requiredetail  "
		sqlStr = sqlStr + " ,d.issailitem, d.bonuscouponidx, d.itemcouponidx"
		sqlStr = sqlStr + " ,s.divname as songjangdivname, s.findurl"
		sqlStr = sqlStr + " , IsNull(d.orgitemcost, 0) as orgitemcost "
		sqlStr = sqlStr + " , IsNull(d.itemcostCouponNotApplied, 0) as itemcostCouponNotApplied "
		sqlStr = sqlStr + " , IsNull(d.plusSaleDiscount, 0) as plusSaleDiscount "
		sqlStr = sqlStr + " , IsNull(d.specialshopDiscount, 0) as specialshopDiscount "
		sqlStr = sqlStr + " , IsNull(d.etcDiscount, 0) as etcDiscount "
		sqlStr = sqlStr + " , d.odlvType, d.odlvfixday "
	    sqlStr = sqlStr + " from [db_threepl].[dbo].tbl_order_detail d "
		sqlStr = sqlStr + "     left join [db_threepl].[dbo].tbl_item i on d.prdcode=i.prdcode"
		sqlStr = sqlStr + "     left join [db_threepl].[dbo].tbl_songjang_div s on d.songjangdiv=s.divcd"
		sqlStr = sqlStr + " where d.orderserial='" + CStr(FRectOrderSerial) + "'"
        sqlStr = sqlStr + " order by d.isupchebeasong, d.makerid, d.itemid, d.itemoption"

        'response.write sqlStr &"<br>"
		rsget_TPL.Open sqlStr,dbget_TPL,1

		FTotalCount = rsget_TPL.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		FTotItemKind = 0
		FTotItemNo = 0
		do until rsget_TPL.eof
			set FItemList(i) = new COrderDetailItem

			FItemList(i).Forderserial = CStr(FRectOrderSerial)
			FItemList(i).Fidx         = rsget_TPL("idx")
			FItemList(i).Fmakerid     = rsget_TPL("makerid")
			FItemList(i).Fprdcode     = rsget_TPL("prdcode")
			FItemList(i).Fitemid      = rsget_TPL("itemid")
			FItemList(i).Fitemoption  = rsget_TPL("itemoption")
			FItemList(i).Fitemno      = rsget_TPL("itemno")
			FItemList(i).Fitemcost    = rsget_TPL("itemcost")
			FItemList(i).Fmileage     = rsget_TPL("mileage")
			FItemList(i).Fcancelyn    = rsget_TPL("cancelyn")

			FItemList(i).Forgsuplycash     			= rsget_TPL("orgsuplycash")
			FItemList(i).FbuycashCouponNotApplied   = rsget_TPL("buycashCouponNotApplied")
			FItemList(i).Fbuycash     				= rsget_TPL("buycash")

			FItemList(i).FItemName    = db2html(rsget_TPL("itemname"))

			FItemList(i).FSmallImage  = rsget_TPL("mainimageurl")

			if IsNull(rsget_TPL("itemoptionname")) then
				FItemList(i).FItemoptionName = "-"
			else
				FItemList(i).FItemoptionName = db2html(rsget_TPL("itemoptionname"))
			end if

			FItemList(i).Fcurrstate         = rsget_TPL("currstate")
			FItemList(i).Fsongjangdiv       = rsget_TPL("songjangdiv")
			FItemList(i).Fsongjangno        = rsget_TPL("songjangno")
			FItemList(i).Fbeasongdate       = rsget_TPL("beasongdate")
			FItemList(i).Fisupchebeasong    = rsget_TPL("isupchebeasong")
			FItemList(i).Fissailitem        = rsget_TPL("issailitem")
			FItemList(i).Fupcheconfirmdate    = rsget_TPL("upcheconfirmdate")

			FItemList(i).Frequiredetail    = rsget_TPL("requiredetail")
            FItemList(i).Fsongjangdivname  = db2html(rsget_TPL("songjangdivname"))
            FItemList(i).Ffindurl          = db2html(rsget_TPL("findurl"))

            FItemList(i).Forgprice          = rsget_TPL("orgprice")
            FItemList(i).Fissailitem        = rsget_TPL("issailitem")
            FItemList(i).Fbonuscouponidx    = rsget_TPL("bonuscouponidx")
            FItemList(i).Fitemcouponidx     = rsget_TPL("itemcouponidx")
            FItemList(i).FreducedPrice      = rsget_TPL("reducedPrice")

            FItemList(i).Forgitemcost      			= rsget_TPL("orgitemcost")
            FItemList(i).FitemcostCouponNotApplied  = rsget_TPL("itemcostCouponNotApplied")
            FItemList(i).FplusSaleDiscount      	= rsget_TPL("plusSaleDiscount")
            FItemList(i).FspecialshopDiscount      	= rsget_TPL("specialshopDiscount")
			FItemList(i).FetcDiscount		      	= rsget_TPL("etcDiscount")
			FItemList(i).Foitemdiv			      	= rsget_TPL("oitemdiv")
			FItemList(i).FodlvType			      	= rsget_TPL("odlvType")
			FItemList(i).fodlvfixday			      	= rsget_TPL("odlvfixday")

            if Not IsNULL(FItemList(i).Fsongjangno) then
               FItemList(i).Fsongjangno = replace(FItemList(i).Fsongjangno,"-","")
            end if

			IF FItemList(i).Fitemid <> 0 THEN
				FTotItemNo = FTotItemNo + FItemList(i).Fitemno
				FTotItemKind = FTotItemKind + 1
			END IF
			rsget_TPL.movenext
			i=i+1
		loop
		rsget_TPL.close
	end sub

    public function GetOneOrderDetail
        dim sqlStr, i
	    dim mastertable, detailtable

	    if (FRectOldOrder<>"") then
			mastertable = "[db_log].[dbo].tbl_old_order_master_2003"
			detailtable	= "[db_log].[dbo].tbl_old_order_detail_2003"
		else
			mastertable = "[db_order].[dbo].tbl_order_master"
			detailtable	= "[db_order].[dbo].tbl_order_detail"
		end if

		sqlStr =	" SELECT d.idx, d.itemid, d.itemoption, d.itemno, d.itemoptionname, d.itemcost," &_
					" d.itemname, d.itemcost, d.makerid, d.currstate, replace(d.songjangno,'-','') as songjangno, d.songjangdiv," &_
					" d.cancelyn, d.isupchebeasong, d.mileage, Replace(d.requiredetail, '，', ',') as requiredetail, d.oitemdiv, d.beasongdate, d.issailitem, d.upcheconfirmdate," &_
					" d.bonuscouponidx, d.itemcouponidx, d.reducedPrice," &_
					" i.smallimage, i.listimage, i.brandname, i.itemdiv, i.orgprice" &_
					" ,s.divname,s.findurl ,s.tel as DeliveryTel" &_
					" FROM " + detailtable + " d " &_
					" JOIN [db_item].[dbo].tbl_item i" &_
					"		ON d.itemid=i.itemid " &_
					" LEFT JOIN db_order.[dbo].tbl_songjang_div s " &_
					"		ON d.songjangdiv = s.divcd " &_
					" WHERE d.orderserial='" + FRectOrderserial + "'" &_
					" and d.idx=" & FRectDetailIdx &_
					" and d.itemid<>0" &_
					" and d.cancelyn<>'Y'" &_
					" order by i.deliverytype"
		rsget.Open sqlStr,dbget,1

		FTotalcount = rsget.Recordcount
		FResultcount = FTotalcount


        if Not rsget.Eof then
			set FOneItem = new COrderDetailItem
			FOneItem.Forderserial = CStr(FRectOrderSerial)
			FOneItem.Fidx         = rsget("idx")
			FOneItem.Fmakerid     = rsget("makerid")
			FOneItem.Fitemid      = rsget("itemid")
			FOneItem.Fitemoption  = rsget("itemoption")
			FOneItem.Fitemno      = rsget("itemno")
			FOneItem.Fitemcost    = rsget("itemcost")
			FOneItem.Fmileage     = rsget("mileage")
			FOneItem.Fcancelyn    = rsget("cancelyn")

			FOneItem.FItemName    = db2html(rsget("itemname"))
			FOneItem.FSmallImage  = webImgUrl + "/image/small/" + GetImageSubFolderByItemID(FOneItem.Fitemid) + "/" + rsget("smallimage")

			if IsNull(rsget("itemoptionname")) then
				FOneItem.FItemoptionName = "-"
			else
				FOneItem.FItemoptionName = db2html(rsget("itemoptionname"))
			end if

			FOneItem.Fcurrstate         = rsget("currstate")
			FOneItem.Fsongjangdiv       = rsget("songjangdiv")
			FOneItem.Fsongjangno        = rsget("songjangno")
			FOneItem.Fbeasongdate       = rsget("beasongdate")
			FOneItem.Fisupchebeasong    = rsget("isupchebeasong")
			FOneItem.Fissailitem        = rsget("issailitem")
			FOneItem.Fupcheconfirmdate    = rsget("upcheconfirmdate")

			FOneItem.Frequiredetail    = rsget("requiredetail")
            FOneItem.Fsongjangdivname  = db2html(rsget("divname"))
            FOneItem.Ffindurl          = db2html(rsget("findurl"))

            FOneItem.Forgprice          = rsget("orgprice")
            FOneItem.Fissailitem        = rsget("issailitem")
            FOneItem.Fbonuscouponidx    = rsget("bonuscouponidx")
            FOneItem.Fitemcouponidx     = rsget("itemcouponidx")

            FOneItem.FreducedPrice      = rsget("reducedPrice")
            if Not IsNULL(FOneItem.Fsongjangno) then
               FOneItem.Fsongjangno = replace(FOneItem.Fsongjangno,"-","")
            end if

		end if
		rsget.close
    end function

    public function getOrderItemSummary()
        dim sqlStr
		sqlStr = " select "
		sqlStr = sqlStr + "	sum(case when isupchebeasong <> 'Y' then itemno else 0 end) as tenbeacnt "
		sqlStr = sqlStr + "		, sum(case when isupchebeasong = 'Y' then itemno else 0 end) as upbeacnt "
		sqlStr = sqlStr + "		, count(distinct (case when isupchebeasong = 'Y' then makerid else '' end)) as brandcnt "
	    sqlStr = sqlStr + " from [db_threepl].[dbo].tbl_order_detail m"
		sqlStr = sqlStr + "	where orderserial = '" + CStr(FRectOrderserial) + "' and itemid <> 0 and cancelyn <> 'Y' "
		rsget_TPL.Open sqlStr,dbget_TPL,1

		set FOneItem = new COrderItemSummaryItem

		if Not rsget_TPL.Eof then
			FOneItem.Ftenbeacnt   = rsget_TPL("tenbeacnt")
			FOneItem.Fupbeacnt   = rsget_TPL("upbeacnt")
			FOneItem.Fbrandcnt   = rsget_TPL("brandcnt")

			if (FOneItem.Ftenbeacnt > 0) then
				FOneItem.Fbrandcnt = FOneItem.Fbrandcnt - 1
			end if
		end if
		rsget_TPL.Close
    end function

    public function getEmsOrderInfo()
        dim sqlStr
        sqlStr = " exec [db_order].[dbo].sp_Ten_OneEmsOrderInfo '" & FRectOrderserial & "'"

        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
            FOneItem.FcountryNameEn   = rsget("countryNameEn")
            FOneItem.FemsAreaCode     = rsget("emsAreaCode")
            FOneItem.FemsZipCode      = rsget("emsZipCode")
            FOneItem.FitemGubunName   = rsget("itemGubunName")
            FOneItem.FgoodNames       = rsget("goodNames")
            FOneItem.FitemWeigth      = rsget("itemWeigth")
            FOneItem.FitemUsDollar    = rsget("itemUsDollar")
            FOneItem.FemsInsureYn     = rsget("InsureYn")
            FOneItem.FemsInsurePrice  = rsget("InsurePrice")

            FOneItem.FemsDlvCost       = rsget("emsDlvCost")
		end if
		rsget.Close
    end function

	public Sub getEtcPaymentList()
		dim sqlStr
		dim i

		sqlStr = " select e.*, d.divnm as acctdivName "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc e "
		sqlStr = sqlStr + " 	left join db_order.dbo.tbl_account_div d "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		e.acctdiv = d.divcd "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and e.orderserial = '" & FRectOrderserial & "' "
		if (FRectIncMainPayment <> "Y") then
			sqlStr = sqlStr + " 	and e.acctdiv not in ('7', '100', '550', '560', '20', '50', '80', '90', '400', '110', '120') "							'OK CASH BAG 은 주결제수단이다. 120=네이버포인트
		end if

        'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CEtcPaymentItem

			FItemList(i).Facctdiv     		= rsget("acctdiv")
			FItemList(i).FacctdivName     	= rsget("acctdivName")
			FItemList(i).Facctamount     	= rsget("acctamount")
			FItemList(i).FrealPayedsum     	= rsget("realPayedsum")
			FItemList(i).FacctAuthCode     	= rsget("acctAuthCode")
			FItemList(i).FacctAuthDate     	= rsget("acctAuthDate")

			rsget.movenext
			i = i + 1
		loop
		rsget.close
	end sub

	'최초 주결제금액(+신용카드 취소관련 정보)
	public Sub getMainPaymentInfo(byval paymethod, byref orgpayment, byref cardcancelok, byref cardcancelerrormsg, byref cardcancelcount, byref cardcancelsum, byref cardcode)
		dim sqlStr

		dim remailpayment, payetcresult
		dim jumundiv, orgorderserial, pggubun
		dim tmpArr

		orgpayment = 0
		cardcancelok = "N"
		cardcancelerrormsg = ""
		cardcancelcount = ""
		cardcode = ""

		'// 교환주문( jumundiv = 6 )이면 원주문에서 결제정보 가져온다.
		sqlStr = " select top 1 m.jumundiv, m.pggubun "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_order.dbo.tbl_order_master m "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.orderserial = '" & FRectOrderserial & "' "
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			jumundiv = rsget("jumundiv")
			pggubun  = rsget("pggubun")
		end if
		rsget.close

		if (jumundiv = "6") then
			sqlStr = " select top 1 c.orgorderserial "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_order.dbo.tbl_change_order c "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and c.chgorderserial = '" & FRectOrderserial & "' "
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				orgorderserial = rsget("orgorderserial")
			end if
			rsget.close
		else
			orgorderserial = FRectOrderserial
		end if

		sqlStr = " select top 1 IsNull(e.acctamount, 0) as orgpayment, IsNull(e.realPayedSum, 0) as remailpayment, IsNull(e.PayEtcResult, '') as payetcresult "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc e "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and e.orderserial = '" & orgorderserial & "' "
		sqlStr = sqlStr + " 	and e.acctdiv in ('7', '100', '550', '560', '20', '50', '80', '90', '400', '110') "							'OK CASH BAG 은 주결제수단이다.

        'response.write sqlStr &"<br>"
        IF (paymethod="110") then
            sqlStr = " select sum(IsNull(e.acctamount, 0)) as orgpayment, sum(IsNull(e.realPayedSum, 0)) as remailpayment, '' as payetcresult "
    		sqlStr = sqlStr + " from "
    		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc e "
    		sqlStr = sqlStr + " where "
    		sqlStr = sqlStr + " 	1 = 1 "
    		sqlStr = sqlStr + " 	and e.orderserial = '" & orgorderserial & "' "
    		sqlStr = sqlStr + " 	and e.acctdiv in ('100', '110') "
        END IF

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			orgpayment = rsget("orgpayment")
			remailpayment = rsget("remailpayment")
			payetcresult = rsget("payetcresult")

			if Len(payetcresult) = 9 and UBound(Split(payetcresult, "|")) = 3 then
				'// 14|26|0|1 => 14|26|00|1
				tmpArr = Split(payetcresult, "|")
				payetcresult = tmpArr(0) & "|" & tmpArr(1) & "|" & "0" & tmpArr(2) & "|" & tmpArr(3)
			end If

			'// 페이코
			if Len(payetcresult) = 6 and UBound(Split(payetcresult, "|")) = 3 then
				'// ||00|1 => XX|XX|00|1
				tmpArr = Split(payetcresult, "|")
				payetcresult = "XX" & "|" & "XX" & "|" & tmpArr(2) & "|" & tmpArr(3)
			end if
		end if
		rsget.Close

        '' 네이버 페이 관련 추가 (포인트)
        if (pggubun="NP") or (pggubun="PY") then
            sqlStr = " select top 1 IsNull(e.acctamount, 0) as orgpayment, IsNull(e.realPayedSum, 0) as remailpayment, IsNull(e.PayEtcResult, '') as payetcresult "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc e "
            sqlStr = sqlStr + " where "
            sqlStr = sqlStr + " 	1 = 1 "
            sqlStr = sqlStr + " 	and e.orderserial = '" & orgorderserial & "' "
            sqlStr = sqlStr + " 	and e.acctdiv='120'"

            rsget.Open sqlStr,dbget,1
            if Not rsget.Eof then
            	orgpayment = orgpayment + rsget("orgpayment")
            	remailpayment = remailpayment + rsget("remailpayment")

            	if Len(payetcresult) = 7 and UBound(Split(payetcresult, "|")) = 3 then
            		'// 14||0|1 => 14|26|00|1
            		tmpArr = Split(payetcresult, "|")
            		payetcresult = tmpArr(0) & "|" & "XX" & "|" & "0" & tmpArr(2) & "|" & tmpArr(3)
            	end If
            end if
            rsget.close

        end if

		if (paymethod <> "100") then
			if (paymethod = "110") then
				cardcancelerrormsg = "OK+신용(결제 부분취소불가)"
			elseif (paymethod = "20") and (pggubun="NP") then                              ''2016/07/21 추가
			    cardcancelok = "Y"
				cardcancelcount = 0
				cardcode = payetcresult
			else
				cardcancelerrormsg = "신용카드결제 아님"
			end if
		else
			if (orgpayment = 0) or (payetcresult = "") then
				cardcancelerrormsg = "신용카드정보 누락"
			else
				cardcancelok = "Y"
				cardcancelcount = 0
				cardcode = payetcresult
			end if
		end if

        cardcancelcount = 0
        cardcancelsum   = 0
		if (cardcancelok = "Y") and (orgpayment <> remailpayment) then
			sqlStr = " select count(orderserial) as cnt, isNULL(sum(cancelprice),0) as canceltotal "  ''2017/07/10 sum(cancelprice) =>isNULL(sum(cancelprice),0)
			sqlStr = sqlStr + " from db_order.dbo.tbl_card_cancel_log "
			sqlStr = sqlStr + " where orderserial = '" & orgorderserial & "' and resultcode in ('00', '2001') "  '''0000' 다시 제거 2016/07/21 eastone 코드 '00' 으로 바꿈
			rsget.Open sqlStr,dbget,1

			if Not rsget.Eof then
				cardcancelcount = rsget("cnt")
				cardcancelsum   = rsget("canceltotal")
			end if
			rsget.close

			'9회까지 부분취소가 가능하지만 만약을 위한 1번은 남겨놓는다.
			if (cardcancelcount >= 9) and (FRectOrderSerial <> "18080316199") then
				cardcancelok = "N"
				cardcancelerrormsg = "부분취소 횟수 초과"
			end if
		end if

		if (cardcancelok = "Y") then
		    if (LEN(TRIM(cardcode))=10) then
                if (Right(cardcode,1)="1") then
                    ''cardcancelok = "Y"
                elseif (Right(cardcode,1)="0") then
                    cardcancelok = "N"
                    if (cardcancelerrormsg="") then cardcancelerrormsg  = "부분취소 <strong>불가</strong> 거래 (충전식 카드 or 복합거래)"
                end if
            end if

''          cardcode 맨 끝자리로 확인 가능.
'			if (InStr("11|00,06|04,12|00,14|26,01|05,04|00,03|00,16|11,17|81", Left(cardcode, 5)) <= 0) then
'				cardcancelok = "N"
'				cardcancelerrormsg = "부분취소 불가카드"
'
'				if (InStr("06,14,01", Left(cardcode, 2)) > 0) then
'					cardcancelerrormsg = "국민/신한/외환카드의 계열사카드는 부분취소 불가"
'				end if
'			end if
		end if

	end sub

	'최초 주결제금액(+ 휴대폰 취소관련 정보)
	public Sub getMainPaymentInfoPhone(byval paymethod, byref orgpayment, byref phonecancelok, byref phonecancelerrormsg, byref phonecancelcount, byref phonecancelsum, byref phonecode)
		dim sqlStr

		dim remailpayment, payetcresult
		dim jumundiv, orgorderserial

		orgpayment = 0
		phonecancelok = "N"
		phonecancelerrormsg = ""
		phonecancelcount = ""
		phonecode = ""

		'// 교환주문( jumundiv = 6 )이면 원주문에서 결제정보 가져온다.
		sqlStr = " select top 1 m.jumundiv "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_order.dbo.tbl_order_master m "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.orderserial = '" & FRectOrderserial & "' "
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			jumundiv = rsget("jumundiv")
		end if
		rsget.close

		if (jumundiv = "6") then
			sqlStr = " select top 1 c.orgorderserial "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_order.dbo.tbl_change_order c "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and c.chgorderserial = '" & FRectOrderserial & "' "
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				orgorderserial = rsget("orgorderserial")
			end if
			rsget.close
		else
			orgorderserial = FRectOrderserial
		end if

		sqlStr = " select top 1 IsNull(e.acctamount, 0) as orgpayment, IsNull(e.realPayedSum, 0) as remailpayment, IsNull(e.PayEtcResult, '') as payetcresult "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc e "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and e.orderserial = '" & orgorderserial & "' "
		sqlStr = sqlStr + " 	and e.acctdiv in ('7', '100', '550', '560', '20', '50', '80', '90', '400', '110') "							'OK CASH BAG 은 주결제수단이다.

        'response.write sqlStr &"<br>"
        IF (paymethod="110") then
            sqlStr = " select sum(IsNull(e.acctamount, 0)) as orgpayment, sum(IsNull(e.realPayedSum, 0)) as remailpayment, '' as payetcresult "
    		sqlStr = sqlStr + " from "
    		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_PaymentEtc e "
    		sqlStr = sqlStr + " where "
    		sqlStr = sqlStr + " 	1 = 1 "
    		sqlStr = sqlStr + " 	and e.orderserial = '" & orgorderserial & "' "
    		sqlStr = sqlStr + " 	and e.acctdiv in ('100', '110') "
        END IF

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			orgpayment = rsget("orgpayment")
			remailpayment = rsget("remailpayment")
			payetcresult = rsget("payetcresult")
		end if
		rsget.close

		if (paymethod <> "400") then
			phonecancelerrormsg = "휴대폰결제 아님"
		else
			if (orgpayment = 0) then
				phonecancelerrormsg = "휴대폰결제정보 누락"
			else
				phonecancelok = "Y"
				phonecancelcount = 0
				phonecode = payetcresult
			end if
		end if

        phonecancelcount = 0
        phonecancelsum   = 0
		if (phonecancelok = "Y") and (orgpayment <> remailpayment) then
			sqlStr = " select count(orderserial) as cnt, sum(cancelprice) as canceltotal "
			sqlStr = sqlStr + " from db_order.dbo.tbl_card_cancel_log "
			sqlStr = sqlStr + " where orderserial = '" & orgorderserial & "' and resultcode = '0000' "
			rsget.Open sqlStr,dbget,1

			if Not rsget.Eof then
				phonecancelcount = rsget("cnt")
				phonecancelsum   = rsget("canceltotal")
			end if
			rsget.close
		end if

	end sub

	public Sub getUpcheBeasongPayList()
		dim sqlStr
		dim i

		sqlStr = " select distinct "
		sqlStr = sqlStr + " 	d.makerid, IsNull(b.defaultfreebeasonglimit, 0) as defaultfreebeasonglimit, IsNull(b.defaultdeliverpay, 0) as defaultdeliverpay "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d "
		sqlStr = sqlStr + " 	join db_user.dbo.tbl_user_c b "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		d.makerid = b.userid "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and d.orderserial = '" & FRectOrderserial & "' "
		sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 	and d.isupchebeasong <> 'N' "

        'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CUpcheBeasongPayItem

			FItemList(i).Fmakerid     					= rsget("makerid")
			FItemList(i).Fdefaultfreebeasonglimit     	= rsget("defaultfreebeasonglimit")
			FItemList(i).Fdefaultdeliverpay     		= rsget("defaultdeliverpay")

			if (FItemList(i).Fdefaultdeliverpay = 0) then
				'기본배송비 설정 않되어 있으면 2500원(since 2012-06-18)
				FItemList(i).Fdefaultdeliverpay = 2500
			end if

			rsget.movenext
			i = i + 1
		loop
		rsget.close
	end sub

	public Sub getUpcheBeasongMakerList()
		dim sqlStr
		dim i

		''10x10logistics : 물류센터
		sqlStr = " select distinct "
		sqlStr = sqlStr + " 	(case when d.isupchebeasong = 'N' then '10x10logistics' else d.makerid end) as makerid"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and d.orderserial = '" & FRectOrderserial & "' "
		sqlStr = sqlStr + " 	and d.itemid not in (0, 100) "
		sqlStr = sqlStr + " order by (case when d.isupchebeasong = 'N' then '10x10logistics' else d.makerid end) "
        ''response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CUpcheBeasongPayItem

			FItemList(i).Fmakerid     					= rsget("makerid")

			rsget.movenext
			i = i + 1
		loop
		rsget.close
	end sub

	Private Sub Class_Initialize()

		Redim FItemList(0)

		FCurrPage =1
		FPageSize = 20
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
