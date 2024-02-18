<%
'####################################################
' Description :  오프라인 매출 클래스
' History : 2009.04.07 서동석 생성
'			2010.04.28 한용민 수정
'####################################################

Class COffShopJaeGo
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Flastrealjeago
	public FIpChulNo
	public FSellNo
	public FJaeGo
	public Ftotipchulno
	public Ftotsellno
	public FMakerID
	public FInputJaeGo

	public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (FItemId >= 1000000) then
			getBarCode = CStr(Fitemgubun) + Format00(8,FItemId) + Fitemoption
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COffShopSellByTerm
	public Fregdate
	public fitemcnt
	public fIorgsellprice
	public Fidx
	public FItemName
	public FTerm
	public FCount
	public FSum
	public FSpendMile
	public frealsum
	public ftotalsum
	public FGainMile
	public FShopid
	public FShopName
	public FMakerid
	public FIsBrandShop
	public FJungsanID
	public FSelljungsanID
	public Fsitename
	public Fselltotal
	public Fsellcnt
	public Fdpart
	public Faccountdiv
	public maxt
	public maxc
	public FItemGubun
	public FItemNo
	public FItemID
	public FItemOption
	public FItemCost
	public FItemOptionStr
	public FBuycash
	public FCancelyn
	public FYYYYMMDDHHNNSS
	public FChargeDiv
	public FpurchaseType
    public FCashSum
    public FCardSum
    public FgiftcardPaysum
	public FextPaysum
    public Fcashcnt
    public Fcardcnt
    public Fgiftcardcnt
	public fsuplyprice
	public fprofit
	public fmagin
	public Frealsellprice
	public fTenGiftCardPaySum
	public fTenGiftCardPaycount
	public faddtaxcharge
	public FaddTaxChargeSum
	public fIXyyyymmdd
	public fsellprice
	public ftargetmaechul
	public fgpart
	public fz1_in
	public fz1_out
	public fz1_all
	public fz2_in
	public fz2_out
	public fz2_all
	public FWeather
	public fpurchasetypename

	public function IsBrandShop()
		''사용안함.
		IsBrandShop = false
		if FIsBrandShop<>"" then IsBrandShop = true
	end function

	public function GetDpartName()
		if Fdpart=1 then
			GetDpartName = "<font color=#FF0000>일</font>"
		elseif Fdpart=2 then
			GetDpartName = "월"
		elseif Fdpart=3 then
			GetDpartName = "화"
		elseif Fdpart=4 then
			GetDpartName = "수"
		elseif Fdpart=5 then
			GetDpartName = "목"
		elseif Fdpart=6 then
			GetDpartName = "금"
		elseif Fdpart=7 then
			GetDpartName = "<font color=#0000FF>토</font>"
		else
			GetDpartName = ""
		end if
	end function

	Public function JumunMethodName()
		if Cstr(Faccountdiv) = "01" then
			JumunMethodName = "현금"
		elseif Cstr(Faccountdiv) = "02" then
			JumunMethodName = "카드"
		end if
	end function

	public function getChargeDivName()
		if FChargeDiv="2" then
			getChargeDivName = "출고특정" '10x10 특정
		elseif FChargeDiv="4" then
			getChargeDivName = "출고매입" '10x10매입
		elseif FChargeDiv="5" then
			getChargeDivName = "출고매입" '출고분정산
		elseif FChargeDiv="6" then
			getChargeDivName = "업체특정"
		elseif FChargeDiv="8" then
			getChargeDivName = "업체매입"
		elseif FChargeDiv="9" then
			getChargeDivName = "가맹점"
		elseif FChargeDiv="0" then
			getChargeDivName = "통합"
		else
			getChargeDivName = FChargeDiv
		end if
	end function

	' 사용중지. 디비에서 일괄로 쿼리해서 가져 오세요.
	public Function getPurchasetypeName()
    	IF FPurchasetype = "1" then
    	    getPurchasetypeName = "일반유통"
    	ELSEIF FPurchasetype = "4" then
    	    getPurchasetypeName = "사입"
    	ELSEIF FPurchasetype = "5" then
    	    getPurchasetypeName = "OFF사입"
    	ELSEIF FPurchasetype = "6" then
    	    getPurchasetypeName = "수입"
    	ELSEIF FPurchasetype = "7" then
    	    getPurchasetypeName = "브랜드수입"
        ELSEIF FPurchasetype = "8" then
    	    getPurchasetypeName = "제작"
        ELSEIF FPurchasetype = "9" then
    	    getPurchasetypeName = "해외직구"
        ELSEIF FPurchasetype = "10" then
    	    getPurchasetypeName = "B2B"
    	END IF
    end Function

	public function IsAvailJumun()
		IsAvailJumun = Not ((CStr(FCancelyn)="D") or (CStr(FCancelyn)="Y"))
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COffShopSellMasterDetailItem
    public Fidx
	public ForderNo
	public Ftotalsum
	public Frealsum
	public Fjumunmethod
	public Fshopregdate
	public Fitemname
	public Fitemoptionname
	public Fsellprice
	public Frealsellprice
	public Fitemno
	public FMakerID
	public Fpointuserno
    public Fcashsum
    public Fcardsum
    public FgiftcardPaysum
	public FextPaysum

	public Fspendmile
	public FTenGiftCardPaySum

    public FaddTaxCharge
    public FShopid
    public FrefOrderno
    public Fcardappno
	public FmatchCount

	Public function JumunMethodColor()
		if Cstr(Fjumunmethod) = "01" then
			JumunMethodColor = "#000000"
		elseif (Cstr(Fjumunmethod) = "02") or (Cstr(Fjumunmethod) = "06") then
			JumunMethodColor = "#0000FF"
		end if
	end function

	Public function JumunMethodName()
		if Cstr(Fjumunmethod) = "01" then
			JumunMethodName = "현금"
		elseif Cstr(Fjumunmethod) = "02" then
			JumunMethodName = "카드"
	    elseif Cstr(Fjumunmethod) = "03" or Cstr(Fjumunmethod) = "07" then
	        JumunMethodName = "복합"
		elseif Cstr(Fjumunmethod) = "09" then
			JumunMethodName = "기타"
	    elseif Cstr(Fjumunmethod) = "06" then
	        JumunMethodName = "Debit"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COffShopSellDetailItem
	public FIdx
	public FShopID
	public FMakerID
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fitemno
	public Fsellprice
	public Frealsellprice
	public Fsubtotal
	public FShopregDate
	public Fsuplyprice
	public FOrderNo
	public Fjungsanid
	public Fcurrentitemprice
	public FOnlineMwDiv
    public fextbarcode
    public fsellsum
    public FaddTaxCharge
    public fsuplysum
    public Fjcomm_cd
    public Fcomm_name
	public fsmallimage
	public foffimgsmall

	public function GetBarCode()
		GetBarCode = CStr(Fitemgubun) + CStr(Format00(6,FItemId)) + CStr(Fitemoption)
		if (FItemId >= 1000000) then
			getBarCode = CStr(Fitemgubun) + Format00(8,FItemId) + Fitemoption
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COffShopSellMasterItem
	public Fidx
	public Forderno
	public Fshopid
	public Ftotalsum
	public Frealsum
	public Fjumundiv
	public Fjumunmethod
	public Fshopregdate
	public Fcancelyn
	public Fregdate
	public Fshopidx

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COffShopRealJaegoDetail
	public Fidx
	public Fmasteridx
	public Fmakerid
	public Fitemgubun
	public Fshopitemid
	public Fitemoption
	public Frealjeago
	public Fcancelyn

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class CFranJungSanItem
	public FjungsanMasterIdx
	public Fidx
	public Fbaljuid
	public Fchargeuser
	public FYYYYMM
	public FTotNo
	public FTotalSellcash
	public FTotalBuyCash
	public Fcurrstate
	public FChargeDiv
	public Fjumundivcode
	public FDefaultMargin
	public Fipgodate
	public Fshopid
	public FjungsaMasterIdx
	public Fjungsantotitemcnt
	public Fjungsantotsum
	public FminusCharge
	public FChargePercent
	public FRealjungsansum
	public Fbigo
	public Fsegumil
	public Fipkumil
	public FJungsanChargediv
	public FconvCount
	public FconvSum
	public Flinkidx

	public function GetJumunDivName()
		if Fjumundivcode="101" then
			GetJumunDivName = "가맹점용 개별매입"
		elseif Fjumundivcode="111" then
			GetJumunDivName = "가맹점용 개별특정"
		elseif Fjumundivcode="121" then
			GetJumunDivName = "온라인특정재고->가맹점용특정"
		elseif Fjumundivcode="131" then
			GetJumunDivName = "온라인특정재고->가맹점용매입"
		elseif Fjumundivcode="201" then
			GetJumunDivName = "온라인매입재고->가맹점용매입"
		elseif Fjumundivcode="300" then
			GetJumunDivName = "온라인주문"
		elseif Fjumundivcode="501" then
			GetJumunDivName = "직영샾주문"
		elseif Fjumundivcode="502" then
			GetJumunDivName = "수수료샾"
		elseif Fjumundivcode="503" then
			GetJumunDivName = "프랜차이즈"
		else
			GetJumunDivName = ""
		end if
	end function

	public function GetJumunDivColor()
		if Fjumundivcode="101" then
			GetJumunDivColor = "#0000AA"
		elseif Fjumundivcode="111" then
			GetJumunDivColor = "#AA0000"
		elseif Fjumundivcode="121" then
			GetJumunDivColor = "#AA00AA"
		elseif Fjumundivcode="131" then
			GetJumunDivColor = "#00AAAA"
		elseif Fjumundivcode="201" then
			GetJumunDivColor = "#AAAA00"
		elseif Fjumundivcode="300" then
			GetJumunDivColor = "#FF0000"
		elseif Fjumundivcode="501" then
			GetJumunDivColor = "#0000FF"
		elseif Fjumundivcode="502" then
			GetJumunDivColor = "#00FF00"
		elseif Fjumundivcode="503" then
			GetJumunDivColor = "#AAFFAA"
		else
			GetJumunDivColor = "#000000"
		end if
	end function

	public function GetCurrStateName()
		if IsNull(Fcurrstate) or (Fcurrstate="") then
			GetCurrStateName = "미정산"
		elseif Fcurrstate="0" then
			GetCurrStateName = "수정중"
		elseif Fcurrstate="1" then
			GetCurrStateName = "업체확인중"
		elseif Fcurrstate="2" then
			GetCurrStateName = "업체확인완료"
		elseif Fcurrstate="3" then
			GetCurrStateName = "정산확정"
		elseif Fcurrstate="7" then
			GetCurrStateName = "입금완료"
		elseif Fcurrstate="8" then
			GetCurrStateName = "정산안함"
		elseif Fcurrstate="9" then
			GetCurrStateName = "통합정산"
		end if
	end function

	public function GetStateColor()
		if Fcurrstate="0" then
			GetStateColor = "#000000"
		elseif Fcurrstate="1" then
			GetStateColor = "#448888"
		elseif Fcurrstate="2" then
			GetStateColor = "#0000FF"
		elseif Fcurrstate="3" then
			GetStateColor = "#0000FF"
		elseif Fcurrstate="7" then
			GetStateColor = "#FF0000"
		elseif Fcurrstate="9" then
			GetStateColor = "#888844"
		elseif Fcurrstate=" " then
			GetStateColor = "#AAAAAA"
		else

		end if
	end function

	public function getChargeDivName()
		if FChargeDiv="2" then
			getChargeDivName = "10x10 특정"
		elseif FChargeDiv="4" then
			getChargeDivName = "10x10 매입"
		elseif FChargeDiv="5" then
			getChargeDivName = "출고분정산"
		elseif FChargeDiv="6" then
			getChargeDivName = "업체 특정"
		elseif FChargeDiv="8" then
			getChargeDivName = "업체 매입"
		elseif FChargeDiv="9" then
			getChargeDivName = "가맹점"
		elseif FChargeDiv="0" then
			getChargeDivName = "통합"
		else
			getChargeDivName = FChargeDiv
		end if
	end function

	public function getJungSanChargeDivName()
		if FJungsanChargediv="2" then
			getJungSanChargeDivName = "10x10 특정"
		elseif FJungsanChargediv="4" then
			getJungSanChargeDivName = "10x10 매입"
		elseif FJungsanChargediv="5" then
			getJungSanChargeDivName = "출고분정산"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivName = "업체 특정"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivName = "업체 매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivName = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivName = "통합"
		else
			getJungSanChargeDivName = FJungsanChargediv
		end if
	end function

	public function getJungSanChargeDivNameUpcheView()
		if FJungsanChargediv="2" then
			getJungSanChargeDivNameUpcheView = "특정"
		elseif FJungsanChargediv="5" then
			getJungSanChargeDivNameUpcheView = "매입"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivNameUpcheView = "특정"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivNameUpcheView = "매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivNameUpcheView = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivNameUpcheView = "통합"
		else
			getJungSanChargeDivNameUpcheView = FJungsanChargediv
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COffShopJungSanItem
	public Fidx
	public FShopid
	public Fchargeuser
	public Fchargename
	public FYYYYMM
	public Ftotno
	public FSum
	public FjungsaMasterIdx
	public Fcurrstate
	public Fjungsantotitemcnt
	public Fjungsantotsum
	public FminusCharge
	public FChargePercent
	public FRealjungsansum
	public Fbigo
	public Fsegumil
	public Fipkumil
	public Fchargediv
	public Fjungsan_acctname
	public Fjungsan_bank
	public Fjungsan_acctno
	public Fcompany_name
	public FJungsanChargediv
	public Fjungsan_date_off
	public Fjungsan_date_frn
	public FNoFixSum
	public FFixSum
	public FIpkumSum
	public FAutoJungsan
	public FShopName
	public FShopDiv
	public Fdefaultmargin
	public FFranChargeDiv
	public FGroupidx
	public FTaxRegdate
	public FDifferencekey
	public FTaxType
	public FTaxLinkidx
	public Fneotaxno
	public Foffgubun
	public Fonlinedefaultmargine
	public Fonlinemaeipdiv
	public Ftotchulgono
	public FtotchulgoSum

	public function GetOnlineMaeipDivName
		if Fonlinemaeipdiv="M" then
			GetOnlineMaeipDivName = "매입"
		elseif Fonlinemaeipdiv="W" then
			GetOnlineMaeipDivName = "특정"
		elseif Fonlinemaeipdiv="U" then
			GetOnlineMaeipDivName = "업체"
		else
			GetOnlineMaeipDivName = Fonlinemaeipdiv
		end if

	end function

	public function GetShopDivName()
		if IsNull(FShopDiv) then

		elseif FShopDiv="1" then
			GetShopDivName = "직영"
		elseif FShopDiv="2" then
			GetShopDivName = "수수료매장"
		elseif FShopDiv="3" then
			GetShopDivName = "가맹점"
		end if
	end function

	public function GetAutoJungsanName()
		if IsNull(FAutoJungsan) then

		elseif FAutoJungsan="Y" then
			GetAutoJungsanName = "자동"
		elseif FAutoJungsan="N" then
			GetAutoJungsanName = "수기"
		end if
	end function

	public function GetAutoJungsanColor()
		if IsNull(FAutoJungsan) then
			GetAutoJungsanColor = "#000000"
		elseif FAutoJungsan="Y" then
			GetAutoJungsanColor = "#000000"
		elseif FAutoJungsan="N" then
			GetAutoJungsanColor = "#4444AA"
		end if
	end function

	public function GetCurrStateName()
		if IsNull(Fcurrstate) or (Fcurrstate="") then
			GetCurrStateName = "미정산"
		elseif Fcurrstate="0" then
			GetCurrStateName = "수정중"
		elseif Fcurrstate="1" then
			GetCurrStateName = "업체확인중"
		elseif Fcurrstate="2" then
			GetCurrStateName = "업체확인완료"
		elseif Fcurrstate="3" then
			GetCurrStateName = "정산확정"
		elseif Fcurrstate="7" then
			GetCurrStateName = "입금완료"
		elseif Fcurrstate="8" then
			GetCurrStateName = "정산안함"
		elseif Fcurrstate="9" then
			GetCurrStateName = "통합정산"
		end if
	end function

	public function GetStateColor()
		if Fcurrstate="0" then
			GetStateColor = "#000000"
		elseif Fcurrstate="1" then
			GetStateColor = "#448888"
		elseif Fcurrstate="2" then
			GetStateColor = "#0000FF"
		elseif Fcurrstate="3" then
			GetStateColor = "#0000FF"
		elseif Fcurrstate="7" then
			GetStateColor = "#FF0000"
		elseif Fcurrstate="8" then
			GetStateColor = "#AAAAAA"
		elseif Fcurrstate="9" then
			GetStateColor = "#AAAAAA"
		else

		end if
	end function

	public function getChargeDivName()
		if FChargeDiv="2" then
			getChargeDivName = "10x10 특정"
		elseif FChargeDiv="4" then
			getChargeDivName = "10x10 매입"
		elseif FChargeDiv="5" then
			getChargeDivName = "출고분정산"
		elseif FChargeDiv="6" then
			getChargeDivName = "업체 특정"
		elseif FChargeDiv="8" then
			getChargeDivName = "업체 매입"
		elseif FChargeDiv="9" then
			getChargeDivName = "가맹점"
		elseif FChargeDiv="0" then
			getChargeDivName = "통합"
		else
			getChargeDivName = FChargeDiv
		end if
	end function

	public function getJungSanChargeDivName()
		if FJungsanChargediv="2" then
			getJungSanChargeDivName = "10x10 특정"
		elseif FJungsanChargediv="5" then
			getJungSanChargeDivName = "출고분정산"
		elseif FJungsanChargediv="4" then
			getJungSanChargeDivName = "10x10 매입"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivName = "업체 특정"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivName = "업체 매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivName = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivName = "통합"
		else
			getJungSanChargeDivName = FJungsanChargediv
		end if
	end function

	public function getJungSanChargeDivNameUpcheView()
		if FJungsanChargediv="2" then
			getJungSanChargeDivNameUpcheView = "특정"
		elseif FJungsanChargediv="5" then
			getJungSanChargeDivNameUpcheView = "매입"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivNameUpcheView = "특정"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivNameUpcheView = "매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivNameUpcheView = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivNameUpcheView = "통합"
		else
			getJungSanChargeDivNameUpcheView = FJungsanChargediv
		end if
	end function

	public function GetFranChargeDivName()
		if FFranChargeDiv="2" then
			GetFranChargeDivName = "특정"
		elseif FFranChargeDiv="4" then
			GetFranChargeDivName = "매입"
		elseif FFranChargeDiv="5" then
			GetFranChargeDivName = "매입"
		elseif FFranChargeDiv="6" then
			GetFranChargeDivName = "특정"
		elseif FFranChargeDiv="8" then
			GetFranChargeDivName = "매입"
		else
			GetFranChargeDivName = FFranChargeDiv
		end if
	end function

	public function GetSimpleTaxtypeName()
		if Ftaxtype="01" then
			GetSimpleTaxtypeName = "과세"
		elseif Ftaxtype="02" then
			GetSimpleTaxtypeName = "면세"
		elseif Ftaxtype="03" then
			GetSimpleTaxtypeName = "간이"
		end if
	end function

	public function GetTaxtypeNameColor()
		if Ftaxtype="01" then
			GetTaxtypeNameColor = "#000000"
		elseif Ftaxtype="02" then
			GetTaxtypeNameColor = "#FF3333"
		elseif Ftaxtype="03" then
			GetTaxtypeNameColor = "#3333FF"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COffShopJungSanDetailItem
	public Fidx
	public Fmasteridx
	public Forderno
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fsellprice
	public Frealsellprice
	public Fsuplyprice
	public Fitemno
	public Fmakerid
	public Flinkidx
	public Fjungsangubun

	public function getDetailGubunName()
		if Fjungsangubun = "101" then
			getDetailGubunName = "매입"
		elseif Fjungsangubun = "131" then
			getDetailGubunName = "특정재고->매입"
		elseif Fjungsangubun = "111" then
			getDetailGubunName = "특정"
		elseif Fjungsangubun = "121" then
			getDetailGubunName = "특정"
		elseif Fjungsangubun = "801" then
			getDetailGubunName = "off매입"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class COffshopAutoJungsan
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectYYYYMM
	public FRectMakerid
	public FRectJungsanYYYY
	public FRectJungsanMM
	public FRectShopID
	public FRectOnlyMaeipChecked

	public sub GetFranChulgoJungsanList
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = sqlStr + " select d.makerid, d.shopid,"
		sqlStr = sqlStr + " IsNULL(T.totno,0) as totno, IsNULL(T.totsum,0) as totsum,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv,"
		sqlStr = sqlStr + " d.chargediv, d.defaultmargin"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " and j.jungsanid=d.makerid"
			sqlStr = sqlStr + " and j.shopid=d.shopid"

			sqlStr = sqlStr + " left join ("
			sqlStr = sqlStr + " select m.socid as shopid, count(itemno) as totno, sum(buycash*itemno*-1) as totsum "
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m,"
			sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
			sqlStr = sqlStr + " where m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
			sqlStr = sqlStr + " and m.executedt<'" + nextYYYYMMDD + "'"
			sqlStr = sqlStr + " and m.ipchulflag='S'"
			sqlStr = sqlStr + " and m.code=d.mastercode"
			sqlStr = sqlStr + " and m.deldt is null"
			sqlStr = sqlStr + " and Left(m.socid,11)='streetshop8'"
			sqlStr = sqlStr + " and d.imakerid='" + FRectMakerid + "'"
			sqlStr = sqlStr + " and d.mwgubun='C'"
			sqlStr = sqlStr + " and d.deldt is null"
			sqlStr = sqlStr + " group by m.socid"

			sqlStr = sqlStr + " ) T on T.shopid=d.shopid"
		sqlStr = sqlStr + " where Left(d.shopid,11) ='streetshop8'"
		sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and (T.totno<>0 or j.totitemcnt<>0)"
		sqlStr = sqlStr + " and d.chargediv in ('4','5')"
		sqlStr = sqlStr + " order by d.makerid, d.shopid"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem

					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotNo = rsget("totno")
					FItemList(i).FTotalSellcash = rsget("totsum")
					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					FItemList(i).Fdefaultmargin	= rsget("defaultmargin")
					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).Fshopid = rsget("shopid")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public sub GetFranMeaipTargetListConv
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = sqlStr + " select m.id, d.imakerid, sum(d.itemno*-1) as ccnt,"
		sqlStr = sqlStr + " sum(d.itemno*d.sellcash*-1) as totalsellcash,"
		sqlStr = sqlStr + " sum(d.itemno*d.buycash*-1) as totalbuycash, m.executedt"
		sqlStr = sqlStr + " from  [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
		if FRectOnlyMaeipChecked="on" then
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on d.iitemgubun='10' and d.itemid=i.itemid"
		end if
		sqlStr = sqlStr + " where m.code=d.mastercode"
		sqlStr = sqlStr + " and d.imakerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and m.executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " and Left(m.code,2)='SO'"
		sqlStr = sqlStr + " and Left(m.socid,11)='streetshop8'"

		if FRectOnlyMaeipChecked="on" then
			sqlStr = sqlStr + " and i.mwdiv='W'"
		end if

		sqlStr = sqlStr + " group by m.id, d.imakerid,m.executedt"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fidx = rsget("id")
					FItemList(i).Fchargeuser = rsget("imakerid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fipgodate	   = rsget("executedt")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetFranMeaipTargetListByIpgo()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = "select m.socid, m.id , m.totalsellcash"
		sqlStr = sqlStr + " ,m.totalbuycash , m.executedt, m.divcode as jdivcode "
		sqlStr = sqlStr + " ,j.idx as jungsanmasteridx"
		'sqlStr = sqlStr + " ,IsNull(j.totsum,0) as totalsellcash"
		'sqlStr = sqlStr + " ,IsNull(j.realjungsansum,0) as totalbuycash"
		sqlStr = sqlStr + " ,j.currstate"
		sqlStr = sqlStr + " ,d.chargediv,d.defaultmargin"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.jungsanid=m.socid"
			sqlStr = sqlStr + " and j.jungsanid='" + FRectMakerid + "'"
			sqlStr = sqlStr + " and j.shopid='streetshop800'"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
			sqlStr = sqlStr + " on m.socid=d.makerid"
			sqlStr = sqlStr + " and d.shopid='streetshop800'"
		sqlStr = sqlStr + " where socid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and divcode ='801'"
		sqlStr = sqlStr + " and deldt is NULL"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fidx = rsget("id")
					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fchargeuser = rsget("socid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fipgodate	   = rsget("executedt")
					FItemList(i).Fdefaultmargin  = rsget("defaultmargin")
					FItemList(i).Fjumundivcode = rsget("jdivcode")
					FItemList(i).FChargeDiv    = rsget("chargediv")
					FItemList(i).Fcurrstate    = rsget("currstate")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end Sub

	public Sub GetFranMeaipTargetList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = "select m.targetid, m.idx , m.totalsellcash"
		sqlStr = sqlStr + " ,m.totalbuycash , m.ipgodate, m.divcode as jdivcode "
		sqlStr = sqlStr + " ,j.idx as jungsanmasteridx"
		'sqlStr = sqlStr + " ,IsNull(j.totsum,0) as totalsellcash"
		'sqlStr = sqlStr + " ,IsNull(j.realjungsansum,0) as totalbuycash"
		sqlStr = sqlStr + " ,j.currstate"
		sqlStr = sqlStr + " ,d.chargediv,d.defaultmargin"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.jungsanid=m.targetid"
			sqlStr = sqlStr + " and j.jungsanid='" + FRectMakerid + "'"
			sqlStr = sqlStr + " and j.shopid='streetshop800'"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
			sqlStr = sqlStr + " on m.targetid=d.makerid"
			sqlStr = sqlStr + " and d.shopid='streetshop800'"
		sqlStr = sqlStr + " where baljuid='10x10'"
		sqlStr = sqlStr + " and targetid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and statecd='9'"
		sqlStr = sqlStr + " and ipgodate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and ipgodate<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and divcode in ('101','131')"
		sqlStr = sqlStr + " and deldt is NULL"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fidx = rsget("idx")
					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fchargeuser = rsget("targetid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fipgodate	   = rsget("ipgodate")
					FItemList(i).Fdefaultmargin  = rsget("defaultmargin")
					FItemList(i).Fjumundivcode = rsget("jdivcode")
					FItemList(i).FChargeDiv    = rsget("chargediv")
					FItemList(i).Fcurrstate    = rsget("currstate")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end Sub

	public Sub GetFranWitakTargetList()
		dim sqlStr, i
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = sqlStr + " select sd.shopid, sd.makerid "
		sqlStr = sqlStr + " ,count(d.itemno) as totno"
		sqlStr = sqlStr + " ,j.idx as jungsanmasteridx"
		sqlStr = sqlStr + " ,IsNull(j.totsum,0) as totalsellcash"
		sqlStr = sqlStr + " ,IsNull(j.realjungsansum,0) as totalbuycash, j.currstate,j.chargediv as jdivcode"
		sqlStr = sqlStr + " ,sd.chargediv, sd.defaultmargin"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_designer sd"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.jungsanid=sd.makerid"
			sqlStr = sqlStr + " and j.jungsanid='" + FRectMakerid + "'"
			sqlStr = sqlStr + " and Left(j.shopid,11)='streetshop8'"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " where m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.shopid =sd.shopid"
		sqlStr = sqlStr + " and Left(m.shopid,11)='streetshop8'"
		sqlStr = sqlStr + " and d.makerid=sd.makerid"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn='N'"
		sqlStr = sqlStr + " and sd.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and sd.chargediv in ('2','6')"

		sqlStr = sqlStr + " group by sd.shopid, sd.makerid,j.idx,j.totsum,j.realjungsansum,j.currstate,j.chargediv"
		sqlStr = sqlStr + " ,sd.chargediv, sd.defaultmargin"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fshopid = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).FYYYYMM  = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno		= rsget("totno")
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fdefaultmargin  = rsget("defaultmargin")
					FItemList(i).Fjumundivcode = rsget("jdivcode")
					FItemList(i).FChargeDiv    = rsget("chargediv")
					FItemList(i).Fcurrstate    = rsget("currstate")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetFranMeaipTargetWitakList()
		dim sqlStr, i
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select m.shopid, d.makerid, "
		sqlStr = sqlStr + " count(itemno) as totno, sum(realsellprice*itemno) as totsum"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.jungsanid=m.targetid"
		sqlStr = sqlStr + " and j.shopid='streetshop800'"
		sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer u"
		sqlStr = sqlStr + " on u.makerid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and u.makerid=m.targetid and u.shopid='streetshop800' "
		sqlStr = sqlStr + " where m.baljuid='10x10'"
		sqlStr = sqlStr + " and m.targetid='" + FRectMakerid + "'"
		sqlStr = sqlStr + " and m.statecd='9'"
		sqlStr = sqlStr + " and m.ipgodate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and m.ipgodate<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.divcode in ('101','131')"
		sqlStr = sqlStr + " and m.deldt is NULL"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fidx = rsget("idx")
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")
					FItemList(i).Fdefaultmargin  = rsget("defaultmargin")
					FItemList(i).Fjumundivcode = rsget("jdivcode")
					FItemList(i).FChargeDiv    = rsget("chargediv")
					FItemList(i).Fipgodate	   = rsget("ipgodate")
					FItemList(i).Fcurrstate    = rsget("currstate")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close


	end sub

	public Sub GetTargetList()
		dim sqlStr, i
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select T.*, u.chargediv,u.autojungsan,u.defaultmargin,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv,"
		sqlStr = sqlStr + " s.shopname, s.shopdiv"
		sqlStr = sqlStr + " from ("
		sqlStr = sqlStr + " select m.shopid, d.makerid, "
		sqlStr = sqlStr + " count(itemno) as totno, sum(realsellprice*itemno) as totsum"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn='N'"
		sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if

		sqlStr = sqlStr + " group by m.shopid,d.makerid"
		sqlStr = sqlStr + " ) as T"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.jungsanid=T.makerid"
		sqlStr = sqlStr + " and j.shopid=T.shopid"
		sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
		end if

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer u"
		sqlStr = sqlStr + " on u.makerid='" + FRectMakerid + "' and u.makerid=T.makerid and u.shopid=T.shopid "

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s"
		sqlStr = sqlStr + " on s.userid=T.shopid "
		sqlStr = sqlStr + " where s.shopdiv<>'3'"
		sqlStr = sqlStr + " order by T.shopid, u.chargediv desc, T.totsum desc, T.totno desc"

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanItem
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno      = rsget("totno")
					FItemList(i).FSum        = rsget("totsum")
					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).FAutojungsan = rsget("autojungsan")
					FItemList(i).FShopname	 = db2html(rsget("shopname"))
					FItemList(i).FShopDiv	 = rsget("shopdiv")
					FItemList(i).Fdefaultmargin = rsget("defaultmargin")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class COffShopSellReport
	public FItemList()
	public FCountList()
	public FPageCount
	public FOneJeaGoMaster
	public FOneJungSanMaster
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectStartDay
	public FRectEndDay
	public FRectNormalOnly
	public FRectJungsanId
	public FRectDesigner
	public FRectShopID
	public FRectTerms
	public FRectItemGubun
	public FRectItemId
    public FRectItemOption
	public FRectJungsanYYYYMM
	public FRectJungsanYYYY
	public FRectJungsanMM
	public FYYYYMMDDHHNNSS
	public FRectJaegoNo
	public FRectIDX
	public maxt
	public maxt2
	public maxc
	public FRectPointYN
	public FRectOnlymijungsan
	public FRectOnlyUpcheJungSan
	public FRectOnlyFranUpcheJungSan
	public FRectNotIncludeWonChon
	public FRectOnlyIncludeWonChon
	public FRectOnlyIncludeNoTax
	public FRectOnlyShop
	public FRectNotChargeDiv
	public FRectChargeDiv
	public FRectOffgubun
	public FDayTsellsum
	public FDayTea
	public FRectOrder
	public FRectMWgubun
	public FRectnomeachul
	public FRectUpcheWitakOnly
	public FRectJungsanDate
	public FRectOldData
    public frectdatefg
    public FRectmakerid
	public frectitemname
	public frectweekdate
	public FRectOldJumun
	public frectupcheyn
	public frectdategubun
	public frectoffcatecode
	public frectoffmduserid
	public FRectBrandPurchaseType
	public FRectOrdertype
	public FRectextbarcode
	public FRectBanPum
	public frectbuyergubun
	public FRectInc3pl
	public FRectExcMatchFinish
	public FRectPgDataCheck
	public FRectCardPayOnly
	public FRectCardSum
	Public FRectPaySum
	public FRectStartdate
	public FRectEndDate
	public FRectJungSanGubun
    public FRectCommCD
    public FRectShowOrder

	function MaxVal(a,b)
		if (CLng(a)> CLng(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function

	'//admin/offshop/sellreportbrand.asp
	public Sub GetBrandSellSumList()
		dim i,sqlStr ,sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')=''"
	    end if
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum > 0"
		end if

		if FRectNormalOnly="on" then
			sqlsearch = sqlsearch + " and m.cancelyn='N'"
			sqlsearch = sqlsearch + " and d.cancelyn='N'"
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		if FRectOnlyShop<>"" then
			sqlsearch = sqlsearch + " and Left(m.shopid,4)<>'cafe'"
		end if

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid = '"&FRectmakerid&"'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if

		If frectoffcatecode <> "" Then
			sqlsearch = sqlsearch + " and p.offcatecode = '" + CStr(frectoffcatecode) + "' "
		End IF

		If frectoffmduserid <> "" Then
			sqlsearch = sqlsearch + " and p.offmduserid = '" + CStr(frectoffmduserid) + "' "
		End IF

		if FRectJungSanGubun <> "" and FRectShopid<>"" then
			sqlsearch = sqlsearch + " and s.chargediv = " + CStr(FRectJungSanGubun)
		end if

		sqlStr = " SELECT top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " sum(d.itemno * isnull(d.realsellprice,0)) as subtotal"
		''sqlStr = sqlStr + " , isnull(sum(d.itemno * d.addTaxCharge),0) as addTaxChargeSum"
		sqlStr = sqlStr + " , sum(d.itemno) as cnt"
		sqlStr = sqlStr + " , d.makerid"
		sqlStr = sqlStr + " , sum(d.itemno * isnull(d.suplyprice,0)) as suplyprice"
		sqlStr = sqlStr + " , sum(d.itemno * isnull(d.Iorgsellprice,0)) as Iorgsellprice"
		sqlStr = sqlStr + " , sum(d.itemno * isnull(d.realsellprice,0)) - sum(d.itemno * isnull(d.suplyprice,0)) as profit"

		if frectdategubun = "M" then
			sqlStr = sqlStr & " ,convert(varchar(7),m.IXyyyymmdd) as IXyyyymmdd"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " ,s.chargediv"
		end if

		    sqlSTr = sqlStr + " , p.purchaseType, pc.pcomm_name as purchasetypename"

		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m with (nolock)" + vbcrlf
			sqlStr = sqlStr + " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d with (nolock)" + vbcrlf
			sqlStr = sqlStr + " 	on m.orderno = d.orderno"
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m with (nolock)" + vbcrlf
			sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d with (nolock)" + vbcrlf
			sqlStr = sqlStr + " 	on m.orderno = d.orderno"
		end if

		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_user u with (nolock)"
		sqlStr = sqlStr + " 	on m.shopid = u.userid"
		sqlStr = sqlStr + " join [db_partner].[dbo].tbl_partner p with (nolock) on d.makerid = p.id "

		if (FRectBrandPurchaseType<>"") then
			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				sqlStr = sqlStr + " 	and p.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				sqlStr = sqlStr & " 	and p.purchasetype in ('3','5','6')"
			else
				sqlStr = sqlStr + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].tbl_partner_comm_code as pc with (nolock)"
		sqlStr = sqlStr & " 	on pc.pcomm_group='purchasetype' and pc.pcomm_isusing='Y' and p.purchasetype=pc.pcomm_cd"

		if FRectShopid<>"" then
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s with (nolock)"
			sqlStr = sqlStr + " 	on s.shopid='" + FRectShopid + "'"
			sqlStr = sqlStr + " 	and d.makerid=s.makerid"
		end if

		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr & "       on m.shopid=pp.id "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " group by d.makerid, p.purchaseType, pc.pcomm_name"

		if frectdategubun = "M" then
			sqlStr = sqlStr & " ,convert(varchar(7),m.IXyyyymmdd)"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " ,s.chargediv"
		end if

		sqlStr = sqlStr + " order by"

		if frectdategubun = "M" then
			sqlStr = sqlStr & " IXyyyymmdd desc,"
		end if

		Select Case FRectOrdertype
			Case "totalprice"
				'매출순
				sqlStr = sqlStr & " subtotal Desc"
	    	Case "gain"
	    		'수익순
	            sqlStr = sqlStr & " profit Desc"
			Case "ea"
				'수량순
				sqlStr = sqlStr & " cnt Desc, subtotal desc"
			case else
				sqlStr = sqlStr + " subtotal desc"
		end Select

	''rw sqlStr
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopSellByTerm
				FItemList(i).fpurchasetypename  = rsget("purchasetypename")
				FItemList(i).fIorgsellprice  = rsget("Iorgsellprice")
				FItemList(i).FMakerid  = rsget("makerid")
				FItemList(i).FCount = rsget("cnt")
				FItemList(i).FSum   = rsget("subtotal")
				FItemList(i).fsuplyprice  = rsget("suplyprice")
				FItemList(i).fprofit  = rsget("profit")
				''FItemList(i).FaddTaxChargeSum  = rsget("addTaxChargeSum")

				if frectdategubun = "M" then
					FItemList(i).fIXyyyymmdd = rsget("IXyyyymmdd")
				end if

				if FRectShopid<>"" then
					FItemList(i).FChargeDiv = rsget("chargediv")
				end if

                    FItemList(i).FpurchaseType = rsget("purchaseType")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'//admin/offshop/brandshopdetail.asp
	public Sub GetBrandshopSell()
		dim i,sqlStr ,sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')=''"
	    end if
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

		if FRectNormalOnly="on" then
			sqlsearch = sqlsearch + " and m.cancelyn='N'"
			sqlsearch = sqlsearch + " and d.cancelyn='N'"
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		if FRectOnlyShop<>"" then
			sqlsearch = sqlsearch + " and Left(m.shopid,4)<>'cafe'"
		end if

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectmakerid<>"" then
			sqlsearch = sqlsearch + " and d.makerid = '"&FRectmakerid&"'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if

		sqlStr = " SELECT top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " sum(d.itemno * d.realsellprice) as subtotal"
		sqlStr = sqlStr + " , sum(d.itemno * d.addTaxCharge) as addTaxChargeSum"
		sqlStr = sqlStr + " , sum(d.itemno) as cnt,d.makerid "
		sqlStr = sqlStr + " ,sum(d.itemno * d.suplyprice) as suplyprice"
		sqlStr = sqlStr + " ,sum(d.realsellprice*d.itemno-d.suplyprice*d.itemno) as profit"
		sqlStr = sqlStr + " ,s.chargediv ,u.shopname , m.shopid"

		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m with (nolock)" + vbcrlf
			sqlStr = sqlStr + " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d with (nolock)" + vbcrlf
			sqlStr = sqlStr + " 	on m.orderno = d.orderno"
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m with (nolock)" + vbcrlf
			sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d with (nolock)" + vbcrlf
			sqlStr = sqlStr + " 	on m.orderno = d.orderno"
		end if

		if (FRectBrandPurchaseType<>"") then
		    sqlStr = sqlStr + " Join db_partner.dbo.tbl_partner p with (nolock)"
		    sqlStr = sqlStr + " 	on d.makerid=p.id"

			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				sqlStr = sqlStr + " 	and p.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				sqlStr = sqlStr & " 	and p.purchasetype in ('3','5','6')"
			else
				sqlStr = sqlStr + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_user u with (nolock)"
		sqlStr = sqlStr + " 	on m.shopid = u.userid"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s with (nolock)"
		sqlStr = sqlStr + " 	on m.shopid = s.shopid"
		sqlStr = sqlStr + " 	and d.makerid=s.makerid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner pp with (nolock)"
	    sqlStr = sqlStr & "       on m.shopid=pp.id "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " group by d.makerid ,s.chargediv ,u.shopname, m.shopid ,u.shopdiv"
		sqlStr = sqlStr + " order by u.shopdiv asc ,m.shopid asc ,subtotal desc"

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COffShopSellByTerm

				FItemList(i).FMakerid  = rsget("makerid")
				FItemList(i).FCount = rsget("cnt")
				FItemList(i).FSum   = rsget("subtotal")
				FItemList(i).fsuplyprice  = rsget("suplyprice")
				FItemList(i).fprofit  = rsget("profit")
				FItemList(i).FaddTaxChargeSum  = rsget("addTaxChargeSum")
				FItemList(i).FChargeDiv = rsget("chargediv")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fshopname = db2html(rsget("shopname"))

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	public Sub GetNotExistsInSertJungSanMaster()
		dim i,sqlStr, masterExists
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select j.*, u.chargename from [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_chargeuser u"
		sqlStr = sqlStr + " on j.jungsanid=u.chargeuser"
		sqlStr = sqlStr + " where j.yyyymm='" + FRectJungsanYYYYMM + "'"
		sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
		sqlStr = sqlStr + " and j.jungsanid='" + FRectJungsanID + "'"
		rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				masterExists = true
				redim  FItemList(0)

				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).FjungsaMasterIdx = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).Fchargeuser = rsget("jungsanid")
				FItemList(i).Fchargename = rsget("chargename")
				FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
				FItemList(i).Ftotno      = rsget("totitemcnt")
				FItemList(i).FSum        = rsget("totsum")
				FItemList(i).Fcurrstate  = rsget("currstate")

				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")
				FItemList(i).Fbigo           = db2html(rsget("bigo"))
				FItemList(i).Fsegumil        = rsget("segumil")
				FItemList(i).Fipkumil        = rsget("ipkumil")

			else
				masterExists = false
			end if
		rsget.Close


		if Not masterExists then
			sqlStr = " insert into [db_shop].[dbo].tbl_shop_jungsanmaster"
			sqlStr = sqlStr + " (yyyymm,shopid,jungsanid,totitemcnt,totsum,currstate)"
			sqlStr = sqlStr + " select '" + FRectJungsanYYYYMM + "','" + FRectShopID +"','" + FRectJungsanID + "',"
			sqlStr = sqlStr + " count(itemno) as totno, sum(realsellprice*itemno) as totsum, '0'"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
			sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
			sqlStr = sqlStr + " and d.jungsanid='" + FRectJungsanId + "'"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
			rsget.Open sqlStr,dbget,1

			sqlStr = " select j.*, u.chargename from [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_chargeuser u"
			sqlStr = sqlStr + " on j.jungsanid=u.chargeuser"
			sqlStr = sqlStr + " where j.yyyymm='" + FRectJungsanYYYYMM + "'"
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
			sqlStr = sqlStr + " and j.jungsanid='" + FRectJungsanID + "'"
			rsget.Open sqlStr,dbget,1
				if Not rsget.Eof then
					masterExists = true
					redim  FItemList(0)

					set FItemList(i) = new COffShopJungSanItem
					FItemList(i).FjungsaMasterIdx = rsget("idx")
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("jungsanid")
					FItemList(i).Fchargename = rsget("chargename")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno      = rsget("totitemcnt")
					FItemList(i).FSum        = rsget("totsum")
					FItemList(i).Fcurrstate = rsget("currstate")

					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FChargePercent  = rsget("chargepercent")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					FItemList(i).Fbigo           = db2html(rsget("bigo"))
					FItemList(i).Fsegumil        = rsget("segumil")
					FItemList(i).Fipkumil        = rsget("ipkumil")
				else
					masterExists = false
				end if
			rsget.Close
		end if
	end Sub

	public Sub GetDesignerJungsanList()
		dim i,sqlStr
		sqlStr = "select count(idx) as cnt from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where jungsanid='" + FRectJungsanId + "'"
		sqlStr = sqlStr + " and currstate >0"
		sqlStr = sqlStr + " and currstate <8"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectOnlyUpcheJungSan="on" then
			sqlStr = sqlStr + " and chargediv in ('2','6','8')"
		end if

		if FRectOnlyFranUpcheJungSan="on" then
			sqlStr = sqlStr + " and chargediv ='9'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.*, s.shopname "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s on m.shopid=s.userid"
		sqlStr = sqlStr + " where jungsanid='" + FRectJungsanId + "'"

		sqlStr = sqlStr + " and currstate >0"
		sqlStr = sqlStr + " and currstate <8"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectOnlyUpcheJungSan="on" then
			sqlStr = sqlStr + " and chargediv in ('2','6','8')"
		end if

		if FRectOnlyFranUpcheJungSan="on" then
			sqlStr = sqlStr + " and chargediv ='9'"
		end if

		sqlStr = sqlStr + " order by idx desc"

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
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).FjungsaMasterIdx = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Ftotno      = rsget("totitemcnt")
				FItemList(i).FSum        = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")
				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")
				FItemList(i).Fbigo           = db2html(rsget("bigo"))
				FItemList(i).Fsegumil        = rsget("segumil")
				FItemList(i).Fipkumil        = rsget("ipkumil")
				FItemList(i).FJungsanChargediv  = rsget("chargediv")
				FItemList(i).Fshopname = db2html(rsget("shopname"))
				FItemList(i).FFranChargediv = rsget("franchargediv")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public Sub GetMiJungSanList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select m.shopid, u.chargeuser, u.chargename,"
		sqlStr = sqlStr + " count(itemno) as totno, sum(realsellprice*itemno) as totsum,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,IsNull(j.minuscharge,0) as minuscharge,IsNull(j.realjungsansum,0) as realjungsansum"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_chargeuser u"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.jungsanid=u.chargeuser"
		sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn='N'"
		sqlStr = sqlStr + " and d.jungsanid=u.chargeuser"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " group by m.shopid,u.chargeuser,u.chargename,j.idx,j.currstate,j.minuscharge,j.realjungsansum"
		sqlStr = sqlStr + " order by totsum desc,totno desc"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanItem
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("chargeuser")
					FItemList(i).Fchargename = rsget("chargename")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno      = rsget("totno")
					FItemList(i).FSum        = rsget("totsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetJungsanSummaryList()
		dim sqlStr,i
		sqlStr = "select m.yyyymm, m.shopid, sum(IsNull(m.totitemcnt,0)) as totitemcnt,"
		sqlStr = sqlStr + " sum(IsNull(m.totsum,0)) as totsum,"
		sqlStr = sqlStr + " sum(IsNull(m.minuscharge,0)) as minuscharge, "
		sqlStr = sqlStr + " sum(IsNull(m.chargepercent,0)) as chargepercent, "
		sqlStr = sqlStr + " sum(IsNull(m.realjungsansum,0)) as realjungsansum, m.currstate"

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on m.jungsanid=p.id"
		sqlStr = sqlStr + " where m.currstate<7"
		sqlStr = sqlStr + " group by m.yyyymm, m.shopid, m.currstate"
		sqlStr = sqlStr + " order by m.yyyymm desc, m.currstate"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Ftotno      = rsget("totitemcnt")
				FItemList(i).FSum        = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")
				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end Sub

	public Sub GetJungsanSummaryListByChargeDiv()
		dim sqlStr,i
		sqlStr = "select m.yyyymm, m.chargediv,  "
		sqlStr = sqlStr + " Sum (case when m.currstate='7' then m.realjungsansum"
		sqlStr = sqlStr + " else 0 end ) as ipkumsum, "
		sqlStr = sqlStr + " Sum (case when (m.currstate='3') then m.realjungsansum"
		sqlStr = sqlStr + " else 0 end ) as fixsum, "
		sqlStr = sqlStr + " Sum (case when (m.currstate='0') or (m.currstate='1') or (m.currstate='2') then m.realjungsansum"
		sqlStr = sqlStr + " else 0 end ) as nofixsum "

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " where m.regdate>'2004-01-01'"
		sqlStr = sqlStr + " group by m.yyyymm, m.chargediv"
		sqlStr = sqlStr + " order by m.yyyymm desc, m.chargediv"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Fchargediv = rsget("chargediv")

				FItemList(i).FNoFixSum  = rsget("nofixsum")
				FItemList(i).FFixSum    = rsget("fixsum")
				FItemList(i).FIpkumSum  = rsget("ipkumsum")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end Sub

	public Sub GetJungsanFix26MasterList()
		dim i,sqlStr
		sqlStr = " select m.*, p.jungsan_acctname, p.jungsan_bank, p.jungsan_acctno, p.company_name, p.jungsan_date_off, p.jungsan_date_frn"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on m.jungsanid=p.id"
		if FRectJungsanDate="w" then
			sqlStr = sqlStr + " left join (select distinct makerid "
			sqlStr = sqlStr + " 			from [db_shop].[dbo].tbl_shop_designer "
			sqlStr = sqlStr + " 			where chargediv='6' "
			sqlStr = sqlStr + " ) as U on m.jungsanid=U.makerid"
		end if

		sqlStr = sqlStr + " where currstate='3'"

		sqlStr = sqlStr + " and m.chargediv in ('2','6','8','0','9')"

		if FRectOffgubun<>"" then
			sqlStr = sqlStr + " and m.offgubun='" + FRectOffgubun + "'"
		end if

		if (FRectJungsanDate="n") then
			sqlStr = sqlStr + " and (((Left(m.shopid,11)='streetshop0') and (IsNULL(p.jungsan_date_off,'')='')) or  ((Left(m.shopid,11)='streetshop8') and (IsNULL(p.jungsan_date_frn,'')='')))"
		elseif ((FRectJungsanDate="15일") or (FRectJungsanDate="말일")) then
			sqlStr = sqlStr + " and (((Left(m.shopid,11)='streetshop0') and (p.jungsan_date_off='" + FRectJungsanDate + "')) or  ((Left(m.shopid,11)='streetshop8') and (p.jungsan_date_frn='" + FRectJungsanDate + "')))"
		elseif FRectJungsanDate<>"" then
			sqlStr = sqlStr + " and U.makerid is not null"
		end if

		sqlStr = sqlStr + " order by m.yyyymm asc, m.taxregdate, m.jungsanid "

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).Fidx	 = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).FChargeUser = rsget("jungsanid")

				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Ftotno      = rsget("totitemcnt")
				FItemList(i).FSum        = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")

				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")


				FItemList(i).Fjungsan_acctname = rsget("jungsan_acctname")
				FItemList(i).Fjungsan_bank = rsget("jungsan_bank")
				FItemList(i).Fjungsan_acctno = rsget("jungsan_acctno")
				FItemList(i).Fcompany_name = rsget("company_name")
				FItemList(i).Fjungsan_date_off = rsget("jungsan_date_off")
				FItemList(i).Fjungsan_date_frn = rsget("jungsan_date_frn")


				FItemList(i).FGroupidx      = rsget("groupidx")
				FItemList(i).FTaxRegdate    = rsget("taxregdate")
				FItemList(i).FDifferencekey = rsget("differencekey")
				FItemList(i).FTaxType       = rsget("taxtype")
				FItemList(i).FTaxLinkidx    = rsget("taxlinkidx")
				FItemList(i).Fneotaxno      = rsget("neotaxno")
				FItemList(i).Foffgubun      = rsget("offgubun")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Sub GetJungsanFixMasterList()
		dim i,sqlStr
		sqlStr = " select m.*, p.jungsan_acctname, p.jungsan_bank, p.jungsan_acctno,p.company_name "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on m.jungsanid=p.id"
		sqlStr = sqlStr + " where currstate='3'"

		if FRectNotChargeDiv<>"" then
			sqlStr = sqlStr + " and m.chargediv='" + FRectNotChargeDiv + "'"
		end if

		if FRectNotIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun<>'원천징수'"
			sqlStr = sqlStr + " and p.jungsan_gubun<>'면세'"
		end if

		if FRectOnlyIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='원천징수'"
		end if

		if FRectOnlyIncludeNoTax<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='면세'"
		end if

		sqlStr = sqlStr + " order by m.idx desc"
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new COffShopJungSanItem
				FItemList(i).Fidx	 = rsget("idx")
				FItemList(i).FShopid	 = rsget("shopid")
				FItemList(i).FChargeUser = rsget("jungsanid")

				FItemList(i).FYYYYMM     = rsget("yyyymm")
				FItemList(i).Ftotno      = rsget("totitemcnt")
				FItemList(i).FSum        = rsget("totsum")
				FItemList(i).Fcurrstate = rsget("currstate")

				FItemList(i).FminusCharge    = rsget("minuscharge")
				FItemList(i).FChargePercent  = rsget("chargepercent")
				FItemList(i).FRealjungsansum = rsget("realjungsansum")


				FItemList(i).Fjungsan_acctname = rsget("jungsan_acctname")
				FItemList(i).Fjungsan_bank = rsget("jungsan_bank")
				FItemList(i).Fjungsan_acctno = rsget("jungsan_acctno")
				FItemList(i).Fcompany_name = rsget("company_name")

				FItemList(i).FGroupidx      = rsget("groupidx")
				FItemList(i).FTaxRegdate    = rsget("taxregdate")
				FItemList(i).FDifferencekey = rsget("differencekey")
				FItemList(i).FTaxType       = rsget("taxtype")
				FItemList(i).FTaxLinkidx    = rsget("taxlinkidx")
				FItemList(i).Fneotaxno      = rsget("neotaxno")
				FItemList(i).Foffgubun      = rsget("offgubun")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

	public Sub GetOneJungsanMaster()
		dim i,sqlStr
		sqlStr = " select * from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where idx=" + FRectIdx + ""
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
				set FItemList(0) = new COffShopJungSanItem
				FItemList(0).FjungsaMasterIdx = rsget("idx")
				FItemList(0).FShopid	 = rsget("shopid")
				FItemList(0).FYYYYMM     = rsget("yyyymm")
				FItemList(0).Ftotno      = rsget("totitemcnt")
				FItemList(0).FSum        = rsget("totsum")
				FItemList(0).Fcurrstate = rsget("currstate")

				FItemList(0).FminusCharge    = rsget("minuscharge")
				FItemList(0).FChargePercent  = rsget("chargepercent")
				FItemList(0).FRealjungsansum = rsget("realjungsansum")
				FItemList(0).Fbigo           = db2html(rsget("bigo"))
				FItemList(0).Fsegumil        = rsget("segumil")
				FItemList(0).Fipkumil        = rsget("ipkumil")
		end if

		rsget.Close
	end Sub

	public Sub GetOffJungSanDetailSum()
		dim i,sqlStr
		sqlStr = "select itemgubun,itemid,itemoption,itemname,itemoptionname,sellprice,realsellprice,suplyprice,sum(itemno) as itemno"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsandetail"
		sqlStr = sqlStr + " where masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " group by itemgubun,itemid,itemoption,itemname,itemoptionname,sellprice,realsellprice,suplyprice"
		sqlStr = sqlStr + " order by itemgubun,itemid,itemoption"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanDetailItem
					FItemList(i).Fitemgubun      = rsget("itemgubun")
					FItemList(i).Fitemid         = rsget("itemid")
					FItemList(i).Fitemoption     = rsget("itemoption")
					FItemList(i).Fitemname       = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
					FItemList(i).Fsellprice      = rsget("sellprice")
					FItemList(i).Frealsellprice  = rsget("realsellprice")
					FItemList(i).Fsuplyprice     = rsget("suplyprice")
					FItemList(i).Fitemno         = rsget("itemno")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.close
	end sub

	public Sub GetOffJungSanDetail()
		dim i,sqlStr
		dim isEof

		sqlStr = "select top 1 * from [db_shop].[dbo].tbl_shop_jungsanmaster"
		sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
		if FRectJungsanId<>"" then
				sqlStr = sqlStr + " and jungsanid='" + CStr(FRectJungsanId) + "'"
		end if

		rsget.Open sqlStr,dbget,1

		if Not rsget.Eof then
			set FOneJungSanMaster = new COffShopJungsanItem

			FOneJungSanMaster.FShopid	 = rsget("shopid")
			FOneJungSanMaster.Fchargeuser = rsget("jungsanid")
			FOneJungSanMaster.FYYYYMM     = rsget("yyyymm")
			FOneJungSanMaster.Ftotno      = rsget("totitemcnt")
			FOneJungSanMaster.FSum        = rsget("totsum")
			FOneJungSanMaster.Fcurrstate = rsget("currstate")
			FOneJungSanMaster.FminusCharge    = rsget("minuscharge")
			FOneJungSanMaster.FRealjungsansum = rsget("realjungsansum")

			FOneJungSanMaster.Fchargediv = rsget("chargediv")

			FOneJungSanMaster.FSegumil = rsget("segumil")
			FOneJungSanMaster.Fipkumil = rsget("ipkumil")
			'FOneJungSanMaster.Fregdate = rsget("regdate")
		else
			isEof = true
		end if

		rsget.Close

		if isEof then dbget.close()	:	response.End

		sqlStr = "select * from [db_shop].[dbo].tbl_shop_jungsandetail"
		sqlStr = sqlStr + " where masteridx=" + CStr(FRectIdx)
		sqlStr = sqlStr + " order by orderno"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanDetailItem
					FItemList(i).Fidx            = rsget("idx")
					FItemList(i).Fmasteridx      = rsget("masteridx")
					FItemList(i).Forderno        = rsget("orderno")
					FItemList(i).Fitemgubun      = rsget("itemgubun")
					FItemList(i).Fitemid         = rsget("itemid")
					FItemList(i).Fitemoption     = rsget("itemoption")
					FItemList(i).Fitemname       = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
					FItemList(i).Fsellprice      = rsget("sellprice")
					FItemList(i).Frealsellprice  = rsget("realsellprice")
					FItemList(i).Fsuplyprice     = rsget("suplyprice")
					FItemList(i).Fitemno         = rsget("itemno")
					FItemList(i).Fmakerid        = rsget("makerid")
					FItemList(i).Flinkidx        = rsget("linkidx")
					FItemList(i).Fjungsangubun   = rsget("jungsangubun")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.close
	end Sub


	'' 가맹점 특정->매입 출고분 정산내역 검토
	public sub GetFranWitak2MeaipChulgoJungSanAutoList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select C.shopid , C.makerid, C.totcnt ,"
		sqlStr = sqlStr + " C.totalsellcash, "
		sqlStr = sqlStr + " C.totalbuycash,"
		sqlStr = sqlStr + " IsNULL(j.idx,'') as jungsanmasteridx,"
		sqlStr = sqlStr + " IsNULL(j.currstate,'') as currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " IsNull(j.chargediv,'') as jchargediv," 			'' (정산당시 정산 구분)
		sqlStr = sqlStr + " IsNull(d.chargediv,'') as chargediv, "			'' (현재 정산 구분)
		sqlStr = sqlStr + " IsNull(d.defaultmargin,0) as defaultmargin"		'' (현재 마진)
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " ( " '' 가맹점 특정 -> 매입 출고 내역
		sqlStr = sqlStr + " 	select m.socid as shopid, d.imakerid as makerid, sum(d.itemno*-1) as totcnt, sum(d.itemno*d.sellcash*-1) as totalsellcash, sum(d.itemno*d.buycash*-1) as totalbuycash from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " 	where m.code=d.mastercode"
		sqlStr = sqlStr + " 	and m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " 	and m.executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " 	and m.deldt is null"
		sqlStr = sqlStr + " 	and d.deldt is null"
		sqlStr = sqlStr + " 	and m.ipchulflag='S'"
		sqlStr = sqlStr + " 	and Left(m.socid,11)='streetshop8'"
		sqlStr = sqlStr + " 	and d.mwgubun='C'"
		sqlStr = sqlStr + " 	group by m.socid , d.imakerid"
		sqlStr = sqlStr + " ) as C"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d "
		sqlStr = sqlStr + " 	on C.shopid=d.shopid and C.makerid=d.makerid"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " 	on C.shopid=j.shopid"
		sqlStr = sqlStr + " 	and C.makerid=j.jungsanid"
		sqlStr = sqlStr + " 	and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " where d.chargediv<>'2'"
		sqlStr = sqlStr + " and d.chargediv<>'6'"
		sqlStr = sqlStr + " order by C.makerid, C.shopid "

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).Fshopid = rsget("shopid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotNo = rsget("totcnt")
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")

					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).Fdefaultmargin = rsget("defaultmargin")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetFranMeaipJungSanAutoList2()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select C.shopid , C.makerid, C.totcnt ,"
		sqlStr = sqlStr + " C.totalsellcash, "
		sqlStr = sqlStr + " C.totalbuycash,"
		sqlStr = sqlStr + " IsNULL(j.idx,'') as jungsanmasteridx,"
		sqlStr = sqlStr + " IsNULL(j.currstate,'') as currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " IsNull(j.chargediv,'') as jchargediv," 			'' (정산당시 정산 구분)
		sqlStr = sqlStr + " IsNull(d.chargediv,'') as chargediv, "			'' (현재 정산 구분)
		sqlStr = sqlStr + " IsNull(d.defaultmargin,0) as defaultmargin"		'' (현재 마진)
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " ( " '' 가맹점 특정 -> 매입 출고 내역
		sqlStr = sqlStr + " 	select m.socid as shopid, d.imakerid as makerid, sum(d.itemno) as totcnt, sum(d.itemno*d.sellcash) as totalsellcash, sum(d.itemno*d.buycash) as totalbuycash from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " 	where m.code=d.mastercode"
		sqlStr = sqlStr + " 	and m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " 	and m.executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " 	and m.deldt is null"
		sqlStr = sqlStr + " 	and d.deldt is null"
		sqlStr = sqlStr + " 	and m.ipchulflag='I'"
		sqlStr = sqlStr + " 	and m.divcode='801'"
		sqlStr = sqlStr + " 	group by m.socid , d.imakerid"
		sqlStr = sqlStr + " ) as C"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d "
		sqlStr = sqlStr + " 	on C.shopid=d.shopid and C.makerid=d.makerid"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " 	on j.shopid='streetshop800'"
		sqlStr = sqlStr + " 	and C.makerid=j.jungsanid"
		sqlStr = sqlStr + " 	and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " order by C.makerid, C.shopid "


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).Fshopid = "streetshop800"
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotNo = rsget("totcnt")
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")

					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")

					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")

					FItemList(i).Fdefaultmargin = rsget("defaultmargin")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end sub

	public Sub GetFranMeaipJungSanAutoList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = "select  targetid,count(m.idx) as totcnt,"
		sqlStr = sqlStr + " sum(m.totalsellcash) as totalsellcash, sum(m.totalbuycash) as totalbuycash,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv,"
		sqlStr = sqlStr + " d.chargediv, d.defaultmargin"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.jungsanid=m.targetid"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " and j.shopid='streetshop800'"

			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
			sqlStr = sqlStr + " on d.shopid='streetshop800'"
			sqlStr = sqlStr + " and d.makerid=m.targetid"

		sqlStr = sqlStr + " where m.baljuid='10x10'"
		sqlStr = sqlStr + " and m.statecd='9'"

		sqlStr = sqlStr + " and m.ipgodate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and m.ipgodate<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.divcode in ('101','131')"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " group by m.targetid, j.idx, j.currstate, j.totitemcnt"
		sqlStr = sqlStr + " ,j.totsum, j.minuscharge, j.realjungsansum, j.chargediv"
		sqlStr = sqlStr + " ,d.chargediv, d.defaultmargin"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).Fchargeuser = rsget("targetid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotNo = rsget("totcnt")
					FItemList(i).FTotalSellcash = rsget("totalsellcash")
					FItemList(i).FTotalBuyCash  = rsget("totalbuycash")

					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")

					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")

					FItemList(i).Fdefaultmargin = rsget("defaultmargin")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end Sub

	public Sub GetOffJungSanAutoList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select T.*, u.chargediv,u.autojungsan,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " j.totitemcnt as jungsantotitemcnt,"
		sqlStr = sqlStr + " j.totsum as jungsantotsum,"
		sqlStr = sqlStr + " IsNULL(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNULL(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv"
		sqlStr = sqlStr + " from ("

		sqlStr = sqlStr + " 	select m.shopid, d.makerid, "
		sqlStr = sqlStr + " 	count(itemno) as totno, sum(realsellprice*itemno) as totsum"

		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_detail d"

		sqlStr = sqlStr + " 	where m.idx=d.masteridx"
		sqlStr = sqlStr + " 	and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " 	and m.shopregdate<'" + nextYYYYMMDD + "'"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn='N'"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " 	group by m.shopid,d.makerid"
		sqlStr = sqlStr + " ) as T"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.jungsanid=T.makerid"
		sqlStr = sqlStr + " and j.shopid=T.shopid"
		sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
		end if

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer u"
		sqlStr = sqlStr + " on u.makerid=T.makerid and u.shopid=T.shopid "

		sqlStr = sqlStr + " where j.idx is null"

		if FRectChargeDiv<>"" then
			sqlStr = sqlStr + " and u.chargediv='" + FRectChargeDiv + "'"
		end if

		sqlStr = sqlStr + " order by T.shopid, u.chargediv desc, T.totsum desc, T.totno desc"


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanItem
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("makerid")
					'FItemList(i).Fchargename = rsget("chargename")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno      = rsget("totno")
					FItemList(i).FSum        = rsget("totsum")
					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")

					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).FAutojungsan = rsget("autojungsan")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetFranWitakJungSanAutoList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = sqlStr + " select d.makerid, d.shopid,"
		sqlStr = sqlStr + " IsNULL(T.totno,0) as totno, IsNULL(T.totsum,0) as totsum,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv,"
		sqlStr = sqlStr + " d.chargediv, d.defaultmargin"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d"
			sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
			sqlStr = sqlStr + " on j.jungsanid=d.makerid"
			sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
			sqlStr = sqlStr + " and Left(j.shopid,11) ='streetshop8'"
			sqlStr = sqlStr + " and j.shopid=d.shopid"

			sqlStr = sqlStr + " left join ("
			sqlStr = sqlStr + " select d.makerid, m.shopid, count(itemno) as totno, sum(sellprice*itemno) as totsum "
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
			sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"
			sqlStr = sqlStr + " and Left(m.shopid,11) ='streetshop8'"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
			if FRectDesigner<>"" then
				sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
			end if
			sqlStr = sqlStr + " group by d.makerid, m.shopid"
			sqlStr = sqlStr + " ) T on T.makerid=d.makerid and T.shopid=d.shopid"
		sqlStr = sqlStr + " where Left(d.shopid,11) ='streetshop8'"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		end if

	    sqlStr = sqlStr + " and (T.totno<>0 or j.totitemcnt<>0)"
		sqlStr = sqlStr + " and d.chargediv in ('2','6')"
		sqlStr = sqlStr + " order by d.makerid, d.shopid"


		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CFranJungSanItem
					FItemList(i).FjungsanMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).FTotNo = rsget("totno")
					FItemList(i).FTotalSellcash = rsget("totsum")
					'FItemList(i).FTotalBuyCash  = 0
					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					'FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					FItemList(i).Fdefaultmargin	= rsget("defaultmargin")
					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).Fshopid = rsget("shopid")



					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetOffJungSanList()
		dim i,sqlStr
		dim nextYYYYMMDD

		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select S.shopid, S.makerid, IsNULL(T.totno,0) as totno, IsNULL(T.totsum,0) as totsum,"
		sqlStr = sqlStr + " IsNULL(T2.totno,0) as totchulgono, IsNULL(T2.totsum,0) as totchulgosum,"
		sqlStr = sqlStr + " S.chargediv,S.autojungsan,"
		sqlStr = sqlStr + " j.idx as jungsanmasteridx,j.currstate,"
		sqlStr = sqlStr + " IsNull(j.totitemcnt,0) as jungsantotitemcnt,"
		sqlStr = sqlStr + " IsNull(j.totsum,0) as jungsantotsum,"
		sqlStr = sqlStr + " IsNull(j.minuscharge,0) as minuscharge,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum,"
		sqlStr = sqlStr + " j.chargediv as jchargediv, u.defaultmargine, u.maeipdiv "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer S"

		'' 판매내역
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select m.shopid, d.makerid, "
		sqlStr = sqlStr + " 	count(itemno) as totno, sum(realsellprice*itemno) as totsum"

		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shopjumun_detail d"

		sqlStr = sqlStr + " 	where m.idx=d.masteridx"
		sqlStr = sqlStr + " 	and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " 	and m.shopregdate<'" + nextYYYYMMDD + "'"

		sqlStr = sqlStr + " 	and m.cancelyn='N'"
		sqlStr = sqlStr + " 	and d.cancelyn='N'"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		end if
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " 	group by m.shopid,d.makerid"
		sqlStr = sqlStr + " ) as T on S.shopid=T.shopid and S.makerid=T.makerid"

		'' 출고내역
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select  m.socid, d.imakerid,"
		sqlStr = sqlStr + " 	count(d.itemno*-1) as totno, sum(d.buycash*d.itemno*-1) as totsum"
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " 	where m.code=d.mastercode"
		sqlStr = sqlStr + " 	and m.ipchulflag='S'"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " 	and m.socid='" + FRectShopID + "'"
		end if
		sqlStr = sqlStr + " 	and m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " 	and m.executedt<'" + nextYYYYMMDD + "'"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " 	and d.imakerid='" + FRectDesigner + "'"
		end if
		sqlStr = sqlStr + " 	and m.deldt is null"
		sqlStr = sqlStr + " 	and d.deldt is null"
		sqlStr = sqlStr + " 	group by m.socid, d.imakerid"
		sqlStr = sqlStr + " ) as T2 on S.shopid=T2.socid and S.makerid=T2.imakerid"

		'' 정산내역
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.jungsanid=S.makerid"
		sqlStr = sqlStr + " and j.shopid=S.shopid"
		sqlStr = sqlStr + " and j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and j.shopid='" + FRectShopID + "'"
		end if
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and j.jungsanid='" + FRectDesigner + "'"
		end if

		'' 온라인마진
		sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c u"
		sqlStr = sqlStr + " on S.makerid=u.userid "

		'' 쿼리조건
		sqlStr = sqlStr + " where 1=1"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and S.shopid='" + FRectShopID + "'"
		end if
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and S.makerid='" + FRectDesigner + "'"
		end if
		if FRectOnlymijungsan="on" then
			sqlStr = sqlStr + " and j.currstate is NULL"
		end if
		if FRectnomeachul="on" then
			sqlStr = sqlStr + " and IsNULL(T.totsum,0)<>0"
		end if

		sqlStr = sqlStr + " order by S.shopid, S.chargediv desc, T.totsum desc, T.totno desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJungSanItem
					FItemList(i).FShopid	 = rsget("shopid")
					FItemList(i).Fchargeuser = rsget("makerid")
					FItemList(i).FYYYYMM     = FRectJungsanYYYY + "-" + FRectJungsanMM
					FItemList(i).Ftotno      = rsget("totno")
					FItemList(i).FSum        = rsget("totsum")
					FItemList(i).Ftotchulgono     = rsget("totchulgono")
					FItemList(i).FtotchulgoSum    = rsget("totchulgosum")
					FItemList(i).Fjungsantotitemcnt = rsget("jungsantotitemcnt")
					FItemList(i).Fjungsantotsum     = rsget("jungsantotsum")
					FItemList(i).FjungsaMasterIdx = rsget("jungsanmasteridx")
					FItemList(i).Fcurrstate = rsget("currstate")
					FItemList(i).FminusCharge    = rsget("minuscharge")
					FItemList(i).FRealjungsansum = rsget("realjungsansum")
					FItemList(i).Fchargediv = rsget("chargediv")
					FItemList(i).FJungsanChargediv = rsget("jchargediv")
					FItemList(i).FAutojungsan = rsget("autojungsan")
					FItemList(i).Fonlinedefaultmargine 	= rsget("defaultmargine")
					FItemList(i).Fonlinemaeipdiv		= rsget("maeipdiv")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetCurrentJaeGoMinusList()
		dim i,sqlStr
		''#### 입출고리스트 - 판매리스트
		sqlStr = " select i.itemgubun, i.shopitemid, i.itemoption, i.shopitemname, i.shopitemoptionname,"
		sqlStr = sqlStr + " IsNull(S.itemno,0) as ipchulno, IsNull(T.itemno,0) as sellno,"
		sqlStr = sqlStr + " i.makerid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.shopitemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m, [db_shop].[dbo].tbl_shop_ipchul_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			if FRectDesigner<>"" then
				sqlStr = sqlStr + " and d.designerid='" + FRectDesigner + "'"
				sqlStr = sqlStr + " and m.chargeid='" + FRectDesigner + "'"
			end if

			if FRectShopID<>"" then
				sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
			end if
			sqlStr = sqlStr + " and m.deleteyn='N'"
			sqlStr = sqlStr + " and d.deleteyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.shopitemid, d.itemoption) as S"
		sqlStr = sqlStr + " on i.itemgubun=S.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=S.shopitemid"
		sqlStr = sqlStr + " and i.itemoption=S.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.itemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m, [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			if FRectDesigner<>"" then
				sqlStr = sqlStr + " and d.jungsan='" + FRectDesigner + "'"
			end if
			if FRectShopID<>"" then
				sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
			end if
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.itemid, d.itemoption) as T"
		sqlStr = sqlStr + " on i.itemgubun=T.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=T.itemid"
		sqlStr = sqlStr + " and i.itemoption=T.itemoption"
		sqlStr = sqlStr + " where i.shopitemid<>0"
		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and i.makerid='" + FRectDesigner + "'"
		end if

		sqlStr = sqlStr + " and (IsNull(S.itemno,0)<>0 or IsNull(T.itemno,0)<>0)"
		sqlStr = sqlStr + " and ((IsNull(S.itemno,0)-IsNull(T.itemno,0))<" + CStr(FRectJaegoNo) + ")"
		sqlStr = sqlStr + " order by i.makerid"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJaeGo
					FItemList(i).Fitemgubun       = rsget("itemgubun")
					FItemList(i).Fitemid          = rsget("shopitemid")
					FItemList(i).Fitemoption      = rsget("itemoption")
					FItemList(i).Fitemname        = db2Html(rsget("shopitemname"))
					FItemList(i).Fitemoptionname  = db2Html(rsget("shopitemoptionname"))
					FItemList(i).FIpChulNo        = rsget("ipchulno")
					FItemList(i).FSellNo       	  = rsget("sellno")
					FItemList(i).FJaeGo			  = FItemList(i).FIpChulNo - FItemList(i).FSellNo
					FItemList(i).FMakerID		  = rsget("makerid")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetCurrentJaeGoList1()
		dim i,sqlStr
		''#### 입출고리스트 - 판매리스트
		sqlStr = " select i.itemgubun, i.shopitemid, i.itemoption, i.shopitemname, i.shopitemoptionname,"
		sqlStr = sqlStr + " IsNull(S.itemno,0) as ipchulno, IsNull(T.itemno,0) as sellno"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.shopitemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m, [db_shop].[dbo].tbl_shop_ipchul_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and d.designerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and m.deleteyn='N'"
			sqlStr = sqlStr + " and d.deleteyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.shopitemid, d.itemoption) as S"
		sqlStr = sqlStr + " on i.itemgubun=S.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=S.shopitemid"
		sqlStr = sqlStr + " and i.itemoption=S.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.itemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m, [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.itemid, d.itemoption) as T"
		sqlStr = sqlStr + " on i.itemgubun=T.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=T.itemid"
		sqlStr = sqlStr + " and i.itemoption=T.itemoption"

		sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and (IsNull(S.itemno,0)<>0 or IsNull(T.itemno,0)<>0)"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJaeGo
					FItemList(i).Fitemgubun       = rsget("itemgubun")
					FItemList(i).Fitemid          = rsget("shopitemid")
					FItemList(i).Fitemoption      = rsget("itemoption")
					FItemList(i).Fitemname        = db2Html(rsget("shopitemname"))
					FItemList(i).Fitemoptionname  = db2Html(rsget("shopitemoptionname"))
					FItemList(i).FIpChulNo        = rsget("ipchulno")
					FItemList(i).FSellNo       	  = rsget("sellno")
					FItemList(i).FJaeGo			  = FItemList(i).FIpChulNo - FItemList(i).FSellNo
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetCurrentJaeGoList2()
		dim i,sqlStr
		''#### 입출고리스트 - 판매리스트
		sqlStr = " select i.itemgubun, i.shopitemid, i.itemoption, i.shopitemname, i.shopitemoptionname,"
		sqlStr = sqlStr + " IsNull(S.itemno,0) as ipchulno, IsNull(T.itemno,0) as sellno"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.shopitemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m, [db_shop].[dbo].tbl_shop_ipchul_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and d.designerid='" + FRectDesigner + "'"
			'sqlStr = sqlStr + " and m.chargeid='" + FRectDesigner + "'"
			if FRectShopid<>"" then
				sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
			end if
			sqlStr = sqlStr + " and m.deleteyn='N'"
			sqlStr = sqlStr + " and d.deleteyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.shopitemid, d.itemoption) as S"
		sqlStr = sqlStr + " on i.itemgubun=S.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=S.shopitemid"
		sqlStr = sqlStr + " and i.itemoption=S.itemoption"
		sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " ( select d.itemgubun, d.itemid, d.itemoption, sum(d.itemno) as itemno"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m, [db_shop].[dbo].tbl_shopjumun_detail d"
			sqlStr = sqlStr + " where m.idx=d.masteridx"
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
			if FRectShopid<>"" then
				sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
			end if
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
			sqlStr = sqlStr + " group by d.itemgubun, d.itemid, d.itemoption) as T"
		sqlStr = sqlStr + " on i.itemgubun=T.itemgubun"
		sqlStr = sqlStr + " and i.shopitemid=T.itemid"
		sqlStr = sqlStr + " and i.itemoption=T.itemoption"

		sqlStr = sqlStr + " where i.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and i.isusing='Y'"
		sqlStr = sqlStr + " and (IsNull(S.itemno,0)<>0 or IsNull(T.itemno,0)<>0)"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopJaeGo
					FItemList(i).Fitemgubun       = rsget("itemgubun")
					FItemList(i).Fitemid          = rsget("shopitemid")
					FItemList(i).Fitemoption      = rsget("itemoption")
					FItemList(i).Fitemname        = db2Html(rsget("shopitemname"))
					FItemList(i).Fitemoptionname  = db2Html(rsget("shopitemoptionname"))
					FItemList(i).FIpChulNo        = rsget("ipchulno")
					FItemList(i).FSellNo       	  = rsget("sellno")
					FItemList(i).FJaeGo			  = FItemList(i).FIpChulNo - FItemList(i).FSellNo
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetBrandMaeipItemList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select m.shopid,m.idx as orderno, d.idx, d.sellcash as sellprice, d.itemno,d.sellcash as realsellprice, d.designerid as makerid,"
		sqlStr = sqlStr + " m.chargeid as jungsanid, d.itemgubun,d.shopitemid as itemid,d.itemoption, i.shopitemname as itemname,i.shopitemoptionname as itemoptionname, "
		sqlStr = sqlStr + " m.execdt as shopregdate, d.suplycash as suplyprice, oi.mwdiv as onlinemwdiv"

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m"
		sqlStr = sqlStr + " ,[db_shop].[dbo].tbl_shop_ipchul_detail d"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " on d.itemgubun=i.itemgubun and d.shopitemid=i.shopitemid and d.itemoption=i.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item oi"
		sqlStr = sqlStr + " on d.itemgubun='10' and d.shopitemid=oi.itemid"
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.execdt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and m.execdt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.shopid='" + FRectShopId + "'"
		sqlStr = sqlStr + " and m.chargeid='" + FRectJungsanID + "'"
		sqlStr = sqlStr + " and m.statecd>=7"
		sqlStr = sqlStr + " and m.deleteyn='N'"
		sqlStr = sqlStr + " and d.deleteyn='N'"
		sqlStr = sqlStr + " and d.itemno<>0"
		sqlStr = sqlStr + " order by d.idx"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem
					FItemList(i).FIdx			 = rsget("idx")
					FItemList(i).FShopID         = rsget("shopid")
					FItemList(i).FMakerID        = rsget("makerid")
					FItemList(i).Forderno         = rsget("orderno")
					FItemList(i).Fitemgubun      = rsget("itemgubun")
					FItemList(i).Fitemid         = rsget("itemid")
					FItemList(i).Fitemoption     = rsget("itemoption")
					FItemList(i).Fitemname       = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno         = rsget("itemno")
					FItemList(i).Fsellprice      = rsget("sellprice")
					FItemList(i).Frealsellprice  = rsget("realsellprice")
					FItemList(i).FShopregDate       = rsget("shopregdate")
					FItemList(i).Fsuplyprice	= rsget("suplyprice")

					FItemList(i).FOnlineMwDiv = rsget("onlinemwdiv")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub


	public Sub GetBrandWitak2MaeipItemList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select m.socid as shopid,m.code as orderno, d.id, d.sellcash as sellprice, d.itemno,d.sellcash as realsellprice,"
		sqlStr = sqlStr + " d.imakerid as makerid,"
		sqlStr = sqlStr + " d.imakerid as jungsanid, d.iitemgubun,d.itemid as itemid,d.itemoption, d.iitemname as itemname,"
		sqlStr = sqlStr + " d.iitemoptionname as itemoptionname, d.mwgubun, oi.mwdiv as onlinemwdiv,"
		sqlStr = sqlStr + " m.executedt as shopregdate, d.buycash as suplyprice"
		sqlStr = sqlStr + " from  [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " ,[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item oi"
		sqlStr = sqlStr + " on d.iitemgubun='10' and d.itemid=oi.itemid"
		sqlStr = sqlStr + " where m.code=d.mastercode"
		sqlStr = sqlStr + " and m.executedt>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-01'"
		sqlStr = sqlStr + " and m.executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " and d.imakerid='" + FRectJungsanID + "'"
		sqlStr = sqlStr + " and m.ipchulflag='S'"
		sqlStr = sqlStr + " and m.socid='" + FRectShopId + "'"
		sqlStr = sqlStr + " order by d.id "

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem

					FItemList(i).FIdx			 = rsget("id")
					FItemList(i).FShopID         = rsget("shopid")
					FItemList(i).FMakerID        = rsget("makerid")
					FItemList(i).Forderno         = rsget("orderno")
					FItemList(i).Fitemgubun      = rsget("iitemgubun")
					FItemList(i).Fitemid         = rsget("itemid")
					FItemList(i).Fitemoption     = rsget("itemoption")
					FItemList(i).Fitemname       = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno         = rsget("itemno")*-1
					FItemList(i).Fsellprice      = rsget("sellprice")
					FItemList(i).Frealsellprice  = rsget("realsellprice")
					FItemList(i).FShopregDate       = rsget("shopregdate")
					FItemList(i).Fsuplyprice	= rsget("suplyprice")
					FItemList(i).FOnlineMwDiv   = rsget("onlinemwdiv")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetBrandSellItemList()
		dim i,sqlStr
		dim nextYYYYMMDD
		nextYYYYMMDD = CStr(dateserial(Left(FRectJungsanYYYY,4),Right(FRectJungsanMM,2)+1,1))

		sqlStr = " select m.shopid,m.orderno, d.idx, d.sellprice, d.itemno,d.realsellprice, d.makerid,"
		sqlStr = sqlStr + " d.jungsanid, d.itemgubun,d.itemid,d.itemoption,d.itemname,d.itemoptionname, "
		sqlStr = sqlStr + " m.shopregdate, d.suplyprice, IsNULL(i.shopitemprice,0) shopitemprice, oi.mwdiv as onlinemwdiv"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		sqlStr = sqlStr + " ,[db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " on d.itemgubun=i.itemgubun and d.itemid=i.shopitemid and d.itemoption=i.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item oi"
		sqlStr = sqlStr + " on d.itemgubun='10' and d.itemid=oi.itemid"
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn='N'"
		sqlStr = sqlStr + " and m.shopregdate>='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "-" + "01'"
		sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.shopid='" + FRectShopId + "'"
		sqlStr = sqlStr + " and d.makerid='" + FRectJungsanID + "'"
		sqlStr = sqlStr + " order by d.idx"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem

					FItemList(i).FIdx			 = rsget("idx")
					FItemList(i).FShopID         = rsget("shopid")
					FItemList(i).FMakerID        = rsget("makerid")
					FItemList(i).FJungsanId        = rsget("jungsanid")
					FItemList(i).Forderno         = rsget("orderno")
					FItemList(i).Fitemgubun      = rsget("itemgubun")
					FItemList(i).Fitemid         = rsget("itemid")
					FItemList(i).Fitemoption     = rsget("itemoption")
					FItemList(i).Fitemname       = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno         = rsget("itemno")
					FItemList(i).Fsellprice      = rsget("sellprice")
					FItemList(i).Frealsellprice  = rsget("realsellprice")
					FItemList(i).FShopregDate       = rsget("shopregdate")
					FItemList(i).Fsuplyprice	= rsget("suplyprice")
					FItemList(i).Fcurrentitemprice = rsget("shopitemprice")
					FItemList(i).FOnlineMwDiv = rsget("onlinemwdiv")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close

	end Sub

	public Sub GetNotMatchSellChargeIDList()
		dim i,sqlStr
		sqlStr = " select top 500 d.makerid, d.idx, d.itemname, d.itemoption, d.itemno,"
		sqlStr = sqlStr + " d.realsellprice,m.shopid, i.chargeid, d.jungsanid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + " where m.idx=d.masteridx"
		if FRectShopid<>"" then
				sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
			end if
		sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		sqlStr = sqlStr + " and d.itemgubun=i.itemgubun"
		sqlStr = sqlStr + " and d.itemid=i.shopitemid"
		sqlStr = sqlStr + " and d.itemoption=i.itemoption"
		sqlStr = sqlStr + " and d.jungsanid<>i.chargeid"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm
					FItemList(i).Fidx  = rsget("idx")
					FItemList(i).FMakerid  = rsget("makerid")
					FItemList(i).FItemName = db2html(rsget("itemname"))
					FItemList(i).FCount = rsget("itemno")
					FItemList(i).FSum   = rsget("realsellprice")
					FItemList(i).FShopid= rsget("shopid")

					FItemList(i).FjungsanID   = rsget("chargeid")
					FItemList(i).FSelljungsanID = rsget("jungsanid")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetBrandShopSellSumList()
		dim i,sqlStr
		sqlStr = " select sum(d.itemno * d.realsellprice) as subtotal, count(m.idx) as cnt, "
		sqlStr = sqlStr + " m.shopid, c.chargeuser as isbrandshop"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_chargeuser c"

		sqlStr = sqlStr + " where m.idx=d.masteridx"
		sqlStr = sqlStr + " and d.jungsanid=c.chargeuser"

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.shopid='" + CStr(FRectShopid) + "'"
		end if

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		end if

		sqlStr = sqlStr + " group by shopid, c.chargeuser"
		sqlStr = sqlStr + " order by subtotal desc, cnt desc"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm
					FItemList(i).FMakerid  = rsget("isbrandshop")
					FItemList(i).FCount = rsget("cnt")
					FItemList(i).FSum   = rsget("subtotal")
					FItemList(i).FShopid= rsget("shopid")
					FItemList(i).FIsBrandShop= rsget("isbrandshop")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	'///admin/offshop/todayselldetail.asp
	public Sub GetDaylySellItemList()
		dim i,sqlStr ,sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(pp.tplcompanyid,'')=''"
	    end if
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

		if FRectNormalOnly="on" then
			sqlsearch = sqlsearch + " and m.cancelyn='N'"
			sqlsearch = sqlsearch + " and d.cancelyn='N'"
		end if

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectTerms<>"" then
			FRectStartDay = FRectTerms
			FRectEndDay   = Left(CStr(DateAdd("d",1,FRectStartDay)),10)
			sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if
		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and d.makerid='" + FRectDesigner + "'"
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		If frectoffcatecode <> "" Then
			sqlsearch = sqlsearch + " and p.offcatecode = '" + CStr(frectoffcatecode) + "' "
		End IF

		If frectoffmduserid <> "" Then
			sqlsearch = sqlsearch + " and p.offmduserid = '" + CStr(frectoffmduserid) + "' "
		End IF

		If frectitemid <> "" Then
			sqlsearch = sqlsearch + " and d.itemid = "&frectitemid&""
		End IF

		If FRectitemname <> "" Then
			sqlsearch = sqlsearch + " and d.itemname like '%"&FRectitemname&"%'"
		End IF

		If FRectextbarcode <> "" Then
			sqlsearch = sqlsearch + " and i.extbarcode = '"&FRectextbarcode&"'"
		End IF

		If FRectCommCD <> "" Then
			sqlsearch = sqlsearch + " and d.jcomm_cd= '"&FRectCommCD&"'"
		End IF

'if (session("ssBctID")="coolhas") then
'    sqlsearch = sqlsearch + " and d.jcomm_cd<>'B012'"
'    sqlsearch = sqlsearch + " and d.jcomm_cd<>'B013'"
'end if
		sqlStr = " select top 3000" & vbcrlf
		sqlStr = sqlStr + " sum(d.itemno * (d.realsellprice+d.addtaxcharge)) as subtotal" & vbcrlf
		sqlStr = sqlStr + " ,sum(d.itemno * d.suplyprice) as suplysum" & vbcrlf
		sqlStr = sqlStr + " ,sum(d.itemno * (d.sellprice+d.addtaxcharge)) as sellsum, sum(d.itemno) as itemno" & vbcrlf
		sqlStr = sqlStr + " ,sum(d.itemno*(d.realsellprice-d.suplyprice)) as profit" & vbcrlf
		sqlStr = sqlStr + " ,d.sellprice, d.realsellprice, d.addtaxcharge, d.itemname, d.itemoptionname" & vbcrlf
		sqlStr = sqlStr + " ,d.itemgubun, d.itemid, d.itemoption, d.makerid ,d.suplyprice , i.extbarcode" & vbcrlf
        sqlStr = sqlStr + " ,d.jcomm_cd, ii.smallimage, i.offimgsmall" & vbcrlf

		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m with (readuncommitted)" & vbcrlf
			sqlStr = sqlStr + " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d with (readuncommitted)" & vbcrlf
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m with (readuncommitted)" & vbcrlf
			sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d with (readuncommitted)" & vbcrlf
		end if
		sqlStr = sqlStr + " 	on m.idx=d.masteridx" & vbcrlf

		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_user u with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr + " 	on m.shopid = u.userid" & vbcrlf
	    sqlStr = sqlStr + " Join db_partner.dbo.tbl_partner p with (readuncommitted)" & vbcrlf
	    sqlStr = sqlStr + " on d.makerid=p.id" & vbcrlf

		if (FRectBrandPurchaseType<>"") then
			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				sqlStr = sqlStr + " 	and p.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				sqlStr = sqlStr & " 	and p.purchasetype in ('3','5','6')"
			else
				sqlStr = sqlStr + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		sqlStr = sqlStr + " left join [db_shop].dbo.tbl_shop_item i with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr + " 	on d.itemgubun = i.itemgubun and d.itemid = i.shopitemid" & vbcrlf
		sqlStr = sqlStr + " 	and d.itemoption = i.itemoption" & vbcrlf
		sqlStr = sqlStr + " left join db_item.dbo.tbl_item ii with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr + " 	on d.itemgubun = '10'" & vbcrlf
		sqlStr = sqlStr + " 	and d.itemid = ii.itemid" & vbcrlf
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner pp with (readuncommitted)" & vbcrlf
	    sqlStr = sqlStr & "       on m.shopid=pp.id " & vbcrlf
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " group by d.sellprice, d.realsellprice, d.addtaxcharge, d.itemname, d.itemoptionname," & vbcrlf
		sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.makerid ,d.suplyprice, i.extbarcode, d.jcomm_cd" & vbcrlf
		sqlStr = sqlStr + " , ii.smallimage, i.offimgsmall" & vbcrlf

		Select Case FRectOrdertype
			Case "totalprice"
				'매출순
				sqlStr = sqlStr & " order by subtotal Desc"
	    	Case "gain"
	    		'수익순
	            sqlStr = sqlStr & " order by profit Desc"
			Case "unitCost"
				'객단가순
				sqlStr = sqlStr & " order by d.sellprice Desc"
			Case Else
				sqlStr = sqlStr + " order by itemno desc ,subtotal desc"
		end Select

		''response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem

					FItemList(i).fsuplysum     = rsget("suplysum")
					FItemList(i).fsellsum     = rsget("sellsum")
					FItemList(i).fsuplyprice     = rsget("suplyprice")
					FItemList(i).Fitemgubun     = rsget("itemgubun")
					FItemList(i).Fitemid        = rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname      = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno        = rsget("itemno")
					FItemList(i).Fsellprice     = rsget("sellprice")
					FItemList(i).Frealsellprice = rsget("realsellprice")
					FItemList(i).Fsubtotal       = rsget("subtotal")
					FItemList(i).FMakerID		 = rsget("makerid")
					FItemList(i).fextbarcode		 = rsget("extbarcode")
					FItemList(i).faddtaxcharge   = rsget("addtaxcharge")
					FItemList(i).fjcomm_cd      = rsget("jcomm_cd")
					FItemList(i).fsmallimage      = rsget("smallimage")
					FItemList(i).foffimgsmall      = rsget("offimgsmall")
					if not(isnull(FItemList(i).fsmallimage)) and FItemList(i).fsmallimage<>"" then FItemList(i).fsmallimage=webImgUrl&"/image/small/"& GetImageSubFolderByItemid(rsget("itemid")) &"/"& FItemList(i).fsmallimage
					if not(isnull(FItemList(i).foffimgsmall)) and FItemList(i).foffimgsmall<>"" then FItemList(i).foffimgsmall=webImgUrl&"/offimage/offsmall/i"& rsget("itemgubun") &"/"& GetImageSubFolderByItemid(rsget("itemid")) &"/"& FItemList(i).foffimgsmall

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	'//사용안함
	public Sub GetDaylySellItemList3TimeBojung()
		dim i,sqlStr
		sqlStr = " select sum(d.itemno * (d.realsellprice+d.addtaxcharge)) as subtotal, sum(d.itemno) as itemno, "
		sqlStr = sqlStr + " d.sellprice, d.realsellprice, d.addtaxcharge,d.itemname, d.itemoptionname,"
		sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.makerid"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " where m.idx=d.masteridx"

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectTerms<>"" then
			sqlStr = sqlStr + " and convert(varchar(10),dateadd(hh,-5,m.shopregdate),20)='" + FRectTerms + "'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)<'" + CStr(FRectEndDay) + "'"
		end if

		if FRectJungsanId<>"" then
			sqlStr = sqlStr + " and d.jungsanid='" + FRectJungsanId + "'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		end if
		sqlStr = sqlStr + " group by d.sellprice, d.realsellprice, d.addtaxcharge,d.itemname, d.itemoptionname,"
		sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.makerid"
		sqlStr = sqlStr + " order by subtotal desc, itemno desc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem
					FItemList(i).Fitemgubun     = rsget("itemgubun")
					FItemList(i).Fitemid        = rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname      = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno        = rsget("itemno")
					FItemList(i).Fsellprice     = rsget("sellprice")
					FItemList(i).Frealsellprice = rsget("realsellprice")
					FItemList(i).Fsubtotal       = rsget("subtotal")
					FItemList(i).FMakerID		 = rsget("makerid")

					FItemList(i).Faddtaxcharge   = rsget("addtaxcharge")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	public Sub GetDaylySellItemListByShopByItem()
		dim i,sqlStr

		sqlStr = " select top 1000 convert(varchar(10),shopregdate,21) as yyyymmdd, m.shopid, d.itemgubun, d.itemid, d.itemoption, "
		sqlStr = sqlStr + "         d.sellprice, d.realsellprice, d.itemname, d.itemoptionname, d.makerid, "
		sqlStr = sqlStr + "         sum(d.itemno) as itemno "
		sqlStr = sqlStr + "         ,d.jcomm_cd, j.comm_name"

        if (FRectShowOrder <> "") then
            sqlStr = sqlStr + "         ,m.orderno "
        else
            sqlStr = sqlStr + "         , '' as orderno "
        end if

	        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
	        sqlStr = sqlStr + "     Join  [db_shop].[dbo].tbl_shopjumun_detail d "
	        sqlStr = sqlStr + "     on m.idx=d.masteridx "
	        sqlStr = sqlStr + "     left join db_jungsan.dbo.tbl_jungsan_comm_code j"
	        sqlStr = sqlStr + "     on d.jcomm_cd=j.comm_cd"
	        sqlStr = sqlStr + " where 1 = 1 "
	        sqlStr = sqlStr + " and m.cancelyn='N' "
	        sqlStr = sqlStr + " and d.cancelyn='N' "

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and shopregdate<'" + CStr(FRectEndDay) + "'"
		end if

        if FRectItemGubun<>"" then
			sqlStr = sqlStr + " and d.itemgubun='" + CStr(FRectItemGubun) + "'"
		end if

		if FRectItemId<>"" then
			sqlStr = sqlStr + " and d.itemid=" + CStr(FRectItemId)
		end if

        if FRectItemOption<>"" then
			sqlStr = sqlStr + " and d.itemoption='" + CStr(FRectItemOption) + "'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		end if

	    sqlStr = sqlStr + " group by convert(varchar(10),shopregdate,21), m.shopid, d.itemgubun, d.itemid, d.itemoption, d.sellprice, d.realsellprice, d.itemname, d.itemoptionname, d.makerid, d.jcomm_cd, j.comm_name "
        if (FRectShowOrder <> "") then
            sqlStr = sqlStr + "         ,m.orderno "
        end if
	    sqlStr = sqlStr + " order by convert(varchar(10),shopregdate,21), m.shopid, d.itemgubun, d.itemid, d.itemoption, d.sellprice, d.realsellprice, d.itemname, d.itemoptionname, d.makerid "
        if (FRectShowOrder <> "") then
            sqlStr = sqlStr + "         ,m.orderno "
        end if

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopSellDetailItem
					FItemList(i).ForderNo   	= rsget("orderno")
                    FItemList(i).Fshopregdate   = rsget("yyyymmdd")
					FItemList(i).Fshopid        = rsget("shopid")
					FItemList(i).Fitemgubun     = rsget("itemgubun")
					FItemList(i).Fitemid        = rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname      = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
					FItemList(i).Fitemno        = rsget("itemno")
					FItemList(i).Fsellprice     = rsget("sellprice")
					FItemList(i).Frealsellprice = rsget("realsellprice")
					FItemList(i).FMakerID	    = rsget("makerid")
					FItemList(i).Fjcomm_cd      = rsget("jcomm_cd")
					FItemList(i).Fcomm_name     = rsget("comm_name")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	'//사용안함
	public Sub GetDaylySellJumunList3TimeBojung()
		dim i,sqlStr
		sqlStr = " select top 3000 m.idx,m.orderno,m.totalsum,m.realsum, m.jumunmethod,m.shopregdate,m.pointuserno,"
		sqlStr = sqlStr + " d.itemname,d.itemoptionname,d.sellprice,d.realsellprice,d.itemno, d.makerid,"
		sqlStr = sqlStr + " IsNULL(cashsum,0) as cashsum, IsNULL(cardsum,0) as cardsum, IsNULL(giftcardpaysum,0) as giftcardPaysum"
		sqlStr = sqlStr + " ,IsNULL(d.addTaxCharge,0) as addTaxCharge"
		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m"
    		sqlStr = sqlStr + "     Join [db_shoplog].[dbo].tbl_old_shopjumun_detail d"
    		sqlStr = sqlStr + "     on m.idx=d.masteridx"
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
    		sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shopjumun_detail d"
    		sqlStr = sqlStr + "     on m.idx=d.masteridx"
		end if

		sqlStr = sqlStr + " where 1=1"

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
		end if

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectTerms<>"" then
			sqlStr = sqlStr + " and convert(varchar(10),dateadd(hh,-5,m.shopregdate),20)='" + FRectTerms + "'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)<'" + CStr(FRectEndDay) + "'"
		end if

		if FRectJungsanId<>"" then
			sqlStr = sqlStr + " and d.jungsanid='" + FRectJungsanId + "'"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		end if
		sqlStr = sqlStr + " order by m.idx, d.idx"
''rw sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopSellMasterDetailItem
					FItemList(i).Fidx             = rsget("idx")
					FItemList(i).ForderNo         = rsget("orderno")
					FItemList(i).Ftotalsum        = rsget("totalsum")
					FItemList(i).Frealsum         = rsget("realsum")
					FItemList(i).Fshopregdate	  = rsget("shopregdate")
					FItemList(i).Fjumunmethod        = rsget("jumunmethod")
					FItemList(i).Fitemname        = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
					FItemList(i).Fsellprice       = rsget("sellprice")
					FItemList(i).Frealsellprice   = rsget("realsellprice")
					FItemList(i).Fitemno          = rsget("itemno")
					FItemList(i).FMakerID		  = rsget("makerid")
					FItemList(i).Fpointuserno	  = rsget("pointuserno")

                    FItemList(i).Fcashsum           = rsget("cashsum")
                    FItemList(i).Fcardsum           = rsget("cardsum")
                    FItemList(i).FgiftcardPaysum    = rsget("giftcardPaysum")

                    FItemList(i).FaddTaxCharge      = rsget("addTaxCharge")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	'//admin/offshop/todaysellmaster.asp
	public Sub GetDaylySellJumunList()
		dim i,sqlStr , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

		if FRectNormalOnly="on" then
			sqlsearch = sqlsearch + " and m.cancelyn='N'"
			sqlsearch = sqlsearch + " and d.cancelyn='N'"
		end if

		if FRectShopid<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopid + "'"
		end if

		if FRectTerms<>"" then
			FRectStartDay = FRectTerms
			FRectEndDay   = Left(CStr(DateAdd("d",1,FRectStartDay)),10)
			sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if


		if FRectJungsanId<>"" then
			sqlsearch = sqlsearch + " and d.jungsanid='" + FRectJungsanId + "'"
		end if

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch + " and d.makerid='" + FRectDesigner + "'"
		end if

        if (FRectPgDataCheck="on") then
            sqlsearch = sqlsearch + " and m.shopid in (select shopid from db_shop.dbo.tbl_shopjumun_cardApp_log group by shopid)"
        end if

		if (FRectExcMatchFinish <> "") then
			sqlsearch = sqlsearch + " and m.orderno not in (select orderserial from db_shop.dbo.tbl_shopjumun_cardApp_log where orderserial is not NULL) "
		end if

		if (FRectCardPayOnly <> "") then
			'// 현금결제만 제외
			''sqlsearch = sqlsearch + " and m.jumunmethod <> '01' "
			sqlsearch = sqlsearch + " and m.cardsum <> 0 "
		end if

		if (FRectCardSum <> "") then
			sqlsearch = sqlsearch + " and m.cardsum = " + CStr(FRectCardSum) + " "
		end if

		if (FRectPaySum <> "") then
			sqlsearch = sqlsearch + " and (m.cardsum = " + CStr(FRectPaySum) + " or m.cashsum = " + CStr(FRectPaySum) + ") "
		end if



		sqlStr = " select top 4000 m.idx, m.orderno,m.totalsum,m.realsum, m.jumunmethod,m.shopregdate,m.pointuserno,"
		sqlStr = sqlStr + " d.itemname,d.itemoptionname,d.sellprice,d.realsellprice,d.itemno, d.makerid,"
		sqlStr = sqlStr + " IsNULL(cashsum,0) as cashsum, IsNULL(cardsum,0) as cardsum, IsNULL(giftcardpaysum,0) as giftcardPaysum, IsNull(m.spendmile,0) as spendmile, IsNull(m.TenGiftCardPaySum,0) as TenGiftCardPaySum,"
		sqlStr = sqlStr + " IsNULL(m.extPaySum,0) as extPaySum, "
		sqlStr = sqlStr + " m.shopid, m.reforderno, m.cardappno"

		if (FRectPgDataCheck="on") then
    		sqlStr = sqlStr + " , (select count(l.orderserial) as cnt from db_shop.dbo.tbl_shopjumun_cardApp_log l where l.shopid = m.shopid and l.orderserial = m.orderno) as matchCount "
    	end if

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " 	on m.idx=d.masteridx"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & " 	on m.shopid=p.id "
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		''sqlStr = sqlStr + " order by m.idx , d.idx"
		sqlStr = sqlStr + " order by m.orderno , d.idx"

		''response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new COffShopSellMasterDetailItem

					FItemList(i).Fidx             = rsget("idx")
					FItemList(i).ForderNo         = rsget("orderno")
					FItemList(i).Ftotalsum        = rsget("totalsum")
					FItemList(i).Frealsum         = rsget("realsum")
					FItemList(i).Fshopregdate	  = rsget("shopregdate")
					FItemList(i).Fjumunmethod        = rsget("jumunmethod")
					FItemList(i).Fitemname        = db2html(rsget("itemname"))
					FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
					FItemList(i).Fsellprice       = rsget("sellprice")
					FItemList(i).Frealsellprice   = rsget("realsellprice")
					FItemList(i).Fitemno          = rsget("itemno")
					FItemList(i).FMakerID		  = rsget("makerid")
					FItemList(i).Fpointuserno		  = rsget("pointuserno")
                    FItemList(i).Fcashsum           = rsget("cashsum")
                    FItemList(i).Fcardsum           = rsget("cardsum")
                    FItemList(i).FgiftcardPaysum    = rsget("giftcardPaysum")
					FItemList(i).FextPaysum			= rsget("extPaySum")

					FItemList(i).Fspendmile    		= rsget("spendmile")
					FItemList(i).FTenGiftCardPaySum = rsget("TenGiftCardPaySum")

                    FItemList(i).Fshopid            = rsget("shopid")
                    FItemList(i).Freforderno        = rsget("reforderno")
                    FItemList(i).Fcardappno         = rsget("cardappno")
                    if (FRectPgDataCheck="on") then
    					FItemList(i).FmatchCount    	= rsget("matchCount")
                    end if
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

	'//designer/offshop/selllist.asp	'/사용안함
	public Sub GetDaylySumListByJungsanID()
		dim i,sqlStr
		sqlStr = " select top 100 sum(d.itemno * d.realsellprice) as sellsum, count(m.idx) as cnt,m.shopid "

		'//주문일 기준
		if frectdatefg = "jumun" then
			sqlStr = sqlStr + " ,convert(varchar(10),m.shopregdate,20) as selldate"
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sqlStr = sqlStr + " ,(m.IXyyyymmdd) as selldate"
		else
			sqlStr = sqlStr + " ,convert(varchar(10),m.shopregdate,20) as selldate"
		end if

		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
		sqlStr = sqlStr + " where m.idx=d.masteridx"

		sqlStr = sqlStr + " and d.makerid='" + FRectJungsanID + "'"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and m.cancelyn='N'"
			sqlStr = sqlStr + " and d.cancelyn='N'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlStr = sqlStr + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlStr = sqlStr + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

		else
			if FRectStartDay<>"" then
				sqlStr = sqlStr + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlStr = sqlStr + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if

		sqlStr = sqlStr + " group by m.shopid "

		'//주문일 기준
		if frectdatefg = "jumun" then
			sqlStr = sqlStr + " ,convert(varchar(10),m.shopregdate,20)"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sqlStr = sqlStr + " ,m.IXyyyymmdd"
		else
			sqlStr = sqlStr + " ,convert(varchar(10),m.shopregdate,20)"
		end if

		sqlStr = sqlStr + " order by m.shopid, selldate desc"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)

			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm

					FItemList(i).FTerm  = rsget("selldate")
					FItemList(i).FCount = rsget("cnt")
					FItemList(i).FSum   = rsget("sellsum")
					FItemList(i).FShopid= rsget("shopid")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	'/admin/offshop/sellreportday.asp
	public Sub GetDaylySumList()
		dim i,sqlStr , sqlsearch , sqlsearch1, sqldategubun , sqldategubungroup

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch1 = sqlsearch1 & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch1 = sqlsearch1 & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

        '//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " 	and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " 	and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			sqldategubun = sqldategubun & " (convert(varchar(10),m.shopregdate,121)) as selldate"
			sqldategubungroup = sqldategubungroup & " convert(varchar(10),m.shopregdate,121)"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " 	and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " 	and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

			sqldategubun = sqldategubun & " (m.IXyyyymmdd) as selldate"
			sqldategubungroup = sqldategubungroup & " m.IXyyyymmdd"
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " 	and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " 	and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			sqldategubun = sqldategubun & " (convert(varchar(10),m.shopregdate,121)) as selldate"
			sqldategubungroup = sqldategubungroup & " convert(varchar(10),m.shopregdate,121)"
		end if

        if FRectShopID<>"" then
			sqlsearch = sqlsearch & " 	and m.shopid='" + FRectShopID + "'"
		end if
		if FRectOnlyShop<>"" then
			sqlsearch = sqlsearch & " 	and Left(m.shopid,4)<>'cafe'"
		end if
        if FRectmakerid<>"" then
			sqlsearch = sqlsearch & " 	and d.makerid='" + FRectmakerid + "'"
		end if

        sqlStr = " SELECT top " + CStr(FPageSize*FCurrPage)
        sqlStr = sqlStr & "	a.selldate, a.shopid, a.itemcnt, a.sellsum, a.suplyprice ,c.targetmaechul"
        sqlStr = sqlStr & "	,a.ordercnt as cnt"
        sqlStr = sqlStr & "	,(convert(varchar(1),wd.weather) + '||' + convert(varchar(1000),wd.comment)) as weather"
        'sqlStr = sqlStr & " ,convert(varchar(100),([db_shop].[dbo].[uf_getWeather](A.selldate, A.shopid))) as weather"
        sqlStr = sqlStr & "	from ("
        sqlStr = sqlStr & "		select"
        sqlStr = sqlStr & "		"&sqldategubun&",m.shopid ,sum(d.itemno) as itemcnt"
        sqlStr = sqlStr & "		,isnull(sum((d.realsellprice+isnull(d.addtaxcharge,0))*d.itemno),0) as sellsum"
        sqlStr = sqlStr & "		,sum(d.suplyprice*d.itemno) as suplyprice"
        sqlStr = sqlStr & "		,count(distinct m.idx) as ordercnt"

		if FRectOldData="on" then
			sqlStr = sqlStr + " 	from [db_shoplog].[dbo].tbl_old_shopjumun_master m"
	        sqlStr = sqlStr & "		join [db_shoplog].[dbo].tbl_old_shopjumun_detail d"
		else
	        sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master m"
	        sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail d"
		end if

        sqlStr = sqlStr & "			on m.idx=d.masteridx and m.cancelyn='N' and d.cancelyn='N'"
        sqlStr = sqlStr & "		where 1=1 " & sqlsearch
        sqlStr = sqlStr & "		group by "&sqldategubungroup&" ,m.shopid"
        sqlStr = sqlStr & "	) A"
        sqlStr = sqlStr & "	left join db_shop.dbo.tbl_targetmaechul_month_off c"
        sqlStr = sqlStr & "		on convert(varchar(7),a.selldate,121) = c.yyyymm"
        sqlStr = sqlStr & "		and a.shopid = c.shopid and c.gubuntype = 1 and c.gubun = 0"
        sqlStr = sqlStr & "	left join db_shop.dbo.tbl_shop_weather wd"
        sqlStr = sqlStr & "		on A.selldate=wd.wdate"
        sqlStr = sqlStr & "		and A.shopid=wd.shopid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on a.shopid=p.id "
        sqlStr = sqlStr & "	where 1=1 " & sqlsearch1
        sqlStr = sqlStr & "	order by A.shopid asc, A.selldate asc"

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm

					FItemList(i).ftargetmaechul  = rsget("targetmaechul")
					FItemList(i).FTerm  = rsget("selldate")
					FItemList(i).FSum   = rsget("sellsum")
					FItemList(i).FShopid= rsget("shopid")
                    FItemList(i).fsuplyprice = rsget("suplyprice")
                    FItemList(i).FCount = rsget("cnt")
                    FItemList(i).FWeather = rsget("weather")
                    FItemList(i).fitemcnt = rsget("itemcnt")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	'//admin/offshop/sellreportdanga.asp
	public Sub GetReportByDanga()
		dim i,sqlStr, sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectStartDay<>"" then
			sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
		end if
		if FRectEndDay<>"" then
			sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
		end if
		if FRectShopID<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopID + "'"
		end if

		sqlStr = " select"
		sqlStr = sqlStr + " (case"
		sqlStr = sqlStr + " 	when m.realsum < 10000 then '0~10,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 10000 and m.realsum < 20000 then '10,000~20,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 20000 and m.realsum < 30000 then '20,000~30,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 30000 and m.realsum < 40000 then '30,000~40,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 40000 and m.realsum < 50000 then '40,000~50,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 50000 and m.realsum < 60000 then '50,000~60,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 60000 and m.realsum < 70000 then '60,000~70,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 70000 and m.realsum < 80000 then '70,000~80,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 80000 and m.realsum < 90000 then '80,000~90,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 90000 and m.realsum < 100000 then '90,000~100,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 100000 and m.realsum < 150000 then 'z100,000~150,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 150000 and m.realsum < 200000 then 'z150,000~200,000'"
		sqlStr = sqlStr + " 	else 'z200,0000~' end) as gubun"
		sqlStr = sqlStr + " ,count(m.idx) as cnt, sum(m.realsum) as sellsum, sum(m.spendmile) as spendmile"

		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m "
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		end if

		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr + "       on m.shopid=p.id "
		sqlStr = sqlStr + " where m.idx<>0"
		sqlStr = sqlStr + " and m.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr + " group by (case"
		sqlStr = sqlStr + " 	when m.realsum < 10000 then '0~10,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 10000 and m.realsum < 20000 then '10,000~20,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 20000 and m.realsum < 30000 then '20,000~30,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 30000 and m.realsum < 40000 then '30,000~40,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 40000 and m.realsum < 50000 then '40,000~50,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 50000 and m.realsum < 60000 then '50,000~60,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 60000 and m.realsum < 70000 then '60,000~70,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 70000 and m.realsum < 80000 then '70,000~80,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 80000 and m.realsum < 90000 then '80,000~90,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 90000 and m.realsum < 100000 then '90,000~100,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 100000 and m.realsum < 150000 then 'z100,000~150,000'"
		sqlStr = sqlStr + " 	when m.realsum >= 150000 and m.realsum < 200000 then 'z150,000~200,000'"
		sqlStr = sqlStr + " 	else 'z200,0000~' end)"
		sqlStr = sqlStr + " order by gubun asc"

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1

		maxt =0
		maxc =0
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm

					FItemList(i).fspendmile = rsget("spendmile")
					FItemList(i).FTerm  = rsget("gubun")
					FItemList(i).FCount = rsget("cnt")
					FItemList(i).FSum   = rsget("sellsum")

					maxc = maxc + FItemList(i).FCount
					maxt = maxt + (FItemList(i).FSum + FItemList(i).fspendmile)

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	'//사용안함
	public Sub GetDaylySumList3TimeBojung()
		dim i,sqlStr
		sqlStr = " select top 500 count(m.idx) as cnt, sum(m.realsum) as sellsum "
		sqlStr = sqlStr + " ,sum(IsNull(spendmile,0)) as spendmilesum, sum(IsNull(gainmile,0)) as gainmilesum "
		sqlStr = sqlStr + " ,sum(case when m.jumunmethod='01' then m.realsum when m.jumunmethod='03' then m.cashsum else 0 end) as cashsum"
		sqlStr = sqlStr + " ,sum(case when m.jumunmethod='02' then m.realsum when m.jumunmethod='03' then m.cardsum else 0 end) as cardsum"
		sqlStr = sqlStr + " ,sum(case when m.jumunmethod='03' then m.giftcardPaysum else 0 end) as giftcardpaysum"
		sqlStr = sqlStr + " ,convert(varchar(10),dateadd(hh,-5,m.shopregdate),20) as selldate, m.shopid"
		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m "
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m"
		end if

		sqlStr = sqlStr + " where m.idx<>0"
		if FRectShopID<>"" then
			sqlStr = sqlStr + " and m.shopid='" + FRectShopID + "'"
		end if

		if FRectOnlyShop<>"" then
			sqlStr = sqlStr + " and Left(m.shopid,4)<>'cafe'"
		end if

		sqlStr = sqlStr + " and m.cancelyn='N'"

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and dateadd(hh,-5,m.shopregdate)<'" + CStr(FRectEndDay) + "'"
		end if

		sqlStr = sqlStr + " group by convert(varchar(10),dateadd(hh,-5,m.shopregdate),20), m.shopid"
		sqlStr = sqlStr + " order by m.shopid, m.selldate desc"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new COffShopSellByTerm
					FItemList(i).FShopName = "취화선"
					FItemList(i).FTerm  = rsget("selldate")
					FItemList(i).FCount = rsget("cnt")
					FItemList(i).FSum   = rsget("sellsum")
					FItemList(i).FShopid= rsget("shopid")
					FItemList(i).FSpendMile = rsget("spendmilesum")
					FItemList(i).FGainMile = rsget("gainmilesum")

					FItemList(i).FCashSum       = rsget("CashSum")
					FItemList(i).FCardSum       = rsget("CardSum")
					FItemList(i).FGiftCardPaysum= rsget("GiftCardPaysum")
					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end Sub

	public Sub GetJumunMasterList()
		dim i,sqlStr
		sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_shopjumun_master"
		sqlStr = sqlStr + " where idx<>0"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and cancelyn='N'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and shopregdate<'" + CStr(FRectEndDay) + "'"
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select idx,orderno,shopid,totalsum,realsum,jumundiv,jumunmethod,"
		sqlStr = sqlStr + " shopregdate,cancelyn,regdate,shopidx"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master"
		sqlStr = sqlStr + " where idx<>0"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and shopid='" + FRectShopID + "'"
		end if

		if FRectNormalOnly="on" then
			sqlStr = sqlStr + " and cancelyn='N'"
		end if

		if FRectStartDay<>"" then
			sqlStr = sqlStr + " and shopregdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			sqlStr = sqlStr + " and shopregdate<'" + CStr(FRectEndDay) + "'"
		end if
		sqlStr = sqlStr + " order by shopregdate desc"

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
					set FItemList(i) = new COffShopSellMasterItem

					FItemList(i).Fidx        = rsget("idx")
					FItemList(i).Forderno    = rsget("orderno")
					FItemList(i).Fshopid     = rsget("shopid")
					FItemList(i).Ftotalsum   = rsget("totalsum")
					FItemList(i).Frealsum    = rsget("realsum")
					FItemList(i).Fjumundiv   = rsget("jumundiv")
					FItemList(i).Fjumunmethod= rsget("jumunmethod")
					FItemList(i).Fshopregdate= rsget("shopregdate")
					FItemList(i).Fcancelyn   = rsget("cancelyn")
					FItemList(i).Fregdate    = rsget("regdate")
					FItemList(i).Fshopidx    = rsget("shopidx")

					i=i+1
					rsget.moveNext
				loop
			end if
		rsget.Close
	end sub

'	'//admin/offshop/offcardlist.asp	'/사용안함
'	public Sub GetDaylyUserCardList()
'		dim i,sqlStr
'		sqlStr = " select top 100 count(m.idx) as cnt, sum(IsNull(shoppoint,0)) as shoppoint, "
'		sqlStr = sqlStr + " convert(varchar(10),m.regdate,20) as regdate, m.regshopid"
'		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_pointuser m"
'		sqlStr = sqlStr + " where m.idx<>0"
'		if FRectShopID<>"" then
'			sqlStr = sqlStr + " and m.regshopid='" + FRectShopID + "'"
'		end if
'
'		if FRectStartDay<>"" then
'			sqlStr = sqlStr + " and m.regdate>='" + CStr(FRectStartDay) + "'"
'		end if
'
'		if FRectEndDay<>"" then
'			sqlStr = sqlStr + " and m.regdate<'" + CStr(FRectEndDay) + "'"
'		end if
'
'		sqlStr = sqlStr + " group by convert(varchar(10),m.regdate,20), m.regshopid"
'		sqlStr = sqlStr + " order by m.regshopid, m.regdate desc"
'
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'		FResultCount = rsget.RecordCount
'			redim preserve FItemList(FResultCount)
'			i=0
'			if  not rsget.EOF  then
'				rsget.absolutepage = FCurrPage
'				do until rsget.eof
'					set FItemList(i) = new COffShopSellByTerm
'					FItemList(i).FTerm  = rsget("regdate")
'					FItemList(i).FCount = rsget("cnt")
'					FItemList(i).FSum   = rsget("shoppoint")
'					FItemList(i).FShopid= rsget("regshopid")
'
'					i=i+1
'					rsget.moveNext
'				loop
'			end if
'		rsget.Close
'	end Sub

	'//admin/offshop/accountreport.asp		'/사용안함
	public Sub GetJumunMethodReport()
		Dim sql, i, ix
		maxt = -1
		maxt2 = -1
   		maxc = -1

		sql = "select"
		sql = sql + " sum(cashsum) as 'cashsum'"
		sql = sql + " ,sum(case when jumunmethod='01' then 1 else 0 end) as 'cashcnt'"
		sql = sql + " , sum(cardsum) as 'cardsum'"
		sql = sql + " ,sum(case when jumunmethod='02' then 1 else 0 end) as 'cardcnt'"
		sql = sql + " , sum(giftcardPaysum) as 'giftcardPaysum'"
		sql = sql + " ,sum(case when giftcardPaysum<>0 then 1 else 0 end) as 'giftcardcnt'"
		sql = sql + " ,sum(TenGiftCardPaySum) as 'TenGiftCardPaySum'"
		sql = sql + " ,sum(case when isnull(TenGiftCardMatchCode,'')<>'' then 1 else 0 end) as 'TenGiftCardPaycount'"
		sql = sql + " , count(m.idx) as sellcnt"
		sql = sql + " ,sum(realsum) as realsum"
		sql = sql + " ,sum(spendmile) as spendmile"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql + " ,convert(varchar(10),m.shopregdate,20) as yyyymmdd, datepart(w,m.shopregdate) as dpart"
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql + " ,(m.IXyyyymmdd) as yyyymmdd, datepart(w,m.IXyyyymmdd) as dpart"
		else
			sql = sql + " ,convert(varchar(10),m.shopregdate,20) as yyyymmdd, datepart(w,m.shopregdate) as dpart"
		end if

		sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
		sql = sql + " where m.cancelyn='N'" + vbcrlf

		if FRectShopID<>"" then
			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sql = sql + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sql = sql + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sql = sql + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sql = sql + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

		else
			if FRectStartDay<>"" then
				sql = sql + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sql = sql + " and m.shopregdate<'" + CStr(FRectEnDay) + "'"
			end if
		end if

		sql = sql + " group by " + vbcrlf

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql + " convert(varchar(10),m.shopregdate,20), datepart(w,m.shopregdate)"
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql + " m.IXyyyymmdd"
		else
			sql = sql + " convert(varchar(10),m.shopregdate,20), datepart(w,m.shopregdate)"
		end if

		sql = sql + " order by  yyyymmdd desc"

		'response.write sql &"<br>"
		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)

		do until rsget.eof

			set FItemList(i) = new COffShopSellByTerm

			FItemList(i).fTenGiftCardPaycount = rsget("TenGiftCardPaycount")
			FItemList(i).fTenGiftCardPaySum = rsget("TenGiftCardPaySum")
		    FItemList(i).Fsitename = rsget("yyyymmdd")
		    FItemList(i).Fsellcnt           = rsget("sellcnt")
		    FItemList(i).Fcashsum           = rsget("cashsum")
		    FItemList(i).Fcardsum           = rsget("cardsum")
		    FItemList(i).FgiftcardPaysum    = rsget("giftcardPaysum")
		    FItemList(i).Fcashcnt           = rsget("cashcnt")
		    FItemList(i).Fcardcnt           = rsget("cardcnt")
		    FItemList(i).Fgiftcardcnt       = rsget("giftcardcnt")
		    FItemList(i).frealsum       = rsget("realsum")
		    FItemList(i).fspendmile       = rsget("spendmile")

		    FItemList(i).Fselltotal  = FItemList(i).frealsum + FItemList(i).fspendmile

			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end Sub

	'//admin/offshop/accountreport_month.asp
	public function GetJumunMethodReportMonth
		dim i , sql , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if FRectshopid <> "" then
			sqlsearch = sqlsearch & " and m.shopid = '"& FRectshopid &"'"
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDate) + "'"
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDate) + "'"
			end if
		end if

		if FRectBanPum = "Y" then
			sqlsearch = sqlsearch & " and m.totalsum < 0"
		elseif FRectBanPum = "N" then
			sqlsearch = sqlsearch & " and m.totalsum > 0"
		end if

		if frectoffgubun <> "" then
			if frectoffgubun = "90" then
				sqlsearch = sqlsearch & " and u.shopdiv in ('1','3')"
			elseif frectoffgubun = "95" then
				sqlsearch = sqlsearch & " and u.shopdiv not in ('11','12')"
			else
				sqlsearch = sqlsearch & " and u.shopdiv = '"&frectoffgubun&"'"
			end if
		end if

		sql = "SELECT TOP " & Cstr(FPageSize * FCurrPage)

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " Convert(varchar(7),m.shopregdate,121) AS regdate"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " Convert(varchar(7),m.IXyyyymmdd,121) AS regdate"
		end if

		sql = sql & " ,sum(spendmile) as spendmile"
		sql = sql & " ,sum(TenGiftCardPaySum) as TenGiftCardPaySum"
		sql = sql & " ,sum(cardsum) as cardsum"
		sql = sql & " ,sum(cashsum) as cashsum"
		sql = sql & " ,isNull(sum(giftcardPaysum),0) as giftcardPaysum"
		sql = sql & " ,isNull(sum(extPaysum),0) as extPaysum"
		sql = sql & " ,(sum(spendmile) + sum(TenGiftCardPaySum) + sum(cardsum) + sum(cashsum) + isNull(sum(giftcardPaysum),0) + isNull(sum(extPaysum),0)) as selltotal"
		sql = sql & " from db_shop.dbo.tbl_shopjumun_master m"
		sql = sql + " join [db_shop].[dbo].tbl_shop_user u"
		sql = sql + " 	on m.shopid = u.userid"
		sql = sql & " left join db_partner.dbo.tbl_partner p"
	    sql = sql & "       on m.shopid=p.id "
		sql = sql & " where m.cancelyn='N' " & sqlsearch
		sql = sql & " group by"

		'//주문일 기준
		if frectdatefg = "jumun" then
			sql = sql & " 	Convert(varchar(7),m.shopregdate,121)"

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			sql = sql & " 	Convert(varchar(7),m.IXyyyymmdd,121)"
		end if

		sql = sql & " ORDER BY regdate DESC"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FTotalCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof

			set FItemList(i) = new COffShopSellByTerm

			FItemList(i).FRegdate			= rsget("regdate")
			FItemList(i).fspendmile			= rsget("spendmile")
			FItemList(i).fTenGiftCardPaySum			= rsget("TenGiftCardPaySum")
			FItemList(i).fcardsum			= rsget("cardsum")
			FItemList(i).fcashsum			= rsget("cashsum")
			FItemList(i).fgiftcardPaysum			= rsget("giftcardPaysum")
			FItemList(i).FextPaysum		= rsget("extPaysum")
			FItemList(i).fselltotal			= rsget("selltotal")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function

	public Sub GetWeeklySellCount()
		Dim sql, ix

		sql = "select convert(varchar(7),m.shopregdate,20) as yyyymm" + vbcrlf
		sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
		sql = sql + " where convert(varchar(7),m.shopregdate,20) = '" + CStr(FRectStartDay) + "'" + vbcrlf
		if FRectShopID<>"" then
			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " group by convert(varchar(7),m.shopregdate,20)" + vbcrlf
		sql = sql + " order by convert(varchar(7),m.shopregdate,20) desc"

		rsget.Open sql,dbget,1

		FTotalCount = rsget.RecordCount
	    redim preserve FCountList(FTotalCount)
		do until rsget.eof
				set FCountList(ix) = new COffShopSellByTerm
			    FCountList(ix).FYYYYMMDDHHNNSS = rsget("yyyymm")
				rsget.MoveNext
				ix = ix + 1
		loop
		rsget.close
	end Sub

	public Sub GetWeeklySellReport()
		Dim sql, i
		maxt = -1
   		maxc = -1

		sql = "select convert(varchar(10),m.shopregdate,20) as yyyymm," + vbcrlf
		sql = sql + " datepart(w,m.shopregdate) as dpart," + vbcrlf
		if FRectPointYN = "Y" then
		sql = sql + " sum(m.totalsum) as sumtotal," + vbcrlf
		else
		sql = sql + " sum(m.realsum) as sumtotal," + vbcrlf
		end if
		sql = sql + " count(m.idx) as sellcnt" + vbcrlf
		sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
		sql = sql + " where convert(varchar(7),m.shopregdate,20) ='" + CStr(FRectStartDay) + "'" + vbcrlf
		if FRectShopID<>"" then
			sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " group by convert(varchar(10),m.shopregdate,20), datepart(w,m.shopregdate)" + vbcrlf
		sql = sql + " order by convert(varchar(10),m.shopregdate,20) desc, datepart(w,m.shopregdate) asc" + vbcrlf

		'response.write sql &"<br>"
		rsget.Open sql,dbget,1
		FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)
		do until rsget.eof
				set FItemList(i) = new COffShopSellByTerm
			    FItemList(i).Fsitename = rsget("yyyymm")
				FItemList(i).Fselltotal = rsget("sumtotal")
				FItemList(i).Fsellcnt = rsget("sellcnt")
				FItemList(i).Fdpart = rsget("dpart")

				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
		loop
		rsget.close

	end Sub

	'/offshop/report/bestseller.asp		'/admin/offshop/itemsellsum_zoom.asp
	'/admin/offshop/bestseller_zoom.asp
	public Sub ShopJumunListBybestseller()
		dim sqlStr , i , sqlsearch

		if frectitemid <> "" then
			sqlsearch = sqlsearch & " and d.itemid=" + frectitemid + "" + vbcrlf
		end if

		if frectitemname <> "" then
			sqlsearch = sqlsearch & " and d.itemname like '%" + frectitemname + "%'" + vbcrlf
		end if

		if frectitemgubun <> "" then
			sqlsearch = sqlsearch & " and d.itemgubun='" + frectitemgubun + "'" + vbcrlf
		end if

		if FRectShopid="streetshop014" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'" + vbcrlf
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'" + vbcrlf
			end if

		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " and m.shopregdate>='" + CStr(FRectStartDay) + "'" + vbcrlf
			end if

			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " and m.shopregdate<'" + CStr(FRectEndDay) + "'" + vbcrlf
			end if
		end if

		if FRectShopID<>"" then
			sqlsearch = sqlsearch & " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if

		if (FRectOffgubun<>"") then
		    if (FRectOffgubun="1") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('1','2')"
		    elseif (FRectOffgubun="3") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('3','4')"
		    elseif (FRectOffgubun="5") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('5','6')"
		    elseif (FRectOffgubun="7") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('7','8')"
		    elseif (FRectOffgubun="9") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('9')"
		    end if
		end if

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch & " and d.makerid='" + FRectDesigner + "'" + vbcrlf
		end if

		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr & " sum(d.itemno) as sm, d.makerid, d.itemgubun, d.itemid, d.itemoption, d.itemname" + vbcrlf
		sqlStr = sqlStr & " , d.itemoptionname , d.sellprice,d.realsellprice ,d.addtaxcharge, d.suplyprice" + vbcrlf

		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m" + vbcrlf
			sqlStr = sqlStr + " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d" + vbcrlf
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
			sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d" + vbcrlf
		end if

		sqlStr = sqlStr & " 	on m.idx = d.masteridx" +vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N'"
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr + " 	on m.shopid = u.userid"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " group by d.itemgubun, d.itemid, d.itemoption, d.sellprice, d.addtaxcharge,d.itemname, d.makerid" + vbcrlf
		sqlStr = sqlStr & " , d.itemoptionname ,d.realsellprice , d.suplyprice" + vbcrlf

		if FRectOrder="bysum" then
			sqlStr = sqlStr & " order by sum(d.itemno*d.sellprice) Desc"
		elseif FRectOrder="bycnt" then
			sqlStr = sqlStr & " order by sm Desc"
		else
			sqlStr = sqlStr & " order by sm Desc"
		end if

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordCount
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		maxt =0
		maxc =0
		do until rsget.eof
				set FItemList(i) = new COffShopSellByTerm

				FItemList(i).FItemNo       = rsget("sm")
				FItemList(i).FItemGubun		= rsget("itemgubun")
				FItemList(i).FItemID       = rsget("itemid")
				FItemList(i).FItemOption       = rsget("itemoption")
				FItemList(i).FItemCost       = rsget("sellprice")
				FItemList(i).FItemName     = db2html(rsget("itemname"))
				FItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))
				FItemList(i).FMakerid		= rsget("makerid")
				FItemList(i).frealsellprice		= rsget("realsellprice")
				FItemList(i).fsuplyprice		= rsget("suplyprice")
				FItemList(i).faddtaxcharge      = rsget("addtaxcharge")
				maxc = maxc + FItemList(i).FItemNo
				maxt = maxt + FItemList(i).FItemNo*FItemList(i).FItemCost

				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end sub

	'//admin/offshop/dayitemsellsum.asp
	public Sub getdayitemsum()
		dim sqlStr , i , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if

		if frectextbarcode <> "" and len(frectextbarcode)=12 then
			sqlsearch = sqlsearch & " and d.itemgubun= "& left(frectextbarcode,2) &""
			sqlsearch = sqlsearch & " and d.itemid= "& mid(frectextbarcode,3,6) &""
			sqlsearch = sqlsearch & " and d.itemoption= "& right(frectextbarcode,4) &""
		end if

		if frectitemid <> "" then
			sqlsearch = sqlsearch & " and d.itemid=" + frectitemid + "" + vbcrlf
		end if

		if frectitemname <> "" then
			sqlsearch = sqlsearch & " and d.itemname like '%" + frectitemname + "%'" + vbcrlf
		end if

		if frectitemgubun <> "" then
			sqlsearch = sqlsearch & " and d.itemgubun='" + frectitemgubun + "'" + vbcrlf
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if
		end if

		if FRectShopID<>"" then
			sqlsearch = sqlsearch & " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if

		if (FRectOffgubun<>"") then
		    if (FRectOffgubun="1") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('1','2')"
		    elseif (FRectOffgubun="3") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('3','4')"
		    elseif (FRectOffgubun="5") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('5','6')"
		    elseif (FRectOffgubun="7") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('7','8')"
		    elseif (FRectOffgubun="9") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('9')"
		    end if
		end if

		if FRectDesigner<>"" then
			sqlsearch = sqlsearch & " and d.makerid='" + FRectDesigner + "'" + vbcrlf
		end if

		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & " 	select"

		if frectdatefg = "maechul" then
			sqlStr = sqlStr & " 	m.IXyyyymmdd"
		else
			sqlStr = sqlStr & " 	convert(varchar(10),m.shopregdate,121) as IXyyyymmdd"
		end if

		sqlStr = sqlStr & " 	,sum(d.itemno) as itemno, d.makerid, d.itemgubun, d.itemid, d.itemoption, d.itemname" + vbcrlf
		sqlStr = sqlStr & " 	, d.itemoptionname , d.sellprice,d.realsellprice , d.suplyprice" + vbcrlf

		if FRectOldData="on" then
			sqlStr = sqlStr + " 	from [db_shoplog].[dbo].tbl_old_shopjumun_master m" + vbcrlf
			sqlStr = sqlStr + " 	join [db_shoplog].[dbo].tbl_old_shopjumun_detail d" + vbcrlf
		else
			sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
			sqlStr = sqlStr + " 	join [db_shop].[dbo].tbl_shopjumun_detail d" + vbcrlf
		end if

		sqlStr = sqlStr & " 		on m.idx = d.masteridx" +vbcrlf
		sqlStr = sqlStr & " 		and m.cancelyn='N' and d.cancelyn='N'"
		sqlStr = sqlStr + " 	join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr + " 		on m.shopid = u.userid"
		sqlStr = sqlStr & " 	left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       	on m.shopid=p.id "
		sqlStr = sqlStr & " 	where 1=1 " & sqlsearch
		sqlStr = sqlStr & " 	group by"

		if frectdatefg = "maechul" then
			sqlStr = sqlStr & " 	m.IXyyyymmdd"
		else
			sqlStr = sqlStr & " 	convert(varchar(10),m.shopregdate,121)"
		end if

		sqlStr = sqlStr & " 	,d.itemgubun, d.itemid, d.itemoption, d.sellprice,d.itemname, d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	, d.itemoptionname ,d.realsellprice , d.suplyprice" + vbcrlf
		sqlStr = sqlStr & " ) as t"

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit Sub

		sqlStr = "select top " + CStr(FPageSize*FCurrPage)

		if frectdatefg = "maechul" then
			sqlStr = sqlStr & " m.IXyyyymmdd"
		else
			sqlStr = sqlStr & " convert(varchar(10),m.shopregdate,121) as IXyyyymmdd"
		end if

		sqlStr = sqlStr & " ,sum(d.itemno) as itemno, d.makerid, d.itemgubun, d.itemid, d.itemoption, d.itemname" + vbcrlf
		sqlStr = sqlStr & " , d.itemoptionname , d.sellprice,d.realsellprice , d.suplyprice" + vbcrlf

		if FRectOldData="on" then
			sqlStr = sqlStr + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m" + vbcrlf
			sqlStr = sqlStr + " join [db_shoplog].[dbo].tbl_old_shopjumun_detail d" + vbcrlf
		else
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
			sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d" + vbcrlf
		end if

		sqlStr = sqlStr & " 	on m.idx = d.masteridx" +vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N'"
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr + " 	on m.shopid = u.userid"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on m.shopid=p.id "
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " group by"

		if frectdatefg = "maechul" then
			sqlStr = sqlStr & " 	m.IXyyyymmdd"
		else
			sqlStr = sqlStr & " 	convert(varchar(10),m.shopregdate,121)"
		end if

		sqlStr = sqlStr & " ,d.itemgubun, d.itemid, d.itemoption, d.sellprice,d.itemname, d.makerid" + vbcrlf
		sqlStr = sqlStr & " , d.itemoptionname ,d.realsellprice , d.suplyprice" + vbcrlf
		sqlStr = sqlStr & " order by IXyyyymmdd desc ,d.itemid asc"

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
				set FItemList(i) = new COffShopSellByTerm

				FItemList(i).FItemNo       = rsget("itemno")
				FItemList(i).FItemGubun		= rsget("itemgubun")
				FItemList(i).FItemID       = rsget("itemid")
				FItemList(i).FItemOption       = rsget("itemoption")
				FItemList(i).FItemCost       = rsget("sellprice")
				FItemList(i).FItemName     = db2html(rsget("itemname"))
				FItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))
				FItemList(i).FMakerid		= rsget("makerid")
				FItemList(i).frealsellprice		= rsget("realsellprice")
				FItemList(i).fsuplyprice		= rsget("suplyprice")
				FItemList(i).fIXyyyymmdd      = rsget("IXyyyymmdd")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/offshop/timesellsum.asp
	public sub SearchMallSellrePort5()
		Dim sql, i , sqlsearch
		maxt = -1
   		maxc = -1

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

		if (FRectOffgubun<>"") then
		    if (FRectOffgubun="1") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('1','2')"
		    elseif (FRectOffgubun="3") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('3','4')"
		    elseif (FRectOffgubun="5") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('5','6')"
		    elseif (FRectOffgubun="7") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('7','8')"
		    elseif (FRectOffgubun="9") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('9')"
		    end if
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.IXyyyymmdd) = "&frectweekdate&""
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.regdate) = "&frectweekdate&""
			end if
		end if

		if FRectShopID<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		end if

		sql = "select"
		sql = sql + " datepart(hh,m.shopregdate) as gpart"
		sql = sql + " ,sum(m.realsum) as sumtotal, count(m.idx) as sellcnt, sum(m.spendmile) as spendmile"

		if FRectOldJumun="on" then
			sql = sql + " from [db_shoplog].[dbo].tbl_old_shopjumun_master m" + vbcrlf
		else
			sql = sql + " from [db_shop].[dbo].tbl_shopjumun_master m" + vbcrlf
		end if

		sql = sql + " join [db_shop].[dbo].tbl_shop_user u"
		sql = sql + "		on m.shopid = u.userid"
		sql = sql + " left join db_partner.dbo.tbl_partner p"
	    sql = sql + "		on m.shopid=p.id "
		sql = sql + " where m.cancelyn='N' " & sqlsearch
		sql = sql + " group by datepart(hh,m.shopregdate)" + vbcrlf
		sql = sql + " order by datepart(hh,m.shopregdate) asc"

		'response.write sql &"<br>"
		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)
		do until rsget.eof
			set FItemList(i) = new COffShopSellByTerm

			FItemList(i).fspendmile = rsget("spendmile")
			FItemList(i).Fselltotal = rsget("sumtotal")
			FItemList(i).Fsellcnt = rsget("sellcnt")
			FItemList(i).Fgpart = rsget("gpart")

			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal+FItemList(i).fspendmile)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	'//admin/offshop/timesellsum.asp
	public sub getshopguestcountandhoursellreport()
		Dim sql, i , sqlsearch ,sqlsearch2
		maxt = -1
   		maxc = -1

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(mp.tplcompanyid,'')<>''"
	            sqlsearch2 = sqlsearch2 & " and isNULL(cp.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(mp.tplcompanyid,'')=''"
	        sqlsearch2 = sqlsearch2 & " and isNULL(cp.tplcompanyid,'')=''"
	    end if
		if frectbuyergubun <> "" then
			sqlsearch = sqlsearch + " and isnull(m.buyergubun,-1) = "&frectbuyergubun&""
		end if

		if (FRectOffgubun<>"") then
		    if (FRectOffgubun="1") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('1','2')"
		    elseif (FRectOffgubun="3") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('3','4')"
		    elseif (FRectOffgubun="5") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('5','6')"
		    elseif (FRectOffgubun="7") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('7','8')"
		    elseif (FRectOffgubun="9") then
		        sqlsearch = sqlsearch + " and u.shopdiv in ('9')"
		    end if
		end if

		'//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
				sqlsearch2 = sqlsearch2 + " and c.yyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
				sqlsearch2 = sqlsearch2 + " and c.yyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.shopregdate) = "&frectweekdate&""
				sqlsearch2 = sqlsearch2 + " and datepart(w,c.yyyymmdd) = "&frectweekdate&""
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'"
				sqlsearch2 = sqlsearch2 + " and c.yyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'"
				sqlsearch2 = sqlsearch2 + " and c.yyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.IXyyyymmdd) = "&frectweekdate&""
				sqlsearch2 = sqlsearch2 + " and datepart(w,c.yyyymmdd) = "&frectweekdate&""
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(FRectStartDay) + "'"
				sqlsearch2 = sqlsearch2 + " and c.yyyymmdd>='" + CStr(FRectStartDay) + "'"
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(FRectEndDay) + "'"
				sqlsearch2 = sqlsearch2 + " and c.yyyymmdd<'" + CStr(FRectEndDay) + "'"
			end if

			if frectweekdate <> "" then
				sqlsearch = sqlsearch + " and datepart(w,m.shopregdate) = "&frectweekdate&""
				sqlsearch2 = sqlsearch2 + " and datepart(w,c.yyyymmdd) = "&frectweekdate&""
			end if
		end if

		if FRectShopID<>"" then
			sqlsearch = sqlsearch + " and m.shopid='" + FRectShopID + "'"
			sqlsearch2 = sqlsearch2 + " and c.shopid='" + FRectShopID + "'"
		end if

		sql = "select"
		sql = sql + " md.gpart ,isnull(md.z1_in,0) as z1_in, isnull(md.z1_out,0) as z1_out ,isnull(md.z2_in,0) as z2_in, isnull(md.z2_out,0) as z2_out"
		sql = sql & " ,round(convert(float,isnull(md.z1_in,0)+isnull(md.z1_out,0))/2,0) as z1_all"
		sql = sql & " ,round(convert(float,isnull(md.z2_in,0)+isnull(md.z2_out,0))/2,0) as z2_all"
		sql = sql + " ,g.dpart, isnull(g.sumtotal,0) as sumtotal, isnull(g.sellcnt,0) as sellcnt, isnull(g.spendmile,0) as spendmile"
		sql = sql + " from ("
		sql = sql + " 		select"
		sql = sql + " 		datepart(hh,c.yyyymmdd) as gpart"
		sql = sql + " 		,sum(isnull(c.z1_in,0)) as z1_in, sum(isnull(c.z1_out,0)) as z1_out ,sum(isnull(c.z2_in,0)) as z2_in, sum(isnull(c.z2_out,0)) as z2_out"
		sql = sql + " 		from db_shop.dbo.tbl_shop_guestcount c"
		sql = sql + " 		left join db_partner.dbo.tbl_partner cp"
	    sql = sql + "       	on c.shopid=cp.id "
		sql = sql + " 		where 1=1 " & sqlsearch2
		sql = sql + " 		group by datepart(hh,c.yyyymmdd)"
		sql = sql + " ) md"
		sql = sql + " left join ("
		sql = sql + " 		select"
		sql = sql + " 		datepart(hh,m.shopregdate) as dpart"
		sql = sql + " 		,sum(m.realsum) as sumtotal, count(m.idx) as sellcnt, sum(m.spendmile) as spendmile"

		if FRectOldJumun="on" then
			sql = sql + " 		from [db_shoplog].[dbo].tbl_old_shopjumun_master m"
		else
			sql = sql + " 		from [db_shop].[dbo].tbl_shopjumun_master m"
		end if

		sql = sql + " 		join [db_shop].[dbo].tbl_shop_user u"
		sql = sql + " 			on m.shopid = u.userid"
		sql = sql + " 		left join db_partner.dbo.tbl_partner mp"
	    sql = sql + "       	on m.shopid=mp.id "
		sql = sql + " 		where m.cancelyn='N' " & sqlsearch
		sql = sql + " 		group by datepart(hh,m.shopregdate)"
		sql = sql + " ) g"
		sql = sql + " 	on md.gpart = g.dpart"
		sql = sql + " order by md.gpart asc"

		'response.write sql &"<br>"
		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)
		do until rsget.eof
			set FItemList(i) = new COffShopSellByTerm

			FItemList(i).fspendmile = rsget("spendmile")
			FItemList(i).fgpart = rsget("gpart")
			FItemList(i).fz1_in = rsget("z1_in")
			FItemList(i).fz1_out = rsget("z1_out")
			FItemList(i).fz1_all = rsget("z1_all")
			FItemList(i).fz2_in = rsget("z2_in")
			FItemList(i).fz2_out = rsget("z2_out")
			FItemList(i).fz2_all = rsget("z2_all")
			FItemList(i).Fselltotal = rsget("sumtotal")
			FItemList(i).Fsellcnt = rsget("sellcnt")
			FItemList(i).Fdpart = rsget("dpart")

			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal+FItemList(i).fspendmile)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
	end sub

	public sub SearchCategorySellrePort()
	Dim sql, i
    maxt = -1
    maxc = -1

		sql = "select count(d.itemno) as sellcnt, sum(d.realsellprice*d.itemno) as sumtotal," + vbcrlf
		sql = sql + " s.catecdl,l.code_nm" + vbcrlf
		if FRectOldData="on" then
			sql = sql + " from  [db_shoplog].[dbo].tbl_old_shopjumun_master m," + vbcrlf
			sql = sql + " [db_shoplog].[dbo].tbl_old_shopjumun_detail d" + vbcrlf
		else
			sql = sql + " from  [db_shop].[dbo].tbl_shopjumun_master m," + vbcrlf
			sql = sql + " [db_shop].[dbo].tbl_shopjumun_detail d" + vbcrlf
		end if
		sql = sql + " ,[db_shop].[dbo].tbl_shop_item s" + vbcrlf
		sql = sql + " left join [db_item].[dbo].tbl_item_large l " + vbcrlf
		sql = sql + " 	on s.catecdl=l.code_large" + vbcrlf
		sql = sql + " where m.orderno = d.orderno" + vbcrlf
		sql = sql + " and d.itemgubun=s.itemgubun" + vbcrlf
		sql = sql + " and d.itemid=s.shopitemid" + vbcrlf
		sql = sql + " and d.itemoption=s.itemoption" + vbcrlf
		sql = sql + " and m.shopid='" + FRectShopID + "'" + vbcrlf
		sql = sql + " and m.shopregdate>='" + CStr(FRectStartDay) + "'" + vbcrlf
		sql = sql + " and m.shopregdate<'" + CStr(FRectEndDay) + "'" + vbcrlf
		sql = sql + " and m.cancelyn='N'" + vbcrlf
		sql = sql + " and d.cancelyn='N'" + vbcrlf
		sql = sql + " group by s.catecdl,l.code_nm" + vbcrlf
		sql = sql + " order by s.catecdl"

		rsget.Open sql,dbget,1

		FResultCount = rsget.RecordCount


	    redim preserve FItemList(FResultCount)


		do until rsget.eof

			set FItemList(i) = new COffShopSellByTerm
		    FItemList(i).Fsitename = rsget("code_nm")
			FItemList(i).Fselltotal = rsget("sumtotal")
			FItemList(i).Fsellcnt = rsget("sellcnt")

			if IsNULL(FItemList(i).Fsitename) then FItemList(i).Fsitename = "미지정"

			if Not IsNull(FItemList(i).Fselltotal) then
				maxt = MaxVal(maxt,FItemList(i).Fselltotal)
				maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
			end if

			rsget.MoveNext
			i = i + 1
		loop
		rsget.close
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

	Private Sub Class_Initialize()
		redim  FItemList(0)
		redim  FCountList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class
%>
