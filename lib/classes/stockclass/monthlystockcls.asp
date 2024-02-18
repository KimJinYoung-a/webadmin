<%
'###########################################################
' Description : 월별 재고자산 클래스
' History : 이상구 생성
'###########################################################

'' 한글 한글

Const iTsGroupID = "G02799"  ''G02799 (주)텐바이텐[아이띵소]_상품 , G02843 (주)텐바이텐[아이띵소]_출고처

Class CMonthlyStockEtcChulGoItem
    public FSocID
    public FSocName
    public FIpChulMwGubun
    public FItemgubun

    public FtargetGbn
    public Flastmwdiv

    public FTTLCNT
    public FTTLSellSum
    public FTTLBuySum
    public FTTLSuplySum

    public FMayStockPrice


    public FIpChulCode

    public FItemID
    public FItemOption
    public FItemName
    public FItemOptionName
    public FMakerid
    public FMaeipLedgeravgipgoPrice

    public function getLogisticsCode()
        if Len(CStr(FItemId))>6 then
            getLogisticsCode = Fitemgubun+Format00(8,FItemId) + FItemOption
        else
            getLogisticsCode = Fitemgubun+Format00(6,FItemId) + FItemOption
        end if
    end function

    public function getBusiName
        if (FtargetGbn="ON") then
		    getBusiName      = "온라인"
		elseif (FtargetGbn="OF") then
		    getBusiName      = "오프라인"
		elseif (FtargetGbn="IT") then
		    getBusiName      = "아이띵소(구)"
		elseif (FtargetGbn="ET") then
		    getBusiName      = "띵소"
	    elseif (FtargetGbn="EG") then
		    getBusiName      = "EG"
		elseif (FtargetGbn="3P") then
		    getBusiName      = "3PL"
		elseif Not IsNull(FtargetGbn) then
		    getBusiName      = FtargetGbn
		else
		    getBusiName      = "-"
	    end if
	end function

	public function getMaeipGubunName()
		if FMaeIpGubun="M" then
			getMaeipGubunName = "매입"
		elseif FMaeIpGubun="W" then
			getMaeipGubunName = "위탁"
		elseif FMaeIpGubun="U" then
			getMaeipGubunName = "업체"
		elseif FMaeIpGubun="Z" then
			getMaeipGubunName = "-"
		elseif FMaeIpGubun="B011" then
			getMaeipGubunName = "위탁판매"
		elseif FMaeIpGubun="B012" then
			getMaeipGubunName = "업체위탁"
		elseif FMaeIpGubun="B013" then
			getMaeipGubunName = "출고위탁"
		elseif FMaeIpGubun="B021" then
			getMaeipGubunName = "오프매입"
		elseif FMaeIpGubun="B022" then
			getMaeipGubunName = "매장매입"
		elseif FMaeIpGubun="B023" then
			getMaeipGubunName = "가맹점매입"
		elseif FMaeIpGubun="B031" then
			getMaeipGubunName = "출고매입"
		elseif FMaeIpGubun="B032" then
			getMaeipGubunName = "센터매입"
		else
			getMaeipGubunName = FMaeIpGubun
		end if

		IF isNULL(FMaeIpGubun) then
		    getMaeipGubunName ="-"
		end if
	end function

	public function getITemGubunName()
	if Fitemgubun="10" then
			getITemGubunName = "일반"
		elseif Fitemgubun="90" then
			getITemGubunName = "오프전용"
		elseif Fitemgubun="70" then
			getITemGubunName = "소모품"
		elseif Fitemgubun="80" then
			getITemGubunName = "사은품"
		elseif Fitemgubun="85" then
			getITemGubunName = "사은품"
		else
			getITemGubunName = Fitemgubun
		end if
	end function
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CMonthlyStockIpgoItem
    public Fyyyymm
	public FstockPlace
	public Fshopid
	public FtargetGbn
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public FipgoMWdiv
	public FitemMWdiv
	public FitemVatInclude
	public FtotItemNo
	public FtotBuyCash
	public Flastupdate
	public FstockIpgoNo
	public FipgoType
	public Flastmwdiv
	public FlastCenterMWDiv
	public Flastmakerid
	public Flastvatinclude

	public function GetIpgoMWdivName()
		Select Case FipgoMWdiv
			Case "M"
				GetIpgoMWdivName = "<font color='red'>매입</font>"
			Case "W"
				GetIpgoMWdivName = "위탁"
			Case "B031"
				GetIpgoMWdivName = "출고매입"
			Case "B012"
				GetIpgoMWdivName = "업체위탁"
			Case "B013"
				GetIpgoMWdivName = "출고위탁"
			Case Else
				GetIpgoMWdivName = FipgoMWdiv
		End Select
	end function

	public function GetStockPlaceName()
		Select Case FstockPlace
			Case "L"
				GetStockPlaceName = "물류"
			Case "S"
				GetStockPlaceName = "매장"
			Case "O"
				GetStockPlaceName = "온라인정산"
			Case "F"
				GetStockPlaceName = "오프정산"
			Case "A"
				GetStockPlaceName = "핑거스정산"
			Case Else
				GetStockPlaceName = FstockPlace
		End Select
	end function

	public function GetIpgoTypeName()
		Select Case FipgoType
			Case "normal"
				GetIpgoTypeName = "물류 입고"
			Case "shopchulgo"
				GetIpgoTypeName = "물류-&gt;매장"
			Case "shopipgo"
				GetIpgoTypeName = "매장 직입고"
			Case "shop_chulgo_ipgo"
				GetIpgoTypeName = "물류-&gt;매장 + 매장 직입고"
			Case "witakjungsan"
				GetIpgoTypeName = "위탁상품 매입정산"
			Case Else
				GetIpgoTypeName = FipgoType
		End Select
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CMonthlyStockAvgPriceItem
    public Fyyyymm
	public FstockPlace
	public Fshopid
	public Fitemgubun
	public Fitemid
	public Fitemoption

	public FtotsysstockPrev
	public FavgipgoPriceSumPrev
	public FtotsysstockShopPrev
	public FtotsysstockBuySumShopPrev
	public FtotItemNo
	public FtotBuyCash
	public Ftotsysstock
	public FavgipgoPricePrev
	public FavgipgoPrice
	public FlastmwdivPrev
	public Flastmwdiv
	public FmakeridPrev
	public Fmakerid

	public function GetIpgoMWdivName()
		Select Case FipgoMWdiv
			Case "M"
				GetIpgoMWdivName = "<font color='red'>매입</font>"
			Case "W"
				GetIpgoMWdivName = "위탁"
			Case Else
				GetIpgoMWdivName = FipgoMWdiv
		End Select
	end function

	public function GetStockPlaceName()
		Select Case FstockPlace
			Case "L"
				GetStockPlaceName = "물류"
			Case "S"
				GetStockPlaceName = "매장"
			Case Else
				GetStockPlaceName = FstockPlace
		End Select
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CMonthlyStockSum
    public FBusiName
    public FtargetGbn
	public FTotCount
	public FTotBuySum
	public fpurchasetypename
	Public FTotRealStockCount
	Public FTotRealStockBuySum

	public FTotPreCount
	public FTotPreBuySum
	public FTotPreRealStockCount
	public FTotPreRealStockBuySum
	public FTotIpCount
	public FTotIpBuySum

	public FTotShopBuySum
	public FTotSellSum
	public FavgIpgoPriceSum

	public FMaeIpGubun
	public FItemMaeIpGubun

	public Fitemgubun
	public FItemId
	public Fregdate
	public FItemOption
	public FItemName
	public FItemOptionname

	public FIsUsing
	public FOptionUsing
	public FMakerid
	public FMakerUsing
    public FSocName
	public FlastIpgoDate
    public Fshopid
    public FShopName
    public FCurrshopitemprice

    public FTotRealCount
	public FTotRealBuySum
	public FTotRealSellSum

    public FTotLossCount
    public FTotLossBuySum

    public FTotSellCount
    public FTotSellBuySum
    public FTotOffChulCount
    public FTotOffChulBuySum
    public FTotEtcChulCount
    public FTotEtcChulBuySum
    public FTotCsChulCount
    public FTotCsChulBuySum

	public FTotErrRealCheckCount
    public FTotErrRealCheckBuySum

	Public FTotErrBadItemCount
	Public FTotErrBadItemBuySum

	Public FTotMoveItemCount
	Public FTotMoveItemBuySum

    public Fetcjungsantype
	Public FpurchaseType

	public FErrItemCnt

	' 사용중지. 디비에서 일괄로 쿼리해서 가져 오세요.
	public function getPurchaseTypeName
		Select Case FpurchaseType
			Case "1"
				getPurchaseTypeName = "일반유통"
			Case "4"
				getPurchaseTypeName = "사입"
			Case "5"
				getPurchaseTypeName = "OFF사입"
			Case "6"
				getPurchaseTypeName = "수입"
			Case "7"
				getPurchaseTypeName = "브랜드수입"
			Case "8"
				getPurchaseTypeName = "제작"
			Case "9"
				GetPurchaseTypeName = "해외직구"
			Case "10"
				GetPurchaseTypeName = "B2B"
			Case Else
				getPurchaseTypeName = "Err"
		End Select
	End Function

    public function getEtcJungsanTypeName
        if isNULL(Fetcjungsantype) then
            getEtcJungsanTypeName = ""
            Exit function
        end if

        if (Fetcjungsantype="1") then
            getEtcJungsanTypeName = "판매분정산"
        elseif (Fetcjungsantype="2") then
            getEtcJungsanTypeName = "출고분정산"
        elseif (Fetcjungsantype="3") then
            getEtcJungsanTypeName = "가맹점정산"
        elseif (Fetcjungsantype="4") then
            getEtcJungsanTypeName = "직영점정산"
        else
            getEtcJungsanTypeName = Fetcjungsantype
        end if
    end function

    public function getCalcuCurSysStock
        getCalcuCurSysStock = FTotPreCount + FTotIpCount + FTotMoveItemCount + FTotSellCount + FTotOffChulCount + FTotEtcChulCount + FTotCsChulCount + FTotLossCount
    end function

    public function getCalcuCurSysBuySum
        getCalcuCurSysBuySum = FTotPreBuySum + FTotIpBuySum + FTotMoveItemBuySum + FTotSellBuySum + FTotOffChulBuySum + FTotEtcChulBuySum + FTotCsChulBuySum + FTotLossBuySum
    end function

    public function getCalcuCurRealStock
        getCalcuCurRealStock = FTotPreRealStockCount + FTotIpCount + FTotMoveItemCount + FTotSellCount + FTotOffChulCount + FTotEtcChulCount + FTotCsChulCount + FTotLossCount
    end function

    public function getCalcuCurRealBuySum
        getCalcuCurRealBuySum = FTotPreRealStockBuySum + FTotIpBuySum + FTotMoveItemBuySum + FTotSellBuySum + FTotOffChulBuySum + FTotEtcChulBuySum + FTotCsChulBuySum + FTotLossBuySum
    end function

    public function getLogisticsCode()
        if Len(CStr(FItemId))>6 then
            getLogisticsCode = Fitemgubun+Format00(8,FItemId) + FItemOption
        else
            getLogisticsCode = Fitemgubun+Format00(6,FItemId) + FItemOption
        end if
    end function

    public function getLossAssignedWongaCnt()
        getLossAssignedWongaCnt = getWongaCnt+FTotLossCount
    end function

    public function getLossAssignedWongaSum()
        getLossAssignedWongaSum = getWongaSum+FTotLossBuySum
    end function

    public function getWongaCnt()
        getWongaCnt = FTotPreCount+FTotIpCount-FTotCount
    end function

    public function getWongaSum()
        getWongaSum = FTotPreBuySum+FTotIpBuySum-FTotBuySum
    end function

    public function getBusiName
        if (FtargetGbn="ON") then
		    getBusiName      = "온라인"
		elseif (FtargetGbn="OF") then
		    getBusiName      = "오프라인"
		elseif (FtargetGbn="IT") then
		    getBusiName      = "아이띵소(구)"
		elseif (FtargetGbn="ET") then
		    getBusiName      = "3PL(아이띵소)"
		elseif (FtargetGbn="EG") then
		    getBusiName      = "3PL(유그레잇)"
        elseif (FtargetGbn="3P") then
		    getBusiName      = "3PL"
		else
		    getBusiName      = "-"
	    end if
	end function

	public function getMaeipGubunName()
		if FMaeIpGubun="M" then
			getMaeipGubunName = "매입"
		elseif FMaeIpGubun="W" then
			getMaeipGubunName = "위탁"
		elseif FMaeIpGubun="U" then
			getMaeipGubunName = "업체"
		elseif FMaeIpGubun="Z" then
			getMaeipGubunName = "-"
		elseif FMaeIpGubun="B011" then
			getMaeipGubunName = "위탁판매"
		elseif FMaeIpGubun="B012" then
			getMaeipGubunName = "업체위탁"
		elseif FMaeIpGubun="B013" then
			getMaeipGubunName = "출고위탁"
		elseif FMaeIpGubun="B021" then
			getMaeipGubunName = "오프매입"
		elseif FMaeIpGubun="B022" then
			getMaeipGubunName = "매장매입"
		elseif FMaeIpGubun="B023" then
			getMaeipGubunName = "가맹점매입"
		elseif FMaeIpGubun="B031" then
			getMaeipGubunName = "출고매입"
		elseif FMaeIpGubun="B032" then
			getMaeipGubunName = "센터매입"
		else
			getMaeipGubunName = FMaeIpGubun
		end if

		IF isNULL(FMaeIpGubun) then
		    getMaeipGubunName ="-"
		end if
	end function

	public function getITemGubunName()
	if Fitemgubun="10" then
			getITemGubunName = "일반"
		elseif Fitemgubun="90" then
			getITemGubunName = "오프전용"
		elseif Fitemgubun="70" then
			getITemGubunName = "소모품"
		elseif Fitemgubun="80" then
			getITemGubunName = "사은품"
		elseif Fitemgubun="85" then
			getITemGubunName = "사은품"
		else
			getITemGubunName = Fitemgubun
		end if
	end function

	public function getITemMaeipGubunName()
		if FItemMaeIpGubun = "M" then
			getITemMaeipGubunName = "<font color=red>매입</font>"
		elseif FItemMaeIpGubun = "W" then
			getITemMaeipGubunName = "위탁"
		elseif IsNull(FItemMaeIpGubun) then
			getITemMaeipGubunName = "<font color=red>ERR</font>"
		elseif Trim(FItemMaeIpGubun) = "" then
			getITemMaeipGubunName = "<font color=red>ERR</font>"
		else
			getITemMaeipGubunName = FItemMaeIpGubun
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CStockOverValueSum
    public FtargetGbn
	public Fitemgubun
	public FMaeIpGubun

	public FpurchasetypeStr

	public FTotBuySum1
	public FTotBuySum2
	public FTotBuySum3
	public FTotBuySum4
	public FTotBuySum5
	public FTotBuySum6
	public FTotBuySum7
	public FTotBuySum8

	public FTotBuySum11
	public FTotBuySum12
	public FTotBuySum13
	public FTotBuySum14

	public FTotBuySum

	Public Fshopid
	public Fmakerid
	public FlastIpgoDate
	public FtotStockNo

	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public FbuyPrice

	public FLstComm_cd


	public function getOverValueStockPrice
		' * 최종입고월을 기준으로 재고월령을 산정합니다.
		' * 재고월령이 1년을 넘는 상품의 경우 재고평가충당금(재고평가손실)을 적용합니다.
		' * 재고월령이 1년-2년 사이인 경우 매입가 대비 50% 의 평가충당금을 산정합니다.
		' * 재고월령이 2년을 넘는 경우 매입가 대비 100% 의 평가충당금을 산정합니다.
		getOverValueStockPrice = (FTotBuySum4 / 2) + FTotBuySum5 + FTotBuySum6
	end function

	public function getOverValueStockPriceYear
		' * 최종입고월을 기준으로 재고월령을 산정합니다.
		' * 재고월령이 1년을 넘는 상품의 경우 재고평가충당금(재고평가손실)을 적용합니다.
		' * 재고월령이 1년-2년 사이인 경우 매입가 대비 50% 의 평가충당금을 산정합니다.
		' * 재고월령이 2년을 넘는 경우 매입가 대비 100% 의 평가충당금을 산정합니다.
		getOverValueStockPriceYear = (FTotBuySum12 / 2) + FTotBuySum13 + FTotBuySum14 + FTotBuySum6
	end function

	public function getLastCommCD
		''B011 : 위탁판매, B012 : 업체위탁, B013 : 출고위탁, B021 : 오프매입, B022 : 매장매입, B023 : 가맹점매입, B031 : 출고매입, B032 : 센터매입

		select case FLstComm_cd
			case "B011"
				getLastCommCD = "위탁판매"
			case "B012"
				getLastCommCD = "업체위탁"
			case "B013"
				getLastCommCD = "출고위탁"
			case "B021"
				getLastCommCD = "오프매입"
			case "B022"
				getLastCommCD = "매장매입"
			case "B023"
				getLastCommCD = "가맹점매입"
			case "B031"
				getLastCommCD = "출고매입"
			case "B032"
				getLastCommCD = "센터매입"
			case else
				getLastCommCD = FLstComm_cd
		end select

	end function

	public function GetlastIpgoDate
		if (FlastIpgoDate = "1") then
			GetlastIpgoDate = "1개월~3개월"
		elseif (FlastIpgoDate = "2") then
			GetlastIpgoDate = "4개월~6개월"
		elseif (FlastIpgoDate = "3") then
			GetlastIpgoDate = "7개월~12개월"
		elseif (FlastIpgoDate = "4") then
			GetlastIpgoDate = "1년~2년"
		elseif (FlastIpgoDate = "5") then
			GetlastIpgoDate = "2년초과"
		elseif (FlastIpgoDate = "6") then
			GetlastIpgoDate = "NULL"
		elseif (FlastIpgoDate = "7") then
			GetlastIpgoDate = "13개월~18개월"
		elseif (FlastIpgoDate = "8") then
			GetlastIpgoDate = "19개월~24개월"
		else
			GetlastIpgoDate = FlastIpgoDate
		end if
	end function

    public function getBusiName
        if (FtargetGbn="ON") then
		    getBusiName      = "온라인"
		elseif (FtargetGbn="OF") then
		    getBusiName      = "오프라인"
		elseif (FtargetGbn="IT") then
		    getBusiName      = "아이띵소(구)"
		elseif (FtargetGbn="ET") then
		    getBusiName      = "3PL(아이띵소)"
		elseif (FtargetGbn="EG") then
		    getBusiName      = "3PL(유그레잇)"
		elseif (FtargetGbn="3P") then
		    getBusiName      = "3PL"
		else
		    getBusiName      = "-"
	    end if
	end function

	public function getMaeipGubunName()
		if FMaeIpGubun="M" then
			getMaeipGubunName = "매입"
		elseif FMaeIpGubun="W" then
			getMaeipGubunName = "위탁"
		elseif FMaeIpGubun="U" then
			getMaeipGubunName = "업체"
		elseif FMaeIpGubun="Z" then
			getMaeipGubunName = "-"
		elseif FMaeIpGubun="B011" then
			getMaeipGubunName = "위탁판매"
		elseif FMaeIpGubun="B012" then
			getMaeipGubunName = "업체위탁"
		elseif FMaeIpGubun="B013" then
			getMaeipGubunName = "출고위탁"
		elseif FMaeIpGubun="B021" then
			getMaeipGubunName = "오프매입"
		elseif FMaeIpGubun="B022" then
			getMaeipGubunName = "매장매입"
		elseif FMaeIpGubun="B023" then
			getMaeipGubunName = "가맹점매입"
		elseif FMaeIpGubun="B031" then
			getMaeipGubunName = "출고매입"
		elseif FMaeIpGubun="B032" then
			getMaeipGubunName = "센터매입"
		else
			getMaeipGubunName = FMaeIpGubun
		end if

		IF isNULL(FMaeIpGubun) then
		    getMaeipGubunName ="-"
		end if
	end function

	public function getITemGubunName()
		if Fitemgubun="10" then
			getITemGubunName = "일반"
		elseif Fitemgubun="90" then
			getITemGubunName = "오프전용"
		elseif Fitemgubun="70" then
			getITemGubunName = "소모품"
		elseif Fitemgubun="80" then
			getITemGubunName = "사은품"
		elseif Fitemgubun="85" then
			getITemGubunName = "사은품"
		else
			getITemGubunName = Fitemgubun
		end if
	end function

	public function getITemMaeipGubunName()
		if FItemMaeIpGubun = "M" then
			getITemMaeipGubunName = "<font color=red>매입</font>"
		elseif FItemMaeIpGubun = "W" then
			getITemMaeipGubunName = "위탁"
		elseif IsNull(FItemMaeIpGubun) then
			getITemMaeipGubunName = "<font color=red>ERR</font>"
		elseif Trim(FItemMaeIpGubun) = "" then
			getITemMaeipGubunName = "<font color=red>ERR</font>"
		else
			getITemMaeipGubunName = FItemMaeIpGubun
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CMonthlyStockItem
	public FYYYYMM
	public FItemGubun
	public FItemId
	public FItemoption

	public Fshopid
	public Fmwdiv
	Public Fcentermwdiv
	public Fmakerid
	public FavgipgoPrice
	public FbuyPrice
	public Fregdate
	public Flastupdate
	public Flastvatinclude
	public FlastIpgoDate
	public FlastIpgoDateLogics
	public FfirstIpgoDate

	public function getMaeipGubunName()
		if Fmwdiv="M" then
			getMaeipGubunName = "매입"
		elseif Fmwdiv="W" then
			getMaeipGubunName = "위탁"
		elseif Fmwdiv="U" then
			getMaeipGubunName = "업체"
		elseif Fmwdiv="Z" then
			getMaeipGubunName = "-"
		elseif Fmwdiv="B011" then
			getMaeipGubunName = "위탁판매"
		elseif Fmwdiv="B012" then
			getMaeipGubunName = "업체위탁"
		elseif Fmwdiv="B013" then
			getMaeipGubunName = "출고위탁"
		elseif Fmwdiv="B021" then
			getMaeipGubunName = "오프매입"
		elseif Fmwdiv="B022" then
			getMaeipGubunName = "매장매입"
		elseif Fmwdiv="B023" then
			getMaeipGubunName = "가맹점매입"
		elseif Fmwdiv="B031" then
			getMaeipGubunName = "출고매입"
		elseif Fmwdiv="B032" then
			getMaeipGubunName = "센터매입"
		else
			getMaeipGubunName = Fmwdiv
		end if

		IF isNULL(Fmwdiv) then
		    getMaeipGubunName ="-"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CMonthlyErrorStockItem
	public FMAX_YYYYMM
	public FMIN_YYYYMM

	public FMaeipCount
	public FWitakCount
	public FErrorCount

    public FminIpgodate
    public FmaxIpgodate
    public FnullCNT


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CMonthlyStock

	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

	public FRectGubun
	public FRectYYYYMM
	public FRectYYYYMMDD
	public FRectIsUsing
	public FRectMwDiv
	public FRectMakerid
	public FRectNewItem
	public FRectVatYn

	public FRectPlaceGubun

    public FRectItemGubun
	public FRectItemId
	public FRectItemOption
    public FRectShopid
	Public FRectShowShopid
    public FRectShowMinus
	public FRectShowMinusOnly
    public FRectOFFReturn2OnStock
    public FRectMinusInclude         '' 마이너스 포함여부
    public FRectITSOnlyOrNot         '' iTS포함여부 Y,N,""
	public FRectPurchaseType
    public FRectShopSuplyPrice       '' 공급액으로 표시 10/11(과세)
    public FRectGroupbyType          '' 브랜드로 그루핑 1
    public FRectOrdTp                '' 정렬순서

	public FRectMonthGubun			''월령

	public FRectPriceGubun
	public FRectIpgoMWdiv
	public FRectItemMwDiv
	public FRectLastCenterMWDiv

    public FRectIsFix ''월말 저장값.
    public FRectTargetGbn
    public FRectShowItemList
	Public FRectDispCate

    public FRectLastIpgoGBN
    public FRectetcjungsantype
	Public FRectBrandUseYN

    public FRectSocID
    public FRectIpChulCode
    public FRectListType
	public FRectChulgoGubun
	Public FRectGrpType

	public FRectStartYYYYMM
	public FRectEndYYYYMM

	Public FRectIpgoType
	public FRectShowUpbae

	public FRectStartDate
	public FRectEndDate

    public function getComonSubQueryLogisticsFIX(iYYYYMM,onOffGbn,igrouping)
        Dim retStr,stockColNm, valPrice, valPricePre
        Dim isPreMonth : isPreMonth = (FRectYYYYMM<>iYYYYMM)

        ''isPreMonth = FALSE ''무조건 상품별 조인, 현재달 기준으로 구함.

        if FRectGubun="sys" then
		    stockColNm = "s.totsysstock"
		else
		    stockColNm = "(s.realstock-s.errbaditemno)" ''불량재고도 실재고임.
	    end if

	    if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(isNULL(isNULL(s.lastvatinclude,i.vatinclude),si.vatinclude), 'Y') = 'Y' then (s.lastbuyPrice*10/11) else s.lastbuyPrice end)"
			valPricePre = "(case when IsNull(isNULL(isNULL(s.lastvatinclude,i.vatinclude),si.vatinclude), 'Y') = 'Y' then (s.lastbuyPrice*10/11) else Ns.lastbuyPrice end)"
		else
			valPrice = "s.lastbuyPrice"
			valPricePre = "Ns.lastbuyPrice"
		end if


	    retStr = ""

	    if (igrouping="itemgubun") then
            retStr = retStr & "	select s.itemgubun as itemgubun"
            retStr = retStr & "	, IsNULL(s.lastmwdiv,'Z') as mwdiv "
            retStr = retStr & "	, s.targetGbn "
        elseif (igrouping="makerid") then
            retStr = retStr & "	select s.lastmakerid as makerid "
        elseif (igrouping="item") then
            retStr = retStr & "	select s.itemgubun,s.itemid,s.itemoption"
            retStr = retStr & "	, isNULL(i.regdate,si.regdate) as regdate"
            retStr = retStr & "	, isNULL(i.itemname,si.shopitemname) as itemname, IsNULL(s.lastmwdiv,'Z') as mwdiv, IsNULL(o.optionname,'') as itemoptionname"
            retStr = retStr & "	, isNULL(i.isusing,si.isusing) as isusing, IsNULL(o.isusing,'Y') as optionusing"
            retStr = retStr & "	, ("&stockColNm&") as totno "
            retStr = retStr & "	, ("&stockColNm&"*IsNULL("&valPrice&",0)) as buysum "
            retStr = retStr & "	, (s.totipgono) as totipgono "
            if (isPreMonth) then
                retStr = retStr & "	, (s.totipgono*IsNULL("&valPricePre&","&valPrice&")) as ipgobuysum "  ''2013/03/18 매입가 변경되는 CASE
            else
                retStr = retStr & "	, (s.totipgono*IsNULL("&valPrice&",0)) as ipgobuysum "
            end if
            retStr = retStr & "	, (s.lossno) as totlossno "
            retStr = retStr & "	, (s.lossno*IsNULL("&valPrice&",0)) as lossbuysum "
            retStr = retStr & "	, s.lastmakerid, s.targetGbn"
        end if

        if (igrouping<>"item") then
            retStr = retStr & "	, sum("&stockColNm&") as totno "
            retStr = retStr & "	, Sum("&stockColNm&"*IsNULL("&valPrice&",0)) as buysum "
            retStr = retStr & "	, sum(s.totipgono) as totipgono "
            if (isPreMonth) then
                retStr = retStr & "	, Sum(s.totipgono*IsNULL("&valPricePre&","&valPrice&")) as ipgobuysum "  ''2013/03/18 매입가 변경되는 CASE
            else
                retStr = retStr & "	, Sum(s.totipgono*IsNULL("&valPrice&",0)) as ipgobuysum "
            end if
            retStr = retStr & "	, sum(s.lossno) as totlossno "
            retStr = retStr & "	, Sum(s.lossno*IsNULL("&valPrice&",0)) as lossbuysum "
        end if
        retStr = retStr & "	from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s "

		if FRectPurchaseType <> "" then
			retStr = retStr & "	left join db_partner.dbo.tbl_partner p "
			retStr = retStr & "	on "
			retStr = retStr & "		s.lastmakerid = p.id "
		end if

        if (isPreMonth) then  ''이전달 내역은 현재고 기준 Join
            retStr = retStr & "	Join [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary Ns"  ''이달재고.
            retStr = retStr & "	on Ns.yyyymm='"&FRectYYYYMM&"'"
            retStr = retStr & "	and s.itemgubun=Ns.itemgubun"
            retStr = retStr & "	and s.itemid=Ns.itemid"
            retStr = retStr & "	and s.itemoption=Ns.itemoption"

            if (FRectMakerid<>"") then
                retStr = retStr + " and Ns.lastmakerid='" + FRectMakerid + "'"
            end if
            if (FRectItemGubun<>"") then
                retStr = retStr & "	and Ns.itemgubun='"&FRectItemGubun&"'"
            end if
            if (FRectTargetGbn<>"") then
				if (FRectTargetGbn = "3X") then
					retStr = retStr + " and Ns.targetGbn not in ('ET', 'EG') "
				else
					retStr = retStr + " and Ns.targetGbn='" + FRectTargetGbn + "'"
				end if
            end if

            if (FRectMinusInclude="P") then
                if FRectGubun="sys" then
                    retStr = retStr & "	and N"&stockColNm&">0 "
                else
                    retStr = retStr & "	and "&replace(stockColNm,"s.","Ns.")&">0 "
                end if
            elseif (FRectMinusInclude="M") then
                if FRectGubun="sys" then
                    retStr = retStr & "	and N"&stockColNm&"<0 "
                else
                    retStr = retStr & "	and "&replace(stockColNm,"s.","Ns.")&"<0 "
                end if
            else

            end if
        end if

        if (igrouping="item") or (FRectIsUsing<>"") or (FRectVatYn<>"") or (FRectNewItem<>"") then
            retStr = retStr & "		left join [db_item].[dbo].tbl_item i "
            retStr = retStr & "		on s.yyyymm='"&iYYYYMM&"'"
            retStr = retStr & "		and s.itemgubun='10' "
            retStr = retStr & "		and s.itemid=i.itemid "

            retStr = retStr & "		left join [db_shop].[dbo].tbl_shop_item si "
            retStr = retStr & "		on s.yyyymm='"&iYYYYMM&"'"
            retStr = retStr & "		and s.itemgubun<>'10' "
            retStr = retStr & "		and s.itemgubun=si.itemgubun"
            retStr = retStr & "		and s.itemid=si.shopitemid "
            retStr = retStr & "		and s.itemoption=si.itemoption "
        end if

        if (igrouping="item") then
            retStr = retStr + "     left join [db_item].[dbo].tbl_item_option o on s.yyyymm='"&iYYYYMM&"' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"
        end if

        retStr = retStr & "	where s.yyyymm='"&iYYYYMM&"'"
        retStr = retStr & "	and not (s.itemgubun='10' and s.itemid  in (0,11406,6400))"


        IF (Not isPreMonth) then
            if (FRectMakerid<>"") then
                retStr = retStr + " and s.lastmakerid='" + FRectMakerid + "'"
            end if
            if (FRectItemGubun<>"") then
                retStr = retStr & "	and s.itemgubun='"&FRectItemGubun&"'"
            end if
            if (FRectTargetGbn<>"") then
				if (FRectTargetGbn = "3X") then
					retStr = retStr + " and s.targetGbn not in ('ET', 'EG') "
				else
					retStr = retStr + " and s.targetGbn='" + FRectTargetGbn + "'"
				end if
            end if

            if (FRectMinusInclude="P") then
                retStr = retStr & "	and "&stockColNm&">0 "
            elseif (FRectMinusInclude="M") then
                retStr = retStr & "	and "&stockColNm&"<0 "
            else

            end if
        end if
'''수정요망..
'       if FRectNewItem<>"" then
'			retStr = retStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
'		end if
    if (Not isPreMonth) then
		if FRectIsUsing<>"" then
			retStr = retStr + " and isNULL(i.isusing,si.isusing)='" + FRectIsUsing + "'"
		end if
        if (FRectVatYn<>"") then
		    retStr = retStr + " and isNULL(isNULL(s.lastvatinclude,i.vatinclude),si.vatinclude)='" + FRectVatYn + "'"
		end if

		if FRectMwDiv<>"" then
			retStr = retStr + " and IsNULL(s.lastmwdiv,'Z')='" + FRectMwDiv + "'"
		end if

		'구매유형
		if FRectPurchaseType <> "" then
			if FRectPurchasetype="101" then
				retStr = retStr + " and p.purchasetype<>1 "
			elseif FRectPurchasetype="102" then
				retStr = retStr + " and p.purchasetype in (3,5,6) "
			else
				retStr = retStr + " and p.purchasetype = " & FRectPurchaseType & " "
			end if
		end if
    end if
		if (igrouping="itemgubun") then
            retStr = retStr & "	group by s.itemgubun, IsNULL(s.lastmwdiv,'Z'), s.targetGbn "
        elseif (igrouping="makerid") then
            retStr = retStr & "	group by s.lastmakerid "
        end if

        getComonSubQueryLogisticsFix = retStr
    end function

    public function getComonSubQueryLogistics(iYYYYMM,onOffGbn,igrouping)
        Dim retStr,stockColNm

        if FRectGubun="sys" then
		    stockColNm = "s.totsysstock"
		else
		    stockColNm = "(s.realstock-s.errbaditemno)" ''불량재고도 실재고임.
	    end if

	    retStr = ""

	    ''if (onOffGbn="ON") then

            if (igrouping="itemgubun") then
                retStr = retStr & "	select ( CASE WHEN p.id is NULL THEN 'ON' ELSE 'IT' END) as targetGbn, '10' as itemgubun"
                retStr = retStr & "	, IsNULL(s.lastmwdiv,'Z') as mwdiv "
            elseif (igrouping="makerid") then
                retStr = retStr & "	select i.makerid "
            elseif (igrouping="item") then
                retStr = retStr & "	select s.itemgubun,s.itemid,s.itemoption"
                retStr = retStr & "	, i.regdate, i.itemname, IsNULL(s.lastmwdiv,'Z') as mwdiv, IsNULL(o.optionname,'') as itemoptionname"
                retStr = retStr & "	, i.isusing, IsNULL(o.isusing,'Y') as optionusing"
                retStr = retStr & "	, ("&stockColNm&") as totno "
                retStr = retStr & "	, ("&stockColNm&"*i.orgsuplycash) as buysum "
                retStr = retStr & "	, (s.totipgono) as totipgono "
                retStr = retStr & "	, (s.totipgono*i.orgsuplycash) as ipgobuysum "
                retStr = retStr & "	, (s.lossno) as totlossno "
                retStr = retStr & "	, (s.lossno*i.orgsuplycash) as lossbuysum "
            end if

            if (igrouping<>"item") then
                retStr = retStr & "	, sum("&stockColNm&") as totno "
                retStr = retStr & "	, Sum("&stockColNm&"*i.orgsuplycash) as buysum "
                retStr = retStr & "	, sum(s.totipgono) as totipgono "
                retStr = retStr & "	, Sum(s.totipgono*i.orgsuplycash) as ipgobuysum "
                retStr = retStr & "	, sum(s.lossno) as totlossno "
                retStr = retStr & "	, Sum(s.lossno*i.orgsuplycash) as lossbuysum "
            end if
            retStr = retStr & "	from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s "
            retStr = retStr & "		join [db_item].[dbo].tbl_item i "
            retStr = retStr & "		on s.yyyymm='"&iYYYYMM&"'"
            retStr = retStr & "		and s.itemgubun='10' "
            retStr = retStr & "		and s.itemid=i.itemid "

            if (FRectItemGubun<>"") then
                retStr = retStr & "	and s.itemgubun='"&FRectItemGubun&"'"
            end if
            if (FRectTargetGbn<>"") then
				if (FRectTargetGbn = "3X") then
					retStr = retStr + " and s.targetGbn not in ('ET', 'EG') "
				else
					retStr = retStr + " and s.targetGbn='" + FRectTargetGbn + "'"
				end if
            end if
            if (igrouping="item") then
                retStr = retStr + "     and i.makerid='" + FRectMakerid + "'"
                retStr = retStr + "     left join [db_item].[dbo].tbl_item_option o on s.yyyymm='"&iYYYYMM&"' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"
            end if

            if (FRectITSOnlyOrNot="N") then
                retStr = retStr & "		left"
            end if
            if (FRectITSOnlyOrNot<>"") then
                retStr = retStr & "		Join db_partner.dbo.tbl_partner p"
                retStr = retStr & "		on p.groupid='"&iTsGroupID&"'"
                retStr = retStr & "		and i.makerid=p.id"
            end if

			if FRectPurchaseType <> "" then
				retStr = retStr & "	left join db_partner.dbo.tbl_partner pp "
				retStr = retStr & "	on "
				retStr = retStr & "		i.makerid = pp.id "
			end if

            retStr = retStr & "	where s.yyyymm='"&iYYYYMM&"'"
            retStr = retStr & "	and i.itemid not in (0,11406,6400)"
            if (FRectMinusInclude="P") then
                retStr = retStr & "	and "&stockColNm&">0 "
            elseif (FRectMinusInclude="M") then
                retStr = retStr & "	and "&stockColNm&"<0 "
            else

			'구매유형
			if FRectPurchaseType <> "" then
				if FRectPurchasetype="101" then
					retStr = retStr + " and pp.purchasetype<>1 "
				elseif FRectPurchasetype="102" then
					retStr = retStr + " and pp.purchasetype in (3,5,6) "
				else
					retStr = retStr + " and pp.purchasetype = " & FRectPurchaseType & " "
				end if
			end if

            end if
'            if (FRectITSOnlyOrNot="N") then
'                retStr = retStr & "	and p.id is NULL"
'            end if
            if FRectNewItem<>"" then
    			retStr = retStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    		end if
    		if FRectIsUsing<>"" then
    			retStr = retStr + " and i.isusing='" + FRectIsUsing + "'"
    		end if
    		if FRectMwDiv<>"" then
    			retStr = retStr + " and IsNULL(s.lastmwdiv,'Z')='" + FRectMwDiv + "'"
    		end if
    		if (FRectVatYn<>"") then
    		    retStr = retStr + " and isNULL(s.lastvatinclude,i.vatinclude)='" + FRectVatYn + "'"
    		end if

    		if (igrouping="itemgubun") then
                retStr = retStr & "	group by ( CASE WHEN p.id is NULL THEN 'ON' ELSE 'IT' END) , IsNULL(s.lastmwdiv,'Z') "
            elseif (igrouping="makerid") then
                retStr = retStr & "	group by i.makerid "
            end if
        retStr = retStr & "	union"
        ''elseif (onOffGbn="OF") then
            if (igrouping="itemgubun") then
                retStr = retStr & "	select ( CASE WHEN p.id is NULL THEN 'OF' ELSE 'IT' END) as targetGbn, s.itemgubun as itemgubun"
                retStr = retStr & "	, IsNULL(s.lastmwdiv,'Z') as mwdiv "
            elseif (igrouping="makerid") then
                retStr = retStr & "	select i.makerid "
            elseif (igrouping="item") then
                retStr = retStr & "	select s.itemgubun,s.itemid,s.itemoption"
                retStr = retStr & "	, i.regdate, i.shopitemname as itemname, IsNULL(s.lastmwdiv,'Z') as mwdiv, i.shopitemoptionname as itemoptionname"
                retStr = retStr & "	, i.isusing, i.isusing as optionusing"
                retStr = retStr & "	, ("&stockColNm&") as totno "
                retStr = retStr & "	, ("&stockColNm&"*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buysum "
                retStr = retStr & "	, (s.totipgono) as totipgono "
                retStr = retStr & "	, (s.totipgono*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as ipgobuysum "
                retStr = retStr & "	, (s.lossno) as totlossno "
                retStr = retStr & "	, (s.lossno*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as lossbuysum "
            end if

            if (igrouping<>"item") then
                retStr = retStr & "	, sum("&stockColNm&") as totno "
                retStr = retStr & "	, Sum("&stockColNm&"*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buysum "
                retStr = retStr & "	, sum(s.totipgono) as totipgono "
                retStr = retStr & "	, Sum(s.totipgono*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as ipgobuysum "
                retStr = retStr & "	, sum(s.lossno) as totlossno "
                retStr = retStr & "	, Sum(s.lossno*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as lossbuysum "
            end if

            retStr = retStr & "	from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s "
            retStr = retStr & "		join [db_shop].[dbo].tbl_shop_item i "
            retStr = retStr & "		on s.yyyymm='"&iYYYYMM&"'"
            retStr = retStr & "		and s.itemgubun<>'10'"
            retStr = retStr & "		and s.itemgubun=i.itemgubun "
            retStr = retStr & "		and s.itemid=i.shopitemid "
            retStr = retStr & "		and s.itemoption=i.itemoption  "
            if (FRectItemGubun<>"") then
                retStr = retStr & "	and s.itemgubun='"&FRectItemGubun&"'"
            end if
            if (FRectTargetGbn<>"") then
				if (FRectTargetGbn = "3X") then
					retStr = retStr + " and s.targetGbn not in ('ET', 'EG') "
				else
					retStr = retStr + " and s.targetGbn='" + FRectTargetGbn + "'"
				end if
            end if
            if (igrouping="item") then
                retStr = retStr + "     and i.makerid='" + FRectMakerid + "'"
            end if
            retStr = retStr & "		left Join db_shop.dbo.tbl_shop_designer d "
            retStr = retStr & "		on d.shopid='streetshop000' "
            retStr = retStr & "		and d.makerid=i.makerid "

            if (FRectITSOnlyOrNot="N") then
                retStr = retStr & "		left"
            end if
            if (FRectITSOnlyOrNot<>"") then
                retStr = retStr & "		Join db_partner.dbo.tbl_partner p"
                retStr = retStr & "		on p.groupid='"&iTsGroupID&"'"
                retStr = retStr & "		and i.makerid=p.id"
            end if

			if FRectPurchaseType <> "" then
				retStr = retStr & "	left join db_partner.dbo.tbl_partner pp "
				retStr = retStr & "	on "
				retStr = retStr & "		i.makerid = pp.id "
			end if

			retStr = retStr & "	where s.yyyymm='"&iYYYYMM&"'"
            ''retStr = retStr & "	and i.itemid not in (0)"
            if (FRectMinusInclude="P") then
                retStr = retStr & "	and "&stockColNm&">0 "
            elseif (FRectMinusInclude="M") then
                retStr = retStr & "	and "&stockColNm&"<0 "
            else

			'구매유형
			if FRectPurchaseType <> "" then
				if FRectPurchasetype="101" then
					retStr = retStr + " and pp.purchasetype<>1 "
				elseif FRectPurchasetype="102" then
					retStr = retStr + " and pp.purchasetype in (3,5,6) "
				else
					retStr = retStr + " and pp.purchasetype = " & FRectPurchaseType & " "
				end if
			end if

            end if
'            if (FRectITSOnlyOrNot="N") then
'                retStr = retStr & "	and p.id is NULL"
'            end if
            if FRectNewItem<>"" then
    			retStr = retStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    		end if
    		if FRectIsUsing<>"" then
    			retStr = retStr + " and i.isusing='" + FRectIsUsing + "'"
    		end if
    		if FRectMwDiv<>"" then
    			retStr = retStr + " and IsNULL(s.lastmwdiv,'Z')='" + FRectMwDiv + "'"
    		end if
    		if (FRectVatYn<>"") then
    		    retStr = retStr + " and isNULL(s.lastvatinclude,i.vatinclude)='" + FRectVatYn + "'"
    		end if

    		if (igrouping="itemgubun") then
                retStr = retStr & "	group by ( CASE WHEN p.id is NULL THEN 'OF' ELSE 'IT' END) , s.itemgubun, IsNULL(s.lastmwdiv,'Z') "
            elseif (igrouping="makerid") then
                retStr = retStr & "	group by i.makerid"
            end if
        ''end if
        getComonSubQueryLogistics = retStr
    end function

    ''상품별 합계
    public Sub GetMonthlyRealJeagoDetailByMakerWithPreMonth()
        dim sqlStr
        Dim stockColNm
        Dim PreYYYYMM
        PreYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)

        sqlStr = ""
        sqlStr = " select A.itemgubun, A.itemid, A.itemoption, A.itemname, A.itemoptionname, A.isusing, A.optionusing,A.regdate,A.mwdiv"
        sqlStr = sqlStr & " ,A.totno,A.buysum,isNULL(B.totno,0) as ptotno,isNULL(B.buysum,0) as pbuysum"
        sqlStr = sqlStr & " ,A.totipgono-isNULL(B.totipgono,0) as ipno"
        sqlStr = sqlStr & " ,A.ipgobuysum-isNULL(B.ipgobuysum,0) as ipbuysum"
        sqlStr = sqlStr & " ,A.totlossno-isNULL(B.totlossno,0) as totlossno"
        sqlStr = sqlStr & " ,A.lossbuysum-isNULL(B.lossbuysum,0) as lossbuysum"
        sqlStr = sqlStr & " from ("
        if (FRectIsFix="on") then
            sqlStr = sqlStr &   getComonSubQueryLogisticsFIX(FRectYYYYMM,"","item")
        else
            sqlStr = sqlStr &   getComonSubQueryLogistics(FRectYYYYMM,"","item")
        end if
        sqlStr = sqlStr & " ) A"
        sqlStr = sqlStr & " left join ("
        if (FRectIsFix="on") then
            sqlStr = sqlStr &   getComonSubQueryLogisticsFIX(PreYYYYMM,"","item")
        else
            sqlStr = sqlStr &   getComonSubQueryLogistics(PreYYYYMM,"","item")
        end if
        sqlStr = sqlStr & " ) B"
        sqlStr = sqlStr & " on A.itemgubun=B.itemgubun"
        sqlStr = sqlStr & " and A.itemid=B.itemid"
        sqlStr = sqlStr & " and A.itemoption=B.itemoption"
        if (FRectOrdTp="S") then
            sqlStr = sqlStr & " order by A.totno desc,A.itemgubun,A.itemid,A.itemoption"
        else
            sqlStr = sqlStr & " order by A.itemgubun,A.itemid,A.itemoption"
        end if
''rw sqlStr
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
				set FItemList(i) = new CMonthlyStockSum
				if (FRectITSOnlyOrNot="N") then
				    FItemList(i).FBusiName      = "온라인"
				elseif (FRectITSOnlyOrNot="O") then
				    FItemList(i).FBusiName      = "아이띵소"
				else
				    FItemList(i).FBusiName      = "온라인+ITS"
			    end if

				FItemList(i).Fitemgubun     = rsget("itemgubun")
				FItemList(i).Fitemid 		= rsget("itemid")
				FItemList(i).Fregdate 		= rsget("regdate")
				FItemList(i).Fitemoption 	= rsget("itemoption")
				FItemList(i).Fitemname 		= db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).FMaeIpGubun	= rsget("mwdiv")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")

				FItemList(i).FTotPreCount   = rsget("ptotno")
            	FItemList(i).FTotPreBuySum  = rsget("pbuysum")
            	FItemList(i).FTotIpCount    = rsget("ipno")
            	FItemList(i).FTotIpBuySum   = rsget("ipbuysum")
            	FItemList(i).FTotLossCount    = rsget("totlossno")
            	FItemList(i).FTotLossBuySum   = rsget("lossbuysum")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end Sub

    ''브랜드별 합계 // ponybrown CASE 11월 전체 위탁으로 변경됨 10월에는 매입인 내역있음..
    public Sub GetMonthlyRealJeagoDetailWithPreMonth()
        dim sqlStr
        Dim stockColNm
        Dim PreYYYYMM
        PreYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)

        sqlStr = ""
        sqlStr = " select A.lastmakerid as makerid,sum(A.totno) as totno,sum(A.buysum) as buysum,sum(isNULL(B.totno,0)) as ptotno,sum(isNULL(B.buysum,0)) as pbuysum"
        sqlStr = sqlStr & " ,sum(A.totipgono)-sum(isNULL(B.totipgono,0)) as ipno"
        sqlStr = sqlStr & " ,sum(A.ipgobuysum)-sum(isNULL(B.ipgobuysum,0)) as ipbuysum"
        sqlStr = sqlStr & " ,sum(A.totlossno)-sum(isNULL(B.totlossno,0)) as totlossno"
        sqlStr = sqlStr & " ,sum(A.lossbuysum)-sum(isNULL(B.lossbuysum,0)) as lossbuysum"
        sqlStr = sqlStr & " , c.socname, c.isusing"
        sqlStr = sqlStr & " from ("
        if (FRectIsFix="on") then
            sqlStr = sqlStr &   getComonSubQueryLogisticsFIX(FRectYYYYMM,"","item")
        else
            sqlStr = sqlStr &   getComonSubQueryLogistics(FRectYYYYMM,"","item")
        end if
        sqlStr = sqlStr & " ) A"
        sqlStr = sqlStr & " left join ("
        if (FRectIsFix="on") then
            sqlStr = sqlStr &   getComonSubQueryLogisticsFIX(PreYYYYMM,"","item")
        else
            sqlStr = sqlStr &   getComonSubQueryLogistics(PreYYYYMM,"","item")
        end if
        sqlStr = sqlStr & " ) B"
        sqlStr = sqlStr & " on A.itemgubun=B.itemgubun"
        sqlStr = sqlStr & " and A.itemid=B.itemid"
        sqlStr = sqlStr & " and A.itemoption=B.itemoption"
        sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on A.lastmakerid=c.userid"
        sqlStr = sqlStr + " group by A.lastmakerid, c.socname, c.isusing"

        sqlStr = sqlStr + " having sum(A.totno)<>0 "
        sqlStr = sqlStr + "     or sum(isNULL(B.totno,0))<>0"
        sqlStr = sqlStr + "     or sum(A.totipgono)-sum(isNULL(B.totipgono,0))<>0"
        sqlStr = sqlStr & "     or sum(A.totlossno)-sum(isNULL(B.totlossno,0))<>0"
        'sqlStr = sqlStr + " where A.totno<>0 "
        'sqlStr = sqlStr + "     or isNULL(B.totno,0)<>0"
        'sqlStr = sqlStr + "     or A.totipgono-isNULL(B.totipgono,0)<>0"
        'sqlStr = sqlStr & "     or A.totlossno-isNULL(B.totlossno,0)<>0"

        sqlStr = sqlStr & " order by totno desc, makerid"
''rw sqlStr
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
				set FItemList(i) = new CMonthlyStockSum
				if (FRectITSOnlyOrNot="N") then
				    FItemList(i).FBusiName      = "온라인"
				elseif (FRectITSOnlyOrNot="O") then
				    FItemList(i).FBusiName      = "아이띵소"
				else
				    FItemList(i).FBusiName      = "온라인+ITS"
			    end if

				FItemList(i).Fmakerid     = rsget("makerid")
				''FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")

				FItemList(i).FTotPreCount   = rsget("ptotno")
            	FItemList(i).FTotPreBuySum  = rsget("pbuysum")
            	FItemList(i).FTotIpCount    = rsget("ipno")
            	FItemList(i).FTotIpBuySum   = rsget("ipbuysum")
            	FItemList(i).FTotLossCount    = rsget("totlossno")
            	FItemList(i).FTotLossBuySum   = rsget("lossbuysum")
            	FItemList(i).FMakerUsing	= rsget("isusing")
            	FItemList(i).Fsocname       = db2HTML(rsget("socname"))
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end Sub

    public Sub GetMonthlyEtcChulgoList
        dim sqlStr, i
        if (FRectSocID="") and (FRectIpChulCode="") then
            FRectListType ="1"
        elseif (FRectIpChulCode<>"") then
            FRectListType ="3"
        elseif (FRectSocID<>"") then
            FRectListType ="2"
        else
            FRectListType ="1"
        end if


		sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_EtcChulgoList] '"&FRectYYYYMM&"', "&FRectListType&", '"&FRectItemGubun&"','"&FRectMwDiv&"', '"&FRectTargetGbn&"', '"&FRectSocID&"', '"&FRectIpChulCode&"','"&FRectShopSuplyPrice&"','"&FRectPlaceGubun&"','"&FRectShopID&"', '" & FRectChulgoGubun & "', '" & FRectGrpType & "' "
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic     '' adOpenForwardOnly?
		rsget.CursorType = adLockReadOnly
		''rsget.LockType = adLockOptimistic '' 느림 ==> adLockReadOnly (디폴트)
rw sqlStr
'response.end

		dbget.CommandTimeout = 60*2   ' 2분
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
				set FItemList(i) = new CMonthlyStockEtcChulGoItem
                FItemList(i).FSocID         = rsget("socid")
                FItemList(i).FSocName       = db2HTML(rsget("socname_kor"))
                FItemList(i).FIpChulMwGubun = rsget("MwGubun")
                FItemList(i).FItemgubun     = rsget("itemgubun")

                FItemList(i).FtargetGbn     = rsget("targetGbn")
				FItemList(i).Flastmwdiv 	= rsget("lastmwdiv")

                FItemList(i).FTTLCNT        = rsget("TTLCNT")
                FItemList(i).FTTLSellSum    = CLNG(rsget("TTLSellSum"))
                FItemList(i).FTTLBuySum     = CLNG(rsget("TTLBuySum"))
                FItemList(i).FTTLSuplySum   = CLNG(rsget("TTLSuplySum"))

                FItemList(i).FMayStockPrice     = (rsget("MayStockPrice"))
                FItemList(i).FMaeipLedgeravgipgoPrice = (rsget("MaeipLedgeravgipgoPrice"))
                if isNULL(FItemList(i).FMaeipLedgeravgipgoPrice) then FItemList(i).FMaeipLedgeravgipgoPrice=0

                if isNULL(FItemList(i).FMayStockPrice) then
                    FItemList(i).FMayStockPrice=0
                else
                    FItemList(i).FMayStockPrice=CLNG(FItemList(i).FMayStockPrice)
                end if

				if (FRectPriceGubun = "V") then
				    if isNULL(rsget("TTLBuySumAvg")) then
				        FItemList(i).FMayStockPrice     = 0
				    else
					    FItemList(i).FMayStockPrice     = CLNG(rsget("TTLBuySumAvg"))
				    end if
				end if


                if (FRectListType>="1") then
                    FItemList(i).FIpChulCode    = rsget("code")
					FItemList(i).FMakerid       = rsget("makerid")
                end if

                ''상세.
                if (FRectListType>"2") then
                    FItemList(i).FItemID        = rsget("itemid")
                    FItemList(i).FItemOption    = rsget("itemoption")
                    FItemList(i).FMakerid       = rsget("makerid")
                   ' FItemList(i).FItemName      = db2HTML(rsget("iitemname"))
    			   ' FItemList(i).FItemOptionName= db2HTML(rsget("iitemOptionname"))
                end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    End Sub

    public Sub GetMonthlyJeagoSumSummary ''재고 서머리 신버전(201305이후)
        dim sqlStr

		if (FRectPlaceGubun = "") then
			FRectPlaceGubun = "L"
		end if

        if (FRectMakerid<>"") then '' 상품별
            sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Stockledger_ItemList] '"&FRectYYYYMM&"','" + CStr(FRectPlaceGubun) + "','" + CStr(FRectShopID) + "','"&FRectTargetGbn&"','"&FRectItemGubun&"','"&FRectMwDiv&"','"&FRectVatYn&"','"&FRectPurchaseType&"','"&FRectShopSuplyPrice&"','"&FRectMakerid&"','"&FRectetcjungsantype&"'"
        ''elseif (FRectGroupbyType=1) then ''브랜드별
        ''    sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Stockledger_BrandList] '"&FRectYYYYMM&"','L','"&FRectTargetGbn&"','"&FRectItemGubun&"','"&FRectMwDiv&"','"&FRectVatYn&"','"&FRectPurchaseType&"','"&FRectShopSuplyPrice&"'"
        else
            sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Stockledger_List] '"&FRectYYYYMM&"','" + CStr(FRectPlaceGubun) + "','" + CStr(FRectShopID) + "','"&FRectTargetGbn&"','"&FRectItemGubun&"','"&FRectMwDiv&"','"&FRectVatYn&"','"&FRectPurchaseType&"','"&FRectShopSuplyPrice&"','"&FRectGroupbyType&"','"&FRectetcjungsantype&"', '" & FRectBrandUseYN & "'"
        end if

		response.write sqlStr
		''response.end
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		''rsget.LockType = adLockOptimistic  ''==> adLockReadOnly

''response.write sqlStr
		''response.end
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
				set FItemList(i) = new CMonthlyStockSum

			    FItemList(i).FtargetGbn     = rsget("targetGbn")
				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FMaeIpGubun 	= rsget("lastmwdiv")
				FItemList(i).Fshopid 		= rsget("shopid")

				FItemList(i).FTotCount 		= rsget("curSysStockNo")
				FItemList(i).FTotBuySum 	= rsget("curSysStockSum")

				FItemList(i).FTotRealStockCount 	= rsget("curRealStockNo")
				FItemList(i).FTotRealStockBuySum 	= rsget("curRealStockSum")

				FItemList(i).FTotPreCount   = rsget("preSysStockNo")
            	FItemList(i).FTotPreBuySum  = rsget("preSysStockSum")
				FItemList(i).FTotPreRealStockCount   = rsget("preRealStockNo")
            	FItemList(i).FTotPreRealStockBuySum  = rsget("preRealStockSum")
            	FItemList(i).FTotIpCount    = rsget("curIpNo")
            	FItemList(i).FTotIpBuySum   = rsget("curIpSum")
            	FItemList(i).FTotLossCount    = rsget("curLossNo")
            	FItemList(i).FTotLossBuySum   = rsget("curLossSum")

                FItemList(i).FTotSellCount         = rsget("curSellNo")
                FItemList(i).FTotSellBuySum        = rsget("curSellSum")
                FItemList(i).FTotOffChulCount      = rsget("curoffChulNo")
                FItemList(i).FTotOffChulBuySum     = rsget("curoffChulSum")
                FItemList(i).FTotEtcChulCount      = rsget("curEtcChulNo")
                FItemList(i).FTotEtcChulBuySum     = rsget("curEtcChulSum")
                FItemList(i).FTotCsChulCount       = rsget("curCsNo")
                FItemList(i).FTotCsChulBuySum      = rsget("curCsSum")
                FItemList(i).FTotErrRealCheckCount = rsget("curErrRealCheckNo")
                FItemList(i).FTotErrRealCheckBuySum= rsget("curErrRealCheckSum")

				FItemList(i).FTotErrBadItemCount = rsget("curErrBadItemNo")
				FItemList(i).FTotErrBadItemBuySum= rsget("curErrBadItemSum")

				FItemList(i).FTotMoveItemCount = rsget("curMoveItemNo")
				FItemList(i).FTotMoveItemBuySum= rsget("curMoveItemSum")

				if (FRectPriceGubun = "V") then
					FItemList(i).FTotBuySum 	= rsget("curSysStockSumAvg")
					FItemList(i).FTotPreBuySum  = rsget("preSysStockSumAvg")

					FItemList(i).FTotIpBuySum   = rsget("curIpSumAvg")
					FItemList(i).FTotLossBuySum   = rsget("curLossSumAvg")
					FItemList(i).FTotSellBuySum        = rsget("curSellSumAvg")
					FItemList(i).FTotOffChulBuySum     = rsget("curoffChulSumAvg")
					FItemList(i).FTotEtcChulBuySum     = rsget("curEtcChulSumAvg")
					FItemList(i).FTotCsChulBuySum      = rsget("curCsSumAvg")
					FItemList(i).FTotErrRealCheckBuySum= rsget("curErrRealCheckSumAvg")
					FItemList(i).FTotErrBadItemBuySum  = rsget("curErrBadItemSumAvg")
					FItemList(i).FTotMoveItemBuySum    = rsget("curMoveItemSumAvg")
					FItemList(i).FTotRealStockBuySum   = rsget("curRealStockSumAvg")
				end if

                ''일단 추가 물류는 기타출고에 로스가 포함되어 있음
                if (FRectPlaceGubun = "L") then
                    FItemList(i).FTotEtcChulCount = FItemList(i).FTotEtcChulCount-FItemList(i).FTotLossCount
                    FItemList(i).FTotEtcChulBuySum = FItemList(i).FTotEtcChulBuySum-FItemList(i).FTotLossBuySum
                end if

                if (FRectMakerid<>"") then
                    FItemList(i).FShopId        = rsget("shopid")
                    FItemList(i).FItemid     = rsget("itemid")
                    FItemList(i).FItemOption     = rsget("itemoption")
                else
                    FItemList(i).FMakerid= rsget("makerid")
                end if

				if (FRectMakerid<>"") then
					FItemList(i).FlastIpgoDate = rsget("lastIpgoDate")
					FItemList(i).fitemname = rsget("itemname")
				end if

                if (FRectMakerid="") then ''??
				    FItemList(i).FpurchaseType = rsget("purchaseType")
					FItemList(i).fpurchasetypename = rsget("purchasetypename")
                end if

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end Sub

    ''물류 일반상품 신버전(201210이후)
    public Sub GetMonthlyJeagoSumWithPreMonth()
        dim sqlStr
        Dim stockColNm
        Dim PreYYYYMM

        PreYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)

        sqlStr = ""
        sqlStr = " select A.targetGbn, A.itemgubun,A.mwdiv,sum(A.totno) as totno,sum(A.buysum) as buysum,sum(isNULL(B.totno,0)) as ptotno,sum(isNULL(B.buysum,0)) as pbuysum"
        sqlStr = sqlStr & " ,sum(A.totipgono)-sum(isNULL(B.totipgono,0)) as ipno"
        sqlStr = sqlStr & " ,sum(A.ipgobuysum)-sum(isNULL(B.ipgobuysum,0)) as ipbuysum"
        sqlStr = sqlStr & " ,sum(A.totlossno)-sum(isNULL(B.totlossno,0)) as totlossno"
        sqlStr = sqlStr & " ,sum(A.lossbuysum)-sum(isNULL(B.lossbuysum,0)) as lossbuysum"
        sqlStr = sqlStr & " from ("
        if (FRectIsFix="on") then
            sqlStr = sqlStr &   getComonSubQueryLogisticsFIX(FRectYYYYMM,"","item")
        else
            sqlStr = sqlStr &   getComonSubQueryLogistics(FRectYYYYMM,"","item")
        end if
        sqlStr = sqlStr & " ) A"
        sqlStr = sqlStr & " left join ("
        if (FRectIsFix="on") then
            sqlStr = sqlStr &   getComonSubQueryLogisticsFIX(PreYYYYMM,"","item")
        else
            sqlStr = sqlStr &   getComonSubQueryLogistics(PreYYYYMM,"","item")
        end if
        sqlStr = sqlStr & " ) B"
        sqlStr = sqlStr & " on A.itemgubun=B.itemgubun"
        sqlStr = sqlStr & " and A.itemid=B.itemid"
        sqlStr = sqlStr & " and A.itemoption=B.itemoption"

        ''sqlStr = sqlStr & " on A.itemgubun=B.itemgubun "
        ''sqlStr = sqlStr & " and A.mwdiv=B.MwDiv"
        ''sqlStr = sqlStr & " and isNULL(A.targetGbn,'')=isNULL(B.targetGbn,'')"
        sqlStr = sqlStr & " group by A.targetGbn, A.itemgubun,A.mwdiv"
        sqlStr = sqlStr & " order by A.targetGbn desc, A.itemgubun,A.mwdiv"

		''response.write sqlStr
'' 서머리로 변경
'        sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Acc_LogisStockList] '"&FRectYYYYMM&"','"&FRectShopid&"','"&FRectMwDiv&"','"&FRectGubun&"',"&CHKIIF(FRectShowMinus<>"on","0","1")&",'"&FRectIsUsing&"','"&FRectVatYn&"'"
'
'        rsget.CursorLocation = adUseClient
'		rsget.CursorType = adOpenStatic
'		rsget.LockType = adLockOptimistic  '' ==> adLockReadOnly


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
				set FItemList(i) = new CMonthlyStockSum

			    FItemList(i).FtargetGbn     = rsget("targetGbn")
				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")

				FItemList(i).FTotPreCount   = rsget("ptotno")
            	FItemList(i).FTotPreBuySum  = rsget("pbuysum")
            	FItemList(i).FTotIpCount    = rsget("ipno")
            	FItemList(i).FTotIpBuySum   = rsget("ipbuysum")
            	FItemList(i).FTotLossCount    = rsget("totlossno")
            	FItemList(i).FTotLossBuySum   = rsget("lossbuysum")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end Sub

	'// 재고평가충당금 : 물류
    public Sub GetJeagoOverValueSum()
    	dim sqlStr
		dim colName, valPrice, valPriceFieldName

		valPriceFieldName = " s.lastbuyPrice "
		if (FRectPriceGubun = "V") then
			''valPriceFieldName = " IsNull(s.avgipgoPrice, s.lastbuyPrice) "
			valPriceFieldName = " IsNull(s.avgipgoPrice, 0) "
		end if

		if (FRectGubun = "sys") then
			colName = "s.totsysstock"
		else
			colName = "(s.realstock - s.errbaditemno)"			'// 불량재고 포함
		end if

		if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(isNULL(isNULL(s.lastvatinclude,i.vatinclude),si.vatinclude), 'Y') = 'Y' then (" + CStr(valPriceFieldName) + "*10/11) else " + CStr(valPriceFieldName) + " end)"
			if (FRectPriceGubun = "V") then
				valPrice = "(case when IsNull(isNULL(isNULL(s.lastvatinclude,i.vatinclude),si.vatinclude), 'Y') = 'Y' then Round((" + CStr(valPriceFieldName) + "*10/11), 0) else " + CStr(valPriceFieldName) + " end)"
			end if
		else
			valPrice = valPriceFieldName
		end if

		sqlStr = " SELECT "
		sqlStr = sqlStr & " 	s.targetGbn, s.itemgubun, IsNULL(s.lastmwdiv,'Z') as mwdiv, "
		sqlStr = sqlStr & " 	SUM("+CStr(colName)+") as totStockNo,"
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun1, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun2, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun3, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun4, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun5, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN IsNull(s.lastIpgoDate, '') = '' "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun6, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun7, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun8, "

		'// 연도별
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun11, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun12, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun13, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun14, "

		sqlStr = sqlStr & " 	SUM(" + CStr(colName) + " * " + CStr(valPrice) + ") as MonthGubunSUM "
		sqlStr = sqlStr & " FROM db_summary.dbo.tbl_monthly_accumulated_logisstock_summary s with (nolock)"

		if FRectPurchaseType <> "" then
			sqlStr = sqlStr & "	left join db_partner.dbo.tbl_partner p with (nolock)"
			sqlStr = sqlStr & "	on "
			sqlStr = sqlStr & "		s.lastmakerid = p.id "
		end if

        if FRectVatYn<>"" or (FRectShopSuplyPrice = "Y") Or (FRectShowUpbae <> "") then
            sqlStr = sqlStr & "	left join [db_item].[dbo].tbl_item i with (nolock)"
            sqlStr = sqlStr & "		on s.yyyymm='"&FRectYYYYMM&"'"
            sqlStr = sqlStr & "		and s.itemgubun='10' "
            sqlStr = sqlStr & "		and s.itemid=i.itemid "

            sqlStr = sqlStr & "	left join [db_shop].[dbo].tbl_shop_item si with (nolock)"
            sqlStr = sqlStr & "		on s.yyyymm='"&FRectYYYYMM&"'"
            sqlStr = sqlStr & "		and s.itemgubun<>'10' "
            sqlStr = sqlStr & "		and s.itemgubun=si.itemgubun"
            sqlStr = sqlStr & "		and s.itemid=si.shopitemid "
            sqlStr = sqlStr & "		and s.itemoption=si.itemoption "
        end if

		sqlStr = sqlStr & " WHERE "
		sqlStr = sqlStr & " 	1 = 1 "
		sqlStr = sqlStr & " 	and s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		sqlStr = sqlStr & "		and not (s.itemgubun='10' and s.itemid  in (0,11406,6400))"
		''sqlStr = sqlStr & " 	and IsNULL(s.lastmwdiv,'Z') <> 'W' "

		if (FRectMwDiv <> "") then
			sqlStr = sqlStr & " 	and IsNULL(s.lastmwdiv,'Z') = '" + CStr(FRectMwDiv) + "' "
		end if

		if (FRectItemGubun <> "") then
			sqlStr = sqlStr & " 	and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
		end If

		if (FRectShowUpbae <> "") then
			sqlStr = sqlStr & " 	and i.mwdiv = 'U' "
		end if

		'구매유형
		if FRectPurchaseType <> "" then
			if FRectPurchasetype="101" then
				sqlStr = sqlStr + " and p.purchasetype<>1 "
			elseif FRectPurchasetype="102" then
				sqlStr = sqlStr + " and p.purchasetype in (3,5,6) "
			else
				sqlStr = sqlStr + " and p.purchasetype = " & FRectPurchaseType & " "
			end if
		end if

        if (FRectVatYn<>"") then
		    sqlStr = sqlStr + " and IsNull(isNULL(isNULL(s.lastvatinclude,i.vatinclude),si.vatinclude), 'Y')='" + FRectVatYn + "'"
		end if

        if (FRectTargetGbn<>"") then
			if (FRectTargetGbn = "3X") then
				sqlStr = sqlStr + " and s.targetGbn not in ('ET', 'EG') "
			else
				sqlStr = sqlStr + " and s.targetGbn='" + FRectTargetGbn + "'"
			end if
		end If

		if FRectDispCate<>"" Then
			sqlStr = sqlStr + " and s.itemgubun = '10' "
		    if LEN(FRectDispCate)>3 then
		         sqlStr = sqlStr + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"
		    end if
			sqlStr = sqlStr + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

		sqlStr = sqlStr & " 	and " + CStr(colName) + " <> 0 "
		sqlStr = sqlStr & " GROUP BY "
		sqlStr = sqlStr & " 	s.targetGbn, s.itemgubun, s.lastMWDiv "
		'sqlStr = sqlStr & " ORDER BY s.targetGbn DESC, s.itemgubun asc, IsNULL(s.lastmwdiv,'Z') asc"
		sqlStr = sqlStr & " ORDER BY s.itemgubun asc, IsNULL(s.lastmwdiv,'Z') asc"	' 2023.08.04 이문재이사님 요청(10 / 55 /90 위로 75 /85 등 저장품 아래로)

		'response.write sqlStr & "<br>"
		''response.end
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
			do until rsget.eof
				set FItemList(i) = new CStockOverValueSum

			    FItemList(i).FtargetGbn     = rsget("targetGbn")
				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")
				FItemList(i).FTotBuySum1 	= rsget("MonthGubun1")
				FItemList(i).FTotBuySum2 	= rsget("MonthGubun2")
				FItemList(i).FTotBuySum3 	= rsget("MonthGubun3")
				FItemList(i).FTotBuySum4 	= rsget("MonthGubun4")
				FItemList(i).FTotBuySum5 	= rsget("MonthGubun5")
				FItemList(i).FTotBuySum6 	= rsget("MonthGubun6")
				FItemList(i).FTotBuySum7 	= rsget("MonthGubun7")
				FItemList(i).FTotBuySum8 	= rsget("MonthGubun8")
				FItemList(i).FTotBuySum11 	= rsget("MonthGubun11")
				FItemList(i).FTotBuySum12 	= rsget("MonthGubun12")
				FItemList(i).FTotBuySum13 	= rsget("MonthGubun13")
				FItemList(i).FTotBuySum14 	= rsget("MonthGubun14")
				FItemList(i).FTotBuySum 	= rsget("MonthGubunSUM")
                FItemList(i).FtotStockNo    = rsget("totStockNo")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

	'// 재고평가충당금 : 물류 : 상세
    public Sub GetJeagoOverValueDetailSum()
    	dim sqlStr, sqlAdd
		dim colName, valPrice, valPriceFieldName
        Dim isPagingReq : isPagingReq=False

		valPriceFieldName = " s.lastbuyPrice "
		if (FRectPriceGubun = "V") then
			valPriceFieldName = " IsNull(s.avgipgoPrice, s.lastbuyPrice) "
		end if

        if (FRectMakerid <> "")or(FRectShowItemList<>"") then
            isPagingReq = TRUE
        end if

		if (FRectGubun = "sys") then
			colName = "s.totsysstock"
		else
			colName = "(s.realstock - s.errbaditemno)"			'// 불량재고 포함
		end if

		if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(isNULL(isNULL(s.lastvatinclude,i.vatinclude),si.vatinclude), 'Y') = 'Y' then (" + CStr(valPriceFieldName) + "*10/11) else " + CStr(valPriceFieldName) + " end)"
			if (FRectPriceGubun = "V") then
				valPrice = "(case when IsNull(isNULL(isNULL(s.lastvatinclude,i.vatinclude),si.vatinclude), 'Y') = 'Y' then Round((" + CStr(valPriceFieldName) + "*10/11), 0) else " + CStr(valPriceFieldName) + " end)"
			end if
		else
			valPrice = valPriceFieldName
		end if

        sqlAdd = ""

        if (FRectMonthGubun = "1") then
			'// 1-3개월
			sqlAdd = sqlAdd & " 	and DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 "
		elseif (FRectMonthGubun = "2") then
			'// 4-6개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) "
		elseif (FRectMonthGubun = "3") then
			'// 7-12개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) "
		elseif (FRectMonthGubun = "4") then
			'// 13-24개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		elseif (FRectMonthGubun = "5") then
			'// 25개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) "
		elseif (FRectMonthGubun = "6") then
			'// 미지정
			sqlAdd = sqlAdd & " 	and IsNull(s.lastIpgoDate, '') = '' "
		elseif (FRectMonthGubun = "7") then
			'// 13-18개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) "
		elseif (FRectMonthGubun = "8") then
			'// 19-24개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		elseif (FRectMonthGubun = "11") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0) "
		elseif (FRectMonthGubun = "12") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1) "
		elseif (FRectMonthGubun = "13") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2) "
		elseif (FRectMonthGubun = "14") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3) "
		end if

		if (FRectMwDiv <> "") then
			sqlAdd = sqlAdd & " 	and IsNULL(s.lastmwdiv,'Z') = '" + CStr(FRectMwDiv) + "' "
		end if

		If (FRectStartDate <> "") And IsNumeric(FRectStartDate) Then
			sqlAdd = sqlAdd + " and DateDiff(m, s.lastIpgoDate + '-01', getdate()) >= " & FRectStartDate & " "
		End If

		If (FRectEndDate <> "") And IsNumeric(FRectEndDate) Then
			sqlAdd = sqlAdd + " and DateDiff(m, s.lastIpgoDate + '-01', getdate()) <= " & FRectEndDate & " "
		End If

		if (FRectItemGubun <> "") then
			sqlAdd = sqlAdd & " 	and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
		end if

		'구매유형
        If FRectPurchasetype <> "" Then
            Select Case FRectPurchasetype
                Case "101"
                    '// 일반유통 제외
                    sqlAdd = sqlAdd & " and p.purchasetype <> 1 "
                Case "102"
                    '// 전략상품만
                    sqlAdd = sqlAdd & " and p.purchasetype in (3,5,6) "
                Case Else
                    sqlAdd = sqlAdd & " and p.purchasetype = '" & FRectPurchasetype &"'"
            End Select
		End IF

        if (FRectVatYn<>"") then
		    sqlAdd = sqlAdd + " and IsNull(isNULL(isNULL(s.lastvatinclude,i.vatinclude),si.vatinclude), 'Y')='" + FRectVatYn + "'"
		end if

        if (FRectTargetGbn<>"") then
			if (FRectTargetGbn = "3X") then
				sqlAdd = sqlAdd + " and s.targetGbn not in ('ET', 'EG') "
			else
				sqlAdd = sqlAdd + " and s.targetGbn='" + FRectTargetGbn + "'"
			end if
		end if

        if (FRectMakerid<>"") then
		    sqlAdd = sqlAdd + " and s.lastmakerid='" + FRectMakerid + "'"
		end If

		if FRectDispCate<>"" Then
			sqlStr = sqlStr + " and s.itemgubun = '10' "
		    if LEN(FRectDispCate)>3 then
		         sqlAdd = sqlAdd + " and i.dispcate1='"&LEFT(FRectDispCate,3)&"'"
		    end if
			sqlAdd = sqlAdd + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end If

		if (FRectShowUpbae <> "") then
			sqlAdd = sqlAdd & " 	and i.mwdiv = 'U' "
		end if

        if (isPagingReq) then
            sqlStr = " SELECT count(*) as CNT, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
            sqlStr = sqlStr & " FROM "
    		sqlStr = sqlStr & " 	db_summary.dbo.tbl_monthly_accumulated_logisstock_summary s "
    		sqlStr = sqlStr & "	left join db_partner.dbo.tbl_partner p "
    		sqlStr = sqlStr & "	on "
    		sqlStr = sqlStr & "		s.lastmakerid = p.id "
    		sqlStr = sqlStr & "	left join db_partner.dbo.tbl_partner_comm_code pc "
    		sqlStr = sqlStr & "	on "
    		sqlStr = sqlStr & "		1 = 1 "
    		sqlStr = sqlStr & "		and pc.pcomm_group = 'purchasetype' "
    		sqlStr = sqlStr & "		and pc.pcomm_cd = p.purchasetype "

		    sqlStr = sqlStr & "		left join [db_item].[dbo].tbl_item i "
            sqlStr = sqlStr & "		on s.yyyymm='"&FRectYYYYMM&"'"
            sqlStr = sqlStr & "		and s.itemgubun='10' "
            sqlStr = sqlStr & "		and s.itemid=i.itemid "

			sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on s.yyyymm='"&FRectYYYYMM&"' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"

            sqlStr = sqlStr & "		left join [db_shop].[dbo].tbl_shop_item si "
            sqlStr = sqlStr & "		on s.yyyymm='"&FRectYYYYMM&"'"
            sqlStr = sqlStr & "		and s.itemgubun<>'10' "
            sqlStr = sqlStr & "		and s.itemgubun=si.itemgubun"
            sqlStr = sqlStr & "		and s.itemid=si.shopitemid "
            sqlStr = sqlStr & "		and s.itemoption=si.itemoption "

            sqlStr = sqlStr & " WHERE "
    		sqlStr = sqlStr & " 	1 = 1 "
    		sqlStr = sqlStr & " 	and s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
    		sqlStr = sqlStr & "		and not (s.itemgubun='10' and s.itemid  in (0,11406,6400))"
    		''sqlStr = sqlStr & " 	and IsNULL(s.lastmwdiv,'Z') <> 'W' "

    		sqlStr = sqlStr & sqlAdd
            sqlStr = sqlStr & " 	and " + CStr(colName) + " <> 0 "		'// 재고없는 것도 표시, 2021-05-11, skyer9

    		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
    		rsget.Close

    		'지정페이지가 전체 페이지보다 클 때 함수종료
    		if Cint(FCurrPage)>Cint(FTotalPage) then
    			FResultCount = 0
    			exit sub
    		end if
        end if

		sqlStr = " SELECT top "&FPageSize*FCurrPage&" "
		sqlStr = sqlStr & " 	s.targetGbn, s.itemgubun, IsNULL(s.lastmwdiv,'Z') as mwdiv, s.lastmakerid as makerid, IsNull(pc.pcomm_name, '') as purchasetypeStr, "
		sqlStr = sqlStr & " 	SUM(" + CStr(colName) + ") as totStockNo, "

		if (FRectMakerid <> "") or (FRectShowItemList<>"") then
			sqlStr = sqlStr & " s.itemgubun,s.itemid,s.itemoption, isNULL(i.itemname,si.shopitemname) as itemname, IsNULL(o.optionname,'') as itemoptionname, IsNull(s.lastIpgoDate, '') as lastIpgoDate, " + CStr(valPrice) + " as buyPrice, "
		else
			sqlStr = sqlStr & " '" + CStr(FRectMonthGubun) + "' as lastIpgoDate, "
		end if

		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun1, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun2, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun3, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun4, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun5, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN IsNull(s.lastIpgoDate, '') = '' "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun6, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun7, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) and (DateDiff(m, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun8, "

		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun11, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun12, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun13, "
		sqlStr = sqlStr & " 	SUM(CASE "
		sqlStr = sqlStr & " 		WHEN (DateDiff(yyyy, (s.lastIpgoDate + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3) "
		sqlStr = sqlStr & " 		THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 END) as MonthGubun14, "

		sqlStr = sqlStr & " 	SUM(" + CStr(colName) + " * " + CStr(valPrice) + ") as MonthGubunSUM "

		sqlStr = sqlStr & " FROM "
		sqlStr = sqlStr & " 	db_summary.dbo.tbl_monthly_accumulated_logisstock_summary s "
		sqlStr = sqlStr & "	left join db_partner.dbo.tbl_partner p "
		sqlStr = sqlStr & "	on "
		sqlStr = sqlStr & "		s.lastmakerid = p.id "
		sqlStr = sqlStr & "	left join db_partner.dbo.tbl_partner_comm_code pc "
		sqlStr = sqlStr & "	on "
		sqlStr = sqlStr & "		1 = 1 "
		sqlStr = sqlStr & "		and pc.pcomm_group = 'purchasetype' "
		sqlStr = sqlStr & "		and pc.pcomm_cd = p.purchasetype "

        if (FRectVatYn<>"") or (FRectMakerid <> "") or (FRectShowItemList<>"") or (FRectShopSuplyPrice = "Y") Or (FRectShowUpbae <> "") then
            sqlStr = sqlStr & "		left join [db_item].[dbo].tbl_item i "
            sqlStr = sqlStr & "		on s.yyyymm='"&FRectYYYYMM&"'"
            sqlStr = sqlStr & "		and s.itemgubun='10' "
            sqlStr = sqlStr & "		and s.itemid=i.itemid "

			sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on s.yyyymm='"&FRectYYYYMM&"' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"

            sqlStr = sqlStr & "		left join [db_shop].[dbo].tbl_shop_item si "
            sqlStr = sqlStr & "		on s.yyyymm='"&FRectYYYYMM&"'"
            sqlStr = sqlStr & "		and s.itemgubun<>'10' "
            sqlStr = sqlStr & "		and s.itemgubun=si.itemgubun"
            sqlStr = sqlStr & "		and s.itemid=si.shopitemid "
            sqlStr = sqlStr & "		and s.itemoption=si.itemoption "
        end if

		sqlStr = sqlStr & " WHERE "
		sqlStr = sqlStr & " 	1 = 1 "
		sqlStr = sqlStr & " 	and s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		sqlStr = sqlStr & "		and not (s.itemgubun='10' and s.itemid  in (0,11406,6400))"
	''sqlStr = sqlStr & " 	and IsNULL(s.lastmwdiv,'Z') <> 'W' "

		sqlStr = sqlStr & sqlAdd

		sqlStr = sqlStr & " 	and " + CStr(colName) + " <> 0 "		'// 재고없는 것도 표시, 2021-05-11, skyer9
		sqlStr = sqlStr & " GROUP BY "
		sqlStr = sqlStr & " 	s.targetGbn, s.itemgubun, s.lastMWDiv, s.lastmakerid, IsNull(pc.pcomm_name, '') "

		if (FRectMakerid <> "")or(FRectShowItemList<>"") then
			sqlStr = sqlStr & " ,s.itemgubun,s.itemid,s.itemoption, isNULL(i.itemname,si.shopitemname), IsNULL(o.optionname,''), IsNull(s.lastIpgoDate, ''), " + CStr(valPrice) + " "
		end if

        if ((FRectMakerid <> "")or(FRectShowItemList<>"")) then
            if (FRectOrdTp="S") then
       		    sqlStr = sqlStr & " ORDER BY "
        		sqlStr = sqlStr & " 	totStockNo desc, s.targetGbn DESC, s.itemgubun, IsNULL(s.lastmwdiv,'Z'),  s.lastmakerid "
        		sqlStr = sqlStr & " ,s.itemid,s.itemoption "
            else
        		sqlStr = sqlStr & " ORDER BY "
        		sqlStr = sqlStr & " 	s.itemgubun "
        		sqlStr = sqlStr & " ,s.itemid,s.itemoption "
            end if
        else
            sqlStr = sqlStr & " ORDER BY "
        	sqlStr = sqlStr & " 	totStockNo desc, s.targetGbn DESC, s.itemgubun"
        end if
		''response.write sqlStr
		''response.end

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
				set FItemList(i) = new CStockOverValueSum

			    FItemList(i).FtargetGbn     = rsget("targetGbn")
				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).FpurchasetypeStr     = rsget("purchasetypeStr")



				FItemList(i).Fmakerid     	= rsget("makerid")
				FItemList(i).FlastIpgoDate  = rsget("lastIpgoDate")
				FItemList(i).FtotStockNo    = rsget("totStockNo")

				if (FRectMakerid <> "")or(FRectShowItemList<>"") then
					FItemList(i).Fitemid     	= rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname     	= rsget("itemname")
					FItemList(i).Fitemoptionname = rsget("itemoptionname")
					FItemList(i).FbuyPrice 		= rsget("buyPrice")
				end if

				FItemList(i).FTotBuySum1 	= rsget("MonthGubun1")
				FItemList(i).FTotBuySum2 	= rsget("MonthGubun2")
				FItemList(i).FTotBuySum3 	= rsget("MonthGubun3")
				FItemList(i).FTotBuySum4 	= rsget("MonthGubun4")
				FItemList(i).FTotBuySum5 	= rsget("MonthGubun5")
				FItemList(i).FTotBuySum6 	= rsget("MonthGubun6")
				FItemList(i).FTotBuySum7 	= rsget("MonthGubun7")
				FItemList(i).FTotBuySum8 	= rsget("MonthGubun8")

				FItemList(i).FTotBuySum11 	= rsget("MonthGubun11")
				FItemList(i).FTotBuySum12 	= rsget("MonthGubun12")
				FItemList(i).FTotBuySum13 	= rsget("MonthGubun13")
				FItemList(i).FTotBuySum14 	= rsget("MonthGubun14")

				FItemList(i).FTotBuySum 	= rsget("MonthGubunSUM")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end Sub

	'// 재고평가충당금 : 매장
    public Sub GetJeagoOverValueSum_Shop()
    	dim sqlStr
		dim colName, valPrice, colNameIpgoDate, valPriceFieldName

		valPriceFieldName = " s.LstBuyCash "
		if (FRectPriceGubun = "V") then
			valPriceFieldName = " IsNull(s.avgShopIpgoPrice, s.LstBuyCash) "
		end if

		''G02799 : 아이띵소 그룹코드
		''B011 : 위탁판매, B012 : 업체위탁, B013 : 출고위탁, B021 : 오프매입, B022 : 매장매입, B023 : 가맹점매입, B031 : 출고매입, B032 : 센터매입

		if (FRectGubun = "sys") then
			colName = "s.sysstockno"
		else
			colName = "(s.realstockno - s.errbaditemno)"			'// 불량재고 포함
		end if

        if (FRectLastIpgoGBN="S") then
            colNameIpgoDate = "s.lastIpgoDate"
		elseif (FRectLastIpgoGBN="M") then
		   '' rw "TTT"
			''colNameIpgoDate = "(case when IsNull(s.LstCenterMwDiv, 'W') = 'M' then s.lastIpgodateLogics else s.lastIpgoDate end)"
			colNameIpgoDate = "(case when IsNull(s.LstCenterMwDiv, 'W') = 'M' then d.lastipgoDateByMW else s.lastIpgoDate end)"
        else
            colNameIpgoDate = "s.lastIpgodateLogics"
        end if

		if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(isNULL(isNULL(s.lstvatinclude,i.vatinclude),si.vatinclude), 'Y') = 'Y' then (" + CStr(valPriceFieldName) + "*10/11) else " + CStr(valPriceFieldName) + " end)"
			if (FRectPriceGubun = "V") then
				valPrice = "(case when IsNull(isNULL(isNULL(s.lstvatinclude,i.vatinclude),si.vatinclude), 'Y') = 'Y' then Round((" + CStr(valPriceFieldName) + "*10/11), 0) else " + CStr(valPriceFieldName) + " end)"
			end if
		else
			valPrice = valPriceFieldName
		end if

		sqlStr = " SELECT "
'		sqlStr = sqlStr & " 	(CASE "
'		sqlStr = sqlStr & " 		WHEN ((IsNull(p.groupid, '') = 'G02799' and s.LstComm_cd='B013') or IsNull(sp.sellbizCD, '') = '0000000301') THEN 'IT' "
'		sqlStr = sqlStr & " 		ELSE 'OF' "
'		sqlStr = sqlStr & " 	END) as targetGbn, "
        sqlStr = sqlStr & " 	isNULL(s.targetGbn,'OF') as targetGbn, "
		sqlStr = sqlStr & " 	s.itemgubun "
		sqlStr = sqlStr & " 	, IsNull(s.LstComm_cd, 'Z') as mwdiv "
		sqlStr = sqlStr & " 	, SUM("+CStr(colName)+") as totStockNo"
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun1 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun2 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun3 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun4 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun5 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN IsNull("&colNameIpgoDate&", '') = '' THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun6 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun7 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE  "
		sqlStr = sqlStr & " 		WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 		ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun8 "

		'// 연도별
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE WHEN DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") ELSE 0 END), 0) AS MonthGubun11 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE WHEN DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") ELSE 0 END), 0) AS MonthGubun12 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE WHEN DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") ELSE 0 END), 0) AS MonthGubun13 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE WHEN DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") ELSE 0 END), 0) AS MonthGubun14 "

		sqlStr = sqlStr & " 	,IsNull(SUM(" + CStr(colName) + " * " + CStr(valPrice) + "), 0) AS MonthGubunSUM "
		sqlStr = sqlStr & " FROM db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s with (nolock)"
		if (FRectLastIpgoGBN="M") then
		    sqlStr = sqlStr & " JOin db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail d with (nolock)"
        	sqlStr = sqlStr & " 	on s.yyyymm=d.yyyymm"
			sqlStr = sqlStr & " 	and d.stockPlace in ('A','E','F','L','N','O','S') "
        	sqlStr = sqlStr & " 	and s.shopid=d.shopid"
        	sqlStr = sqlStr & " 	and s.itemgubun=d.itemgubun"
        	sqlStr = sqlStr & " 	and s.itemid=d.itemid"
        	sqlStr = sqlStr & " 	and s.itemoption=d.itemoption"
		end if
		sqlStr = sqlStr & "	left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr & "		on "
		sqlStr = sqlStr & " 		1 = 1 "
		sqlStr = sqlStr & "			and s.itemgubun='10' "
		sqlStr = sqlStr & "			and s.itemid=i.itemid "
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_item si with (nolock)"
		sqlStr = sqlStr & " 	on "
		sqlStr = sqlStr & " 		1 = 1 "
		sqlStr = sqlStr & "			and s.itemgubun<>'10' "
		sqlStr = sqlStr & " 		and s.itemgubun = si.itemgubun "
		sqlStr = sqlStr & " 		and s.itemid = si.shopitemid "
		sqlStr = sqlStr & " 		and s.itemoption = si.itemoption "
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p with (nolock)"
		sqlStr = sqlStr & " 	on s.LstMakerid=p.id "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner sp with (nolock)"
		sqlStr = sqlStr & " 	ON s.shopid = sp.id "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " 	AND s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		sqlStr = sqlStr & " 	AND NOT (s.itemgubun = '10' AND s.itemid IN (0,11406,6400)) "
		sqlStr = sqlStr & " 	and s.itemid <> 0 "
		sqlStr = sqlStr & " 	AND " + CStr(colName) + " <> 0 "		'// 재고없는 것도 표시, 2021-05-11, skyer9

		if (FRectMwDiv <> "") then
			if (FRectMwDiv = "M") then
				sqlStr = sqlStr & " 	and s.LstComm_cd in ('B021','B022','B031','B032','B013') "
				sqlStr = sqlStr & " 	and (isNULL(s.targetGbn,'OF') not in ('ET','EG'))" ''NOT ((s.targetGbn='ET') and (s.LstComm_cd='B013'))"
			elseif (FRectMwDiv = "W") then
				sqlStr = sqlStr & " 	and s.LstComm_cd in ('B011','B012') "
			elseif (FRectMwDiv = "Z") then
				sqlStr = sqlStr & " 	and IsNULL(s.LstComm_cd,'Z') not in ('B021','B022','B031','B032','B011','B012','B013') "
			else
				sqlStr = sqlStr & " 	and IsNULL(s.LstComm_cd,'Z') = '" + CStr(FRectMwDiv) + "' "
			end if
		end if

		if (FRectItemGubun <> "") then
			sqlStr = sqlStr & " 	and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
		end if

		'구매유형
		if FRectPurchaseType <> "" then
			if FRectPurchasetype="101" then
				sqlStr = sqlStr + " and p.purchasetype<>1 "
			elseif FRectPurchasetype="102" then
				sqlStr = sqlStr + " and p.purchasetype in (3,5,6) "
			else
				sqlStr = sqlStr + " and p.purchasetype = " & FRectPurchaseType & " "
			end if
		end if

        if (FRectVatYn<>"") then
		    sqlStr = sqlStr + " and IsNull(isNULL(isNULL(s.lstvatinclude,i.vatinclude),si.vatinclude), 'Y')='" + FRectVatYn + "'"
		end if

        if (FRectTargetGbn<>"") then
		    ''sqlStr = sqlStr + " and (CASE WHEN ((IsNull(p.groupid, '') = 'G02799' and s.LstComm_cd='B013') or IsNull(sp.sellbizCD, '') = '0000000301') THEN 'IT' ELSE 'OF' END) = '" + CStr(FRectTargetGbn) + "' "
		    'sqlStr = sqlStr & " and isNULL(s.targetGbn,'OF')= '" + CStr(FRectTargetGbn) + "' "

			if (FRectTargetGbn = "3X") then
				sqlStr = sqlStr + " and isNULL(s.targetGbn,'OF') not in ('ET', 'EG') "
			else
				sqlStr = sqlStr + " and isNULL(s.targetGbn,'OF')='" + FRectTargetGbn + "'"
			end if
		end if

		if (FRectShopID <> "") then
			sqlStr = sqlStr & " 	and s.shopid = '" + CStr(FRectShopID) + "' "
		end if

        IF (FRectetcjungsantype<>"") then
    	    if (FRectetcjungsantype="41") then
    	        sqlStr = sqlStr + " and sp.etcjungsantype in ('4','1')"
    	    else
    	        sqlStr = sqlStr + " and sp.etcjungsantype='"&FRectetcjungsantype&"'"
    	    end if
    	end if

		sqlStr = sqlStr & " GROUP BY "
'		sqlStr = sqlStr & " 	(CASE "
'		sqlStr = sqlStr & " 		WHEN ((IsNull(p.groupid, '') = 'G02799' and s.LstComm_cd='B013') or IsNull(sp.sellbizCD, '') = '0000000301') THEN 'IT' "
'		sqlStr = sqlStr & " 		ELSE 'OF' "
'		sqlStr = sqlStr & " 	END), "
        sqlStr = sqlStr & " isNULL(s.targetGbn,'OF'), "
		sqlStr = sqlStr & " 	s.itemgubun "
		sqlStr = sqlStr & " 	, IsNull(s.LstComm_cd, 'Z') "
		'sqlStr = sqlStr & " ORDER BY targetGbn DESC, s.itemgubun asc, IsNull(s.LstComm_cd, 'Z') asc"
		sqlStr = sqlStr & " ORDER BY s.itemgubun asc, IsNull(s.LstComm_cd, 'Z') asc"

		''response.write sqlStr
		''response.end
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
			do until rsget.eof
				set FItemList(i) = new CStockOverValueSum

			    FItemList(i).FtargetGbn     = rsget("targetGbn")
				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")
				FItemList(i).FTotBuySum1 	= rsget("MonthGubun1")
				FItemList(i).FTotBuySum2 	= rsget("MonthGubun2")
				FItemList(i).FTotBuySum3 	= rsget("MonthGubun3")
				FItemList(i).FTotBuySum4 	= rsget("MonthGubun4")
				FItemList(i).FTotBuySum5 	= rsget("MonthGubun5")
				FItemList(i).FTotBuySum6 	= rsget("MonthGubun6")
				FItemList(i).FTotBuySum7 	= rsget("MonthGubun7")
				FItemList(i).FTotBuySum8 	= rsget("MonthGubun8")
				FItemList(i).FTotBuySum11 	= rsget("MonthGubun11")
				FItemList(i).FTotBuySum12 	= rsget("MonthGubun12")
				FItemList(i).FTotBuySum13 	= rsget("MonthGubun13")
				FItemList(i).FTotBuySum14 	= rsget("MonthGubun14")
				FItemList(i).FTotBuySum 	= rsget("MonthGubunSUM")
                FItemList(i).FtotStockNo    = rsget("totStockNo")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

	'// 재고평가충당금 : 매장 : 상세
    public Sub GetJeagoOverValueDetailSum_Shop()
    	dim sqlStr, sqlAdd
		dim colName, valPrice, colNameIpgoDate, valPriceFieldName
        Dim isPagingReq : isPagingReq=False

		valPriceFieldName = " s.LstBuyCash "
		if (FRectPriceGubun = "V") then
			valPriceFieldName = " IsNull(s.avgShopIpgoPrice, s.LstBuyCash) "
		end if

        if (FRectMakerid <> "")or(FRectShowItemList<>"") then
            isPagingReq = TRUE
        end if

		''G02799 : 아이띵소 그룹코드
		''B011 : 위탁판매, B012 : 업체위탁, B013 : 출고위탁, B021 : 오프매입, B022 : 매장매입, B023 : 가맹점매입, B031 : 출고매입, B032 : 센터매입

		if (FRectGubun = "sys") then
			colName = "s.sysstockno"
		else
			colName = "(s.realstockno - s.errbaditemno)"			'// 불량재고 포함
		end if

        if (FRectLastIpgoGBN="S") then
            colNameIpgoDate = "s.lastIpgoDate"
		elseif (FRectLastIpgoGBN="M") then
			''colNameIpgoDate = "(case when IsNull(s.LstCenterMwDiv, 'W') = 'M' then s.lastIpgodateLogics else s.lastIpgoDate end)"
			colNameIpgoDate = "(case when IsNull(s.LstCenterMwDiv, 'W') = 'M' then d.lastipgoDateByMW else s.lastIpgoDate end)"
        else
            colNameIpgoDate = "s.lastIpgodateLogics"
        end if

		if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			''valPrice = "(case when IsNull(isNULL(isNULL(s.lstvatinclude,i.vatinclude),si.vatinclude), 'Y') = 'Y' then (" + CStr(valPriceFieldName) + "*10/11) else " + CStr(valPriceFieldName) + " end)"
			valPrice = "(case when isNULL(s.lstvatinclude, 'Y') = 'Y' then (" + CStr(valPriceFieldName) + "*10/11) else " + CStr(valPriceFieldName) + " end)"
			if (FRectPriceGubun = "V") then
				''valPrice = "(case when IsNull(isNULL(isNULL(s.lstvatinclude,i.vatinclude),si.vatinclude), 'Y') = 'Y' then Round((" + CStr(valPriceFieldName) + "*10/11), 0) else " + CStr(valPriceFieldName) + " end)"
				valPrice = "(case when isNULL(s.lstvatinclude, 'Y') = 'Y' then Round((" + CStr(valPriceFieldName) + "*10/11), 0) else " + CStr(valPriceFieldName) + " end)"
			end if
		else
			valPrice = valPriceFieldName
		end if

        sqlAdd = ""
        if (FRectMwDiv <> "") then
			if (FRectMwDiv = "M") then
				sqlAdd = sqlAdd & " 	and s.LstComm_cd in ('B021','B022','B031','B032','B013') "
				sqlAdd = sqlAdd & " 	and (isNULL(s.targetGbn,'OF') not in ('ET','EG')) " ''NOT ((s.targetGbn='ET') and (s.LstComm_cd='B013'))" ''
			elseif (FRectMwDiv = "W") then
				sqlAdd = sqlAdd & " 	and s.LstComm_cd in ('B011','B012') "
			elseif (FRectMwDiv = "Z") then
				sqlAdd = sqlAdd & " 	and IsNULL(s.LstComm_cd,'Z') not in ('B021','B022','B031','B032','B011','B012','B013') "
			else
				sqlAdd = sqlAdd & " 	and IsNULL(s.LstComm_cd,'Z') = '" + CStr(FRectMwDiv) + "' "
			end if
		end if

		if (FRectItemGubun <> "") then
			sqlAdd = sqlAdd & " 	and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
		end if

		'구매유형
		if FRectPurchaseType <> "" then
			if FRectPurchasetype="101" then
				sqlAdd = sqlAdd + " and p.purchasetype<>1 "
			elseif FRectPurchasetype="102" then
				sqlAdd = sqlAdd + " and p.purchasetype in (3,5,6) "
			else
				sqlAdd = sqlAdd + " and p.purchasetype = " & FRectPurchaseType & " "
			end if
		end if

        if (FRectVatYn<>"") then
		    sqlAdd = sqlAdd + " and IsNull(isNULL(isNULL(s.lstvatinclude,i.vatinclude),si.vatinclude), 'Y')='" + FRectVatYn + "'"
		end if

        if (FRectTargetGbn<>"") then
		    'sqlStr = sqlStr + " and (CASE WHEN ((IsNull(p.groupid, '') = 'G02799' and s.LstComm_cd='B013') or IsNull(sp.sellbizCD, '') = '0000000301') THEN 'IT' ELSE 'OF' END) = '" + CStr(FRectTargetGbn) + "' "
			'sqlAdd = sqlAdd & " and isNULL(s.targetGbn,'OF')= '" + CStr(FRectTargetGbn) + "' "

			if (FRectTargetGbn = "3X") then
				sqlAdd = sqlAdd + " and isNULL(s.targetGbn,'OF') not in ('ET', 'EG') "
			else
				sqlAdd = sqlAdd + " and isNULL(s.targetGbn,'OF')='" + FRectTargetGbn + "'"
			end if
		end if

		if (FRectShopID <> "") then
			sqlAdd = sqlAdd & " 	and s.shopid = '" + CStr(FRectShopID) + "' "
		end if

		if (FRectMonthGubun = "1") then
			'// 1-3개월
			sqlAdd = sqlAdd & " 	and DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 "
		elseif (FRectMonthGubun = "2") then
			'// 4-6개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) "
		elseif (FRectMonthGubun = "3") then
			'// 7-12개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) "
		elseif (FRectMonthGubun = "4") then
			'// 13-24개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		elseif (FRectMonthGubun = "5") then
			'// 25개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) "
		elseif (FRectMonthGubun = "6") then
			'// 미지정
			sqlAdd = sqlAdd & " 	and IsNull("&colNameIpgoDate&", '') = '' "
		elseif (FRectMonthGubun = "7") then
			'// 13-18개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) "
		elseif (FRectMonthGubun = "8") then
			'// 19-24개월
			sqlAdd = sqlAdd & " 	and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) and (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) "
		elseif (FRectMonthGubun = "11") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0) "
		elseif (FRectMonthGubun = "12") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1) "
		elseif (FRectMonthGubun = "13") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2) "
		elseif (FRectMonthGubun = "14") then
			'// 올해
			sqlAdd = sqlAdd & " 	and (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3) "
		end if

        if (FRectMakerid<>"") then
		    sqlAdd = sqlAdd + " and s.LstMakerid='" + FRectMakerid + "'"
		end if

        IF (FRectetcjungsantype<>"") then
    	    if (FRectetcjungsantype="41") then
    	        sqlAdd = sqlAdd + " and sp.etcjungsantype in ('4','1')"
    	    else
    	        sqlAdd = sqlAdd + " and sp.etcjungsantype='"&FRectetcjungsantype&"'"
    	    end if
    	end if

        if (isPagingReq) then
            sqlStr = " SELECT count(*) as CNT, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
            sqlStr = sqlStr & " FROM "
            sqlStr = sqlStr & " 	db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s "
    		sqlStr = sqlStr & " 	LEFT JOIN [db_item].[dbo].tbl_item i "
    		sqlStr = sqlStr & " 	ON "
    		sqlStr = sqlStr & " 		1 = 1 "
    		sqlStr = sqlStr & " 		AND s.itemgubun = '10' "
    		sqlStr = sqlStr & " 		AND s.itemid = i.itemid "

    		sqlStr = sqlStr + "     LEFT JOIN [db_item].[dbo].tbl_item_option o on s.yyyymm='"&FRectYYYYMM&"' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"

    		sqlStr = sqlStr & " 	LEFT JOIN db_shop.dbo.tbl_shop_item si "
    		sqlStr = sqlStr & " 	ON "
    		sqlStr = sqlStr & " 		1 = 1 "
    		sqlStr = sqlStr & " 		AND s.itemgubun <> '10' "
    		sqlStr = sqlStr & " 		AND s.itemgubun = si.itemgubun "
    		sqlStr = sqlStr & " 		AND s.itemid = si.shopitemid "
    		sqlStr = sqlStr & " 		AND s.itemoption = si.itemoption "
    		sqlStr = sqlStr & " 	LEFT JOIN db_partner.dbo.tbl_partner p "
    		sqlStr = sqlStr & " 	ON "
    		sqlStr = sqlStr & " 		s.LstMakerid = p.id "
    		sqlStr = sqlStr & " 	LEFT JOIN db_partner.dbo.tbl_partner sp "
    		sqlStr = sqlStr & " 	ON "
    		sqlStr = sqlStr & " 		s.shopid = sp.id "
    		sqlStr = sqlStr & " WHERE 1 = 1 "
    		sqlStr = sqlStr & " 	AND s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
    		sqlStr = sqlStr & " 	AND NOT (s.itemgubun = '10' AND s.itemid IN (0,11406,6400)) "
    		sqlStr = sqlStr & " 	AND s.itemid <> 0 "
    		sqlStr = sqlStr & " 	AND " + CStr(colName) + " <> 0 "
    		sqlStr = sqlStr & sqlAdd
            rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
    		rsget.Close

    		'지정페이지가 전체 페이지보다 클 때 함수종료
    		if Cint(FCurrPage)>Cint(FTotalPage) then
    			FResultCount = 0
    			exit sub
    		end if
        end if

		sqlStr = " SELECT top "&FPageSize*FCurrPage&" "
		sqlStr = sqlStr & " 	isNULL(s.targetGbn,'OF') as targetGbn "
		sqlStr = sqlStr & " 	,s.itemgubun "
		sqlStr = sqlStr & " 	,IsNull(s.LstComm_cd, 'Z') AS mwdiv "
		sqlStr = sqlStr & " 	,s.LstMakerid as makerid "
		sqlStr = sqlStr & " 	,SUM(" + CStr(colName) + ") as totStockNo "

		if (FRectMakerid <> "")or(FRectShowItemList<>"") then
			sqlStr = sqlStr & " ,s.shopid,s.itemgubun,s.itemid,s.itemoption, isNULL(i.itemname,si.shopitemname) as itemname, IsNULL(o.optionname,'') as itemoptionname, IsNull("&colNameIpgoDate&", '') as lastIpgoDate, " + CStr(valPrice) + " as buyPrice, s.LstComm_cd "
		else
			sqlStr = sqlStr & " ,'" + CStr(FRectMonthGubun) + "' as lastIpgoDate "
		end if

		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' then 0 "
		sqlStr = sqlStr & " 				WHEN DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 2 THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun1 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' then 0 "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 2) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 5) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun2 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' then 0 "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 5) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 11) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun3 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' then 0 "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun4 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' then 0 "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun5 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun6 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' then 0 "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 11) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 17) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun7 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' then 0 "
		sqlStr = sqlStr & " 				WHEN (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') > 17) AND (DateDiff(m, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') <= 23) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun8 "

		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' then 0 "
		sqlStr = sqlStr & " 				WHEN (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 0) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun11 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' then 0 "
		sqlStr = sqlStr & " 				WHEN (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 1) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun12 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' then 0 "
		sqlStr = sqlStr & " 				WHEN (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') = 2) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun13 "
		sqlStr = sqlStr & " 	,IsNull(SUM(CASE "
		sqlStr = sqlStr & " 				WHEN IsNull("&colNameIpgoDate&", '') = '' then 0 "
		sqlStr = sqlStr & " 				WHEN (DateDiff(yyyy, ("&colNameIpgoDate&" + '-01'), '" + CStr(FRectYYYYMM) + "' + '-01') >= 3) THEN (" + CStr(colName) + " * " + CStr(valPrice) + ") "
		sqlStr = sqlStr & " 				ELSE 0 "
		sqlStr = sqlStr & " 	END), 0) AS MonthGubun14 "

		sqlStr = sqlStr & " 	,IsNull(SUM(" + CStr(colName) + " * " + CStr(valPrice) + "), 0) AS MonthGubunSUM "
		sqlStr = sqlStr & " FROM "
		sqlStr = sqlStr & " 	db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s "
		if (FRectLastIpgoGBN="M") then
		    sqlStr = sqlStr & " 	JOin db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail d"
        	sqlStr = sqlStr & " 	on s.yyyymm=d.yyyymm"
        	sqlStr = sqlStr & " 	and d.stockPlace in ('A','E','F','L','N','O','S') "
			sqlStr = sqlStr & " 	and s.shopid=d.shopid"
        	sqlStr = sqlStr & " 	and s.itemgubun=d.itemgubun"
        	sqlStr = sqlStr & " 	and s.itemid=d.itemid"
        	sqlStr = sqlStr & " 	and s.itemoption=d.itemoption"
		end if

		if (FRectMakerid <> "")or(FRectShowItemList<>"") then
			sqlStr = sqlStr & " 	LEFT JOIN [db_item].[dbo].tbl_item i "
			sqlStr = sqlStr & " 	ON "
			sqlStr = sqlStr & " 		1 = 1 "
			sqlStr = sqlStr & " 		AND s.itemgubun = '10' "
			sqlStr = sqlStr & " 		AND s.itemid = i.itemid "
			sqlStr = sqlStr + "     LEFT JOIN [db_item].[dbo].tbl_item_option o on s.yyyymm='"&FRectYYYYMM&"' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"
			sqlStr = sqlStr & " 	LEFT JOIN db_shop.dbo.tbl_shop_item si "
			sqlStr = sqlStr & " 	ON "
			sqlStr = sqlStr & " 		1 = 1 "
			sqlStr = sqlStr & " 		AND s.itemgubun <> '10' "
			sqlStr = sqlStr & " 		AND s.itemgubun = si.itemgubun "
			sqlStr = sqlStr & " 		AND s.itemid = si.shopitemid "
			sqlStr = sqlStr & " 		AND s.itemoption = si.itemoption "
		end if

		sqlStr = sqlStr & " 	LEFT JOIN db_partner.dbo.tbl_partner p "
		sqlStr = sqlStr & " 	ON "
		sqlStr = sqlStr & " 		s.LstMakerid = p.id "
		sqlStr = sqlStr & " 	LEFT JOIN db_partner.dbo.tbl_partner sp "
		sqlStr = sqlStr & " 	ON "
		sqlStr = sqlStr & " 		s.shopid = sp.id "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " 	AND s.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		sqlStr = sqlStr & " 	AND NOT (s.itemgubun = '10' AND s.itemid IN (0,11406,6400)) "
		sqlStr = sqlStr & " 	AND s.itemid <> 0 "
		sqlStr = sqlStr & " 	AND " + CStr(colName) + " <> 0 "

		sqlStr = sqlStr & sqlAdd

		sqlStr = sqlStr & " GROUP BY isNULL(s.targetGbn,'OF') "
		sqlStr = sqlStr & " 	,s.itemgubun "
		sqlStr = sqlStr & " 	,IsNull(s.LstComm_cd, 'Z') "
		sqlStr = sqlStr & " 	,s.LstMakerid "

		if (FRectMakerid <> "")or(FRectShowItemList<>"") then
			sqlStr = sqlStr & " ,s.shopid,s.itemgubun,s.itemid,s.itemoption, isNULL(i.itemname,si.shopitemname), IsNULL(o.optionname,''), IsNull("&colNameIpgoDate&", ''), " + CStr(valPrice) + ", s.LstComm_cd "
		end if

        if ((FRectMakerid <> "")or(FRectShowItemList<>"")) then
            if (FRectOrdTp="S") then
                sqlStr = sqlStr & " ORDER BY totStockNo desc, targetGbn DESC "
        		sqlStr = sqlStr & " 	,s.itemgubun "
        		sqlStr = sqlStr & " ,s.shopid,s.itemid,s.itemoption "
            else
        		sqlStr = sqlStr & " ORDER BY s.shopid,s.itemgubun "
        		sqlStr = sqlStr & " ,s.itemid,s.itemoption "
        	end if
        else
            sqlStr = sqlStr & " ORDER BY targetGbn DESC "
    		sqlStr = sqlStr & " 	,s.itemgubun "
    		sqlStr = sqlStr & " 	,IsNull(s.LstComm_cd, 'Z') "
    		sqlStr = sqlStr & " 	,IsNull(SUM(" + CStr(colName) + " * " + CStr(valPrice) + "), 0) desc, s.LstMakerid "
        end if

		''response.write sqlStr
		''response.end
        rsget.pagesize = FPageSize
       rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
'   if (FResultCount>=3000) then  ''서동석 임시
'        rw sqlStr
'        dbget.close() : response.end
'   end if
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		    rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CStockOverValueSum

			    FItemList(i).FtargetGbn     = rsget("targetGbn")
				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).Fmakerid     	= rsget("makerid")
				FItemList(i).FlastIpgoDate  = rsget("lastIpgoDate")
				FItemList(i).FtotStockNo    = rsget("totStockNo")

				if (FRectMakerid <> "")or(FRectShowItemList<>"") then
					FItemList(i).Fshopid     	= rsget("shopid")
					FItemList(i).Fitemid     	= rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname     	= rsget("itemname")
					FItemList(i).Fitemoptionname = rsget("itemoptionname")
					FItemList(i).FbuyPrice 		= rsget("buyPrice")
					FItemList(i).FLstComm_cd	= rsget("LstComm_cd")
				end if

				FItemList(i).FTotBuySum1 	= rsget("MonthGubun1")
				FItemList(i).FTotBuySum2 	= rsget("MonthGubun2")
				FItemList(i).FTotBuySum3 	= rsget("MonthGubun3")
				FItemList(i).FTotBuySum4 	= rsget("MonthGubun4")
				FItemList(i).FTotBuySum5 	= rsget("MonthGubun5")
				FItemList(i).FTotBuySum6 	= rsget("MonthGubun6")
				FItemList(i).FTotBuySum7 	= rsget("MonthGubun7")
				FItemList(i).FTotBuySum8 	= rsget("MonthGubun8")

				FItemList(i).FTotBuySum11 	= rsget("MonthGubun11")
				FItemList(i).FTotBuySum12 	= rsget("MonthGubun12")
				FItemList(i).FTotBuySum13 	= rsget("MonthGubun13")
				FItemList(i).FTotBuySum14 	= rsget("MonthGubun14")

				FItemList(i).FTotBuySum 	= rsget("MonthGubunSUM")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end Sub

    ''물류 오프재고
    function GetOFFMonthlyJeagoSumWithPreMonth()
        dim sqlStr
        Dim stockColNm
        Dim PreYYYYMM

        PreYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)


        sqlStr = ""

        sqlStr = " select A.itemgubun,A.mwdiv,A.totno,A.buysum,isNULL(B.totno,0) as ptotno,isNULL(B.buysum,0) as pbuysum"
        sqlStr = sqlStr & " ,A.totipgono-isNULL(B.totipgono,0) as ipno"
        sqlStr = sqlStr & " ,A.ipgobuysum-isNULL(B.ipgobuysum,0) as ipbuysum"
        sqlStr = sqlStr & " ,A.totlossno-isNULL(B.totlossno,0) as totlossno"
        sqlStr = sqlStr & " ,A.lossbuysum-isNULL(B.lossbuysum,0) as lossbuysum"
        sqlStr = sqlStr & " from ("
        sqlStr = sqlStr &   getComonSubQueryLogistics(FRectYYYYMM,"OF","itemgubun")
        sqlStr = sqlStr & " ) A"
        sqlStr = sqlStr & " left join ("
        sqlStr = sqlStr &   getComonSubQueryLogistics(PreYYYYMM,"OF","itemgubun")
        sqlStr = sqlStr & " ) B"
        sqlStr = sqlStr & " on A.itemgubun=B.itemgubun"
        sqlStr = sqlStr & " and A.mwdiv=B.MwDiv"

        sqlStr = sqlStr & " order by A.itemgubun, A.mwdiv"
''rw sqlStr
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
				set FItemList(i) = new CMonthlyStockSum
				if (FRectITSOnlyOrNot="N") then
				    FItemList(i).FBusiName      = "오프라인"
				elseif (FRectITSOnlyOrNot="O") then
				    FItemList(i).FBusiName      = "아이띵소(OFF)"
				else
				    FItemList(i).FBusiName      = "오프라인+ITS"
			    end if

				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")

				FItemList(i).FTotPreCount   = rsget("ptotno")
            	FItemList(i).FTotPreBuySum  = rsget("pbuysum")
            	FItemList(i).FTotIpCount    = rsget("ipno")
            	FItemList(i).FTotIpBuySum   = rsget("ipbuysum")
            	FItemList(i).FTotLossCount    = rsget("totlossno")
            	FItemList(i).FTotLossBuySum   = rsget("lossbuysum")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    ''물류 일반상품 :: 구버전(201210이전)
    public Sub GetMonthlyJeagoSumNew()
		dim sqlStr
        Dim stockColNm
		if FRectGubun="sys" then
		    stockColNm = "s.totsysstock"
		else
		    stockColNm = "(s.realstock-s.errbaditemno)" ''불량재고도 실재고임.
	    end if

        sqlStr = ""

        IF (FRectOFFReturn2OnStock<>"") then
            sqlStr = sqlStr + " select itemgubun , mwdiv, sum(totno) as totno, sum(buysum) as buysum, sum(sellsum) as sellsum"
            sqlStr = sqlStr + " ,sum(avgIpgoPriceSum) as avgIpgoPriceSum"
            sqlStr = sqlStr + " from ("
            sqlStr = sqlStr + " select '10' as itemgubun, IsNULL(s.lastmwdiv,'Z') as mwdiv "
            sqlStr = sqlStr + "  , sum("&stockColNm&") as totno "
            sqlStr = sqlStr + "  , Sum("&stockColNm&"*si.shopsuplycash) as buysum"
            sqlStr = sqlStr + "  , Sum("&stockColNm&"*si.shopitemprice) as sellsum "
            sqlStr = sqlStr + "  , Sum("&stockColNm&"*IsNULL(s.avgIpgoPrice,si.shopitemprice)) as avgIpgoPriceSum "
            sqlStr = sqlStr + "  from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
            sqlStr = sqlStr + " 	 join db_shop.dbo.tbl_shop_item si"
            sqlStr = sqlStr + "  	 on si.itemgubun='90'"
            sqlStr = sqlStr + " 	 and si.shopitemid=1385"
            sqlStr = sqlStr + " 	 and s.yyyymm='" + FRectYYYYMM + "'"
            sqlStr = sqlStr + " 	 and s.itemgubun=si.itemgubun"
            sqlStr = sqlStr + " 	 and s.itemid=si.shopitemid"
            sqlStr = sqlStr + " 	 and s.itemoption=si.itemoption"
            sqlStr = sqlStr + "  where 1=1"
            if (FRectVatYn<>"") then
    		    sqlStr = sqlStr + " and isNULL(s.lastvatinclude,si.vatinclude)='" + FRectVatYn + "'"
    		end if
            if FRectMwDiv<>"" then
                sqlStr = sqlStr + "  and IsNULL(s.lastmwdiv,'Z')='"&FRectMwDiv&"'"
            end if
            sqlStr = sqlStr + "  group by IsNULL(s.lastmwdiv,'Z')" &VbCRLF
            sqlStr = sqlStr + "  union "
        end if

		sqlStr = sqlStr + " select '10' as itemgubun, IsNULL(s.lastmwdiv,'Z') as mwdiv" ''(CASE WHEN i.mwdiv<>'U' THEN i.mwdiv ELSE IsNULL(s.lastmwdiv,'Z') END)
		sqlStr = sqlStr + " , sum("&stockColNm&") as totno "
		sqlStr = sqlStr + " , Sum("&stockColNm&"*i.orgsuplycash) as buysum"     ''sellsum ,buysum
		sqlStr = sqlStr + " , Sum("&stockColNm&"*i.orgprice) as sellsum"        ''orgprice,orgsuplycash    '''매입인경우 할인하여 떨구는 경우가 많으므로 원 소비자가로.
		sqlStr = sqlStr + " , Sum("&stockColNm&"*IsNULL(s.avgIpgoPrice,i.buycash)) as avgIpgoPriceSum "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
		sqlStr = sqlStr + "     join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + "     on  s.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + "     and s.itemgubun='10'"
		sqlStr = sqlStr + "     and s.itemid=i.itemid"

		sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
		if FRectNewItem<>"" then
			sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
		end if
		if FRectMwDiv<>"" then
			sqlStr = sqlStr + " and IsNULL(s.lastmwdiv,'Z')='" + FRectMwDiv + "'"
		end if
		if (FRectVatYn<>"") then
		    sqlStr = sqlStr + " and isNULL(s.lastvatinclude,i.vatinclude)='" + FRectVatYn + "'"
		end if
		sqlStr = sqlStr + " and i.itemid<>0"
		sqlStr = sqlStr + " and i.itemid<>11406"  	''포장비
		sqlStr = sqlStr + " and i.itemid<>6400"		''배송비
		sqlStr = sqlStr + " and "&stockColNm&">0"
		sqlStr = sqlStr + " group by IsNULL(s.lastmwdiv,'Z')"  ''i.mwdiv
		IF (FRectOFFReturn2OnStock<>"") then
		    sqlStr = sqlStr + ") T group by itemgubun , mwdiv order by mwdiv"
		ELSE
		    sqlStr = sqlStr + " order by mwdiv"
	    ENd IF
'rw 	sqlStr

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
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FItemgubun     = rsget("Itemgubun")
				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")
				FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


	''물류 오프 전용 상품
	public Sub GetOFFMonthlyJeagoSumNew()
		dim sqlStr
        Dim stockColNm
		if FRectGubun="sys" then
		    stockColNm = "s.totsysstock"
		else
		    stockColNm = "(s.realstock-s.errbaditemno)"
	    end if

		sqlStr = "select i.itemgubun, IsNULL(s.lastmwdiv,'Z') as mwdiv"  ''(CASE WHEN i.centermwdiv<>'U' THEN i.centermwdiv ELSE IsNULL(s.lastmwdiv,'Z') END)
		sqlStr = sqlStr + " ,sum("&stockColNm&") as totno "
		sqlStr = sqlStr + " ,Sum("&stockColNm&"*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buysum"
		sqlStr = sqlStr + " , Sum("&stockColNm&"*i.shopitemprice) as sellsum "
		sqlStr = sqlStr + " , Sum("&stockColNm&"*IsNULL(s.avgIpgoPrice,(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END))) as avgIpgoPriceSum "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
		sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + "     on s.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + "     and s.itemgubun<>'10'"
		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
		if (FRectOFFReturn2OnStock<>"") then
		    sqlStr = sqlStr + " and Not (s.itemgubun='90' and s.itemid=1385)"
		end if
		sqlStr = sqlStr + "     left Join db_shop.dbo.tbl_shop_designer d"
		sqlStr = sqlStr + "     on d.shopid='streetshop000'"
		sqlStr = sqlStr + "     and d.makerid=i.makerid"
		sqlStr = sqlStr + " where 1=1"
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
		end if
		if FRectMwDiv<>"" then
			sqlStr = sqlStr + " and IsNULL(s.lastmwdiv,'Z')='" + FRectMwDiv + "'"
		end if
		if (FRectVatYn<>"") then
		    sqlStr = sqlStr + " and isNULL(s.lastvatinclude,i.vatinclude)='" + FRectVatYn + "'"
		end if
		sqlStr = sqlStr + " and i.shopitemid<>0"
		sqlStr = sqlStr + " and "&stockColNm&">0"
		sqlStr = sqlStr + " group by i.itemgubun, IsNULL(s.lastmwdiv,'Z')"
		sqlStr = sqlStr + " order by  mwdiv ,i.itemgubun"

'rw sqlStr
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
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")
				FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).Fitemgubun		= rsget("itemgubun")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	public sub GetMonthlyMoveDiffList
		dim sqlStr
		dim prevYYYYMM

		prevYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)

		sqlStr = " SELECT TOP " & FPageSize & " '" & FRectYYYYMM & "' AS yyyymm "
		sqlStr = sqlStr + " 	,m.itemgubun "
		sqlStr = sqlStr + " 	,m.makerid "
		sqlStr = sqlStr + " 	,m.itemid "
		sqlStr = sqlStr + " 	,m.itemoption "
		sqlStr = sqlStr + " 	,sum(CASE  "
		sqlStr = sqlStr + " 			WHEN m.yyyymm = '" & FRectYYYYMM & "' AND ( "
		sqlStr = sqlStr + " 					(isNULL(m.isMove, 0) <> 0) "
		sqlStr = sqlStr + " 					OR "
		sqlStr = sqlStr + " 					(m.stockPlace <> 'T' AND m.makerid IN ('ithinkso','grandmintfestival','beautifulmintlife') AND m.yyyymm >= '2012-01' AND m.yyyymm < '2012-10') "
		sqlStr = sqlStr + " 					OR "
		sqlStr = sqlStr + " 					(m.stockPlace = 'S' AND i.ipgomwdiv IS NOT NULL AND i.lastcentermwdiv IS NOT NULL AND IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') AND i.ipgomwdiv = 'M') "
		sqlStr = sqlStr + " 				) "
		sqlStr = sqlStr + " 				THEN IsNull(stIpgoNo, 0) "
		sqlStr = sqlStr + " 			ELSE 0 "
		sqlStr = sqlStr + " 			END) "
		sqlStr = sqlStr + " 			+ "
		sqlStr = sqlStr + " 			sum(CASE "
		sqlStr = sqlStr + " 				WHEN m.yyyymm = '" & FRectYYYYMM & "' THEN IsNull(OffChulMoveNo, 0) "
		sqlStr = sqlStr + " 				ELSE 0 "
		sqlStr = sqlStr + " 			END) AS MoveNo "
		sqlStr = sqlStr + " FROM db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail m "
		sqlStr = sqlStr + " LEFT JOIN db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum i ON 1 = 1 "
		sqlStr = sqlStr + " 	AND i.yyyymm = '" & FRectYYYYMM & "' "
		sqlStr = sqlStr + " 	AND m.yyyymm = i.yyyymm "
		sqlStr = sqlStr + " 	AND m.stockPlace = i.stockPlace "
		sqlStr = sqlStr + " 	AND m.shopid = i.shopid "
		sqlStr = sqlStr + " 	AND m.itemgubun = i.itemgubun "
		sqlStr = sqlStr + " 	AND m.itemid = i.itemid "
		sqlStr = sqlStr + " 	AND m.itemoption = i.itemoption "
		sqlStr = sqlStr + " 	AND i.ipgoMWDIV = 'M' "
		sqlStr = sqlStr + " WHERE 1 = 1 "
		sqlStr = sqlStr + " 	AND m.yyyymm >= '" & prevYYYYMM & "' "
		sqlStr = sqlStr + " 	AND m.yyyymm <= '" & FRectYYYYMM & "' "
		sqlStr = sqlStr + " 	AND m.targetGbn NOT IN ('ET','EG') "
		sqlStr = sqlStr + " 	AND m.etcjungsantype IN (1,4) "
		sqlStr = sqlStr + " 	AND NOT (m.lastmwdiv = 'B013' AND m.targetGbn <> 'IT') "
		sqlStr = sqlStr + " 	AND m.lastmwdiv NOT IN ('W','B012','B011') "
		sqlStr = sqlStr + " 	AND m.stockPlace in ('L', 'S') "
		sqlStr = sqlStr + " GROUP BY "
		sqlStr = sqlStr + " 	m.itemgubun "
		sqlStr = sqlStr + " 	,m.makerid "
		sqlStr = sqlStr + " 	,m.itemid "
		sqlStr = sqlStr + " 	,m.itemoption "
		sqlStr = sqlStr + " having sum(CASE  "
		sqlStr = sqlStr + " 			WHEN m.yyyymm = '" & FRectYYYYMM & "' AND ( "
		sqlStr = sqlStr + " 					(isNULL(m.isMove, 0) <> 0) "
		sqlStr = sqlStr + " 					OR "
		sqlStr = sqlStr + " 					(m.stockPlace <> 'T' AND m.makerid IN ('ithinkso','grandmintfestival','beautifulmintlife') AND m.yyyymm >= '2012-01' AND m.yyyymm < '2012-10') "
		sqlStr = sqlStr + " 					OR "
		sqlStr = sqlStr + " 					(m.stockPlace = 'S' AND i.ipgomwdiv IS NOT NULL AND i.lastcentermwdiv IS NOT NULL AND IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') AND i.ipgomwdiv = 'M') "
		sqlStr = sqlStr + " 				) "
		sqlStr = sqlStr + " 				THEN IsNull(stIpgoNo, 0) "
		sqlStr = sqlStr + " 			ELSE 0 "
		sqlStr = sqlStr + " 			END) "
		sqlStr = sqlStr + " 			+ "
		sqlStr = sqlStr + " 			sum(CASE "
		sqlStr = sqlStr + " 				WHEN m.yyyymm = '" & FRectYYYYMM & "' THEN IsNull(OffChulMoveNo, 0) "
		sqlStr = sqlStr + " 				ELSE 0 "
		sqlStr = sqlStr + " 			END) <> 0 "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	m.itemgubun "
		sqlStr = sqlStr + " 	,m.itemid "
		sqlStr = sqlStr + " 	,m.itemoption "

		'rw sqlStr
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
				set FItemList(i) = new CMonthlyStockIpgoItem

				FItemList(i).Fyyyymm 		= rsget("yyyymm")
				FItemList(i).Fitemgubun 	= rsget("itemgubun")
				FItemList(i).Flastmakerid	= rsget("makerid")
				FItemList(i).Fitemid 		= rsget("itemid")
				FItemList(i).Fitemoption 	= rsget("itemoption")
				FItemList(i).FtotItemNo 	= rsget("MoveNo")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Sub

    ''매장 월말 재고액 합계 시스템재고 기준 실사재고표시
    public Sub GetShopMonthlyJeagoSumSysWithReal
        dim sqlStr
	    Dim stockColNm
		if FRectGubun="sys" then
		    stockColNm = "sysstockno"
		else
		    stockColNm = "realstockno"
	    end if

		sqlStr = "select u.userid, u.shopname, IsNULL(d.Comm_cd,'Z') as mwdiv"   ''i.itemgubun,
		sqlStr = sqlStr + " ,sum(s.sysstockno) as totSysno "
		sqlStr = sqlStr + " ,sum(s.realstockno) as totRealno "

		sqlStr = sqlStr + " , Sum(s.sysstockno*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buySyssum"
		sqlStr = sqlStr + " , Sum(s.realstockno*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buyRealsum"

		sqlStr = sqlStr + " , Sum(s.sysstockno*i.shopitemprice) as sellSyssum "
		sqlStr = sqlStr + " , Sum(s.realstockno*i.shopitemprice) as sellRealsum "

		sqlStr = sqlStr + " , 0 as avgIpgoPriceSum "
		sqlStr = sqlStr + " from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s"
		sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + "     on s.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
		''직영샵만.
		sqlStr = sqlStr + "     left Join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr + "     on s.shopid=u.userid"
		'''sqlStr = sqlStr + "     and u.shopdiv='1'"

		sqlStr = sqlStr + "     left join db_summary.dbo.tbl_monthly_shop_designer d" ''left join
		sqlStr = sqlStr + "     on d.yyyymm='"&FRectYYYYMM&"'"
		sqlStr = sqlStr + "     and s.shopid=d.shopid"
		sqlStr = sqlStr + "     and i.makerid=d.makerid"
		''sqlStr = sqlStr + "     and d.Comm_cd<>'B012'"  '''오재고 너무 많음.

		sqlStr = sqlStr + " where 1=1"
		if (FRectShopid<>"") then
		    sqlStr = sqlStr + " and s.shopid='" + FRectShopid + "'"
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
		end if
		if FRectMwDiv<>"" then
		    IF FRectMwDiv="M" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B021','B022','B031','B032')" ''B023
		    ELSEIF FRectMwDiv="W" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B011','B012','B013')"
		    ELSE
    			sqlStr = sqlStr + " and IsNULL(d.Comm_cd,'Z')='" + FRectMwDiv + "'"
    		END IF
		end if
		sqlStr = sqlStr + " and i.shopitemid<>0"
		IF (FRectShowMinus<>"on") then
    		sqlStr = sqlStr + " and s.sysstockno>0"   ''시스템재고/실사재고
    	End IF
		sqlStr = sqlStr + " group by u.userid, u.shopname,IsNULL(d.Comm_cd,'Z')" ''i.itemgubun,
		sqlStr = sqlStr + " order by u.userid, mwdiv "

        ''response.write "<!--" & sqlStr & "-->"
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
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FTotCount 		= rsget("totSysno")
				FItemList(i).FTotBuySum 	= rsget("buySyssum")
				'FItemList(i).FTotShopBuySum 	= rsget("shopbuysum")
				FItemList(i).FTotSellSum 	= rsget("sellSyssum")
				FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

                FItemList(i).Fshopid 	= rsget("userid")
                FItemList(i).Fshopname 	= rsget("shopname")

                FItemList(i).FTotRealCount   = rsget("totRealno")
                FItemList(i).FTotRealBuySum  = rsget("buyRealsum")
                FItemList(i).FTotRealSellSum = rsget("sellRealsum")

				''FItemList(i).Fitemgubun		= rsget("itemgubun")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub


	''매장 월말 재고액 합계
	public Sub GetShopMonthlyJeagoSumNew
	    dim sqlStr
	    Dim stockColNm, valPrice, valPriceSupp
		if FRectGubun="sys" then
		    stockColNm = "s.sysstockno"
		else
		    stockColNm = "(s.realstockno - s.errbaditemno)"
	    end if

        if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(isNULL(s.lstvatinclude,i.vatinclude), 'Y') = 'Y' then (isNULL(s.LstBuycash,0)*10/11) else isNULL(s.LstBuycash,0) end)"
			valPriceSupp = "(case when IsNull(isNULL(s.lstvatinclude,i.vatinclude), 'Y') = 'Y' then (isNULL(s.LstSuplycash,0)*10/11) else isNULL(s.LstSuplycash,0) end)"
		else
			valPrice = "isNULL(s.LstBuycash,0)"
			valPriceSupp = "isNULL(s.LstSuplycash,0)"
		end if

        sqlStr = ""
		sqlStr = sqlStr + " select isNULL(s.targetGbn,'OF') as targetGbn,u.userid, u.shopname, IsNULL(s.LstComm_cd,'Z') as mwdiv"   ''i.itemgubun,
		sqlStr = sqlStr + " ,sum("&stockColNm&") as totno "

		''sqlStr = sqlStr + " ,Sum("&stockColNm&"*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buysum"
		''sqlStr = sqlStr + " ,Sum("&stockColNm&"*(CASE WHEN i.shopbuyprice=0 THEN convert(money,(100-IsNULL(d.defaultsuplymargin,30))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as shopbuysum"

		'sqlStr = sqlStr + " , Sum("&stockColNm&"*(CASE WHEN "&valPrice&" is Not NULL THEN "&valPrice&" WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buysum"
		'sqlStr = sqlStr + " , Sum("&stockColNm&"*(CASE WHEN i.shopBuyprice=0 THEN convert(money,(100-IsNULL(d.defaultSuplymargin,30))/100*i.shopitemprice) ELSE i.shopBuyprice END)) as shopbuysum"
        sqlStr = sqlStr + " , Sum("&stockColNm&"*"&valPrice&") as buysum"
		sqlStr = sqlStr + " , Sum("&stockColNm&"*"&valPriceSupp&") as shopbuysum"

		sqlStr = sqlStr + " , Sum("&stockColNm&"*IsNull(i.shopitemprice,0)) as sellsum "
		sqlStr = sqlStr + " , 0 as avgIpgoPriceSum "
		sqlStr = sqlStr + " , sp.etcjungsantype"
		sqlStr = sqlStr + " from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s"
		sqlStr = sqlStr + "     Left Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + "     on 1 = 1 "
		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
        sqlStr = sqlStr + " 	and i.shopitemid<>0"

		sqlStr = sqlStr + "     left Join db_shop.dbo.tbl_shop_user u"
		sqlStr = sqlStr + "     on s.shopid=u.userid"
        sqlStr = sqlStr + "     left join db_partner.dbo.tbl_partner sp"
        sqlStr = sqlStr + "     on s.shopid=sp.id"

'        if FRectMwDiv<>"" and FRectMwDiv<>"Z" then
'            sqlStr = sqlStr + "     join db_summary.dbo.tbl_monthly_shop_designer d"
'        else
'		    sqlStr = sqlStr + "     left join db_summary.dbo.tbl_monthly_shop_designer d" ''left join
'		end if
'		sqlStr = sqlStr + "     on d.yyyymm='"&FRectYYYYMM&"'"
'		sqlStr = sqlStr + "     and s.shopid=d.shopid"
'		sqlStr = sqlStr + "     and i.makerid=d.makerid"
		''sqlStr = sqlStr + "     and d.Comm_cd<>'B012'"  '''오재고 너무 많음.

		sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and s.yyyymm='" + FRectYYYYMM + "' "
		''sqlStr = sqlStr + " and i.makerid not in ('yougreat')"
        if (FRectTargetGbn<>"") then
            if (FRectTargetGbn="TN") or (FRectTargetGbn = "3X") then
                sqlStr = sqlStr + " and (isNULL(s.targetGbn,'OF') not in ('ET','EG') ) "
            else
                sqlStr = sqlStr + " and isNULL(s.targetGbn,'OF')='"&FRectTargetGbn&"'"
            end if
        end if

		if (FRectShopid<>"") then
		    sqlStr = sqlStr + " and s.shopid='" + FRectShopid + "'"
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
		end if
        if (FRectVatYn<>"") then
            sqlStr = sqlStr + " and isNULL(isNULL(s.lstvatinclude,i.vatinclude),'Y')='" + FRectVatYn + "'"
        end if
		if FRectMwDiv<>"" then
		    IF FRectMwDiv="M" then
		        sqlStr = sqlStr + " and (s.LstComm_cd in ('B021','B022','B031','B032') or (isNULL(s.targetGbn,'OF') = 'IT' and s.LstComm_cd = 'B013'))" '''B023',
		        ''sqlStr = sqlStr + " and isNULL(s.targetGbn,'OF') not in ('ET','EG')"
		    ELSEIF FRectMwDiv="W" then
		        sqlStr = sqlStr + " and s.LstComm_cd in ('B011','B012','B013')"
				sqlStr = sqlStr + " and not (isNULL(s.targetGbn,'OF') = 'IT' and s.LstComm_cd = 'B013')"
		    ELSEIF FRectMwDiv="C" then
		        sqlStr = sqlStr + " and s.LstComm_cd in ('B013')"
		    ELSE
    			sqlStr = sqlStr + " and IsNULL(s.LstComm_cd,'Z')='" + FRectMwDiv + "'"
    		END IF
		end if

		IF (FRectShowMinus<>"on") then
    		sqlStr = sqlStr + " and "&stockColNm&">0"   ''시스템재고/실사재고
    	End IF
		IF (FRectShowMinusOnly = "on") then
    		sqlStr = sqlStr + " and "&stockColNm&"<0"   ''시스템재고/실사재고
    	End IF
    	IF (FRectetcjungsantype<>"") then
    	    if (FRectetcjungsantype="41") then
    	        sqlStr = sqlStr + " and (sp.etcjungsantype in ('4','1') or (s.shopid='streetshop803' and isNULL(s.targetGbn,'OF') = 'IT' and s.LstComm_cd = 'B013') ) "  ''or (isNULL(s.targetGbn,'OF') = 'IT' and s.LstComm_cd = 'B013') 추가
    	    else
    	        sqlStr = sqlStr + " and sp.etcjungsantype='"&FRectetcjungsantype&"'"
    	    end if
    	end if
		sqlStr = sqlStr + " group by isNULL(s.targetGbn,'OF'), u.userid, u.shopname,IsNULL(s.LstComm_cd,'Z') ,sp.etcjungsantype"
		sqlStr = sqlStr + " order by targetGbn desc, u.userid, mwdiv "

		''response.write "<!--" & sqlStr & "-->"
        ''response.end
        ''sqlStr =""
        ''sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Acc_ShopstockList] '"&FRectYYYYMM&"','"&FRectShopid&"','"&FRectMwDiv&"','"&FRectGubun&"',"&CHKIIF(FRectShowMinus<>"on","0","1")&",'"&FRectIsUsing&"','"&FRectVatYn&"',"&CHKIIF(FRectShopSuplyPrice="Y","1","0")&",'"&FRectTargetGbn&"'"
'rw 	sqlStr
'response.write "수정중"
'response.end
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		''rsget.LockType = adLockOptimistic '' ==> adLockReadOnly

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
				set FItemList(i) = new CMonthlyStockSum

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotShopBuySum 	= rsget("shopbuysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")
				FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

                FItemList(i).Fshopid 	= rsget("userid")
                FItemList(i).Fshopname 	= rsget("shopname")
                FItemList(i).FtargetGbn  = rsget("targetGbn")
                FItemList(i).Fetcjungsantype = rsget("etcjungsantype")
				''FItemList(i).Fitemgubun		= rsget("itemgubun")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

    public Sub GetShopMonthlyRealJeagoDetailByMakerSysWithReal()
	    dim sqlStr

		Dim stockColNm
		if FRectGubun="sys" then
		    stockColNm = "sysstockno"
		else
		    stockColNm = "realstockno"
	    end if
	    ''최대 N건 제한

		sqlStr = "select top "& FPageSize*FCurrPage&" i.itemgubun, i.shopitemid as itemid, i.regdate, i.shopitemname as itemname"
	    sqlStr = sqlStr + " , IsNULL(d.Comm_cd,'Z') as mwdiv"
	    sqlStr = sqlStr + " , IsNULL(i.itemoption,'0000') as itemoption"
	    sqlStr = sqlStr + " , IsNULL(i.shopitemoptionname,'') as itemoptionname"
	    sqlStr = sqlStr + " , i.isusing, IsNULL(i.isusing,'Y') as optionusing"
		sqlStr = sqlStr + " ,s.sysstockno as totno"
		sqlStr = sqlStr + " ,s.realstockno as totRealno"
		sqlStr = sqlStr + " ,s.sysstockno*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END) as buysum"
		sqlStr = sqlStr + " ,s.realstockno*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END) as buyRealsum"
		sqlStr = sqlStr + " , s.sysstockno*i.shopitemprice as sellsum "
		sqlStr = sqlStr + " , s.realstockno*i.shopitemprice as sellRealsum "
		sqlStr = sqlStr + " , 0 as avgIpgoPriceSum "
		sqlStr = sqlStr + " , i.shopitemprice "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_shopstock_summary s"
		sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + "     on s.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
		sqlStr = sqlStr + "     and i.makerid='"&FRECTMakerid&"'"
		sqlStr = sqlStr + "     left join db_summary.dbo.tbl_monthly_shop_designer d" ''left join
		sqlStr = sqlStr + "     on d.yyyymm='"&FRectYYYYMM&"'"
		sqlStr = sqlStr + "     and s.shopid=d.shopid"
		sqlStr = sqlStr + "     and i.makerid=d.makerid"
		sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
		IF (FRectItemGubun<>"") then
    		sqlStr = sqlStr + " and s.itemgubun='"&FRectItemGubun&"'"
    	ENd IF
        sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
        sqlStr = sqlStr + " and s.itemid=i.shopitemid"
        sqlStr = sqlStr + " and s.itemoption=i.itemoption"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and s.shopid='"&FRectShopID&"'"
		end if

		if FRectNewItem<>"" then
			sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
		end if

		if FRectIsUsing<>"" then'
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
		end if

		if FRectMwDiv<>"" then
		    IF FRectMwDiv="M" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B021','B022','B031','B032')" '''B023',
		    ELSEIF FRectMwDiv="W" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B011','B012')"
		    ELSEIF FRectMwDiv="C" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B013')"
		    ELSE
    			sqlStr = sqlStr + " and IsNULL(d.Comm_cd,'Z')='" + FRectMwDiv + "'"
    		END IF
		end if
		sqlStr = sqlStr + " and i.shopitemid<>0"
		if (FRectShowMinus<>"on") then
		    sqlStr = sqlStr + " and s.sysstockno>0"
	    end if
		sqlStr = sqlStr + " order by totno desc"
		''response.write sqlStr
''rw sqlStr
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
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).Fitemgubun     = rsget("itemgubun")
				FItemList(i).Fitemid 		= rsget("itemid")
				FItemList(i).Fregdate 		= rsget("regdate")
				FItemList(i).Fitemoption 	= rsget("itemoption")
				FItemList(i).Fitemname 		= db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).FMaeIpGubun	= rsget("mwdiv")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")

				FItemList(i).FTotRealCount 		= rsget("totRealno")
				FItemList(i).FTotRealBuySum 	= rsget("buyRealsum")
				FItemList(i).FTotRealSellSum 	= rsget("sellRealsum")

				FItemList(i).FIsUsing	= rsget("isusing")
				FItemList(i).FOptionUsing	= rsget("optionusing")

				FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				FITemList(i).FCurrshopitemprice= rsget("shopitemprice")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub


	public Sub GetShopMonthlyRealJeagoDetailByMakerNew()
	    dim sqlStr

		Dim stockColNm,valPrice,valPriceSupp
		if FRectGubun="sys" then
		    stockColNm = "sysstockno"
		else
		    stockColNm = "realstockno"
	    end if
	    if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(isNULL(s.lstvatinclude,i.vatinclude), 'Y') = 'Y' then (s.LstBuycash*10/11) else s.LstBuycash end)"
			valPriceSupp = "(case when IsNull(isNULL(s.lstvatinclude,i.vatinclude), 'Y') = 'Y' then (s.LstSuplycash*10/11) else s.LstSuplycash end)"
		else
			valPrice = "s.LstBuycash"
			valPriceSupp = "s.LstSuplycash"
		end if

	    ''최대 N건 제한

		sqlStr = "select top "& FPageSize*FCurrPage&" s.itemgubun, s.itemid, i.regdate, i.shopitemname as itemname"
	    sqlStr = sqlStr + " , IsNULL(s.LstComm_cd,'Z') as mwdiv"
	    sqlStr = sqlStr + " , IsNULL(s.lstcentermwdiv, 'Z') AS centermwdiv "
	    sqlStr = sqlStr + " , IsNULL(s.itemoption,'0000') as itemoption"
	    sqlStr = sqlStr + " , IsNULL(i.shopitemoptionname,'') as itemoptionname"
	    sqlStr = sqlStr + " , i.isusing, IsNULL(i.isusing,'Y') as optionusing"
		sqlStr = sqlStr + " ,s."&stockColNm&" as totno"
		'sqlStr = sqlStr + " ,s."&stockColNm&"*(CASE WHEN s.LstBuycash is Not NULL THEN s.LstBuycash WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END) as buysum"
		'sqlStr = sqlStr + " ,s."&stockColNm&"*(CASE WHEN i.shopBuyprice=0 THEN convert(money,(100-IsNULL(d.defaultSuplymargin,30))/100*i.shopitemprice) ELSE i.shopBuyprice END) as Shopbuysum"
		sqlStr = sqlStr + " , (s."&stockColNm&"*"&valPrice&") as buysum"
        sqlStr = sqlStr + " , (s."&stockColNm&"*"&valPriceSupp&") as shopbuysum"
		sqlStr = sqlStr + " , s."&stockColNm&"*IsNull(i.shopitemprice,0) as sellsum "
		sqlStr = sqlStr + " , 0 as avgIpgoPriceSum "
		sqlStr = sqlStr + " , IsNull(i.shopitemprice,0) as shopitemprice "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_shopstock_summary s"
		sqlStr = sqlStr + "     Left Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + "     on 1 = 1 "
		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
        sqlStr = sqlStr + " 	and i.shopitemid<>0"
		'sqlStr = sqlStr + "     left join db_summary.dbo.tbl_monthly_shop_designer d" ''left join
		'sqlStr = sqlStr + "     on d.yyyymm='"&FRectYYYYMM&"'"
		'sqlStr = sqlStr + "     and s.shopid=d.shopid"
		'sqlStr = sqlStr + "     and i.makerid=d.makerid"
		sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
        sqlStr = sqlStr + "     and s.LstMakerid='"&FRECTMakerid&"'"

		IF (FRectItemGubun<>"") then
    		sqlStr = sqlStr + " and s.itemgubun='"&FRectItemGubun&"'"
    	ENd IF

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and s.shopid='"&FRectShopID&"'"
		end if

		if FRectNewItem<>"" then
			sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
		end if

		if FRectIsUsing<>"" then'
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
		end if
        if (FRectVatYn<>"") then
            sqlStr = sqlStr + " and isNULL(isNULL(s.lstvatinclude,i.vatinclude),'Y')='" + FRectVatYn + "'"
        end if

		if FRectItemMwDiv<>"" then
			sqlStr = sqlStr + " and IsNULL(s.LstComm_cd, 'Z') = '" + FRectItemMwDiv + "'"
		end if

		if FRectMwDiv<>"" then
		    IF FRectMwDiv="M" then
		        sqlStr = sqlStr + " and s.LstComm_cd in ('B021','B022','B031','B032','B013')" '''B023',
		    ELSEIF FRectMwDiv="W" then
		        sqlStr = sqlStr + " and s.LstComm_cd in ('B011','B012')"
		    ELSEIF FRectMwDiv="C" then
		        sqlStr = sqlStr + " and s.LstComm_cd in ('B013')"
		    ELSE
    			sqlStr = sqlStr + " and IsNULL(s.LstComm_cd,'Z')='" + FRectMwDiv + "'"
    		END IF
		end if

		if (FRectShowMinus<>"on") then
		    sqlStr = sqlStr + " and s."&stockColNm&">0"
	    end if
		IF (FRectShowMinusOnly = "on") then
    		sqlStr = sqlStr + " and s."&stockColNm&"<0"   ''시스템재고/실사재고
    	End IF
	    if (FRectOrdTp="S") then
	        sqlStr = sqlStr + " order by totno desc,s.itemgubun,itemid,s.itemoption"
	    else
		    sqlStr = sqlStr + " order by s.itemgubun,itemid,s.itemoption"
        end if
		response.write "<!--" & sqlStr & "-->"


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
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).Fitemgubun     = rsget("itemgubun")
				FItemList(i).Fitemid 		= rsget("itemid")
				FItemList(i).Fregdate 		= rsget("regdate")
				FItemList(i).Fitemoption 	= rsget("itemoption")
				FItemList(i).Fitemname 		= db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))

				FItemList(i).FMaeIpGubun		= rsget("mwdiv")			'// 계약조건
				FItemList(i).FItemMaeIpGubun	= rsget("centermwdiv")		'// 상품속성

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotShopBuySum 	= rsget("Shopbuysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")

				FItemList(i).FIsUsing	= rsget("isusing")
				FItemList(i).FOptionUsing	= rsget("optionusing")

				FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				FITemList(i).FCurrshopitemprice= rsget("shopitemprice")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub


    public Sub GetShopMonthlyRealJeagoDetailSysWithReal()
	    dim sqlStr

		Dim stockColNm
		if FRectGubun="sys" then
		    stockColNm = "sysstockno"
		else
		    stockColNm = "realstockno"
	    end if

	    sqlStr = "select i.makerid"
	    sqlStr = sqlStr + " , sum(s.sysstockno) as totno"
	    sqlStr = sqlStr + " , sum(s.realstockno) as totRealno"
		sqlStr = sqlStr + " , Sum(s.sysstockno*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buysum"
		sqlStr = sqlStr + " , Sum(s.realstockno*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buyRealsum"
		sqlStr = sqlStr + " , Sum(s.sysstockno*i.shopitemprice) as sellsum, Sum(s.realstockno*i.shopitemprice) as sellRealsum, c.isusing "
		sqlStr = sqlStr + " , 0 as avgIpgoPriceSum "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_shopstock_summary s"
		sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + "     on s.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
		sqlStr = sqlStr + "     left join db_summary.dbo.tbl_monthly_shop_designer d" ''left join
		sqlStr = sqlStr + "     on d.yyyymm='"&FRectYYYYMM&"'"
		sqlStr = sqlStr + "     and s.shopid=d.shopid"
		sqlStr = sqlStr + "     and i.makerid=d.makerid"
		sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
		IF (FRectItemGubun<>"") then
    		sqlStr = sqlStr + " and s.itemgubun='"&FRectItemGubun&"'"
    	end if
        sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
        sqlStr = sqlStr + " and s.itemid=i.shopitemid"
        sqlStr = sqlStr + " and s.itemoption=i.itemoption"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and s.shopid='"&FRectShopID&"'"
		end if

		if FRectNewItem<>"" then
			sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
		end if

		if FRectIsUsing<>"" then'
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
		end if

		if FRectMwDiv<>"" then
		    IF FRectMwDiv="M" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B021','B022','B031','B032')" '''B023',
		    ELSEIF FRectMwDiv="W" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B011','B012')"
		    ELSEIF FRectMwDiv="C" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B013')"
		    ELSE
    			sqlStr = sqlStr + " and IsNULL(d.Comm_cd,'Z')='" + FRectMwDiv + "'"
    		END IF
		end if
		sqlStr = sqlStr + " and i.shopitemid<>0"
		if (FRectShowMinus<>"on") then
    		sqlStr = sqlStr + " and s.sysstockno>0"
    	end if
		sqlStr = sqlStr + " group by i.makerid,  c.isusing"
		sqlStr = sqlStr + " order by totno desc"

		rsget.Open sqlStr,dbget,1
''response.write sqlStr
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FMakerid 	= rsget("makerid")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")

				FItemList(i).FTotRealCount 		= rsget("totRealno")
				FItemList(i).FTotRealBuySum 	= rsget("buyRealsum")
				FItemList(i).FTotRealSellSum 	= rsget("sellRealsum")

                FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				FItemList(i).FMakerUsing	= rsget("isusing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

    public Sub GetShopMonthlyRealJeagoDetailByShopidSysWithReal()
	    dim sqlStr

		Dim stockColNm
		if FRectGubun="sys" then
		    stockColNm = "sysstockno"
		else
		    stockColNm = "realstockno"
	    end if

	    sqlStr = "select s.shopid, n.socname as shopname"
	    sqlStr = sqlStr + " , sum(s.sysstockno) as totno"
	    sqlStr = sqlStr + " , sum(s.realstockno) as totRealno"
		sqlStr = sqlStr + " , Sum(s.sysstockno*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buysum"
		sqlStr = sqlStr + " , Sum(s.realstockno*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buyRealsum"
		sqlStr = sqlStr + " , Sum(s.sysstockno*i.shopitemprice) as sellsum, Sum(s.realstockno*i.shopitemprice) as sellRealsum, c.isusing "
		sqlStr = sqlStr + " , 0 as avgIpgoPriceSum "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_shopstock_summary s"
		sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + "     on s.yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
		sqlStr = sqlStr + "     left join db_summary.dbo.tbl_monthly_shop_designer d" ''left join
		sqlStr = sqlStr + "     on d.yyyymm='"&FRectYYYYMM&"'"
		sqlStr = sqlStr + "     and s.shopid=d.shopid"
		sqlStr = sqlStr + "     and i.makerid=d.makerid"
		sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c n on s.shopid=n.userid"
		sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
		IF (FRectItemGubun<>"") then
    		sqlStr = sqlStr + " and s.itemgubun='"&FRectItemGubun&"'"
    	end if
        sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
        sqlStr = sqlStr + " and s.itemid=i.shopitemid"
        sqlStr = sqlStr + " and s.itemoption=i.itemoption"

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and s.shopid='"&FRectShopID&"'"
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid='"&FRectMakerid&"' "
		end if

		if FRectNewItem<>"" then
			sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
		end if

		if FRectIsUsing<>"" then'
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
		end if

		if FRectMwDiv<>"" then
		    IF FRectMwDiv="M" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B021','B022','B031','B032')" '''B023',
		    ELSEIF FRectMwDiv="W" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B011','B012')"
		    ELSEIF FRectMwDiv="C" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B013')"
		    ELSE
    			sqlStr = sqlStr + " and IsNULL(d.Comm_cd,'Z')='" + FRectMwDiv + "'"
    		END IF
		end if
		sqlStr = sqlStr + " and i.shopitemid<>0"
		if (FRectShowMinus<>"on") then
    		sqlStr = sqlStr + " and s.sysstockno>0"
    	end if
		sqlStr = sqlStr + " group by s.shopid, n.socname, c.isusing"
		sqlStr = sqlStr + " order by totno desc"

		rsget.Open sqlStr,dbget,1
''response.write sqlStr
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).Fshopid 	= rsget("shopid")
				FItemList(i).Fshopname 	= rsget("shopname")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")

				FItemList(i).FTotRealCount 		= rsget("totRealno")
				FItemList(i).FTotRealBuySum 	= rsget("buyRealsum")
				FItemList(i).FTotRealSellSum 	= rsget("sellRealsum")

                FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				FItemList(i).FMakerUsing	= rsget("isusing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

	public Sub GetShopMonthlyRealJeagoDetailNew()
	    dim sqlStr

		Dim stockColNm,valPrice,valPriceSupp

		if FRectGubun="sys" then
		    stockColNm = "sysstockno"
		else
		    stockColNm = "realstockno"
	    end if

        if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(isNULL(s.lstvatinclude,i.vatinclude), 'Y') = 'Y' then (isNULL(s.LstBuycash,0)*10/11) else isNULL(s.LstBuycash,0) end)"
			valPriceSupp = "(case when IsNull(isNULL(s.lstvatinclude,i.vatinclude), 'Y') = 'Y' then (isNULL(s.LstSuplycash,0)*10/11) else isNULL(s.LstSuplycash,0) end)"
		else
			valPrice = "isNULL(s.LstBuycash,0)"
			valPriceSupp = "isNULL(s.LstSuplycash,0)"
		end if

	    sqlStr = "select isNULL(s.targetGbn,'OF') as targetGbn,s.LstMakerid as makerid, sum(s."&stockColNm&") as totno"
		'sqlStr = sqlStr + " , Sum(s."&stockColNm&"*(CASE WHEN s.LstBuycash is Not NULL THEN s.LstBuycash WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buysum"
		'sqlStr = sqlStr + " , Sum(s."&stockColNm&"*(CASE WHEN i.shopBuyprice=0 THEN convert(money,(100-IsNULL(d.defaultSuplymargin,30))/100*i.shopitemprice) ELSE i.shopBuyprice END)) as shopbuysum"
		sqlStr = sqlStr + " , Sum(s."&stockColNm&"*"&valPrice&") as buysum"
        sqlStr = sqlStr + " , Sum(s."&stockColNm&"*"&valPriceSupp&") as shopbuysum"
		sqlStr = sqlStr + " , Sum(s."&stockColNm&"*IsNull(i.shopitemprice,0)) as sellsum, c.isusing "
		sqlStr = sqlStr + " , 0 as avgIpgoPriceSum "
		sqlStr = sqlStr + " , s.LstComm_cd as Comm_cd"
		sqlStr = sqlStr + " , sum(case when s.errrealcheckno<>0 then 1 else 0 end) as ErrItemCnt "
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_shopstock_summary s"
		sqlStr = sqlStr + "     Left Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + "     on 1 = 1 "
		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
        sqlStr = sqlStr + " 	and i.shopitemid<>0"
		'sqlStr = sqlStr + "     left join db_summary.dbo.tbl_monthly_shop_designer d" ''left join
		'sqlStr = sqlStr + "     on d.yyyymm='"&FRectYYYYMM&"'"
		'sqlStr = sqlStr + "     and s.shopid=d.shopid"
		'sqlStr = sqlStr + "     and i.makerid=d.makerid"
		sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c c on s.LstMakerid=c.userid"
		sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
		if (FRectTargetGbn<>"") then
			if (FRectTargetGbn = "3X") then
				sqlStr = sqlStr + " and isNULL(s.targetGbn,'OF') not in ('ET', 'EG') "
			else
				sqlStr = sqlStr + " and isNULL(s.targetGbn,'OF')='" + FRectTargetGbn + "'"
			end if
        end if
		IF (FRectItemGubun<>"") then
    		sqlStr = sqlStr + " and s.itemgubun='"&FRectItemGubun&"'"
    	end if

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and s.shopid='"&FRectShopID&"'"
		end if

		if FRectNewItem<>"" then
			sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
		end if

		if FRectIsUsing<>"" then'
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
		end if
        if (FRectVatYn<>"") then
            sqlStr = sqlStr + " and isNULL(isNULL(s.lstvatinclude,i.vatinclude),'Y')='" + FRectVatYn + "'"
        end if

		if FRectMwDiv<>"" then
		    IF FRectMwDiv="M" then
		        sqlStr = sqlStr + " and s.LstComm_cd in ('B021','B022','B031','B032','B013')" '''B023',
		    ELSEIF FRectMwDiv="W" then
		        sqlStr = sqlStr + " and s.LstComm_cd in ('B011','B012')"
		    ELSEIF FRectMwDiv="C" then
		        sqlStr = sqlStr + " and s.LstComm_cd in ('B013')"
		    ELSE
    			sqlStr = sqlStr + " and IsNULL(s.LstComm_cd,'Z')='" + FRectMwDiv + "'"
    		END IF
		end if

		if (FRectShowMinus<>"on") then
    		sqlStr = sqlStr + " and s."&stockColNm&">0"
    	end if
		IF (FRectShowMinusOnly = "on") then
    		sqlStr = sqlStr + " and s."&stockColNm&"<0"   ''시스템재고/실사재고
    	End IF
		sqlStr = sqlStr + " group by isNULL(s.targetGbn,'OF') ,s.LstMakerid,  s.LstComm_cd, c.isusing"
		sqlStr = sqlStr + " order by targetGbn desc, totno desc"
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
			do until rsget.eof
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FMakerid 	= rsget("makerid")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotShopBuySum 	= rsget("shopbuysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")
                FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				FItemList(i).FMakerUsing	= rsget("isusing")

				FItemList(i).FMaeIpGubun    = rsget("Comm_cd")
				FItemList(i).FtargetGbn     = rsget("targetGbn")

				FItemList(i).FErrItemCnt    = rsget("ErrItemCnt")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

	public Sub GetShopMonthlyRealJeagoDetailByShopidNew()
	    dim sqlStr

		Dim stockColNm,valPrice,valPriceSupp
		if FRectGubun="sys" then
		    stockColNm = "sysstockno"
		else
		    stockColNm = "realstockno"
	    end if
        if (FRectShopSuplyPrice = "Y") then
			'// 공급가 표시(세금 제외)
			valPrice = "(case when IsNull(isNULL(s.lstvatinclude,i.vatinclude), 'Y') = 'Y' then (s.LstBuycash*10/11) else s.LstBuycash end)"
			valPriceSupp = "(case when IsNull(isNULL(s.lstvatinclude,i.vatinclude), 'Y') = 'Y' then (s.LstSuplycash*10/11) else s.LstSuplycash end)"
		else
			valPrice = "s.LstBuycash"
			valPriceSupp = "s.LstSuplycash"
		end if

	    sqlStr = "select s.shopid, n.socname as shopname, sum(s."&stockColNm&") as totno"
		'sqlStr = sqlStr + " , Sum(s."&stockColNm&"*(CASE WHEN s.LstBuycash is Not NULL THEN s.LstBuycash WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buysum"
		'sqlStr = sqlStr + " , Sum(s."&stockColNm&"*(CASE WHEN i.shopBuyprice=0 THEN convert(money,(100-IsNULL(d.defaultSuplymargin,30))/100*i.shopitemprice) ELSE i.shopBuyprice END)) as shopbuysum"
		sqlStr = sqlStr + " , Sum(s."&stockColNm&"*"&valPrice&") as buysum"
        sqlStr = sqlStr + " , Sum(s."&stockColNm&"*"&valPriceSupp&") as shopbuysum"
		sqlStr = sqlStr + " , Sum(s."&stockColNm&"*IsNull(i.shopitemprice,0)) as sellsum, c.isusing "
		sqlStr = sqlStr + " , 0 as avgIpgoPriceSum "
		sqlStr = sqlStr + " , d.Comm_cd"
		sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_shopstock_summary s"
		sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shop_item i"
		sqlStr = sqlStr + "     on 1 = 1 "
		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
        sqlStr = sqlStr + " 	and i.shopitemid<>0"
		sqlStr = sqlStr + "     left join db_summary.dbo.tbl_monthly_shop_designer d" ''left join
		sqlStr = sqlStr + "     on d.yyyymm='"&FRectYYYYMM&"'"
		sqlStr = sqlStr + "     and s.shopid=d.shopid"
		sqlStr = sqlStr + "     and i.makerid=d.makerid"
		sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
		sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c n on s.shopid=n.userid"
		sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
		IF (FRectItemGubun<>"") then
    		sqlStr = sqlStr + " and s.itemgubun='"&FRectItemGubun&"'"
    	end if

		if FRectShopID<>"" then
			sqlStr = sqlStr + " and s.shopid='"&FRectShopID&"'"
		end if

		if FRectMakerid<>"" then
			sqlStr = sqlStr + " and i.makerid='"&FRectMakerid&"' "
		end if

		if FRectNewItem<>"" then
			sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
		end if

		if FRectIsUsing<>"" then'
			sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
		end if
        if (FRectVatYn<>"") then
            sqlStr = sqlStr + " and isNULL(isNULL(s.lstvatinclude,i.vatinclude),'Y')='" + FRectVatYn + "'"
        end if
		if FRectMwDiv<>"" then
		    IF FRectMwDiv="M" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B021','B022','B031','B032')" '''B023',
		    ELSEIF FRectMwDiv="W" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B011','B012')"
		    ELSEIF FRectMwDiv="C" then
		        sqlStr = sqlStr + " and d.Comm_cd in ('B013')"
		    ELSE
    			sqlStr = sqlStr + " and IsNULL(d.Comm_cd,'Z')='" + FRectMwDiv + "'"
    		END IF
		end if

		if (FRectShowMinus<>"on") then
    		sqlStr = sqlStr + " and s."&stockColNm&">0"
    	end if
		IF (FRectShowMinusOnly = "on") then
    		sqlStr = sqlStr + " and s."&stockColNm&"<0"   ''시스템재고/실사재고
    	End IF
		sqlStr = sqlStr + " group by s.shopid, n.socname, c.isusing, d.Comm_cd"
		sqlStr = sqlStr + " order by totno desc"
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
			do until rsget.eof
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).Fshopid 	= rsget("shopid")
				FItemList(i).Fshopname 	= rsget("shopname")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotShopBuySum 	= rsget("shopbuysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")
                FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				FItemList(i).FMakerUsing	= rsget("isusing")

				FItemList(i).FMaeIpGubun    = rsget("Comm_cd")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

	''---상세 내역

	public Sub GetMonthlyRealJeagoDetailNew()
		dim sqlStr

		Dim stockColNm
		if FRectGubun="sys" then
		    stockColNm = "s.totsysstock"
		else
		    stockColNm = "(s.realstock-s.errbaditemno)"
	    end if

		IF FRectItemGubun<>"10" then
		    sqlStr = "select i.makerid, c.socname, sum("&stockColNm&") as totno"
			sqlStr = sqlStr + " , Sum("&stockColNm&"*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as buysum"
			sqlStr = sqlStr + " , Sum("&stockColNm&"*i.shopitemprice) as sellsum, c.isusing "
			sqlStr = sqlStr + " , Sum("&stockColNm&"*IsNULL(s.avgIpgoPrice,(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END))) as avgIpgoPriceSum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
			sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shop_item i"
    		sqlStr = sqlStr + "     on s.yyyymm='" + FRectYYYYMM + "'"
    		sqlStr = sqlStr + "     and s.itemgubun<>'10'"
    		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
    		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
    		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
    		if (FRectOFFReturn2OnStock<>"") then
    		    sqlStr = sqlStr + " and Not (s.itemgubun='90' and s.itemid=1385)"
    		end if
    		sqlStr = sqlStr + "     left Join db_shop.dbo.tbl_shop_designer d"
    		sqlStr = sqlStr + "     on d.shopid='streetshop000'"
    		sqlStr = sqlStr + "     and d.makerid=i.makerid"
			sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun='"&FRectItemGubun&"'"
            sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
            sqlStr = sqlStr + " and s.itemid=i.shopitemid"
            sqlStr = sqlStr + " and s.itemoption=i.itemoption"

			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if

			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
            if (FRectVatYn<>"") then
    		    sqlStr = sqlStr + " and isNULL(s.lastvatinclude,i.vatinclude)='" + FRectVatYn + "'"
    		end if
			if FRectMwDiv<>"" then
				sqlStr = sqlStr + " and IsNULL(s.lastmwdiv,'Z')='" + FRectMwDiv + "'"
				''sqlStr = sqlStr + " and (CASE WHEN i.centermwdiv<>'U' THEN i.centermwdiv ELSE IsNULL(s.lastmwdiv,'Z') END)='" + FRectMwDiv + "'"
			end if

			sqlStr = sqlStr + " and i.shopitemid<>0"
			sqlStr = sqlStr + " and "&stockColNm&">0"
			sqlStr = sqlStr + " group by i.makerid,  c.socname, c.isusing"
			sqlStr = sqlStr + " order by totno desc"
		ELSE
		    sqlStr = ""
		    IF (FRectOFFReturn2OnStock<>"") then
		        sqlStr = sqlStr + " select makerid , socname, sum(totno) as totno, sum(buysum) as buysum, sum(sellsum) as sellsum"
                sqlStr = sqlStr + " , isusing, sum(avgIpgoPriceSum) as avgIpgoPriceSum"
                sqlStr = sqlStr + " From ("
		        sqlStr = sqlStr + " select i.shopitemoptionname as makerid, c.socname, sum("&stockColNm&") as totno"
    			sqlStr = sqlStr + " , Sum("&stockColNm&"*i.shopsuplycash) as buysum"
    			sqlStr = sqlStr + " , Sum("&stockColNm&"*i.shopitemprice) as sellsum, 'N' as isusing "
    			sqlStr = sqlStr + " , Sum("&stockColNm&"*IsNULL(s.avgIpgoPrice,i.shopsuplycash)) as avgIpgoPriceSum "
    			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
    			sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shop_item i"
        		sqlStr = sqlStr + "     on s.yyyymm='" + FRectYYYYMM + "'"
        		sqlStr = sqlStr + "     and s.itemgubun='90' and s.itemid=1385 "
        		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
        		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
        		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
    			sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c c on i.shopitemoptionname=c.userid"
    			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"

    			if FRectNewItem<>"" then
    				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    			end if

    			if FRectIsUsing<>"" then
    				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
    			end if
                if (FRectVatYn<>"") then
        		    sqlStr = sqlStr + " and isNULL(s.lastvatinclude,i.vatinclude)='" + FRectVatYn + "'"
        		end if
    			if FRectMwDiv<>"" then
    				sqlStr = sqlStr + " and IsNULL(s.lastmwdiv,'Z')='" + FRectMwDiv + "'"
    			end if

    			sqlStr = sqlStr + " and i.shopitemid<>0"
    			sqlStr = sqlStr + " and "&stockColNm&">0"
    			sqlStr = sqlStr + " group by i.shopitemoptionname,  c.socname, c.isusing"
    			sqlStr = sqlStr + " union "
		    END IF
			sqlStr = sqlStr + " select i.makerid, c.socname, sum("&stockColNm&") as totno "
			sqlStr = sqlStr + " , Sum("&stockColNm&"*i.orgsuplycash) as buysum"
			sqlStr = sqlStr + " , Sum("&stockColNm&"*i.orgprice) as sellsum, c.isusing "
			sqlStr = sqlStr + " , Sum("&stockColNm&"*IsNULL(s.avgIpgoPrice,i.orgsuplycash)) as avgIpgoPriceSum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
			sqlStr = sqlStr + "     join [db_item].[dbo].tbl_item i"
    		sqlStr = sqlStr + "     on  s.yyyymm='" + FRectYYYYMM + "'"
    		sqlStr = sqlStr + "     and s.itemgubun='10'"
    		sqlStr = sqlStr + "     and s.itemid=i.itemid"
			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
			sqlStr = sqlStr + " where 1=1"
			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if

			if FRectIsUsing<>"" then
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
            if (FRectVatYn<>"") then
    		    sqlStr = sqlStr + " and isNULL(s.lastvatinclude,i.vatinclude)='" + FRectVatYn + "'"
    		end if
			if FRectMwDiv<>"" then
				sqlStr = sqlStr + " and IsNULL(s.lastmwdiv,'Z')='" + FRectMwDiv + "'"
				''sqlStr = sqlStr + " and (CASE WHEN i.mwdiv<>'U' THEN i.mwdiv ELSE IsNULL(s.lastmwdiv,'Z') END)='" + FRectMwDiv + "'"
			end if

			sqlStr = sqlStr + " and i.itemid<>0"
			sqlStr = sqlStr + " and i.itemid<>11406"  	''포장비
			sqlStr = sqlStr + " and i.itemid<>6400"		''배송비
			sqlStr = sqlStr + " and "&stockColNm&">0"
			sqlStr = sqlStr + " group by i.makerid, c.socname, c.isusing"
			IF (FRectOFFReturn2OnStock<>"") then
			    sqlStr = sqlStr + " ) T"
			    sqlStr = sqlStr + " group by makerid, socname, isusing"
			end if
			sqlStr = sqlStr + " order by totno desc"
	    END IF


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
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FMakerid 	= rsget("makerid")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")
                FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				FItemList(i).FMakerUsing	= rsget("isusing")
				FItemList(i).FSocName       = db2HTML(rsget("socname"))
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end Sub

	public Sub GetMonthlyRealJeagoDetailByMakerNew()
		dim sqlStr

        Dim stockColNm
		if FRectGubun="sys" then
		    stockColNm = "s.totsysstock"
		else
		    stockColNm = "(s.realstock-s.errbaditemno)"
	    end if

		IF FRectItemGubun<>"10" then
		    sqlStr = "select s.itemgubun,i.shopitemid as itemid, i.regdate, i.shopitemname as itemname"
		    sqlStr = sqlStr + " , IsNULL(s.lastmwdiv,'Z') as mwdiv"  ''(CASE WHEN IsNULL(i.centermwdiv,'')<>'' THEN i.centermwdiv ELSE IsNULL(s.lastmwdiv,'Z') END)
		    sqlStr = sqlStr + " , IsNULL(i.itemoption,'0000') as itemoption"
		    sqlStr = sqlStr + " , IsNULL(i.shopitemoptionname,'') as itemoptionname"
		    sqlStr = sqlStr + " , i.isusing, IsNULL(i.isusing,'Y') as optionusing"
			sqlStr = sqlStr + " ,"&stockColNm&" as totno"
			sqlStr = sqlStr + " ,"&stockColNm&"*(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END) as buysum"
			sqlStr = sqlStr + " , "&stockColNm&"*i.shopitemprice as sellsum "
			sqlStr = sqlStr + " , "&stockColNm&"*IsNULL(s.avgIpgoPrice,(CASE WHEN i.shopsuplycash=0 THEN convert(money,(100-IsNULL(d.defaultmargin,35))/100*i.shopitemprice) ELSE i.shopsuplycash END)) as avgIpgoPriceSum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
			sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shop_item i"
    		sqlStr = sqlStr + "     on s.yyyymm='" + FRectYYYYMM + "'"
    		''sqlStr = sqlStr + "     and s.itemgubun<>'10'"
    		sqlStr = sqlStr + "     and s.itemgubun='"&FRectItemGubun&"'"
    		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
    		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
    		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"
    		if (FRectOFFReturn2OnStock<>"") then
    		    sqlStr = sqlStr + " and Not (s.itemgubun='90' and s.itemid=1385)"
    		end if
    		sqlStr = sqlStr + "     left Join db_shop.dbo.tbl_shop_designer d"
    		sqlStr = sqlStr + "     on d.shopid='streetshop000'"
    		sqlStr = sqlStr + "     and d.makerid=i.makerid"
			sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
			sqlStr = sqlStr + " where 1=1"
			sqlStr = sqlStr + " and i.makerid='" + FRectMakerid + "'"

			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if

			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
            if (FRectVatYn<>"") then
    		    sqlStr = sqlStr + " and isNULL(s.lastvatinclude,i.vatinclude)='" + FRectVatYn + "'"
    		end if

			if FRectMwDiv<>"" then
				sqlStr = sqlStr + " and IsNULL(s.lastmwdiv,'Z')='" + FRectMwDiv + "'"
                ''sqlStr = sqlStr + " and (CASE WHEN i.centermwdiv<>'U' THEN i.centermwdiv ELSE IsNULL(s.lastmwdiv,'Z') END)='" + FRectMwDiv + "'"
			end if

			sqlStr = sqlStr + " and i.shopitemid<>0"
			sqlStr = sqlStr + " and "&stockColNm&">0"
			''sqlStr = sqlStr + " order by totno desc"
            sqlStr = sqlStr + " order by s.itemgubun,itemid,itemoption"
		ELSE
		    sqlStr = ""
		    if (FRectOFFReturn2OnStock<>"") then
    		    sqlStr = sqlStr + " select '10' as itemgubun,i.shopitemid as itemid, i.regdate, i.shopitemname as itemname"
    		    sqlStr = sqlStr + " , IsNULL(s.lastmwdiv,'Z') as mwdiv"
    		    sqlStr = sqlStr + " , IsNULL(i.itemoption,'0000') as itemoption"
    		    sqlStr = sqlStr + " , '' as itemoptionname"
    		    sqlStr = sqlStr + " , i.isusing, IsNULL(i.isusing,'Y') as optionusing"
    			sqlStr = sqlStr + " ,"&stockColNm&" as totno"
    			sqlStr = sqlStr + " ,"&stockColNm&"*i.shopsuplycash as buysum"
    			sqlStr = sqlStr + " , "&stockColNm&"*i.shopitemprice as sellsum "
    			sqlStr = sqlStr + " , "&stockColNm&"*IsNULL(s.avgIpgoPrice,i.shopsuplycash) as avgIpgoPriceSum "
    			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
    			sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shop_item i"
        		sqlStr = sqlStr + "     on s.yyyymm='" + FRectYYYYMM + "'"
        		sqlStr = sqlStr + "     and s.itemgubun='90' and s.itemid=1385"
        		sqlStr = sqlStr + "     and s.itemgubun=i.itemgubun"
        		sqlStr = sqlStr + "     and s.itemid=i.shopitemid"
        		sqlStr = sqlStr + "     and s.itemoption=i.itemoption"

    			sqlStr = sqlStr + "     left join [db_user].[dbo].tbl_user_c c on i.shopitemname=c.userid"
    			sqlStr = sqlStr + " where 1=1"
    			sqlStr = sqlStr + " and i.shopitemoptionname='" + FRectMakerid + "'"

    			if FRectNewItem<>"" then
    				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    			end if

    			if FRectIsUsing<>"" then'
    				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
    			end if
                if (FRectVatYn<>"") then
        		    sqlStr = sqlStr + " and isNULL(s.lastvatinclude,i.vatinclude)='" + FRectVatYn + "'"
        		end if

    			if FRectMwDiv<>"" then
    				sqlStr = sqlStr + " and IsNULL(s.lastmwdiv,'Z')='" + FRectMwDiv + "'"
    			end if

    			sqlStr = sqlStr + " and i.shopitemid<>0"
    			sqlStr = sqlStr + " and "&stockColNm&">0"
    			sqlStr = sqlStr + " union"
    		end if

			sqlStr = sqlStr + " select s.itemgubun,i.itemid, i.regdate, i.itemname"
			sqlStr = sqlStr + " , IsNULL(s.lastmwdiv,'Z') as mwdiv"  ''(CASE WHEN i.mwdiv<>'U' THEN i.mwdiv ELSE IsNULL(s.lastmwdiv,'Z') END)
			sqlStr = sqlStr + " , IsNULL(o.itemoption,'0000') as itemoption, IsNULL(o.optionname,'') as itemoptionname"
			sqlStr = sqlStr + " , i.isusing, IsNULL(o.isusing,'Y') as optionusing,"
			sqlStr = sqlStr + " "&stockColNm&" as totno ,"
			sqlStr = sqlStr + " "&stockColNm&"*i.orgsuplycash as buysum, "&stockColNm&"*i.orgprice as sellsum "
			sqlStr = sqlStr + " ,"&stockColNm&"*IsNULL(s.avgIpgoPrice,i.orgsuplycash) as avgIpgoPriceSum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
			sqlStr = sqlStr + "     Join [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + "     on s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + "     and s.itemgubun='10'"
			sqlStr = sqlStr + "     and s.itemid=i.itemid"
			sqlStr = sqlStr + "     and i.makerid='" + FRectMakerid + "'"

			sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o on s.yyyymm='" + FRectYYYYMM + "' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"
			sqlStr = sqlStr + " where 1=1"

			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if

			if FRectIsUsing<>"" then
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
            if (FRectVatYn<>"") then
    		    sqlStr = sqlStr + " and isNULL(s.lastvatinclude,i.vatinclude)='" + FRectVatYn + "'"
    		end if

			if FRectMwDiv<>"" then
				sqlStr = sqlStr + " and IsNULL(s.lastmwdiv,'Z')='" + FRectMwDiv + "'"
				''sqlStr = sqlStr + " and (CASE WHEN i.mwdiv<>'U' THEN i.mwdiv ELSE IsNULL(s.lastmwdiv,'Z') END)='" + FRectMwDiv + "'"
			end if

			sqlStr = sqlStr + " and i.itemid<>0"
			sqlStr = sqlStr + " and i.itemid<>11406"  	''포장비
			sqlStr = sqlStr + " and i.itemid<>6400"		''배송비
			sqlStr = sqlStr + " and "&stockColNm&">0"
			''sqlStr = sqlStr + " order by totno desc"
			sqlStr = sqlStr + " order by s.itemgubun,s.itemid,s.itemoption"
		end IF

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
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).Fitemgubun     = rsget("itemgubun")
				FItemList(i).Fitemid 		= rsget("itemid")
				FItemList(i).Fregdate 		= rsget("regdate")
				FItemList(i).Fitemoption 	= rsget("itemoption")
				FItemList(i).Fitemname 		= db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).FMaeIpGubun	= rsget("mwdiv")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")

				FItemList(i).FIsUsing	= rsget("isusing")
				FItemList(i).FOptionUsing	= rsget("optionusing")

				FITemList(i).FavgIpgoPriceSum = rsget("avgIpgoPriceSum")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end Sub


''================기존 방식 =======================================================================================

	public Sub GetMonthlyRealJeagoDetailByMaker()
		dim sqlStr

		if FRectGubun="sys" then
			''시스템재고
			IF FRectItemGubun<>"10" then
			    sqlStr = "select i.shopitemid as itemid, i.regdate, i.shopitemname as itemname, i.centermwdiv as mwdiv, IsNULL(i.itemoption,'0000') as itemoption, IsNULL(i.shopitemoptionname,'') as itemoptionname, i.isusing, IsNULL(i.isusing,'Y') as optionusing,"
    			sqlStr = sqlStr + " s.totsysstock as totno ,"
    			sqlStr = sqlStr + " s.totsysstock*i.shopsuplycash as buysum, s.totsysstock*i.shopitemprice as sellsum "
    			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
    			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item i"
    			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
    			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
    			sqlStr = sqlStr + " and s.itemgubun='"&FRectItemGubun&"'"
                sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
                sqlStr = sqlStr + " and s.itemid=i.shopitemid"
                sqlStr = sqlStr + " and s.itemoption=i.itemoption"

    			sqlStr = sqlStr + " and i.makerid='" + FRectMakerid + "'"

    			if FRectNewItem<>"" then
    				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    			end if

    			if FRectIsUsing<>"" then'
    				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
    			end if

    			if FRectMwDiv<>"" then
    				sqlStr = sqlStr + " and i.centermwdiv='" + FRectMwDiv + "'"
    			end if

    			sqlStr = sqlStr + " and i.shopitemid<>0"
    			sqlStr = sqlStr + " and s.totsysstock>0"
    			sqlStr = sqlStr + " order by totno desc"
			ELSE
    			sqlStr = "select i.itemid, i.regdate, i.itemname, i.mwdiv, IsNULL(o.itemoption,'0000') as itemoption, IsNULL(o.optionname,'') as itemoptionname, i.isusing, IsNULL(o.isusing,'Y') as optionusing,"
    			sqlStr = sqlStr + " s.totsysstock as totno ,"
    			sqlStr = sqlStr + " s.totsysstock*i.buycash as buysum, s.totsysstock*i.sellcash as sellsum "
    			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
    			sqlStr = sqlStr + " [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
    			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on s.yyyymm='" + FRectYYYYMM + "' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"
    			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
    			sqlStr = sqlStr + " and s.itemgubun='10'"
    			sqlStr = sqlStr + " and s.itemid=i.itemid"
    			sqlStr = sqlStr + " and i.makerid='" + FRectMakerid + "'"

    			if FRectNewItem<>"" then
    				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    			end if

    			if FRectIsUsing<>"" then
    				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
    			end if

    			if FRectMwDiv<>"" then
    				sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
    			end if

    			sqlStr = sqlStr + " and i.itemid<>0"
    			sqlStr = sqlStr + " and i.itemid<>11406"  	''포장비
    			sqlStr = sqlStr + " and i.itemid<>6400"		''배송비
    			sqlStr = sqlStr + " and s.totsysstock>0"
    			sqlStr = sqlStr + " order by totno desc"
    		end IF
		else
			''실사재고
            IF FRectItemGubun<>"10" then
			    sqlStr = "select i.shopitemid as itemid, i.regdate, i.shopitemname as itemname, i.centermwdiv as mwdiv, IsNULL(i.itemoption,'0000') as itemoption, IsNULL(i.shopitemoptionname,'') as itemoptionname, i.isusing, IsNULL(i.isusing,'Y') as optionusing,"
    			sqlStr = sqlStr + " s.realstock as totno ,"
    			sqlStr = sqlStr + " s.realstock*i.shopsuplycash as buysum, s.realstock*i.shopitemprice as sellsum "
    			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
    			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item i"
    			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
    			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
    			sqlStr = sqlStr + " and s.itemgubun='"&FRectItemGubun&"'"
                sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
                sqlStr = sqlStr + " and s.itemid=i.shopitemid"
                sqlStr = sqlStr + " and s.itemoption=i.itemoption"

    			sqlStr = sqlStr + " and i.makerid='" + FRectMakerid + "'"

    			if FRectNewItem<>"" then
    				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    			end if

    			if FRectIsUsing<>"" then'
    				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
    			end if

    			if FRectMwDiv<>"" then
    				sqlStr = sqlStr + " and i.centermwdiv='" + FRectMwDiv + "'"
    			end if

    			sqlStr = sqlStr + " and i.shopitemid<>0"
    			sqlStr = sqlStr + " and s.realstock>0"
    			sqlStr = sqlStr + " order by totno desc"
			ELSE
    			sqlStr = "select i.itemid, i.regdate, i.itemname, i.mwdiv, IsNULL(o.itemoption,'0000') as itemoption, IsNULL(o.optionname,'') as itemoptionname, i.isusing, IsNULL(o.isusing,'Y') as optionusing,"
    			sqlStr = sqlStr + " s.realstock as totno ,"
    			sqlStr = sqlStr + " s.realstock*i.buycash as buysum, s.realstock*i.sellcash as sellsum "
    			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i,"
    			sqlStr = sqlStr + " [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
    			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on s.yyyymm='" + FRectYYYYMM + "' and s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"
    			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
    			sqlStr = sqlStr + " and s.itemgubun='10'"
    			sqlStr = sqlStr + " and s.itemid=i.itemid"
    			sqlStr = sqlStr + " and i.makerid='" + FRectMakerid + "'"

    			if FRectNewItem<>"" then
    				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    			end if

    			if FRectIsUsing<>"" then
    				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
    			end if

    			if FRectMwDiv<>"" then
    				sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
    			end if

    			sqlStr = sqlStr + " and i.itemid<>0"
    			sqlStr = sqlStr + " and i.itemid<>11406"  	''포장비
    			sqlStr = sqlStr + " and i.itemid<>6400"		''배송비
    			sqlStr = sqlStr + " and s.realstock>0"
    			sqlStr = sqlStr + " order by totno desc"
            end if
		end if
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
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).Fitemid 		= rsget("itemid")
				FItemList(i).Fregdate 		= rsget("regdate")
				FItemList(i).Fitemoption 	= rsget("itemoption")
				FItemList(i).Fitemname 		= db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).FMaeIpGubun	= rsget("mwdiv")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")

				FItemList(i).FIsUsing	= rsget("isusing")
				FItemList(i).FOptionUsing	= rsget("optionusing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end Sub

	public Sub GetMonthlyRealJeagoDetail()
		dim sqlStr

		if FRectGubun="sys" then
			''시스템재고
			IF FRectItemGubun<>"10" then
			    sqlStr = "select i.makerid, sum(s.totsysstock) as totno ,"
    			sqlStr = sqlStr + " Sum(s.totsysstock*i.shopsuplycash) as buysum, Sum(s.totsysstock*i.shopitemprice) as sellsum, c.isusing "
    			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
    			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item i"
    			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
    			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
    			sqlStr = sqlStr + " and s.itemgubun='"&FRectItemGubun&"'"
                sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
                sqlStr = sqlStr + " and s.itemid=i.shopitemid"
                sqlStr = sqlStr + " and s.itemoption=i.itemoption"

    			if FRectNewItem<>"" then
    				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    			end if

    			if FRectIsUsing<>"" then'
    				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
    			end if

    			if FRectMwDiv<>"" then
    				sqlStr = sqlStr + " and i.centermwdiv='" + FRectMwDiv + "'"
    			end if

    			sqlStr = sqlStr + " and i.shopitemid<>0"
    			sqlStr = sqlStr + " and s.totsysstock>0"
    			sqlStr = sqlStr + " group by i.makerid, c.isusing"
    			sqlStr = sqlStr + " order by totno desc"
			ELSE
    			sqlStr = "select i.makerid, sum(s.totsysstock) as totno ,"
    			sqlStr = sqlStr + " Sum(s.totsysstock*i.buycash) as buysum, Sum(s.totsysstock*i.sellcash) as sellsum, c.isusing "
    			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
    			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
    			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
    			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
    			sqlStr = sqlStr + " and s.itemgubun='10'"
    			sqlStr = sqlStr + " and s.itemid=i.itemid"

    			if FRectNewItem<>"" then
    				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    			end if

    			if FRectIsUsing<>"" then'
    				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
    			end if

    			if FRectMwDiv<>"" then
    				sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
    			end if

    			sqlStr = sqlStr + " and i.itemid<>0"
    			sqlStr = sqlStr + " and i.itemid<>11406"  	''포장비
    			sqlStr = sqlStr + " and i.itemid<>6400"		''배송비
    			sqlStr = sqlStr + " and s.totsysstock>0"
    			sqlStr = sqlStr + " group by i.makerid, c.isusing"
    			sqlStr = sqlStr + " order by totno desc"
		    END IF
		else
			''실사재고
			IF FRectItemGubun<>"10" then
			    sqlStr = "select i.makerid, sum(s.realstock) as totno ,"
    			sqlStr = sqlStr + " Sum(s.realstock*i.shopsuplycash) as buysum, Sum(s.realstock*i.shopitemprice) as sellsum, c.isusing "
    			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
    			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item i"
    			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
    			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
    			sqlStr = sqlStr + " and s.itemgubun='"&FRectItemGubun&"'"
                sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
                sqlStr = sqlStr + " and s.itemid=i.shopitemid"
                sqlStr = sqlStr + " and s.itemoption=i.itemoption"

    			if FRectNewItem<>"" then
    				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    			end if

    			if FRectIsUsing<>"" then'
    				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
    			end if

    			if FRectMwDiv<>"" then
    				sqlStr = sqlStr + " and i.centermwdiv='" + FRectMwDiv + "'"
    			end if

    			sqlStr = sqlStr + " and i.shopitemid<>0"
    			sqlStr = sqlStr + " and s.realstock>0"
    			sqlStr = sqlStr + " group by i.makerid, c.isusing"
    			sqlStr = sqlStr + " order by totno desc"

			ELSE
    			sqlStr = "select i.makerid, sum(s.realstock) as totno ,"
    			sqlStr = sqlStr + " Sum(s.realstock*i.buycash) as buysum, Sum(s.realstock*i.sellcash) as sellsum, c.isusing "
    			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
    			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
    			sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on i.makerid=c.userid"
    			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
    			sqlStr = sqlStr + " and s.itemgubun='10'"
    			sqlStr = sqlStr + " and s.itemid=i.itemid"

    			if FRectNewItem<>"" then
    				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
    			end if

    			if FRectIsUsing<>"" then'
    				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
    			end if

    			if FRectMwDiv<>"" then
    				sqlStr = sqlStr + " and i.mwdiv='" + FRectMwDiv + "'"
    			end if

    			sqlStr = sqlStr + " and i.itemid<>0"
    			sqlStr = sqlStr + " and i.itemid<>11406"
    			sqlStr = sqlStr + " and i.itemid<>6400"
    			sqlStr = sqlStr + " and s.realstock>0"
    			sqlStr = sqlStr + " group by i.makerid, c.isusing"
    			sqlStr = sqlStr + " order by totno desc"
    		end if
		end if

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
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FMakerid 	= rsget("makerid")

				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")

				FItemList(i).FMakerUsing	= rsget("isusing")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end Sub


	public Sub GetOFFMonthlyJeagoSum()
		dim sqlStr

		if FRectGubun="sys" then
			''시스템재고
			sqlStr = "select i.itemgubun, i.centermwdiv as mwdiv, sum(s.totsysstock) as totno ,"
			sqlStr = sqlStr + " Sum(s.totsysstock*i.shopsuplycash) as buysum, Sum(s.totsysstock*i.shopitemprice) as sellsum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item i"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun<>'10'"
			sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
			sqlStr = sqlStr + " and s.itemid=i.shopitemid"
			sqlStr = sqlStr + " and s.itemoption=i.itemoption"
			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
			sqlStr = sqlStr + " and i.shopitemid<>0"
			sqlStr = sqlStr + " and s.totsysstock>0"
			sqlStr = sqlStr + " group by i.itemgubun, i.centermwdiv"
			sqlStr = sqlStr + " order by IsNULL(i.centermwdiv,'Z') asc,i.itemgubun"
		else
			''실사재고
			sqlStr = "select i.itemgubun, i.centermwdiv as mwdiv, sum(s.realstock) as totno ,"
			sqlStr = sqlStr + " Sum(s.realstock*i.shopsuplycash) as buysum, Sum(s.realstock*i.shopitemprice) as sellsum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
			sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item i"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun<>'10'"
			sqlStr = sqlStr + " and s.itemgubun=i.itemgubun"
			sqlStr = sqlStr + " and s.itemid=i.shopitemid"
			sqlStr = sqlStr + " and s.itemoption=i.itemoption"
			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
			sqlStr = sqlStr + " and i.shopitemid<>0"
			sqlStr = sqlStr + " and s.realstock>0"
			sqlStr = sqlStr + " group by i.itemgubun, i.centermwdiv"
			sqlStr = sqlStr + " order by IsNULL(i.centermwdiv,'Z') asc, i.itemgubun"
		end if


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
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")

				FItemList(i).Fitemgubun		= rsget("itemgubun")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


	public Sub GetMonthlyJeagoSum()
		dim sqlStr

		if FRectGubun="sys" then
			''시스템재고
			sqlStr = "select i.mwdiv, sum(s.totsysstock) as totno ,"
			sqlStr = sqlStr + " Sum(s.totsysstock*i.buycash) as buysum, Sum(s.totsysstock*i.sellcash) as sellsum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s"
			sqlStr = sqlStr + "     join [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + "     on  s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + "     and s.itemgubun='10'"
			sqlStr = sqlStr + "     and s.itemid=i.itemid"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun='10'"
			sqlStr = sqlStr + " and s.itemid=i.itemid"
			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if
			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
			sqlStr = sqlStr + " and i.itemid<>0"
			sqlStr = sqlStr + " and i.itemid<>11406"  	''포장비
			sqlStr = sqlStr + " and i.itemid<>6400"		''배송비
			sqlStr = sqlStr + " and s.totsysstock>0"
			sqlStr = sqlStr + " group by i.mwdiv"
			sqlStr = sqlStr + " order by i.mwdiv"
		else
			''실사재고
			sqlStr = "select i.mwdiv, sum(s.realstock) as totno ,"
			sqlStr = sqlStr + " Sum(s.realstock*i.buycash) as buysum, Sum(s.realstock*i.sellcash) as sellsum "
			sqlStr = sqlStr + " from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary s,"
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where s.yyyymm='" + FRectYYYYMM + "'"
			sqlStr = sqlStr + " and s.itemgubun='10'"
			sqlStr = sqlStr + " and s.itemid=i.itemid"
			if FRectNewItem<>"" then
				sqlStr = sqlStr + " and (datediff(m, i.regdate ,'" + FRectYYYYMMDD + "') <= 3)"
			end if
			if FRectIsUsing<>"" then'
				sqlStr = sqlStr + " and i.isusing='" + FRectIsUsing + "'"
			end if
			sqlStr = sqlStr + " and i.itemid<>0"
			sqlStr = sqlStr + " and i.itemid<>11406"
			sqlStr = sqlStr + " and i.itemid<>6400"
			sqlStr = sqlStr + " and s.realstock>0"
			sqlStr = sqlStr + " group by i.mwdiv"
			sqlStr = sqlStr + " order by i.mwdiv"
		end if


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
				set FItemList(i) = new CMonthlyStockSum
				FItemList(i).FTotCount 		= rsget("totno")
				FItemList(i).FTotBuySum 	= rsget("buysum")
				FItemList(i).FTotSellSum 	= rsget("sellsum")
				FItemList(i).FMaeIpGubun 	= rsget("mwdiv")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'// 월별 입고내역
	public Sub GetMonthlyIpgoList()
    	dim sqlStr, sqlAdd

		if (FRectPlaceGubun = "") then
			FRectPlaceGubun = "L"
		end if

		sqlAdd = " from "
		sqlAdd = sqlAdd + " 	db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum s "
		sqlAdd = sqlAdd + " where "
		sqlAdd = sqlAdd + " 	1 = 1 "
		sqlAdd = sqlAdd + " 	and s.yyyymm >= '" + CStr(FRectStartYYYYMM) + "' "
		sqlAdd = sqlAdd + " 	and s.yyyymm <= '" + CStr(FRectEndYYYYMM) + "' "
		sqlAdd = sqlAdd + " 	and s.stockPlace = '" + CStr(FRectPlaceGubun) + "' "

		if (FRectTargetGbn <> "") then
			if (FRectTargetGbn = "3X") then
				sqlAdd = sqlAdd + " and s.targetGbn not in ('ET', 'EG') "
			else
				sqlAdd = sqlAdd + " and s.targetGbn='" + FRectTargetGbn + "'"
			end if
		end if

		if (FRectItemGubun <> "") then
			sqlAdd = sqlAdd + " 	and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
		end if

		if (FRectIpgoMWdiv <> "") then
			if (FRectIpgoMWdiv = "X") then
				sqlAdd = sqlAdd + " 	and s.ipgoMWdiv not in ('M', 'W') "
			else
				sqlAdd = sqlAdd + " 	and s.ipgoMWdiv = '" + CStr(FRectIpgoMWdiv) + "' "
			end if
		end if

		if (FRectMwDiv <> "") then
			if (FRectMwDiv = "X") then
				sqlAdd = sqlAdd + " 	and s.lastmwdiv not in ('M', 'W') "
			else
				sqlAdd = sqlAdd + " 	and s.lastmwdiv = '" + CStr(FRectMwDiv) + "' "
			end if
		end if

		if (FRectLastCenterMWDiv <> "") then
			if (FRectLastCenterMWDiv = "X") then
				sqlAdd = sqlAdd + " 	and IsNull(s.lastCenterMWDiv, 'Z') not in ('M', 'W') "
			else
				sqlAdd = sqlAdd + " 	and s.lastCenterMWDiv = '" + CStr(FRectLastCenterMWDiv) + "' "
			end if
		end if

		if (FRectItemMWdiv <> "") then
			sqlAdd = sqlAdd + " 	and s.itemMWdiv = '" + CStr(FRectItemMWdiv) + "' "
		end if

		if (FRectItemid <> "") then
			sqlAdd = sqlAdd + " 	and s.itemid = '" + CStr(FRectItemid) + "' "
		end if

		if (FRectMakerid <> "") then
			sqlAdd = sqlAdd + " 	and s.lastmakerid = '" + CStr(FRectMakerid) + "' "
		end if

		if (FRectShopid <> "") then
			sqlAdd = sqlAdd + " 	and s.shopid = '" + CStr(FRectShopid) + "' "
		end If

		if (FRectIpgoType <> "") then
			sqlAdd = sqlAdd + " 	and s.ipgoType = '" + CStr(FRectIpgoType) + "' "
		end if

		sqlStr = " SELECT count(*) as CNT, CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg "
		sqlStr = sqlStr + sqlAdd
		''response.write sqlStr
		''response.end

    	rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		FTotalPage = rsget("totPg")
    	rsget.Close

    	'지정페이지가 전체 페이지보다 클 때 함수종료
    	if CLng(FCurrPage)>CLng(FTotalPage) then
    		FResultCount = 0
    		exit sub
    	end if

		sqlStr = " SELECT top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " s.yyyymm, s.stockPlace, s.shopid, s.targetGbn, s.itemgubun, s.itemid, s.itemoption, s.ipgoMWdiv, s.itemMWdiv, s.itemVatInclude, s.totItemNo, s.totBuyCash, s.lastupdate, s.ipgoType, s.lastmwdiv, s.lastmakerid, s.lastvatinclude, s.lastCenterMWDiv "
		sqlStr = sqlStr + sqlAdd

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	s.yyyymm desc, s.stockPlace, s.targetGbn desc, s.shopid, s.itemgubun, s.itemid, s.itemoption, s.ipgoMWdiv "

		''response.write sqlStr
		''response.end

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
				set FItemList(i) = new CMonthlyStockIpgoItem

			    FItemList(i).Fyyyymm     		= rsget("yyyymm")
				FItemList(i).FstockPlace    	= rsget("stockPlace")
				FItemList(i).Fshopid     		= rsget("shopid")
				FItemList(i).FtargetGbn    		= rsget("targetGbn")
				FItemList(i).Fitemgubun     	= rsget("itemgubun")
				FItemList(i).Fitemid     		= rsget("itemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).FipgoMWdiv     	= rsget("ipgoMWdiv")
				FItemList(i).FitemMWdiv     	= rsget("itemMWdiv")
				FItemList(i).FitemVatInclude    = rsget("itemVatInclude")
				FItemList(i).FtotItemNo     	= rsget("totItemNo")
				FItemList(i).FtotBuyCash     	= rsget("totBuyCash")
				FItemList(i).Flastupdate     	= rsget("lastupdate")
				FItemList(i).FipgoType     		= rsget("ipgoType")

				FItemList(i).Flastmwdiv     	= rsget("lastmwdiv")
				FItemList(i).FlastCenterMWDiv  	= rsget("lastCenterMWDiv")
				FItemList(i).Flastmakerid     	= rsget("lastmakerid")
				FItemList(i).Flastvatinclude    = rsget("lastvatinclude")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'// 월별 입고내역(서머리)
	public Sub GetMonthlyIpgoSum()
    	dim sqlStr, sqlAdd

		if (FRectPlaceGubun = "") then
			FRectPlaceGubun = "L"
		end if

		sqlAdd = " from "
		sqlAdd = sqlAdd + " 	db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum s "
		sqlAdd = sqlAdd + " where "
		sqlAdd = sqlAdd + " 	1 = 1 "
		sqlAdd = sqlAdd + " 	and s.yyyymm >= '" + CStr(FRectStartYYYYMM) + "' "
		sqlAdd = sqlAdd + " 	and s.yyyymm <= '" + CStr(FRectEndYYYYMM) + "' "
		sqlAdd = sqlAdd + " 	and s.stockPlace = '" + CStr(FRectPlaceGubun) + "' "

		if (FRectTargetGbn <> "") then
			if (FRectTargetGbn = "3X") then
				sqlAdd = sqlAdd + " and s.targetGbn not in ('ET', 'EG') "
			else
				sqlAdd = sqlAdd + " and s.targetGbn='" + FRectTargetGbn + "'"
			end if
		end if

		if (FRectItemGubun <> "") then
			sqlAdd = sqlAdd + " 	and s.itemgubun = '" + CStr(FRectItemGubun) + "' "
		end if

		if (FRectIpgoMWdiv <> "") then
			if (FRectIpgoMWdiv = "X") then
				sqlAdd = sqlAdd + " 	and s.ipgoMWdiv not in ('M', 'W') "
			else
				sqlAdd = sqlAdd + " 	and s.ipgoMWdiv = '" + CStr(FRectIpgoMWdiv) + "' "
			end if
		end if

		if (FRectMwDiv <> "") then
			if (FRectMwDiv = "X") then
				sqlAdd = sqlAdd + " 	and s.lastmwdiv not in ('M', 'W') "
			else
				sqlAdd = sqlAdd + " 	and s.lastmwdiv = '" + CStr(FRectMwDiv) + "' "
			end if
		end if

		if (FRectLastCenterMWDiv <> "") then
			if (FRectLastCenterMWDiv = "X") then
				sqlAdd = sqlAdd + " 	and IsNull(s.lastCenterMWDiv, 'Z') not in ('M', 'W') "
			else
				sqlAdd = sqlAdd + " 	and s.lastCenterMWDiv = '" + CStr(FRectLastCenterMWDiv) + "' "
			end if
		end if

		if (FRectItemMWdiv <> "") then
			sqlAdd = sqlAdd + " 	and s.itemMWdiv = '" + CStr(FRectItemMWdiv) + "' "
		end if

		if (FRectItemid <> "") then
			sqlAdd = sqlAdd + " 	and s.itemid = '" + CStr(FRectItemid) + "' "
		end if

		if (FRectMakerid <> "") then
			sqlAdd = sqlAdd + " 	and s.lastmakerid = '" + CStr(FRectMakerid) + "' "
		end If

		if (FRectShopid <> "") then
			sqlAdd = sqlAdd + " 	and s.shopid = '" + CStr(FRectShopid) + "' "
		end If

		if (FRectIpgoType <> "") then
			sqlAdd = sqlAdd + " 	and s.ipgoType = '" + CStr(FRectIpgoType) + "' "
		end if

		sqlStr = " SELECT "
		sqlStr = sqlStr + " s.yyyymm, s.stockPlace, s.targetGbn, s.itemgubun, s.ipgoMWdiv, s.lastmwdiv, IsNull(s.lastCenterMWDiv,'Z') as lastCenterMWDiv, s.lastvatinclude, sum(s.totItemNo) as totItemNo, sum(s.totBuyCash) as totBuyCash "
		If (FRectShowShopid <> "") Then
			sqlStr = sqlStr + " , s.shopid "
		End If
		sqlStr = sqlStr + sqlAdd

		sqlStr = sqlStr + " group by "
		sqlStr = sqlStr + " 	s.yyyymm, s.stockPlace, s.targetGbn, s.itemgubun, s.ipgoMWdiv, s.lastvatinclude, s.lastmwdiv, IsNull(s.lastCenterMWDiv,'Z') "
		If (FRectShowShopid <> "") Then
			sqlStr = sqlStr + " , s.shopid "
		End If
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	s.yyyymm desc, s.stockPlace, s.targetGbn desc, s.itemgubun, s.ipgoMWdiv, s.lastvatinclude desc, s.lastmwdiv, IsNull(s.lastCenterMWDiv,'Z') "
		If (FRectShowShopid <> "") Then
			sqlStr = sqlStr + " , s.shopid "
		End If

		''response.write sqlStr
		''response.end

        rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CMonthlyStockIpgoItem

			    FItemList(i).Fyyyymm     		= rsget("yyyymm")
				FItemList(i).FstockPlace    	= rsget("stockPlace")
				If (FRectShowShopid <> "") Then
					FItemList(i).Fshopid     		= rsget("shopid")
				End If
				FItemList(i).FtargetGbn    		= rsget("targetGbn")
				FItemList(i).Fitemgubun     	= rsget("itemgubun")
				FItemList(i).FipgoMWdiv     	= rsget("ipgoMWdiv")

				FItemList(i).Flastmwdiv     	= rsget("lastmwdiv")
				FItemList(i).FlastCenterMWDiv   = rsget("lastCenterMWDiv")

				FItemList(i).Flastvatinclude   	= rsget("lastvatinclude")
				FItemList(i).FtotItemNo     	= rsget("totItemNo")
				FItemList(i).FtotBuyCash     	= rsget("totBuyCash")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'// 월별 입고내역-재고테이블 입고수량 불일치
	public Sub GetMonthlyIpgoDiff()
    	dim sqlStr, sqlAdd

		if (FRectPlaceGubun = "") then
			FRectPlaceGubun = "L"
		end if

		sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthlyLogisstock_ipgoSumDiff] '" + CStr(FRectYYYYMM) + "','" + CStr(FRectPlaceGubun) + "' "
		''response.write sqlStr
		''response.end

        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		''rsget.LockType = adLockOptimistic  ''==> adLockReadOnly

		rsget.Open sqlStr,dbget,1

		FTotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FTotalPage = FTotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CMonthlyStockIpgoItem

			    FItemList(i).Fyyyymm     		= rsget("yyyymm")
				FItemList(i).FstockPlace    	= rsget("stockPlace")
				FItemList(i).Fshopid    		= rsget("shopid")
				FItemList(i).Fitemgubun     	= rsget("itemgubun")
				FItemList(i).Fitemid     		= rsget("itemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).FtotItemNo     	= rsget("totItemNo")
				FItemList(i).FtotBuyCash     	= rsget("totBuyCash")
				FItemList(i).FstockIpgoNo     	= rsget("stockIpgoNo")		'// 재고테이블 입고수량

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'// 상품별 평균매입가(물류)
	public Sub GetMonthlyAvgPriceLogics()
    	dim sqlStr, sqlAdd

		if (FRectPlaceGubun = "") then
			FRectPlaceGubun = "L"
		end if

		if (FRectItemGubun = "") then
			FRectItemGubun = "10"
		end if

		sqlAdd = " from "
		sqlAdd = sqlAdd + " 	db_summary.dbo.tbl_monthly_accumulated_logisstock_summary a"
		sqlAdd = sqlAdd + " 	left join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary p"
		sqlAdd = sqlAdd + " 	on"
		sqlAdd = sqlAdd + " 		1 = 1"
		sqlAdd = sqlAdd + " 		and p.yyyymm = convert(varchar(7), DATEADD (mm, -1, a.yyyymm + '-01'), 121)"
		sqlAdd = sqlAdd + " 		and p.itemgubun = a.itemgubun"
		sqlAdd = sqlAdd + " 		and p.itemid = a.itemid"
		sqlAdd = sqlAdd + " 		and p.itemoption = a.itemoption"
		sqlAdd = sqlAdd + " 	left join ("
		sqlAdd = sqlAdd + " 		select i.yyyymm, i.stockPlace, i.itemgubun, i.itemid, i.itemoption, i.ipgoMWdiv, sum(i.totItemNo) as totItemNo, sum(i.totBuyCash) as totBuyCash"
		sqlAdd = sqlAdd + " 		from"
		sqlAdd = sqlAdd + " 			db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum i"
		sqlAdd = sqlAdd + " 		where"
		sqlAdd = sqlAdd + " 			1 = 1"
		sqlAdd = sqlAdd + " 			and i.yyyymm >= '" + CStr(FRectStartYYYYMM) + "'"
		sqlAdd = sqlAdd + " 			and i.yyyymm <= '" + CStr(FRectEndYYYYMM) + "'"
		sqlAdd = sqlAdd + " 			and i.stockPlace = 'L'"
		sqlAdd = sqlAdd + " 			and i.itemgubun = '" + CStr(FRectItemGubun) + "'"
		sqlAdd = sqlAdd + " 			and i.itemid = " + CStr(FRectItemid) + ""

		if (FRectItemOption <> "") then
			sqlAdd = sqlAdd + " 			and i.itemoption = '" + CStr(FRectItemOption) + "'"
		end if

		sqlAdd = sqlAdd + " 		group by"
		sqlAdd = sqlAdd + " 			i.yyyymm, i.stockPlace, i.itemgubun, i.itemid, i.itemoption, i.ipgoMWdiv"
		sqlAdd = sqlAdd + " 	) I"
		sqlAdd = sqlAdd + " 	on"
		sqlAdd = sqlAdd + " 		1 = 1"
		sqlAdd = sqlAdd + " 		and a.yyyymm = i.yyyymm"
		sqlAdd = sqlAdd + " 		and a.itemgubun = i.itemgubun"
		sqlAdd = sqlAdd + " 		and a.itemid = i.itemid"
		sqlAdd = sqlAdd + " 		and a.itemoption = i.itemoption"
		sqlAdd = sqlAdd + " 		and a.lastmwdiv = i.ipgoMWdiv"
		sqlAdd = sqlAdd + " 	left join ("
		sqlAdd = sqlAdd + " 		select"
		sqlAdd = sqlAdd + " 			a.yyyymm, 'S' as stockPlace"
		sqlAdd = sqlAdd + " 			, a.itemgubun, a.itemid, a.itemoption"
		sqlAdd = sqlAdd + " 			, sum(IsNull(a.sysstockno, 0)) as totsysstock"
		sqlAdd = sqlAdd + " 			, sum(IsNull(a.sysstockno, 0)*IsNull(IsNull(a.avgShopIpgoPrice, a.LstBuyCash), 0)) as totsysstockBuySum"
		sqlAdd = sqlAdd + " 			, (case"
		sqlAdd = sqlAdd + " 							when IsNull(a.LstComm_cd, 'Z') in ('B011', 'B012', 'B013', 'W') then 'W'"
		sqlAdd = sqlAdd + " 							when IsNull(a.LstComm_cd, 'Z') in ('Z', '') then 'Z'"
		sqlAdd = sqlAdd + " 							else 'M' end) as lastmwdiv"
		sqlAdd = sqlAdd + " 		from"
		sqlAdd = sqlAdd + " 			db_summary.dbo.tbl_monthly_accumulated_shopstock_summary a"
		sqlAdd = sqlAdd + " 		where"
		sqlAdd = sqlAdd + " 			1 = 1"
		sqlAdd = sqlAdd + " 			and a.yyyymm >= convert(varchar(7), DATEADD (mm, -1, '" + CStr(FRectStartYYYYMM) + "' + '-01'), 121)"
		sqlAdd = sqlAdd + " 			and a.yyyymm <= convert(varchar(7), DATEADD (mm, -1, '" + CStr(FRectEndYYYYMM) + "' + '-01'), 121)"
		sqlAdd = sqlAdd + " 			and a.itemgubun = '" + CStr(FRectItemGubun) + "'"
		sqlAdd = sqlAdd + " 			and a.itemid = " + CStr(FRectItemid) + ""

		if (FRectItemOption <> "") then
			sqlAdd = sqlAdd + " 			and a.itemoption = '" + CStr(FRectItemOption) + "'"
		end if

		sqlAdd = sqlAdd + " 		group by"
		sqlAdd = sqlAdd + " 			a.yyyymm"
		sqlAdd = sqlAdd + " 			, a.itemgubun, a.itemid, a.itemoption"
		sqlAdd = sqlAdd + " 			, (case"
		sqlAdd = sqlAdd + " 							when IsNull(a.LstComm_cd, 'Z') in ('B011', 'B012', 'B013', 'W') then 'W'"
		sqlAdd = sqlAdd + " 							when IsNull(a.LstComm_cd, 'Z') in ('Z', '') then 'Z'"
		sqlAdd = sqlAdd + " 							else 'M' end)"
		sqlAdd = sqlAdd + " 	) S"
		sqlAdd = sqlAdd + " 	on"
		sqlAdd = sqlAdd + " 		1 = 1"
		sqlAdd = sqlAdd + " 		and S.yyyymm = convert(varchar(7), DATEADD (mm, -1, a.yyyymm + '-01'), 121)"
		sqlAdd = sqlAdd + " 		and S.itemgubun = a.itemgubun"
		sqlAdd = sqlAdd + " 		and S.itemid = a.itemid"
		sqlAdd = sqlAdd + " 		and S.itemoption = a.itemoption"
		sqlAdd = sqlAdd + " 		and S.lastmwdiv = a.lastmwdiv"
		sqlAdd = sqlAdd + " where"
		sqlAdd = sqlAdd + " 	1 = 1"
		sqlAdd = sqlAdd + " 	and a.yyyymm >= '" + CStr(FRectStartYYYYMM) + "'"
		sqlAdd = sqlAdd + " 	and a.yyyymm <= '" + CStr(FRectEndYYYYMM) + "'"
		sqlAdd = sqlAdd + " 	and a.itemgubun = '" + CStr(FRectItemGubun) + "'"
		sqlAdd = sqlAdd + " 	and a.itemid = " + CStr(FRectItemid) + ""

		if (FRectItemOption <> "") then
			sqlAdd = sqlAdd + " 	and a.itemoption = '" + CStr(FRectItemOption) + "'"
		end if

		if (FRectMwDiv <> "") then
			if (FRectMwDiv = "X") then
				sqlAdd = sqlAdd + " 	and IsNull(s.lastmwdiv, '') not in ('M', 'W') "
			else
				sqlAdd = sqlAdd + " 	and s.lastmwdiv = '" + CStr(FRectMwDiv) + "' "
			end if
		end if


		sqlStr = " SELECT count(*) as CNT, CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg "
		sqlStr = sqlStr + sqlAdd
		''response.write sqlStr
		''dbget.close
		''response.end

    	rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		FTotalPage = rsget("totPg")
    	rsget.Close

    	'지정페이지가 전체 페이지보다 클 때 함수종료
    	if CLng(FCurrPage)>CLng(FTotalPage) then
    		FResultCount = 0
    		exit sub
    	end if


		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	a.yyyymm, '" + CStr(FRectPlaceGubun) + "' as stockPlace, '' as shopid "
		sqlStr = sqlStr + " 	, a.itemgubun, a.itemid, a.itemoption "
		sqlStr = sqlStr + " 	, IsNull(p.totsysstock, 0) as totsysstockPrev "
		sqlStr = sqlStr + " 	, IsNull(p.avgipgoPrice, p.lastbuyPrice)*IsNull(p.totsysstock, 0) as avgipgoPriceSumPrev "
		sqlStr = sqlStr + " 	, IsNull(S.totsysstock, 0) as totsysstockShopPrev "
		sqlStr = sqlStr + " 	, IsNull(S.totsysstockBuySum, 0) as totsysstockBuySumShopPrev "
		sqlStr = sqlStr + " 	, IsNull(I.totItemNo, 0) as totItemNo "
		sqlStr = sqlStr + " 	, IsNull(I.totBuyCash, 0) as totBuyCash "
		sqlStr = sqlStr + " 	, IsNull(a.totsysstock, 0) as totsysstock "
		sqlStr = sqlStr + " 	, IsNull(p.avgipgoPrice, p.lastbuyPrice) as avgipgoPricePrev "
		sqlStr = sqlStr + " 	, IsNull(a.avgipgoPrice, a.lastbuyPrice) as avgipgoPrice "
		sqlStr = sqlStr + " 	, p.lastmwdiv as lastmwdivPrev "
		sqlStr = sqlStr + " 	, IsNull(a.lastmwdiv, 'Z') as lastmwdiv "
		sqlStr = sqlStr + " 	, p.lastmakerid as makeridPrev "
		sqlStr = sqlStr + " 	, IsNull(a.lastmakerid, 'Z') as makerid "
		sqlStr = sqlStr + sqlAdd

		sqlStr = sqlStr + " order by"
		sqlStr = sqlStr + " 	a.yyyymm, a.itemgubun, a.itemid, a.itemoption"

		''response.write sqlStr
		''response.end

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
				set FItemList(i) = new CMonthlyStockAvgPriceItem

			    FItemList(i).Fyyyymm     			= rsget("yyyymm")
				FItemList(i).FstockPlace     		= rsget("stockPlace")
				FItemList(i).Fshopid     			= rsget("shopid")
				FItemList(i).Fitemgubun     		= rsget("itemgubun")
				FItemList(i).Fitemid     			= rsget("itemid")
				FItemList(i).Fitemoption     		= rsget("itemoption")
				FItemList(i).FtotsysstockPrev     			= rsget("totsysstockPrev")
				FItemList(i).FavgipgoPriceSumPrev     		= rsget("avgipgoPriceSumPrev")
				FItemList(i).FtotsysstockShopPrev     		= rsget("totsysstockShopPrev")
				FItemList(i).FtotsysstockBuySumShopPrev     = rsget("totsysstockBuySumShopPrev")
				FItemList(i).FtotItemNo     		= rsget("totItemNo")
				FItemList(i).FtotBuyCash     		= rsget("totBuyCash")
				FItemList(i).Ftotsysstock     		= rsget("totsysstock")
				FItemList(i).FavgipgoPricePrev 		= rsget("avgipgoPricePrev")
				FItemList(i).FavgipgoPrice     		= rsget("avgipgoPrice")
				FItemList(i).FlastmwdivPrev     	= rsget("lastmwdivPrev")
				FItemList(i).Flastmwdiv     		= rsget("lastmwdiv")
				FItemList(i).FmakeridPrev     		= rsget("makeridPrev")
				FItemList(i).Fmakerid     			= rsget("makerid")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'// 상품 매입구분 로그(물류)
	public Sub GetMonthlyMWDivHistoryLogics()
    	dim sqlStr, sqlAdd

		FCurrPage = 1
		FPageSize = 100

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " a.yyyymm, a.itemgubun, a.ItemId, a.Itemoption, '' as shopid, a.lastmwdiv as mwdiv, a.lastmakerid as makerid, a.avgipgoPrice, a.lastbuyPrice as buyPrice"
		sqlStr = sqlStr + ", convert(varchar(19),a.regdate,21) as regdate, convert(varchar(19),a.lastupdate,21) as lastupdate, a.lastvatinclude, a.lastIpgoDate, a.firstIpgoDate "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_summary.dbo.tbl_monthly_accumulated_logisstock_summary a "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and a.itemgubun = '" + CStr(FRectItemGubun) + "' "
		sqlStr = sqlStr + " 	and a.itemid = " + CStr(FRectItemId) + " "
		sqlStr = sqlStr + " 	and a.itemoption = '" + CStr(FRectItemOption) + "' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	a.yyyymm desc "

		''response.write sqlStr
		''response.end

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount

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
				set FItemList(i) = new CMonthlyStockItem

			    FItemList(i).Fyyyymm     			= rsget("yyyymm")
				FItemList(i).FItemGubun     		= rsget("ItemGubun")
				FItemList(i).FItemId     			= rsget("ItemId")
				FItemList(i).FItemoption     		= rsget("Itemoption")
				FItemList(i).Fshopid     			= rsget("shopid")
				FItemList(i).Fmwdiv     			= rsget("mwdiv")
				FItemList(i).Fmakerid     			= rsget("makerid")
				FItemList(i).FavgipgoPrice     		= rsget("avgipgoPrice")
				FItemList(i).FbuyPrice     			= rsget("buyPrice")
				FItemList(i).Fregdate     			= rsget("regdate")
				FItemList(i).Flastupdate     		= rsget("lastupdate")
				FItemList(i).Flastvatinclude   		= rsget("lastvatinclude")
				FItemList(i).FlastIpgoDate   		= rsget("lastIpgoDate")
				FItemList(i).FfirstIpgoDate   		= rsget("firstIpgoDate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

	'// 상품 매입구분 로그(매장)
	public Sub GetMonthlyMWDivHistoryShop()
    	dim sqlStr, sqlAdd

		FCurrPage = 1
		FPageSize = 100

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " a.yyyymm, a.itemgubun, a.ItemId, a.Itemoption, a.shopid, a.LstComm_cd as mwdiv, a.LstCenterMwDiv as centermwdiv, a.LstMakerid as makerid, a.avgShopIpgoPrice as avgipgoPrice, a.LstBuyCash as buyPrice"
		sqlStr = sqlStr + ", convert(Varchar(19),a.regdate,21) as regdate"
		sqlStr = sqlStr + ", convert(Varchar(19),a.lastupdate,21) as lastupdate, a.lastIpgoDate, a.lastIpgoDateLogics "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_summary.dbo.tbl_monthly_accumulated_shopstock_summary a "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and a.itemgubun = '" + CStr(FRectItemGubun) + "' "
		sqlStr = sqlStr + " 	and a.itemid = " + CStr(FRectItemId) + " "
		sqlStr = sqlStr + " 	and a.itemoption = '" + CStr(FRectItemOption) + "' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	a.yyyymm desc, a.shopid desc "

		''response.write sqlStr
		''response.end

        rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount

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
				set FItemList(i) = new CMonthlyStockItem

			    FItemList(i).Fyyyymm     			= rsget("yyyymm")
				FItemList(i).FItemGubun     		= rsget("ItemGubun")
				FItemList(i).FItemId     			= rsget("ItemId")
				FItemList(i).FItemoption     		= rsget("Itemoption")
				FItemList(i).Fshopid     			= rsget("shopid")
				FItemList(i).Fmwdiv     			= rsget("mwdiv")
				FItemList(i).Fcentermwdiv  			= rsget("centermwdiv")
				FItemList(i).Fmakerid     			= rsget("makerid")
				FItemList(i).FavgipgoPrice     		= rsget("avgipgoPrice")
				FItemList(i).FbuyPrice     			= rsget("buyPrice")
				FItemList(i).Fregdate     			= rsget("regdate")
				FItemList(i).Flastupdate     		= rsget("lastupdate")

				FItemList(i).FlastIpgoDate     		= rsget("lastIpgoDate")
				FItemList(i).FlastIpgoDateLogics	= rsget("lastIpgoDateLogics")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub

    public Sub GetMonthlyNullIpgoInfo()
		dim sqlStr

        sqlStr = "select max(yyyymm) as maxyyyymm, min(yyyymm) as minyyyymm "
        sqlStr = sqlStr + " ,MIN(lastipgodate) as minIpgodate,MAX(lastipgodate) as maxIpgodate"
        sqlStr = sqlStr + " ,sum(CASE WHEN lastipgodate is NULL THEN 1 ELSE 0 END) nullCNT"
        sqlStr = sqlStr + " from"
        sqlStr = sqlStr + " [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and itemgubun = '" + CStr(FRectItemGubun) + "' "
		sqlStr = sqlStr + " 	and itemid =" + CStr(FRectItemId) + " "
		sqlStr = sqlStr + " 	and itemoption ='" + CStr(FRectItemOption) + "' "
 	    rsget.Open sqlStr,dbget,1

		set FOneItem = new CMonthlyErrorStockItem

		if  not rsget.EOF  then
			FOneItem.FMAX_YYYYMM = rsget("maxyyyymm")
			FOneItem.FMIN_YYYYMM = rsget("minyyyymm")

			FOneItem.FminIpgodate = rsget("minIpgodate")
			FOneItem.FmaxIpgodate = rsget("maxIpgodate")
			FOneItem.FnullCNT     = rsget("nullCNT")
		end if
		rsget.Close
	end Sub

	public Sub GetMonthlyErrorInfo()
		dim sqlStr

		sqlStr = " select "
		sqlStr = sqlStr + " 	max(yyyymm) as maxyyyymm, min(yyyymm) as minyyyymm "
		sqlStr = sqlStr + " 	, sum(case when IsNull(lastmwdiv,'') = 'M' then 1 else 0 end) as Mcnt "
		sqlStr = sqlStr + " 	, sum(case when IsNull(lastmwdiv,'') = 'W' then 1 else 0 end) as Wcnt "
		sqlStr = sqlStr + " 	, sum(case when IsNull(lastmwdiv,'') not in ('M', 'W') then 1 else 0 end) as Ecnt "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and itemgubun = '" + CStr(FRectItemGubun) + "' "
		sqlStr = sqlStr + " 	and itemid =" + CStr(FRectItemId) + " "
		sqlStr = sqlStr + " 	and itemoption ='" + CStr(FRectItemOption) + "' "

		rsget.Open sqlStr,dbget,1

		set FOneItem = new CMonthlyErrorStockItem

		if  not rsget.EOF  then
			FOneItem.FMAX_YYYYMM = rsget("maxyyyymm")
			FOneItem.FMIN_YYYYMM = rsget("minyyyymm")

			FOneItem.FMaeipCount = rsget("Mcnt")
			FOneItem.FWitakCount = rsget("Wcnt")
			FOneItem.FErrorCount = rsget("Ecnt")
		end if
		rsget.Close
	end Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 100
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

Sub drawSelectBoxAccShop(yyyymm, makerid, selectBoxName, selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value="" <%if selectedId="" then response.write " selected"%>>ALL</option><%
   query1 = " select distinct shopid from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary "
   query1 = query1 + " where yyyymm = '" + CStr(yyyymm) + "' "

   if (makerid <> "") then
	   query1 = query1 + " and IsNull(LstMakerid, 'Z') = '" + CStr(makerid) + "' "
   end if

   query1 = query1 + " order by shopid "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("shopid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("shopid")&"' "&tmp_str&">" + rsget("shopid") + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawSelectBoxBuseoGubunWith3PL(selectBoxName, selectedId)
%>
	<select class="select" name="<%=selectBoxName%>" class="select">
		<option value="">선택</option>
		<option value="3X" <% if (selectedId="3X") then response.write "selected" %> >텐바이텐(3PL제외)</option>
		<option value="ON" <% if (selectedId="ON") then response.write "selected" %> >온라인</option>
		<option value="OF" <% if (selectedId="OF") then response.write "selected" %> >오프라인</option>
		<option value="IT" <% if (selectedId="IT") then response.write "selected" %> >아이띵소(구)</option>
		<option value="ET" <% if (selectedId="ET") then response.write "selected" %> >3PL(아이띵소)</option>
		<option value="EG" <% if (selectedId="EG") then response.write "selected" %> >3PL(유그레잇)</option>
        <option value="3P" <% if (selectedId="3P") then response.write "selected" %> >3PL</option>
	</select>
<%
End Sub

%>
